"""Image processing helpers for conversion."""

from __future__ import annotations

import base64
import re
import uuid
from io import BytesIO
from pathlib import Path

import httpx
from docx import Document
from docx.image.exceptions import UnrecognizedImageError
from docx.shared import Inches
from PIL import Image

from ..config import Config
from .common import print_error, print_info


def process_image_content(image_content: bytes, url: str, local_dir: str = "./images") -> str:
    """Process image content, convert format and save, then return the local path."""
    Path(local_dir).mkdir(parents=True, exist_ok=True)

    image = Image.open(BytesIO(image_content))
    original_format = image.format.lower() if image.format else "png"

    supported_formats = ["png", "jpeg", "jpg"]
    if original_format not in supported_formats:
        if image.mode in ("RGBA", "LA") or "transparency" in image.info:
            target_format = "png"
        else:
            target_format = "jpeg"
    else:
        target_format = original_format

    url_filename = url.split("/")[-1].split("?")[0]
    name_without_ext = Path(url_filename).stem if url_filename else str(uuid.uuid4())
    local_filename = f"{name_without_ext}.{target_format}"
    local_path = Path(local_dir) / local_filename

    if original_format != target_format:
        if target_format == "jpeg" and image.mode in ("RGBA", "LA"):
            background = Image.new("RGB", image.size, (255, 255, 255))
            if image.mode == "RGBA":
                background.paste(image, mask=image.split()[-1])
            else:
                background.paste(image)
            image = background

        image.save(local_path, format=target_format.upper())
        print_info(f"Downloaded and converted image: {url} ({original_format} -> {target_format}) -> {local_path}")
    else:
        with open(local_path, "wb") as f:
            f.write(image_content)
        print_info(f"Downloaded image: {url} -> {local_path}")

    return str(local_path)


def download_image(url: str, config: Config) -> str | None:
    """Download an image and return the local file path."""
    local_dir = config.image_local_dir
    headers = {"User-Agent": config.image_user_agent}
    timeout = config.image_download_timeout

    try:
        Path(local_dir).mkdir(parents=True, exist_ok=True)

        with httpx.Client() as client:
            response = client.get(url, timeout=timeout, headers=headers, follow_redirects=True)
            response.raise_for_status()
            image_content = response.content

        return process_image_content(image_content, url, local_dir=local_dir)
    except Exception as e:
        print_error(f"Failed to download image {url}: {e}")
        return None


def ensure_local_image_compatible(image_path: str, local_dir: str = "./images") -> str | None:
    """Ensure a local image is in a docx-supported format."""
    path = Path(image_path)
    if not path.exists():
        print_error(f"Local image not found: {image_path}")
        return None

    try:
        image_content = path.read_bytes()
    except Exception as e:
        print_error(f"Failed to read local image {image_path}: {e}")
        return None

    try:
        image = Image.open(BytesIO(image_content))
        original_format = image.format.lower() if image.format else "png"
        image.verify()
    except Exception as e:
        print_error(f"Cannot recognize local image {image_path}: {e}")
        return None

    if original_format in ("png", "jpeg", "jpg"):
        return str(path)

    try:
        return process_image_content(image_content, path.name, local_dir=local_dir)
    except Exception as e:
        print_error(f"Failed to convert local image {image_path}: {e}")
        return None


def decode_data_uri_image(data_uri: str, local_dir: str = "./images") -> str | None:
    """Decode a data URI and save it as a local image."""
    if not data_uri.startswith("data:") or "base64," not in data_uri:
        return None

    try:
        _, b64_data = data_uri.split("base64,", 1)
        image_content = base64.b64decode(b64_data)
    except Exception as e:
        print_error(f"Failed to decode data URI: {e}")
        return None

    try:
        name_hint = f"inline_{uuid.uuid4().hex}"
        return process_image_content(image_content, name_hint, local_dir=local_dir)
    except Exception as e:
        print_error(f"Failed to process data URI image: {e}")
        return None


def _extract_img_attr(tag: str, attr: str) -> str | None:
    """Extract an attribute from an img tag."""
    match = re.search(rf'{attr}\s*=\s*(["\'])(.*?)\1', tag, flags=re.IGNORECASE)
    if match:
        return match.group(2)
    match = re.search(rf"{attr}\s*=\s*([^\s>]+)", tag, flags=re.IGNORECASE)
    if match:
        return match.group(1)
    return None


def _replace_img_src(tag: str, new_src: str) -> str:
    """Replace the src attribute in an img tag."""
    replacement = f'src="{new_src}"'
    updated = re.sub(r'\bsrc\s*=\s*([\'"])(.*?)\1', lambda _m: replacement, tag, flags=re.IGNORECASE)
    if updated != tag:
        return updated
    updated = re.sub(r"\bsrc\s*=\s*([^\s>]+)", lambda _m: replacement, tag, flags=re.IGNORECASE)
    if updated != tag:
        return updated

    alt = _extract_img_attr(tag, "alt")
    if alt:
        return f'<img src="{new_src}" alt="{alt}">'
    return f'<img src="{new_src}">'


def sanitize_html_images(html_content: str, config: Config) -> str:
    """Process HTML images and ensure they are usable."""
    img_pattern = re.compile(r"<img\b[^>]*>", flags=re.IGNORECASE)
    local_dir = config.image_local_dir

    def replace_img(match: re.Match[str]) -> str:
        tag = match.group(0)
        src = _extract_img_attr(tag, "src")
        alt = _extract_img_attr(tag, "alt") or ""

        if not src:
            return alt
        if src.startswith(("http://", "https://")):
            local_path = download_image(src, config)
            if local_path:
                return _replace_img_src(tag, local_path)
            print_info(f"Image download failed, skipping: {src}")
            return alt
        if src.startswith("data:"):
            local_path = decode_data_uri_image(src, local_dir=local_dir)
            if local_path:
                return _replace_img_src(tag, local_path)
            print_info("Data URI image processing failed, skipping")
            return alt

        compatible_path = ensure_local_image_compatible(src, local_dir=local_dir)
        if compatible_path:
            return _replace_img_src(tag, compatible_path)

        print_info(f"Local image unavailable, skipping: {src}")
        return alt

    return img_pattern.sub(replace_img, html_content)


def is_docx_image_supported(image_path: str) -> bool:
    """Check whether an image can be recognized by python-docx."""
    try:
        test_doc = Document()
        test_doc.add_picture(image_path)
        return True
    except UnrecognizedImageError:
        return False
    except Exception as e:
        print_error(f"Failed to check image {image_path}: {e}")
        return False


def filter_unrecognized_images(html_content: str) -> str:
    """Remove image tags that docx cannot recognize."""
    img_pattern = re.compile(r"<img\b[^>]*>", flags=re.IGNORECASE)

    def replace_img(match: re.Match[str]) -> str:
        tag = match.group(0)
        src = _extract_img_attr(tag, "src")
        alt = _extract_img_attr(tag, "alt") or ""

        if not src:
            return alt
        if src.startswith(("http://", "https://", "data:")):
            print_info(f"Unprocessed image link, skipping: {src}")
            return alt
        if not is_docx_image_supported(src):
            print_info(f"Image cannot be recognized, skipping: {src}")
            return alt
        return tag

    return img_pattern.sub(replace_img, html_content)


def process_markdown_images(markdown_content: str, config: Config) -> str:
    """Process markdown image links, downloading remote images to local files."""
    image_pattern = r"!\[([^\]]*)\]\(([^)]+)\)"

    def replace_image(match: re.Match[str]) -> str:
        alt_text = match.group(1)
        image_url = match.group(2)

        if image_url.startswith(("http://", "https://")):
            local_path = download_image(image_url, config)
            if local_path:
                return f"![{alt_text}]({local_path})"
            print_info(f"Image download failed, skipping: {image_url}")
            return alt_text or ""
        return match.group(0)

    return re.sub(image_pattern, replace_image, markdown_content)


def resize_images_in_document(document, max_width_inches: float = 6.0) -> None:
    """Resize images in a document to fit the configured maximum width."""
    try:
        for shape in document.inline_shapes:
            if hasattr(shape, "type") and "PICTURE" in str(shape.type):
                current_width_inches = shape.width.inches
                current_height_inches = shape.height.inches

                if current_width_inches > max_width_inches:
                    scale_ratio = max_width_inches / current_width_inches
                    new_height_inches = current_height_inches * scale_ratio

                    shape.width = Inches(max_width_inches)
                    shape.height = Inches(new_height_inches)

                    print_info(
                        f"Resized image: {current_width_inches:.2f}x{current_height_inches:.2f} -> "
                        f"{max_width_inches:.2f}x{new_height_inches:.2f} inches"
                    )
    except Exception as e:
        print_error(f"Error resizing images: {e}")


__all__ = [
    "decode_data_uri_image",
    "download_image",
    "ensure_local_image_compatible",
    "filter_unrecognized_images",
    "is_docx_image_supported",
    "process_image_content",
    "process_markdown_images",
    "resize_images_in_document",
    "sanitize_html_images",
]
