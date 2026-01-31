"""
Core converter module for md2word.
Converts Markdown content to Word documents.
"""

from __future__ import annotations

import base64
import re
import uuid
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING

import httpx
import markdown2
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.image.exceptions import UnrecognizedImageError
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from html4docx import HtmlToDocx
from PIL import Image

from .config import Config, StyleConfig
from .latex import extract_latex_formulas, replace_formula_placeholders

if TYPE_CHECKING:
    pass


def print_info(message: str) -> None:
    """Print info message."""
    print(f"[INFO] {message}")


def print_error(message: str) -> None:
    """Print error message."""
    print(f"[ERROR] {message}")


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Convert hex color to RGB tuple."""
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (r, g, b)


# Chinese number mapping
CHINESE_NUMBERS = [
    "零",
    "一",
    "二",
    "三",
    "四",
    "五",
    "六",
    "七",
    "八",
    "九",
    "十",
    "十一",
    "十二",
    "十三",
    "十四",
    "十五",
    "十六",
    "十七",
    "十八",
    "十九",
    "二十",
]


def number_to_chinese(n: int) -> str:
    """Convert number to Chinese."""
    if n <= 20:
        return CHINESE_NUMBERS[n]
    return str(n)


class HeadingNumbering:
    """Heading numbering manager."""

    FORMATS = {
        "chapter": "第{n}章",
        "section": "第{n}节",
        "chinese": "{n}、",
        "chinese_paren": "（{n}）",
        "arabic": "{n}.",
        "arabic_paren": "({n})",
        "arabic_bracket": "[{n}]",
        "roman": "{n}.",
        "roman_lower": "{n}.",
        "letter": "{n}.",
        "letter_lower": "{n}.",
        "circle": "{n}",
        "none": "",
    }

    ROMAN_NUMERALS = [
        "",
        "I",
        "II",
        "III",
        "IV",
        "V",
        "VI",
        "VII",
        "VIII",
        "IX",
        "X",
        "XI",
        "XII",
        "XIII",
        "XIV",
        "XV",
        "XVI",
        "XVII",
        "XVIII",
        "XIX",
        "XX",
    ]

    CIRCLE_NUMBERS = [
        "⓪",
        "①",
        "②",
        "③",
        "④",
        "⑤",
        "⑥",
        "⑦",
        "⑧",
        "⑨",
        "⑩",
        "⑪",
        "⑫",
        "⑬",
        "⑭",
        "⑮",
        "⑯",
        "⑰",
        "⑱",
        "⑲",
        "⑳",
    ]

    def __init__(self):
        self.counters = {}

    def reset(self, level: int | None = None):
        """Reset counters."""
        if level is None:
            self.counters = {}
        else:
            for lvl in list(self.counters.keys()):
                if lvl >= level:
                    self.counters[lvl] = 0

    def get_number(self, level: int, format_name: str | None) -> str:
        """Get numbering for specified level."""
        if not format_name or format_name == "none":
            return ""

        if level not in self.counters:
            self.counters[level] = 0
        self.counters[level] += 1

        for lvl in list(self.counters.keys()):
            if lvl > level:
                self.counters[lvl] = 0

        n = self.counters[level]

        if format_name in ("chapter", "section"):
            chinese_n = number_to_chinese(n)
            return self.FORMATS[format_name].format(n=chinese_n)
        elif format_name in ("chinese", "chinese_paren"):
            chinese_n = number_to_chinese(n)
            return self.FORMATS[format_name].format(n=chinese_n)
        elif format_name in ("arabic", "arabic_paren", "arabic_bracket"):
            return self.FORMATS[format_name].format(n=n)
        elif format_name == "roman":
            roman = self.ROMAN_NUMERALS[n] if n <= 20 else str(n)
            return f"{roman}."
        elif format_name == "roman_lower":
            roman = self.ROMAN_NUMERALS[n].lower() if n <= 20 else str(n)
            return f"{roman}."
        elif format_name == "letter":
            letter = chr(ord("A") + n - 1) if n <= 26 else str(n)
            return f"{letter}."
        elif format_name == "letter_lower":
            letter = chr(ord("a") + n - 1) if n <= 26 else str(n)
            return f"{letter}."
        elif format_name == "circle":
            return self.CIRCLE_NUMBERS[n] if n <= 20 else f"({n})"
        else:
            try:
                return format_name.format(n=n, cn=number_to_chinese(n))
            except (KeyError, ValueError):
                return f"{n}. "


# Image processing functions
def process_image_content(image_content: bytes, url: str, local_dir: str = "./images") -> str:
    """Process image content, convert format and save, return local path."""
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
    """Download image and return local file path."""
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
    """Ensure local image is in docx-supported format."""
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
    """Decode data URI and save as local image."""
    if not data_uri.startswith("data:"):
        return None
    if "base64," not in data_uri:
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
    """Extract attribute from img tag."""
    match = re.search(rf'{attr}\s*=\s*(["\'])(.*?)\1', tag, flags=re.IGNORECASE)
    if match:
        return match.group(2)
    match = re.search(rf"{attr}\s*=\s*([^\s>]+)", tag, flags=re.IGNORECASE)
    if match:
        return match.group(1)
    return None


def _replace_img_src(tag: str, new_src: str) -> str:
    """Replace src attribute in img tag."""
    replacement = f'src="{new_src}"'
    updated = re.sub(r'\bsrc\s*=\s*([\'"])(.*?)\1', lambda m: replacement, tag, flags=re.IGNORECASE)
    if updated != tag:
        return updated
    updated = re.sub(r"\bsrc\s*=\s*([^\s>]+)", lambda m: replacement, tag, flags=re.IGNORECASE)
    if updated != tag:
        return updated
    alt = _extract_img_attr(tag, "alt")
    if alt:
        return f'<img src="{new_src}" alt="{alt}">'
    return f'<img src="{new_src}">'


def sanitize_html_images(html_content: str, config: Config) -> str:
    """Process images in HTML, ensure they are usable."""
    img_pattern = re.compile(r"<img\b[^>]*>", flags=re.IGNORECASE)
    local_dir = config.image_local_dir

    def replace_img(match):
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
    """Check if image can be recognized by docx."""
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

    def replace_img(match):
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
    """Process image links in markdown, download to local and replace paths."""
    image_pattern = r"!\[([^\]]*)\]\(([^)]+)\)"

    def replace_image(match):
        alt_text = match.group(1)
        image_url = match.group(2)

        if image_url.startswith(("http://", "https://")):
            local_path = download_image(image_url, config)
            if local_path:
                return f"![{alt_text}]({local_path})"
            else:
                print_info(f"Image download failed, skipping: {image_url}")
                return alt_text or ""
        else:
            return match.group(0)

    return re.sub(image_pattern, replace_image, markdown_content)


def resize_images_in_document(document, max_width_inches: float = 6.0) -> None:
    """Resize all images in document to fit page width."""
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


# Style application functions
def apply_style_to_run(run, style_config: StyleConfig) -> None:
    """Apply style configuration to run."""
    run.font.name = style_config.font_name
    run.font.size = Pt(style_config.font_size)
    run.font.bold = style_config.bold
    run.font.italic = style_config.italic

    r, g, b = hex_to_rgb(style_config.color)
    run.font.color.rgb = RGBColor(r, g, b)

    # Set Chinese font
    if run._element.rPr is not None:
        rFonts = run._element.rPr.rFonts
        if rFonts is not None:
            rFonts.set(qn("w:eastAsia"), style_config.font_name)


def apply_style_to_paragraph(paragraph, style_config: StyleConfig) -> None:
    """Apply style configuration to paragraph."""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(style_config.space_before)
    pf.space_after = Pt(style_config.space_after)

    # Alignment
    alignment_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    if style_config.alignment in alignment_map:
        pf.alignment = alignment_map[style_config.alignment]

    # Line spacing
    if style_config.line_spacing_rule == "exact" and style_config.line_spacing_value:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(style_config.line_spacing_value)
    elif style_config.line_spacing_rule == "at_least" and style_config.line_spacing_value:
        pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        pf.line_spacing = Pt(style_config.line_spacing_value)
    elif style_config.line_spacing_rule == "single":
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    elif style_config.line_spacing_rule == "1.5":
        pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    elif style_config.line_spacing_rule == "double":
        pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    elif style_config.line_spacing_rule == "multiple":
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        if style_config.line_spacing_value:
            pf.line_spacing = style_config.line_spacing_value
        elif style_config.line_spacing > 0:
            pf.line_spacing = style_config.line_spacing
    elif style_config.line_spacing > 0:
        pf.line_spacing = style_config.line_spacing

    # Left indent
    if style_config.left_indent > 0:
        pf.left_indent = Inches(style_config.left_indent)

    # First line indent (in characters)
    if style_config.first_line_indent > 0:
        indent_pt = style_config.first_line_indent * style_config.font_size
        pf.first_line_indent = Pt(indent_pt)


def get_heading_level(paragraph) -> int | None:
    """Get heading level of paragraph, returns None if not a heading."""
    style_name = paragraph.style.name if paragraph.style else ""
    if style_name.startswith("Heading"):
        try:
            return int(style_name.replace("Heading ", "").replace("Heading", ""))
        except ValueError:
            return None
    return None


def apply_styles_to_document(document, config: Config) -> None:
    """Apply style configuration to document."""
    numbering = HeadingNumbering()

    for paragraph in document.paragraphs:
        heading_level = get_heading_level(paragraph)

        if heading_level is not None:
            style_name = f"heading_{heading_level}"
            style_config = config.get_style(style_name)

            # Add heading numbering
            if style_config.numbering_format and paragraph.runs:
                number_text = numbering.get_number(heading_level, style_config.numbering_format)
                if number_text:
                    first_run = paragraph.runs[0]
                    original_text = first_run.text
                    first_run.text = number_text + original_text
        else:
            style_config = config.get_style("body")

        apply_style_to_paragraph(paragraph, style_config)

        for run in paragraph.runs:
            apply_style_to_run(run, style_config)

    # Process tables
    for table in document.tables:
        for i, row in enumerate(table.rows):
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if i == 0:
                        style_config = config.get_style("table_header")
                    else:
                        style_config = config.get_style("table_cell")

                    apply_style_to_paragraph(paragraph, style_config)
                    for run in paragraph.runs:
                        apply_style_to_run(run, style_config)


def add_toc(document, title: str = "目录", max_level: int = 3) -> None:
    """Add table of contents at the beginning of document."""
    toc_title = document.paragraphs[0].insert_paragraph_before(title)
    toc_title.style = document.styles["Heading 1"]

    toc_paragraph = toc_title.insert_paragraph_before("")
    run = toc_paragraph.add_run()

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_char_begin)

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = f' TOC \\o "1-{max_level}" \\h \\z \\u '
    run._r.append(instr_text)

    fld_char_separate = OxmlElement("w:fldChar")
    fld_char_separate.set(qn("w:fldCharType"), "separate")
    run._r.append(fld_char_separate)

    placeholder_run = toc_paragraph.add_run("Right-click here and select 'Update Field' to generate TOC")
    placeholder_run.italic = True
    placeholder_run.font.color.rgb = RGBColor(128, 128, 128)

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char_end)

    page_break_paragraph = toc_title.insert_paragraph_before("")
    page_break_run = page_break_paragraph.add_run()
    page_break_run.add_break(WD_BREAK.PAGE)

    print_info(f"Added TOC (levels 1-{max_level})")


def convert(
    markdown_content: str,
    output_path: str | Path,
    config: Config | None = None,
    toc: bool = False,
    toc_title: str = "目录",
    toc_max_level: int = 3,
) -> Path:
    """
    Convert Markdown content to Word document.

    Args:
        markdown_content: Markdown text content
        output_path: Output file path
        config: Configuration object (uses defaults if None)
        toc: Whether to add table of contents
        toc_title: TOC title
        toc_max_level: Maximum heading level for TOC

    Returns:
        Path to the output file
    """
    if config is None:
        config = Config()

    output_path = Path(output_path)

    # Extract LaTeX formulas
    processed_content, formulas = extract_latex_formulas(markdown_content)
    if formulas:
        print_info(f"Detected {len(formulas)} LaTeX formulas")

    # Process markdown images
    processed_content = process_markdown_images(processed_content, config)

    # Convert to HTML
    html_content = markdown2.markdown(
        processed_content,
        extras=["tables", "cuddled-lists", "fenced-code-blocks", "header-ids"],
    )

    # Process HTML images
    html_content = sanitize_html_images(html_content, config)

    # Create Word document
    document = Document()
    new_parser = HtmlToDocx()

    try:
        new_parser.add_html_to_document(html_content, document)
    except UnrecognizedImageError as e:
        print_error(f"UnrecognizedImageError, retrying without problematic images: {e}")
        html_filtered = filter_unrecognized_images(html_content)
        document = Document()
        new_parser = HtmlToDocx()
        try:
            new_parser.add_html_to_document(html_filtered, document)
        except UnrecognizedImageError as e2:
            print_error(f"Still failing, removing all images: {e2}")
            html_without_images = re.sub(r"<img[^>]*>", "", html_filtered, flags=re.IGNORECASE)
            document = Document()
            new_parser = HtmlToDocx()
            new_parser.add_html_to_document(html_without_images, document)

    # Replace formula placeholders
    if formulas:
        replace_formula_placeholders(document, formulas)

    # Apply styles
    apply_styles_to_document(document, config)

    # Resize images
    resize_images_in_document(document, config.max_image_width_inches)

    # Add TOC
    if toc and len(document.paragraphs) > 0:
        add_toc(document, title=toc_title, max_level=toc_max_level)

    # Save document
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))
    print_info(f"Document saved: {output_path}")

    return output_path


def convert_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
    config: Config | str | Path | None = None,
    toc: bool = False,
    toc_title: str = "目录",
    toc_max_level: int = 3,
) -> Path:
    """
    Convert Markdown file to Word document.

    Args:
        input_path: Input Markdown file path
        output_path: Output file path (defaults to input with .docx extension)
        config: Configuration object or path to config file
        toc: Whether to add table of contents
        toc_title: TOC title
        toc_max_level: Maximum heading level for TOC

    Returns:
        Path to the output file
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        output_path = input_path.with_suffix(".docx")
    else:
        output_path = Path(output_path)

    # Load config
    if config is None:
        config = Config()
    elif isinstance(config, (str, Path)):
        config = Config.from_file(config)

    # Read markdown content
    markdown_content = input_path.read_text(encoding="utf-8")

    # Remove markdown code block wrapper if present
    if markdown_content.startswith("```markdown") and markdown_content.endswith("```"):
        markdown_content = markdown_content[12:-3]

    return convert(
        markdown_content,
        output_path,
        config,
        toc=toc,
        toc_title=toc_title,
        toc_max_level=toc_max_level,
    )
