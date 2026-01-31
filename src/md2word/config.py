"""Configuration classes for md2word."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

# Chinese font size mapping (font name -> point size)
CHINESE_FONT_SIZE_MAP: dict[str, float] = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}


def parse_font_size(value: int | float | str) -> float:
    """Parse font size, supporting both numeric (points) and Chinese font sizes."""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        value = value.strip()
        if value in CHINESE_FONT_SIZE_MAP:
            return CHINESE_FONT_SIZE_MAP[value]
        try:
            return float(value)
        except ValueError:
            pass
    print(f"[WARN] Unrecognized font size: {value}, using default 10.5 (五号)")
    return 10.5


@dataclass
class StyleConfig:
    """Style configuration for document elements."""

    font_name: str = "微软雅黑"
    font_size: float = 11
    bold: bool = False
    italic: bool = False
    color: str = "000000"
    space_before: int = 0
    space_after: int = 6
    line_spacing: float = 1.0
    left_indent: float = 0
    background_color: str | None = None
    # Extended configuration
    alignment: str = "left"  # left, center, right, justify
    line_spacing_rule: str = "multiple"  # single, 1.5, double, multiple, exact, at_least
    line_spacing_value: float | None = None  # Line spacing value (points or multiple)
    first_line_indent: float = 0  # First line indent (in characters)
    is_heading: bool = True  # Whether to treat as heading (for TOC)
    numbering_format: str | None = None  # Numbering format

    @classmethod
    def from_dict(cls, data: dict[str, Any], default_font: str = "微软雅黑") -> StyleConfig:
        """Create StyleConfig from dictionary."""
        return cls(
            font_name=data.get("font_name", default_font),
            font_size=parse_font_size(data.get("font_size", 11)),
            bold=data.get("bold", False),
            italic=data.get("italic", False),
            color=data.get("color", "000000"),
            space_before=data.get("space_before", 0),
            space_after=data.get("space_after", 6),
            line_spacing=data.get("line_spacing", 1.0),
            left_indent=data.get("left_indent", 0),
            background_color=data.get("background_color"),
            alignment=data.get("alignment", "left"),
            line_spacing_rule=data.get("line_spacing_rule", "multiple"),
            line_spacing_value=data.get("line_spacing_value"),
            first_line_indent=data.get("first_line_indent", 0),
            is_heading=data.get("is_heading", True),
            numbering_format=data.get("numbering_format"),
        )

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary."""
        return {
            "font_name": self.font_name,
            "font_size": self.font_size,
            "bold": self.bold,
            "italic": self.italic,
            "color": self.color,
            "space_before": self.space_before,
            "space_after": self.space_after,
            "line_spacing": self.line_spacing,
            "left_indent": self.left_indent,
            "background_color": self.background_color,
            "alignment": self.alignment,
            "line_spacing_rule": self.line_spacing_rule,
            "line_spacing_value": self.line_spacing_value,
            "first_line_indent": self.first_line_indent,
            "is_heading": self.is_heading,
            "numbering_format": self.numbering_format,
        }


@dataclass
class TableConfig:
    """Table style configuration."""

    # Border settings
    border_style: str = "single"  # single, double, dotted, dashed, none
    border_color: str = "000000"
    border_width: int = 4  # in eighths of a point (4 = 0.5pt, 8 = 1pt)

    # Background colors
    header_background_color: str | None = None  # e.g., "D9E2F3" for light blue
    cell_background_color: str | None = None
    alternating_row_color: str | None = None  # For zebra striping

    # Cell padding (in points)
    cell_padding_top: float = 2
    cell_padding_bottom: float = 2
    cell_padding_left: float = 5
    cell_padding_right: float = 5

    # Table width
    width_mode: str = "auto"  # auto, full (100% page width), fixed
    width_inches: float | None = None  # Used when width_mode is "fixed"

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> TableConfig:
        """Create TableConfig from dictionary."""
        return cls(
            border_style=data.get("border_style", "single"),
            border_color=data.get("border_color", "000000"),
            border_width=data.get("border_width", 4),
            header_background_color=data.get("header_background_color"),
            cell_background_color=data.get("cell_background_color"),
            alternating_row_color=data.get("alternating_row_color"),
            cell_padding_top=data.get("cell_padding_top", 2),
            cell_padding_bottom=data.get("cell_padding_bottom", 2),
            cell_padding_left=data.get("cell_padding_left", 5),
            cell_padding_right=data.get("cell_padding_right", 5),
            width_mode=data.get("width_mode", "auto"),
            width_inches=data.get("width_inches"),
        )

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary."""
        return {
            "border_style": self.border_style,
            "border_color": self.border_color,
            "border_width": self.border_width,
            "header_background_color": self.header_background_color,
            "cell_background_color": self.cell_background_color,
            "alternating_row_color": self.alternating_row_color,
            "cell_padding_top": self.cell_padding_top,
            "cell_padding_bottom": self.cell_padding_bottom,
            "cell_padding_left": self.cell_padding_left,
            "cell_padding_right": self.cell_padding_right,
            "width_mode": self.width_mode,
            "width_inches": self.width_inches,
        }


@dataclass
class Config:
    """Global configuration for md2word converter."""

    default_font: str = "微软雅黑"
    page_width_inches: float = 8.5
    page_height_inches: float = 11
    max_image_width_inches: float = 6.0
    image_local_dir: str = "./images"
    image_download_timeout: int = 30
    image_user_agent: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    styles: dict[str, StyleConfig] = field(default_factory=dict)
    table: TableConfig = field(default_factory=TableConfig)

    @classmethod
    def from_file(cls, config_path: str | Path) -> Config:
        """Load configuration from JSON file."""
        config_path = Path(config_path)
        if not config_path.exists():
            print(f"Config file not found: {config_path}, using defaults")
            return cls()

        with open(config_path, encoding="utf-8") as f:
            data = json.load(f)

        return cls.from_dict(data)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Config:
        """Create Config from dictionary."""
        config = cls()

        # Document configuration
        doc_config = data.get("document", {})
        config.default_font = doc_config.get("default_font", config.default_font)
        config.page_width_inches = doc_config.get("page_width_inches", config.page_width_inches)
        config.page_height_inches = doc_config.get("page_height_inches", config.page_height_inches)
        config.max_image_width_inches = doc_config.get("max_image_width_inches", config.max_image_width_inches)

        # Image configuration
        img_config = data.get("image", {})
        config.image_local_dir = img_config.get("local_dir", config.image_local_dir)
        config.image_download_timeout = img_config.get("download_timeout", config.image_download_timeout)
        config.image_user_agent = img_config.get("user_agent", config.image_user_agent)

        # Style configuration
        styles_data = data.get("styles", {})
        for style_name, style_config in styles_data.items():
            config.styles[style_name] = StyleConfig.from_dict(style_config, config.default_font)

        # Table configuration
        table_data = data.get("table", {})
        if table_data:
            config.table = TableConfig.from_dict(table_data)

        return config

    def get_style(self, style_name: str) -> StyleConfig:
        """Get style configuration by name, returns default if not found."""
        return self.styles.get(style_name, StyleConfig(font_name=self.default_font))

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary."""
        return {
            "document": {
                "default_font": self.default_font,
                "page_width_inches": self.page_width_inches,
                "page_height_inches": self.page_height_inches,
                "max_image_width_inches": self.max_image_width_inches,
            },
            "image": {
                "local_dir": self.image_local_dir,
                "download_timeout": self.image_download_timeout,
                "user_agent": self.image_user_agent,
            },
            "styles": {name: style.to_dict() for name, style in self.styles.items()},
            "table": self.table.to_dict(),
        }

    def save(self, path: str | Path) -> None:
        """Save configuration to JSON file."""
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=4)


# Default configuration template
DEFAULT_CONFIG = {
    "document": {
        "default_font": "仿宋",
        "page_width_inches": 8.5,
        "page_height_inches": 11,
        "max_image_width_inches": 6.0,
    },
    "styles": {
        "heading_1": {
            "font_name": "黑体",
            "font_size": "三号",
            "bold": True,
            "alignment": "center",
            "line_spacing_rule": "exact",
            "line_spacing_value": 28,
            "first_line_indent": 0,
            "space_before": 24,
            "space_after": 12,
            "numbering_format": "chapter",
        },
        "heading_2": {
            "font_name": "黑体",
            "font_size": "三号",
            "bold": True,
            "alignment": "left",
            "line_spacing_rule": "exact",
            "line_spacing_value": 28,
            "first_line_indent": 2,
            "space_before": 12,
            "space_after": 6,
            "numbering_format": "section",
        },
        "heading_3": {
            "font_name": "黑体",
            "font_size": "三号",
            "bold": True,
            "alignment": "center",
            "line_spacing_rule": "exact",
            "line_spacing_value": 28,
            "first_line_indent": 0,
            "space_before": 6,
            "space_after": 6,
            "numbering_format": "chinese",
        },
        "body": {
            "font_name": "仿宋",
            "font_size": 11,
            "alignment": "justify",
            "line_spacing_rule": "multiple",
            "line_spacing_value": 1.5,
            "first_line_indent": 2,
            "space_before": 0,
            "space_after": 10,
        },
        "code": {
            "font_name": "Consolas",
            "font_size": 10,
            "alignment": "left",
            "line_spacing_rule": "single",
            "first_line_indent": 0,
            "background_color": "f5f5f5",
        },
        "blockquote": {
            "font_name": "仿宋",
            "font_size": 11,
            "italic": True,
            "color": "666666",
            "alignment": "left",
            "line_spacing_rule": "multiple",
            "line_spacing_value": 1.5,
            "left_indent": 0.5,
            "first_line_indent": 0,
        },
        "table_header": {
            "font_name": "仿宋",
            "font_size": 11,
            "bold": True,
            "alignment": "center",
            "line_spacing_rule": "single",
        },
        "table_cell": {
            "font_name": "仿宋",
            "font_size": 11,
            "alignment": "left",
            "line_spacing_rule": "single",
        },
    },
    "image": {
        "local_dir": "./images",
        "download_timeout": 30,
        "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    },
}
