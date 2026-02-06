# md2word

[中文文档](README_zh.md) | English

Convert Markdown files to Word documents (.docx) with extensive customization options.

## Features

- Convert Markdown to Word documents
- Support for tables, code blocks, images, and more
- Automatic download and embedding of web images
- Automatic conversion of unsupported image formats (e.g., WebP)
- LaTeX formula support (converted to native Word equations)
- Configurable styles for headings, body text, and other elements
- Chinese font size support (e.g., "四号", "小四")
- Automatic heading numbering with multiple formats
- Optional table of contents generation

## Installation

### Using uv (recommended)

```bash
# Install globally
uv tool install md2word

# Or run directly without installation
uvx md2word input.md
```

### Using pip

```bash
pip install md2word
```

### Downloadable desktop app (no Python required)

If you prefer a GUI and do not want to install Python, you can download prebuilt executables from the GitHub Releases page.
Look for the desktop app artifacts for Windows, macOS, and Linux, then unzip and run.

## Usage

### Command Line

```bash
# Basic conversion (outputs input.docx)
md2word input.md

# Specify output file
md2word input.md -o output.docx

# Use custom config file
md2word input.md -c my_config.json

# Add table of contents
md2word input.md --toc

# Custom TOC title and level
md2word input.md --toc --toc-title "Contents" --toc-level 4

# Generate default config file
md2word --init-config
```

### As a Python Library

```python
import md2word

# Simple conversion
md2word.convert_file("input.md", "output.docx")

# With custom configuration
config = md2word.Config.from_file("config.json")
md2word.convert_file("input.md", "output.docx", config=config, toc=True)

# Convert from string
markdown_content = "# Hello World\n\nThis is a test."
md2word.convert(markdown_content, "output.docx")

# Programmatic configuration
config = md2word.Config()
config.default_font = "Arial"
config.styles["heading_1"] = md2word.StyleConfig(
    font_name="Arial",
    font_size=24,
    bold=True,
    alignment="center",
    numbering_format="chapter",
)
md2word.convert_file("input.md", "output.docx", config=config)
```

## Configuration

Create a `config.json` file to customize document styles.

### Example Configuration

```json
{
    "document": {
        "default_font": "Arial",
        "max_image_width_inches": 6.0
    },
    "styles": {
        "heading_1": {
            "font_name": "Arial",
            "font_size": 24,
            "bold": true,
            "alignment": "center",
            "line_spacing_rule": "exact",
            "line_spacing_value": 28,
            "numbering_format": "chapter"
        },
        "body": {
            "font_name": "Times New Roman",
            "font_size": 12,
            "alignment": "justify",
            "line_spacing_rule": "multiple",
            "line_spacing_value": 1.5,
            "first_line_indent": 2
        }
    }
}
```

### Style Properties

| Property | Type | Description |
|----------|------|-------------|
| `font_name` | string | Font name |
| `font_size` | number/string | Font size (points or Chinese size name) |
| `bold` | boolean | Bold text |
| `italic` | boolean | Italic text |
| `color` | string | Font color (hex, e.g., "000000") |
| `alignment` | string | Paragraph alignment: `left`/`center`/`right`/`justify` |
| `line_spacing_rule` | string | Line spacing mode (see below) |
| `line_spacing_value` | number | Line spacing value |
| `first_line_indent` | number | First line indent (in characters) |
| `left_indent` | float | Left indent (in inches) |
| `space_before` | number | Space before paragraph (points) |
| `space_after` | number | Space after paragraph (points) |
| `numbering_format` | string | Heading numbering format (see below) |

### Line Spacing Modes

| Value | Description |
|-------|-------------|
| `single` | Single line spacing |
| `1.5` | 1.5 line spacing |
| `double` | Double line spacing |
| `multiple` | Multiple line spacing (use `line_spacing_value` as multiplier) |
| `exact` | Exact line spacing (use `line_spacing_value` in points) |
| `at_least` | Minimum line spacing (use `line_spacing_value` in points) |

### Numbering Formats

| Format | Example |
|--------|---------|
| `chapter` | 第一章, 第二章, 第三章... |
| `section` | 第一节, 第二节, 第三节... |
| `chinese` | 一、二、三... |
| `chinese_paren` | （一）（二）（三）... |
| `arabic` | 1. 2. 3... |
| `arabic_paren` | (1) (2) (3)... |
| `arabic_bracket` | [1] [2] [3]... |
| `roman` | I. II. III... |
| `roman_lower` | i. ii. iii... |
| `letter` | A. B. C... |
| `letter_lower` | a. b. c... |
| `circle` | ① ② ③... |
| `none` | No numbering |

Custom format strings are also supported using `{n}` for Arabic numbers and `{cn}` for Chinese numbers.

### Table Configuration

Configure table appearance in the `table` section:

```json
{
    "table": {
        "border_style": "single",
        "border_color": "000000",
        "border_width": 4,
        "header_background_color": "D9E2F3",
        "cell_background_color": null,
        "alternating_row_color": "F2F2F2",
        "cell_padding_top": 2,
        "cell_padding_bottom": 2,
        "cell_padding_left": 5,
        "cell_padding_right": 5,
        "width_mode": "full",
        "width_inches": null
    }
}
```

| Property | Type | Description |
|----------|------|-------------|
| `border_style` | string | Border style: `single`/`double`/`dotted`/`dashed`/`none` |
| `border_color` | string | Border color (hex, e.g., "000000") |
| `border_width` | number | Border width in 1/8 points (4 = 0.5pt, 8 = 1pt) |
| `header_background_color` | string | Header row background color (hex) |
| `cell_background_color` | string | Cell background color (hex) |
| `alternating_row_color` | string | Alternating row color for zebra striping (hex) |
| `cell_padding_*` | number | Cell padding in points (top/bottom/left/right) |
| `width_mode` | string | Table width mode: `auto`/`full`/`fixed` |
| `width_inches` | number | Fixed width in inches (when `width_mode` is "fixed") |

### Chinese Font Sizes

| Name | Points | Name | Points |
|------|--------|------|--------|
| 初号 | 42 | 小初 | 36 |
| 一号 | 26 | 小一 | 24 |
| 二号 | 22 | 小二 | 18 |
| 三号 | 16 | 小三 | 15 |
| 四号 | 14 | 小四 | 12 |
| 五号 | 10.5 | 小五 | 9 |
| 六号 | 7.5 | 小六 | 6.5 |
| 七号 | 5.5 | 八号 | 5 |

## Requirements

- Python >= 3.10
- markdown2
- python-docx
- html-for-docx
- httpx
- Pillow
- latex2mathml
- mathml2omml
- lxml

## License

MIT
