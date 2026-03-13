---
name: md2word
description: >
  Convert Markdown text or files to professionally styled Word (.docx) documents
  using the md2word Python library. Use this skill whenever the user needs to:
  (1) convert Markdown content to Word format,
  (2) generate .docx files from .md files or Markdown strings,
  (3) create Word documents with custom styling (fonts, colors, spacing, numbering),
  (4) produce Chinese-style official documents (公文) from Markdown,
  (5) convert documents containing LaTeX formulas, tables, code blocks, or images to Word,
  (6) customize Word output with specific heading numbering formats (第一章, 一、, ①, etc.),
  (7) generate table of contents in Word documents.
  Even if the user doesn't mention "md2word" explicitly, trigger this skill when they
  want Markdown-to-Word conversion with any level of formatting control.
---

# md2word — Markdown to Word Converter

## What md2word Does

md2word converts Markdown to Word documents with fine-grained control over every
aspect of the output styling. It handles the full spectrum of Markdown syntax —
tables, code blocks, LaTeX math, images, nested lists — and produces clean .docx
files with configurable fonts, colors, spacing, numbering, and layout.

The library is especially strong for Chinese typography and academic/official
document formatting, but works equally well for English documents.

## Installation

```bash
pip install md2word
# or
uvx md2word
```

## Two Ways to Use md2word

### 1. CLI (quick conversions)

```bash
md2word input.md                                    # → input.docx
md2word input.md -o report.docx                     # custom output path
md2word input.md -c config.json                     # custom styling config
md2word input.md --toc --toc-title "Contents"       # with table of contents
md2word --init-config                               # generate default config.json (includes table block)
md2word --show-config                               # print effective config as JSON
md2word --show-config -c my.json                    # print merged config from file
md2word --list-formats                              # list heading numbering formats
md2word --validate-config -c my.json                # validate config file
echo "# Hello" | md2word - -o output.docx           # read from stdin
```

### 2. Python API (programmatic control)

```python
import md2word

# Simple conversion
md2word.convert_file("input.md", "output.docx")

# From string
md2word.convert("# Hello\n\nSome text.", "output.docx")

# With config
config = md2word.Config.from_file("config.json")
md2word.convert_file("input.md", "output.docx", config=config, toc=True)
```

## Supported Markdown Features

md2word handles all common Markdown elements:

- **Text**: bold, italic, bold-italic, inline code
- **Headings**: H1–H6 with 12+ automatic numbering formats
- **Lists**: ordered and unordered, with nesting
- **Tables**: pipe-delimited with headers, configurable borders/colors/padding
- **Code blocks**: fenced with language tag, background color, monospace font
- **Blockquotes**: with left border styling
- **Images**: local files, HTTP/HTTPS URLs, data URIs, auto-resize to page width
- **LaTeX formulas**: inline `$...$` and block `$$...$$`, converted to native Word equations (OMML)
- **Links**: preserved as hyperlinks

## Configuration System

md2word uses a JSON config file to control every aspect of Word output.
For the full configuration reference with all available options, read
`references/config_reference.md` in this skill directory.

### Quick Config Overview

The config has four sections:

1. **document** — page size, default font, max image width
2. **styles** — per-element styling (heading_1–4, body, code, blockquote, table_header, table_cell)
3. **table** — borders, colors, padding, width mode
4. **image** — download settings for remote images

Each style element supports: `font_name`, `font_size` (points or Chinese size like "四号"),
`bold`, `italic`, `color`, `alignment`, `line_spacing_rule`, `line_spacing_value`,
`first_line_indent`, `left_indent`, `space_before`, `space_after`, `numbering_format`,
`background_color`.

### Heading Numbering Formats

| Format | Example Output |
|---|---|
| `chapter` | 第一章, 第二章, 第三章 |
| `section` | 第一节, 第二节, 第三节 |
| `chinese` | 一、二、三 |
| `chinese_paren` | （一）（二）（三） |
| `arabic` | 1. 2. 3. |
| `arabic_paren` | (1) (2) (3) |
| `arabic_bracket` | [1] [2] [3] |
| `roman` | I. II. III. |
| `roman_lower` | i. ii. iii. |
| `letter` | A. B. C. |
| `letter_lower` | a. b. c. |
| `circle` | ① ② ③ |
| `none` | No numbering |
| Custom string | Use `{n}` for Arabic, `{cn}` for Chinese numbers |

### Chinese Font Sizes

md2word supports traditional Chinese font size names:

| Name | Points | Name | Points |
|------|--------|------|--------|
| 初号 | 42 | 小初 | 36 |
| 一号 | 26 | 小一 | 24 |
| 二号 | 22 | 小二 | 18 |
| 三号 | 16 | 小三 | 15 |
| 四号 | 14 | 小四 | 12 |
| 五号 | 10.5 | 小五 | 9 |

## Workflow Guide

### Basic: Convert a Markdown file

```bash
md2word input.md -o output.docx
```

### With custom styling

1. Generate a default config: `md2word --init-config`
2. Edit `config.json` to adjust fonts, sizes, colors, spacing
3. Convert: `md2word input.md -c config.json -o output.docx`

### Chinese official document (公文) style

Use the Chinese config template from `examples/config_chinese.json` in the project.
It sets up 仿宋 body text, 黑体 headings, 第一章/第一节/一、 numbering hierarchy,
exact 28pt line spacing, and 2-character first-line indent.

### With table of contents

```bash
md2word input.md --toc --toc-title "目录" --toc-level 3
```

### Programmatic with custom config

```python
import md2word

config = md2word.Config()
config.default_font = "Arial"
config.styles["heading_1"] = md2word.StyleConfig(
    font_name="Arial", font_size=24, bold=True,
    alignment="center", numbering_format="chapter"
)
config.styles["body"] = md2word.StyleConfig(
    font_name="Times New Roman", font_size=12,
    alignment="justify", line_spacing_rule="multiple",
    line_spacing_value=1.5, first_line_indent=2
)
md2word.convert_file("input.md", "output.docx", config=config, toc=True)
```

## Key API Reference

### Functions

- `md2word.convert(markdown_content, output_path, config=None, toc=False, toc_title="目录", toc_max_level=3)` — convert Markdown string
- `md2word.convert_file(input_path, output_path=None, config=None, toc=False, toc_title="目录", toc_max_level=3)` — convert Markdown file
- `md2word.extract_latex_formulas(text)` — extract LaTeX formulas from text
- `md2word.latex_to_omml(latex)` — convert LaTeX to Word OMML format

### Classes

- `md2word.Config` — main configuration, load with `Config.from_file("config.json")`
- `md2word.StyleConfig` — per-element style settings
- `md2word.CHINESE_FONT_SIZE_MAP` — dict mapping Chinese size names to points
- `md2word.DEFAULT_CONFIG` — default configuration dictionary

## Troubleshooting

- If `md2word` command not found: use `uvx md2word` or `python -m md2word`
- LaTeX conversion requires `latex2mathml` and `mathml2omml` packages (included as dependencies)
- Remote images need network access; set `image.download_timeout` in config if slow
- WebP images are auto-converted to PNG/JPEG for Word compatibility
