# md2word Configuration Reference

This is the complete reference for md2word's JSON configuration file.
Generate a default config with `md2word --init-config`.
View effective config with `md2word --show-config` or `md2word --show-config -c your_config.json`.
Validate a config with `md2word --validate-config -c your_config.json`.

## Config Structure

```json
{
    "document": { ... },
    "styles": { ... },
    "table": { ... },
    "image": { ... }
}
```

## document

| Field | Type | Default | Description |
|---|---|---|---|
| `default_font` | string | `"仿宋"` | Default font for the entire document |
| `page_width_inches` | float | `8.5` | Page width in inches |
| `page_height_inches` | float | `11` | Page height in inches |
| `max_image_width_inches` | float | `6.0` | Max image width (auto-resize) |

## styles

Each key maps to a `StyleConfig` object. Supported keys:
`heading_1`, `heading_2`, `heading_3`, `heading_4`, `body`,
`code`, `blockquote`, `table_header`, `table_cell`.

### StyleConfig Fields

| Field | Type | Default | Description |
|---|---|---|---|
| `font_name` | string | null | Font name (e.g. "Arial", "黑体") |
| `font_size` | number/string | null | Points (12) or Chinese name ("四号") |
| `bold` | bool | false | Bold text |
| `italic` | bool | false | Italic text |
| `color` | string | null | Hex color without # (e.g. "FF0000") |
| `alignment` | string | null | "left", "center", "right", "justify" |
| `line_spacing_rule` | string | null | "single", "1.5", "double", "multiple", "exact", "at_least" |
| `line_spacing_value` | float | null | Multiplier (for "multiple") or points (for "exact"/"at_least") |
| `first_line_indent` | int | null | First line indent in characters (e.g. 2 for Chinese) |
| `left_indent` | float | null | Left indent in inches |
| `space_before` | float | null | Space before paragraph in points |
| `space_after` | float | null | Space after paragraph in points |
| `numbering_format` | string | null | Heading numbering format (see SKILL.md) |
| `background_color` | string | null | Background hex color (for code blocks) |

## table

| Field | Type | Default | Description |
|---|---|---|---|
| `border_style` | string | `"single"` | Border style |
| `border_color` | string | `"000000"` | Border color (hex) |
| `border_width` | int | `4` | Border width in eighth-points |
| `header_background_color` | string | null | Header row background |
| `cell_background_color` | string | null | Default cell background |
| `alternating_row_color` | string | null | Alternating row background |
| `cell_padding_top` | int | `2` | Cell top padding (points) |
| `cell_padding_bottom` | int | `2` | Cell bottom padding |
| `cell_padding_left` | int | `5` | Cell left padding |
| `cell_padding_right` | int | `5` | Cell right padding |
| `width_mode` | string | `"auto"` | "auto", "full", or "fixed" |
| `width_inches` | float | null | Fixed width (when width_mode="fixed") |

## image

| Field | Type | Default | Description |
|---|---|---|---|
| `local_dir` | string | `"./images"` | Directory to save downloaded images |
| `download_timeout` | int | `30` | HTTP download timeout in seconds |
| `user_agent` | string | (browser UA) | User-Agent for image downloads |

## Example: Chinese Official Document Config

```json
{
    "document": {
        "default_font": "仿宋",
        "page_width_inches": 8.5,
        "page_height_inches": 11,
        "max_image_width_inches": 6.0
    },
    "styles": {
        "heading_1": {
            "font_name": "黑体",
            "font_size": "三号",
            "bold": true,
            "alignment": "center",
            "line_spacing_rule": "exact",
            "line_spacing_value": 28,
            "numbering_format": "chapter"
        },
        "heading_2": {
            "font_name": "黑体",
            "font_size": "三号",
            "bold": true,
            "alignment": "left",
            "first_line_indent": 2,
            "numbering_format": "section"
        },
        "heading_3": {
            "font_name": "黑体",
            "font_size": "四号",
            "bold": true,
            "first_line_indent": 2,
            "numbering_format": "chinese"
        },
        "body": {
            "font_name": "仿宋",
            "font_size": "四号",
            "alignment": "justify",
            "line_spacing_rule": "exact",
            "line_spacing_value": 28,
            "first_line_indent": 2
        }
    }
}
```

## Example: English Academic Paper Config

```json
{
    "document": {
        "default_font": "Times New Roman"
    },
    "styles": {
        "heading_1": {
            "font_name": "Arial",
            "font_size": 16,
            "bold": true,
            "alignment": "left",
            "numbering_format": "arabic",
            "space_before": 24,
            "space_after": 12
        },
        "heading_2": {
            "font_name": "Arial",
            "font_size": 14,
            "bold": true,
            "numbering_format": "arabic",
            "space_before": 18,
            "space_after": 6
        },
        "body": {
            "font_name": "Times New Roman",
            "font_size": 12,
            "alignment": "justify",
            "line_spacing_rule": "double",
            "first_line_indent": 0
        },
        "code": {
            "font_name": "Consolas",
            "font_size": 10,
            "background_color": "f5f5f5"
        }
    }
}
```
