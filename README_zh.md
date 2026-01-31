# md2word

将 Markdown 文件转换为 Word 文档（.docx），支持丰富的自定义配置选项。

## 功能特性

- Markdown 转 Word 文档
- 支持表格、代码块、图片等常见元素
- 自动下载网络图片并嵌入文档
- 自动转换不支持的图片格式（如 WebP）
- LaTeX 公式支持（转换为 Word 原生公式）
- 可配置各级标题和正文的字体、字号、加粗、斜体等样式
- 支持中文字号（如"四号"、"小四"）
- 自动标题编号，支持多种格式
- 可选生成目录

## 安装

### 使用 uv（推荐）

```bash
# 全局安装
uv tool install md2word

# 或直接运行无需安装
uvx md2word input.md
```

### 使用 pip

```bash
pip install md2word
```

## 使用方法

### 命令行

```bash
# 基本转换（输出 input.docx）
md2word input.md

# 指定输出文件
md2word input.md -o output.docx

# 使用自定义配置文件
md2word input.md -c my_config.json

# 添加目录
md2word input.md --toc

# 自定义目录标题和层级
md2word input.md --toc --toc-title "目 录" --toc-level 4

# 生成默认配置文件
md2word --init-config
```

### 作为 Python 库使用

```python
import md2word

# 简单转换
md2word.convert_file("input.md", "output.docx")

# 使用自定义配置
config = md2word.Config.from_file("config.json")
md2word.convert_file("input.md", "output.docx", config=config, toc=True)

# 从字符串转换
markdown_content = "# 你好世界\n\n这是一个测试。"
md2word.convert(markdown_content, "output.docx")

# 编程式配置
config = md2word.Config()
config.default_font = "仿宋"
config.styles["heading_1"] = md2word.StyleConfig(
    font_name="黑体",
    font_size=16,  # 三号
    bold=True,
    alignment="center",
    numbering_format="chapter",
)
md2word.convert_file("input.md", "output.docx", config=config)
```

## 配置文件

创建 `config.json` 文件来自定义文档样式。

### 配置示例

```json
{
    "document": {
        "default_font": "仿宋",
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
        "body": {
            "font_name": "仿宋",
            "font_size": 11,
            "alignment": "justify",
            "line_spacing_rule": "multiple",
            "line_spacing_value": 1.5,
            "first_line_indent": 2
        }
    }
}
```

### 样式属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `font_name` | string | 字体名称 |
| `font_size` | number/string | 字号（磅值或中文字号） |
| `bold` | boolean | 是否加粗 |
| `italic` | boolean | 是否斜体 |
| `color` | string | 字体颜色（十六进制，如 "000000"） |
| `alignment` | string | 对齐方式：`left`/`center`/`right`/`justify` |
| `line_spacing_rule` | string | 行距模式（见下表） |
| `line_spacing_value` | number | 行距数值 |
| `first_line_indent` | number | 首行缩进（字符数） |
| `left_indent` | float | 左缩进（英寸） |
| `space_before` | number | 段前间距（磅） |
| `space_after` | number | 段后间距（磅） |
| `numbering_format` | string | 标题编号格式（见下表） |

### 行距模式（line_spacing_rule）

| 值 | 说明 |
|------|------|
| `single` | 单倍行距 |
| `1.5` | 1.5倍行距 |
| `double` | 双倍行距 |
| `multiple` | 多倍行距（使用 `line_spacing_value` 作为倍数） |
| `exact` | 固定值（使用 `line_spacing_value` 作为磅值） |
| `at_least` | 最小值（使用 `line_spacing_value` 作为最小磅值） |

### 序号格式（numbering_format）

| 格式 | 效果示例 |
|------|------|
| `chapter` | 第一章、第二章、第三章... |
| `section` | 第一节、第二节、第三节... |
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
| `none` | 无编号 |

也支持自定义格式字符串，使用 `{n}` 表示阿拉伯数字，`{cn}` 表示中文数字。例如：`"第{cn}部分"` 会生成 "第一部分"、"第二部分"...

### 中文字号对照表

| 字号 | 磅值 | 字号 | 磅值 |
|------|------|------|------|
| 初号 | 42 | 小初 | 36 |
| 一号 | 26 | 小一 | 24 |
| 二号 | 22 | 小二 | 18 |
| 三号 | 16 | 小三 | 15 |
| 四号 | 14 | 小四 | 12 |
| 五号 | 10.5 | 小五 | 9 |
| 六号 | 7.5 | 小六 | 6.5 |
| 七号 | 5.5 | 八号 | 5 |

## 依赖

- Python >= 3.10
- markdown2
- python-docx
- html-for-docx
- httpx
- Pillow
- latex2mathml
- mathml2omml
- lxml

## 许可证

MIT
