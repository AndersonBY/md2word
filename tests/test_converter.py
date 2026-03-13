"""Tests for md2word converter module."""

import json
import os
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

from md2word import Config, convert, convert_file
from md2word.__main__ import main
from md2word.converter import (
    HeadingNumbering,
    hex_to_rgb,
    number_to_chinese,
)


class TestHexToRgb:
    """Tests for hex_to_rgb function."""

    def test_black(self):
        """Test black color."""
        assert hex_to_rgb("000000") == (0, 0, 0)

    def test_white(self):
        """Test white color."""
        assert hex_to_rgb("FFFFFF") == (255, 255, 255)

    def test_red(self):
        """Test red color."""
        assert hex_to_rgb("FF0000") == (255, 0, 0)

    def test_with_hash(self):
        """Test color with hash prefix."""
        assert hex_to_rgb("#FF0000") == (255, 0, 0)


class TestNumberToChinese:
    """Tests for number_to_chinese function."""

    def test_basic_numbers(self):
        """Test basic number conversion."""
        assert number_to_chinese(1) == "一"
        assert number_to_chinese(2) == "二"
        assert number_to_chinese(10) == "十"

    def test_teens(self):
        """Test teen numbers."""
        assert number_to_chinese(11) == "十一"
        assert number_to_chinese(15) == "十五"

    def test_twenty(self):
        """Test twenty."""
        assert number_to_chinese(20) == "二十"

    def test_beyond_twenty(self):
        """Test numbers beyond 20 return string."""
        assert number_to_chinese(21) == "21"
        assert number_to_chinese(100) == "100"


class TestHeadingNumbering:
    """Tests for HeadingNumbering class."""

    def test_chapter_format(self):
        """Test chapter numbering format."""
        numbering = HeadingNumbering()
        assert numbering.get_number(1, "chapter") == "第一章"
        assert numbering.get_number(1, "chapter") == "第二章"
        assert numbering.get_number(1, "chapter") == "第三章"

    def test_section_format(self):
        """Test section numbering format."""
        numbering = HeadingNumbering()
        assert numbering.get_number(2, "section") == "第一节"
        assert numbering.get_number(2, "section") == "第二节"

    def test_chinese_format(self):
        """Test Chinese numbering format."""
        numbering = HeadingNumbering()
        assert numbering.get_number(3, "chinese") == "一、"
        assert numbering.get_number(3, "chinese") == "二、"

    def test_arabic_format(self):
        """Test Arabic numbering format."""
        numbering = HeadingNumbering()
        assert numbering.get_number(4, "arabic") == "1."
        assert numbering.get_number(4, "arabic") == "2."

    def test_arabic_paren_format(self):
        """Test Arabic parenthesis numbering format."""
        numbering = HeadingNumbering()
        assert numbering.get_number(5, "arabic_paren") == "(1)"
        assert numbering.get_number(5, "arabic_paren") == "(2)"

    def test_circle_format(self):
        """Test circle numbering format."""
        numbering = HeadingNumbering()
        assert numbering.get_number(6, "circle") == "①"
        assert numbering.get_number(6, "circle") == "②"

    def test_roman_format(self):
        """Test Roman numeral format."""
        numbering = HeadingNumbering()
        assert numbering.get_number(1, "roman") == "I."
        assert numbering.get_number(1, "roman") == "II."
        assert numbering.get_number(1, "roman") == "III."

    def test_letter_format(self):
        """Test letter format."""
        numbering = HeadingNumbering()
        assert numbering.get_number(1, "letter") == "A."
        assert numbering.get_number(1, "letter") == "B."

    def test_none_format(self):
        """Test none format returns empty string."""
        numbering = HeadingNumbering()
        assert numbering.get_number(1, "none") == ""
        assert numbering.get_number(1, None) == ""

    def test_level_reset(self):
        """Test that sub-levels reset when parent level increments."""
        numbering = HeadingNumbering()
        numbering.get_number(1, "arabic")  # 1.
        numbering.get_number(2, "arabic")  # 1.
        numbering.get_number(2, "arabic")  # 2.
        numbering.get_number(1, "arabic")  # 2. (level 1 increments)
        assert numbering.get_number(2, "arabic") == "1."  # level 2 resets

    def test_custom_format(self):
        """Test custom format string."""
        numbering = HeadingNumbering()
        assert numbering.get_number(1, "Part {n}") == "Part 1"
        assert numbering.get_number(1, "第{cn}部分") == "第二部分"


class TestConvert:
    """Tests for convert function."""

    def test_simple_conversion(self):
        """Test simple markdown conversion."""
        markdown = "# Hello World\n\nThis is a test."
        temp_dir = tempfile.mkdtemp()
        output_path = Path(temp_dir) / "output.docx"
        try:
            result = convert(markdown, output_path)
            assert result.exists()
            assert result.suffix == ".docx"
        finally:
            output_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)

    def test_with_config(self):
        """Test conversion with custom config."""
        markdown = "# Test\n\nContent"
        config = Config()
        config.default_font = "Arial"
        temp_dir = tempfile.mkdtemp()
        output_path = Path(temp_dir) / "output.docx"
        try:
            result = convert(markdown, output_path, config=config)
            assert result.exists()
        finally:
            output_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)

    def test_with_toc(self):
        """Test conversion with table of contents."""
        markdown = "# Chapter 1\n\n## Section 1.1\n\nContent"
        temp_dir = tempfile.mkdtemp()
        output_path = Path(temp_dir) / "output.docx"
        try:
            result = convert(markdown, output_path, toc=True)
            assert result.exists()
        finally:
            output_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)

    def test_with_table(self):
        """Test conversion with table."""
        markdown = """
# Test

| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |
"""
        temp_dir = tempfile.mkdtemp()
        output_path = Path(temp_dir) / "output.docx"
        try:
            result = convert(markdown, output_path)
            assert result.exists()
        finally:
            output_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)

    def test_with_code_block(self):
        """Test conversion with code block."""
        markdown = """
# Test

```python
def hello():
    print("Hello")
```
"""
        temp_dir = tempfile.mkdtemp()
        output_path = Path(temp_dir) / "output.docx"
        try:
            result = convert(markdown, output_path)
            assert result.exists()
        finally:
            output_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)


class TestConvertFile:
    """Tests for convert_file function."""

    def test_file_conversion(self):
        """Test file-based conversion."""
        temp_dir = tempfile.mkdtemp()
        md_path = Path(temp_dir) / "input.md"
        docx_path = Path(temp_dir) / "input.docx"
        try:
            md_path.write_text("# Test\n\nHello World", encoding="utf-8")
            result = convert_file(md_path)
            assert result.exists()
            assert result == docx_path
        finally:
            md_path.unlink(missing_ok=True)
            docx_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)

    def test_custom_output_path(self):
        """Test conversion with custom output path."""
        temp_dir = tempfile.mkdtemp()
        md_path = Path(temp_dir) / "input.md"
        docx_path = Path(temp_dir) / "custom_output.docx"
        try:
            md_path.write_text("# Test", encoding="utf-8")
            result = convert_file(md_path, docx_path)
            assert result.exists()
            assert result == docx_path
        finally:
            md_path.unlink(missing_ok=True)
            docx_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)

    def test_nonexistent_file_raises(self):
        """Test that non-existent file raises error."""
        with pytest.raises(FileNotFoundError):
            convert_file("/nonexistent/file.md")


class TestCLI:
    """Tests for CLI entry point."""

    def test_version(self, capsys):
        """Test --version flag."""
        with patch("sys.argv", ["md2word", "-v"]):
            assert main() == 0
        assert "md2word" in capsys.readouterr().out

    def test_show_config_default(self, capsys):
        """Test --show-config with default config."""
        with patch("sys.argv", ["md2word", "--show-config"]):
            assert main() == 0
        output = capsys.readouterr().out
        data = json.loads(output)
        assert "document" in data
        assert "styles" in data
        assert "table" in data

    def test_show_config_with_file(self, capsys, tmp_path):
        """Test --show-config with a config file."""
        cfg = tmp_path / "cfg.json"
        cfg.write_text(json.dumps({
            "document": {"default_font": "Arial"},
        }), encoding="utf-8")
        with patch("sys.argv", ["md2word", "--show-config", "-c", str(cfg)]):
            assert main() == 0
        data = json.loads(capsys.readouterr().out)
        assert data["document"]["default_font"] == "Arial"

    def test_list_formats(self, capsys):
        """Test --list-formats flag."""
        with patch("sys.argv", ["md2word", "--list-formats"]):
            assert main() == 0
        output = capsys.readouterr().out
        assert "chapter" in output
        assert "section" in output
        assert "chinese" in output
        assert "第一章" in output

    def test_init_config_creates_file(self, capsys, tmp_path):
        """Test --init-config creates config with table block."""
        cfg = tmp_path / "config.json"
        with patch("sys.argv", ["md2word", "--init-config", "-c", str(cfg)]):
            assert main() == 0
        assert cfg.exists()
        data = json.loads(cfg.read_text(encoding="utf-8"))
        assert "table" in data
        assert "border_style" in data["table"]

    def test_init_config_refuses_overwrite(self, capsys, tmp_path):
        """Test --init-config refuses to overwrite existing file."""
        cfg = tmp_path / "config.json"
        cfg.write_text("{}", encoding="utf-8")
        with patch("sys.argv", ["md2word", "--init-config", "-c", str(cfg)]):
            assert main() == 1
        assert "already exists" in capsys.readouterr().out

    def test_config_not_found_explicit_warns(self, capsys, tmp_path):
        """Test explicit -c with nonexistent file prints warning."""
        md = tmp_path / "test.md"
        md.write_text("# Hello", encoding="utf-8")
        with patch("sys.argv", ["md2word", str(md), "-c", "nonexistent.json"]):
            assert main() == 0
        output = capsys.readouterr().out
        assert "[WARN]" in output
        assert "nonexistent.json" in output

    def test_config_not_found_implicit_silent(self, capsys, tmp_path):
        """Test no -c flag uses default config silently."""
        md = tmp_path / "test.md"
        md.write_text("# Hello", encoding="utf-8")
        with patch("sys.argv", ["md2word", str(md), "-o", str(tmp_path / "out.docx")]):
            assert main() == 0
        output = capsys.readouterr().out
        assert "[WARN]" not in output
        assert "Using default config" in output

    def test_validate_config_valid(self, capsys, tmp_path):
        """Test --validate-config with valid config."""
        cfg = tmp_path / "cfg.json"
        cfg.write_text(json.dumps({
            "styles": {
                "heading_1": {"font_name": "黑体", "bold": True},
            },
        }), encoding="utf-8")
        with patch("sys.argv", ["md2word", "--validate-config", "-c", str(cfg)]):
            assert main() == 0
        assert "[OK]" in capsys.readouterr().out

    def test_validate_config_invalid(self, capsys, tmp_path):
        """Test --validate-config with invalid config."""
        cfg = tmp_path / "cfg.json"
        cfg.write_text(json.dumps({
            "styles": {
                "heading_99": {"alignment": "middle"},
            },
        }), encoding="utf-8")
        with patch("sys.argv", ["md2word", "--validate-config", "-c", str(cfg)]):
            assert main() == 1
        output = capsys.readouterr().out
        assert "heading_99" in output
        assert "middle" in output

    def test_stdin_mode(self, capsys, tmp_path):
        """Test reading from stdin with '-'."""
        out = tmp_path / "output.docx"
        with patch("sys.argv", ["md2word", "-", "-o", str(out)]):
            with patch("sys.stdin") as mock_stdin:
                mock_stdin.read.return_value = "# Hello\n\nWorld"
                assert main() == 0
        assert out.exists()

    def test_stdin_requires_output(self, capsys):
        """Test stdin mode requires -o flag."""
        with patch("sys.argv", ["md2word", "-"]):
            assert main() == 1
        assert "-o/--output is required" in capsys.readouterr().out

    def test_no_input_shows_help(self, capsys):
        """Test no arguments shows help."""
        with patch("sys.argv", ["md2word"]):
            assert main() == 1

    def test_config_source_shown(self, capsys, tmp_path):
        """Test config source is printed during conversion."""
        md = tmp_path / "test.md"
        md.write_text("# Hello", encoding="utf-8")
        cfg = tmp_path / "my.json"
        cfg.write_text("{}", encoding="utf-8")
        with patch("sys.argv", ["md2word", str(md), "-c", str(cfg)]):
            assert main() == 0
        assert "Using config:" in capsys.readouterr().out
