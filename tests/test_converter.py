"""Tests for md2word converter module."""

import os
import tempfile
from pathlib import Path

import pytest

from md2word import Config, convert, convert_file
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
