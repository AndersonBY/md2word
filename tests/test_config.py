"""Tests for md2word configuration module."""

import json
import os
import tempfile
from pathlib import Path

from md2word.config import (
    CHINESE_FONT_SIZE_MAP,
    DEFAULT_CONFIG,
    Config,
    StyleConfig,
    parse_font_size,
)


class TestParseFontSize:
    """Tests for parse_font_size function."""

    def test_numeric_int(self):
        """Test parsing integer font size."""
        assert parse_font_size(12) == 12.0

    def test_numeric_float(self):
        """Test parsing float font size."""
        assert parse_font_size(10.5) == 10.5

    def test_numeric_string(self):
        """Test parsing numeric string font size."""
        assert parse_font_size("14") == 14.0
        assert parse_font_size("10.5") == 10.5

    def test_chinese_font_size(self):
        """Test parsing Chinese font size names."""
        assert parse_font_size("四号") == 14
        assert parse_font_size("小四") == 12
        assert parse_font_size("三号") == 16
        assert parse_font_size("五号") == 10.5

    def test_all_chinese_sizes(self):
        """Test all Chinese font sizes are parseable."""
        for name, size in CHINESE_FONT_SIZE_MAP.items():
            assert parse_font_size(name) == size

    def test_invalid_returns_default(self):
        """Test invalid font size returns default."""
        assert parse_font_size("invalid") == 10.5
        assert parse_font_size("") == 10.5


class TestStyleConfig:
    """Tests for StyleConfig class."""

    def test_default_values(self):
        """Test default values are set correctly."""
        style = StyleConfig()
        assert style.font_name == "微软雅黑"
        assert style.font_size == 11
        assert style.bold is False
        assert style.italic is False
        assert style.alignment == "left"

    def test_from_dict(self):
        """Test creating StyleConfig from dictionary."""
        data = {
            "font_name": "黑体",
            "font_size": "三号",
            "bold": True,
            "alignment": "center",
            "numbering_format": "chapter",
        }
        style = StyleConfig.from_dict(data)
        assert style.font_name == "黑体"
        assert style.font_size == 16  # 三号 = 16pt
        assert style.bold is True
        assert style.alignment == "center"
        assert style.numbering_format == "chapter"

    def test_to_dict(self):
        """Test converting StyleConfig to dictionary."""
        style = StyleConfig(font_name="Arial", font_size=12, bold=True)
        data = style.to_dict()
        assert data["font_name"] == "Arial"
        assert data["font_size"] == 12
        assert data["bold"] is True


class TestConfig:
    """Tests for Config class."""

    def test_default_values(self):
        """Test default values are set correctly."""
        config = Config()
        assert config.default_font == "微软雅黑"
        assert config.page_width_inches == 8.5
        assert config.max_image_width_inches == 6.0

    def test_from_dict(self):
        """Test creating Config from dictionary."""
        data = {
            "document": {
                "default_font": "仿宋",
                "max_image_width_inches": 5.0,
            },
            "styles": {
                "heading_1": {
                    "font_name": "黑体",
                    "font_size": 24,
                    "bold": True,
                }
            },
        }
        config = Config.from_dict(data)
        assert config.default_font == "仿宋"
        assert config.max_image_width_inches == 5.0
        assert "heading_1" in config.styles
        assert config.styles["heading_1"].font_name == "黑体"

    def test_get_style_existing(self):
        """Test getting existing style."""
        config = Config()
        config.styles["body"] = StyleConfig(font_name="Arial")
        style = config.get_style("body")
        assert style.font_name == "Arial"

    def test_get_style_nonexistent(self):
        """Test getting non-existent style returns default."""
        config = Config()
        config.default_font = "Times"
        style = config.get_style("nonexistent")
        assert style.font_name == "Times"

    def test_from_file(self):
        """Test loading config from file."""
        # Create temp directory and file
        temp_dir = tempfile.mkdtemp()
        config_path = Path(temp_dir) / "config.json"
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "document": {"default_font": "Arial"},
                        "styles": {"body": {"font_size": 12}},
                    },
                    f,
                )
            config = Config.from_file(config_path)
            assert config.default_font == "Arial"
        finally:
            config_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)

    def test_from_nonexistent_file(self):
        """Test loading from non-existent file returns default config."""
        config = Config.from_file("/nonexistent/path/config.json")
        assert config.default_font == "微软雅黑"

    def test_save_and_load(self):
        """Test saving and loading config."""
        config = Config()
        config.default_font = "TestFont"
        config.styles["test"] = StyleConfig(font_name="TestStyle")

        temp_dir = tempfile.mkdtemp()
        config_path = Path(temp_dir) / "config.json"
        try:
            config.save(config_path)
            loaded = Config.from_file(config_path)
            assert loaded.default_font == "TestFont"
            assert loaded.styles["test"].font_name == "TestStyle"
        finally:
            config_path.unlink(missing_ok=True)
            os.rmdir(temp_dir)


class TestDefaultConfig:
    """Tests for DEFAULT_CONFIG."""

    def test_default_config_structure(self):
        """Test DEFAULT_CONFIG has expected structure."""
        assert "document" in DEFAULT_CONFIG
        assert "styles" in DEFAULT_CONFIG
        assert "image" in DEFAULT_CONFIG

    def test_default_config_loadable(self):
        """Test DEFAULT_CONFIG can be loaded as Config."""
        config = Config.from_dict(DEFAULT_CONFIG)
        assert config.default_font == "仿宋"
        assert "heading_1" in config.styles
        assert "body" in config.styles
