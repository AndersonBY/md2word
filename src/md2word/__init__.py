"""
md2word - Convert Markdown to Word documents.

A Python library and CLI tool for converting Markdown files to Word documents (.docx)
with extensive customization options.
"""

from .config import CHINESE_FONT_SIZE_MAP, DEFAULT_CONFIG, Config, StyleConfig
from .converter import convert, convert_file
from .latex import extract_latex_formulas, latex_to_omml

__version__ = "0.1.0"
__all__ = [
    "Config",
    "StyleConfig",
    "DEFAULT_CONFIG",
    "CHINESE_FONT_SIZE_MAP",
    "convert",
    "convert_file",
    "extract_latex_formulas",
    "latex_to_omml",
]
