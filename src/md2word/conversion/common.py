"""Shared conversion helpers."""

from __future__ import annotations


def print_info(message: str) -> None:
    """Print info message."""
    print(f"[INFO] {message}")


def print_error(message: str) -> None:
    """Print error message."""
    print(f"[ERROR] {message}")


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Convert hex color to an RGB tuple."""
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (r, g, b)


__all__ = ["hex_to_rgb", "print_error", "print_info"]
