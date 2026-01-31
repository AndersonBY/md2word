"""
CLI entry point for md2word.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import DEFAULT_CONFIG, Config
from .converter import convert_file


def main() -> int:
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        prog="md2word",
        description="Convert Markdown files to Word documents (.docx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  md2word input.md                    Convert to input.docx
  md2word input.md -o output.docx     Specify output file
  md2word input.md --toc              Add table of contents
  md2word input.md -c config.json     Use custom config file
  md2word --init-config               Generate default config file
        """,
    )
    parser.add_argument("input", nargs="?", help="Input Markdown file path")
    parser.add_argument("-o", "--output", help="Output Word file path (default: input with .docx extension)")
    parser.add_argument("-c", "--config", default="config.json", help="Config file path (default: config.json)")
    parser.add_argument("--toc", action="store_true", help="Add table of contents at the beginning")
    parser.add_argument("--toc-title", default="目录", help="TOC title (default: 目录)")
    parser.add_argument("--toc-level", type=int, default=3, help="Maximum heading level for TOC (default: 3)")
    parser.add_argument("--init-config", action="store_true", help="Generate default config file")
    parser.add_argument("-v", "--version", action="store_true", help="Show version")

    args = parser.parse_args()

    if args.version:
        from . import __version__

        print(f"md2word {__version__}")
        return 0

    if args.init_config:
        import json

        config_path = Path(args.config)
        if config_path.exists():
            print(f"[ERROR] Config file already exists: {config_path}")
            return 1
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        print(f"[INFO] Config file created: {config_path}")
        return 0

    if not args.input:
        parser.print_help()
        return 1

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERROR] Input file not found: {input_path}")
        return 1

    # Load config
    config_path = Path(args.config)
    if config_path.exists():
        config = Config.from_file(config_path)
    else:
        config = Config()

    try:
        convert_file(
            input_path,
            args.output,
            config,
            toc=args.toc,
            toc_title=args.toc_title,
            toc_max_level=args.toc_level,
        )
        return 0
    except Exception as e:
        print(f"[ERROR] Conversion failed: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
