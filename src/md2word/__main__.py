"""
CLI entry point for md2word.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .config import DEFAULT_CONFIG, Config
from .converter import HeadingNumbering, convert, convert_file


def _load_config(args: argparse.Namespace) -> tuple[Config, str | None]:
    """Load config based on CLI args.

    Returns (config, config_source) where config_source is the file path
    string if loaded from file, or None if using defaults.
    """
    explicit_config = args.config is not None
    config_path_str = args.config if explicit_config else "config.json"
    config_path = Path(config_path_str)

    if config_path.exists():
        return Config.from_file(config_path), str(config_path)

    if explicit_config:
        print(f"[WARN] Config file not found: {config_path}, using default config")

    return Config(), None


def _cmd_show_config(args: argparse.Namespace) -> int:
    config, _ = _load_config(args)
    print(json.dumps(config.to_dict(), ensure_ascii=False, indent=4))
    return 0


def _cmd_list_formats() -> int:
    print("Available heading numbering formats:")
    for fmt_name in HeadingNumbering.FORMATS:
        if fmt_name == "none":
            print(f"  {fmt_name:<20}(no numbering)")
            continue
        numbering = HeadingNumbering()
        examples = [numbering.get_number(1, fmt_name) for _ in range(3)]
        print(f"  {fmt_name:<20}{', '.join(examples)}")
    print("\nCustom format strings with {n} (arabic) and {cn} (chinese) are also supported.")
    print("  Example: \"Part {n}\" -> Part 1, Part 2, ...")
    print("  Example: \"第{cn}部分\" -> 第一部分, 第二部分, ...")
    return 0


def _cmd_validate_config(args: argparse.Namespace) -> int:
    config, source = _load_config(args)
    if source:
        print(f"[INFO] Validating config: {source}")
    else:
        print("[INFO] Validating default config")
    warnings = config.validate()
    if warnings:
        for w in warnings:
            print(f"[WARN] {w}")
        return 1
    print("[OK] Config is valid")
    return 0


def _cmd_init_config(args: argparse.Namespace) -> int:
    config_path = Path(args.config if args.config else "config.json")
    if config_path.exists():
        print(f"[ERROR] Config file already exists: {config_path}")
        return 1
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
    print(f"[INFO] Config file created: {config_path}")
    return 0


def main() -> int:
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        prog="md2word",
        description="Convert Markdown files to Word documents (.docx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Examples:
  md2word input.md                    Convert to input.docx
  md2word input.md -o output.docx     Specify output file
  md2word input.md --toc              Add table of contents
  md2word input.md -c config.json     Use custom config file
  md2word --init-config               Generate default config file
  md2word --show-config               Show effective config (JSON)
  md2word --list-formats              List heading numbering formats
  md2word - -o output.docx            Read Markdown from stdin
  md2word --validate-config -c x.json Validate a config file
        """,
    )
    parser.add_argument(
        "input", nargs="?",
        help="Input Markdown file path (use '-' to read from stdin)",
    )
    parser.add_argument("-o", "--output", help="Output Word file path (default: input with .docx extension)")
    parser.add_argument("-c", "--config", default=None, help="Config file path (default: config.json if exists)")
    parser.add_argument("--toc", action="store_true", help="Add table of contents at the beginning")
    parser.add_argument("--toc-title", default="目录", help="TOC title (default: 目录)")
    parser.add_argument("--toc-level", type=int, default=3, help="Maximum heading level for TOC (default: 3)")
    parser.add_argument("--init-config", action="store_true", help="Generate default config file")
    parser.add_argument("--show-config", action="store_true", help="Show effective config as JSON")
    parser.add_argument("--list-formats", action="store_true", help="List available heading numbering formats")
    parser.add_argument("--validate-config", action="store_true", help="Validate config file")
    parser.add_argument("-v", "--version", action="store_true", help="Show version")

    args = parser.parse_args()

    if args.version:
        from . import __version__

        print(f"md2word {__version__}")
        return 0

    if args.list_formats:
        return _cmd_list_formats()

    if args.init_config:
        return _cmd_init_config(args)

    if args.show_config:
        return _cmd_show_config(args)

    if args.validate_config:
        return _cmd_validate_config(args)

    if not args.input:
        parser.print_help()
        return 1

    # stdin mode
    if args.input == "-":
        if not args.output:
            print("[ERROR] -o/--output is required when reading from stdin")
            return 1
        config, config_source = _load_config(args)
        if config_source:
            print(f"[INFO] Using config: {config_source}")
        else:
            print("[INFO] Using default config")
        markdown_content = sys.stdin.read()
        try:
            convert(
                markdown_content,
                args.output,
                config=config,
                toc=args.toc,
                toc_title=args.toc_title,
                toc_max_level=args.toc_level,
            )
            return 0
        except Exception as e:
            print(f"[ERROR] Conversion failed: {e}")
            return 1

    # file mode
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERROR] Input file not found: {input_path}")
        return 1

    config, config_source = _load_config(args)
    if config_source:
        print(f"[INFO] Using config: {config_source}")
    else:
        print("[INFO] Using default config")

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
