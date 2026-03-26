"""Heading numbering helpers."""

from __future__ import annotations


CHINESE_NUMBERS = [
    "零",
    "一",
    "二",
    "三",
    "四",
    "五",
    "六",
    "七",
    "八",
    "九",
    "十",
    "十一",
    "十二",
    "十三",
    "十四",
    "十五",
    "十六",
    "十七",
    "十八",
    "十九",
    "二十",
]


def number_to_chinese(n: int) -> str:
    """Convert number to Chinese."""
    if n <= 20:
        return CHINESE_NUMBERS[n]
    return str(n)


class HeadingNumbering:
    """Heading numbering manager."""

    FORMATS = {
        "chapter": "第{n}章",
        "section": "第{n}节",
        "chinese": "{n}、",
        "chinese_paren": "（{n}）",
        "arabic": "{n}.",
        "arabic_paren": "({n})",
        "arabic_bracket": "[{n}]",
        "roman": "{n}.",
        "roman_lower": "{n}.",
        "letter": "{n}.",
        "letter_lower": "{n}.",
        "circle": "{n}",
        "none": "",
    }

    ROMAN_NUMERALS = [
        "",
        "I",
        "II",
        "III",
        "IV",
        "V",
        "VI",
        "VII",
        "VIII",
        "IX",
        "X",
        "XI",
        "XII",
        "XIII",
        "XIV",
        "XV",
        "XVI",
        "XVII",
        "XVIII",
        "XIX",
        "XX",
    ]

    CIRCLE_NUMBERS = [
        "⓪",
        "①",
        "②",
        "③",
        "④",
        "⑤",
        "⑥",
        "⑦",
        "⑧",
        "⑨",
        "⑩",
        "⑪",
        "⑫",
        "⑬",
        "⑭",
        "⑮",
        "⑯",
        "⑰",
        "⑱",
        "⑲",
        "⑳",
    ]

    def __init__(self):
        self.counters: dict[int, int] = {}

    def reset(self, level: int | None = None):
        """Reset counters."""
        if level is None:
            self.counters = {}
            return

        for lvl in list(self.counters.keys()):
            if lvl >= level:
                self.counters[lvl] = 0

    def get_number(self, level: int, format_name: str | None) -> str:
        """Get numbering for specified level."""
        if not format_name or format_name == "none":
            return ""

        if level not in self.counters:
            self.counters[level] = 0
        self.counters[level] += 1

        for lvl in list(self.counters.keys()):
            if lvl > level:
                self.counters[lvl] = 0

        n = self.counters[level]

        if format_name in ("chapter", "section", "chinese", "chinese_paren"):
            chinese_n = number_to_chinese(n)
            return self.FORMATS[format_name].format(n=chinese_n)
        if format_name in ("arabic", "arabic_paren", "arabic_bracket"):
            return self.FORMATS[format_name].format(n=n)
        if format_name == "roman":
            roman = self.ROMAN_NUMERALS[n] if n <= 20 else str(n)
            return f"{roman}."
        if format_name == "roman_lower":
            roman = self.ROMAN_NUMERALS[n].lower() if n <= 20 else str(n)
            return f"{roman}."
        if format_name == "letter":
            letter = chr(ord("A") + n - 1) if n <= 26 else str(n)
            return f"{letter}."
        if format_name == "letter_lower":
            letter = chr(ord("a") + n - 1) if n <= 26 else str(n)
            return f"{letter}."
        if format_name == "circle":
            return self.CIRCLE_NUMBERS[n] if n <= 20 else f"({n})"

        try:
            return format_name.format(n=n, cn=number_to_chinese(n))
        except (KeyError, ValueError):
            return f"{n}. "


__all__ = ["CHINESE_NUMBERS", "HeadingNumbering", "number_to_chinese"]
