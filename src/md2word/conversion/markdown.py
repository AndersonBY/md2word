"""Markdown compatibility helpers."""

from __future__ import annotations

import html
import re
import unicodedata
from collections.abc import Sequence

import markdown2
from bs4 import BeautifulSoup
from bs4.element import NavigableString, PageElement, Tag


def _markdown2_has_punctuated_emphasis_regression() -> bool:
    """Detect markdown2 versions that miss strong/emphasis around punctuation-adjacent spans."""
    probe = 'a**“b”**c'
    try:
        return "<strong>" not in markdown2.markdown(probe)
    except Exception:
        return False


_MARKDOWN2_PUNCTUATED_EMPHASIS_COMPAT = _markdown2_has_punctuated_emphasis_regression()
_LEFTOVER_STRONG_RE = re.compile(r"(?<!\\)(?<!\*)\*\*(?=\S)(?P<content>.+?\S)\*\*(?!\*)")
_LEFTOVER_EM_RE = re.compile(r"(?<!\\)(?<!\*)\*(?=\S)(?P<content>.+?\S)\*(?!\*)")
_HTML_SKIP_TAGS = {"code", "pre", "script", "style"}


def _is_punctuation_char(char: str) -> bool:
    """Return whether a character is Unicode punctuation."""
    return bool(char) and unicodedata.category(char).startswith("P")


def _should_fix_leftover_emphasis(text: str, match: re.Match[str]) -> bool:
    """Fix only the punctuation-adjacent emphasis spans markdown2 2.5.5 regressed on."""
    content = match.group("content")
    if not content:
        return False

    if not (_is_punctuation_char(content[0]) or _is_punctuation_char(content[-1])):
        return False

    before = text[match.start() - 1] if match.start() > 0 else ""
    after = text[match.end()] if match.end() < len(text) else ""
    return before.isalnum() or after.isalnum()


def _replace_leftover_emphasis(text: str, pattern: re.Pattern[str], tag: str) -> str:
    """Convert markdown emphasis syntax that markdown2 left in a text node into HTML tags."""

    def repl(match: re.Match[str]) -> str:
        if not _should_fix_leftover_emphasis(text, match):
            return match.group(0)
        content = html.escape(match.group("content"))
        return f"<{tag}>{content}</{tag}>"

    return pattern.sub(repl, text)


def _node_text(node: PageElement) -> str:
    """Return rendered text for a direct child node."""
    if isinstance(node, NavigableString):
        return str(node)
    if isinstance(node, Tag):
        return node.get_text()
    return str(node)


def _is_delimiter_at(text: str, pos: int, marker: str) -> bool:
    """Return whether marker at pos is a standalone emphasis delimiter."""
    if text[pos : pos + len(marker)] != marker:
        return False
    if pos > 0 and text[pos - 1] == "\\":
        return False

    end = pos + len(marker)
    if marker == "**":
        if pos > 0 and text[pos - 1] == "*":
            return False
        if end < len(text) and text[end] == "*":
            return False
    else:
        if pos > 0 and text[pos - 1] == "*":
            return False
        if end < len(text) and text[end] == "*":
            return False

    return True


def _find_delimiter(text: str, marker: str, start: int = 0) -> int:
    """Find the next standalone emphasis delimiter in a text node."""
    pos = text.find(marker, start)
    while pos != -1:
        if _is_delimiter_at(text, pos, marker):
            return pos
        pos = text.find(marker, pos + 1)
    return -1


def _char_before(children: Sequence[PageElement], idx: int, pos: int) -> str:
    """Return the rendered character immediately before a marker."""
    text = _node_text(children[idx])
    if pos > 0:
        return text[pos - 1]
    for prev_idx in range(idx - 1, -1, -1):
        prev_text = _node_text(children[prev_idx])
        if prev_text:
            return prev_text[-1]
    return ""


def _char_after(children: Sequence[PageElement], idx: int, pos: int) -> str:
    """Return the rendered character immediately after a marker."""
    text = _node_text(children[idx])
    if pos < len(text):
        return text[pos]
    for next_idx in range(idx + 1, len(children)):
        next_text = _node_text(children[next_idx])
        if next_text:
            return next_text[0]
    return ""


def _should_fix_leftover_emphasis_chars(
    before: str, after: str, content_first: str, content_last: str
) -> bool:
    """Apply the markdown2 regression guard using explicit boundary characters."""
    if not content_first or not content_last:
        return False
    if content_first.isspace() or content_last.isspace():
        return False
    if not (_is_punctuation_char(content_first) or _is_punctuation_char(content_last)):
        return False
    return before.isalnum() or after.isalnum()


def _append_child(wrapper: Tag, node: PageElement) -> None:
    """Append node into wrapper, flattening redundant nested emphasis tags."""
    if isinstance(node, Tag) and node.name == wrapper.name:
        for child in list(node.contents):
            wrapper.append(child.extract())
        node.decompose()
        return
    wrapper.append(node)


def _range_crosses_skip_tag(children: Sequence[PageElement], start_idx: int, end_idx: int) -> bool:
    """Return whether an emphasis range crosses tags that should stay literal."""
    return any(
        isinstance(node, Tag) and node.name in _HTML_SKIP_TAGS for node in children[start_idx + 1 : end_idx]
    )


def _find_cross_node_emphasis_range(
    parent: Tag, marker: str
) -> tuple[int, int, int, int] | None:
    """Find a leftover emphasis span that crosses direct child nodes."""
    children = list(parent.contents)

    for open_idx, child in enumerate(children):
        if not isinstance(child, NavigableString):
            continue

        text = str(child)
        open_pos = _find_delimiter(text, marker)
        while open_pos != -1:
            before = _char_before(children, open_idx, open_pos)
            content_first = _char_after(children, open_idx, open_pos + len(marker))
            if content_first.isspace():
                open_pos = _find_delimiter(text, marker, open_pos + 1)
                continue

            for close_idx in range(open_idx + 1, len(children)):
                if _range_crosses_skip_tag(children, open_idx, close_idx):
                    break
                close_child = children[close_idx]
                if not isinstance(close_child, NavigableString):
                    continue

                close_text = str(close_child)
                close_pos = _find_delimiter(close_text, marker)
                while close_pos != -1:
                    content_last = _char_before(children, close_idx, close_pos)
                    after = _char_after(children, close_idx, close_pos + len(marker))
                    if _should_fix_leftover_emphasis_chars(before, after, content_first, content_last):
                        return (open_idx, open_pos, close_idx, close_pos)
                    close_pos = _find_delimiter(close_text, marker, close_pos + 1)

            open_pos = _find_delimiter(text, marker, open_pos + 1)

    return None


def _wrap_cross_node_emphasis(
    soup: BeautifulSoup, parent: Tag, marker: str, tag_name: str, match: tuple[int, int, int, int]
) -> None:
    """Wrap a leftover emphasis span that crosses direct child nodes."""
    open_idx, open_pos, close_idx, close_pos = match
    children = list(parent.contents)
    open_node = children[open_idx]
    close_node = children[close_idx]
    open_text = str(open_node)
    close_text = str(close_node)

    prefix = open_text[:open_pos]
    start_content = open_text[open_pos + len(marker) :]
    end_content = close_text[:close_pos]
    suffix = close_text[close_pos + len(marker) :]

    wrapper = soup.new_tag(tag_name)
    if prefix:
        open_node.insert_before(NavigableString(prefix))
    open_node.insert_before(wrapper)

    if start_content:
        wrapper.append(NavigableString(start_content))

    current = open_node.next_sibling
    while current is not None and current is not close_node:
        next_sibling = current.next_sibling
        _append_child(wrapper, current.extract())
        current = next_sibling

    if end_content:
        wrapper.append(NavigableString(end_content))

    open_node.extract()
    close_node.extract()

    if suffix:
        wrapper.insert_after(NavigableString(suffix))


def _repair_cross_node_emphasis(soup: BeautifulSoup, marker: str, tag_name: str) -> None:
    """Repair leftover emphasis that spans multiple inline nodes."""
    while True:
        updated = False
        for parent in soup.find_all(True):
            if parent.name in _HTML_SKIP_TAGS or any(ancestor.name in _HTML_SKIP_TAGS for ancestor in parent.parents):
                continue

            match = _find_cross_node_emphasis_range(parent, marker)
            if match is None:
                continue

            _wrap_cross_node_emphasis(soup, parent, marker, tag_name, match)
            updated = True
            break

        if not updated:
            return


def fix_markdown2_punctuated_emphasis_html(html_content: str) -> str:
    """Repair punctuation-adjacent emphasis spans left unparsed by markdown2 2.5.5+."""
    if not _MARKDOWN2_PUNCTUATED_EMPHASIS_COMPAT or "*" not in html_content:
        return html_content

    soup = BeautifulSoup(html_content, "html.parser")

    for text_node in list(soup.find_all(string=True)):
        if any(parent.name in _HTML_SKIP_TAGS for parent in text_node.parents):
            continue

        original = str(text_node)
        replaced = _replace_leftover_emphasis(original, _LEFTOVER_STRONG_RE, "strong")
        replaced = _replace_leftover_emphasis(replaced, _LEFTOVER_EM_RE, "em")
        if replaced == original:
            continue

        fragment = BeautifulSoup(replaced, "html.parser")
        new_nodes = list(fragment.contents)
        if not new_nodes:
            continue

        first = new_nodes[0]
        text_node.replace_with(first)
        current = first
        for node in new_nodes[1:]:
            current.insert_after(node)
            current = node

    _repair_cross_node_emphasis(soup, "**", "strong")
    _repair_cross_node_emphasis(soup, "*", "em")

    return str(soup)


__all__ = ["_MARKDOWN2_PUNCTUATED_EMPHASIS_COMPAT", "fix_markdown2_punctuated_emphasis_html"]
