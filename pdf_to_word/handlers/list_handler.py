"""
List item detection and marker normalization handler.

Responsibilities:
  1. Detect bullet and numbered list items from paragraph text and glyph spans.
  2. Strip bullet characters and numbering prefixes from the text spans so that
     Word's native "List Bullet" and "List Number" styles do not produce
     duplicate markers (e.g. "• • Item text" or "1. 1. Item text").
  3. Correctly handle cases where the bullet glyph is in its own isolated span
     (common in PDF drawing streams).
"""
from __future__ import annotations

import re

from ..models.document_schema import Block, ElementType

# Single bullet characters recognized at start of paragraph or isolated in first span
_BULLET_CHARS = frozenset([
    "\u2022",  # •
    "\u2023",  # ‣
    "\u25e6",  # ◦
    "\u2043",  # ⁃
    "\u2219",  # ∙
    "\u25cf",  # ●
    "\u25cb",  # ○
    "\u25a0",  # ■
    "\u25aa",  # ▪
    "\u25b8",  # ▸
    "\u2713",  # ✓
    "\u2714",  # ✔
    "-",
    "\u2013",  # en dash
    "\u2014",  # em dash
    "*",
    ">",
])

# Regex for bullet followed by space
_BULLET_PREFIX_RE = re.compile(
    r"^([\u2022\u2023\u25e6\u2043\u2219\u25cf\u25cb\u25a0\u25aa\u25b8\u2713\u2714\*\-\u2013\u2014>])\s+"
)

# Regex for numbered list markers: "1. ", "1) ", "(1) ", "a. ", "a) ", "(a) ", "i. ", "(i) "
_NUMBER_PREFIX_RE = re.compile(
    r"^(\(?\d{1,3}[\.\)]|\(?[a-zA-Z][\.\)]|\(?[ivxlcdmIVXLCDM]{1,4}[\.\)])\s+"
)


class ListHandler:
    @classmethod
    def process_lists(cls, blocks: list[Block]) -> list[Block]:
        """
        Scans PARAGRAPH blocks. If a block starts with a list marker:
          - Sets element_type = ElementType.LIST_ITEM
          - Sets list_type = 'bullet' | 'numbered'
          - Strips the marker text so Word native list styles handle the bullets cleanly.
        """
        for block in blocks:
            if block.element_type != ElementType.PARAGRAPH or not block.spans:
                continue

            first_span = block.spans[0]
            raw_text = first_span.text
            stripped = raw_text.lstrip()

            if not stripped:
                continue

            # Case A: First span is exclusively a bullet glyph (e.g. Symbol/Wingdings font)
            if stripped in _BULLET_CHARS:
                block.element_type = ElementType.LIST_ITEM
                block.list_type = "bullet"
                if len(block.spans) > 1:
                    block.spans.pop(0)
                    # Strip leading spaces from next span if any
                    block.spans[0].text = block.spans[0].text.lstrip()
                else:
                    first_span.text = ""
                continue

            # Case B: Numbered list marker at start of first span
            num_match = _NUMBER_PREFIX_RE.match(stripped)
            if num_match:
                # Disqualify common false positives like "3.14", "e.g.", "i.e."
                matched_str = num_match.group(0)
                marker_core = num_match.group(1).lower().strip("().")
                if marker_core not in ("eg", "ie"):
                    block.element_type = ElementType.LIST_ITEM
                    block.list_type = "numbered"
                    # Strip marker prefix from text
                    remainder = stripped[len(matched_str):]
                    first_span.text = remainder
                    if not first_span.text.strip() and len(block.spans) > 1:
                        block.spans.pop(0)
                    continue

            # Case C: Bullet glyph + space at start of first span
            bullet_match = _BULLET_PREFIX_RE.match(stripped)
            if bullet_match:
                matched_str = bullet_match.group(0)
                block.element_type = ElementType.LIST_ITEM
                block.list_type = "bullet"
                remainder = stripped[len(matched_str):]
                first_span.text = remainder
                if not first_span.text.strip() and len(block.spans) > 1:
                    block.spans.pop(0)
                continue

        return blocks
