"""
Font name normalizer and standard font mapper.

PDF files frequently embed custom subset fonts with mangled prefixes
(e.g. "ABCDEF+Arial-BoldMT", "XYZABC+TimesNewRomanPS") or use non-standard
internal names.  This module cleans those strings and maps them to standard,
widely-available system fonts so that Word renders them accurately without
falling back to unexpected default fonts.
"""
from __future__ import annotations

import re


class FontMapper:
    # Standard fallback families
    FALLBACK_SERIF = "Times New Roman"
    FALLBACK_SANS  = "Arial"
    FALLBACK_MONO  = "Courier New"

    # Known family mapping (all keys lowercase and alphanumeric)
    FONT_MAP: dict[str, str] = {
        # Sans-serif
        "helvetica":      "Arial",
        "arial":          "Arial",
        "calibri":        "Calibri",
        "segoe":          "Segoe UI",
        "segoeui":        "Segoe UI",
        "roboto":         "Arial",
        "opensans":       "Arial",
        "lato":           "Arial",
        "verdana":        "Verdana",
        "tahoma":         "Tahoma",
        "trebuchet":      "Trebuchet MS",
        "trebuchetms":    "Trebuchet MS",
        "centurygothic":  "Century Gothic",
        "franklingothic": "Franklin Gothic Medium",
        "aptos":          "Aptos",

        # Serif
        "times":          "Times New Roman",
        "timesnewroman":  "Times New Roman",
        "cambria":        "Cambria",
        "georgia":        "Georgia",
        "garamond":       "Garamond",
        "palatino":       "Palatino Linotype",
        "baskerville":    "Baskerville",
        "minion":         "Minion Pro",
        "bookantiqua":    "Book Antiqua",

        # Monospace
        "courier":        "Courier New",
        "couriernew":     "Courier New",
        "consolas":       "Consolas",
        "monaco":         "Courier New",
        "menlo":          "Courier New",
        "dejavusansmono": "Courier New",
        "sourcecodepro":  "Consolas",
    }

    # Style suffixes to strip iteratively
    _STYLE_SUFFIXES_RE = re.compile(
        r"[-_, ]+(bold|italic|oblique|regular|roman|mt|ps|pro|condensed|black|heavy|light|medium|semibold|demibold)\b",
        flags=re.IGNORECASE,
    )

    @classmethod
    def clean_font_name(cls, raw_font: str | None) -> str:
        """
        Normalizes a PDF font name to a clean system font name.

        Examples:
          "ABCDEF+Arial-BoldMT"       -> "Arial"
          "XYZ123+TimesNewRoman,Bold" -> "Times New Roman"
          "Consolas-Italic"           -> "Consolas"
          "F1"                        -> "Arial"
        """
        if not raw_font:
            return cls.FALLBACK_SANS

        clean = raw_font.strip()

        # 1. Remove 6-letter subset prefix (e.g. "ABCDEF+Arial" -> "Arial")
        clean = re.sub(r"^[A-Z]{6}\+", "", clean)

        # 2. Iteratively strip style suffixes (handles "Arial-BoldItalicMT")
        prev = ""
        while prev != clean:
            prev = clean
            clean = cls._STYLE_SUFFIXES_RE.sub("", clean).strip()

        # Clean base without punctuation
        lookup_key = re.sub(r"[^a-zA-Z0-9]", "", clean).lower()
        if not lookup_key:
            return cls.FALLBACK_SANS

        # 3. Direct or prefix/substring lookup in map
        if lookup_key in cls.FONT_MAP:
            return cls.FONT_MAP[lookup_key]

        for key, target in cls.FONT_MAP.items():
            if key in lookup_key or lookup_key.startswith(key):
                return target

        # 4. Fallback heuristics based on common hints in name
        if any(hint in lookup_key for hint in ("mono", "code", "console")):
            return cls.FALLBACK_MONO
        if any(hint in lookup_key for hint in ("serif", "roman", "times")):
            return cls.FALLBACK_SERIF
        if any(hint in lookup_key for hint in ("sans", "arial", "gothic")):
            return cls.FALLBACK_SANS

        # 5. If it looks like a genuine named font (> 3 chars, not generic "f1", "tt0"),
        # preserve the cleaned base with standard title casing
        if len(clean) > 3 and not re.match(r"^[a-zA-Z]{1,2}\d+$", clean):
            return clean

        return cls.FALLBACK_SANS
