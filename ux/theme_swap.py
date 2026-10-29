#!/usr/bin/env python3
"""Swap a .pptx, .docx or .xlsx between the Acme and Violet themes.

    python theme_swap.py in.pptx out.pptx --to violet
    python theme_swap.py in.docx out.docx --to violet
    python theme_swap.py in.xlsx out.xlsx --to acme

The file type is detected from the input extension. PowerPoint and Word
packages get colour, font and vocabulary swaps; Excel packages get colour and
vocabulary swaps only, since workbook fonts are not theme fonts and must not
change. Embedded PNG images are recoloured in every format. The Violet copy
drops the package thumbnail and Excel's saved folder path, which a reverse
conversion cannot restore and Office writes again on the next save.
"""

from __future__ import annotations

import argparse
from collections import Counter
import os
from pathlib import Path
import re
import tempfile
from xml.sax.saxutils import escape, unescape
import zipfile
import zlib


# Acme -> Violet. Both sides must remain unique and disjoint so the
# conversion can be reversed without losing the original package contents.
COLOURS = {
    "003946": "241733",
    "71777A": "6E6A7C",
    "FFCB05": "C9A6FF",
    "FEBE10": "A100FF",
    "FC9E2D": "7A1FA2",
    "5BB54B": "2F8F6B",
    "FFF8E0": "F4EFFA",
    "FFF4C2": "EADFF7",
    "FFE994": "DCC9FA",
    "E0DCD1": "DCD6E4",
    "F0E9E0": "EFEBF4",
    "A59E8C": "9C95A8",
    "F1F7F7": "F8F6FB",
    "EAEBEB": "E5E2EA",
    "DDE6E1": "E3DEEC",
    "B0E3DB": "D6CCE6",
    "E6EFEF": "EFEBF5",
    "4A7283": "5E536F",
    "D6E3E6": "DBD5E5",
    "8BBAC4": "927BAA",
    "AC6100": "5B246E",
    "6ED2DF": "8E75D8",
    "004F73": "3B2452",
    "38758E": "674779",
    "FFDB52": "D6A62C",
    "D9541A": "C43C5A",
    "F6F1EB": "F7F3FA",
    "E8E5DC": "EAE5F0",
    "B5AFA0": "B1A9BC",
    "315F6E": "4B3D5E",
    "A37C2C": "7D4BA8",
}

# Shared neutral colours are valid in both themes and are intentionally left
# unchanged. Keeping them outside COLOURS preserves the disjoint mapping rule.
UNCHANGED_COLOURS = frozenset({"000000", "FFFFFF"})

# Acme font families, grouped by where they are used. Each source keeps its
# own distinct Violet target so the reverse conversion restores exactly the
# name the file started with.
FONTS = {
    # Slides
    "Sun Life New Display": "Franklin Gothic Medium",
    "Sun Life New Text": "Franklin Gothic Book",
    # Documents
    "Sun Life Serif": "Georgia",
    "Sun Life Sans": "Segoe UI",
}

# Organization vocabulary, Acme -> Violet. Violet values are bracketed
# placeholders such as "[company]", one per term, so the swap stays one-to-one.
# The rules match COLOURS: each side unique, the two sides disjoint, and no
# entry contained in an entry from the other side. Matching is exact,
# case-sensitive and on whole words, longest entry first, so "Acme Holdings"
# is matched before "Acme". Only text content changes; tags, attributes and
# style names are never touched.
#
# To extend, add one entry per spelling that must change, including plurals,
# abbreviations and domain names, and give each its own placeholder. Text that
# Word or PowerPoint has split across several runs is only matched if the
# whole term sits in a single run.
TERMS = {
    # Organization
    "Sun Life": "[organization]",
    "sunlife.ca": "[organization_domain]",
    "SLF": "[organization_abbreviation]",
    "SLC": "[business_group_1]",
    "Asia": "[business_group_2]",
    "ASIA": "[BUSINESS_GROUP_2]",
    "Asset Management": "[business_group_1_name]",
    "Asset Mgmt.": "[business_group_1_short]",
    "Canada": "[business_group_3]",
    "US": "[business_group_4]",
    "U.S.": "[business_group_4_abbreviation]",
    "Corporate": "[business_group_5]",
    "Business Groups": "[business_units]",
    "Business Group": "[business_unit]",
    # Business lines
    "Group Benefits": "[business_line_1]",
    "HR&S": "[business_line_2]",
    # Governance bodies, teams and internal programmes
    "ETAB": "[architecture_board]",
    "DBTS": "[leadership_team]",
    "ATG": "[data_product]",
    "Enterprise Data Architecture": "[architecture_team]",
    "ENTERPRISE DATA ARCHITECTURE": "[ARCHITECTURE_TEAM]",
    "Data & Analytics capability model": "[capability_model]",
    "Imagineers": "[innovation_team]",
    "Agentic Control Plane": "[agent_control_platform]",
    "Data Control Plane": "[data_control_platform]",
    # People (document properties store the author as "Last First")
    "Saad Tadili": "[first_name_last_name]",
    "Tadili Saad": "[last_name_first_name]",
    "Melisa Muraca": "[first_name_last_name_2]",
    # Respondents. Reviewer comments address a respondent by first name.
    "Eric Robson": "[respondent_1]",
    "Eric": "[respondent_1_first_name]",
    # Platforms and vendors
    "SharePoint": "[document_platform]",
    "Copilot": "[ai_assistant]",
    "ServiceNow": "[itsm_platform]",
    "Ardoq": "[architecture_tool]",
    "Snowflake Horizon": "[data_platform_governance]",
    "Snowflake": "[data_platform]",
    "AWS S3": "[object_storage]",
    "AWS": "[cloud_provider]",
    "Collibra": "[data_catalog]",
    "Tableau": "[bi_tool]",
    # Regulators and laws
    "OSFI Guideline E-23": "[regulator_model_risk_guideline]",
    "OSFI B-13": "[regulator_technology_risk_guideline]",
    "OSFI": "[regulator]",
    "PIPEDA": "[federal_privacy_law]",
    "Quebec Law 25": "[provincial_privacy_law]",
}

# Embedded PNG images (icons, logos) are recoloured with COLOURS. A pixel
# within this many levels of a mapped colour on every channel is shifted by
# that colour's offset, alpha unchanged, so anti-aliased edges follow and the
# reverse shift restores every pixel exactly.
IMAGE_TOLERANCE = 4

# Package profiles. Which parts are converted, which parts must exist, and
# whether fonts are swapped all depend on the package format.
FORMATS = {
    ".pptx": {
        "label": "PowerPoint",
        "media_prefix": "ppt/media/",
        "part_prefixes": (
            "docProps/",
            "ppt/charts/",
            "ppt/diagrams/",
            "ppt/handoutMasters/",
            "ppt/notesMasters/",
            "ppt/notesSlides/",
            "ppt/slideLayouts/",
            "ppt/slideMasters/",
            "ppt/slides/",
            "ppt/theme/",
        ),
        "part_files": frozenset({"ppt/tableStyles.xml"}),
        "required_parts": frozenset({"[Content_Types].xml", "ppt/presentation.xml"}),
        "swap_fonts": True,
    },
    ".docx": {
        "label": "Word",
        "media_prefix": "word/media/",
        "part_prefixes": (
            "docProps/",
            "word/charts/",
            "word/comments",
            "word/diagrams/",
            "word/document",
            "word/endnotes",
            "word/fontTable",
            "word/footer",
            "word/footnotes",
            "word/header",
            "word/numbering",
            "word/styles",
            "word/theme/",
        ),
        "part_files": frozenset(),
        "required_parts": frozenset({"[Content_Types].xml", "word/document.xml"}),
        "swap_fonts": True,
    },
    ".xlsx": {
        "label": "Excel",
        "media_prefix": "xl/media/",
        "part_prefixes": (
            "docProps/",
            "xl/charts/",
            "xl/chartsheets/",
            "xl/drawings/",
            "xl/tables/",
            "xl/theme/",
            "xl/worksheets/",
        ),
        "part_files": frozenset({"xl/styles.xml", "xl/sharedStrings.xml"}),
        "required_parts": frozenset({"[Content_Types].xml", "xl/workbook.xml"}),
        "swap_fonts": False,
    },
}

# DrawingML colours (slides, themes, charts, drawings): srgbClr val="RRGGBB".
SRGB_RE = re.compile(r'(srgbClr\s+)val="([0-9A-Fa-f]{6})"')
# SpreadsheetML colours (styles, fills, borders, tab colours, conditional
# formats, rich text runs): rgb="AARRGGBB". The alpha byte is preserved.
ARGB_RE = re.compile(r'(rgb=")([0-9A-Fa-f]{2})([0-9A-Fa-f]{6})(")')
TYPEFACE_RE = re.compile(r'typeface="([^"]*)"')
# Per-script fallback fonts in a theme (the default Office theme maps Lisu to
# Segoe UI, for example). They are not theme choices, so they are neither
# counted nor swapped; otherwise a stock theme would read as mixed.
SCRIPT_FONT_RE = re.compile(r'<a:font\s+script="[^"]*"[^>]*/>')
LAST_MODIFIED_RE = re.compile(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>")
# WordprocessingML colours: <w:color w:val="RRGGBB"/> for text, and the fill
# and color attributes used by shading and borders.
WORD_COLOR_RE = re.compile(r'(<w:color\b[^>]*?\bw:val=")([0-9A-Fa-f]{6})(")')
WORD_ATTR_RE = re.compile(r'(\bw:(?:fill|color)=")([0-9A-Fa-f]{6})(")')
# WordprocessingML fonts: run fonts and the font table.
WORD_FONT_RE = re.compile(r'(\bw:(?:ascii|hAnsi|cs|eastAsia|name)=")([^"]*)(")')
# Text content between tags. Vocabulary swaps are confined to these spans.
TEXT_RE = re.compile(r">([^<]+)<")
# The package thumbnail is a picture of the first page in the original theme,
# text included. It cannot be recoloured, so the Violet copy drops it; Office
# writes a new one on the next save.
THUMBNAIL_PREFIX = "docProps/thumbnail."
THUMBNAIL_REL_RE = re.compile(r'<Relationship\b[^>]*Target="/?docProps/thumbnail\.[^"]*"[^>]*/>')
# Excel records the folder a workbook was last saved in, user name included.
# The Violet copy drops it; Excel writes the new location on the next save.
ABSPATH_RE = re.compile(
    r"<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<x15ac:absPath\b[^>]*/>"
    r"\s*</mc:Choice>\s*</mc:AlternateContent>"
)
PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


def term_pattern(terms: list[str]) -> re.Pattern[str] | None:
    """Match any of the given terms as whole words, longest first."""
    if not terms:
        return None
    ordered = sorted(terms, key=len, reverse=True)
    alternatives = "|".join(re.escape(escape(term)) for term in ordered)
    return re.compile(rf"(?<![0-9A-Za-z])(?:{alternatives})(?![0-9A-Za-z])")


def detect_format(path: Path) -> dict:
    profile = FORMATS.get(path.suffix.lower())
    if profile is None:
        supported = ", ".join(sorted(FORMATS))
        raise ValueError(
            f"Unsupported file type '{path.suffix}' for {path.name}; "
            f"supported: {supported}"
        )
    return profile


def validate_mapping() -> None:
    """Ensure each direction is deterministic and reversible."""
    colour_sources = set(COLOURS)
    colour_targets = set(COLOURS.values())
    font_sources = set(FONTS)
    font_targets = set(FONTS.values())

    if len(colour_targets) != len(COLOURS):
        raise ValueError("Violet colour values must be unique")
    if colour_sources & colour_targets:
        raise ValueError("Acme and Violet colour values must be disjoint")
    if UNCHANGED_COLOURS & (colour_sources | colour_targets):
        raise ValueError("Unchanged colours must not appear in the mapping")
    if len(font_targets) != len(FONTS):
        raise ValueError("Violet font values must be unique")
    if font_sources & font_targets:
        raise ValueError("Acme and Violet font values must be disjoint")

    term_sources = set(TERMS)
    term_targets = set(TERMS.values())
    if "" in term_sources | term_targets:
        raise ValueError("Vocabulary entries must not be empty")
    if len(term_targets) != len(TERMS):
        raise ValueError("Violet vocabulary values must be unique")
    # Terms match whole words only, so "US" inside "[BUSINESS_GROUP_2]" is
    # harmless; what must never happen is a whole-word match across sides.
    for source in term_sources:
        for target in term_targets:
            if term_pattern([source]).search(escape(target)) or term_pattern(
                [target]
            ).search(escape(source)):
                raise ValueError(
                    f"Acme and Violet vocabulary must be disjoint: "
                    f"'{source}' overlaps '{target}'"
                )


def build(
    target: str, swap_fonts: bool
) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    """Return the colour, font and vocabulary mappings for the target theme."""
    if target == "violet":
        colour_map = dict(COLOURS)
        font_map = dict(FONTS)
        term_map = dict(TERMS)
    else:
        colour_map = {value: key for key, value in COLOURS.items()}
        font_map = {value: key for key, value in FONTS.items()}
        term_map = {value: key for key, value in TERMS.items()}
    if not swap_fonts:
        font_map = {}
    # Terms are matched against XML text, so compare in escaped form.
    term_map = {escape(key): escape(value) for key, value in term_map.items()}
    return colour_map, font_map, term_map


def is_convertible_part(name: str, profile: dict) -> bool:
    return name.endswith(".xml") and (
        name in profile["part_files"] or name.startswith(profile["part_prefixes"])
    )


def is_image_part(name: str, profile: dict) -> bool:
    return name.startswith(profile["media_prefix"]) and name.lower().endswith(".png")


def png_chunks(data: bytes) -> list[tuple[bytes, bytes]]:
    if not data.startswith(PNG_SIGNATURE):
        raise ValueError("not a PNG")
    chunks = []
    position = len(PNG_SIGNATURE)
    while position < len(data):
        length = int.from_bytes(data[position : position + 4], "big")
        kind = data[position + 4 : position + 8]
        chunks.append((kind, data[position + 8 : position + 8 + length]))
        position += 12 + length
    return chunks


def png_decode(data: bytes) -> tuple[list[tuple[bytes, bytes]], dict, bytearray] | None:
    """Return chunks, header and raw pixels of an 8-bit non-interlaced RGB,
    RGBA or palette PNG; None for any other kind, which is left unchanged."""
    chunks = png_chunks(data)
    header = chunks[0][1]
    width, height = int.from_bytes(header[0:4], "big"), int.from_bytes(header[4:8], "big")
    depth, kind, interlace = header[8], header[9], header[12]
    if depth != 8 or interlace != 0 or kind not in (2, 3, 6):
        return None
    info = {"width": width, "height": height, "kind": kind, "channels": {2: 3, 3: 1, 6: 4}[kind]}
    if kind == 3:
        return chunks, info, bytearray()
    raw = zlib.decompress(b"".join(body for name, body in chunks if name == b"IDAT"))
    bpp = info["channels"]
    stride = width * bpp
    pixels = bytearray(stride * height)
    previous = bytearray(stride)
    position = 0
    for row in range(height):
        method = raw[position]
        line = bytearray(raw[position + 1 : position + 1 + stride])
        position += 1 + stride
        for i in range(stride):
            left = line[i - bpp] if i >= bpp else 0
            up = previous[i]
            if method == 1:
                line[i] = (line[i] + left) & 0xFF
            elif method == 2:
                line[i] = (line[i] + up) & 0xFF
            elif method == 3:
                line[i] = (line[i] + ((left + up) >> 1)) & 0xFF
            elif method == 4:
                corner = previous[i - bpp] if i >= bpp else 0
                estimate = left + up - corner
                pa, pb, pc = abs(estimate - left), abs(estimate - up), abs(estimate - corner)
                predictor = left if pa <= pb and pa <= pc else (up if pb <= pc else corner)
                line[i] = (line[i] + predictor) & 0xFF
        pixels[row * stride : (row + 1) * stride] = line
        previous = line
    return chunks, info, pixels


def png_chunk(kind: bytes, body: bytes) -> bytes:
    return (
        len(body).to_bytes(4, "big")
        + kind
        + body
        + zlib.crc32(kind + body).to_bytes(4, "big")
    )


def png_encode(chunks: list[tuple[bytes, bytes]], info: dict, pixels: bytearray) -> bytes:
    """Rebuild the PNG with new pixels (or a new palette), other chunks kept."""
    out = [PNG_SIGNATURE]
    written = False
    for kind, body in chunks:
        if kind == b"IDAT" and info["kind"] != 3:
            if not written:
                stride = info["width"] * info["channels"]
                raw = b"".join(
                    b"\x00" + bytes(pixels[row * stride : (row + 1) * stride])
                    for row in range(info["height"])
                )
                out.append(png_chunk(b"IDAT", zlib.compress(raw, 9)))
                written = True
            continue
        out.append(png_chunk(kind, body))
    return b"".join(out)


def nearby(rgb: tuple[int, int, int], colours: dict[str, tuple[int, int, int]]) -> list[str]:
    return [
        name
        for name, ref in colours.items()
        if max(abs(rgb[0] - ref[0]), abs(rgb[1] - ref[1]), abs(rgb[2] - ref[2]))
        <= IMAGE_TOLERANCE
    ]


def as_rgb(value: str) -> tuple[int, int, int]:
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def recolour_png(data: bytes, colour_map: dict[str, str], name: str) -> tuple[bytes, int]:
    """Shift pixels near a mapped colour by that colour's offset.

    Raises if a pixel is near two mapped colours, would leave the 0-255
    range, or already sits near a target colour: in each case the reverse
    conversion could not restore the original exactly."""
    decoded = png_decode(data)
    if decoded is None:
        return data, 0
    chunks, info, pixels = decoded
    sources = {key: as_rgb(key) for key in colour_map}
    targets = {value: as_rgb(value) for value in colour_map.values()}
    cache: dict[tuple[int, int, int], tuple[int, int, int] | None] = {}

    def shifted(rgb: tuple[int, int, int]) -> tuple[int, int, int] | None:
        if rgb in cache:
            return cache[rgb]
        near_source = nearby(rgb, sources)
        if len(near_source) > 1:
            raise ValueError(f"{name}: pixel {rgb} is near several mapped colours")
        if not near_source:
            if nearby(rgb, targets):
                raise ValueError(f"{name}: pixel {rgb} is already near a target colour")
            cache[rgb] = None
            return None
        source = sources[near_source[0]]
        target = targets[colour_map[near_source[0]]]
        result = tuple(rgb[i] + target[i] - source[i] for i in range(3))
        if min(result) < 0 or max(result) > 255 or nearby(result, targets) != [
            colour_map[near_source[0]]
        ]:
            raise ValueError(f"{name}: pixel {rgb} cannot be shifted reversibly")
        cache[rgb] = result
        return result

    changes = 0
    if info["kind"] == 3:
        chunks = list(chunks)
        for index, (kind, body) in enumerate(chunks):
            if kind != b"PLTE":
                continue
            palette = bytearray(body)
            for offset in range(0, len(palette), 3):
                result = shifted(tuple(palette[offset : offset + 3]))
                if result is not None:
                    palette[offset : offset + 3] = bytes(result)
                    changes += 1
            chunks[index] = (kind, bytes(palette))
    else:
        step = info["channels"]
        for offset in range(0, len(pixels), step):
            result = shifted((pixels[offset], pixels[offset + 1], pixels[offset + 2]))
            if result is not None:
                pixels[offset : offset + 3] = bytes(result)
                changes += 1
    if not changes:
        return data, 0
    return png_encode(chunks, info, pixels), changes


def image_colours(data: bytes) -> set[str]:
    """Mapped colours of either theme found in a PNG, for the mixed-input check."""
    decoded = png_decode(data)
    if decoded is None:
        return set()
    chunks, info, pixels = decoded
    every = {key: as_rgb(key) for key in list(COLOURS) + list(COLOURS.values())}
    if info["kind"] == 3:
        body = b"".join(body for kind, body in chunks if kind == b"PLTE")
        colours = {tuple(body[i : i + 3]) for i in range(0, len(body), 3)}
    else:
        step = info["channels"]
        colours = {
            (pixels[i], pixels[i + 1], pixels[i + 2]) for i in range(0, len(pixels), step)
        }
    found: set[str] = set()
    for rgb in colours:
        found.update(nearby(rgb, every))
    return found


def theme_inventory(path: Path, profile: dict) -> tuple[Counter, Counter]:
    """Count mapped Acme and Violet values in a package, value by value."""
    acme_colours = set(COLOURS)
    violet_colours = set(COLOURS.values())
    acme_fonts = set(FONTS) if profile["swap_fonts"] else set()
    violet_fonts = set(FONTS.values()) if profile["swap_fonts"] else set()
    acme_terms = term_pattern(list(TERMS))
    violet_terms = term_pattern(list(TERMS.values()))
    acme: Counter = Counter()
    violet: Counter = Counter()

    def colour(value: str) -> None:
        value = value.upper()
        if value in acme_colours:
            acme[f"colour {value}"] += 1
        if value in violet_colours:
            violet[f"colour {value}"] += 1

    def font(value: str) -> None:
        if value in acme_fonts:
            acme[f"font {value}"] += 1
        if value in violet_fonts:
            violet[f"font {value}"] += 1

    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if is_image_part(name, profile):
                for value in image_colours(package.read(name)):
                    side = acme if value in acme_colours else violet
                    side[f"colour {value} in {name}"] += 1
                continue
            if not is_convertible_part(name, profile):
                continue
            # Office rewrites the last editor's name on every save, so it says
            # nothing about which theme the package is in.
            text = LAST_MODIFIED_RE.sub("", package.read(name).decode("utf-8"))
            for match in SRGB_RE.findall(text):
                colour(match[1])
            for match in ARGB_RE.findall(text):
                colour(match[2])
            for regex in (WORD_COLOR_RE, WORD_ATTR_RE):
                for match in regex.findall(text):
                    colour(match[1])
            for match in TYPEFACE_RE.findall(SCRIPT_FONT_RE.sub("", text)):
                font(match)
            for match in WORD_FONT_RE.findall(text):
                font(match[1])
            for name_ in acme_fonts | violet_fonts:
                for _ in range(text.count(f"<vt:lpstr>{name_}</vt:lpstr>")):
                    font(name_)
            # Font names listed in document properties also appear as text,
            # so vocabulary is counted only after those entries are removed.
            for name_ in acme_fonts | violet_fonts:
                text = text.replace(f"<vt:lpstr>{name_}</vt:lpstr>", "")
            for span in TEXT_RE.findall(text):
                for pattern, found in ((acme_terms, acme), (violet_terms, violet)):
                    if pattern is not None:
                        for term in pattern.findall(span):
                            found[f'text "{unescape(term)}" in {name}'] += 1

    return acme, violet


def validate_theme_state(path: Path, profile: dict) -> None:
    """Reject mixed inputs because collapsing both sides is not reversible."""
    acme, violet = theme_inventory(path, profile)
    if acme and violet:
        # Name the minority side: those are the values to fix by hand.
        side, found = ("Acme", acme) if sum(acme.values()) <= sum(violet.values()) else ("Violet", violet)
        listed = "; ".join(f"{key} (x{count})" for key, count in found.most_common(10))
        raise ValueError(
            "Input contains mapped values from both Acme and Violet; "
            f"conversion would not be reversible. {side} values found: {listed}"
        )


def convert(
    text: str,
    colour_map: dict[str, str],
    font_map: dict[str, str],
    term_map: dict[str, str],
) -> tuple[str, int, int, int]:
    """Convert theme values in one XML part and return replacement counts."""
    colour_changes = 0
    font_changes = 0
    term_changes = 0

    def replace_srgb(match: re.Match[str]) -> str:
        nonlocal colour_changes
        replacement = colour_map.get(match.group(2).upper())
        if replacement is None:
            return match.group(0)
        colour_changes += 1
        return f'{match.group(1)}val="{replacement}"'

    def replace_argb(match: re.Match[str]) -> str:
        nonlocal colour_changes
        replacement = colour_map.get(match.group(3).upper())
        if replacement is None:
            return match.group(0)
        colour_changes += 1
        return f"{match.group(1)}{match.group(2)}{replacement}{match.group(4)}"

    def replace_word_colour(match: re.Match[str]) -> str:
        nonlocal colour_changes
        replacement = colour_map.get(match.group(2).upper())
        if replacement is None:
            return match.group(0)
        colour_changes += 1
        return f"{match.group(1)}{replacement}{match.group(3)}"

    def replace_typeface(match: re.Match[str]) -> str:
        nonlocal font_changes
        replacement = font_map.get(match.group(1))
        if replacement is None:
            return match.group(0)
        font_changes += 1
        return f'typeface="{replacement}"'

    def replace_word_font(match: re.Match[str]) -> str:
        nonlocal font_changes
        replacement = font_map.get(match.group(2))
        if replacement is None:
            return match.group(0)
        font_changes += 1
        return f"{match.group(1)}{replacement}{match.group(3)}"

    text = SRGB_RE.sub(replace_srgb, text)
    text = ARGB_RE.sub(replace_argb, text)
    text = WORD_COLOR_RE.sub(replace_word_colour, text)
    text = WORD_ATTR_RE.sub(replace_word_colour, text)
    if font_map:
        # Swap typefaces everywhere except inside per-script fallback entries.
        pieces = []
        position = 0
        for script_font in SCRIPT_FONT_RE.finditer(text):
            pieces.append(
                TYPEFACE_RE.sub(replace_typeface, text[position : script_font.start()])
            )
            pieces.append(script_font.group(0))
            position = script_font.end()
        pieces.append(TYPEFACE_RE.sub(replace_typeface, text[position:]))
        text = "".join(pieces)
        text = WORD_FONT_RE.sub(replace_word_font, text)
        for original, replacement in font_map.items():
            source = f"<vt:lpstr>{original}</vt:lpstr>"
            occurrences = text.count(source)
            if occurrences:
                text = text.replace(
                    source, f"<vt:lpstr>{replacement}</vt:lpstr>"
                )
                font_changes += occurrences

    # Vocabulary runs last, after font names in document properties have been
    # swapped, so a font name is never rewritten as ordinary text.
    if term_map:
        # term_map keys are already escaped, so they are compiled directly
        # rather than through term_pattern, which escapes its input.
        ordered = sorted(term_map, key=len, reverse=True)
        pattern = re.compile(
            r"(?<![0-9A-Za-z])(?:"
            + "|".join(re.escape(key) for key in ordered)
            + r")(?![0-9A-Za-z])"
        )

        def replace_term(match: re.Match[str]) -> str:
            nonlocal term_changes
            term_changes += 1
            return term_map[match.group(0)]

        def replace_span(match: re.Match[str]) -> str:
            return f">{pattern.sub(replace_term, match.group(1))}<"

        text = TEXT_RE.sub(replace_span, text)

    return text, colour_changes, font_changes, term_changes


def validate_package(path: Path, profile: dict) -> None:
    """Confirm the result is a readable Office package of the expected type."""
    if not zipfile.is_zipfile(path):
        raise ValueError(f"Not a valid {profile['label']} package: {path}")

    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        missing = profile["required_parts"] - names
        if missing:
            raise ValueError(
                f"{profile['label']} package is missing: "
                + ", ".join(sorted(missing))
            )
        corrupt_part = package.testzip()
        if corrupt_part is not None:
            raise ValueError(f"Corrupt package part: {corrupt_part}")


def swap(
    src: str | Path, dst: str | Path, target: str
) -> tuple[Path, int, int, int, int, int, list[str]]:
    """Write a validated conversion without risking either input or output."""
    source = Path(src).expanduser().resolve()
    destination = Path(dst).expanduser().resolve()

    if source == destination:
        raise ValueError("Input and output paths must be different")
    if not source.is_file():
        raise FileNotFoundError(source)

    profile = detect_format(source)
    if destination.suffix.lower() != source.suffix.lower():
        raise ValueError(
            "Input and output must have the same extension: "
            f"{source.suffix} vs {destination.suffix}"
        )

    validate_mapping()
    validate_package(source, profile)
    validate_theme_state(source, profile)
    colour_map, font_map, term_map = build(target, profile["swap_fonts"])

    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.name}.", suffix=".tmp", dir=destination.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)

    converted_parts = 0
    colour_changes = 0
    font_changes = 0
    term_changes = 0
    image_changes = 0
    removed: list[str] = []

    try:
        with zipfile.ZipFile(source) as input_package, zipfile.ZipFile(
            temporary, "w", zipfile.ZIP_DEFLATED
        ) as output_package:
            for item in input_package.infolist():
                data = input_package.read(item.filename)
                if target == "violet" and item.filename.startswith(THUMBNAIL_PREFIX):
                    removed.append(f"thumbnail {item.filename}")
                    continue
                if target == "violet" and item.filename == "_rels/.rels":
                    data = THUMBNAIL_REL_RE.sub("", data.decode("utf-8")).encode("utf-8")
                if target == "violet" and item.filename == "xl/workbook.xml":
                    text, found = ABSPATH_RE.subn("", data.decode("utf-8"))
                    if found:
                        removed.append("saved folder path in xl/workbook.xml")
                        data = text.encode("utf-8")
                if is_image_part(item.filename, profile):
                    data, pixels = recolour_png(data, colour_map, item.filename)
                    if pixels:
                        image_changes += 1
                elif is_convertible_part(item.filename, profile):
                    text, colours, fonts, terms = convert(
                        data.decode("utf-8"), colour_map, font_map, term_map
                    )
                    data = text.encode("utf-8")
                    if colours or fonts or terms:
                        converted_parts += 1
                        colour_changes += colours
                        font_changes += fonts
                        term_changes += terms
                output_package.writestr(item, data)

        validate_package(temporary, profile)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)

    return (
        destination,
        converted_parts,
        colour_changes,
        font_changes,
        term_changes,
        image_changes,
        removed,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Source .pptx, .docx or .xlsx file")
    parser.add_argument("output", help="Destination file, same extension as input")
    parser.add_argument("--to", choices=("acme", "violet"), required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output, parts, colours, fonts, terms, images, removed = swap(
        args.input, args.output, args.to
    )
    print(f"{args.to.title()} theme conversion complete")
    print(f"XML parts changed: {parts}")
    print(f"Colour references changed: {colours}")
    print(f"Font references changed: {fonts}")
    print(f"Text references changed: {terms}")
    print(f"Images recoloured: {images}")
    for item in removed:
        print(f"Removed {item} (Office writes it again on save)")
    print(f"Wrote: {output}")


if __name__ == "__main__":
    main()
