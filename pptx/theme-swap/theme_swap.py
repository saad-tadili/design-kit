#!/usr/bin/env python3
"""Swap a .pptx, .docx or .xlsx between the Acme and Violet themes.

    python theme_swap.py in.pptx out.pptx --to violet
    python theme_swap.py in.docx out.docx --to violet
    python theme_swap.py in.xlsx out.xlsx --to acme

The file type is detected from the input extension. PowerPoint and Word
packages get colour, font and vocabulary swaps; Excel packages get colour and
vocabulary swaps only, since workbook fonts are not theme fonts and must not
change.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import re
import tempfile
from xml.sax.saxutils import escape
import zipfile


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
    # Governance bodies and internal programmes
    "ETAB": "[architecture_board]",
    "DBTS": "[leadership_team]",
    "ATG": "[data_product]",
    # People
    "Saad Tadili": "[first_name_last_name]",
    # Platforms and vendors
    "ServiceNow": "[itsm_platform]",
    "Ardoq": "[architecture_tool]",
    "Snowflake Horizon": "[data_platform_governance]",
    "Snowflake": "[data_platform]",
    "AWS S3": "[object_storage]",
    "AWS": "[cloud_provider]",
    "Collibra": "[data_catalog]",
    "Tableau": "[bi_tool]",
    # Regulators
    "OSFI": "[regulator]",
}

# Package profiles. Which parts are converted, which parts must exist, and
# whether fonts are swapped all depend on the package format.
FORMATS = {
    ".pptx": {
        "label": "PowerPoint",
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
# WordprocessingML colours: <w:color w:val="RRGGBB"/> for text, and the fill
# and color attributes used by shading and borders.
WORD_COLOR_RE = re.compile(r'(<w:color\b[^>]*?\bw:val=")([0-9A-Fa-f]{6})(")')
WORD_ATTR_RE = re.compile(r'(\bw:(?:fill|color)=")([0-9A-Fa-f]{6})(")')
# WordprocessingML fonts: run fonts and the font table.
WORD_FONT_RE = re.compile(r'(\bw:(?:ascii|hAnsi|cs|eastAsia|name)=")([^"]*)(")')
# Text content between tags. Vocabulary swaps are confined to these spans.
TEXT_RE = re.compile(r">([^<]+)<")


def term_pattern(terms: list[str]) -> re.Pattern[str] | None:
    """Match any of the given terms as whole words, longest first."""
    if not terms:
        return None
    ordered = sorted(terms, key=len, reverse=True)
    alternatives = "|".join(re.escape(escape(term)) for term in ordered)
    return re.compile(rf"(?<![0-9A-Za-z])(?:{alternatives})(?![0-9A-Za-z])")


def count_terms(text: str, pattern: re.Pattern[str] | None) -> int:
    if pattern is None:
        return 0
    return sum(len(pattern.findall(span)) for span in TEXT_RE.findall(text))


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
    for source in term_sources:
        for target in term_targets:
            if source in target or target in source:
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


def theme_inventory(path: Path, profile: dict) -> tuple[int, int]:
    """Count mapped Acme and Violet values in a package."""
    acme_colours = set(COLOURS)
    violet_colours = set(COLOURS.values())
    acme_fonts = set(FONTS) if profile["swap_fonts"] else set()
    violet_fonts = set(FONTS.values()) if profile["swap_fonts"] else set()
    acme_terms = term_pattern(list(TERMS))
    violet_terms = term_pattern(list(TERMS.values()))
    acme_count = 0
    violet_count = 0

    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if not is_convertible_part(name, profile):
                continue
            text = package.read(name).decode("utf-8")
            for colour in SRGB_RE.findall(text):
                value = colour[1].upper()
                acme_count += value in acme_colours
                violet_count += value in violet_colours
            for colour in ARGB_RE.findall(text):
                value = colour[2].upper()
                acme_count += value in acme_colours
                violet_count += value in violet_colours
            for regex in (WORD_COLOR_RE, WORD_ATTR_RE):
                for colour in regex.findall(text):
                    value = colour[1].upper()
                    acme_count += value in acme_colours
                    violet_count += value in violet_colours
            for font in TYPEFACE_RE.findall(text):
                acme_count += font in acme_fonts
                violet_count += font in violet_fonts
            for font in WORD_FONT_RE.findall(text):
                acme_count += font[1] in acme_fonts
                violet_count += font[1] in violet_fonts
            for font in acme_fonts:
                acme_count += text.count(f"<vt:lpstr>{font}</vt:lpstr>")
            for font in violet_fonts:
                violet_count += text.count(f"<vt:lpstr>{font}</vt:lpstr>")
            # Font names listed in document properties also appear as text,
            # so vocabulary is counted only after those entries are removed.
            for font in acme_fonts | violet_fonts:
                text = text.replace(f"<vt:lpstr>{font}</vt:lpstr>", "")
            acme_count += count_terms(text, acme_terms)
            violet_count += count_terms(text, violet_terms)

    return acme_count, violet_count


def validate_theme_state(path: Path, profile: dict) -> None:
    """Reject mixed inputs because collapsing both sides is not reversible."""
    acme_count, violet_count = theme_inventory(path, profile)
    if acme_count and violet_count:
        raise ValueError(
            "Input contains mapped values from both Acme and Violet; "
            "conversion would not be reversible"
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
        text = TYPEFACE_RE.sub(replace_typeface, text)
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
) -> tuple[Path, int, int, int, int]:
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

    try:
        with zipfile.ZipFile(source) as input_package, zipfile.ZipFile(
            temporary, "w", zipfile.ZIP_DEFLATED
        ) as output_package:
            for item in input_package.infolist():
                data = input_package.read(item.filename)
                if is_convertible_part(item.filename, profile):
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

    return destination, converted_parts, colour_changes, font_changes, term_changes


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Source .pptx, .docx or .xlsx file")
    parser.add_argument("output", help="Destination file, same extension as input")
    parser.add_argument("--to", choices=("acme", "violet"), required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output, parts, colours, fonts, terms = swap(args.input, args.output, args.to)
    print(f"{args.to.title()} theme conversion complete")
    print(f"XML parts changed: {parts}")
    print(f"Colour references changed: {colours}")
    print(f"Font references changed: {fonts}")
    print(f"Text references changed: {terms}")
    print(f"Wrote: {output}")


if __name__ == "__main__":
    main()
