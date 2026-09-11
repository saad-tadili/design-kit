#!/usr/bin/env python3
"""Swap a .pptx or .xlsx between the Acme and Violet themes.

    python theme_swap.py in.pptx out.pptx --to violet
    python theme_swap.py in.xlsx out.xlsx --to acme

The file type is detected from the input extension. PowerPoint packages get
colour and font swaps; Excel packages get colour swaps only, since workbook
fonts are not theme fonts and must not change.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import re
import tempfile
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

# Both brand font families are Acme-side sources: decks built with the real
# names and decks built with the DLP-safe placeholder names both convert.
# Each source keeps its own distinct Violet target so the reverse conversion
# restores exactly the name the deck started with.
FONTS = {
    "Sun Life New Display": "Franklin Gothic Medium",
    "Sun Life New Text": "Franklin Gothic Book",
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


def build(target: str, swap_fonts: bool) -> tuple[dict[str, str], dict[str, str]]:
    """Return the colour and font mappings required for the target theme."""
    if target == "violet":
        colour_map = dict(COLOURS)
        font_map = dict(FONTS)
    else:
        colour_map = {value: key for key, value in COLOURS.items()}
        font_map = {value: key for key, value in FONTS.items()}
    if not swap_fonts:
        font_map = {}
    return colour_map, font_map


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
            for font in TYPEFACE_RE.findall(text):
                acme_count += font in acme_fonts
                violet_count += font in violet_fonts
            for font in acme_fonts:
                acme_count += text.count(f"<vt:lpstr>{font}</vt:lpstr>")
            for font in violet_fonts:
                violet_count += text.count(f"<vt:lpstr>{font}</vt:lpstr>")

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
    text: str, colour_map: dict[str, str], font_map: dict[str, str]
) -> tuple[str, int, int]:
    """Convert theme values in one XML part and return replacement counts."""
    colour_changes = 0
    font_changes = 0

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

    def replace_typeface(match: re.Match[str]) -> str:
        nonlocal font_changes
        replacement = font_map.get(match.group(1))
        if replacement is None:
            return match.group(0)
        font_changes += 1
        return f'typeface="{replacement}"'

    text = SRGB_RE.sub(replace_srgb, text)
    text = ARGB_RE.sub(replace_argb, text)
    if font_map:
        text = TYPEFACE_RE.sub(replace_typeface, text)
        for original, replacement in font_map.items():
            source = f"<vt:lpstr>{original}</vt:lpstr>"
            occurrences = text.count(source)
            if occurrences:
                text = text.replace(
                    source, f"<vt:lpstr>{replacement}</vt:lpstr>"
                )
                font_changes += occurrences

    return text, colour_changes, font_changes


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


def swap(src: str | Path, dst: str | Path, target: str) -> tuple[Path, int, int, int]:
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
    colour_map, font_map = build(target, profile["swap_fonts"])

    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.name}.", suffix=".tmp", dir=destination.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)

    converted_parts = 0
    colour_changes = 0
    font_changes = 0

    try:
        with zipfile.ZipFile(source) as input_package, zipfile.ZipFile(
            temporary, "w", zipfile.ZIP_DEFLATED
        ) as output_package:
            for item in input_package.infolist():
                data = input_package.read(item.filename)
                if is_convertible_part(item.filename, profile):
                    text, colours, fonts = convert(
                        data.decode("utf-8"), colour_map, font_map
                    )
                    data = text.encode("utf-8")
                    if colours or fonts:
                        converted_parts += 1
                        colour_changes += colours
                        font_changes += fonts
                output_package.writestr(item, data)

        validate_package(temporary, profile)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)

    return destination, converted_parts, colour_changes, font_changes


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Source .pptx or .xlsx file")
    parser.add_argument("output", help="Destination file, same extension as input")
    parser.add_argument("--to", choices=("acme", "violet"), required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output, parts, colours, fonts = swap(args.input, args.output, args.to)
    print(f"{args.to.title()} theme conversion complete")
    print(f"XML parts changed: {parts}")
    print(f"Colour references changed: {colours}")
    print(f"Font references changed: {fonts}")
    print(f"Wrote: {output}")


if __name__ == "__main__":
    main()
