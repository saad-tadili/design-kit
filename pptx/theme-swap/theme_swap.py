#!/usr/bin/env python3
"""Swap a .pptx between theme A and theme B.

    python theme_swap.py in.pptx out.pptx --to a
    python theme_swap.py in.pptx out.pptx --to b
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import re
import tempfile
import zipfile


# Theme A -> Theme B. Both sides must remain unique and disjoint so the
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

FONTS = {
    "Acme New Display": "Franklin Gothic Medium",
    "Acme New Text": "Franklin Gothic Book",
}

PART_PREFIXES = (
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
)
PART_FILES = frozenset({"ppt/tableStyles.xml"})
REQUIRED_PARTS = frozenset({"[Content_Types].xml", "ppt/presentation.xml"})

SRGB_RE = re.compile(r'(srgbClr\s+)val="([0-9A-Fa-f]{6})"')
TYPEFACE_RE = re.compile(r'typeface="([^"]*)"')


def validate_mapping() -> None:
    """Ensure each direction is deterministic and reversible."""
    colour_sources = set(COLOURS)
    colour_targets = set(COLOURS.values())
    font_sources = set(FONTS)
    font_targets = set(FONTS.values())

    if len(colour_targets) != len(COLOURS):
        raise ValueError("Theme B colour values must be unique")
    if colour_sources & colour_targets:
        raise ValueError("Theme A and Theme B colour values must be disjoint")
    if UNCHANGED_COLOURS & (colour_sources | colour_targets):
        raise ValueError("Unchanged colours must not appear in the mapping")
    if len(font_targets) != len(FONTS):
        raise ValueError("Theme B font values must be unique")
    if font_sources & font_targets:
        raise ValueError("Theme A and Theme B font values must be disjoint")


def build(target: str) -> tuple[dict[str, str], dict[str, str]]:
    """Return the colour and font mappings required for the target theme."""
    if target == "b":
        return dict(COLOURS), dict(FONTS)
    return (
        {value: key for key, value in COLOURS.items()},
        {value: key for key, value in FONTS.items()},
    )


def is_convertible_part(name: str) -> bool:
    return name.endswith(".xml") and (
        name in PART_FILES or name.startswith(PART_PREFIXES)
    )


def theme_inventory(path: Path) -> tuple[int, int]:
    """Count mapped Theme A and Theme B values in a package."""
    theme_a_colours = set(COLOURS)
    theme_b_colours = set(COLOURS.values())
    theme_a_fonts = set(FONTS)
    theme_b_fonts = set(FONTS.values())
    theme_a_count = 0
    theme_b_count = 0

    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if not is_convertible_part(name):
                continue
            text = package.read(name).decode("utf-8")
            for colour in SRGB_RE.findall(text):
                value = colour[1].upper()
                theme_a_count += value in theme_a_colours
                theme_b_count += value in theme_b_colours
            for font in TYPEFACE_RE.findall(text):
                theme_a_count += font in theme_a_fonts
                theme_b_count += font in theme_b_fonts
            for font in theme_a_fonts:
                theme_a_count += text.count(f"<vt:lpstr>{font}</vt:lpstr>")
            for font in theme_b_fonts:
                theme_b_count += text.count(f"<vt:lpstr>{font}</vt:lpstr>")

    return theme_a_count, theme_b_count


def validate_theme_state(path: Path) -> None:
    """Reject mixed inputs because collapsing both sides is not reversible."""
    theme_a_count, theme_b_count = theme_inventory(path)
    if theme_a_count and theme_b_count:
        raise ValueError(
            "Input contains mapped values from both Theme A and Theme B; "
            "conversion would not be reversible"
        )


def convert(
    text: str, colour_map: dict[str, str], font_map: dict[str, str]
) -> tuple[str, int, int]:
    """Convert theme values in one XML part and return replacement counts."""
    colour_changes = 0
    font_changes = 0

    def replace_colour(match: re.Match[str]) -> str:
        nonlocal colour_changes
        original = match.group(2)
        replacement = colour_map.get(original.upper())
        if replacement is None:
            return match.group(0)
        colour_changes += 1
        return f'{match.group(1)}val="{replacement}"'

    def replace_typeface(match: re.Match[str]) -> str:
        nonlocal font_changes
        original = match.group(1)
        replacement = font_map.get(original)
        if replacement is None:
            return match.group(0)
        font_changes += 1
        return f'typeface="{replacement}"'

    text = SRGB_RE.sub(replace_colour, text)
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


def validate_package(path: Path) -> None:
    """Confirm the result is a readable PowerPoint package."""
    if not zipfile.is_zipfile(path):
        raise ValueError(f"Not a valid .pptx package: {path}")

    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        missing = REQUIRED_PARTS - names
        if missing:
            raise ValueError(
                "PowerPoint package is missing: " + ", ".join(sorted(missing))
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

    validate_mapping()
    validate_package(source)
    validate_theme_state(source)
    colour_map, font_map = build(target)

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
                if is_convertible_part(item.filename):
                    text, colours, fonts = convert(
                        data.decode("utf-8"), colour_map, font_map
                    )
                    data = text.encode("utf-8")
                    if colours or fonts:
                        converted_parts += 1
                        colour_changes += colours
                        font_changes += fonts
                output_package.writestr(item, data)

        validate_package(temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)

    return destination, converted_parts, colour_changes, font_changes


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Source .pptx file")
    parser.add_argument("output", help="Destination .pptx file")
    parser.add_argument("--to", choices=("a", "b"), required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output, parts, colours, fonts = swap(args.input, args.output, args.to)
    print(f"Theme {args.to.upper()} conversion complete")
    print(f"XML parts changed: {parts}")
    print(f"Colour references changed: {colours}")
    print(f"Font references changed: {fonts}")
    print(f"Wrote: {output}")


if __name__ == "__main__":
    main()
