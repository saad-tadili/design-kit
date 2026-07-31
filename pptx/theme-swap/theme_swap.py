#!/usr/bin/env python3
"""Swap a .pptx between theme A and theme B.

    python theme_swap.py in.pptx out.pptx --to a
    python theme_swap.py in.pptx out.pptx --to b
"""
import re
import sys
import zipfile

COLOURS = {
    "003946": "241733",
    "71777A": "6E6A7C",
    "FFCB05": "C9A6FF",
    "FEBE10": "A100FF",
    "FC9E2D": "7A1FA2",
    "5BB54B": "2F8F6B",
    "FFF8E0": "F4EFFA",
    "FFF4C2": "EADFF7",
    "E0DCD1": "DCD6E4",
    "F0E9E0": "EFEBF4",
    "A59E8C": "9C95A8",
    "EAEBEB": "E5E2EA",
    "DDE6E1": "E3DEEC",
}

FONTS = {
    "Acme New Display": "Franklin Gothic Medium",
    "Acme New Text": "Franklin Gothic Book",
}

PARTS = ("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/",
         "ppt/theme/", "ppt/notesSlides/", "docProps/")


def build(theme):
    if theme == "a":
        return dict(COLOURS), dict(FONTS)
    return ({v: k for k, v in COLOURS.items()},
            {v: k for k, v in FONTS.items()})


def convert(text, cmap, fmap):
    text = re.sub(r'(srgbClr\s+)val="([0-9A-Fa-f]{6})"',
                  lambda m: f'{m.group(1)}val="{cmap.get(m.group(2).upper(), m.group(2))}"',
                  text)
    text = re.sub(r'typeface="([^"]*)"',
                  lambda m: f'typeface="{fmap.get(m.group(1), m.group(1))}"',
                  text)
    for a, b in fmap.items():
        text = text.replace(f"<vt:lpstr>{a}</vt:lpstr>", f"<vt:lpstr>{b}</vt:lpstr>")
    return text


def swap(src, dst, theme):
    cmap, fmap = build(theme)
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") and item.filename.startswith(PARTS):
                data = convert(data.decode("utf8"), cmap, fmap).encode("utf8")
            zout.writestr(item, data)
    zin.close()
    return dst


if __name__ == "__main__":
    # both sides must stay unique or the swap is not reversible
    assert not (set(COLOURS) & set(COLOURS.values()))
    assert not (set(FONTS) & set(FONTS.values()))
    if len(sys.argv) != 5 or sys.argv[3] != "--to" or sys.argv[4] not in ("a", "b"):
        print(__doc__)
        sys.exit(1)
    print(f"wrote {swap(sys.argv[1], sys.argv[2], sys.argv[4])}")
