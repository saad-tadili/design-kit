# theme-swap

Converts a `.pptx` between two colour and font themes. Theme A uses the values
on the left side of the mappings in `theme_swap.py`; Theme B uses the values on
the right.

The mappings are one-to-one and reversible. Converting from one theme to the
other and back restores the original uncompressed PowerPoint package contents.
The `.pptx` file itself may not be byte-identical because its ZIP container is
recompressed.

## Usage

```text
python theme_swap.py in.pptx out.pptx --to b
python theme_swap.py out.pptx restored.pptx --to a
```

Use different input and output paths. The script rejects in-place conversion
and writes through a temporary file so a failed conversion cannot damage an
existing presentation.

Python 3.11 or later. Standard library only, nothing to install.

## What it changes

The script converts configured colours and exact font names in:

- slides, layouts and masters;
- presentation themes and table styles;
- chart and SmartArt XML;
- notes and handout masters;
- document properties.

Layout, text, geometry and embedded images are unchanged. Code colours remain
unchanged unless they are explicitly added to the mapping.

After conversion, the script validates the PowerPoint package and reports how
many XML parts, colour references and font references changed.

## Configuring

Edit `COLOURS` and `FONTS` at the top of the script. Each `COLOURS` entry maps a
Theme A value to its Theme B equivalent. Each `FONTS` entry follows the same
direction. `--to b` applies the mappings as written; `--to a` reverses them.

Font names are matched exactly as PowerPoint stores them. Both names must be
spelled precisely; a near-match is deliberately left unchanged.

Every source and target value must remain unique, and the two sides must remain
disjoint. The script validates these constraints on every run. Shared neutral
colours such as white and black belong in `UNCHANGED_COLOURS`, not in the
reversible mapping.

An input containing mapped values from both themes is rejected because merging
the two sides would not be reversible. Normalize the presentation to one theme
before converting it.

## Notes

The theme carries font names only. Rendering a theme as intended still requires
those fonts to be installed on the machine.

Embedded images are copied unchanged. If an image must differ between themes,
provide and replace it separately.

