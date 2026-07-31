# theme-swap

Converts a `.pptx` between two colour and font themes. The mapping is
one-to-one in both directions, so the conversion is lossless and can be run any
number of times: converting to A and back to B reproduces the original file
byte for byte.

## Usage

```
python theme_swap.py in.pptx out.pptx --to a
python theme_swap.py in.pptx out.pptx --to b
```

Python 3.6 or later. Standard library only, nothing to install.

## What it changes

Panel fills, accent colours, text colours and font names, across slides,
layouts, masters, theme and document properties. Layout, text and geometry are
untouched.

Code blocks are out of scope by design: the dark panel background and the syntax
colours inside it are the same in both themes, so code always reads the same
way. `F1F7F7` is also left alone, because a deck may use that one value both as
a pale panel fill and as code body text.

## Configuring

Edit `COLOURS` and `FONTS` at the top of the script. Fonts are matched on their
exact name as PowerPoint stores it, so both sides must be spelled precisely —
a near-miss silently changes nothing.

Both sides of every row must stay unique. If a value appeared as both a source
and a target the swap would no longer be reversible; the script asserts this on
every run.

## Notes

The theme carries font names only. Rendering a theme as intended still requires
those fonts installed on the machine.
