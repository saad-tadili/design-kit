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

Hi all,

Following the BG feedback on our AI-Ready Data Assessment, we would like to bring more business context to the findings and connect them to Asia’s priorities and intended AI use.

The material shared outlines five AI transformation pillars:

* Advisor enablement
* Interactive client portal
* Client Contact Centre
* Digital Onboarding and Policy Issuance
* Claims automation

To make that connection, could you please identify **up to two priority AI initiatives within each pillar most critical to Asia’s near-term business commitments**? Please focus on specific business use cases planned or underway, indicating the relevant market or entity. There is no need to cover every pillar or provide a full portfolio inventory.

For each selected initiative, please share any existing material describing:

* **Business purpose and function:** intended users, the problem addressed and expected outcome.
* **What AI does:** what it receives, produces, recommends or executes, including where human review or approval occurs.
* **Main data and content needed:** key datasets, documents and source systems or repositories.
* **Delivery stage and timing:** current stage, next milestone and planned deployment scope.

This will help us identify the relevant questionnaire capabilities and the stakes of each intended AI use, so we can explain what the assessment findings mean for Asia’s priorities.

**There is no required format.** Existing presentations, initiative summaries, business cases, links or brief answers are welcome. Please share what already exists rather than creating or reformatting material.

**Requested by: Friday, October 2 (EOD)** to support completion of the assessment and preparation of enterprise-level findings.

Thank you for your support.

