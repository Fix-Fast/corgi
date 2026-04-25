# Format Rules

These rules describe how the finished document should look.

## Document parts

These insurance-policy documents usually contain:

- a running header
- a title
- section headings
- subheadings
- coverage headings or insuring-agreement headings
- body text
- legal list items

Coverage and insuring-agreement headings do not need to appear in a
fixed sequence. For example, a document may contain only `Coverage B`
without also containing `Coverage A` or `Coverage C`.

## Rule 1: Text hierarchy

The document should have a clear text hierarchy:

- title
- section heading
- subheading
- coverage heading or insuring-agreement heading
- body text

These should appear as real document structure, not just visual
formatting.

Structure:

- title -> `Heading 1`
- section heading -> `Heading 2`
- subheading, coverage heading, or insuring-agreement heading -> `Heading 3`
- body text -> ordinary paragraph

Running-header content should not be treated as body text.

Title styling:

- `Bricolage Grotesque`
- `26pt`
- bold
- centered
- `18pt` space after

Section heading styling:

- `Bricolage Grotesque`
- `14pt`
- bold
- left-aligned
- `16pt` space before
- `8pt` space after

Subheading styling:

- `Bricolage Grotesque`
- `12pt`
- bold
- left-aligned
- `10pt` space before
- `6pt` space after

Body text styling:

- `Inter`
- `11pt`
- black text
- left-aligned
- `6pt` space after

Script:

- `rule_1.py`

## Rule 2: Lists

List content should appear as a native multilevel list.

Marker sequence:

- `1)`
- `a)`
- `i)`
- `(1)`
- `(a)`
- `(i)`

If list content is still written as typed markers in plain text, it
should be converted into a native list.

If a paragraph contains multiple embedded list markers, it should be
split into separate list paragraphs.

If a paragraph continues a list item, it should stay attached to that
list item.

The list should use these level formats:

- level 0: decimal with `)`
- level 1: lower-alpha with `)`
- level 2: lower-roman with `)`
- level 3: decimal with surrounding parentheses
- level 4: lower-alpha with surrounding parentheses
- level 5: lower-roman with surrounding parentheses

The list should use this indentation ladder:

- level 0: left indent `720 twips`, hanging indent `360 twips`
- level 1: left indent `1440 twips`, hanging indent `360 twips`
- level 2: left indent `2160 twips`, hanging indent `360 twips`
- level 3: left indent `2880 twips`, hanging indent `360 twips`
- level 4: left indent `3600 twips`, hanging indent `360 twips`
- level 5: left indent `4320 twips`, hanging indent `360 twips`

The space between the list marker and body text should be a single
`space`, not a tab.

Paragraph-level indentation overrides should be removed from list
paragraphs.

List markers should be:

- `Inter`
- `11pt`
- black text

Script:

- `rule_2.py`

## Rule 3: Running header and page layout

Every section should use:

- top margin `0.7"`
- bottom margin `0.7"`
- left margin `1.0"`
- right margin `1.0"`
- header distance `0.35"`

The running header should be:

- `<Title><TAB><Policy Code>`

Example:

- `Commercial General Liability Policy<TAB>CORGI-TECH-1234`

If only one of these values is available, the header should still use
the available value.

The header paragraph should be:

- left-aligned
- with a right-aligned tab stop at `6.3"`
- in `Inter`
- `10pt`
- gray text (`RGB 128,128,128`)

Script:

- `rule_3.py`

## Full Formatter

The full formatter is the composition of:

1. `rule_1.py`
2. `rule_3.py`
3. `rule_2.py`
