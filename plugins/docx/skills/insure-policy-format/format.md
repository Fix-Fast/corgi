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

## Rule 0: Outline normalization (pre-pass)

Some source documents use non-canonical outline markers — for example,
uppercase letters at the top level of a section:

```
A. Foo
  1) bar
  2) baz
B. Qux
```

Rule 0 rewrites such outlines to the canonical sequence
(`1)` -> `a)` -> `i)` -> `(1)` -> `(a)` -> `(i)`). The example above
becomes:

```
1) Foo
  a) bar
  b) baz
2) Qux
```

The rewrite is content-level, not styling. After Rule 0 finishes, the
document looks as if it had been authored in canonical form, so
Rules 1/2/3 see only canonical markers and stay commutative with each
other.

Rule 0 only fires for sections that are explicitly flagged as needing
normalization. The flagging mechanism, target-section identification,
and source-marker description live in the skill's manifest (see
`SKILL.md`), not here.

Script:

- `rule_0.py`

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

- `1)` or `1.`
- `a)` or `a.`
- `i)` or `i.`
- `(1)`
- `(a)`
- `(i)`

A marker at the top three levels may end in either a closing
parenthesis `)` or a period `.`. Both forms mean the same thing. The
parenthesized forms (`(1)`, `(a)`, `(i)`) are matched only with
parentheses on both sides.

For example, all of these are recognized as the same kind of list item:

- `1) Allocation: ...`
- `1. Allocation: ...`
- `a) Defense Costs: ...`
- `a. Defense Costs: ...`

If list content is still written as typed markers in plain text, it
should be converted into a native list.

A paragraph is a list item only when its **leading** text is a list
marker. Markers embedded mid-paragraph are left as plain text — this
matches Word's native behavior (typing `1) Foo and 2) Bar` mid-list
gives you one item, not two). Authors who want separate items must
write them as separate paragraphs in the source.

A piece of text only counts as a leading marker if it stands on its
own — at the start of the paragraph, separated from the body text by a
space (or a tab to be normalized to a space). Citations like
`officer(s)`, `§4958(c)`, and `Section 4958(a)(2)` are left alone
because the `(s)`, `(c)`, and `(a)` are not in leading position.

If a paragraph continues a list item, it should stay attached to that
list item.

```
SECTION II: INSURING AGREEMENTS

1) Allocation: ...

   a) Defense Costs: The covered portion of Defense Costs incurred in
      defending a Demand will be considered covered Loss, ...

      Notwithstanding the foregoing, we may, at our discretion, allocate
      Defense Costs between covered and uncovered matters.
```

The continuation paragraph aligns with the body text of its parent item,
not with the page margin or the marker column.

List counters restart at every section heading (Heading 2). Two list
items at the same level under different sections are numbered
independently — `1)` under SECTION II is unrelated to `1)` under
SECTION III.

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

1. `rule_0.py` (only fires for sections listed in `outline_normalizations`)
2. `rule_1.py`
3. `rule_2.py`
4. `rule_3.py`

Each rule is `(doc, parts) -> doc`: it mutates the OOXML tree in place.
Downstream rules read upstream effects from the doc itself — Rule 2
reads `pStyle` to know which paragraphs are headings, Rule 3 reads
`sectPr` to find sections. No out-of-band state is passed between
rules.
