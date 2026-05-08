# Format Rules

These rules describe how a finished document for a Corgi Insurance Policy should look.

## Rule 1: Text hierarchy

The document should have a clear text hierarchy:

- a running header
- a title
- a notices block
- section headings
- subheadings (including coverage and insuring-agreement headings)
- body text
- legal list items

> Note: these documents may also have content that should be ignored (paragraphs that exist in the source but should not appear in the output — e.g. running-header content baked into the body, or duplicated title fragments). These paragraphs are removed from the output entirely.

These should appear as real document structure, not just visual formatting.

Structure:

- title -> `Heading 1`
- section heading -> `Heading 2`
- subheading, coverage heading, or insuring-agreement heading -> `Heading 3`
- body text -> ordinary paragraph

Running-header content should not be treated as body text.

### Title styling:

- `Bricolage Grotesque ExtraBold` (the heavy-weight font face, not a regular face with the bold attribute)
- `23pt`
- centered
- `10pt` space after

### Section heading styling:

- `Bricolage Grotesque`
- `14pt`
- bold
- left-aligned
- `16pt` space before
- `8pt` space after

### Notices block styling:

- `Inter` (inherits from doc defaults)
- `13pt`
- bold
- black text
- left-aligned
- `10pt` space after
- a hard page break is appended inside the LAST paragraph in the
  block, so the body of the policy starts on a fresh page
- if the source follows the notices block with its own empty
  page-break paragraph(s), those redundant breaks are stripped so the
  output renders one page break, not two (no blank pages)

The notices block sits between the title and the first section heading
and typically contains a regulatory disclosure paragraph (e.g., a
risk-retention-group notice) plus the coverage-form-type paragraph
(e.g., `THIS IS A "CLAIMS-MADE" LIABILITY COVERAGE FORM...` or
`THIS IS AN OCCURRENCE-BASED LIABILITY COVERAGE FORM...`).

The `NOTICES:` or `IMPORTANT NOTICE:` subheading itself is styled per
the subheading rules below; this rule covers only the prose paragraphs
that follow it.

Example (Directors & Officers Liability):

```
IMPORTANT NOTICE:

This Policy is issued by a Risk Retention Group (RRG). A Risk
Retention Group is a state-chartered insurance company that
enjoys certain federal preemptions under the Liability Risk
Retention Act (15 U.S.C. §3901 et seq.). As such, it is not
subject to all the insurance laws and regulations of your state.

THIS IS A "CLAIMS-MADE" LIABILITY COVERAGE FORM. This Policy
provides coverage for Claims first made against an Insured
during the Policy Period (or any applicable Extended Reporting
Period). Defense Costs reduce the Limit of Liability (unless
otherwise stated) and may be applied against the Retention.
Please read the entire Policy carefully.

───── page break ─────
```

### Subheading styling:

- `Bricolage Grotesque`
- `13pt`
- bold
- left-aligned
- `10pt` space after

### Body text styling:

- `Inter`
- `11pt`
- black text
- left-aligned
- `10pt` space after

## Rule 2: Lists

List content should appear as a native multilevel list.

### Marker ladder

Each outline level uses a single, distinct marker form. There is no
sharing of markers across levels and no alternate forms.

- L0: `A.`
- L1: `1.`
- L2: `a.`
- L3: `(1)`
- L4: `(a)`
- L5: `(i)`

Elsewhere in this document, this is referred to as the **standard list
format**. The top three levels are period-suffixed; the bottom three are
parenthesized on both sides. Because each level has its own unique
form, there is never any ambiguity about which level a marker belongs
to.

A paragraph is a list item only when its **leading** text is one of the
six marker forms above, separated from the body text by a space (a tab
is normalized to a space). Markers embedded mid-paragraph are left as
plain text — this matches Word's native behavior (typing `1. Foo and 2.
Bar` mid-list gives you one item, not two).

Citations and parentheticals are never treated as markers, because
they don't appear in leading position: `officer(s)`, `§4958(c)`,
`Section 4958(a)(2)`, `Section IV.A.`, and `sixty (60) days` all stay
as plain text.

List counters restart at every section heading (Heading 2). Two list
items at the same level under different sections are numbered
independently — `A.` under SECTION II is unrelated to `A.` under
SECTION III.

### Outline marker normalization (pre-pass)

Some source documents use outline markers that don't match the
standard list format — for example, uppercase Roman numerals at the
top level of a section:

```
I. Foo
  1) bar
  2) baz
II. Qux
  1) baz-bar
```

These are rewritten to match the standard list format. The example
becomes:

```
A. Foo
  1. bar
  2. baz
B. Qux
  1. baz-bar
```

(Note that the L1 counter restarts at `1.` under each new top-level
item.)

The rewrite is content-level, not styling: after it finishes, the
document looks as if it had been authored using the standard list
format.

### List structure

If list content is still written as typed markers in plain text, it
should be converted into a native list.

If a paragraph continues a list item, it should stay attached to that
list item. Continuation behavior persists across consecutive unmarked
paragraphs: every body paragraph that follows a list item, until a new
list marker or heading appears, is indented to that item's body-text
column.

```
SECTION II: INSURING AGREEMENTS

A. Allocation: ...

   1. Defense Costs: The covered portion of Defense Costs incurred in
      defending a Demand will be considered covered Loss, ...

      Notwithstanding the foregoing, we may, at our discretion, allocate
      Defense Costs between covered and uncovered matters.
```

The continuation paragraph aligns with the body text of its parent
item, not with the page margin or the marker column.

### Indentation

The list should use this indentation ladder:

- level 0: left indent `720 twips`, hanging indent `360 twips`
- level 1: left indent `1440 twips`, hanging indent `360 twips`
- level 2: left indent `2160 twips`, hanging indent `360 twips`
- level 3: left indent `2880 twips`, hanging indent `360 twips`
- level 4: left indent `3600 twips`, hanging indent `360 twips`
- level 5: left indent `4320 twips`, hanging indent `360 twips`

The separator between the list marker and body text is a tab. The
hanging indent ladder above places the tab stop at the body-text
column, so wrapped lines align under the first character of body text.

Paragraph-level indentation overrides should be removed from list
paragraphs.

List markers inherit their formatting from the body style (Inter,
11pt, black).

## Rule 3: Running header and page layout

Every section should use:

- top margin `1.0"`
- bottom margin `1.0"`
- left margin `1.0"`
- right margin `1.0"`
- header distance `0.5"`
- footer distance `0.5"`

The running header should be:

- `<Title><TAB><Policy Code>`

Example:

- `Commercial General Liability Policy<TAB>CORGI-TECH-1234`

If only one of these values is available, the header should still use
the available value (no tab in that case). If neither is available, the
header is an empty paragraph (still styled per the rules below, just
with no text).

The header paragraph should be:

- left-aligned
- with a right-aligned tab stop at `6.3"`
- in `Inter`
- `10pt`
- gray text (`RGB 128,128,128`)
