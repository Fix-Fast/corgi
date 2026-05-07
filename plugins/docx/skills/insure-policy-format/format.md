# Format Rules

These rules describe how the finished document should look.

## Document parts

These insurance-policy documents usually contain:

- a running header
- a title
- a notices block
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
uppercase Roman numerals at the top level of a section:

```
I. Foo
  1) bar
  2) baz
II. Qux
```

Rule 0 rewrites such outlines to the canonical sequence
(`A.` -> `1.` -> `a.` -> `1.` -> `a.` -> `i.`). The example above
becomes:

```
A. Foo
  1. bar
  2. baz
B. Qux
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
- notices block
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

- `Bricolage Grotesque ExtraBold` (the heavy-weight font face, not a regular face with the bold attribute)
- `23pt`
- centered
- `10pt` space after

Section heading styling:

- `Bricolage Grotesque`
- `14pt`
- bold
- left-aligned
- `16pt` space before
- `8pt` space after

The section heading formatting is carried by the `Heading 2` style
definition in `styles.xml`, not by run-level character overrides. Rule 1
injects/overwrites the `Heading 2` style def to match these values and
leaves the runs themselves bare.

Notices block styling:

- `Inter` (inherits from doc defaults)
- `13pt`
- bold
- black text
- left-aligned
- `10pt` space after
- a hard page break is appended inside the LAST paragraph in the
  block, so the body of the policy starts on a fresh page

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

Subheading styling:

- `Bricolage Grotesque`
- `13pt`
- bold
- left-aligned
- `10pt` space after

Body text styling:

- `Inter`
- `11pt`
- black text
- left-aligned
- `10pt` space after

Body styling is also installed at the OOXML doc-default level
(`<w:docDefaults><w:rPrDefault>`) so anything that doesn't override —
notably the list markers, whose canonical level definitions omit
run formatting — inherits Inter / 11pt / black.

Script:

- `rule_1.py`

## Rule 2: Lists

List content should appear as a native multilevel list.

Marker sequence (canonical):

- `A.`
- `1.`
- `a.`
- `1.`
- `a.`
- `i.`

The ladder is period-delimited and cycles `decimal -> lowerLetter`
between levels — list counters reset under their parent so two `1.`
items at L1 and L3 don't conflict in practice.

Recognized authoring forms (any of these typed as the leading marker
on a paragraph become the canonical native list item at the
corresponding level):

- L0: `A.` or `A)`
- L1: `1.` or `1)`
- L2: `a.` or `a)`
- L3: `1.` or `1)` or `(1)`
- L4: `a.` or `a)` or `(a)`
- L5: `i.` or `i)` or `(i)`

A marker at the top three levels may end in either a closing
parenthesis `)` or a period `.`. Both forms mean the same thing. The
parenthesized forms (`(1)`, `(a)`, `(i)`) are matched only with
parentheses on both sides.

For example, all of these are recognized as the same kind of list item:

- `A. Allocation: ...`
- `A) Allocation: ...`
- `1. Defense Costs: ...`
- `1) Defense Costs: ...`

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

A. Allocation: ...

   1. Defense Costs: The covered portion of Defense Costs incurred in
      defending a Demand will be considered covered Loss, ...

      Notwithstanding the foregoing, we may, at our discretion, allocate
      Defense Costs between covered and uncovered matters.
```

The continuation paragraph aligns with the body text of its parent item,
not with the page margin or the marker column.

List counters restart at every section heading (Heading 2). Two list
items at the same level under different sections are numbered
independently — `A.` under SECTION II is unrelated to `A.` under
SECTION III.

The list should use these level formats:

- level 0: upper-alpha with `.`
- level 1: decimal with `.`
- level 2: lower-alpha with `.`
- level 3: decimal with `.`
- level 4: lower-alpha with `.`
- level 5: lower-roman with `.`

The list should use this indentation ladder:

- level 0: left indent `720 twips`, hanging indent `360 twips`
- level 1: left indent `1440 twips`, hanging indent `360 twips`
- level 2: left indent `2160 twips`, hanging indent `360 twips`
- level 3: left indent `2880 twips`, hanging indent `360 twips`
- level 4: left indent `3600 twips`, hanging indent `360 twips`
- level 5: left indent `4320 twips`, hanging indent `360 twips`

The separator between the list marker and body text is a `tab` (the
OOXML default — the canonical list definition omits `<w:suff>` so Word
falls back to tab). The hanging indent ladder above places the tab stop
at the body-text column so wrapped lines align under the first character
of body text.

Paragraph-level indentation overrides should be removed from list
paragraphs.

List markers inherit their formatting from the paragraph's body style
(Inter, 11pt, black). The canonical list definition emits no run-level
overrides on level entries; future body-style changes carry the markers
along.

Script:

- `rule_2.py`

## Rule 3: Running header and page layout

Every section should use:

- top margin `1.0"`
- bottom margin `1.0"`
- left margin `1.0"`
- right margin `1.0"`
- header distance `0.5"`

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
