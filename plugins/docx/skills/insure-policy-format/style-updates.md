# Style updates: gap analysis vs. user reference

Comparing our formatter output against the user's preferred CGL policy
format (`CGL POLICY (1) (1).docx`, treated here as the reference).

Tested against `insure-policy-format` v0.8.0 (post-OOXML rewrite).
Source input: `CGL POLICY (1).docx`. Output: `CGL POLICY (1) - corgi.docx`.

## Already matching (no change needed)

- **Body font size** — both 11pt. The user reference declares no explicit
  size on 98.5% of body chars; they inherit the doc default of 22
  half-points = 11pt. Same as our `Body` style. The two visible 13pt
  paragraphs at the top are intentional emphasis (see #2 below).
- **Alignment** — both inherit default LEFT. The reference has only 5
  explicit JUSTIFY and 5 explicit CENTER paragraphs out of ~200; ours
  has 1 CENTER (the title). Effectively the same.
- **Inline list-marker splitting** — both keep "Claim means: (1) ...
  (2) ..." and "Coverage Territory means: (1) ... (2) ..." as a single
  paragraph with markers inline. Already addressed in `rule_2.py` lines
  104-107 (parenthesized embedded markers explicitly not split, only
  bare markers).

## Gaps to close (ordered by impact)

### 1. Outline scheme

Our ladder is paren-rooted with decimal at L0:

```
1) ... a) ... i) ... (1) ... (a) ... (i)
```

Reference is period-rooted with upper-letter at L0 (or upper-roman for
endorsements like Hired/Non-Owned Auto):

```
A. ... 1. ... a. ... 1. ... a. ... i.
```

Most-used reference definition (`numbering.xml` abstractNum 9):

| Level | numFmt        | lvlText |
|-------|---------------|---------|
| 0     | upperLetter   | `%1.`   |
| 1     | decimal       | `%2.`   |
| 2     | lowerLetter   | `%3.`   |
| 3     | decimal       | `%4.`   |
| 4     | lowerLetter   | `%5.`   |
| 5     | lowerRoman    | `%6.`   |

A handful of sections (e.g., the Hired/Non-Owned Auto endorsement) use
`upperRoman` at L0 instead. Worth deciding whether the ladder is
section-aware or always upper-letter rooted with the section overriding
explicitly.

Rule 0 (outline normalization) and rule_2 (canonical multilevel
abstractNum) both reflect the current paren-rooted scheme; both would
need updating.

### 2. Intro emphasis paragraphs at 13pt

The two paragraphs immediately after the title (NOTICES body and the
"THIS IS AN OCCURRENCE-BASED LIABILITY COVERAGE FORM..." preamble) are
13pt in the reference. Body elsewhere stays 11pt.

This is a targeted size bump on a known set of paragraphs, not a global
body-size change.

### 3. Body space-after

| | ours | reference |
|---|---|---|
| Body para space-after | 6pt | 10pt |

Reference uses 10pt uniformly across 184 of ~200 paragraphs. Ours uses
6pt for body, 8pt for section headings, 18pt for title. Bumping body
from 6→10pt would close most of the gap.

### 4. Title

| | ours | reference |
|---|---|---|
| Size | 26pt | 23pt |
| Font | Bricolage Grotesque + bold attr | Bricolage Grotesque ExtraBold (separate face) |
| Space-after | 18pt | 10pt |

The font face differs: reference uses the ExtraBold weight as its own
font name, not a regular weight with `bold=true`. If the system has
Bricolage Grotesque ExtraBold installed, the visual result is heavier
than what `bold=true` on the regular face produces.

### 5. Section heading

| | ours | reference |
|---|---|---|
| Size / family | explicit 14pt Bricolage Grotesque, bold | Heading 2 style default (no per-run override) |
| Space before/after | 16pt / 8pt | ~12pt / ~12pt |

Reference relies on the doc-level `Heading 2` style rather than
applying run-level overrides on every section heading. Ours stamps the
font/size on each run.

### 6. Subheading (e.g., NOTICES:)

| | ours | reference |
|---|---|---|
| Size | 12pt | 13pt |
| Space before/after | 10pt / 6pt | 0pt / 10pt |

### 7. Page margins

| | ours | reference |
|---|---|---|
| Top / bottom | 0.7" | 1.0" |
| Left / right | 1.0" | 1.0" |

### 8. Header distance

| | ours | reference |
|---|---|---|
| Header distance | 0.35" | 0.5" |

### 9. Running header content

Ours: `Commercial General Liability Policy` left-aligned, Inter 10pt.
Reference: empty.

This one is policy-specific — `header_title_text` is opt-in via
`parts.json`, so it's already controllable per-run. May not need a
spec change, but worth confirming whether the canonical default should
be empty or "title-of-policy".

## Out of scope (not differences in style)

- Tail content matches: both end at `[list additional insureds]` after
  the ADDITIONAL INSUREDS ENDORSEMENT. Reference has trailing empty
  paragraphs; ours does not. No extra endorsements either way.
- File size delta (reference 173KB vs ours 37KB) is embedded fonts and
  images in the reference, not body content.

## Pre-existing bug surfaced during this comparison

Rule 0's embedded-marker scan over `[A-Z]\.` patterns chews up real
text containing abbreviations and sentence-ending uppercase letters
("U.S.", sentence-ending "Coverage A."). Five hits in the test doc —
e.g., `2)  S. antitrust laws` from "U.S. antitrust laws". Boundary
`(?<![A-Za-z0-9).(])` only checks the char before the match; needs a
trailing lookahead `(?![A-Za-z])` to skip cases where the period is
followed by another letter. Tracked separately from the spec changes
above.
