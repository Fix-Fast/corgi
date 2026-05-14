---
name: insure-policy-format
description: Deterministically reformat a Corgi-Tech insurance policy `.docx` into Corgi's canonical heading, list, and layout conventions via a bundled Python CLI. Claude is expected to inspect the source document first, write a `parts.json` manifest, and pass that into the CLI.
---

# Corgi insure-policy-format

Deterministic reformatter for Corgi-Tech insurance policies. Ships with
this plugin as a self-contained Python CLI; `uv` resolves its deps on
first run via PEP 723 inline script metadata — no install step.

Packaged docs:

- Human-facing rules: `format.md`

## When to use

- The user hands you a Corgi-Tech insurance policy `.docx` and wants it
  put into canonical form.
- You can identify the important document parts yourself and pass them
  to the formatter as `parts.json`.

## When NOT to use

- Any document that is not a Corgi-Tech insurance policy — the
  formatter still assumes insurance-policy structure and may mis-format
  other document families.
- Live edits on an open Word document — use the `word-bridge` skill for
  that.
- Generic DOCX creation or editing — use Anthropic's `docx` skill.

## Requirements

- `uv` on PATH (https://docs.astral.sh/uv/). The script declares its own
  Python and dependency requirements inline via PEP 723; `uv` handles
  environment setup automatically on first run.

## Usage

```bash
uv run "${CLAUDE_PLUGIN_ROOT}/skills/insure-policy-format/scripts/format.py" \
  /abs/path/to/input.docx \
  -o /abs/path/to/output.docx \
  --parts-in /abs/path/to/policy.parts.json
```

The pipeline composes three pure rules (`(doc, parts) -> doc`) in
numerical order, mutating the OOXML tree directly:

- `rule_1.py` — text hierarchy and body/heading styling
- `rule_2.py` — outline marker normalization (pre-pass; only fires when
  `outline_normalizations` is set) followed by list structure and list
  formatting
- `rule_3.py` — page layout and running header

## Claude prompt contract

Before running the formatter:

1. Read `format.md`.
2. Inspect the source document.
3. Decide which source paragraphs are:
   - running-header content that should not become body text
   - the title
   - section headings
   - subheadings, including coverage headings and insuring-agreement
     headings
4. Decide what the running-header title text should be.
5. Decide what the policy code should be, if one is present.
6. Write those decisions to `parts.json`.
7. Run the deterministic formatter CLI with `--parts-in`.

If the document already has explicit heading structure, use that.

Coverage and insuring-agreement headings do not need to appear in a
fixed sequence. For example, a document may contain only `Coverage B`
without also containing `Coverage A` or `Coverage C`.

Each `*_texts` entry is a string drawn verbatim from the source. The
match rules:

1. The string is searched against the source corpus (all paragraphs
   joined by `\n`). It must match at exactly one position.
2. The target paragraph is the one **containing the start of the
   match**.

For unambiguous cases, just write a substring of the target paragraph
(short is fine — like `"NOTICES:"`). For paragraphs whose text alone is
ambiguous (e.g. duplicated lines), extend the string across paragraph
boundaries with `\n` and include enough text from a following neighbor
to make the match position unique. Use the next-paragraph context so
the match still STARTS in the target.

The formatter rejects no-match and multi-match cases with a message
that names the offending field, shows the candidate paragraphs, and
tells the LLM to extend the string with neighbor text.

`parts.json` may look like:

```json
{
  "ignored_body_texts": [
    "SEIC-DO-0100",
    "DIRECTORS AND OFFICERS LIABILITY\nSPORTS",
    "SPORTS AND ENTERTAINMENT ORGANIZATION\nSECTION I."
  ],
  "title_texts": [
    "DIRECTORS AND OFFICERS LIABILITY\nINSURANCE",
    "INSURANCE POLICY",
    "SPORTS AND ENTERTAINMENT ORGANIZATION\nMANAGEMENT",
    "MANAGEMENT LIABILITY COVERAGE"
  ],
  "section_heading_texts": [
    "SECTION I. INSURING AGREEMENTS",
    "SECTION II. DEFINITIONS"
  ],
  "subheading_texts": [
    "NOTICES:",
    "Insuring Agreement A – Non-Indemnifiable",
    "Insuring Agreement B – Indemnifiable"
  ],
  "outline_normalizations": [
    {
      "section_text": "SECTION IV. LIMITATIONS",
      "source_levels": [
        {"pattern": "([A-Z])\\.", "sequence": "upper_alpha"},
        {"pattern": "(\\d+)\\)", "sequence": "decimal"},
        {"pattern": "([a-z])\\)", "sequence": "lower_alpha"}
      ]
    }
  ],
  "header_title_text": "Directors and Officers Liability Policy",
  "policy_code": "SEIC-DO-0100"
}
```

`outline_normalizations` is optional. Include an entry per section
that uses outline markers that don't match the standard list format
(e.g. uppercase Roman numerals at the top level). Each entry tells the
formatter what the source's marker style looks like at each level so
it can rewrite to the standard list format
(`A.`/`1.`/`a.`/`(1)`/`(a)`/`(i)`) before Rules 1/2/3 process the doc.

Schema:

- `section_text`: a substring that resolves to a unique source section
  heading, using the same string-match rules as `section_heading_texts`
  (verbatim text from the source; extend across `\n` if needed). The
  rewrite covers the range from that heading up to (but not including)
  the next section heading.
- `source_levels`: ordered list (top-down) describing the source's
  marker style at each outline level. Each level entry is an object:
  - `pattern`: a Python regex with **exactly one capture group** for
    the enumeration token. Examples:
    - `"([A-Z])\\."` for `A. ... B. ... C. ...`
    - `"(\\d+)\\)"` for `1) ... 2) ...`
    - `"([a-z])\\)"` for `a) ... b) ...`
    - `"\\(([ivx]+)\\)"` for `(i) ... (ii) ...`
  - `sequence`: one of `decimal`, `upper_alpha`, `lower_alpha`,
    `upper_roman`, `lower_roman` — declares how the captured token
    should be parsed for ordering. Mismatched sequences (e.g. pattern
    `([A-Z])` with `sequence: decimal`) are rejected.

The formatter automatically anchors the regex at paragraph start (with
optional leading whitespace) for leading-marker scans, and prepends
boundary `(?<![A-Za-z0-9).(])` for embedded scans. This means citation
references like `Section IV.A.` and parentheticals like `sixty (60)
days` are NOT falsely matched.

The Rule 0 rewrite walks paragraphs in the section, tracks per-level
counters with parent-aware resets (any higher-level marker resets the
counters below it), and substitutes each matched marker with the
standard-list-format marker for that level: level 0 → `A.` `B.` ...,
level 1 → `1.` `2.` ..., level 2 → `a.` `b.` ..., level 3 → `(1)`
`(2)` ..., level 4 → `(a)` `(b)` ..., level 5 → `(i)` `(ii)` ....
Both the leading marker of a paragraph and any embedded markers inside
it are rewritten — so an inline `1) X ... or 2) Y` inside an A-level
item becomes inline `a. X ... or b. Y` in the standard form.

In the example above:

- `"NOTICES:"` is a unique substring → matches one position; target is
  the paragraph containing it.
- `"DIRECTORS AND OFFICERS LIABILITY\nINSURANCE"` distinguishes the
  title's first line from its body-repeat: the title is followed by
  `INSURANCE POLICY`, the body-repeat is followed by
  `SPORTS AND ENTERTAINMENT ORGANIZATION`. So `"...\nINSURANCE"` matches
  only the title position, and the target is the paragraph at the start
  of that match.
- `"DIRECTORS AND OFFICERS LIABILITY\nSPORTS"` similarly anchors the
  body-repeat.

## Composing with the other plugin skills

- To pick up the freshly formatted DOCX inside an open Word session:
  run the CLI, then use the `word-bridge` skill to open or refresh the
  document.
- For offline DOCX work not covered by this skill (generic edits,
  content extraction), use Anthropic's `docx` skill from the
  `document-skills` plugin.

## What the pipeline does

Each rule has the uniform signature `apply(doc: Doc, parts: ResolvedParts) -> Doc`,
where `Doc` bundles the document/numbering/styles/header etree roots.
`format.py` builds a `Doc` from the source `.docx` zip, threads it
through the rules in order, then serializes the trees back out:

1. `rule_1.py` — classify paragraphs into title / section heading /
   subheading / notices block / body via parts.json indices, remove
   ignored paragraphs, and apply the corresponding pStyle and run
   formatting. Also installs body-level doc defaults and the Heading 2
   style def.
2. `rule_2.py` — two phases:
   1. Outline marker normalization (conditional, runs only when
      `outline_normalizations` is set): rewrite non-canonical source
      markers (e.g. `I./A./1)`) to the canonical ladder
      (`A./1./a./1./a./i.`) within flagged sections.
   2. List detection: walk body paragraphs (skipping headings), detect
      leading list markers, attach `numPr` at the correct level, and
      install a single canonical multilevel numbering definition with
      one `<w:num>` per section so list counters restart at every
      Heading 2.
3. `rule_3.py` — set page margins, build the running header (`<w:hdr>`
   element on `Doc.header`), and wire up the section properties.

Output is formatting-deterministic: for the same input DOCX, the
formatter produces the same document structure and styling. Container
metadata such as `docProps/core.xml` timestamps may still vary by run.

## Notes

- The formatter does not infer title, section-heading, subheading, or
  running-header values on its own. Claude is expected to supply them
  via `--parts-in`.

## Maintaining this skill

`format.md` is the canonical spec of what the output should look like.
The rule scripts are a materialized view of `format.md` — when the spec
and a script disagree, the script is the bug. The golden tests in
`tests/golden/` (relative to the **corgi git root**, not the skill
directory) are the regression gate. Run them from the git root.

### Spec → script index

Use this map to find the script(s) affected by a `format.md` edit:

| `format.md` section | Owning script |
|---|---|
| Document parts (input categories) | `scripts/parts.py` |
| Rule 1 — Text hierarchy | `scripts/rule_1.py` |
| Rule 2 — Lists (outline normalization, marker recognition, list structure) | `scripts/rule_2.py` |
| Rule 3 — Running header and page layout | `scripts/rule_3.py` |

Cross-cutting (rarely edited via `format.md` changes alone):
- `scripts/_docx.py` — low-level OOXML helpers and the `Doc` bundle type
- `scripts/format.py` — pipeline orchestration

### The loop

1. Edit `format.md`. `format.md` is human-authored — when the user asks
   to update the formatting rules, open it in their editor first
   (macOS: `open plugins/docx/skills/insure-policy-format/format.md`,
   which uses their default rich-text/markdown editor) and wait for
   them to make changes before proceeding. The agent's job is
   propagation, not authorship.
2. Use the index above to find the affected script(s).
3. Edit the script(s) to match.
4. Run the goldens from the corgi git root (not the skill dir):
   `cd "$(git rev-parse --show-toplevel)" && uv run tests/golden/test_formatter_golden.py`.
   Add `--keep` to preserve produced `.docx` files in `/tmp/...` for
   manual inspection.
5. **If goldens pass**: the change had no observable effect on
   cgl/seic_do — done. Commit `format.md` + script changes together.
6. **If goldens fail**: do NOT silently regenerate. Surface the
   divergence to the user so they can see what changed.

   The default path is visual comparison via `word-bridge`: open the
   produced `.docx` (from `--keep`) in Word so the user can compare
   against the prior golden in real Word rendering. Visual review
   catches things that XML diffs gloss over.

   Only fall back to summarizing the canonical XML diff in plain
   language ("section headings are now 16pt instead of 14pt", etc.)
   if `word-bridge` is unavailable or the user explicitly asks for a
   text-only diff.

   Then, **only after the user confirms intent**:
   - If the diff matches the spec change → regenerate the goldens from
     the corgi git root with
     `cd "$(git rev-parse --show-toplevel)" && uv run tests/golden/test_formatter_golden.py --regenerate`, then
     commit the regenerated `.docx` files as a **separate,
     explicitly-labeled commit** (`regenerate cgl golden: <one-liner
     reason>`). Never bundle regeneration with the script change — the
     audit trail depends on these being separable.
   - If the diff doesn't match intent → fix the script, leave the
     golden untouched.

### When `format.md` introduces a new structural concept

If a `format.md` change introduces a new *category* of document part
(like `notices_block` was), the change spans more than one file:

- `scripts/parts.py` — add the new field to the JSON schema and resolver
- This `SKILL.md` — document the new field in the "Claude prompt
  contract" section above
- The relevant rule script — handle the new category

Worth scanning all three when the spec gains a new concept.

## Coding Info

The section heading formatting is carried by the `Heading 2` style
definition in `styles.xml`, not by run-level character overrides. Rule 1
injects/overwrites the `Heading 2` style def to match these values and
leaves the runs themselves bare.

Body styling is also installed at the OOXML doc-default level
(`<w:docDefaults><w:rPrDefault>`) so anything that doesn't override —
notably the list markers, whose canonical level definitions omit
run formatting — inherits Inter / 11pt / black.