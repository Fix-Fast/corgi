---
name: insure-policy-format
description: Deterministically reformat a Corgi-Tech insurance policy `.docx` (a document bearing a `CORGI-TECH-*` or `CORG-TECH-*` running-header code) into Corgi's canonical heading, list, and layout conventions via a bundled Python CLI. Use ONLY for these policy documents — the classifier is hard-coded to Corgi insurance-policy structure (SECTION I:, Coverage A —, typed 1)/a)/(i) list markers) and will mis-format any other document.
---

# Corgi insure-policy-format

Deterministic reformatter for Corgi-Tech insurance policies. Ships with
this plugin as a self-contained Python CLI; `uv` resolves its deps on
first run via PEP 723 inline script metadata — no install step.

## When to use

- The user hands you a Corgi-Tech insurance policy `.docx` and wants it
  put into canonical form.
- Running-header code matches `CORGI-TECH-*` / `CORG-TECH-*` or title
  matches `… INSURANCE POLICY` in all-caps.

## When NOT to use

- Any document that is not a Corgi-Tech insurance policy — the
  classifier assumes specific structure (`SECTION I:`, `Coverage A —`,
  typed list markers like `1)`, `a)`, `(i)`) and will mis-format other
  documents.
- Live edits on an open Word document — use the `word-bridge` skill for
  that.
- Generic DOCX creation or editing — use Anthropic's `docx` skill.

## Requirements

- `uv` on PATH (https://docs.astral.sh/uv/). The script declares its own
  Python and dependency requirements inline via PEP 723; `uv` handles
  environment setup automatically on first run.
- `pandoc` on PATH. The pipeline shells out to pandoc for DOCX → JSON
  AST conversion.

## Usage

```bash
uv run "${CLAUDE_PLUGIN_ROOT}/skills/insure-policy-format/scripts/format.py" \
  /abs/path/to/input.docx \
  -o /abs/path/to/output.docx
```

Debug run (writes `*.blocks.json`, `*.ast.json`, `*.report.json`):

```bash
uv run "${CLAUDE_PLUGIN_ROOT}/skills/insure-policy-format/scripts/format.py" \
  /abs/path/to/input.docx \
  -o /abs/path/to/output.docx \
  --artifacts-dir /abs/path/to/debug
```

Replay from a previously captured blocks file:

```bash
uv run "${CLAUDE_PLUGIN_ROOT}/skills/insure-policy-format/scripts/format.py" \
  /abs/path/to/input.docx \
  -o /abs/path/to/output.docx \
  --blocks-in /abs/path/to/prev.blocks.json
```

## Composing with the other plugin skills

- To pick up the freshly formatted DOCX inside an open Word session:
  run the CLI, then use the `word-bridge` skill to open or refresh the
  document.
- For offline DOCX work not covered by this skill (generic edits,
  content extraction), use Anthropic's `docx` skill from the
  `document-skills` plugin.

## What the pipeline does

1. Shell out to `pandoc` to convert the source DOCX to a JSON AST.
2. Flatten paragraphs, strip running-header junk, split paragraphs on
   embedded list markers.
3. Classify each paragraph as title / section heading / subheading /
   list item / body; parse typed markers (`1)`, `a)`, `(i)`, `(1)`,
   `(a)`, `(i)`) into a six-level ordered-list hierarchy.
4. Compose a canonical pandoc AST; render back to DOCX via pandoc.
5. Apply Corgi styles (Bricolage Grotesque headings, Inter body, sized
   margins, running header with policy code right-aligned).
6. Patch OOXML: canonicalize the numbering.xml level definitions, set
   list suffixes to `space`, strip paragraph-level indent overrides,
   fix list-marker font.

Output is byte-deterministic: same input DOCX → same output DOCX.

## Notes

- The formatter raises on malformed input rather than silently
  producing a broken output. If it fails, the error message will point
  at the specific structural assumption that was violated.
- Artifacts (`--artifacts-dir`) are for debugging only; production runs
  should omit that flag.
