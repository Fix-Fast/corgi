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
  -o /abs/path/to/output.docx \
  --parts-in /abs/path/to/policy.parts.json
```

Rule-oriented scripts are available as the public interface:

- `rule_1.py` — converge text hierarchy and body/heading styling
- `rule_2.py` — converge list structure and list formatting
- `rule_3.py` — converge page layout and running header

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

Do not assume coverage headings appear consecutively. For example, a
document may contain only `Coverage B`.

`parts.json` should look like:

```json
{
  "ignored_body_indexes": [0, 1],
  "title_indexes": [2],
  "section_heading_indexes": [8, 41],
  "subheading_indexes": [5, 12, 19],
  "header_title_text": "Commercial General Liability Policy",
  "policy_code": "CORGI-TECH-1234"
}
```

## Composing with other skills

- For offline DOCX work not covered by this skill (generic edits,
  content extraction), use Anthropic's `docx` skill from the
  `document-skills` plugin.

## What the pipeline does

1. `rule_1.py`: rebuild canonical document structure from the supplied
   document parts and apply heading/body text styling.
2. `rule_3.py`: apply section layout and the running header from the
   supplied document parts.
3. `rule_2.py`: normalize list numbering, suffix spacing, indentation,
   and list-marker styling.

Output is formatting-deterministic: for the same input DOCX, the
formatter produces the same document structure and styling. Container
metadata such as `docProps/core.xml` timestamps may still vary by run.

## Notes

- The formatter does not infer title, section-heading, subheading, or
  running-header values on its own. Claude is expected to supply them
  via `--parts-in`.
