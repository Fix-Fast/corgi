---
name: insure-policy-draft-md
description: Compose a Corgi-Tech insurance policy as a markdown document and render it to canonical .docx in one shot via pandoc + insure-policy-format. Use this when the user wants to draft a new CGL policy and does not already have a live Word doc open, or wants reproducible non-interactive output. Prefer this over insure-policy-draft when the user says things like "draft a policy", "compose a CGL", "generate a policy doc", or "make a new policy for [insured]" and there's no open Word session. Also prefer this when the task is mostly boilerplate with at most a couple of per-insured decision points.
---

# insure-policy-draft-md

Compose a Corgi-Tech CGL policy as a single markdown file, then render to canonical `.docx` via `pandoc` and `insure-policy-format`. No live Word session required — no `word-bridge`, no snapshot/save dance, no mid-draft formatter friction.

## When to use

- User wants a new Corgi-Tech policy and no live Word doc is open.
- User says *"draft a policy for [insured]"*, *"compose a CGL"*, *"generate a policy skeleton"*, *"make a policy doc"*.
- Reproducibility matters — markdown + deterministic formatter is byte-stable.

## When NOT to use

- User wants to watch the doc fill up live in Word, make mid-draft redlines, or interactively redirect the drafting surface → `insure-policy-draft`.
- Reformatting an existing `.docx` → `insure-policy-format`.
- Narrow edit to an already-drafted doc → `word-bridge`.
- Non-Corgi-Tech policy → this skill's references don't fit.

## Requirements

- `pandoc` on PATH (macOS: `brew install pandoc`).
- `uv` on PATH (for the formatter CLI).

## Why this flow exists

The original `insure-policy-draft` works against a live Word document via `word-bridge`. That's great for interactive editing, but for policy composition it's mostly friction: the agent inserts paragraphs one by one, the user has to save, the formatter then runs post-hoc, and there's a gotcha where `pandoc` doesn't extract Word's running header into its AST (so the form code has to live in the body anyway). For this skill, markdown-first is simpler: compose once, render once, done.

The reference canonical text is already markdown. The formatter already shells out to pandoc internally. This skill just collapses the pipeline.

## Core rules

**1. Boilerplate text is source of truth.** Insert reference canonical text character-for-character. Do not paraphrase. If something in a reference looks wrong, surface it to the user before changing. (Same rule as `insure-policy-draft`; the references are shared.)

**2. Do not draft custom legal language.** At the per-insured decision points (Section III definitions, Section V exclusions), describe what's needed and pause. The user provides the text. If the user asks you to propose, be explicit that any proposed text must be reviewed (ideally by counsel) before it's treated as final.

**3. Don't format as you draft.** Your markdown is intentionally flat — plain paragraphs separated by blank lines, with list markers like `1)`, `a)`, `(i)` appearing as literal text at the start of a paragraph. The formatter classifies paragraphs and builds the list hierarchy during rendering. No markdown headings, no bold/italic, no list syntax.

**4. Form code goes in the body.** Write the running-header form code (e.g. `CORG-TECH-0100`) as the very first body paragraph in the markdown. The formatter reads it from there (pandoc does not extract Word's running header) and strips it during rendering, writing it into the canonical right-aligned gray running header on output.

## Workflow

### Step 1: Kickoff

Gather from the user (or default if they say "defaults"):

- **Insured name** — context for the Section III/V pause points. Not inserted inline.
- **Form code** — default `CORG-TECH-0100`.
- **Policy title** — default `COMMERCIAL GENERAL LIABILITY INSURANCE POLICY`.
- **Coverages in scope** — default A, B, C all in. If any is dropped, flag that Section II and Section V exclusion groupings both need to change.
- **Working directory** — where the intermediate `.md` and the final `.docx` go. If the user doesn't specify, default to `~/Documents/` with filenames based on a slug of the insured name (or `policy-draft` if no name).

### Step 2: Compose the markdown buffer

Read the canonical references from the sibling skill `insure-policy-draft/references/` (files `00-header-preamble.md` through `07-section-vii-notices.md`). These files exist as the single source of truth for boilerplate text and are shared between `insure-policy-draft-md` and `insure-policy-draft`.

For each reference:
- Skip the header frontmatter block (everything up to and including the `## Canonical text (insert verbatim)` heading and its `---` separator).
- Take the remaining content as-is.

Compose the output markdown as:

```
CORG-TECH-0100

<reference 00 canonical text>

<reference 01 canonical text>

...
```

The reference files already contain the literal list markers (`1)`, `a)`, `(i)`) at the start of paragraphs and use indentation for visual hierarchy. **Strip the leading indentation** when moving to the output markdown — the formatter doesn't use indentation as a signal; it uses the marker syntax. Leading spaces would also be parsed by pandoc as code-block indent or list continuation.

One terse way to do this in bash (per reference file):

```bash
awk '/^## Canonical text/{found=1; next} found' "$f" | sed 's/^ *//'
```

This skips everything up to and including the `## Canonical text...` heading, then strips leading whitespace from remaining lines.

### Non-interactive mode

If you're invoked in a context where there is no user to respond in real time (e.g. another agent spawned you with a fully-specified task, or the prompt itself resolves every input), treat both pause points as *"skip"* and leave the Section VII contact block as `——` placeholders. The prompt that spawned you should have resolved any per-insured content upfront; if it didn't, default to skeleton-only output and note in the close-out that no per-insured customization was applied.

### Step 3: Pause at Section III

After emitting the Section III canonical definitions (items 1–38, alphabetized), stop and ask:

> *"Standard definitions inserted. Are there any insured-specific terms to add — an industry-specific concept this policy will use in exclusions or coverage carve-outs? If you need research first, go do it and come back with the term + draft definition."*

If the user provides text: splice into the markdown buffer in alphabetical order among the existing definitions, renumbering accordingly. (Or insert unnumbered with a clear marker if renumbering is risky — flag the choice.)

If the user says "skip" / "no add-ons": continue.

### Step 4: Pause at Section V

After emitting the Section V canonical exclusions (groups 1, 23, 24, 25), stop and ask:

> *"Standard exclusions inserted. For this insured, are there: (a) additional exclusions to add, (b) existing exclusions to carve out or narrow, or (c) exclusions to remove entirely?"*

Same rules: user provides text, you insert verbatim. If the user asks you to propose, shape proposals off the parallel structure of neighboring canonical entries and be explicit the proposal is a first pass for review.

### Step 5: Handle the insurer contact block

Section VII contains `Address: ——`, `Fax Number: ——`, `Email: ——` as templated placeholders. If the user provides values, replace them. Otherwise leave `——` and flag at close-out that they must be filled before the policy issues.

### Step 6: Handle the Section III item-28 defect

The reference `03-section-iii-definitions.md` has a known defect: item 28 is a mis-numbered top-level definition that is actually a continuation of item 27 (Products-Completed Operations Hazard). Emit it as an unnumbered continuation paragraph under 27. Numbers 29–38 remain intact. This mirrors the decision made in `insure-policy-draft`. (If the reference gets patched upstream, revisit.)

Concrete fix: after composing `draft.md`, strip the `28) ` prefix from the one line that begins with it. A sed one-liner works:

```bash
sed -i.bak 's/^28) Bodily Injury or Property Damage arising out of the transportation/Bodily Injury or Property Damage arising out of the transportation/' <working-dir>/draft.md
rm <working-dir>/draft.md.bak
```

The match anchor is specific enough that it only hits the one intended line — no other paragraph starts with that exact prefix. Verify with `grep -n '^28) Bodily Injury' draft.md` (should return nothing after the sed).

### Step 7: Write and render

Write the composed buffer to `<working-dir>/draft.md`.

Run pandoc to convert markdown to DOCX. **Critical flag:** disable pandoc's `fancy_lists` and `startnum` extensions — without this, pandoc interprets `a) Defense Costs`, `1) Allocation`, `(i) ...` as ordered lists, which produces Word list items. The formatter expects plain paragraphs with literal markers; it drops items that arrive as Word lists. So:

```bash
pandoc -f markdown-fancy_lists-startnum <working-dir>/draft.md -o <working-dir>/draft.docx
```

Run the formatter (same plugin, sibling skill — absolute path is fine):

```bash
uv run /Users/milanb/Workable/corgi/plugins/docx/skills/insure-policy-format/scripts/format.py \
  <working-dir>/draft.docx \
  -o <output-path>
```

Where `<output-path>` is the final location the user asked for (e.g. `~/Documents/acme-robotics-policy.docx`), not necessarily in the working directory.

### Step 8: Close

Before summarizing, run two sanity checks against the final `.docx`:

```bash
pandoc <output-path> -t plain | grep -c '^CORG-TECH'   # Should be 0 — formatter strips the code from the body and moves it to the running header. If this is ≥1, something went wrong in the formatter step.
pandoc <output-path> -t plain | grep -c '——'           # Count of unfilled contact placeholders. If > 0, tell the user which fields (Address / Fax / Email) are still blank.
```

Then summarize for the user:
- Sections emitted, which had pause-point content added (or skipped).
- Form code (reminder: it's in the running header now, not the body — formatter moved it).
- Insurer contact block: filled or left as `——` (surface the count from the grep above).
- Any reference-source defects noted for counsel review.
- Path to final `.docx`.
- Path to intermediate files (`<working-dir>/draft.md`, `<working-dir>/draft.docx`) — these are useful if the user wants to re-render after tweaks, so point at them explicitly rather than leaving them buried.

## Anti-patterns

- **Paraphrasing canonical text.** Even tiny rewording drifts legal meaning.
- **Inventing exclusion or definition language.** Pause and let the user drive.
- **Trying to write markdown headings or list syntax.** The formatter expects flat paragraphs with literal markers; markdown lists/headings will misclassify.
- **Omitting the form code body paragraph.** Formatter will fail with *"no CORGI-TECH-\* / CORG-TECH-\* policy code found."*
- **Opening Word to watch it build.** That's the other skill's flow. This one is one-shot.
- **Leaving intermediate files around without pointing them out.** Tell the user where `draft.md` and `draft.docx` live — they may want to inspect or re-render.

## Composition with other skills

- **`insure-policy-format`** — called internally as the final step. Takes the pandoc output and canonicalizes it.
- **`insure-policy-draft`** — the live-Word alternative. Prefer that one if the user wants interactive mid-draft control.
- **`word-bridge`** — not used here. This flow doesn't touch a live Word session.
- **Research skills / external tools (user-directed)** — at the pause points, the user drives whatever research they need.

## Shared references

The canonical text lives in `../insure-policy-draft/references/`. Read from there; do not copy. Single source of truth prevents drift between the two drafting skills.
