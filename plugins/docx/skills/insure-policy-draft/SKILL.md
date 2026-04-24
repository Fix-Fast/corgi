---
name: insure-policy-draft
description: Build a new Corgi-Tech insurance policy document by walking the user through the canonical CGL skeleton section-by-section, inserting vetted boilerplate verbatim into a live Microsoft Word document via the word-bridge skill. Use whenever the user asks to draft, author, build, create, or start a new insurance policy — even if they don't mention CGL by name — as long as the target is a Corgi-Tech-style policy. Pauses for per-insured decisions (Section V exclusion customizations, Section III definition add-ons, templated values like form code and insurer contact) and hands off to the user for research; the skill itself does not draft custom policy language.
---

# insure-policy-draft

Iteratively build a Corgi-Tech insurance policy in a live Word document, one section at a time. The canonical skeleton — based on the CGL reference — lives in this skill's `references/` directory. Each section's text is inserted verbatim into the open Word document via the `word-bridge` skill; the user decides when to continue and where per-insured customization happens.

## When to use

- User asks to draft, build, author, create, or start a new Corgi-Tech insurance policy.
- User opens a (blank or shell) Word doc and wants it populated with the canonical skeleton.
- User says things like *"make a new CGL for [insured]"* or *"set up a policy doc for this insured."*

## When NOT to use

- User wants to *reformat* an existing policy → `insure-policy-format`.
- User wants a narrow targeted edit to a doc (change a heading, redline a paragraph) → `word-bridge` directly.
- User wants to draft a document that isn't a Corgi-Tech policy → this skill's skeleton won't fit and the references aren't appropriate.

## Prerequisites

- `word-bridge` is healthy. Call `bridge_status` first. If unreachable, ask the user to start Carson desktop (it hosts the bridge on `127.0.0.1:3137`) before continuing.
- A Word document is open and visible to the bridge (`list_documents` returns at least one). If not, ask the user to open one — the skill writes *into* whatever is currently active.

## Core rules

**1. Boilerplate text is source of truth.** When a reference file says "insert verbatim," insert it character-for-character via `word-bridge`. Do not paraphrase, re-tone, re-order, or "improve" the wording. The canonical language is legally reviewed; any deviation is liability. If something in a reference looks wrong, surface it to the user before changing anything.

**2. This skill does not draft custom legal language.** At the per-insured decision points (Section V exclusion customizations, Section III definition add-ons), describe what's needed and pause. The user will either hand-write the text, invoke their own research workflow and return, or say "skip." If the user asks you to *propose* language, be explicit that any proposed text must be reviewed (ideally by counsel) before it's treated as final, and shape proposals off the parallel structure of neighboring canonical entries rather than inventing new legal theories.

**3. Snapshot before each section.** Call `snapshot_document` with a label like `before-section-v` before a section's inserts, then `after-section-v` when done. Makes rollback cheap.

**4. One section at a time; confirm between.** After a section is inserted, summarize what landed and ask *"ready to continue to Section X, or revise this one?"* Do not chain sections together without the user's green light.

**5. Don't fight the formatter.** Inserted text may look visually off (indent ladder, list numbering, running header missing) during drafting. That's expected — the closing step runs `insure-policy-format` which canonicalizes indents, numbering, fonts, and running header. Don't attempt to format as you draft; it wastes time and conflicts with the deterministic formatter.

## Drafting workflow

### Kickoff

1. `bridge_status` → confirm healthy.
2. `list_documents` → confirm a target doc is active. If the doc has existing content, ask: append at the end, replace the whole body, or start from a specific anchor?
3. Gather minimal kickoff info:
   - **Insured name** (for context, to ground the Section V pause-point — not inserted inline).
   - **Form code** for the running header (e.g. `CORG-TECH-0100`). The `insure-policy-format` pass wires this into the right-aligned gray header later; for now, note it and remind the user at close.
   - **Policy title** — default to `COMMERCIAL GENERAL LIABILITY INSURANCE POLICY` unless the user indicates another line.
   - **Coverages in scope** — confirm A (Bodily Injury / Property Damage), B (Personal & Advertising), and C (Medical Expenses) all apply. If the user is dropping one, flag it: Section II's coverages list and Section V's exclusion subgroupings both change.
4. Offer to turn on Track Changes (`set_track_changes(enabled=true)`) if the user wants a reviewable redline trail; otherwise proceed without.

### Section loop

For each section in order, read the matching reference file, follow its header guidance, and insert the canonical text via `word-bridge`. The references are the source of truth for content — this table is just the index.

| # | Section | Reference | Kind | Pause? |
|---|---|---|---|---|
| 0 | Title + NOTICES + Preamble | `references/00-header-preamble.md` | Title templated; body boilerplate | no |
| 1 | Section I — Policy Terms and Conditions | `references/01-section-i-terms.md` | Boilerplate | no |
| 2 | Section II — Insuring Agreements | `references/02-section-ii-insuring.md` | Boilerplate (confirm coverages) | brief |
| 3 | Section III — Definitions | `references/03-section-iii-definitions.md` | Boilerplate + optional add-ons | **yes** |
| 4 | Section IV — Limits & Retentions | `references/04-section-iv-limits.md` | Boilerplate | no |
| 5 | Section V — Exclusions | `references/05-section-v-exclusions.md` | Boilerplate + **per-insured customizations** | **yes — main decision point** |
| 6 | Section VI — Defense & Settlement | `references/06-section-vi-defense.md` | Boilerplate | no |
| 7 | Section VII — Notices & Conditions | `references/07-section-vii-notices.md` | Boilerplate + templated insurer contact block | brief |

### Pause points (where the user has to decide)

**Section III — per-insured definitions.** After inserting the standard definitions, ask:
> *"Standard definitions inserted. Are there any insured-specific terms to add — an industry-specific concept this policy will use in exclusions or coverage carve-outs? If you need research first, go do it and come back with the term + draft definition."*

If the user provides text, insert it in alphabetical order among the existing definitions. If the user asks you to propose, shape proposals on the parallel of existing definitions (e.g. *"X means …"*) and be explicit the proposal is a first pass for review.

**Section V — per-insured exclusions.** After inserting the standard exclusions, ask:
> *"Standard exclusions inserted. For this insured, are there: (a) additional exclusions to add, (b) existing exclusions to carve out or narrow, or (c) exclusions to remove entirely? If you need research on comparable policies or this insured's risk profile, do that now — I'll wait."*

Same rules: user provides text; you insert verbatim. Proposals, if requested, are first-pass only.

### Closing

After Section VII is complete:

1. `snapshot_document` with label `complete-draft`.
2. `verify_doc` to catch broken styles.
3. Summarize what landed — coverages in scope, whether custom exclusions / definitions were added, whether the insurer contact block was filled or left as `——`, and the form code noted at kickoff.
4. Offer: *"Draft is complete. The layout (list numbering, running header with form code `CORG-TECH-0100`, fonts) isn't canonicalized yet — that's what `insure-policy-format` is for. Want me to save the doc and hand off to that skill?"* If yes, compose: save via `word-bridge`, then invoke `insure-policy-format` with the saved path.

## Composition with other skills in this plugin

- **`word-bridge`** — underlying edit transport. Every insert, snapshot, and verify call goes through it. This skill does not call office.js directly.
- **`insure-policy-format`** — final canonicalization pass. Run after drafting to fix list indents, numbering glyphs, suffix spacing, fonts, and the running header. Expects a saved `.docx` on disk; byte-deterministic.
- **Research skills / external tools (user-directed)** — this skill does not do research. At each pause-point, the user drives whatever research they need and returns with facts or draft text.

## Anti-patterns

- **Paraphrasing canonical text.** Even tiny rewording drifts legal meaning.
- **Inventing exclusion or definition language.** Pause and let the user drive.
- **Chaining sections without confirmation.** Each section completes → user confirms → next.
- **Skipping snapshots on large inserts.** No snapshot = no rollback path when something goes wrong.
- **Formatting as you draft.** Leave layout to `insure-policy-format`.
- **Drafting the Declarations page.** The main policy doc refers to a separate Declarations artifact for limits, named insured, policy period, retroactive date, etc. That's out of scope for this skill.
