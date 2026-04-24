---
name: insure-policy-author
description: Author a new Corgi-Tech insurance policy form for a product line that does NOT yet have canonical boilerplate (e.g., a new Cyber Liability, Crime, Representations & Warranties, or other novel form) by adapting from existing in-corpus policies as structural and stylistic precedent, supplemented by SERFF and web research at pause points. Use whenever the user wants to draft, create, or design a new TYPE of policy — a new product form rather than a new instance of an existing form. Trigger phrases include "new policy form", "new product line", "author a cyber policy", "draft a new [X] liability policy we don't have yet", or "we want to add [X] as a product". Distinct from insure-policy-draft-md, which composes from canned canonical references; this skill PRODUCES the canonical references by making authorial choices grounded in precedent. If the target policy type already has a references/<type>/ bundle, prefer insure-policy-draft-md instead.
---

# insure-policy-author

Author a new Corgi-Tech policy form from scratch — not by inserting canned boilerplate (that's `insure-policy-draft-md`) but by **deriving a house-style skeleton from existing policies in the corpus** and adapting clauses section by section with user adjudication at every substantive decision. Output is a markdown draft + canonical `.docx`; if the draft is later adjudicated as canonical, its sections become the seed reference bundle for `insure-policy-draft-md` to compose from on subsequent drafts of that product line.

## When to use

- User wants to draft a **new policy type** the company doesn't yet have canonical text for (new Cyber form, new Crime form, new R&W form, etc.).
- User has a corpus of existing in-house policies that can serve as house-style precedent.
- User is willing to make (or adjudicate) authorial and legal judgment calls section by section. This skill is *not* push-button.

## When NOT to use

- User wants to draft a new instance of an **existing** product line (e.g., another CGL variant) → `insure-policy-draft-md` (if references exist) or `insure-policy-draft` (live Word flow).
- User wants to reformat an existing policy `.docx` → `insure-policy-format`.
- User wants a narrow edit to an already-drafted policy → `word-bridge`.
- There is no in-house precedent corpus to draw from. This skill assumes precedent. Without it, the output would be pure invention — outside scope.

## Requirements

- Read access to the customer's policy corpus (a directory of `.pdf` or `.docx` policy forms).
- `pandoc` on PATH.
- `insure-policy-format` available in this plugin for the final canonicalization pass.
- *Optional but strongly recommended*: the customer's SERFF search skill. SERFF (NAIC's form-filing system) gives access to comparable policies filed by other carriers. The customer has programmatic access via a separate skill; invoke it at pause points when a clause needs market grounding.
- *Optional*: WebSearch / WebFetch for secondary research (regulatory references, trade press on emerging risks, etc.).

## Why this skill is different from `insure-policy-draft-md`

`insure-policy-draft-md` assumes the canonical text already exists (in `references/`) and the agent's job is to assemble and customize it. That's the right shape for drafting another CGL (Commercial General Liability) — canonical text is known, pause points are narrow.

For a policy type we've never drafted before, the canonical text doesn't exist yet. Trying to "compose" it is the wrong frame; we need to **author** it, which means: derive the skeleton from analogous forms, adapt language clause-by-clause, flag every judgment call, and ground novel provisions in external research (SERFF, regulatory sources). Once authoring is complete and adjudicated, the output becomes the canonical text — at which point future drafting of the same product line flips back to the `-draft-md` flow.

## Core rules

**1. Every clause has a precedent citation.** For each clause in the output, record which precedent policy (and which section within it) the clause was adapted from. If a clause has no direct precedent — because the new product covers exposure no existing form addresses — flag it explicitly as *"authored de novo, grounded in [SERFF filing / regulatory source / user input]"*. A clause with no documented lineage is a red flag; surface it to the user.

**2. Do not invent novel legal theories.** Reuse language patterns from precedents. If the new product requires a genuinely novel provision (e.g., an "Incident Response Costs" definition for a Cyber form when no precedent uses that term), ground it in a SERFF-filed form or regulatory citation that the user has surfaced. Do not fabricate legal constructs.

**3. Section taxonomy is a deliberate choice, not an assumption.** Different policies in the same corpus may partition content differently (e.g., some put Notice of Claims under Section I; others promote it to its own Section II). Survey the precedent corpus, propose a taxonomy, and explicitly confirm with the user before authoring.

**4. Pause at every substantive section, not just two.** Unlike `-draft-md`, which only pauses at Section III definitions and Section V exclusions, this skill pauses after each section's draft for user review. Authoring new legal language warrants more adjudication surface, not less.

**5. Output is seed, not finished product.** The `.docx` this skill produces is a first pass for counsel review. The close-out message must list every authorial judgment call, every de-novo clause, and every SERFF/research gap. Do not imply the output is shippable without review.

**6. Respect house style.** Running-header code format, defined-term capitalization, list-marker conventions (`1)`, `a)`, `(i)`), notice-block phrasing, disclaimers — all of these are house-style signals the skill learns from the precedent corpus. Match them exactly in the output. The formatter's classifier depends on these being consistent.

**7. Always reference by label *and* title.** Never cite a section, coverage, or numbered provision by bare label alone. Write *"Section I (Policy Terms and Conditions)"*, not *"Section I"*. Write *"Coverage A (Bodily Injury and Property Damage Liability)"*, not *"Coverage A"*. Write *"Section V §22 (Recording and Distribution of Material or Information in Violation of Law)"*, not *"Section V-22"*. This applies to **three** places:
- **This SKILL.md and any reference files** it loads — so the agent reading the skill always sees what the label means.
- **Precedent citations recorded during authoring** — *"adapted from D&O Section I (Policy Terms and Conditions) §7 (Allocation)"*, not *"D&O §I-7"*.
- **Cross-references inside the authored policy body** — when the drafted policy refers to its own sections, both label and title must appear at least on first use. Bare labels are fine on second-and-subsequent reference within a paragraph, but every fresh paragraph that cites a section reintroduces the title.

The rationale: bare labels are brittle. If someone later renumbers sections, every bare *"Section V"* becomes stale silently. *"Section V (Exclusions)"* makes the intent explicit and lets reviewers catch mismatches. It also makes precedent citations self-documenting when you come back to the draft six months later.

## Workflow

### Step 1: Kickoff

Gather from the user:

- **Target policy type** — plain-English name (e.g., "Cyber Liability", "Crime", "Representations & Warranties").
- **Precedent corpus location** — directory containing the in-house policies to draw from. Ask the user for this path if it is not already known in the current workspace; do not assume a developer-local fixture path.
- **Scope hints** — what exposures the new product should cover. First-party, third-party, both? Claims-made or occurrence? New-entity coverage if acquired during the policy period? Any early constraints the user wants to bake in (e.g., "must include breach-response costs", "no regulatory defense coverage").
- **Form code** — if the user has a code assignment convention (the corpus suggests `CORG-TECH-0100` for CGL, `CORG-TECH-0200` for Cyber, etc.), let the user pick a new code. If they don't have one, propose next-available.
- **Output path** — where to save the final `.docx`.
- **Working directory** — for intermediates (`draft.md`, intermediate `.docx`). Default: `~/Documents/<policy-slug>-workdir/`.

### Step 2: Precedent survey

Read **all** policy files in the precedent corpus. For each, extract:

- **Section taxonomy**: what sections appear, in what order, with what titles.
- **Top-level provisions in each section**: numbered provisions, titled as in the form.
- **House-style markers**: list marker syntax (`1)` vs `(1)` vs `1.`), running-header code format, defined-term capitalization, notice/disclaimer phrasing.
- **Coverage trigger basis**: claims-made vs occurrence, whether there's a Retroactive Date.
- **Closeness to target**: for the new product, rank precedents by structural affinity. Example: for Cyber Liability (claims-made, professional-exposure-adjacent), Technology E&O (Errors and Omissions) and Media Liability are closest; CGL (Commercial General Liability — occurrence, general third-party) is furthest.

The reference file `references/section-taxonomy-patterns.md` documents the taxonomies observed in the current fixture corpus. Update it as new precedents are added.

### Step 3: Propose and confirm section taxonomy

Based on the survey, propose a section taxonomy for the new product. Always present sections with both label and title so the user can review intent, not just structure. Template:

> *"Based on the corpus, for a \[target\] policy I'd structure it as:*
> *- Section I (Policy Terms and Conditions) — 17 provisions, following the D&O (Directors and Officers Liability) / E&O (Errors and Omissions) majority convention.*
> *- Section II (Insuring Agreements) — X clauses covering first-party breach response, third-party liability, and regulatory defense.*
> *- Section III (Definitions).*
> *- Section IV (Exclusions).*
> *- Section V (Notices and Conditions).*
>
> *Alternative: the real Cyber Liability market sometimes promotes Notice of Claims to its own Section II (Notice of Claims and Potential Claims) rather than nesting it inside Section I (Policy Terms and Conditions). Happy to follow either convention; majority is Section I.*
>
> *Confirm before I start authoring, or redirect."*

Do not proceed to authoring until the user signs off on the taxonomy.

### Step 4: Per-section authoring loop

For each section in the confirmed taxonomy:

**a. Survey the analogs.** For this section, identify which precedents have the closest analog. E.g., for Cyber Section I (Policy Terms and Conditions), the E&O (Errors and Omissions) and D&O (Directors and Officers Liability) Section I (Policy Terms and Conditions) are closest; Media Liability is second tier; CGL (Commercial General Liability) is off-precedent.

**b. Derive provision list.** List the top-level provisions the section should contain. For each provision, note which precedents have it and how they phrase the heading.

**c. Draft each provision in precedent-adapted language.** For each provision:
- Pull the closest-precedent language as starting point.
- Adapt for the target product's specifics (e.g., swap `Wrongful Act` for `Security Incident` where appropriate).
- Preserve house-style markers and convention.
- Record the precedent citation **by label and title, not bare label**. Example: *"adapted from E&O (Errors and Omissions) Section I (Policy Terms and Conditions) §7 (Allocation), with cyber-specific additions for Incident Response Costs"*. Not *"E&O §I-7"*.

**d. Flag additions and de-novo clauses.** If the section needs a provision no precedent has:
- Describe what's needed in plain English.
- Ask the user: *"No precedent in the corpus has [X]. I can draft this from market convention or from a SERFF search of comparable forms. Which?"*
- If the user invokes SERFF (via their separate skill), consume the result as precedent. If the user asks you to draft from market convention, be explicit that the draft needs counsel review.

**e. Pause for user review.** After drafting the section, present a summary:
- Section name (label and title), provision list.
- Precedent citations per provision.
- Any de-novo clauses, flagged.
- Gaps or uncertainties for counsel.
- Length in lines of generated markdown.

Wait for the user to approve, redirect, or substitute language before moving to the next section.

**f. Intra-section pauses for long sections.** Sections with more than ~10 top-level provisions warrant an intermediate pause partway through, not just at the end. Two common cases:
- **Section III (Definitions)** tends to have 30–45 alphabetized terms. Pause after drafting the first half (roughly A–M), present the term list + precedent lineage, get approval, then draft the second half.
- **Section IV or V (Exclusions)** tends to have 15–25 exclusions grouped into subgroups. Pause after each subgroup (e.g., after *Applicable to All Coverages*, before starting *Applicable to Coverage A and Coverage C*).

The rationale: a single end-of-section pause on a 45-item list means the user has to review 45 provisions at once, and any substitution requires unwinding the whole section. Intra-section pauses limit the blast radius of any redirect. Use judgment — if a section is purely mechanical (e.g., a 5-provision Notices and Conditions section), a single end-pause is fine.

### Step 5: Compose the markdown buffer

After all sections are drafted and approved, assemble the composed markdown the same way `insure-policy-draft-md` does:

- **First line: the form code** as a bare paragraph (formatter reads it from body; pandoc doesn't extract Word headers). Example: `CORG-TECH-0200`.
- **Second substantive paragraph: the ALL-CAPS policy title** that ends in `INSURANCE POLICY`. The `insure-policy-format` classifier requires an ALL-CAPS title paragraph matching `/INSURANCE POLICY/` to recognize the document as a Corgi-Tech policy; without it the formatter fails with *"no title block classified"*. Example: `CYBER LIABILITY INSURANCE POLICY`.
- Each section appended in order, separated by blank lines.
- Plain paragraphs with literal list markers (`1)`, `a)`, `(i)`). No markdown headings, no list syntax, no bold/italic — the formatter handles all of that.

### Step 6: Render

```bash
pandoc -f markdown-fancy_lists-startnum <working-dir>/draft.md -o <working-dir>/draft.docx
```

Then invoke the `insure-policy-format` skill to canonicalize `<working-dir>/draft.docx` into `<output-path>`. Do not run the formatter script directly from this skill; the formatter skill owns its own CLI path, dependencies, and execution details. The `-f markdown-fancy_lists-startnum` flag is **critical** — without it pandoc promotes `a) ...`, `1) ...` to Word ordered lists and the formatter drops them. See `insure-policy-draft-md` for the full explanation of this gotcha.

Sanity-check the output:

```bash
pandoc <output-path> -t plain | grep -c '^CORG-TECH'   # Must be 0; code should be in running header only
pandoc <output-path> -t plain | grep -c '——'           # Count of unfilled templated placeholders, if any
```

### Step 7: Close-out summary

Deliver to the user:

- **Path to final `.docx`**, intermediate `draft.md`, intermediate pandoc `draft.docx`.
- **Section taxonomy chosen** and majority-vs-minority rationale.
- **Precedent lineage map** — one line per section summarizing which precedents were drawn from.
- **Authorial judgment calls for counsel review** — every de-novo clause, every language adaptation that went beyond mechanical substitution, every flag raised during the per-section pause points.
- **Research gaps** — any provisions the skill couldn't ground in precedent or external research, and what the user would need to fill in.
- **Next step**: *"Once this draft is adjudicated as canonical, it can be converted into a reference bundle under `insure-policy-draft-md/references/<type>/` so future instances of this form can use the cheaper compose-only flow."*

## Anti-patterns

- **Authoring without precedent citation.** Every clause has lineage or it's flagged de-novo. No exceptions.
- **Guessing house style from one precedent.** Survey the whole corpus; majority convention usually wins (flag minority choices).
- **Proceeding past taxonomy without user sign-off.** Section structure is a load-bearing decision; don't skip the confirmation.
- **Collapsing the pause points.** The `-draft-md` skill pauses twice because canonical text is known; this skill pauses per section because it isn't. Don't copy the lean pause model from the cheaper skill.
- **Treating output as shippable.** Close-out must frame the `.docx` as a first-pass draft for counsel review. Downstream flows (adjudication, re-rendering from the adjudicated version as canonical references) are out of scope for this skill.
- **Fabricating legal theories.** If the exposure is genuinely novel, the user brings the precedent (via SERFF or otherwise); the skill doesn't invent.
- **Skipping the precedent survey.** The temptation is to jump to authoring Section I of Cyber by reading one E&O policy. Don't — the taxonomy step requires surveying all of them, because section partition is a per-policy choice and majority matters.

## Composition with other skills

- **`insure-policy-draft-md`** — this skill's downstream successor once its output is adjudicated. The `.docx` this skill produces can be decomposed into `references/<type>/0N-section-N.md` files that `-draft-md` consumes on future invocations of the same product line.
- **`insure-policy-format`** — hand off after creating `<working-dir>/draft.docx`. Same canonicalization pass as `-draft-md`.
- **`insure-policy-draft`** — not used. This skill is markdown-first, not live-Word-first.
- **`word-bridge`** — not used.
- **Customer's SERFF search skill** — invoked by the user at pause points when a clause needs market-filed-form grounding. This skill does not call SERFF directly.
- **WebSearch / WebFetch** — invoked by the user or this skill at pause points for regulatory citations, trade-press context, or emerging-risk characterization.

## Shared reference

The file `references/section-taxonomy-patterns.md` catalogs section structures observed in the current fixture corpus. It's a living document — update it as new policies are authored or adjudicated.
