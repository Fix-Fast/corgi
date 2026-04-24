# Section Taxonomy Patterns

Observed section structures across the in-house Corgi-Tech policy corpus. This is a living document — update it when new policies are authored or existing ones are restructured.

Every section is written here **by label and title** (e.g., *"Section I (Policy Terms and Conditions)"*) because bare labels drift silently when sections get renumbered. Same rule applies in the authored output.

## Taxonomy inventory

| Policy | Coverage basis | Has Coverage A/B/C? | Section count | Notes |
|---|---|---|---|---|
| CGL (Commercial General Liability) | Occurrence | Yes | 7 | The only occurrence-basis form; has a distinct *Section IV (Limitations of Liability and Retentions)* and *Section VI (Defense and Settlement)* that the claims-made forms fold elsewhere. |
| D&O (Directors and Officers Liability) | Claims-made | No (single insuring agreement) | ~5 | Section I (Policy Terms and Conditions) is heavy — 15+ provisions including all the claims-handling machinery that CGL scatters across Sections IV and VI. |
| E&O (Technology Errors and Omissions) | Claims-made | No (single insuring agreement) | ~5 | Closest structural twin to Cyber Liability. Section I (Policy Terms and Conditions) has 14 provisions. |
| EPL (Employment Practices Liability) | Claims-made | No | ~5 | Structurally similar to D&O and E&O. |
| Fiduciary Liability | Claims-made | No | ~5 | ERISA exposure adds some definitional and exclusion content. Structurally similar to D&O. |
| Media Liability | Claims-made | No | ~5 | Structurally similar to E&O; content focus is publishing, advertising, content-creation offenses. |
| Cyber Liability | Claims-made (with first-party components) | No | ~5 | **Outlier**: Section I (Policy Terms and Conditions) is minimal (only Duty to Defend + Cooperation/Settlement/Consent). Section II is elevated to *Section II (Notice of Claims and Potential Claims)*. Majority of other claims-made forms put Notice of Claims under Section I. |

## Common CGL taxonomy (occurrence-basis)

Observed in `COMMERCIAL GENERAL LIABILITY INSURANCE POLICY.pdf`:

1. *Section I (Policy Terms and Conditions)* — 1 top-level anchor (§1 Allocation) with a-j sub-items covering Allocation, Changes in Exposure, Loss of Subsidiary Status, Assignment, References to Laws, Most Favorable Jurisdiction, Coverage Territory, Premium Adjustments.
2. *Section II (Insuring Agreements)* — three coverage parts: *Coverage A (Bodily Injury and Property Damage Liability)*, *Coverage B (Personal and Advertising Injury Liability)*, *Coverage C (Medical Expenses)*.
3. *Section III (Definitions)* — 38 alphabetized defined terms.
4. *Section IV (Limitations of Liability and Retentions)* — Limits block (a-h) + Retention.
5. *Section V (Exclusions)* — four subgroups: *Applicable to All Coverages (Coverage A, B, and C)* (23 entries), *Applicable to Coverage A and Coverage C* (9 entries), *Applicable to Coverage B* (8 entries), *Applicable to Coverage C* (1 entry).
6. *Section VI (Defense and Settlement)* — 5 provisions.
7. *Section VII (Notices and Conditions)* — change-in-control, bankruptcy, notices-provision with insurer contact block, cooperation sub-provisions (a-r), entire-agreement clause.

## Common claims-made taxonomy (majority: D&O / E&O / EPL / Fiduciary / Media)

Observed across D&O, E&O, EPL, Fiduciary, Media forms. Majority convention:

1. *Section I (Policy Terms and Conditions)* — 14–17 provisions. Typical provisions, in approximately this order:
   - §1 (Defense of Claims and Settlement) *or split into* §1 (Duty to Defend) + §2 (Cooperation, Settlement and Consent)
   - §N (Notice of Claims and Potential Claims) — **this is the provision Cyber promotes to its own Section II**
   - §N (Cooperation)
   - §N (Limitations of Liability)
   - §N (Retentions)
   - §N (Related Claims)
   - §N (Coverage Territory)
   - §N (Extended Reporting Period)
   - §N (Allocation)
   - §N (Changes in Exposure)
   - §N (Application — Representations and Severability)
   - §N (Action Against the Insurer)
   - §N (Assignment)
   - §N (Named Insured's Authority)
   - §N (Cancellation)
   - §N (Bankruptcy or Insolvency)
   - §N (Additional Conditions)
2. *Section II (Insuring Agreements)* — single or multi-clause insuring agreement depending on form. E&O has Insuring Agreement A-C; some others have one.
3. *Section III (Definitions)*.
4. *Section IV (Exclusions)*.
5. *Section V (Notices and Conditions)* — **authorial choice, not a corpus majority.** The claims-made forms in this corpus (D&O, E&O, EPL, Fiduciary, Media) actually fold Notices and Conditions into *Section I (Policy Terms and Conditions)*; only CGL (Commercial General Liability) breaks them out into *Section VII (Notices and Conditions)*. For a new claims-made form authored in this corpus, the defensible call is either (a) follow claims-made majority and keep Notices/Conditions inside Section I, or (b) follow CGL's precedent of a dedicated Notices section at the end — particularly if the new form has many Notice-adjacent provisions (insurer contact block, cancellation mechanics, appraisal, vendor panel) that would bloat Section I. When proposing the taxonomy in Step 3, surface this as an explicit fork.

## Cyber Liability minority taxonomy

Observed in `CYBER LIABILITY POLICY (CLEAN).pdf`:

1. *Section I (Policy Terms and Conditions)* — **only 2 provisions**: §1 (Duty to Defend) and §2 (Cooperation, Settlement and Consent). Defense machinery only.
2. *Section II (Notice of Claims and Potential Claims)* — promoted to its own section rather than nested in Section I. Internally has numbered sub-provisions (Cooperation, Limitations of Liability, Retentions, Related Claims, Coverage Territory, Extended Reporting Period, Allocation, Changes in Exposure, Application, Action Against the Insurer, Assignment, Named Insured's Authority, Cancellation, Bankruptcy, Additional Conditions) — structurally equivalent to what D&O/E&O put under Section I.
3. *Section III (Insuring Agreements)* — first-party + third-party coverages specific to cyber exposure.
4. *Section IV (Definitions)*.
5. *Section IV (Exclusions)* — **note: the source PDF shows "Section IV" twice, apparently a numbering error**. When drafting a new cyber form, use *Section V (Exclusions)*.

## Section-taxonomy decisions the author must make

When authoring a new policy type, these are the forks to surface to the user:

1. **Claims-made vs occurrence.** Drives whether Retroactive Date, Extended Reporting Period, and claims-handling machinery are load-bearing.
2. **Single vs multi-coverage insuring agreement.** CGL is multi (A/B/C); D&O/E&O/Cyber tend to be single, though E&O has sub-coverages.
3. **Notice of Claims placement.** Nested under *Section I (Policy Terms and Conditions)* (majority) or promoted to its own section (Cyber's choice)?
4. **Defense and Settlement placement.** Dedicated section (CGL's *Section VI (Defense and Settlement)*) or folded into *Section I (Policy Terms and Conditions)* (claims-made convention)?
5. **Limitations of Liability and Retentions placement.** Dedicated section (CGL's *Section IV (Limitations of Liability and Retentions)*) or nested under *Section I (Policy Terms and Conditions)* (claims-made convention)?
6. **First-party components.** If the new product covers first-party loss (like Cyber's Incident Response Costs), the insuring-agreement structure needs to accommodate — and Section I (Policy Terms and Conditions) may need provisions for first-party triggers (e.g., Notice of Security Incidents) that pure third-party forms lack.

When presenting the proposed taxonomy to the user, always cite sections by label and title together, and flag whether each choice is majority or minority convention so the user can redirect.
