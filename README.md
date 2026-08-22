# Hidden Champions — with Validation

A variant of the [Hidden Champions HVT pipeline](../HCP/README.md) that adds a 5th
pipeline step: an **independent, automated second opinion** on every claim the base
pipeline already drafted. This document explains what that addition is, why it
exists, and — since this variant is a direct upgrade of the base pipeline rather
than a different tool — exactly what it does better and how, for a reader
completely new to both.

## 1. The problem this variant solves

The base Hidden Champions pipeline (`HCP/apps_script/`) already does a lot: it
searches the web, structures what it finds into four pillar claims
(differentiation, moat, product-improvement, alignment) via Gemini, and computes a
`triage_status` flag (`needs-review` vs `auto-confirmed`) so a human reviewer knows
which rows are worth a manual look.

But that `triage_status` comes from the **same pass that wrote the claims in the
first place** — the structuring step grading its own homework. At the project's
real operating scale (50–70 companies processed a day), a human genuinely cannot
manually re-verify every claim on every company. The base pipeline's honesty signal
is useful, but it's a single, self-reported opinion.

**This variant adds a second, independent opinion** — a separate Gemini pass whose
entire job is to try to find a reason the first pass's claim is wrong, not to
confirm it. It reproduces, in code, the manual verification discipline a human
reviewer would apply by hand: check the claim is actually about the right company,
re-fetch its cited source live (not trust the old cached copy), actively look for a
second independent source, and flag anything that doesn't hold up — all without
requiring a human to do this per-row.

## 2. What's added, concretely

### A brand-new pipeline step: Gemini Validation

One new file, `Validation.js`, adds `runGeminiValidation()` as **Step 5** in the
menu (after the base pipeline's Step 4, "Run Gemini Structure"). For each of the
three pillar claims already drafted (differentiation, moat, product_improvement —
alignment is deliberately excluded, since its source is always forced to match its
already-verified parent pillar), it:

1. **Re-fetches the claim's cited source live**, right now, rather than trusting
   whatever was scraped and cached when the claim was first drafted — a source page
   can change or go stale between when a claim was written and when someone
   eventually reviews it.
2. **Checks for corroboration.** If nothing already staged for that pillar comes
   from a source independent of the one being cited, it fires one fresh,
   deliberately different-shaped search (not just re-searching the company name
   again) specifically looking for outside confirmation.
3. **Checks identity coherence** — a pure, no-API-call sanity check: does the
   claim's cited source domain actually appear anywhere among search results
   scoped to this company's own name? If not, that's flagged as a real risk the
   claim describes a *different* company that happens to share a name (this exact
   failure mode was found in production — see §5).
4. **Asks Gemini for an adversarial verdict** — explicitly instructed to look for
   reasons the claim might be wrong, not reasons to accept it — landing on one of:
   `Verified`, `Needs Verify`, `Rejected`, or `Contradiction Found`.
5. **Writes the verdict to new, separate columns** (`*_auto_review`,
   `*_auto_notes`) that sit right next to the existing human-only `*_review`/
   `*_notes` columns, but never touch them.
6. **Attempts to fill any pillar that's still completely empty**, if — and only
   if — real evidence for it already exists. The fill is clearly labeled
   `Inferred — Low Confidence` rather than presented as an equally solid claim, and
   a pillar with genuinely zero evidence is left empty (there's nothing to guess
   from — filling it in anyway would be exactly the kind of fabrication the base
   pipeline's prompt rules already forbid).
7. **Computes a second, independent triage flag** — `auto_triage_status` — sitting
   alongside the base pipeline's own `triage_status`, so a row can be flagged by
   either signal without the two being merged into one and losing which check
   actually caught the problem.

### Schema additions (visible in the sheet)

Six new per-pillar columns (two per pillar × three pillars), immediately after
each pillar's existing human `*_review`/`*_notes` pair:

| Column | What it holds |
|---|---|
| `differentiation_auto_review` / `_auto_notes` | The independent verdict + full reasoning for the differentiation claim |
| `moat_auto_review` / `_auto_notes` | Same, for the moat claim |
| `product_improvement_auto_review` / `_auto_notes` | Same, for the product-improvement claim |

Plus one new row-level column, `auto_triage_status`, next to the existing
`triage_status`.

The `*_auto_review` verdict uses its **own enum**
(`Not Run`/`Verified`/`Needs Verify`/`Rejected`/`Contradiction Found`/`Inferred —
Low Confidence`), deliberately different from the human `*_review` enum, so a
machine verdict can never be visually mistaken for a human one even by someone
skimming the sheet without reading headers carefully.

### Everything else that changed to support it

- **Menu.js** — adds "5. Run Gemini Validation" (renumbering Company Summary to
  6th), plus a new **"⚠ Wipe ALL Company Data (Testing Reset)"** item for clearing
  a test sheet back to empty (requires typing the literal word `WIPE` to confirm —
  the only irreversible bulk-delete action in either variant of this project).
- **SheetSetup.js / SheetIO.js** — the schema migration and column-backfill logic
  was extended to also create/validate/default the six new `*_auto_review`
  columns, using the same safe, never-overwrite backfill pattern the base pipeline
  already uses for the human review columns (defaults to `Not Run` rather than
  `Pending`, so the two column families stay visually distinguishable even before
  either pipeline stage has run).
- **Batch.js** — the batch orchestrator's step list gained a `validation` step
  (`BATCH_STEPS = [..., 'validation']`), so a full unattended overnight batch run
  now validates every company automatically, not just structures them.
- **Triage.js (REVIEW_QUEUE)** — the review queue now lists a row if **either**
  `triage_status` or `auto_triage_status` says needs-review, with a new "Why"
  column showing which signal(s) fired (`structuring`, `validation`, or both) and
  an "Auto-Review Flags" column showing exactly which pillar(s) the validation
  pass flagged and why — so a reviewer can see at a glance whether this is a
  self-reported doubt, an independently-caught problem, or both.
- **PipelineSummary.js** — the summary sheet's tracked columns were extended to
  include the six new auto-review columns and `auto_triage_status`.

## 3. Why this is better than the base pipeline

| | Base Hidden Champions | With Validation |
|---|---|---|
| Who checks a claim | The same Gemini pass that wrote it | A separate Gemini pass, instructed to look for reasons it's wrong |
| Source freshness | Trusts the originally-scraped cached text | Re-fetches the cited source live before judging it |
| Corroboration | Whatever evidence happened to be gathered during the original search | Actively searches for independent confirmation if none already exists |
| Same-name company mix-ups | Not checked | Explicitly checked — flags if a claim's source never appeared under this company's own name search |
| Empty pillars | Stay blank | Filled with a clearly-labeled low-confidence guess, if real evidence supports one — never fabricated from nothing |
| Review signal | One flag (`triage_status`), self-reported | Two independent flags (`triage_status` + `auto_triage_status`), so an honest reviewer can see whether structuring, validation, or both raised concern |
| Review queue | Lists rows flagged by structuring only | Lists rows flagged by either signal, with a "Why" and per-pillar detail column |
| Batch/overnight runs | Structure only | Structure **and** independently re-check, unattended |

The net effect: at the 50–70-companies-a-day scale where a human physically cannot
spot-check most rows, the `needs-review` flag in this variant is a genuinely
stronger signal — a row that's clean here has been checked twice, by two
independently-instructed passes, not once.

## 4. What this addition deliberately does NOT do

These boundaries are enforced in code, not just by prompt instruction, precisely
because soft prompt-only rules have failed to hold elsewhere in this project even
after rewording:

- **Never writes the human `*_review`/`*_notes` columns.** Those stay
  human-only, forever — the validation pass only ever writes to the separate
  `*_auto_*` columns.
- **Never writes `cin`/`ownership_type`.** Company identity stays a manual
  registry check, same discipline as the base pipeline.
- **Fails closed, always.** An unreachable source, a malformed Gemini response, or
  missing evidence all resolve to `Needs Verify` — never silently to `Verified`.
- **An identity-coherence warning can never coexist with a `Verified` verdict** —
  enforced in code as a hard override, even if Gemini's own response says
  otherwise.
- **A failed re-fetch can never justify a `Rejected` verdict on its own** — an
  unreachable page is not evidence a claim is false; it's just unconfirmable this
  pass. Also enforced in code, not just prompted.
- **An `Inferred — Low Confidence` fill never counts as needs-review** — it's
  treated as an accepted, visibly-labeled judgment call, not an unresolved doubt.
- **Verified no longer requires 2+ independent sources by default.** This is a
  deliberate loosening from an earlier, stricter rule: requiring multiple sources
  for every claim produced a review queue full of ordinary small companies with
  nothing actually wrong, which drowned out the rows that genuinely needed
  attention. A single, plausible, uncontradicted, identity-coherent source is now
  allowed to be `Verified` — the corroboration tier is still recorded on every
  verdict for reference, it just no longer gates the verdict by itself.

## 5. Why this exists — the real incident that prompted it

The identity-coherence check specifically targets a contamination pattern found in
production: a pillar claim for one company ("Vikalp") was actually sourced from an
unrelated same-name entity ("Nostrain") — a claim that read as perfectly plausible
on its own, but was simply about the wrong company. Checking whether a claim's
cited source domain ever appeared in search results scoped to the company's actual
name catches exactly this pattern, cheaply, with zero extra API cost (it's a pure
set-membership check against evidence already gathered).

## 6. Cost and re-run behavior

- **Re-checking is skip-unchanged by default.** If a pillar's claim text hasn't
  changed since its last validation pass, that pillar is skipped on the next
  batch/automated run — a fresh Gemini call on identical input would be pure
  waste. A short fingerprint of the claim text is stored inside `*_auto_notes`
  itself specifically to detect this, without needing a separate hidden column.
- **A manual, single-row click always force-reruns**, regardless of whether the
  claim changed — because a deliberate click from a reviewer means "check this
  now," and because a prompt/logic fix to the validation code itself wouldn't
  otherwise trigger a fresh check on an unchanged claim (the staleness check can
  only detect "the claim changed," not "the validation logic changed"). Batch runs
  keep the cost-saving skip-unchanged behavior, since re-validating 50–70
  companies a day at full cost every single day isn't sustainable.
- **Source re-fetches are cache-first**, going through the exact same
  cache-then-scrape path the rest of the pipeline already uses — a re-run costs
  nothing extra unless the cache was explicitly cleared for that row.
- **Corroboration searches only fire when genuinely needed** — only when nothing
  already staged for that pillar comes from an independent source, so this doesn't
  double the Serper cost on every single validation pass.

## 7. How to run it

Same operational pattern as the base pipeline (see the base
[README](../HCP/README.md) for full first-time setup — Script Properties, `clasp
push`, etc.), with one addition:

1. Run the base pipeline's steps 1–4 first (Serper Grounding → Resolve Identity →
   Sector-Context Grounding → Gemini Structure) so there are claims to validate.
2. Run **"5. Run Gemini Validation"** from the menu for a single company (cursor on
   its row), or let a batch/overnight run include it automatically — it's now part
   of the default `BATCH_STEPS` sequence.
3. Check the new `*_auto_review`/`*_auto_notes` columns, or open **REVIEW_QUEUE**
   to see every row flagged by either signal, with the reason(s) why.

## 8. Relationship to the base pipeline

This folder was split out of `HCP/apps_script_with_validation/` into its own
top-level location on 2026-08-22, to keep it physically separate from the base
(non-validation) pipeline in `HCP/apps_script/`. It has its own `.clasp.json`, so
it's bound to its own separate Google Sheet deployment — **not** the same live
script as the base `HCP/apps_script/`. The downstream profile-writing stages
(`profile_generation/`, `target_companies/`, `temp_profile_creator/`) are not
duplicated here — they remain in `HCP/` and are shared by both pipeline variants.

## 9. File reference

| File | Role |
|---|---|
| `Validation.js` | **New.** The entire validation step described above — re-fetching, corroboration search, identity-coherence check, the adversarial Gemini prompt, verdict sanitization, and empty-pillar fill. |
| `Menu.js` | Adds the "Run Gemini Validation" menu item and the testing-reset wipe function. |
| `SheetSetup.js` / `SheetIO.js` | Extended to create/migrate/default the six new `*_auto_review`/`*_auto_notes` columns and `auto_triage_status`, using the same safe backfill mechanism the base pipeline already relies on. |
| `Batch.js` | Adds `validation` to the automated batch step sequence. |
| `Triage.js` | REVIEW_QUEUE now reflects both triage signals, with a "Why" and per-pillar flag detail column. |
| `PipelineSummary.js` | Tracks the new columns in the summary dashboard. |
| `Config.js`, `Gemini.js`, `Identity.js`, `Serper.js`, `Menu.js` (base items), and everything else | Unchanged from the base pipeline — see `../HCP/README.md` for what each does. |
