# Hidden Champions Pipeline — How It Works

*Last verified against live code: 2026-08-19 (HCP/apps_script/*.js read directly, not from memory).*

This document explains the Hidden Champions ("HC") system end to end: what
goes in, what comes out, what each file does, and how the pieces connect.
It has two parts, per what was asked:

1. **What it currently does** — the working pipeline today, including two
   known gaps that are not yet fixed.
2. **How to implement this in DeepThought's CRM** — a conceptual plan for
   moving this from "Google Sheet + Apps Script" into a real CRM module,
   written without assuming any specifics of the CRM's stack (none were
   available to reference — see the open questions at the end).

---

## Part 1 — What It Currently Does

### The big picture

Hidden Champions identifies and profiles India's "hidden champion"
manufacturers — pre-qualified, often under-the-radar companies with a real
competitive moat — and turns research into a structured, source-cited
company profile. The population is **pre-qualified before this system ever
sees it**: an upstream project already filtered ~1,100 companies "with
potential" out of ~12,000. This system does not gatekeep or score
pass/fail — it characterizes *how much of the company's wealth-creation
engine is present and verified*, and how big the unrealized opportunity
(the "execution gap") is.

There are two stages, two different tools:

```
Stage 1: CHARACTERIZE                    Stage 2: PROFILE
company_name + website                   a characterized HVT row
        │                                        │
        ▼                                        ▼
┌─────────────────────────┐            ┌──────────────────────────┐
│  apps_script/            │  ──────▶  │  target_companies/ or     │
│  (Google Sheet + Apps    │  (human    │  temp_profile_creator/   │
│   Script, automated)     │  reviews   │  (Claude Code session,   │
│                          │  first)    │   hand-built JSON+HTML)  │
└─────────────────────────┘            └──────────────────────────┘
```

This document focuses on **Stage 1** (`apps_script/`) since that's the
part the "validation logic" and "Serper cost log" items below apply to.
Stage 2 is a separate, largely manual Claude Code workflow — see
`HCP/README.md` and `profile_generation/README.md` if that needs its own
write-up later.

### The engine model (what it's actually scoring)

Every company is assessed against a causal chain, not a checklist
(`HCP/context/data-infra-spec.md`):

```
DIFFERENTIATION → protected by a MOAT → kept current by PRODUCT-IMPROVEMENT
CAPABILITY → ALIGNED to a market need → wealth
                                              ↑
        the wealth lives in ALIGNMENT, the upside lives in the EXECUTION GAP
```

Four pillars, each a first-class object with a claim, a source, and a
credibility tier:

| Pillar | What it captures |
|---|---|
| **differentiation** | What makes the offering non-commodity — a *positive* finding only, never inferred from silence |
| **moat** | Typed (`switching-cost \| qualification-lock-in \| IP \| brand \| integration \| scale \| network`) and rated for durability (`compounding \| stable \| eroding`) |
| **product_improvement** | THE discriminator — R&D-engine recognition (DSIR/registry), patents, NPD pipeline, history of product evolution |
| **alignment[]** (up to 3) | An explicit relation: this capability/moat → this specific market need, with a pull strength |
| **gap** | What the capability *could* monetize minus what it captures today — the opportunity |

Plus **identity** fields (CIN, ownership type, sector) that anchor which
real company is being talked about.

### Input

Per company, typed directly into the bound Google Sheet ("Hidden
Champions HVT"):
- `company_name` (column A)
- `website` (column B)

Nothing else is required to start — everything else is filled in by the
pipeline.

### Output

A fully populated row (48 columns, `Config.js` → `HVT_HEADERS`) per
company:
- The 4 pillar claims, each with its own `_source`, `_credibility`,
  `_needs_verify` columns (differentiation/moat/product_improvement), plus
  `moat_type`/`moat_durability`.
- Up to 3 alignment entries (`align1..3_need`, `_source`), each pinned to
  the pillar claim it relates to.
- `gap_summary`, `gap_currently_monetized`, `gap_activation_path`,
  `gap_source`.
- Identity: `cin`, `ownership_type`, `sector` (never `segment_ref` — that's
  a human taxonomy call, not discoverable).
- `unresolved_pillars` and `triage_status` (`auto-confirmed` /
  `needs-review`) — computed in code, never written by the LLM.
- 6 manual review columns (`<pillar>_review`, `<pillar>_notes`) — a human
  marks `Pending / Verified / Rejected` here; the pipeline never writes to
  these itself except to default them to `Pending` on a new row.

Supporting sheets, all rebuilt/maintained by the pipeline: `RAW_EVIDENCE`
(staged search results), `SEARCH_CACHE` (no-TTL cache of past Serper
calls), `GEMINI_LOG` / `SERPER_LOG` (per-call cost ledger),
`COST_SUMMARY`, `PIPELINE_SUMMARY`, `THIN_ROWS`, `REVIEW_QUEUE`,
`EVENT_LOG`.

### How it runs — the menu, in order

Under the **Hidden Champions** custom menu in the Sheet, for a single
company (click its row first):

1. **Run Serper Grounding** — fires ~18-20 web searches (Google-powered,
   via Serper.dev) grouped by pillar, plus a Patents-type and a News-type
   call, plus site-restricted variants once `website` is known. Scrapes
   the #1 result of each query for full page text. Writes everything to
   `RAW_EVIDENCE`.
2. **Resolve Identity** — reads identity-tagged evidence + fetches the
   website directly, drafts `cin` / `ownership_type` / `sector`. Validates
   `cin` against a CIN-format regex; cross-source conflicts get appended
   to `notes` rather than silently resolved one way.
3. **Run Sector-Context Grounding** — a second, additive search pass (only
   enabled once `sector` is known) that searches by *sector name* instead
   of company name, specifically to fill `alignment` and `moat`
   evidence that a company-name-scoped search structurally can't surface
   (generic market-trend/certification content).
4. **Run Gemini Structure** — reads all staged evidence, sends it to
   Gemini to draft the 4 pillars + alignment + gap, then runs it through a
   chain of code-level checks (below) before writing to the row.
5. **Run Company Summary** (optional, not in the default batch) — a 5th
   Gemini call for a narrative paragraph, mostly redundant with what
   `REVIEW_QUEUE` already shows.

For many companies at once, **Batch** menu items do the same 4 steps
company-by-company, working around Apps Script's 6-minute execution cap
by processing one `(company, step)` unit at a time, checkpointing to
Script Properties after every unit, and rescheduling itself via a
time-based trigger. A crash or quota error loses at most one unit, never
the whole batch.

### File-by-file (`HCP/apps_script/`)

| File | Responsibility |
|---|---|
| `Config.js` | Schema (`HVT_HEADERS`), closed vocabularies (moat types/durability, credibility tiers, ownership types...), Serper query templates per pillar, `SELF_PUBLISHED_DOMAINS` / `OFFICIAL_REGISTRY_DOMAINS` guardrail lists |
| `Vocabularies.js` | Closed-vocabulary enums (split out from Config.js — check both before assuming which owns a given list) |
| `Menu.js` | Builds the "Hidden Champions" custom menu; reads `company_name`/`website` off whichever row is selected — no dialogs |
| `SheetSetup.js` | One-time `setupHVTSheet()` — migrates rows **by column name**, never blanket-clears existing data |
| `SheetIO.js` | Row lookup/creation by `company_name`; defaults new rows' `_review` columns to `Pending` |
| `Serper.js` | All web search + scraping. Cache-check → circuit-breaker check → rate-limit → live call → scrape #1 result → write cache + cost log. Tries a **free direct page fetch first**, only pays for Serper's paid scraper if that fails. |
| `Identity.js` | `runIdentityResolution` — drafts CIN/ownership/sector from identity-tagged evidence + a direct site fetch |
| `Gemini.js` | `runGeminiStructure` — the structuring step + the full validation/sanitization chain (next section) |
| `Batch.js` | Multi-company orchestrator: checkpointed, self-rescheduling, skips a failed company rather than halting |
| `CircuitBreaker.js` / `RateLimiter.js` | Halts Serper/Gemini calls on sustained API errors; batch pauses (not hard-stops) and resumes via cooldown |
| `BatchFailures.js` | Buffers failed `(company, stage)` pairs so "Retry Failed Companies" only re-runs the broken handful |
| `TokenLog.js` | Ledger for every Gemini/Serper call (`GEMINI_LOG`/`SERPER_LOG`), including failed calls — a failed call still spent tokens |
| `CostSummary.js` | Rebuilds `COST_SUMMARY` from the logs; also powers "Estimate Batch Cost" (projects cost for N companies from the real observed per-company average) |
| `PipelineSummary.js` | `PIPELINE_SUMMARY` (totals, review-queue count, cost totals) + `THIN_ROWS` (companies with too many empty substantive columns) |
| `Triage.js` | `REVIEW_QUEUE` sheet — just the rows flagged `needs-review` |
| `EventLog.js` | Writes to `EVENT_LOG` so batch progress is visible to anyone with the Sheet open, not just the developer's execution log |
| `Calibration.js` | Calibration-run tooling — check the file directly, not documented in depth here |
| `Test.js` | `testApiKeys()`, `listGeminiModels()` — script-editor-only diagnostics |

### Validation logic that already exists (in `Gemini.js`)

This is the part worth being precise about, since "add validation logic"
is one of the open items below — a meaningful amount already exists. The
`Gemini Structure` step runs the LLM's draft through, in order:

1. `validateAndSanitizeDraft_` — every enum field (moat type/durability,
   credibility tier, `needs_verify`) is checked against the actual closed
   vocabulary; a violation is cleared and logged, never written as-is.
   Also enforces the `SELF_PUBLISHED_DOMAINS` denylist (YouTube, LinkedIn,
   Facebook, X, Medium, PR wires → forced to `self-claimed`) and the
   `OFFICIAL_REGISTRY_DOMAINS` **allowlist** (only a claim sourced from an
   actual registry domain can earn `qualification-gated`/
   `registry-confirmed` — every other domain is forced down regardless of
   what the model said).
2. `upgradeToStrongestSourceForPillars_`
3. `checkPillarSourceFidelity_` — does the drafted claim's language
   actually appear in the cited source's scraped text?
4. `resolveAndFilterAlignment_` — drops any alignment entry whose
   capability doesn't match an already-drafted pillar claim, and
   **force-pins** that entry's source to the matched pillar's own source
   (Gemini is no longer trusted to supply alignment sources itself).
5. `verifyGapAgainstEvidence_` — clears the gap claim if its
   financial-sounding language isn't actually present in the cited
   source's own text.
6. `checkForMissedMoatSignal_` — diagnostic-only: logs a warning if moat
   came back empty but scale/government-backing signal words were seen
   elsewhere in the evidence.
7. `checkGapSourceFidelity_` — word-overlap check between the gap claim
   and the actual scraped content of `gap_source`; logs a warning below
   30% overlap.

`unresolved_pillars` and `triage_status` are then computed in code purely
from the resulting `needs_verify` flags — never LLM-drafted.

### Known gaps (yet to do)

**1. Identity ↔ pillar coherence is not checked.**
A correct, registry-confirmed CIN does not currently protect the pillar
content from being about a *different* company. This has already been
observed in real data: one row carried a genuinely correct CIN (resolved
to the right registered entity, address matched) while all three pillars
described an unrelated company that merely shares a brand name. A second
case: a `product_improvement` patent claim was actually a namesake
individual's unrelated patents (assignee ≠ the resolved company), not
caught because inventor-name matching was treated as if it were
entity-matching. Identity resolution (step 2) and pillar drafting (step 4)
currently run independently and never cross-check each other.

What's needed: (a) treat any inventor/person-name patent search as a
name-match, not an entity-match — only accept it into `product_improvement`
if the patent's assignee field matches the resolved company; (b) add a
coherence check between the resolved CIN's registered business activity
and the drafted pillar claims before writing a row, flagging (not
silently writing) when they plainly diverge; (c) never trust a bare brand
name as the identity anchor — CIN or registered address only, since short
generic Indian company names collide often (at least 4 unrelated real
companies trade as "Vikalp").

**2. The Serper cost log doesn't count actual credits — it uses a flat
per-call estimate.**
`TokenLog.js`:
```js
const SERPER_COST_PER_CALL = 0.001;
```
Every single Serper call — a plain `/search`, a `/patents` or `/news`
typed call, a `scrape.serper.dev` call — is logged at the same flat
$0.001, regardless of what it actually consumed. Serper's actual billing
is credit-based and different endpoints/result-counts can consume
different amounts; a flat per-call number is a rough estimate, not a
measurement. `CostSummary.js`'s own comments already flag this ("confirm
these against your real Gemini/Serper invoices before trusting
COST_SUMMARY numbers for a real budget decision"). This is what's meant
by "credits are not counted" — the fix is to read Serper's actual
response for a credit-usage field (if the API surfaces one) or otherwise
reconcile the estimate against Serper's own dashboard, rather than
assuming every call costs the same.

---

## Part 2 — How to Implement This in DeepThought's CRM

No specifics about DeepThought's CRM (stack, data model, existing API
surface) were available to reference for this document, so this is
written conceptually — what needs to move, what needs to exist, what
decisions have to be made — rather than against a concrete
framework/schema. Treat it as a planning checklist to work through with
whoever owns the CRM's architecture, not a spec ready to execute.

### What's actually being moved

Not "the Google Sheet," but three things currently living inside it:

1. **A schema** — the 48-column HVT structure, currently enforced only by
   `Config.js`'s `HVT_HEADERS` array and column-name-based migration.
2. **A pipeline** — Serper search → identity resolution → sector-context
   search → Gemini structuring → the validation chain — currently
   triggered by Apps Script menu clicks / a self-rescheduling trigger loop.
3. **A review workflow** — the `_review`/`_notes` columns, `triage_status`,
   `REVIEW_QUEUE` — currently a human editing cells directly in the Sheet.

Each of these maps to a different kind of CRM work, and they don't have
to move together.

### Conceptual architecture

```
┌────────────────┐      ┌──────────────────────────┐      ┌─────────────────┐
│ CRM data model  │◀────▶│  Characterization engine  │◀────▶│ External APIs    │
│ (company record,│      │  (the actual pipeline —   │      │ Serper, Gemini   │
│  pillar fields, │      │  can be a re-hosted/       │      │ (same calls as   │
│  review state)  │      │  re-implemented version    │      │  today)          │
└────────────────┘      │  of Serper.js/Identity.js/ │      └─────────────────┘
        ▲                 Gemini.js's logic)          
        │               └──────────────────────────┘
        │
┌────────────────┐
│ Review UI       │   (replaces "type Verified/Rejected into a cell")
│ + REVIEW_QUEUE  │
│ equivalent view │
└────────────────┘
```

### Decisions that need making before implementation starts

- **Where does the pipeline execute?** Apps Script's 6-minute cap and
  self-rescheduling-trigger pattern is a workaround specific to Google's
  runtime — it should not be carried over as-is. A CRM-hosted job
  queue/worker (whatever the CRM already uses for background jobs, or a
  small dedicated service if it doesn't have one) is the natural
  replacement, and removes the whole checkpoint-and-reschedule mechanism
  in `Batch.js` — a real backend process doesn't need to fake
  statelessness the way an Apps Script trigger does.
- **Does the schema become first-class CRM fields, or a structured
  blob attached to a company record?** The 4 pillars + alignment + gap
  are naturally nested/repeating (up to 3 alignment entries, each with its
  own source) — worth checking whether the CRM's data model handles
  nested/structured custom fields natively, or whether this is better
  represented as a linked child-record type ("pillar claims" as records
  related to a company) rather than forced into flat custom fields.
- **How does the review workflow become a UI instead of spreadsheet
  cells?** Minimum needed: per-pillar Verified/Rejected/Pending state,
  a notes field, and a filtered view equivalent to `REVIEW_QUEUE` (only
  `needs-review` rows). This is the most CRM-native part of the whole
  system — likely the easiest piece to build well, since "a queue of
  records awaiting human review" is a common CRM pattern already.
- **Where do the API keys and cost logs live?** `SERPER_API_KEY` /
  `GEMINI_API_KEY` currently sit in Apps Script's Script Properties — in
  a CRM this becomes whatever secret-management the CRM already uses.
  Cost logging (`TokenLog.js`/`CostSummary.js`) should be rebuilt from
  scratch anyway once fix #2 from Part 1 is addressed — no reason to port
  the flat-rate estimate forward into a new system.
- **Does batching stay batch-shaped, or become event-driven?** The
  current design processes companies in an explicit batch because Apps
  Script forces that shape. A CRM might instead trigger characterization
  the moment a company record is created/imported (one job per company,
  queued normally) — worth deciding whether "run the whole HVT population
  through in a batch" is still a real use case in the CRM, or whether
  it becomes purely per-company, triggered on demand.
- **What happens to the caching layer (`SEARCH_CACHE`)?** Its whole
  premise is "no TTL, stays stable until explicitly cleared" — reasonable
  for a 1,100-company one-time characterization pass. Confirm this
  assumption still holds in the CRM's usage pattern (e.g. if companies get
  re-characterized periodically, a no-TTL cache would silently serve stale
  search results forever).
- **The two Part-1 gaps should be fixed before or during the port, not
  after.** Porting the identity/pillar-coherence gap as-is means the CRM
  inherits the same failure mode (a correct CIN sitting next to pillar
  content about an unrelated company) at whatever scale the CRM runs at.
  Same for the cost-log estimate — better to build accurate cost tracking
  into the CRM version from day one than migrate a known-inaccurate one.

### Suggested sequencing

1. Fix the two Part-1 gaps in the current Apps Script version first (small,
   contained, already well-understood — see Part 1 for the concrete fix
   each needs).
2. Re-implement the pipeline logic (Serper calls, identity resolution,
   Gemini structuring + validation chain) as a standalone
   service/module, independent of Google Sheets — this is the reusable
   core regardless of which CRM it plugs into.
3. Design the CRM-side data model for the schema (flat fields vs. child
   records) and the review UI, in parallel with step 2.
4. Wire the pipeline to write into the CRM's data model instead of sheet
   cells; wire the review UI to read/write the same records the pipeline
   populates.
5. Build cost tracking into the new service from the start, using actual
   Serper/Gemini usage data rather than porting `TokenLog.js`'s flat-rate
   estimate.

### Open questions to answer with the CRM owner

- What background-job/queue mechanism does the CRM already have (if any)?
- Does the CRM support structured/nested custom fields, or only flat ones?
- Is there an existing "review queue" or approval-workflow pattern in the
  CRM to reuse, or does one need to be built?
- Where do API secrets live for other CRM integrations today?
- Is per-company on-demand characterization the target usage pattern, or
  batch runs over an imported list (or both)?
