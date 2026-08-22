const HVT_SHEET_NAME = 'HVT';
const RAW_EVIDENCE_SHEET_NAME = 'RAW_EVIDENCE';
// Same column shape as RAW_EVIDENCE, deliberately a SEPARATE sheet rather
// than overloading RAW_EVIDENCE itself — RAW_EVIDENCE's disposability is
// load-bearing elsewhere (setupHVTSheet() clears it fully; runSerperGrounding
// clears it per-company on every deliberate re-run so a human gets fresh
// results on demand). A cache needs the opposite semantics (persist, reuse),
// so keeping them separate means neither call site has to know which mode
// it's in. No TTL — deliberately: since the same query has been confirmed to
// return different live results run-to-run, a no-TTL cache also converts
// "noisy every run" into "noisy once, then stable" as a side benefit.
const SEARCH_CACHE_SHEET_NAME = 'SEARCH_CACHE';

const REVIEW_STATUS_OPTIONS = ['Pending', 'Verified', 'Rejected', 'Needs registry check'];
const REVIEW_COLUMNS = ['differentiation_review', 'moat_review', 'product_improvement_review'];

// Machine-verdict enum for the Gemini Validation round (Validation.js) — kept
// visually and structurally distinct from REVIEW_STATUS_OPTIONS (human-only)
// so a *_auto_review verdict can never be mistaken for a human *_review
// verdict, even by someone skimming the sheet without reading headers.
// Gemini Validation NEVER writes to REVIEW_COLUMNS/*_review/*_notes.
// 'Inferred — Low Confidence' is a labeled best-guess fill for a pillar that
// had NO claim at all and thin/partial (but non-zero) evidence — see
// Validation.js's attemptEmptyPillarFill_(). Deliberately NOT a "needs
// review" trigger on its own: it is an accepted, visible judgment call, not
// an unresolved doubt like Needs Verify/Rejected/Contradiction Found are.
const AUTO_REVIEW_STATUS_OPTIONS = ['Not Run', 'Verified', 'Needs Verify', 'Rejected', 'Contradiction Found', 'Inferred — Low Confidence'];
const AUTO_REVIEW_COLUMNS = ['differentiation_auto_review', 'moat_auto_review', 'product_improvement_auto_review'];

// Full §3 schema, with the 6 review/notes columns inserted right after each
// pillar's existing columns (differentiation, moat, product_improvement),
// plus 6 auto_review/auto_notes columns (Validation.js's machine verdicts)
// immediately after each pillar's human review pair, and auto_triage_status
// alongside triage_status — a second, independent honesty signal computed
// from the validation pass rather than the structuring pass (see Validation.js).
const HVT_HEADERS = [
  'company_name', 'website', 'cin', 'ownership_type', 'sector', 'segment_ref',
  'differentiation', 'differentiation_source', 'differentiation_credibility', 'differentiation_needs_verify',
  'differentiation_review', 'differentiation_notes', 'differentiation_auto_review', 'differentiation_auto_notes',
  'moat_type', 'moat_durability', 'moat', 'moat_source', 'moat_credibility', 'moat_needs_verify',
  'moat_review', 'moat_notes', 'moat_auto_review', 'moat_auto_notes',
  'product_improvement', 'product_improvement_source', 'product_improvement_credibility', 'product_improvement_needs_verify',
  'product_improvement_review', 'product_improvement_notes', 'product_improvement_auto_review', 'product_improvement_auto_notes',
  'align1_capability', 'align1_need', 'align1_pull', 'align1_source',
  'align2_capability', 'align2_need', 'align2_pull', 'align2_source',
  'align3_capability', 'align3_need', 'align3_pull', 'align3_source',
  'gap_latent', 'gap_currently_monetized', 'gap_summary', 'gap_activation_path', 'gap_source',
  'proof_type', 'proof_status', 'proof_source', 'proof_pillar',
  'unresolved_pillars', 'triage_status', 'auto_triage_status', 'company_summary', 'notes'
];

// Computed in code from unresolved_pillars, never LLM-drafted — same
// discipline as unresolved_pillars itself. Lets a human work a REVIEW_QUEUE
// of only the needs-review rows instead of opening all 1,100 by hand.
const TRIAGE_STATUS = { AUTO_CONFIRMED: 'auto-confirmed', NEEDS_REVIEW: 'needs-review' };

const RAW_EVIDENCE_HEADERS = ['company_name', 'pillar', 'query', 'source_url', 'title', 'snippet', 'full_text', 'retrieved_at'];

// §4 query groups, plus an `identity` group whose results are leads only
// (staged to RAW_EVIDENCE, never auto-written to cin/ownership_type — those
// stay a manual MCA check per §5). {company} is substituted with the exact
// company name.
const SERPER_QUERIES_BY_PILLAR = {
  differentiation: [
    '"{company}" patents',
    '"{company}" DSIR',
    '"{company}" "first in India" OR "only" OR "pioneer"',
    '"{company}" proprietary technology'
  ],
  moat: [
    '"{company}" certification OR approval OR AS9100 OR IATF OR GOTS',
    '"{company}" sole supplier OR qualified vendor',
    // Added after Elixir Ingredients (2026-08-21): the two queries above are
    // certification-shaped and never surfaced that Elixir's jaggery process
    // is licensed from a SPECIFIC IIT Bombay patent (Grant No. 525046,
    // rnd.iitb.ac.in) — a real qualification-lock-in/IP moat that a
    // certification-only query can't find, since a technology-licensing
    // moat is framed as "patent"/"licensed technology", not "certified".
    // Distinct from differentiation's '"{company}" patents' query above:
    // that one looks for the COMPANY holding its own patents; this looks
    // for the company being LICENSED to use someone else's — a different
    // relationship (per data-infra-spec.md, only resulting lock-in counts
    // as moat, so this must find the licensing/exclusivity language, not
    // just any patent mention, to actually be moat-shaped evidence).
    '"{company}" patent OR licensed technology OR "technology license" OR exclusive license'
  ],
  product_improvement: [
    '"{company}" R&D center',
    '"{company}" DSIR recognized',
    '"{company}" new product launch'
  ],
  alignment_gap: [
    '"{company}" export',
    '"{company}" expansion OR hiring OR new facility'
  ],
  identity: [
    '"{company}" CIN',
    '"{company}" "Ministry of Corporate Affairs" OR Zauba OR Tofler',
    '"{company}" GST OR GSTIN OR Udyam OR proprietorship'
  ]
};

// Site-restricted variants, fired only when a website is supplied. Pushes the
// company's own deeper pages (e.g. a real technology page) ahead of thin
// social posts, which is what generic web search kept surfacing instead.
const SERPER_SITE_QUERIES_BY_PILLAR = {
  differentiation: ['site:{domain} patent OR proprietary OR technology'],
  moat: ['site:{domain} certification OR approval OR "sole supplier"'],
  product_improvement: ['site:{domain} R&D OR "research and development" OR "new product"']
};

function extractDomain_(website) {
  return String(website).replace(/^https?:\/\//i, '').replace(/^www\./i, '').replace(/\/.*$/, '');
}

// Sector-aware queries, fired only once `sector` is known (via Identity
// Resolution or manual entry) — searched by SECTOR NAME, not company name, so
// results surface genuine industry/market-trend and certification content
// instead of only what the company says about itself. This is one mechanism
// for two confirmed, repeated gaps: the alignment-search gap (empty/weak on
// Tatva Chintan, Sandstorm, Natesan — alignment_gap queries never reached
// real market-trend content) and the moat-certification gap (OMRI missed for
// Jaishil, an agrochemical company, by the auto/aerospace-biased
// AS9100/IATF/GOTS keyword list). Rather than pre-enumerating every possible
// certification acronym or market-trend phrase per sector — which can never
// be complete — this lets the search engine's own relevance ranking do the
// sector-specific mapping: "{sector} certification" naturally surfaces
// whatever registry actually matters for that sector.
const SERPER_SECTOR_QUERIES = {
  alignment_gap: ['"{sector}" market demand OR growth OR outlook India'],
  moat: ['"{sector}" certification OR standard OR compliance OR approval']
};

// Closed vocabularies from data-infra-spec.md §2. Given to Gemini in the
// prompt AND enforced in code after the response comes back — a single
// prompt instruction degrades with model drift; a code-level check doesn't.
const MOAT_TYPES = ['switching-cost', 'qualification-lock-in', 'IP', 'brand', 'integration', 'scale', 'network'];
const MOAT_DURABILITY = ['compounding', 'stable', 'eroding'];
const CREDIBILITY_TIERS = ['self-claimed', 'third-party', 'qualification-gated', 'registry-confirmed'];
const NEEDS_VERIFY_OPTIONS = ['yes', 'no'];

// ownershipType enum from data-infra-spec.md's "Identity & sourcing" section,
// plus 'sole-proprietorship' (added 2026-07-07) — the original 5 assume a
// Companies-Act-registered entity; a sole proprietorship (no CIN, GST/Udyam-
// registered instead) doesn't fit any of them. Found on Sandstorm Equipment Co.
// Plus 'partnership' (added 2026-07-08) — distinct from sole-proprietorship:
// no CIN either, but 2+ owners rather than one individual. Found on Jaishil,
// where evidence showed no CIN (same GST/no-CIN pattern as a proprietorship)
// but also no named individual owner — the model had been forced into
// "sole-proprietorship" as the closest available value despite no evidence
// actually confirming single ownership.
const OWNERSHIP_TYPES = ['listed', 'unlisted-public', 'private', 'government-PSU', 'group-holding', 'sole-proprietorship', 'partnership'];

// Real Indian CIN format: 1 letter (L/U) + 5 digits + 2-letter state code +
// 4-digit year + 3-letter type (PLC/OPC/PTC/...) + 6 digits = 21 chars.
// e.g. U00000CH2004PLC027625. A structural check, not proof the CIN is
// correct — just proof it isn't obviously fabricated.
const CIN_FORMAT_REGEX = /^[LU]\d{5}[A-Z]{2}\d{4}[A-Z]{3}\d{6}$/;

// data-infra-spec.md: alignment[] carries pullStrength (must-have | strong |
// nice-to-have). proof_status per §3 is `not-found | claimed-verify |
// confirmed | registry-access-blocked` — but "confirmed" and
// "registry-access-blocked" can only ever be set by a human who actually did
// the registry check (§5), so Gemini's draft is restricted to the other two.
const PULL_STRENGTHS = ['must-have', 'strong', 'nice-to-have'];
const PROOF_STATUS_OPTIONS = ['not-found', 'claimed-verify', 'confirmed', 'registry-access-blocked'];
const PROOF_STATUS_DRAFTABLE = ['not-found', 'claimed-verify'];

// Evidence tags Gemini is shown (as informational metadata, not authoritative
// classification — see buildGeminiPrompt_ rule 3). Distinct from
// GEMINI_DRAFT_PILLARS (Gemini.js), which is the set of OUTPUT fields.
const EVIDENCE_INPUT_TAGS = ['differentiation', 'moat', 'product_improvement', 'news', 'patents', 'alignment_gap'];

function getScriptProp_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('Missing Script Property: ' + key + '. Set it via Project Settings > Script Properties in the Apps Script editor.');
  }
  return value;
}
