// Fast-growing keyword/domain lists, split out of Config.js — these change on
// their own schedule (a new registry domain or keyword phrasing found nearly
// every company tested) separate from schema/logic constants. Config.js
// keeps the schema; this file is where you look first when a new domain or
// phrasing needs adding.

// Platforms where hosting ≠ independence: a claim here is still the
// company's own words unless the content shows a named independent
// journalist/analyst doing original reporting. Enforced in code so a model
// that mislabels a YouTube video as "third-party" gets overridden every time,
// not just when the prompt happens to catch it.
const SELF_PUBLISHED_DOMAINS = [
  'youtube.com', 'youtu.be', 'facebook.com', 'linkedin.com', 'twitter.com', 'x.com',
  'instagram.com', 'medium.com', 'prnewswire.com', 'businesswire.com', 'globenewswire.com'
];

function isSelfPublishedDomain_(url) {
  return SELF_PUBLISHED_DOMAINS.some(function (domain) {
    return String(url).toLowerCase().indexOf(domain) !== -1;
  });
}

// "qualification-gated" / "registry-confirmed" credibility requires the
// source to actually BE the certifying/registry body — a company's own site
// describing a certification is self-claimed, not qualification-gated, per
// data-infra-spec.md: "Company sites and aggregators cannot settle it."
// Enforced as an allowlist (only these domains can earn the higher tier),
// not a denylist, since the space of legitimate registries is small and
// enumerable while the space of non-registry sites is not.
// Expanded 2026-07-10 after Eldorado Agritech's product_improvement claim
// (sourced from indiascienceandtechnology.gov.in, a real DST portal listing
// DSIR-registered institutions) got force-downgraded to self-claimed simply
// because that specific domain wasn't on the list yet — same failure mode
// that added s3waas.gov.in/nstmis-dst.org earlier for Natesan. Added the
// adjacent government registries proactively (BIS, Udyam, DST itself) rather
// than waiting to hit each one company-by-company, plus omri.org and
// nabl-india.org — non-.gov.in but the same category: the actual certifying
// body's own domain, already seen in real claims (Jaishil's OMRI listing,
// Natesan's NABL certification).
// texmin.gov.in/pli.texmin.gov.in added 2026-08-07 for TechnoSport, whose moat
// is confirmed by the Ministry of Textiles' own published PLI applicants list
// (pli.texmin.gov.in/Applicants_list, Row 69, naming Techno Sportswear and its
// subsidiary against one approved project at Sipcot Perundurai). Same failure
// mode as Eldorado and Natesan above — a genuine approving-authority domain
// that would otherwise be force-downgraded to self-claimed, except here it
// would have destroyed the ONLY registry-confirmed pillar on the row. Both
// forms are listed for the same self-documenting reason as bis.gov.in /
// standards.bis.gov.in, even though the substring match makes the parent
// domain sufficient on its own. First non-R&D registry on this list: PLI is an
// incentive-approval body rather than a certification or IP registry, so it
// widens what "registry-confirmed" covers — worth remembering when the next
// sector-specific approval scheme (Ministry of Textiles or otherwise) shows up.
const OFFICIAL_REGISTRY_DOMAINS = [
  'dsir.gov.in', 'ipindia.gov.in', 'ipindiaservices.gov.in', 'mca.gov.in', 'gst.gov.in',
  's3waas.gov.in', 'nstmis-dst.org', 'indiascienceandtechnology.gov.in', 'dst.gov.in',
  'bis.gov.in', 'standards.bis.gov.in', 'udyamregistration.gov.in', 'omri.org', 'nabl-india.org',
  'texmin.gov.in', 'pli.texmin.gov.in'
];

function isOfficialRegistryDomain_(url) {
  return OFFICIAL_REGISTRY_DOMAINS.some(function (domain) {
    return String(url).toLowerCase().indexOf(domain) !== -1;
  });
}

// If a pillar's own claim TEXT mentions one of these, needs_verify must be
// "yes" regardless of what the model set — a deterministic backstop for rule
// 6, since the model doesn't reliably re-scan its own drafted claim for
// registry-trigger words when they appear as a qualifier rather than being
// the whole claim (seen: "patented" inside a differentiation sentence not
// tripping needs_verify, even though the identical fact correctly did when
// it was the entire product_improvement claim).
// 'dgft' added 2026-07-14 — Garware's moat claim ("Four Star Export House by
// DGFT, Ministry of Commerce") is a genuine, checkable government
// classification (Directorate General of Foreign Trade) but stayed
// needs_verify: no since only dsir/mca were covered, not other government
// registry bodies. 'uidai'/'stqc' added same day — Mantra Softech's moat
// claim ("UIDAI-certified and STQC-tested") is two more real government
// certification bodies (Unique Identification Authority of India,
// Standardisation Testing and Quality Certification) missed the same way.
const REGISTRY_TRIGGER_KEYWORDS = ['patent', 'patented', 'patents', 'dsir', 'mca', 'dgft', 'uidai', 'stqc', 'as9100', 'iatf', 'gots'];

// Market-exclusivity/superlative claims are just as much an unverified lead
// as a registry-checkable one — they need a different verification channel
// (industry reports, competitor benchmarking) rather than a government
// registry, but they're still self-claimed facts, not confirmed ones. Found
// on Tatva Chintan: "sole commercial supplier" / "world's second-largest
// manufacturer" claims came back needs_verify: no because they don't mention
// patent/DSIR/MCA — this closes that gap the same way REGISTRY_TRIGGER_KEYWORDS
// does, just for a different class of claim.
// 'largest' is deliberately a single bare word, not multi-word phrases like
// "world's largest" or "largest producer" — those are redundant since they
// all contain 'largest' as a substring anyway, and the narrower phrasing
// missed Jaishil's "one of largest" variant entirely. One root word catches
// every surrounding phrasing instead of needing to enumerate each one.
// 'leading' and "india's first" added 2026-07-10 — Spray Engineering's
// "India's first 4G ethanol plant" (reversed word order from 'first in
// india') and Garware's "world's leading fishing nets supplier" ('leading'
// as a near-synonym of largest/biggest) both came back needs_verify: no
// despite being unverified superlative claims, same gap class as the
// 'largest' bare-word fix below.
const EXCLUSIVITY_TRIGGER_KEYWORDS = [
  'sole supplier', 'sole commercial supplier', 'sole licensee', 'only manufacturer',
  'only company', 'largest', 'biggest', 'leading', 'first in india', "india's first",
  'first privately developed', 'pioneer'
];

// gap_currently_monetized/gap_summary kept asserting financial distress
// ("cash flow remains under pressure") across four straight runs despite a
// targeted prompt fix — the model's tendency wasn't tied to one inference
// path, so a single negative example didn't hold. This is the code-level
// cross-check: if gap text uses this language but the cited source's own
// content doesn't, the claim is treated as unverified and cleared.
const GAP_FINANCIAL_KEYWORDS = ['cash flow', 'profit', 'loss', 'revenue', 'financial', 'distress', 'debt', 'liquidity'];

// Signals that a "scale"-type moat (rule 4 — uniquely large/first/backed at a
// scale competitors can't quickly replicate) was likely available in the
// evidence but got used elsewhere (e.g. gap_activation_path) instead of
// being drafted as a moat entry. Used only for a diagnostic warning, not to
// auto-fill moat — that judgment stays with the analyst.
const MOAT_SCALE_SIGNAL_KEYWORDS = [
  'government approval', 'government-backed', 'government grant', 'financial assistance',
  'crore', 'sole supplier', 'sole licensee', 'first privately developed', 'first in india',
  'exclusive', 'only company'
];

// Diagnostic only, same posture as checkForMissedMoatSignal_ (Gemini.js) —
// scans a company's evidence for .gov.in domains not already on the
// OFFICIAL_REGISTRY_DOMAINS allowlist, so a new registry gap surfaces as a
// logged signal instead of only being noticed by chance during a manual
// review (how every prior addition to this list was actually found).
function checkForUnknownRegistryDomain_(snippets, companyName, runId) {
  const seen = {};
  snippets.forEach(function (s) {
    const url = String(s.source_url || '').toLowerCase();
    if (!url || seen[url]) return;
    seen[url] = true;
    if (url.indexOf('.gov.in') === -1) return;
    if (isOfficialRegistryDomain_(url)) return;

    const domainMatch = url.match(/^https?:\/\/([^/]+)/);
    const domain = domainMatch ? domainMatch[1] : url;
    const message = 'Evidence for "' + companyName + '" cites a .gov.in domain not on OFFICIAL_REGISTRY_DOMAINS: ' + domain + ' (' + s.source_url + ') — consider adding it if it is a genuine registry/certifying body.';
    Logger.log('INFO: ' + message);
    if (typeof EventLog !== 'undefined') EventLog.info(runId, companyName, '', 'registry_domain_check', 'unknown-domain', message);
  });
}
