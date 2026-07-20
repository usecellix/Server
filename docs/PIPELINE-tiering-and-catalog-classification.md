# CELLIX — Agent Pipeline Redesign: Does This Use-Case Catalog Need Planner→Executor→Verifier?

*Answers, using your actual `cellix-basic-usecases.html` catalog: (1) would the previous fixes get you to Shortcut/Vesence/Crunched-level smoothness, and (2) which of these 180+ operations genuinely need the 3-agent pipeline.*

---

## 1. Short answer to both questions

**No, the earlier fixes alone would not get you there — and now we can say exactly why, with numbers.** Your own catalog is dominated by single-action operations. Running "freeze row 1," "sum column D," "flag rows where GSTIN doesn't start with 32," and "calculate 18% GST on column D" through the *same* Planner-decompose → Executor → Verifier pipeline as a genuinely compound multi-sheet task is architecturally treating 95% of your product surface as if it were the hard 5%. That's not a bug to patch — it's a scope mismatch between the pipeline's design (built for decomposition) and the traffic it's actually serving (mostly atomic).

**Answer to "do we need the pipeline for these use cases": for the overwhelming majority, no.** Section 2 classifies your entire catalog. The rough split:

| Runtime tier | Share of your catalog | Needs Planner decomposition? |
|---|---|---|
| Tier 0 — deterministic, no LLM | ~25-30% | No |
| Tier 1 — single LLM call, no verify | ~20-25% | No |
| Tier 2 — generate + verify (2-hop, no decompose) | ~35-40% | No |
| Tier 3 — full Planner→Executor→Verifier | ~5-10%, and only when *composed* | Yes |

Almost nothing in your catalog is Tier 3 **as a category**. Tier 3 only shows up when a user *chains* several catalog items in one request ("clean duplicates, then sort, then add a pivot table") — it's an emergent property of composition, not a property of any single use case in your document.

---

## 2. Full classification of your catalog

### Tier 0 — Deterministic, zero LLM calls (should already be regex/shortcut lane, or needs to be added to it)

Pure structural operations with no interpretation required once the target is identified — these are 1:1 mappable to a single Office.js call.

| Your category | Examples from your doc |
|---|---|
| Sheet Operations | create/rename/delete/hide/show/color sheet, reorder sheets |
| Row & Column Operations | insert/delete row or column, resize, hide/unhide |
| Basic Cell Formatting | bold, color, borders, number format, font |
| Named Ranges (creation) | define a named range for an explicit selection |

**These should never touch an LLM at all if the target is explicit** ("bold A1:C1", "freeze row 1", "hide column F") — pure regex/keyword extraction, same lane as your existing `INSTANT_SHORTCUT_PATTERNS`. If the target is *implicit* ("bold the header row" — which row is that?), it needs exactly one small LLM call to resolve the target, not three agents.

### Tier 1 — One LLM call, no verification needed (low-stakes, easily reversible, non-numeric)

| Your category | Examples |
|---|---|
| Copy, Paste & Fill | fill down, copy formatting, paste special |
| Data Sorting | sort by column, multi-column sort |
| Data Filtering | filter by condition |
| Find & Replace | replace text/value |
| Conditional Formatting | highlight cells meeting a condition |
| Q&A.1 Explain Requests | "what does this formula do", "describe this sheet" — already read-only, no write at all |

One model call turns the message into one action. No decomposition (there's only ever one action), and verification adds little value here because these operations are visually self-evident and trivially reversible — a misapplied sort or a wrong highlight color is obvious and undoable in one undo, unlike a wrong tax figure buried in a cell.

### Tier 2 — Generate + Verify (two agents, no Planner/decompose step)

This is your largest bucket, and it's the one most worth designing carefully, because **"generate a formula" is not the same architectural problem as "decompose a multi-step task"** — yet it currently goes through the same pipeline as if it were.

| Your category | Examples |
|---|---|
| Basic Math / Logical Formulas | `=D2*0.18`, `=IF(...)` |
| Lookup Formulas | INDEX/MATCH, VLOOKUP |
| Text Formulas | SEARCH/SUBSTITUTE combos |
| Date & Time Formulas | MONTH/YEAR extraction |
| Conditional Formulas (SUMIF/COUNTIF/AVERAGEIF) | your Q&A.5 table is literally this |
| Array & Dynamic Formulas | array formulas |
| Error Detection & Fix (Q&A.3) | diagnose #REF/#N/A/#VALUE, propose fix |
| Data Validation | dropdown/rule creation |
| Duplicate Detection & Cleanup | find + optionally remove duplicates |
| Q&A.4 Data Analysis Questions | sum/count/max/average/percentage — read-only, single computed answer |
| Pivot Tables (Basic) | one Office.js `PivotTable` creation call with defined fields |
| Basic Charts | one Office.js `Chart` creation call with a defined range/type |

Every single row in your own Q&A.5 table (`"Calculate 18% GST" → =D2*0.18`, verified by testing on sample rows) is **exactly the two-hop shape**: generate the formula, verify it against sample data, apply. There is nothing to decompose — it's one formula, applied to one range. Sending this through `PlannerAgent.plan()` first produces a plan with exactly one subtask, which is pure overhead (an LLM call whose entire output is "there is one thing to do"), and — per the earlier bug investigation — is also where an over-cautious clarification prompt gets an unnecessary opportunity to fire.

**Note on pivot tables/charts:** these look "complex" but Office.js exposes them as single structured API calls once you know the source range, fields, and chart type — the complexity is in correctly specifying that one call, not in sequencing multiple calls. Tier 2, not Tier 3.

### Tier 3 — Genuinely needs Planner→Executor→Verifier

This only happens when a request **spans multiple targets with real interdependency** — not because any one operation is hard, but because the *sequence and dependencies between operations* matter and the LLM needs to reason about ordering, cross-sheet consistency, or a target that depends on a prior step's output.

Real examples of what actually belongs here (none of them are individually novel — they're compositions of Tier 0-2 items):
- "Clean up duplicates in column A, then sort by date, then add a total row, then create a pivot summarizing by supplier" — 4 dependent steps, later steps need earlier steps' output.
- "Find all invoices with invalid GSTIN across every sheet in this workbook and consolidate them into a new 'Exceptions' sheet" — cross-sheet, creates a new target that downstream steps write into.
- Anything from your **future** GST/TDS/Ind-AS domain engine work — reconciliation, ITC computation — which is inherently multi-stage (normalize → match → categorize → compute → write → cite).

**Nothing in your current basic-use-cases document, taken as a single request, requires this.** It only becomes necessary when a user's *single message* asks for several of these chained together.

---

## 3. The redesigned pipeline

Replace the current "every write → Planner→Executor→Verifier" default with a **complexity-gated dispatcher**. This extends (doesn't replace) the `write-direct` / `write-planned` split from the previous design doc, now with your full catalog mapped onto it so the classifier has concrete ground truth to be built and tested against.

```
route='write' (from LlmRouterService)
        │
        ▼
Complexity classifier
  (regex pre-filter, §3 below, backed by the router LLM's new "complexity" field)
        │
        ├─ TIER 0: target fully explicit, maps to one known action type
        │          → direct Office.js call, NO LLM AT ALL
        │
        ├─ TIER 1: single action, target needs light resolution, low stakes
        │          → ONE LLM call → action → apply (no verify)
        │
        ├─ TIER 2: single formula/computation/structured object (pivot/chart)
        │          → Generate agent → Verify agent → preview → apply
        │          (2 agents, NOT 3 — no Planner, because there's nothing to decompose)
        │
        └─ TIER 3: multiple dependent actions, cross-sheet, or genuinely compound
                   → Planner → Executor → Verifier (existing pipeline, unchanged)
                   (reserved for compositions and future domain-engine work)
```

**Why Tier 2 should be "Generate → Verify," not "Planner → Executor → Verifier":** the Planner's job is producing a `subtasks[]` array with `dependsOn` ordering. A single formula has no subtasks to order. Skipping straight to a Generate agent (equivalent to today's `ExecutorAgent`, called directly) followed by the existing `VerifierAgent` gets you the exact same safety property (independent verification before it touches the sheet) for one-third less latency and one fewer opportunity for a clarification-prone prompt to fire.

**Where the classifier lives:** extend `LlmRouterService`. You already have the regex fast lane pattern (`INSTANT_SHORTCUT_PATTERNS`) and the `quickDataCheck` keyword lane — this is a third lane of the same shape, sitting between them and the LLM router:

```typescript
// New: maps directly to Tier 0/1/2 action types, bypassing Planner entirely.
const SINGLE_ACTION_PATTERNS: Array<{ pattern: RegExp; tier: 0 | 1 | 2; actionHint: string }> = [
  { pattern: /\b(sort|filter)\b.*\bby\b/i, tier: 1, actionHint: 'SORT_OR_FILTER' },
  { pattern: /\bfind\s*(and)?\s*replace\b/i, tier: 1, actionHint: 'FIND_REPLACE' },
  { pattern: /\b(highlight|conditional format)\b/i, tier: 1, actionHint: 'CONDITIONAL_FORMAT' },
  { pattern: /\bcalculate\b.*%|\b=|\bformula\b|\bif\s.*then\b/i, tier: 2, actionHint: 'FORMULA_GEN' },
  { pattern: /\bpivot table\b/i, tier: 2, actionHint: 'PIVOT_TABLE' },
  { pattern: /\bchart\b|\bgraph\b/i, tier: 2, actionHint: 'CHART' },
  { pattern: /\bduplicate\b/i, tier: 2, actionHint: 'DUPLICATE_CHECK' },
  { pattern: /\bvalidation\b|\bdropdown\b/i, tier: 2, actionHint: 'DATA_VALIDATION' },
  // ... one row per catalog category from §2, Tier 0/1/2 buckets
];

// Compound signal — if ANY of these match alongside a single-action pattern,
// escalate to Tier 3 regardless of what else matched.
const COMPOUND_SIGNALS = /\band then\b|\bafter that\b|,\s*(then|and)\s|\bfor each sheet\b|\backross (all|every) sheets?\b/i;
```

This is not meant to be exhaustive on day one — it's meant to cover the ~180 operations you've already documented, since you already know their shape. Anything that doesn't match falls through to the LLM router's `complexity` field (from the previous design doc) as the semantic fallback, and only *that* uncertain remainder should ever reach the full Planner.

---

## 4. What this does to "smoothness" — concrete latency budget

Rough LLM-call counts per tier, which is the dominant cost of perceived latency (each hop is a network round trip + inference time):

| Tier | LLM calls | Rough latency (typical) | Current behavior |
|---|---|---|---|
| 0 | 0 | <50ms | Currently: 3 calls, 3-8s (if it reaches Planner at all) |
| 1 | 1 | 150-400ms | Currently: 3 calls, 3-8s |
| 2 | 2 (generate + verify) | 800ms-2s | Currently: 3 calls, 3-8s (Planner call is pure overhead) |
| 3 | 3+ (planner + N executor waves + verify) | 3-8s, justified by actual complexity | Same as now — this is correct usage |

If ~85% of real traffic (Tiers 0-2, per your catalog) currently takes 3-8 seconds and could take under 2 seconds — often under 500ms — **that gap alone is a large fraction of "why doesn't this feel like Shortcut."** Shortcut's own published behavior (Plan Mode as an *optional*, explicitly-invoked mode, not a default gate on every write) supports this: they don't force full planning on simple requests either.

---

## 5. Where verification still matters, per tier — don't lose safety chasing speed

Speed-driven redesigns tend to over-correct and drop verification where it's still needed. Be precise about this:

- **Tier 0**: no verification needed — these are structural/cosmetic, trivially reversible, no numeric risk.
- **Tier 1**: no verification needed for the same reason, **except** find & replace on numeric/financial columns — a wrong replace across 800 rows is not "trivially reversible" in the same way a wrong sort is. Route that specific case to Tier 2 instead.
- **Tier 2**: verification is mandatory, not optional — this is exactly where `VerifierAgent` earns its keep: formula correctness, sample-row testing (your own Q&A.5 table already specifies this per formula type), and the `isHardcoded` check from the earlier gap analysis should gate here specifically. This tier is where wrong-number risk actually concentrates in your catalog (GST %, SUMIFS thresholds, date-range formulas) — don't skip verification here even though it's "just one formula."
- **Tier 3**: unchanged — full pipeline, full verification, exactly as today.

---

## 6. Migration plan (safe order, doesn't require rewriting the pipeline first)

1. **Build the classifier as an additive layer in front of the existing pipeline** — don't touch `PlannerAgent`/`ExecutorAgent`/`VerifierAgent` yet. Route Tier 3 (compound-signal matches) to the existing pipeline unchanged; route Tier 0/1 to new lightweight handlers; route Tier 2 to a new "Generate→Verify" path that reuses `ExecutorAgent` (single subtask, no Planner call) and the existing `VerifierAgent` directly.
2. **Backfill the pattern table (§3) using your actual catalog** — you already have 180+ documented use cases with example phrasings; each one is a test case for the classifier. This effectively gives you the eval harness recommended in the previous conversation, for free, as a byproduct of this migration.
3. **Instrument tier distribution in production** — log which tier every write request actually lands in. Confirm the ~85%-in-Tier-0/1/2 hypothesis holds against real usage, not just the catalog's theoretical shape; adjust pattern coverage based on what's actually landing in the LLM-router fallback (meaning the regex layer missed it).
4. **Only then, revisit the Planner prompt patch** from the earlier design doc — once Tier 3 traffic is genuinely restricted to compound/compositional requests, the "reserve clarification for consequential ambiguity" rule will have much higher signal, because everything reaching the Planner will actually be complex enough to warrant that posture.
5. **Domain engine work builds on Tier 3, correctly this time** — GST reconciliation, ITC computation, etc. are inherently multi-stage, so they belong in Tier 3 by design, not because the pipeline defaults everything there. This is the corrected foundation to build the domain engine on top of.

---

## 7. Direct answer to your closing question

Build the domain engine *after* this tiering exists, not before — for two reasons. First, the domain engine's own multi-stage logic (normalize → match → categorize → compute → cite) needs a Planner that's well-behaved and reserved for genuinely complex work; building it on top of today's "everything goes through Planner, Planner over-clarifies" pipeline means the domain engine inherits that exact problem on every GST query, which is the worst place for it to show up. Second, this migration is small relative to the domain engine and directly de-risks it — you'll enter that work with a pipeline that already knows the difference between "one formula" and "a five-step reconciliation," which is precisely the distinction the domain engine will need to make correctly, at a higher stakes level, from day one.
