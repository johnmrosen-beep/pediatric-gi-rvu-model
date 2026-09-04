# Handoff: Pediatric GI Clinic RVU Production Model

Context for whoever (Claude Code included) picks this up next — this was
built iteratively in a chat session; this file is the memory that session
doesn't carry over automatically. Read this before making changes.

## What this is

A client-only web tool  to model
professional wRVU production across two clinic templates — a general
clinic and a second, independently-configurable "specialty" clinic (named
"Motility" by default, but the label, slot length, and every assumption
are editable so it can represent any second clinic type). Built to support
practice-plan / clinic-strategy conversations within the group.

**This is Model 1**: professional wRVU production only, optimized around
time-based E/M billing for all-telehealth visits. Facility RVU and total
(direct + downstream) revenue are explicitly out of scope until
revenue cycle confirms the clinic's provider-based billing status for
telehealth — see the in-app disclaimer on the Production Model tab.

## What's built (current state)

- **Three tabs**: Assumptions, Reference (CPT/wRVU table), Production Model
  (scenario table + headline chart).
- **Two clinic panels**, each independently configurable: templated slot
  length, new-vs-follow-up split (or "always follow-up" toggle), and a full
  Level 2–5 E&M code-mix table per visit type (99202–99205 for new,
  99212–99215 for follow-up), with live 100%-sum validation per table.
- **Slot-length flag**: dynamically checks the entered E&M mix against each
  clinic's slot length and flags any code whose CMS time threshold exceeds
  the template (e.g., a 30-min slot can't reach 99204's 45–59 min floor
  without work extending past the visit block). This is a real compliance/
  modeling check, not decorative.
- **Historical no-show / DNKA rates**, entered per visit type per clinic,
  blended by the new/follow-up split, applied to effective (completed)
  visits per session — scales throughput without touching the blended
  wRVU-per-visit (case-mix) figures.
- **Exports**: Excel (SheetJS, 3-sheet snapshot: Assumptions,
  CPT_Reference, Production_Model — a data snapshot, not a recalculating
  template), PDF (browser print-to-PDF via a dedicated print-only summary
  view), and JSON save/load (round-trips through this app itself).
- **State-in-URL**: assumptions are encoded into the URL hash fragment and
  kept in sync automatically (debounced) as the user edits, so refreshing,
  bookmarking, or copying the address bar preserves exact inputs. "Copy
  link" button, with a visible-input fallback if clipboard access is
  blocked. No backend — see the "Sharing / persistence" section of
  README.md for the privacy/isolation reasoning; it's worth reading before
  changing this, since the guarantees it makes (no user can see another
  user's inputs, no server ever sees the data) come specifically from using
  the hash fragment and having zero server-side state. Don't casually swap
  this for a query string or add a backend without re-deriving those
  guarantees.

## Architecture notes

- **Zero backend, fully static.** Everything computes client-side. This
  was a deliberate choice tied to the privacy/isolation requirement above
  — don't add a server without discussing the tradeoff.
- **Single source of truth for calculations**: `clinicMetrics()` and
  `computeModel()` in `src/App.jsx`. The Production Model tab, the Excel
  export (`buildProductionRows`), and the PDF summary (`PrintSummary`) all
  call `computeModel()` — don't fork this logic when adding a new export
  or view; extend `computeModel()`'s return shape instead so everything
  stays consistent by construction.
- **CPT reference data** (`NEW_CODES`, `EST_CODES`) is the single place
  2026 CMS work RVUs and time thresholds live. Update here first if the
  fee schedule changes for a new year.
- **Tailwind here is a full, normal setup** (JIT, arbitrary values, custom
  colors all fine). The original in-chat artifact preview version was
  restricted to Tailwind's core utility classes only (no arbitrary
  values) — that constraint does **not** apply to this repo. Feel free to
  use arbitrary values / a custom theme if extending the design; no need
  to preserve that restriction.
- **`xlsx` is pinned to `^0.18.5`** — the last version SheetJS published
  to the public npm registry under this package name; newer versions moved
  to SheetJS's own CDN. Check where a version actually lives before
  bumping.
- **No test suite yet.** Calculation logic has been manually verified
  against hand-computed values and cross-checked against a parallel Excel
  workbook built earlier in the same project, but there's no automated
  regression coverage.
- Recharts + xlsx together push the minified bundle to ~830KB (Vite warns
  about this at build). Not a problem yet; if it grows, look at dynamic
  `import()` for the export/chart code paths before reaching for anything
  more drastic.

## Not yet done

- **Git repo not initialized.** Do that here, with your own identity, not
  a placeholder one.
- **Not pushed to GitHub, not deployed anywhere.** See README.md for the
  push sequence and hosting options (GitHub Pages / Vercel / Netlify).
- **Facility RVU / total revenue (Model 2)** — blocked on confirming
  provider-based billing status for telehealth with revenue cycle.
  There's a placeholder input for facility RVU/encounter on the
  Assumptions tab and disclaimers throughout flagging this is unmodeled;
  wire it in once that's confirmed rather than guessing at a number.
- A parallel Excel workbook (`.xlsx`, not this repo) was built earlier as
  a standalone deliverable with live formulas — it is **not**
  in this repo and is not the same thing as the Excel *export* this app
  produces (which is a static snapshot). If asked to reconcile the two,
  check with the user about which one is meant.

## Suggested first steps

1. `npm install && npm run dev` — confirm it runs as-is before changing
   anything.
2. `git init && git add . && git commit -m "Initial pediatric GI RVU production model"`
3. Create the GitHub repo and push (README.md has the exact commands).
4. Decide on hosting and deploy (README.md has options).
5. Only then take on new feature work — the facility-RVU/Model 2 layer is
   the obvious next thing, gated on the revenue-cycle confirmation above.
