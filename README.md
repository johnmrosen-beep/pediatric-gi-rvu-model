# Pediatric GI Clinic — RVU Production Model

Interactive web tool for modeling professional wRVU production across clinic
templates (e.g., General GI vs. a specialty clinic such as Motility), built
for practice-plan planning.

Three views:
- **Assumptions** — clinic schedule, per-clinic slot length, new/follow-up
  split, historical no-show/DNKA rates, and a full Level 2–5 E&M code-mix
  table per visit type (validates that each mix sums to 100%, and flags
  codes whose time threshold exceeds the templated slot length).
- **Reference** — the 2026 CMS work RVU / time-threshold table the model is
  built on.
- **Production Model** — no-show-adjusted wRVU/hour by clinic, a scenario
  table across every schedule mix from 0–100% General GI, and the
  current-baseline vs. RVU-optimum comparison.

## Sharing / persistence

Assumptions are encoded into the page's URL hash (`#...`) and kept in sync
automatically — refreshing, bookmarking, or copying the address bar
preserves whatever's currently entered. The **Copy link** button does that
in one click (with a visible-input fallback if the browser blocks
clipboard access).

This is deliberately backend-free:
- The hash fragment is **never sent in the HTTP request** for the page, so
  it never reaches a server log, CDN, or analytics — it only ever exists
  in the requesting browser.
- Because there's no shared server-side store, **concurrent users are
  isolated by construction** — there's no mechanism by which one user's
  session could read another's. Nothing extra was built for this; it falls
  out of the static-site architecture.
- It is **not encryption** — anyone holding a link can decode it (it's
  just base64url JSON). Treat a link the way you'd treat any other
  "shareable by URL" tool. Fine for clinic-operations assumptions; not
  a fit if this ever needs to carry anything sensitive.

Excel and PDF exports (below) remain the right choice for a static,
point-in-time deliverable rather than a live/editable link.

## Exporting

Nothing is persisted automatically — state lives in the browser tab only.
From the header:
- **Export Excel** — downloads a 3-sheet `.xlsx` snapshot (Assumptions,
  CPT_Reference, Production_Model) of the current inputs and computed
  results, built client-side with SheetJS. This is a data snapshot, not a
  recalculating template — open it in Excel to review or hand off, but
  editing cells there won't recompute anything.
- **Export PDF summary** — opens the browser print dialog with a
  print-only one-page summary (headline stat, clinic assumptions, full
  scenario table); choose "Save as PDF" as the destination.
- **Save / Load assumptions (JSON)** — round-trips the raw input state
  through this tool itself, for continuing work later or sharing a working
  file with someone else using the app.

## Stack

Vite + React 18 + Tailwind CSS + Recharts. No backend; everything computes
client-side.

## Local development

```bash
npm install
npm run dev       # starts a local dev server, prints a URL
npm run build      # production build to dist/
npm run preview    # serve the production build locally
```

## Working on this with Claude Code

This repo is set up so you can open it directly in Claude Code and ask it to
extend the model — e.g., add a facility-RVU / total-revenue view (Model 2),
add a third clinic type, or wire up persistence. The calculation logic lives
in `src/App.jsx`:

- `NEW_CODES` / `EST_CODES` — the CPT reference table (code, work RVU, time
  thresholds). Update here first if CMS rates change for a new fee-schedule
  year.
- `clinicMetrics()` — the core per-clinic calculation (blended wRVU →
  slots/session → wRVU/session → wRVU/hour).
- `ProductionTab` — the scenario table and chart.
- `defaultState()` — starting assumptions; safe to change the placeholder
  numbers without touching the calculation logic.

## Publishing to GitHub

From this folder:

```bash
git init
git add .
git commit -m "Initial pediatric GI RVU production model"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```

## Deploying

Any static host works since this is a client-only build (`npm run build`
outputs static files to `dist/`):
- **GitHub Pages** — add a `gh-pages` deploy step or use
  `vite-plugin-gh-pages`; set `base` in `vite.config.js` to your repo name.
- **Vercel / Netlify** — connect the repo, build command `npm run build`,
  output directory `dist`.

## Planned next steps

- Facility RVU + total (professional + facility) revenue per visit type,
  pending confirmation of provider-based billing status for
  telehealth from revenue cycle.
- A third model layer for access/volume optimization, to weigh against pure
  RVU-production optimization.
