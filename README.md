# Throttle — Inventory Scoring + SOV Suggestion Agent

Scores every supply/demand tag in an ad-inventory log on fill rate and revenue,
then suggests a QPS action for each: **Block**, **Decrease QPS**, **Monitor**, or
**No Change**. Includes a bundle-level blocklist for tags being throttled, and an
optional global revenue-loss cap that reverts the most damaging cuts back to
Monitor if they'd cost too much revenue.

Two ways to run it:

1. **CLI** — `run_agent.py`, for CSV/Excel files, produces a colour-coded `.xlsx` workbook.
2. **Web app** — `docs/index.html`, a static page that runs the exact same Python
   pipeline in-browser via [Pyodide](https://pyodide.org). No server, no data
   leaves the browser tab. Live at `https://<your-username>.github.io/<repo>/`
   once GitHub Pages is enabled (Settings → Pages → source: `docs/`).

Both entry points import the same file, **`engine/pipeline.py`** — the scoring
math is defined exactly once. `docs/assets/pipeline.py` is a byte-identical copy
kept in sync by `scripts/sync_webapp.py` (run manually, or automatically in CI
before every Pages deploy — see `.github/workflows/deploy-pages.yml`).

## CLI usage

```bash
pip install -r requirements.txt

python run_agent.py \
    --input path/to/data.csv \
    --from-date 2026-06-01 --to-date 2026-06-30 \
    --output output/inventory_scoring_output.xlsx \
    [--cost-basis requests --cost-per-1m-req 1.50] \
    [--cost-basis impressions --cost-per-1k-imp 0.20] \
    [--rev-threshold 10] \
    [--max-rev-loss-pct 2]
```

Only pass the bracketed flags for features you want **on** — each one is an
independent Yes/No gate:

| Flag | Gate |
|---|---|
| `--cost-basis` + `--cost-per-1m-req` / `--cost-per-1k-imp` | Server-cost model — blocks any tag with negative net revenue after cost |
| `--rev-threshold` | Revenue-floor override — never fully blocks a tag earning below this $ figure |
| `--max-rev-loss-pct` | Global revenue-loss cap — reverts the highest-scoring cuts back to Monitor once total projected revenue loss would exceed this % of period revenue |

## Input columns

Auto-detected by name (case/whitespace-insensitive), so exact header casing
doesn't matter:

| Column | Aliases matched | Required |
|---|---|---|
| Date | `date` | yes |
| Requests | `requests`, `request`, `ad requests` | yes |
| Revenue | `revenue` | yes |
| Fill Rate *or* Impressions | `fillrate`, `fill_rate`, `fill` / `impressions`, `imps` | one of the two |
| Supply Tag | `supply tag`, `supplytag`, `supply` | no — skips Supply-side sheet if absent |
| Demand Tag | `demand tag`, `demandtag`, `demand` | no — skips Demand-side sheet if absent |
| App Bundle | `app bundle`, `bundle`, `bundle id` | no — skips bundle blocklist if absent |

If your file uses a column name outside these aliases (e.g. a one-off export
naming requests something unexpected), rename that column before running —
the pipeline deliberately does not guess.

## Output

An `.xlsx` with, per side that ran:

- **`<Side> Tag Analysis`** — every tag, colour-coded by action, with QPS
  change %, suggested daily cap, and bundle blocklist
- **`<Side> Top & Bottom 10`** — best/worst performing tags excluding those
  already flagged for throttling
- **`Summary`** — totals, revenue/request impact, action counts, and the
  bundle-blocking rules in plain English

## Repo layout

```
engine/pipeline.py       canonical scoring + workbook-building logic
run_agent.py              CLI wrapper around engine/pipeline.py
docs/index.html           web app (Pyodide-powered, static, GitHub Pages-ready)
docs/assets/pipeline.py    byte-identical copy of engine/pipeline.py for the browser
scripts/sync_webapp.py    copies engine/pipeline.py -> docs/assets/pipeline.py
.github/workflows/        CI: syncs + deploys docs/ to GitHub Pages on push to main
```

## Development

After editing `engine/pipeline.py`, run:

```bash
python scripts/sync_webapp.py
```

before committing, so the web app doesn't drift from the CLI. (CI does this
for you automatically on every push to `main`, but keeping your local copy in
sync avoids confusing diffs.)

## Notes on the scoring model

- No live capping API is assumed — "current cap" is treated as
  `Avg_Daily_Requests`, so all suggested caps and QPS changes are relative to
  that, not a real production ceiling. Hook a real API in at the
  `Daily_Cap` computation in `engine/pipeline.py` if you have one.
- Bundle blocklists are only computed for tags already flagged
  Decrease QPS or Block, and only ever block bundles earning at or below the
  P25 daily revenue within that tag — bundles earning above P25 are never
  auto-blocked.
