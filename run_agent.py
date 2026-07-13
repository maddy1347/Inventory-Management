"""
Run the Inventory Scoring + SOV Suggestion Agent (v4 spec) end to end.

USAGE
-----
python run_agent.py \
    --input path/to/data.csv \
    --from-date 2026-06-01 --to-date 2026-06-30 \
    --output output/inventory_scoring_output.xlsx \
    [--cost-basis requests --cost-per-1m-req 1.50] \
    [--cost-basis impressions --cost-per-1k-imp 0.20] \
    [--rev-threshold 10] \
    [--max-rev-loss-pct 2]

Only pass --cost-per-1m-req / --cost-per-1k-imp if you want the server-cost model on
(Section 1.3). Only pass --rev-threshold if you want the revenue-threshold override on
(Section 1.4). Only pass --max-rev-loss-pct if you want the v4 global revenue-loss cap on
(Section 1.5). Any omitted feature is simply left off, matching the "Yes/No" gate in the spec.

All scoring + workbook-writing logic lives in engine/pipeline.py, which is the same file
the browser app (docs/) runs under Pyodide -- this script is just the CLI wrapper.
"""

import argparse
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'engine'))
from pipeline import Config, build_workbook  # noqa: E402


def load_input(path: str) -> pd.DataFrame:
    if path.lower().endswith(('.xlsx', '.xls')):
        return pd.read_excel(path)
    return pd.read_csv(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    ap.add_argument('--from-date', default=None)
    ap.add_argument('--to-date', default=None)
    ap.add_argument('--cost-basis', choices=['requests', 'impressions'], default=None)
    ap.add_argument('--cost-per-1m-req', type=float, default=0.0)
    ap.add_argument('--cost-per-1k-imp', type=float, default=0.0)
    ap.add_argument('--rev-threshold', type=float, default=None)
    ap.add_argument('--max-rev-loss-pct', type=float, default=None)
    args = ap.parse_args()

    raw = load_input(args.input)

    cfg = Config(
        from_date=pd.to_datetime(args.from_date) if args.from_date else None,
        to_date=pd.to_datetime(args.to_date) if args.to_date else None,
        cost_enabled=args.cost_basis is not None,
        cost_basis=args.cost_basis or 'requests',
        cost_per_1m_req=args.cost_per_1m_req,
        cost_per_1k_imp=args.cost_per_1k_imp,
        rev_threshold_enabled=args.rev_threshold is not None,
        rev_threshold=args.rev_threshold or 0.0,
        max_rev_loss_enabled=args.max_rev_loss_pct is not None,
        max_rev_loss_pct=args.max_rev_loss_pct or 0.0,
    )

    try:
        wb, warnings, _ = build_workbook(raw, cfg)
    except ValueError as e:
        print(str(e), file=sys.stderr)
        sys.exit(1)

    wb.save(args.output)

    for w in warnings:
        print('WARNING:', w)
    print('Wrote', args.output)


if __name__ == '__main__':
    main()
