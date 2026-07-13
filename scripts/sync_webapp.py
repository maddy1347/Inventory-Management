#!/usr/bin/env python3
"""
Copies engine/pipeline.py -> docs/assets/pipeline.py verbatim.

engine/pipeline.py is the single source of truth for the scoring logic.
The web app (docs/index.html) loads docs/assets/pipeline.py into Pyodide
at runtime, so the two files must stay identical. Run this after any edit
to engine/pipeline.py. CI (.github/workflows/deploy-pages.yml) also runs
it automatically before every Pages deploy, so a forgotten manual sync
never ships stale logic to the live site.
"""
import filecmp
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'engine' / 'pipeline.py'
DST = ROOT / 'docs' / 'assets' / 'pipeline.py'


def main():
    if not SRC.exists():
        print(f'FATAL: {SRC} not found', file=sys.stderr)
        sys.exit(1)
    DST.parent.mkdir(parents=True, exist_ok=True)
    changed = not DST.exists() or not filecmp.cmp(SRC, DST, shallow=False)
    shutil.copyfile(SRC, DST)
    print(f'{"Updated" if changed else "Already in sync:"} {DST.relative_to(ROOT)}')


if __name__ == '__main__':
    main()
