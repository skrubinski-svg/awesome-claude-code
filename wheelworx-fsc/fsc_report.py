"""Wheelworx Fuel Surcharge (FSC) report.

python fsc_report.py                    # download EIA, build report
python fsc_report.py --csv prices.csv   # offline: use a manual price file (week,price)
python fsc_report.py --open             # open the report when done
python fsc_report.py --check            # run calculations + golden tests only, no file
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from fsc import eia, schedules, workbook  # noqa: E402

CONFIG = ROOT / "config.json"
TABLE_2025 = ROOT / "data" / "2025_table.json"
REPORTS = ROOT / "reports"


def load_prices(csv_path: str | None):
    if csv_path:
        return eia.load_csv(csv_path)
    return eia.parse_eia_xls(eia.download())


def load_2025():
    if not TABLE_2025.is_file():
        raise eia.FscError(
            f"{TABLE_2025} is missing. It must hold the 111 rows transcribed from the 2025 PDF."
        )
    with open(TABLE_2025, encoding="utf-8") as f:
        footer = json.load(f).get("footer_text", "")
    return schedules.load_2025_table(TABLE_2025), footer


def unique_report_path(today: date) -> Path:
    """Never overwrite: append _2, _3, ... when today's report already exists."""
    base = f"Wheelworx_FSC_Report_{today:%Y-%m-%d}"
    path = REPORTS / f"{base}.xlsx"
    n = 2
    while path.exists():
        path = REPORTS / f"{base}_{n}.xlsx"
        n += 1
    return path


def summary(latest: schedules.WeekRow) -> str:
    return (
        f"Latest EIA week {latest.week:%m/%d/%Y} (effective {latest.effective:%m/%d/%Y})  "
        f"DOE ${latest.raw:.3f} -> ${latest.rounded:.2f}\n"
        f"New Wheelworx  ${latest.new_wheelworx:.2f}/mi   "
        f"DAT flatbed ${latest.dat_flatbed:.2f}/mi   "
        f"2025 table ${latest.wheelworx_2025:.2f}/mi   Gap vs DAT ${latest.gap:.2f}"
    )


def open_file(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", str(path)], check=False)


def run_check(csv_path: str | None) -> int:
    """Golden tests via pytest. Live-data checks run against EIA (or --csv) too."""
    import pytest

    if csv_path:
        os.environ["FSC_EIA_CSV"] = str(Path(csv_path).resolve())
    else:
        os.environ["FSC_EIA_LIVE"] = "1"
    code = pytest.main(["-q", "-rs", str(ROOT / "tests")])
    print("CHECK PASSED" if code == 0 else "CHECK FAILED")
    return int(code)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Build the Wheelworx Fuel Surcharge report.")
    ap.add_argument(
        "--csv", help="manual price file with columns week,price (skips the EIA download)"
    )
    ap.add_argument("--open", action="store_true", help="open the report when done")
    ap.add_argument(
        "--check", action="store_true", help="run calculations + golden tests only, no file"
    )
    args = ap.parse_args(argv)

    if args.check:
        return run_check(args.csv)

    try:
        params = schedules.load_params(CONFIG)
        table_2025, footer_2025 = load_2025()
        prices = load_prices(args.csv)
        weeks = schedules.compute_weeks(prices, params)
        path = unique_report_path(date.today())
        used = workbook.build_workbook(path, weeks, prices, params, table_2025, footer_2025)
        problems = workbook.verify_report(path, used, params, table_2025)
    except eia.FscError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    print(summary(weeks[-1]))
    try:
        shown = path.relative_to(Path.cwd())
    except ValueError:
        shown = path
    print(f"Report: {shown}")
    if problems:
        print("OUTPUT CHECK FAILED:", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        return 3
    print("Output check: OK (values re-read from the file, 2 charts x 3 series)")
    if args.open:
        open_file(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
