"""Golden tests (spec §9). These must pass before any report is trusted."""

from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from fsc import schedules

D = Decimal

TABLE_2025 = Path(__file__).resolve().parent.parent / "data" / "2025_table.json"

# Band -> (New Wheelworx, DAT Flatbed, 2025)
RATE_TABLE = [
    ("3.60", "0.23", "0.47", "0.01"),
    ("3.65", "0.24", "0.48", "0.01"),
    ("4.20", "0.35", "0.59", "0.11"),
    ("6.35", "0.78", "1.02", "0.51"),
    ("6.50", "0.81", "1.05", "0.53"),
    ("9.10", "1.33", "1.57", "1.01"),
    ("11.05", "1.72", "1.96", "1.36"),  # 2025 beyond the printed table: rule value
]

# EIA week, raw, rounded, New, DAT, 2025
WEEKLY = [
    (date(2026, 1, 5), "3.477", "3.48", "0.20", "0.44", "0.00"),
    (date(2026, 3, 2), "3.897", "3.90", "0.29", "0.53", "0.06"),
    (date(2026, 7, 13), "4.796", "4.80", "0.47", "0.71", "0.22"),
    (date(2026, 8, 3), "5.348", "5.35", "0.58", "0.82", "0.32"),
    (date(2026, 8, 31), "5.599", "5.60", "0.63", "0.87", "0.37"),
    (date(2026, 9, 14), "6.285", "6.29", "0.76", "1.00", "0.49"),
    (date(2026, 9, 21), "6.529", "6.53", "0.81", "1.05", "0.53"),
    (date(2026, 9, 28), "6.382", "6.38", "0.78", "1.02", "0.51"),
]

# DAT published 2026 monthly flatbed averages ($/mile), ±1¢
DAT_MONTHLY = [(3, "0.73"), (4, "0.85"), (8, "0.84")]


@pytest.mark.parametrize("band, new, dat, y25", RATE_TABLE)
def test_rate_table_spot_checks(params, band, new, dat, y25):
    b = D(band)
    assert params.new_wheelworx(b) == D(new)
    assert params.dat_flatbed(b) == D(dat)
    assert params.wheelworx_2025(b) == D(y25)


def test_step_and_gap_invariants(params):
    bands = schedules.bands_from(params.table_start, params.rate_table_rows, params.band_width)
    assert len(bands) == 150 and bands[-1] == params.table_end
    for prev, cur in zip(bands, bands[1:], strict=False):
        assert params.new_wheelworx(cur) - params.new_wheelworx(prev) == D("0.01")
        assert params.dat_flatbed(cur) - params.dat_flatbed(prev) == D("0.01")
    for b in bands:
        assert params.dat_flatbed(b) - params.new_wheelworx(b) == D("0.24")


def test_2025_rule_matches_every_published_row(params):
    rows = schedules.load_2025_table(TABLE_2025)
    assert len(rows) == 111
    assert rows[0][0] == D("3.60") and rows[-1][0] == D("9.10")
    for band, surcharge in rows:
        assert params.wheelworx_2025(band) == surcharge, f"band {band}"


@pytest.mark.parametrize("week, raw, rounded, new, dat, y25", WEEKLY)
def test_weekly_checks(params, week, raw, rounded, new, dat, y25):
    row = schedules.compute_week(week, D(raw), params)
    assert (row.rounded, row.new_wheelworx, row.dat_flatbed, row.wheelworx_2025) == (
        D(rounded),
        D(new),
        D(dat),
        D(y25),
    )


# ---- checks against real price data (run by `python fsc_report.py --check`) ----


def test_real_data_matches_golden_weeks(eia_prices):
    by_week = dict(eia_prices)
    for week, raw, *_ in WEEKLY:
        assert week in by_week, f"EIA data has no week {week}"
        assert by_week[week] == D(raw), f"{week}: EIA {by_week[week]} != golden {raw}"


@pytest.mark.parametrize("month, published", DAT_MONTHLY)
def test_dat_monthly_cross_check(params, eia_prices, month, published):
    n_weeks = sum(1 for w, _ in eia_prices if (w.year, w.month) == (2026, month))
    if n_weeks < 4:
        pytest.skip(
            f"only {n_weeks} week(s) of 2026-{month:02d} in the price data; need the full month"
        )
    avg = schedules.monthly_dat_average(eia_prices, 2026, month, params.dat_peg, params.dat_mpg)
    assert abs(avg - D(published)) <= D("0.01"), f"2026-{month:02d}: {avg} vs DAT {published}"
