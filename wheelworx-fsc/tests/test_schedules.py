"""Unit tests for the surcharge math and the EIA/CSV parsers."""

from datetime import date, datetime
from decimal import Decimal

import pandas as pd
import pytest
from fsc import eia, schedules
from fsc.schedules import band_low, round_half_up

D = Decimal


def test_derived_constants(params):
    assert params.ww_peg == D("2.45")
    assert params.gap == D("0.24")
    assert params.step_per_band == D("0.01")
    assert params.rate_table_rows == 150


def test_round_half_up_not_bankers():
    assert round_half_up("6.385") == D("6.39")
    assert round_half_up("6.375") == D("6.38")
    assert round_half_up(3.477) == D("3.48")


@pytest.mark.parametrize(
    "price, band",
    [
        (6.4999999, "6.50"),  # float artefact must not stay in the 6.45 band
        (6.499, "6.50"),  # rounds to 6.50 first, then bands
        (6.494, "6.45"),
        (4.796, "4.80"),
        (4.794, "4.75"),
        (3.60, "3.60"),
        (0.70, "0.70"),
    ],
)
def test_band_rounds_to_cent_first(price, band):
    assert band_low(price, D("0.05")) == D(band)


def test_never_negative_below_peg(params):
    assert params.new_wheelworx(D("2.00")) == D("0.00")
    assert params.dat_flatbed(D("1.00")) == D("0.00")
    assert params.wheelworx_2025(D("3.55")) == D("0.00")


def test_formula_continues_above_table_end(params):
    assert params.new_wheelworx(D("12.00")) == D("1.91")
    assert params.dat_flatbed(D("12.00")) == D("2.15")


def test_effective_date_is_week_plus_one(params):
    row = schedules.compute_week(date(2026, 9, 28), D("6.382"), params)
    assert row.effective == date(2026, 9, 29)


def test_monthly_dat_average():
    prices = [
        (date(2026, 3, 2), D("4.90")),
        (date(2026, 3, 9), D("5.00")),
        (date(2026, 4, 6), D("9.99")),
    ]
    # ((4.90-1.25)/5 + (5.00-1.25)/5) / 2 = (0.73 + 0.75) / 2 = 0.74
    assert schedules.monthly_dat_average(prices, 2026, 3, D("1.25"), D("5.0")) == D("0.74")
    assert schedules.monthly_dat_average(prices, 2026, 5, D("1.25"), D("5.0")) is None


def _frame(keys, rows):
    head = [
        ["Back to Contents"] + [None] * (len(keys) - 1),
        keys,
        ["Date"] + ["desc"] * (len(keys) - 1),
    ]
    return pd.DataFrame(head + rows)


def test_parse_matches_on_key_not_position():
    df = _frame(
        ["Sourcekey", "OTHER_KEY", eia.SERIES_KEY],
        [[datetime(2026, 9, 21), 9.999, 6.529], [datetime(2026, 9, 28), 9.999, 6.382]],
    )
    assert eia.parse_eia_frame(df) == [
        (date(2026, 9, 21), D("6.529")),
        (date(2026, 9, 28), D("6.382")),
    ]


def test_parse_skips_leading_blank_prices_and_sorts():
    df = _frame(
        ["Sourcekey", eia.SERIES_KEY],
        [
            [datetime(1990, 1, 1), None],
            [datetime(2026, 9, 28), 6.382],
            [datetime(2026, 9, 21), 6.529],
        ],
    )
    assert [w for w, _ in eia.parse_eia_frame(df)] == [date(2026, 9, 21), date(2026, 9, 28)]


def test_parse_missing_key_fails_loudly():
    df = _frame(["Sourcekey", "OTHER_KEY"], [[datetime(2026, 9, 28), 6.0]])
    with pytest.raises(eia.FscError, match=eia.SERIES_KEY):
        eia.parse_eia_frame(df)


def test_parse_non_date_row_fails_loudly():
    df = _frame(["Sourcekey", eia.SERIES_KEY], [["not a date", 6.0]])
    with pytest.raises(eia.FscError, match="not a date"):
        eia.parse_eia_frame(df)


def test_parse_empty_fails_loudly():
    df = _frame(["Sourcekey", eia.SERIES_KEY], [[datetime(2026, 9, 28), None]])
    with pytest.raises(eia.FscError, match="no price data"):
        eia.parse_eia_frame(df)


def test_load_csv(tmp_path):
    p = tmp_path / "prices.csv"
    p.write_text("week,price\n09/28/2026,6.382\n2026-09-21,$6.529\n")
    assert eia.load_csv(p) == [(date(2026, 9, 21), D("6.529")), (date(2026, 9, 28), D("6.382"))]


def test_load_csv_bad_header(tmp_path):
    p = tmp_path / "prices.csv"
    p.write_text("date,value\n2026-09-28,6.382\n")
    with pytest.raises(eia.FscError, match="week,price"):
        eia.load_csv(p)


def test_download_failure_message(monkeypatch):
    import requests

    def boom(*a, **k):
        raise requests.ConnectionError("proxy refused")

    monkeypatch.setattr(requests, "get", boom)
    with pytest.raises(
        eia.FscError,
        match=r"Could not download EIA file \(proxy refused\)\. "
        r"Check access to eia.gov or use --csv\.",
    ):
        eia.download()
