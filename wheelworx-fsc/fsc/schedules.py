"""All surcharge math. Pure functions, Decimal only (never raw floats).

Rules (spec §5):
  p    = round_half_up(DOE price, 2)
  band = floor(p / band_width) * band_width
  New Wheelworx / DAT flatbed: max(0, round_half_up((band - peg) / mpg, 2))
  Wheelworx 2025: 0 below first_band, else (floor((band - first_band) / 0.055) + 1) cents
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import ROUND_FLOOR, ROUND_HALF_UP, Decimal
from pathlib import Path

CENT = Decimal("0.01")
ZERO = Decimal("0.00")


def to_decimal(x) -> Decimal:
    """Convert via str() so 3.477 stays 3.477, not 3.47699999..."""
    return x if isinstance(x, Decimal) else Decimal(str(x))


def round_half_up(x, places: int = 2) -> Decimal:
    return to_decimal(x).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def band_low(price, band_width: Decimal) -> Decimal:
    """Round the price to the cent first, then take the $0.05 band floor."""
    p = round_half_up(price, 2)
    n = (p / band_width).to_integral_value(rounding=ROUND_FLOOR)
    return (n * band_width).quantize(CENT)


def peg_surcharge_for_band(band: Decimal, peg: Decimal, mpg: Decimal) -> Decimal:
    """New Wheelworx and DAT flatbed formula: never negative."""
    return max(ZERO, round_half_up((band - peg) / mpg, 2))


def y2025_surcharge_for_band(
    band: Decimal, first_band: Decimal, dollars_per_cent: Decimal
) -> Decimal:
    """The printed 2025 table steps $0.01 per $0.055 (not the $0.05 its footer claims)."""
    if band < first_band:
        return ZERO
    cents = ((band - first_band) / dollars_per_cent).to_integral_value(rounding=ROUND_FLOOR) + 1
    return (cents * CENT).quantize(CENT)


@dataclass(frozen=True)
class Params:
    band_width: Decimal
    table_start: Decimal
    table_end: Decimal
    ww_mpg: Decimal
    ww_start_surcharge: Decimal
    dat_peg: Decimal
    dat_mpg: Decimal
    y2025_first_band: Decimal
    y2025_dollars_per_cent: Decimal
    history_weeks: int
    chart_12wk_weeks: int
    ytd_label_every_n_weeks: int

    @property
    def ww_peg(self) -> Decimal:
        """table_start - start_surcharge * mpg  (3.60 - 0.23 * 5.0 = 2.45)."""
        return (self.table_start - self.ww_start_surcharge * self.ww_mpg).quantize(CENT)

    @property
    def gap(self) -> Decimal:
        """(ww_peg - dat_peg) / mpg  ((2.45 - 1.25) / 5 = 0.24)."""
        return round_half_up((self.ww_peg - self.dat_peg) / self.ww_mpg, 2)

    @property
    def step_per_band(self) -> Decimal:
        """Surcharge change per band ($0.05 / 5.0 MPG = $0.01)."""
        return round_half_up(self.band_width / self.ww_mpg, 2)

    @property
    def rate_table_rows(self) -> int:
        """Rows from table_start to table_end inclusive (3.60..11.05 = 150)."""
        return int((self.table_end - self.table_start) / self.band_width) + 1

    def new_wheelworx(self, band: Decimal) -> Decimal:
        return peg_surcharge_for_band(band, self.ww_peg, self.ww_mpg)

    def dat_flatbed(self, band: Decimal) -> Decimal:
        return peg_surcharge_for_band(band, self.dat_peg, self.dat_mpg)

    def wheelworx_2025(self, band: Decimal) -> Decimal:
        return y2025_surcharge_for_band(band, self.y2025_first_band, self.y2025_dollars_per_cent)

    def band(self, price) -> Decimal:
        return band_low(price, self.band_width)


def load_params(config_path: str | Path) -> Params:
    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f, parse_float=Decimal, parse_int=Decimal)
    ww, dat, y25 = cfg["wheelworx_new"], cfg["dat_flatbed"], cfg["wheelworx_2025"]
    params = Params(
        band_width=cfg["band_width"],
        table_start=cfg["table_start_price"],
        table_end=cfg["table_end_price"],
        ww_mpg=ww["mpg"],
        ww_start_surcharge=ww["start_surcharge_at_table_start"],
        dat_peg=dat["peg"],
        dat_mpg=dat["mpg"],
        y2025_first_band=y25["first_band"],
        y2025_dollars_per_cent=y25["dollars_per_cent"],
        history_weeks=int(cfg["history_weeks"]),
        chart_12wk_weeks=int(cfg["chart_12wk_weeks"]),
        ytd_label_every_n_weeks=int(cfg["ytd_label_every_n_weeks"]),
    )
    if params.ww_mpg != params.dat_mpg:
        raise ValueError(
            f"config.json: wheelworx_new.mpg ({params.ww_mpg}) must equal dat_flatbed.mpg "
            f"({params.dat_mpg}); otherwise the gap vs DAT is not constant."
        )
    return params


@dataclass(frozen=True)
class WeekRow:
    week: date  # EIA week (Monday)
    raw: Decimal  # DOE price, $/gal, 3 decimals
    rounded: Decimal  # to the cent
    band: Decimal
    new_wheelworx: Decimal
    dat_flatbed: Decimal
    wheelworx_2025: Decimal

    @property
    def effective(self) -> date:
        return self.week + timedelta(days=1)

    @property
    def gap(self) -> Decimal:
        return self.dat_flatbed - self.new_wheelworx


def compute_week(week: date, raw_price, params: Params) -> WeekRow:
    raw = to_decimal(raw_price)
    band = params.band(raw)
    return WeekRow(
        week=week,
        raw=raw,
        rounded=round_half_up(raw, 2),
        band=band,
        new_wheelworx=params.new_wheelworx(band),
        dat_flatbed=params.dat_flatbed(band),
        wheelworx_2025=params.wheelworx_2025(band),
    )


def compute_weeks(prices: list[tuple[date, Decimal]], params: Params) -> list[WeekRow]:
    return [compute_week(w, p, params) for w, p in prices]


def bands_from(start: Decimal, count: int, band_width: Decimal) -> list[Decimal]:
    return [(start + band_width * i).quantize(CENT) for i in range(count)]


def monthly_dat_average(
    prices: list[tuple[date, Decimal]], year: int, month: int, peg: Decimal, mpg: Decimal
) -> Decimal | None:
    """DAT cross-check: weekly average of (round(p,2) - peg) / mpg, exact (no banding).

    Returns None when no EIA weeks fall in that month.
    """
    vals = [
        (round_half_up(p, 2) - peg) / mpg for w, p in prices if w.year == year and w.month == month
    ]
    if not vals:
        return None
    return round_half_up(sum(vals) / len(vals), 2)


def load_2025_table(path: str | Path) -> list[tuple[Decimal, Decimal]]:
    """Transcribed 2025 rows: [(band_low, surcharge), ...]."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return [(Decimal(r["band_low"]), Decimal(r["surcharge"])) for r in data["rows"]]
