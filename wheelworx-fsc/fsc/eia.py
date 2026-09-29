"""Download and parse the EIA weekly retail on-highway diesel file (spec §3)."""

from __future__ import annotations

import csv
import io
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
import requests

from .schedules import round_half_up

EIA_URL = "https://www.eia.gov/petroleum/gasdiesel/xls/psw18vwall.xls"
EIA_SHEET = "Data 1"
SERIES_KEY = "EMD_EPD2D_PTE_NUS_DPG"  # U.S. No 2 diesel retail, $/gal
KEY_ROW = 1  # 0-based: spreadsheet row 2 holds the EIA source keys
FIRST_DATA_ROW = 3  # 0-based: spreadsheet row 4 is the first data row
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)


class FscError(Exception):
    """A problem the user must fix; the message is printed as-is."""


def download(url: str = EIA_URL, timeout: int = 60) -> bytes:
    try:
        resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise FscError(
            f"Could not download EIA file ({e}). Check access to eia.gov or use --csv."
        ) from e
    if not resp.content:
        raise FscError(
            "Could not download EIA file (empty response). Check access to eia.gov or use --csv."
        )
    return resp.content


def parse_eia_xls(content: bytes, series_key: str = SERIES_KEY) -> list[tuple[date, Decimal]]:
    try:
        df = pd.read_excel(io.BytesIO(content), sheet_name=EIA_SHEET, header=None, engine="xlrd")
    except ValueError as e:
        raise FscError(f"EIA file has no sheet '{EIA_SHEET}' or is not a valid .xls ({e}).") from e
    except Exception as e:  # xlrd raises its own types for HTML/garbage content
        raise FscError(
            f"EIA file could not be read as .xls ({e}). The download may be an error page."
        ) from e
    return parse_eia_frame(df, series_key)


def parse_eia_frame(df: pd.DataFrame, series_key: str = SERIES_KEY) -> list[tuple[date, Decimal]]:
    """Match the series on the row-2 key, never on column position."""
    if len(df) <= FIRST_DATA_ROW:
        raise FscError("EIA file has no data rows.")
    keys = [str(v).strip() for v in df.iloc[KEY_ROW].tolist()]
    if series_key not in keys:
        raise FscError(
            f"EIA series key '{series_key}' is missing from row 2 of sheet '{EIA_SHEET}'."
        )
    col = keys.index(series_key)

    out: list[tuple[date, Decimal]] = []
    for i in range(FIRST_DATA_ROW, len(df)):
        wk, price = df.iat[i, 0], df.iat[i, col]
        if pd.isna(wk) and pd.isna(price):
            continue  # blank trailing row
        if not isinstance(wk, (datetime, pd.Timestamp)):
            raise FscError(f"EIA sheet row {i + 1}: column A is not a date ({wk!r}).")
        if pd.isna(price):
            continue  # series has no value that week (e.g. before it began)
        out.append((pd.Timestamp(wk).date(), round_half_up(Decimal(repr(float(price))), 3)))
    return _validated(out, "EIA file")


def load_csv(path: str | Path) -> list[tuple[date, Decimal]]:
    """Manual fallback: columns week,price. Dates as YYYY-MM-DD or MM/DD/YYYY."""
    p = Path(path)
    if not p.is_file():
        raise FscError(f"CSV file not found: {p}")
    out: list[tuple[date, Decimal]] = []
    with open(p, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or {"week", "price"} - {
            h.strip().lower() for h in reader.fieldnames
        }:
            raise FscError(f"CSV {p} must have columns 'week,price' (found {reader.fieldnames}).")
        for n, row in enumerate(reader, start=2):
            row = {k.strip().lower(): (v or "").strip() for k, v in row.items()}
            if not row["week"] and not row["price"]:
                continue
            wk = _parse_date(row["week"])
            if wk is None:
                raise FscError(f"CSV {p} line {n}: week '{row['week']}' is not a date.")
            try:
                price = Decimal(row["price"].lstrip("$"))
            except InvalidOperation:
                raise FscError(
                    f"CSV {p} line {n}: price '{row['price']}' is not a number."
                ) from None
            out.append((wk, round_half_up(price, 3)))
    return _validated(out, f"CSV {p}")


def _parse_date(s: str) -> date | None:
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _validated(rows: list[tuple[date, Decimal]], source: str) -> list[tuple[date, Decimal]]:
    if not rows:
        raise FscError(f"{source}: no price data found.")
    rows = sorted(rows)
    weeks = [w for w, _ in rows]
    if len(set(weeks)) != len(weeks):
        raise FscError(f"{source}: duplicate week dates.")
    bad = [(w, p) for w, p in rows if p <= 0]
    if bad:
        raise FscError(f"{source}: non-positive price on {bad[0][0]:%m/%d/%Y}.")
    return rows
