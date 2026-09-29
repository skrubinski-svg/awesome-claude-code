"""Excel writer (spec §6). Values only, never live formulas (spec §6.7)."""

from __future__ import annotations

import math
import re
import zipfile
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
    RichTextProperties,
)
from openpyxl.drawing.text import Font as DrawingFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .eia import EIA_SHEET, EIA_URL, SERIES_KEY
from .schedules import CENT, Params, WeekRow, bands_from, to_decimal

DAT_URL = (
    "https://iq.support.dat.com/1-rateview-d8f11ff2/fuel-surcharge-calculation-in-rateview-81e5c885"
)

NAVY, RED, GREY = "1F3864", "C00000", "7F7F7F"
FONT = "Arial"
MONEY, MONEY3, DATE_FMT = "$0.00", "$0.000", "mm/dd/yy"
LINE_WIDTH_EMU = 28575  # 2.25 pt

SHEET_12WK = "12-Week FSC"
SHEET_YTD = "YTD FSC"
SHEET_RATE = "Rate Table"
SHEET_DAT = "DAT Flatbed FSC"
SHEET_2025 = "2025 Wheelworx FSC"
SHEET_DOE = "DOE Weekly Prices"

HEADER_ROW = 4
FIRST_DATA_ROW = HEADER_ROW + 1
BLOCK_COLS = (1, 5, 9)  # From columns of the 3 side-by-side blocks (A, E, I)
PANEL_LABEL_COL, PANEL_VALUE_COL = 17, 18  # Q, R

_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HIGHLIGHT_FILL = PatternFill("solid", fgColor="FFFF00")


# ---------- small helpers ----------


def _money(d: Decimal) -> str:
    return f"${d:.2f}"


def _num(d: Decimal) -> float:
    return float(d)


def _put(
    ws,
    row,
    col,
    value,
    fmt=None,
    bold=False,
    size=10,
    color=None,
    border=False,
    wrap=False,
    align=None,
):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=FONT, bold=bold, size=size, color=color)
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BORDER
    if wrap or align:
        c.alignment = Alignment(wrap_text=wrap, horizontal=align, vertical="center")
    return c


def _header(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.border = BORDER
    c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    return c


def _title(ws, text, basis=None):
    _put(ws, 1, 1, text, bold=True, size=14, color=NAVY)
    if basis:
        _put(ws, 2, 1, basis, size=9, color="595959")


def _fit_one_page(ws, landscape):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def chart_label(d: date) -> str:
    return f"{d:%b} {d.day}"  # "Jul 13"


# ---------- chart ----------


def _char_props(size_pt, bold=False):
    return CharacterProperties(sz=int(size_pt * 100), b=bold, latin=DrawingFont(typeface=FONT))


def _axis_text(size_pt):
    cp = _char_props(size_pt)
    return RichText(
        bodyPr=RichTextProperties(rot=0, vert="horz"),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)],
    )


def _axis_title(text, size_pt):
    cp = _char_props(size_pt, bold=True)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def _line_chart(ws, first_row, last_row, series_cols, label_col, width_cm, height_cm):
    ch = LineChart()
    ch.width, ch.height = width_cm, height_cm
    for col in series_cols:
        ch.add_data(
            Reference(ws, min_col=col, min_row=HEADER_ROW, max_row=last_row), titles_from_data=True
        )
    ch.set_categories(Reference(ws, min_col=label_col, min_row=first_row, max_row=last_row))
    for s, color in zip(ch.series, (NAVY, RED, GREY), strict=True):
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = LINE_WIDTH_EMU
        s.smooth = False
        s.marker.symbol = "none"
    ch.legend.position = "b"
    ch.y_axis.number_format = MONEY
    ch.y_axis.title = _axis_title("$ per mile", 12)
    ch.x_axis.title = _axis_title("EIA week", 12)
    ch.y_axis.txPr = _axis_text(11)
    ch.x_axis.txPr = _axis_text(11)
    ch.x_axis.tickLblSkip = 1  # labels are pre-thinned in the label column itself
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    return ch


# ---------- weekly sheets ----------


def weekly_headers(params: Params, with_gap: bool) -> list[str]:
    cols = [
        "EIA Week",
        "Effective",
        "DOE Nat'l Avg (to the cent)",
        f"New Wheelworx (DAT flatbed − {_money(params.gap)})",
        f"DAT Flatbed ({_money(params.dat_peg)} peg, {params.dat_mpg} MPG)",
        "Wheelworx 2025 Table",
    ]
    if with_gap:
        cols.append("Gap vs DAT")
    cols.append("Chart label")
    return cols


def _weekly_sheet(
    ws, title, rows: list[WeekRow], labels: list[str], params: Params, with_gap: bool
):
    _title(
        ws,
        title,
        f"Source: EIA U.S. national average on-highway diesel ({EIA_URL}) | "
        f"Effective = EIA week + 1 day | Surcharges in $ per mile",
    )
    headers = weekly_headers(params, with_gap)
    for j, h in enumerate(headers, start=1):
        _header(ws, HEADER_ROW, j, h)
    label_col = len(headers)
    for i, (w, lab) in enumerate(zip(rows, labels, strict=True)):
        r = FIRST_DATA_ROW + i
        _put(ws, r, 1, w.week, DATE_FMT, border=True)
        _put(ws, r, 2, w.effective, DATE_FMT, border=True)
        _put(ws, r, 3, _num(w.rounded), MONEY, border=True)
        _put(ws, r, 4, _num(w.new_wheelworx), MONEY, border=True, bold=True)
        _put(ws, r, 5, _num(w.dat_flatbed), MONEY, border=True)
        _put(ws, r, 6, _num(w.wheelworx_2025), MONEY, border=True)
        if with_gap:
            _put(ws, r, 7, _num(w.gap), MONEY, border=True)
        _put(ws, r, label_col, lab, color="A6A6A6", border=True)
    ws.row_dimensions[HEADER_ROW].height = 45
    for j, width in enumerate(
        [10, 10, 13, 17, 17, 13] + ([11] if with_gap else []) + [10], start=1
    ):
        ws.column_dimensions[get_column_letter(j)].width = width
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    return FIRST_DATA_ROW, FIRST_DATA_ROW + len(rows) - 1, label_col


def ytd_labels(rows: list[WeekRow], every_n: int) -> list[str]:
    return [chart_label(w.week) if i % every_n == 0 else "" for i, w in enumerate(rows)]


# ---------- block tables ----------


def _block_table(ws, bands, surcharges, band_width, highlight_band, headers):
    per_block = math.ceil(len(bands) / len(BLOCK_COLS))
    to_offset = band_width - CENT  # To = From + $0.04: band edges never overlap
    for c0 in BLOCK_COLS:
        for k, h in enumerate(headers):
            _header(ws, HEADER_ROW, c0 + k, h)
        ws.column_dimensions[get_column_letter(c0)].width = 9
        ws.column_dimensions[get_column_letter(c0 + 1)].width = 9
        ws.column_dimensions[get_column_letter(c0 + 2)].width = 12
        if c0 + 3 < BLOCK_COLS[-1]:
            ws.column_dimensions[get_column_letter(c0 + 3)].width = 2
    for i, (band, s) in enumerate(zip(bands, surcharges, strict=True)):
        c0 = BLOCK_COLS[i // per_block]
        r = FIRST_DATA_ROW + i % per_block
        hit = highlight_band is not None and band == highlight_band
        for k, v in enumerate((band, band + to_offset, s)):
            c = _put(ws, r, c0 + k, _num(v), MONEY, bold=hit or k == 2, border=True, align="center")
            if hit:
                c.fill = HIGHLIGHT_FILL
    ws.row_dimensions[HEADER_ROW].height = 30
    return FIRST_DATA_ROW + per_block  # first row after the blocks


def _panel(ws, items):
    ws.column_dimensions[get_column_letter(PANEL_LABEL_COL)].width = 26
    ws.column_dimensions[get_column_letter(PANEL_VALUE_COL)].width = 12
    _header(ws, HEADER_ROW, PANEL_LABEL_COL, "Inputs & current")
    _header(ws, HEADER_ROW, PANEL_VALUE_COL, "Value")
    for i, (label, value, fmt) in enumerate(items):
        r = FIRST_DATA_ROW + i
        _put(ws, r, PANEL_LABEL_COL, label, border=True)
        _put(ws, r, PANEL_VALUE_COL, value, fmt, border=True, bold=True, align="right")


def _footer(ws, row, lines):
    for i, text in enumerate(lines):
        _put(ws, row + i, 1, text, size=9, color="404040" if i else "000000", bold=(i == 0))


# ---------- workbook ----------


def build_workbook(
    path: Path,
    weeks: list[WeekRow],
    prices: list,
    params: Params,
    table_2025: list[tuple[Decimal, Decimal]],
    footer_2025: str,
) -> dict:
    """Write the report. Returns the week lists used, for verification."""
    latest = weeks[-1]
    last12 = weeks[-params.chart_12wk_weeks :]
    ytd = [w for w in weeks if w.week.year == latest.week.year]
    wb = Workbook()

    # 6.1 12-Week FSC
    ws = wb.active
    ws.title = SHEET_12WK
    first, last, label_col = _weekly_sheet(
        ws,
        "Fuel Surcharge per Mile — Last 12 Weeks",
        last12,
        [chart_label(w.week) for w in last12],
        params,
        with_gap=False,
    )
    ch = _line_chart(ws, first, last, (4, 5, 6), label_col, 22, 11)
    ws.add_chart(ch, "I4")
    _fit_one_page(ws, landscape=True)

    # 6.2 YTD FSC
    ws = wb.create_sheet(SHEET_YTD)
    first, last, label_col = _weekly_sheet(
        ws,
        f"Fuel Surcharge per Mile — {latest.week.year} Year to Date, Week by Week",
        ytd,
        ytd_labels(ytd, params.ytd_label_every_n_weeks),
        params,
        with_gap=True,
    )
    ch = _line_chart(ws, first, last, (4, 5, 6), label_col, 34, 14)
    ws.add_chart(ch, "J4")
    _fit_one_page(ws, landscape=True)

    # 6.3 Rate Table
    ws = wb.create_sheet(SHEET_RATE)
    bw = params.band_width
    rate_bands = bands_from(params.table_start, params.rate_table_rows, bw)
    in_table = params.table_start <= latest.band <= params.table_end
    _title(
        ws,
        f"WHEELWORX — Fuel Surcharge Schedule (Flatbed {params.ww_mpg} MPG, "
        f"{_money(params.step_per_band)} per {_money(bw)})",
        f"Basis: EIA U.S. national average diesel | (price − {_money(params.ww_peg)}) / "
        f"{params.ww_mpg} MPG: {_money(params.new_wheelworx(params.table_start))} at "
        f"{_money(params.table_start)}, +{_money(params.step_per_band)} per {_money(bw)} | "
        f"Highlighted = current {_money(latest.rounded)}",
    )
    end_row = _block_table(
        ws,
        rate_bands,
        [params.new_wheelworx(b) for b in rate_bands],
        bw,
        latest.band if in_table else None,
        ("From", "To", "Surcharge per Mile"),
    )
    _panel(
        ws,
        [
            (f"Current DOE price (week {latest.week:%m/%d/%y})", _num(latest.rounded), MONEY),
            ("Table start", _num(params.table_start), MONEY),
            ("Wheelworx peg", _num(params.ww_peg), MONEY),
            ("MPG", _num(params.ww_mpg), "0.0"),
            ("Current surcharge per mile", _num(latest.new_wheelworx), MONEY),
            ("Gap vs DAT flatbed", _num(params.gap), MONEY),
            ("Surcharge at table start", _num(params.new_wheelworx(params.table_start)), MONEY),
        ],
    )
    if not in_table:
        _put(
            ws,
            FIRST_DATA_ROW + 8,
            PANEL_LABEL_COL,
            "Current price is outside the printed table; formula still applies.",
            size=9,
            color=RED,
        )
    _footer(
        ws,
        end_row + 1,
        [
            f"Equals the DAT flatbed table ({_money(params.dat_peg)} peg, {params.dat_mpg} MPG) "
            f"less {_money(params.gap)} per mile at every price "
            f"(Wheelworx peg {_money(params.ww_peg)})",
            f"DAT source: {DAT_URL}",
            f"EIA source: {EIA_URL}",
        ],
    )
    _fit_one_page(ws, landscape=False)

    # 6.4 DAT Flatbed FSC
    ws = wb.create_sheet(SHEET_DAT)
    dat_bands = bands_from(params.dat_peg, params.rate_table_rows, bw)
    _title(
        ws,
        f"DAT Flatbed — Fuel Surcharge Schedule ({_money(params.dat_peg)} peg, "
        f"{params.dat_mpg} MPG, "
        f"{_money(params.step_per_band)} per {_money(bw)})",
        f"(band − {_money(params.dat_peg)}) / {params.dat_mpg} MPG | "
        f"Highlighted = current {_money(latest.rounded)}",
    )
    in_dat = dat_bands[0] <= latest.band <= dat_bands[-1]
    end_row = _block_table(
        ws,
        dat_bands,
        [params.dat_flatbed(b) for b in dat_bands],
        bw,
        latest.band if in_dat else None,
        ("From", "To", "Surcharge per Mile"),
    )
    _panel(
        ws,
        [
            ("Peg (input)", _num(params.dat_peg), MONEY),
            ("MPG (input)", _num(params.dat_mpg), "0.0"),
            (f"Current DOE price (week {latest.week:%m/%d/%y})", _num(latest.rounded), MONEY),
            ("Current DAT surcharge per mile", _num(latest.dat_flatbed), MONEY),
        ],
    )
    _footer(
        ws,
        end_row + 1,
        [
            f"DAT RateView method: DOE price rounded to the cent, {_money(params.dat_peg)} "
            f"baseline, flatbed {params.dat_mpg} MPG.",
            f"DAT source: {DAT_URL}",
            f"EIA source: {EIA_URL}",
        ],
    )
    _fit_one_page(ws, landscape=False)

    # 6.5 2025 Wheelworx FSC
    ws = wb.create_sheet(SHEET_2025)
    _title(
        ws,
        "WHEELWORX 2025 — Fuel Surcharge Table (as published, reference only)",
        f"{len(table_2025)} rows transcribed from the published 2025 table",
    )
    end_row = _block_table(
        ws,
        [b for b, _ in table_2025],
        [s for _, s in table_2025],
        bw,
        None,
        ("From", "To", "Surcharge per Mile"),
    )
    _footer(
        ws,
        end_row + 1,
        [
            f'Published footer: "{footer_2025}"',
            f"Note: the printed rows actually step $0.01 per ${params.y2025_dollars_per_cent}, "
            f"not per {_money(bw)}. Rule that reproduces every row: "
            f"$0.00 below {_money(params.y2025_first_band)}; otherwise "
            f"cents = floor((band − {_money(params.y2025_first_band)}) ÷ "
            f"{params.y2025_dollars_per_cent}) + 1.",
        ],
    )
    _fit_one_page(ws, landscape=False)

    # 6.6 DOE Weekly Prices
    ws = wb.create_sheet(SHEET_DOE)
    _title(ws, "DOE Weekly Prices — U.S. On-Highway Diesel ($/gal)")
    _put(ws, 2, 1, f"Source: {EIA_URL} (sheet '{EIA_SHEET}')", size=9, color="595959")
    _put(ws, 3, 1, f"Series key: {SERIES_KEY}", size=9, color="595959")
    _header(ws, HEADER_ROW, 1, "Week (Monday)")
    _header(ws, HEADER_ROW, 2, "U.S. Average")
    history = prices[-params.history_weeks :]
    for i, (wk, p) in enumerate(history):
        _put(ws, FIRST_DATA_ROW + i, 1, wk, DATE_FMT, border=True)
        _put(ws, FIRST_DATA_ROW + i, 2, _num(p), MONEY3, border=True)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 14
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return {"last12": last12, "ytd": ytd, "history": history}


# ---------- output checks (spec §9) ----------


def _cell_dec(v) -> Decimal:
    """Exact: a stored 0.234 must not pass as 0.23."""
    return to_decimal(v)


def _read_block(ws, n_rows):
    per_block = math.ceil(n_rows / len(BLOCK_COLS))
    out = []
    for i in range(n_rows):
        c0 = BLOCK_COLS[i // per_block]
        r = FIRST_DATA_ROW + i % per_block
        out.append((_cell_dec(ws.cell(r, c0).value), _cell_dec(ws.cell(r, c0 + 2).value)))
    return out


def verify_report(
    path: Path, used: dict, params: Params, table_2025: list[tuple[Decimal, Decimal]]
) -> list[str]:
    """Re-open the saved file and check it. Returns a list of problems (empty = OK)."""
    problems: list[str] = []
    wb = load_workbook(path)

    for name, rows, gap in ((SHEET_12WK, used["last12"], False), (SHEET_YTD, used["ytd"], True)):
        ws = wb[name]
        n = 0
        while ws.cell(FIRST_DATA_ROW + n, 1).value is not None:
            n += 1
        if n != len(rows):
            problems.append(f"{name}: {n} data rows, expected {len(rows)}")
        for i, w in enumerate(rows):
            r = FIRST_DATA_ROW + i
            got = [ws.cell(r, c).value for c in (3, 4, 5, 6) + ((7,) if gap else ())]
            exp = [w.rounded, w.new_wheelworx, w.dat_flatbed, w.wheelworx_2025] + (
                [w.gap] if gap else []
            )
            if any(g is None or _cell_dec(g) != e for g, e in zip(got, exp, strict=True)):
                problems.append(
                    f"{name} row {r}: cells {got} != calculated {[str(e) for e in exp]}"
                )
    if len(used["last12"]) != params.chart_12wk_weeks:
        problems.append(
            f"{SHEET_12WK}: only {len(used['last12'])} weeks of data "
            f"(expected {params.chart_12wk_weeks})"
        )
    ytd_year = used["ytd"][-1].week.year
    if any(w.week.year != ytd_year for w in used["ytd"]):
        problems.append(f"{SHEET_YTD}: rows from outside {ytd_year}")

    n = params.rate_table_rows
    for name, fn in ((SHEET_RATE, params.new_wheelworx), (SHEET_DAT, params.dat_flatbed)):
        for band, s in _read_block(wb[name], n):
            if s != fn(band):
                problems.append(f"{name}: band {band} shows {s}, calculated {fn(band)}")
    for (band, s), (eb, es) in zip(
        _read_block(wb[SHEET_2025], len(table_2025)), table_2025, strict=True
    ):
        if (band, s) != (eb, es):
            problems.append(f"{SHEET_2025}: row {band}/{s} != published {eb}/{es}")

    problems += _verify_charts(path)
    return problems


def _verify_charts(path: Path) -> list[str]:
    """openpyxl does not load charts, so read the chart XML directly."""
    problems = []
    with zipfile.ZipFile(path) as z:
        charts = [n for n in z.namelist() if re.fullmatch(r"xl/charts/chart\d+\.xml", n)]
        found = {}
        for n in charts:
            xml = z.read(n).decode("utf-8")
            n_series = len(re.findall(r"<(?:c:)?ser>", xml))
            refs = set(re.findall(r"<(?:c:)?f>'?([^'!<]+)'?!", xml))
            for sheet in refs:
                found[sheet] = n_series
    for sheet in (SHEET_12WK, SHEET_YTD):
        if sheet not in found:
            problems.append(f"{sheet}: line chart missing")
        elif found[sheet] != 3:
            problems.append(f"{sheet}: chart has {found[sheet]} series, expected 3")
    return problems
