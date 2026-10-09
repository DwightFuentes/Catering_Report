#!/usr/bin/env python3
"""Build the LHSD Activities with Catering report from the Catering Services tab.

The layout follows the Local Health Support Division Power BI report
(overview, to-be-conducted, conducted). Figures are recomputed from the
current sheet, not copied from the 17 July 2026 snapshot.
"""

from __future__ import annotations

import calendar
import html
import math
import re
import subprocess
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "data" / "catering-services.xlsx"
OUT_DIR = ROOT / "output"
HTML_PATH = OUT_DIR / "report.html"
PDF_PATH = OUT_DIR / "LHSD-Activities-with-Catering.pdf"

AS_OF = date.today()
SHEET = "Catering Services"

MONTHS = {
    "january": 1,
    "jan": 1,
    "february": 2,
    "feb": 2,
    "march": 3,
    "mar": 3,
    "april": 4,
    "apr": 4,
    "may": 5,
    "june": 6,
    "jun": 6,
    "july": 7,
    "jul": 7,
    "august": 8,
    "aug": 8,
    "september": 9,
    "sept": 9,
    "sep": 9,
    "october": 10,
    "oct": 10,
    "november": 11,
    "nov": 11,
    "december": 12,
    "dec": 12,
}
MONTH_NAMES = [
    "",
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]
FUND_ORDER = ["GAA", "SAA", "ConAp"]

INK = "#1d5245"
GREEN = "#1f8a70"
ORANGE = "#ef8c2f"
RED = "#e04b45"


@dataclass
class Activity:
    program: str
    fund: str
    title: str
    remarks: str
    po_status: str
    pr_submitted: str
    when: date | None
    when_label: str
    conducted: bool
    billing: str
    excluded: bool


def load_rows(path: Path) -> list[tuple]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[SHEET]
    rows = []
    for index, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        title = row[4]
        if title is None or not str(title).strip():
            continue
        rows.append(row)
    wb.close()
    return rows


def clean(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def fund_name(value) -> str:
    key = clean(value).upper().replace(" ", "")
    return {"GAA": "GAA", "SAA": "SAA", "CONAP": "ConAp"}.get(key, clean(value) or "—")


def parse_when(value) -> tuple[date | None, str]:
    """Return a start date and a display label.

    Text dates whose year ran past 2026 are brought back to 2026. Those
    values are an autofill from August 2026 (August 12-13, 2027, 2028, …).
    Real Excel dates are kept as stored.
    """
    if value is None or (isinstance(value, str) and not value.strip()):
        return None, "—"
    if isinstance(value, datetime):
        day = value.date()
        return day, f"{MONTH_NAMES[day.month]} {day.day}, {day.year}"
    if isinstance(value, date):
        return value, f"{MONTH_NAMES[value.month]} {value.day}, {value.year}"

    text = clean(value)
    slash = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(20\d\d)", text)
    if slash:
        month, day, year = int(slash.group(1)), int(slash.group(2)), int(slash.group(3))
        if 1 <= month <= 12:
            day = min(day, calendar.monthrange(year, month)[1])
            parsed = date(year, month, day)
            return parsed, f"{MONTH_NAMES[month]} {day}, {year}"

    low = text.lower().replace(".", " ")
    month = None
    for name, number in sorted(MONTHS.items(), key=lambda item: -len(item[0])):
        if re.search(rf"\b{name}\b", low):
            month = number
            break
    if month is None:
        return None, text

    years = [int(year) for year in re.findall(r"20\d\d", text)]
    year = years[0] if years else 2026
    if year > 2026:
        year = 2026
    without_year = re.sub(r"20\d\d", " ", text)
    days = [int(day) for day in re.findall(r"\b(\d{1,2})\b", without_year)]
    days = [day for day in days if 1 <= day <= 31]
    start = days[0] if days else 1
    start = min(start, calendar.monthrange(year, month)[1])
    parsed = date(year, month, start)
    if len(days) >= 2 and days[1] != days[0] and days[1] <= 31:
        end = min(days[1], calendar.monthrange(year, month)[1])
        label = f"{MONTH_NAMES[month]} {days[0]}–{end}, {year}"
    elif days:
        label = f"{MONTH_NAMES[month]} {start}, {year}"
    else:
        label = f"{MONTH_NAMES[month]} {year}"
    return parsed, label


def moved_to(remarks: str) -> tuple[date | None, str]:
    match = re.search(
        r"moved to\s+([A-Za-z]+)\.?(?:\s+(\d{1,2}))?(?:,?\s*(20\d\d))?",
        remarks,
        re.I,
    )
    if not match:
        return None, ""
    month = MONTHS.get(match.group(1).lower())
    if not month:
        return None, ""
    year = int(match.group(3) or 2026)
    if match.group(2):
        day = min(int(match.group(2)), calendar.monthrange(year, month)[1])
        return date(year, month, day), f"{MONTH_NAMES[month]} {day}, {year}"
    return date(year, month, 1), f"{MONTH_NAMES[month]} {year}"


def has_purchase_order(value) -> bool:
    text = clean(value)
    if not text:
        return False
    if re.search(r"with\s*p\.?\s*o", text, re.I):
        return True
    digits = re.sub(r"\D", "", text)
    return len(digits) >= 5 and len(text) >= 8


def billing_submitted(remarks: str, finance, paid) -> bool:
    if finance not in (None, ""):
        return True
    if isinstance(paid, (int, float)) and paid != 0:
        return True
    if isinstance(paid, str) and clean(paid) not in {"", "0", "-", "—"}:
        return True
    if re.search(r"\bpaid\b|disburs", remarks, re.I):
        return True
    if re.search(
        r"billing.{0,48}(process|submit|receiv)|bill received|documents received",
        remarks,
        re.I,
    ):
        return True
    return False


def is_set_aside(remarks: str) -> bool:
    return bool(re.search(r"cancel|dovetailed to|recalibrat", remarks, re.I))


def build_activities(rows: list[tuple]) -> list[Activity]:
    activities = []
    for row in rows:
        remarks = clean(row[27])
        program = clean(row[1]) or "—"
        title = clean(row[4])
        fund = fund_name(row[3])
        conducted_raw, conducted_label = parse_when(row[21])
        target_raw, target_label = parse_when(row[16])
        flag = clean(row[30]).upper()
        excluded = is_set_aside(remarks)
        not_conducted = bool(re.search(r"not conducted", remarks, re.I))

        happened = False
        if not excluded and not not_conducted:
            if conducted_raw is not None and conducted_raw <= AS_OF:
                happened = True
            elif flag == "CONDUCTED" and conducted_raw is None:
                happened = True

        if happened:
            when, label = conducted_raw, conducted_label
            if when is None:
                when, label = target_raw, target_label
        else:
            moved_day, moved_label = moved_to(remarks)
            if moved_day is not None and conducted_raw is None:
                when, label = moved_day, moved_label
            elif conducted_raw is not None and conducted_raw > AS_OF:
                when, label = conducted_raw, conducted_label
            else:
                when, label = target_raw, target_label

        pr_day, pr_label = parse_when(row[8])
        activities.append(
            Activity(
                program=program,
                fund=fund,
                title=title,
                remarks=remarks,
                po_status="With PO" if has_purchase_order(row[12]) else "PO in process",
                pr_submitted=pr_label if pr_day or pr_label != "—" else "—",
                when=when,
                when_label=label,
                conducted=happened,
                billing="Submitted" if billing_submitted(remarks, row[25], row[26]) else "Not yet submitted",
                excluded=excluded,
            )
        )
    return activities


def pct(part: int, whole: int, places: int = 1) -> str:
    if not whole:
        return "—"
    return f"{part / whole * 100:.{places}f}%"


def month_key(day: date | None) -> tuple[int, int] | None:
    if day is None:
        return None
    return (day.year, day.month)


def month_heading(key: tuple[int, int]) -> str:
    year, month = key
    return f"{MONTH_NAMES[month][:3]} {year}"


def span_label(days: list[date]) -> str:
    if not days:
        return "No dated activities"
    start, end = min(days), max(days)
    if start.year == end.year and start.month == end.month:
        return f"{MONTH_NAMES[start.month]} {start.year}"
    if start.year == end.year:
        return f"{MONTH_NAMES[start.month][:3]} – {MONTH_NAMES[end.month][:3]} {start.year}"
    return f"{MONTH_NAMES[start.month][:3]} {start.year} – {MONTH_NAMES[end.month][:3]} {end.year}"


def esc(value) -> str:
    return html.escape(str(value), quote=True)


def donut_svg(segments: list[tuple[int, str]], center: str, caption: str, size: int = 132) -> str:
    radius = 46
    circumference = 2 * math.pi * radius
    total = sum(value for value, _color in segments) or 1
    offset = 0.0
    rings = []
    center_pt = size / 2
    for value, color in segments:
        if value <= 0:
            continue
        length = value / total * circumference
        gap = circumference - length
        rings.append(
            f'<circle cx="{center_pt}" cy="{center_pt}" r="{radius}" fill="none" '
            f'stroke="{color}" stroke-width="16" '
            f'stroke-dasharray="{length:.3f} {gap:.3f}" '
            f'stroke-dashoffset="{-offset:.3f}" '
            f'transform="rotate(-90 {center_pt} {center_pt})"/>'
        )
        offset += length
    if not rings:
        rings.append(
            f'<circle cx="{center_pt}" cy="{center_pt}" r="{radius}" fill="none" '
            f'stroke="#e6e8eb" stroke-width="16"/>'
        )
    return f'''<svg class="donut" viewBox="0 0 {size} {size}" width="{size}" height="{size}" role="img">
      {''.join(rings)}
      <text x="{center_pt}" y="{center_pt - 2}" text-anchor="middle" class="donut-value">{esc(center)}</text>
      <text x="{center_pt}" y="{center_pt + 14}" text-anchor="middle" class="donut-caption">{esc(caption)}</text>
    </svg>'''


def status_class(label: str) -> str:
    return {
        "Submitted": "ok",
        "With PO": "ok",
        "Not yet submitted": "bad",
        "PO in process": "wait",
    }.get(label, "")


def sort_key(activity: Activity):
    return (activity.program.lower(), activity.when or date(9999, 12, 31), activity.title.lower())


def matrix(upcoming: list[Activity]) -> tuple[list[tuple[int, int] | None], list[str], dict]:
    keys = sorted({month_key(item.when) for item in upcoming if item.when})
    columns: list[tuple[int, int] | None] = list(keys)
    if any(item.when is None for item in upcoming):
        columns.append(None)
    funds = [fund for fund in FUND_ORDER if any(item.fund == fund for item in upcoming)]
    extra = sorted({item.fund for item in upcoming if item.fund not in funds})
    funds.extend(extra)
    counts: dict[tuple[str, tuple[int, int] | None], int] = defaultdict(int)
    for item in upcoming:
        counts[(item.fund, month_key(item.when))] += 1
    return columns, funds, counts


def render(activities: list[Activity]) -> str:
    active = [item for item in activities if not item.excluded]
    excluded = [item for item in activities if item.excluded]
    conducted = [item for item in active if item.conducted]
    upcoming = [item for item in active if not item.conducted]
    total = len(active)

    conducted_days = [item.when for item in conducted if item.when]
    upcoming_days = [item.when for item in upcoming if item.when]
    conducted_span = span_label(conducted_days)
    upcoming_span = span_label(upcoming_days)

    funds = [fund for fund in FUND_ORDER if any(item.fund == fund for item in active)]
    extra_funds = sorted({item.fund for item in active if item.fund not in funds})
    funds.extend(extra_funds)

    def count(predicate) -> int:
        return sum(1 for item in active if predicate(item))

    fund_rows = []
    for fund in funds:
        fund_items = [item for item in active if item.fund == fund]
        done = sum(1 for item in fund_items if item.conducted)
        fund_rows.append((fund, len(fund_items), done, len(fund_items) - done))

    submitted = sum(1 for item in conducted if item.billing == "Submitted")
    not_submitted = len(conducted) - submitted
    with_po = sum(1 for item in upcoming if item.po_status == "With PO")
    po_process = len(upcoming) - with_po

    conducted_bits = []
    upcoming_bits = []
    for fund, _all, done, left in fund_rows:
        if done:
            conducted_bits.append(f"{done} {fund}")
        if left:
            upcoming_bits.append(f"{left} {fund}")

    columns, month_funds, month_counts = matrix(upcoming)
    conducted.sort(key=sort_key)
    upcoming.sort(key=sort_key)
    excluded.sort(key=lambda item: (item.program.lower(), item.title.lower()))

    as_of_label = f"{AS_OF.strftime('%B')} {AS_OF.day}, {AS_OF.year}"
    generated = as_of_label

    def fund_table() -> str:
        body = []
        for fund, all_count, done, left in fund_rows:
            width = 0 if not all_count else done / all_count * 100
            body.append(
                f"""<tr>
                  <td>{esc(fund)}</td>
                  <td class="num">{all_count}</td>
                  <td class="num">{done}</td>
                  <td class="pct-cell"><span class="track"><span class="fill" style="width:{width:.1f}%"></span></span><span>{pct(done, all_count)}</span></td>
                  <td class="num peach">{left}</td>
                </tr>"""
            )
        return f"""
        <table class="funds">
          <thead>
            <tr>
              <th>Fund Source</th>
              <th>Total Activities</th>
              <th>Conducted</th>
              <th>% Conducted</th>
              <th>Activities to be Conducted</th>
            </tr>
          </thead>
          <tbody>
            {''.join(body)}
            <tr class="total">
              <td>Total</td>
              <td class="num">{total}</td>
              <td class="num">{len(conducted)}</td>
              <td class="num">{pct(len(conducted), total)}</td>
              <td class="num">{len(upcoming)}</td>
            </tr>
          </tbody>
        </table>"""

    def month_table() -> str:
        heads = "".join(
            f"<th>{'No date' if key is None else esc(month_heading(key))}</th>" for key in columns
        )
        body = []
        for fund in month_funds:
            row_total = sum(month_counts[(fund, key)] for key in columns)
            cells = []
            for key in columns:
                value = month_counts[(fund, key)]
                cells.append(f"<td class='num'>{value or ''}</td>")
            body.append(
                f"<tr><td>{esc(fund)}</td><td class='num mint'>{row_total}</td>{''.join(cells)}</tr>"
            )
        totals = []
        grand = 0
        for key in columns:
            value = sum(month_counts[(fund, key)] for fund in month_funds)
            grand += value
            totals.append(f"<td class='num'>{value or ''}</td>")
        undated = sum(1 for item in upcoming if item.when is None)
        note = ""
        if any(key is not None and date(key[0], key[1], calendar.monthrange(key[0], key[1])[1]) < AS_OF for key in columns):
            note = '<p class="fine">Months already past still have activities that are not marked conducted.</p>'
        undated_note = (
            f'<p class="fine">{undated} activities have no conduct date on the sheet.</p>' if undated else ""
        )
        return f"""
        <table class="months">
          <thead>
            <tr>
              <th>Fund Source</th>
              <th>Activities to be Conducted</th>
              {heads}
            </tr>
          </thead>
          <tbody>
            {''.join(body)}
            <tr class="total"><td>Total</td><td class="num">{grand}</td>{''.join(totals)}</tr>
          </tbody>
        </table>
        {note}{undated_note}"""

    def detail_table(items: list[Activity], mode: str) -> str:
        status_head = "Billing Documents Status" if mode == "billing" else "PO status"
        pr_head = "" if mode == "billing" else "<th>Date PR Submitted</th>"
        rows_html = []
        for item in items:
            label = item.billing if mode == "billing" else item.po_status
            pr_cell = "" if mode == "billing" else f"<td>{esc(item.pr_submitted)}</td>"
            rows_html.append(
                f"""<tr>
                  <td>{esc(item.program)}</td>
                  <td>{esc(item.fund)}</td>
                  <td class="title">{esc(item.title)}</td>
                  <td>{esc(item.when_label)}</td>
                  <td class="{status_class(label)}">{esc(label)}</td>
                  {pr_cell}
                </tr>"""
            )
        return f"""
        <table class="details">
          <thead>
            <tr>
              <th>Program</th>
              <th>Fund Source</th>
              <th>Activity Title</th>
              <th>Date of Conduct</th>
              <th>{status_head}</th>
              {pr_head}
            </tr>
          </thead>
          <tbody>
            {''.join(rows_html)}
          </tbody>
        </table>"""

    def aside_table() -> str:
        if not excluded:
            return ""
        rows_html = []
        for item in excluded:
            rows_html.append(
                f"""<tr>
                  <td>{esc(item.program)}</td>
                  <td>{esc(item.fund)}</td>
                  <td class="title">{esc(item.title)}</td>
                  <td>{esc(item.remarks or "—")}</td>
                </tr>"""
            )
        return f"""
        <section class="sheet">
          <div class="band">SET ASIDE</div>
          <p class="lead">These {len(excluded)} rows stay off the totals. The remarks mark them cancelled, dovetailed into another activity, or recalibrated.</p>
          <table class="details">
            <thead>
              <tr><th>Program</th><th>Fund Source</th><th>Activity Title</th><th>Remarks</th></tr>
            </thead>
            <tbody>{''.join(rows_html)}</tbody>
          </table>
        </section>"""

    overall_donut = donut_svg(
        [(len(conducted), GREEN), (len(upcoming), ORANGE)],
        pct(len(conducted), total, 2),
        "Conducted",
        132,
    )
    billing_donut = donut_svg(
        [(submitted, GREEN), (not_submitted, RED)],
        str(len(conducted)),
        "Conducted",
        104,
    )
    po_donut = donut_svg(
        [(with_po, GREEN), (po_process, ORANGE)],
        str(len(upcoming)),
        "Upcoming",
        104,
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <title>Local Health Support Division — Activities with Catering</title>
  <style>
    @font-face {{
      font-family: Inter;
      src: url("file:///usr/share/fonts/truetype/macos/Inter-Regular.ttf");
      font-weight: 400;
    }}
    @font-face {{
      font-family: Inter;
      src: url("file:///usr/share/fonts/truetype/macos/Inter-Medium.ttf");
      font-weight: 500;
    }}
    @font-face {{
      font-family: Inter;
      src: url("file:///usr/share/fonts/truetype/macos/Inter-SemiBold.ttf");
      font-weight: 600;
    }}
    @font-face {{
      font-family: Inter;
      src: url("file:///usr/share/fonts/truetype/macos/Inter-Bold.ttf");
      font-weight: 700;
    }}
    @page {{
      size: 13.333in 8.15in;
      margin: 0.32in 0.42in 0.3in 0.42in;
    }}
    * {{ box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    html, body {{ margin: 0; padding: 0; }}
    body {{
      font-family: Inter, "Noto Sans", sans-serif;
      color: #24312c;
      background: white;
      font-size: 11px;
    }}
    .dash, .upcoming {{ break-after: page; }}
    h1 {{
      margin: 0;
      text-align: center;
      color: {INK};
      font-size: 20px;
      letter-spacing: 0.08em;
      font-weight: 700;
    }}
    h2 {{
      margin: 1px 0 0;
      text-align: center;
      color: {INK};
      font-size: 15px;
      letter-spacing: 0.06em;
      font-weight: 700;
    }}
    .range {{
      margin: 1px 0 6px;
      text-align: center;
      color: #5d6b64;
      font-size: 11px;
    }}
    .total-card {{
      width: 230px;
      margin: 0 auto 6px;
      border: 1px solid #e4e7e4;
      border-radius: 10px;
      padding: 4px 12px 5px;
      display: flex;
      gap: 10px;
      align-items: center;
      justify-content: center;
      box-shadow: 0 1px 2px rgba(20, 40, 30, 0.04);
    }}
    .total-card .num {{
      font-size: 28px;
      line-height: 1;
      font-weight: 700;
      color: {INK};
      font-variant-numeric: tabular-nums;
    }}
    .total-card .lbl {{
      margin-top: 2px;
      font-size: 9px;
      letter-spacing: 0.08em;
      color: #6b7671;
      font-weight: 600;
    }}
    .card {{
      border: 1px solid #e6e8e6;
      border-radius: 10px;
      padding: 6px 10px 7px;
      margin-bottom: 6px;
    }}
    .card h3, .subhead {{
      margin: 0 0 6px;
      color: {INK};
      font-size: 11px;
      letter-spacing: 0.04em;
      font-weight: 700;
    }}
    .status {{
      display: grid;
      grid-template-columns: 1.15fr 1.15fr 1.15fr 0.95fr;
      gap: 8px;
      align-items: center;
    }}
    .stat {{ text-align: center; }}
    .stat .num {{
      font-size: 26px;
      font-weight: 700;
      line-height: 1;
      font-variant-numeric: tabular-nums;
    }}
    .stat .green {{ color: {GREEN}; }}
    .stat .orange {{ color: {ORANGE}; }}
    .stat .name {{ margin-top: 3px; font-weight: 600; color: #31423b; font-size: 11px; }}
    .stat .sub {{ color: #6a756f; font-size: 10px; }}
    .stat .pct {{ margin-top: 2px; font-weight: 700; font-size: 13px; }}
    .legend {{
      display: flex;
      justify-content: center;
      gap: 14px;
      margin-top: 2px;
      color: #4d5a54;
      font-size: 10px;
    }}
    .swatch {{
      display: inline-block;
      width: 8px;
      height: 8px;
      border-radius: 2px;
      margin-right: 4px;
    }}
    .note {{
      border-left: 1px solid #e3e6e3;
      padding: 4px 8px 4px 12px;
      color: #3d4c45;
      font-size: 11px;
      line-height: 1.35;
    }}
    .note strong {{ color: {INK}; }}
    .check {{
      width: 18px; height: 18px; border-radius: 50%;
      background: #e7f6ef; color: {GREEN};
      display: inline-flex; align-items: center; justify-content: center;
      font-weight: 700; margin-bottom: 4px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-variant-numeric: tabular-nums;
    }}
    th {{
      background: {INK};
      color: white;
      font-weight: 600;
      text-align: left;
      padding: 5px 7px;
      font-size: 10px;
    }}
    td {{
      padding: 4px 7px;
      border-bottom: 1px solid #eef0ee;
      vertical-align: top;
    }}
    .funds td {{ font-size: 11px; }}
    .num {{ text-align: right; }}
    th.num, .funds th {{ text-align: left; }}
    .funds th:not(:first-child), .months th:not(:first-child) {{ text-align: right; }}
    .peach {{ background: #fff4e8; }}
    .mint {{ background: #e7f3d4; font-weight: 600; }}
    tr.total td {{ background: #f8efe4; font-weight: 700; }}
    .months tr.total td {{ background: #f3f5f3; }}
    .pct-cell {{ white-space: nowrap; }}
    .track {{
      display: inline-block;
      width: 72px;
      height: 8px;
      background: #e8eee9;
      border-radius: 99px;
      vertical-align: middle;
      margin-right: 6px;
      overflow: hidden;
    }}
    .fill {{ display: block; height: 100%; background: {GREEN}; border-radius: 99px; }}
    .split {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }}
    .mini {{
      display: grid;
      grid-template-columns: 108px 1fr;
      gap: 8px;
      align-items: center;
    }}
    .mini ul {{ list-style: none; margin: 0; padding: 0; }}
    .mini li {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      padding: 3px 0;
      border-bottom: 1px solid #f1f3f1;
      font-size: 11px;
    }}
    .mini .count {{ font-weight: 700; }}
    .ok {{ color: {GREEN}; font-weight: 600; }}
    .bad {{ color: {RED}; font-weight: 600; }}
    .wait {{ color: #d97816; font-weight: 600; }}
    .donut-value {{
      font-family: Inter, sans-serif;
      font-size: 16px;
      font-weight: 700;
      fill: {INK};
    }}
    .donut-caption {{
      font-family: Inter, sans-serif;
      font-size: 7px;
      fill: #66716c;
      letter-spacing: 0.02em;
    }}
    .band {{
      background: {INK};
      color: white;
      text-align: center;
      font-weight: 700;
      letter-spacing: 0.08em;
      font-size: 16px;
      padding: 8px 12px;
      border-radius: 6px;
      margin-bottom: 10px;
    }}
    .kpis {{
      display: grid;
      grid-template-columns: 1.4fr 1fr 1fr;
      border: 1px solid #e6e8e6;
      border-radius: 10px;
      margin-bottom: 10px;
    }}
    .kpi {{
      text-align: center;
      padding: 8px 6px 9px;
    }}
    .kpi + .kpi {{ border-left: 1px solid #e6e8e6; }}
    .kpi .num {{ font-size: 26px; font-weight: 700; color: {INK}; line-height: 1; }}
    .kpi .lbl {{
      margin-top: 3px;
      font-size: 9px;
      letter-spacing: 0.06em;
      color: #6b7671;
      font-weight: 600;
    }}
    .lead {{ margin: 0 0 8px; color: #4c5a54; }}
    .fine {{ margin: 4px 0 0; color: #6d7872; font-size: 9px; }}
    .details th {{ font-size: 9px; }}
    .details td {{ font-size: 9px; line-height: 1.3; padding: 3px 6px; }}
    .details tr {{ break-inside: avoid; }}
    .details thead {{ display: table-header-group; }}
    .details tbody tr:nth-child(even) td {{ background: #f7f8f7; }}
    .title {{ width: 46%; }}
    .source {{
      margin-top: 8px;
      color: #7a847e;
      font-size: 9px;
    }}
    .rule {{
      margin: 2px 0 0;
      color: #66706b;
      font-size: 8px;
      line-height: 1.3;
    }}
  </style>
</head>
<body>
  <section class="sheet dash">
    <h1>LOCAL HEALTH SUPPORT DIVISION</h1>
    <h2>ACTIVITIES WITH CATERING</h2>
    <p class="range">January – December 2026 · as of {esc(as_of_label)}</p>
    <div class="total-card">
      <svg width="28" height="28" viewBox="0 0 28 28" aria-hidden="true">
        <rect x="3" y="5" width="22" height="19" rx="3" fill="#e7f4ef" stroke="{GREEN}" stroke-width="1.4"/>
        <path d="M3 10h22" stroke="{GREEN}" stroke-width="1.4"/>
        <path d="M9 3v5M19 3v5" stroke="{GREEN}" stroke-width="1.6" stroke-linecap="round"/>
      </svg>
      <div>
        <div class="num">{total}</div>
        <div class="lbl">TOTAL ACTIVITIES</div>
      </div>
    </div>

    <div class="card">
      <h3>OVERALL ACTIVITY STATUS</h3>
      <div class="status">
        <div class="stat">
          <div class="num green">{len(conducted)}</div>
          <div class="name">Conducted Activities</div>
          <div class="sub">{esc(conducted_span)}</div>
          <div class="pct green">{pct(len(conducted), total, 2)}</div>
        </div>
        <div>
          {overall_donut}
          <div class="legend">
            <span><i class="swatch" style="background:{GREEN}"></i>Conducted</span>
            <span><i class="swatch" style="background:{ORANGE}"></i>To be Conducted</span>
          </div>
        </div>
        <div class="stat">
          <div class="num orange">{len(upcoming)}</div>
          <div class="name">To be Conducted</div>
          <div class="sub">{esc(upcoming_span)}</div>
          <div class="pct orange">{pct(len(upcoming), total, 2)}</div>
        </div>
        <div class="note">
          <div class="check">✓</div>
          <div>As of {esc(as_of_label)}, <strong>{len(conducted)} of {total}</strong> activities have been conducted.</div>
        </div>
      </div>
    </div>

    <div class="card">
      <h3>FUNDING SOURCE SUMMARY</h3>
      {fund_table()}
    </div>

    <div class="split">
      <div class="card">
        <h3>CONDUCTED ACTIVITIES</h3>
        <div class="subhead" style="font-weight:500;color:#66716c">Billing Documents Submission · {esc(conducted_span)}</div>
        <div class="mini">
          {billing_donut}
          <ul>
            <li><span><i class="swatch" style="background:{GREEN}"></i>Submitted</span><span class="count ok">{submitted} ({pct(submitted, len(conducted), 2)})</span></li>
            <li><span><i class="swatch" style="background:{RED}"></i>Not yet submitted</span><span class="count bad">{not_submitted} ({pct(not_submitted, len(conducted), 2)})</span></li>
          </ul>
        </div>
      </div>
      <div class="card">
        <h3>UPCOMING ACTIVITIES</h3>
        <div class="subhead" style="font-weight:500;color:#66716c">Purchase Order Status · {esc(upcoming_span)}</div>
        <div class="mini">
          {po_donut}
          <ul>
            <li><span><i class="swatch" style="background:{GREEN}"></i>With PO</span><span class="count ok">{with_po} ({pct(with_po, len(upcoming), 2)})</span></li>
            <li><span><i class="swatch" style="background:{ORANGE}"></i>PO in process</span><span class="count wait">{po_process} ({pct(po_process, len(upcoming), 2)})</span></li>
          </ul>
        </div>
      </div>
    </div>
    <p class="rule">Conducted means the date of conduct is on or before {esc(as_of_label)}, or the row is flagged Conducted with no later date. “Not conducted” stays open. Billing counts as submitted when the row is paid, forwarded to finance, or marked billing on process. A PO counts when a PO number is present. Typed target years after 2026 are read as 2026. {len(excluded)} cancelled, dovetailed, or recalibrated rows are listed at the end and are not in these totals.</p>
    <p class="source">Source: LHS Division Fund Utilization Report and Catering Activity Monitoring Sheet · Catering Services tab · Date generated: {esc(generated)}</p>
  </section>

  <section class="sheet upcoming">
    <div class="band">UPCOMING ACTIVITIES</div>
    <div class="kpis">
      <div class="kpi"><div class="num">{len(upcoming)}</div><div class="lbl">ACTIVITIES TO BE CONDUCTED</div></div>
      <div class="kpi"><div class="num">{with_po}</div><div class="lbl">WITH PO</div></div>
      <div class="kpi"><div class="num">{po_process}</div><div class="lbl">PO IN PROCESS</div></div>
    </div>
    <div class="card">
      <h3>UPCOMING ACTIVITIES BY MONTH</h3>
      <div class="subhead" style="font-weight:500;color:#66716c">{esc(upcoming_span)}</div>
      {month_table()}
    </div>
    <h3 class="subhead" style="margin-top:12px">ACTIVITY DETAILS</h3>
    <div class="subhead" style="font-weight:500;color:#66716c">Purchase order status</div>
    {detail_table(upcoming, "po")}
    <p class="source">Source: Catering Services tab · Date generated: {esc(generated)}</p>
  </section>

  <section class="sheet">
    <div class="band">CONDUCTED ACTIVITIES</div>
    <div class="card">
      <div class="status" style="grid-template-columns: 0.9fr 1.3fr">
        <div class="stat" style="text-align:left;padding-left:8px">
          <div class="subhead">CONDUCTED ACTIVITIES</div>
          <div class="sub">{esc(conducted_span)}</div>
          <div class="num green" style="margin-top:4px">{len(conducted)}</div>
          <div class="name">Activities Conducted</div>
          <div class="sub">{esc(" · ".join(conducted_bits))}</div>
        </div>
        <div>
          <div class="subhead">Billing Documents Submission</div>
          <div class="mini">
            {billing_donut}
            <ul>
              <li><span><i class="swatch" style="background:{GREEN}"></i>Submitted</span><span class="count ok">{submitted} ({pct(submitted, len(conducted), 2)})</span></li>
              <li><span><i class="swatch" style="background:{RED}"></i>Not yet submitted</span><span class="count bad">{not_submitted} ({pct(not_submitted, len(conducted), 2)})</span></li>
            </ul>
          </div>
        </div>
      </div>
    </div>
    <h3 class="subhead">ACTIVITY DETAILS</h3>
    <div class="subhead" style="font-weight:500;color:#66716c">Billing documents submission status</div>
    {detail_table(conducted, "billing")}
    <p class="source">Source: Catering Services tab · Date generated: {esc(generated)}</p>
  </section>

  {aside_table()}
</body>
</html>
"""


def write_pdf(html_path: Path, pdf_path: Path) -> None:
    profile = Path("/tmp/chrome-report-pdf")
    if profile.exists():
        subprocess.run(["rm", "-rf", str(profile)], check=False)
    if pdf_path.exists():
        pdf_path.unlink()
    cmd = [
        "google-chrome",
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-extensions",
        f"--user-data-dir={profile}",
        "--no-pdf-header-footer",
        "--virtual-time-budget=12000",
        f"--print-to-pdf={pdf_path}",
        html_path.as_uri(),
    ]
    try:
        subprocess.run(cmd, check=False, timeout=30, capture_output=True)
    except subprocess.TimeoutExpired:
        subprocess.run(["pkill", "-f", "user-data-dir=/tmp/chrome-report-pdf"], check=False)
    if not pdf_path.exists() or pdf_path.stat().st_size < 1000:
        raise SystemExit(f"PDF was not written: {pdf_path}")


def main() -> None:
    activities = build_activities(load_rows(XLSX))
    active = [item for item in activities if not item.excluded]
    conducted = [item for item in active if item.conducted]
    upcoming = [item for item in active if not item.conducted]
    print(f"rows {len(activities)} active {len(active)} conducted {len(conducted)} upcoming {len(upcoming)} excluded {len(activities) - len(active)}")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    HTML_PATH.write_text(render(activities), encoding="utf-8")
    write_pdf(HTML_PATH, PDF_PATH)
    print(PDF_PATH)


if __name__ == "__main__":
    main()
