#!/usr/bin/env python3
"""
Convert a volunteer-scheduling JSON export into an Excel "Services Master
Schedule" laid out like the reference PDF.

Layout (one block per schedule/week, stacked vertically):

    Services Master Schedule
                 +--------------------------------------------------------+
                 |                Nov 8, 2026 - Nov 14, 2026              |
                 |   Sunday School   | AM Worship Service | PM Worship ... |
                 | Sun, Nov 8 9:30 AM| Sun, Nov 8 11:00 AM| ...            |
    AV Monitor   | David White | Accepted | Randy Ayscue | Accepted | ...  |
    ...          |                          |       (blank) | Open   | ... |

* Each event gets two sub-columns: volunteer name (left-aligned) and status
  (right-aligned).
* Accepted = green (#D9EAD3), Declined = red (#F4CCCC), Open / empty = white.
  Header and position-label cells = blue (#CFE2F3).
* A position appears once per volunteer slot (the label repeats on each row,
  as in the PDF).  Positions with no volunteers and no open slots in a given
  week are omitted for that week.
* Open slots are rendered as a blank name plus "Open" in the status column,
  using the JSON's `needed` count for that position.
* Column widths are computed from the text that is actually displayed (using
  Calibri-compatible font metrics when available), so nothing is cut off or
  wrapped.

Usage:
    python schedule_to_xlsx.py schedule.json [more.json ...] -o schedule.xlsx

Each input file may contain a single schedule object or a list of them.  All
schedules are sorted by start date and stacked on one sheet, so passing one
JSON per week yields a multi-week sheet like the PDF.

Requires: openpyxl (Pillow optional, for accurate width measurement).
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
import dateparser
from pathlib import Path
from pprint import pprint

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

CONFIG_PATH = Path(__file__).with_name("config.json")
with CONFIG_PATH.open(encoding="utf-8-sig") as config_file:
    config = json.load(config_file)

TITLE = "Services Master Schedule"
SHEET_NAME = "Services Master Schedule"

# Column order of events (matched case-insensitively by name).  Events that
# are not listed are placed after these, in order of first appearance.
EVENT_ORDER = ["Sunday School", "AM Worship Service", "PM Worship Service", "Prayer Meeting"]

BLUE = "CFE2F3"    # headers + position labels   (sampled from the PDF)
GREEN = "D9EAD3"   # accepted                    (sampled from the PDF)
RED = "F4CCCC"     # declined                    (sampled from the PDF: #F4CCCD)

FONT_NAME = "Calibri"
FONT_SIZE = 11
TITLE_SIZE = 20

# Blank rows between stacked weeks.  The PDF has none.
BLOCK_GAP_ROWS = 0

# Extra width (in Excel character units) added to the measured text width.
PAD_CHARS = 1.6
MIN_STATUS_WIDTH = 9.0

STATUS_LABELS = {"ACCEPTED": "Accepted", "DECLINED": "Declined"}
STATUS_KIND = {"ACCEPTED": "accepted", "DECLINED": "declined"}

DAY_ABBR = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Category numbers for schedules. Get these from the links on https://grbc.ccbchurch.com/goto/scheduling/main/categories
# As of 9/25/2021, 5=Nursery, 16=Service Leaders, 2=Services Master Schedule
# - Also lists the Sunday and Wednesday event names we'll search for (which are the headers at the top of each schedule)
# - Also lists the names of the teams we're looking for (which appear below each "event name" in the columns in each schedule)
category_map = {
    'master': {
        'category_id': '2',
        'sun_event_names': {'Sunday School', 'AM Worship Service', 'PM Worship Service'},
        'wed_event_names': {'Prayer Meeting'},
        'teams_to_track':  {'Audio Visual', 'Bookroom', 'Deacons', 'YouTube Editors'},
    },
    'nursery': {
        'category_id': '5',
        'sun_event_names': {'Sunday School Nursery', 'Sunday AM Nursery', 'Sunday PM Nursery'},
        'wed_event_names': {'Prayer Meeting Nursery'},
        'teams_to_track':  {'AM', 'PM', 'SS', 'Wed'},
    },
    'service_leaders': {
        'category_id': '16',
        'sun_event_names': {'AM Worship Service', 'PM Worship Service'},
        'wed_event_names': {},
        'teams_to_track':  {},
    },
}

# We want the next Sunday schedule (or today, if today is Sunday)
now = datetime.now().date()
target_sunday = now + timedelta(days=7 - now.isoweekday())

login_url = f"{config['ccb_api_url']}/login"
schedule_list_url  = f"{config['ccb_api_url']}/scheduling/categories/{{category_id}}/schedules?page=1&with_metrics=1&after={{lastweek}}"
schedule_fetch_url = f'{config["ccb_api_url"]}/scheduling/categories/{{category_id}}/schedules?get_full_schedules=true&schedule_ids[]={{schedule_id}}'

ccb_session = requests.Session()
authlogin = ccb_session.post(
    login_url,
    data={
        "username": config["ccb_api_username"],
        "password": config["ccb_api_password"],
        "keep_me_logged_in": False,
        "accept_privacy_policy": False,
    },
    timeout=30,
)
authlogin.raise_for_status()
cookies = authlogin.cookies

print(f'[*] Getting all schedule IDs')
schedule_found = False
lastweek = (now - timedelta(days=7)).strftime('%Y-%m-%d')
print(f'[*] Using lastweek date of "{lastweek}"')

req = schedule_list_url.format(category_id=category_map['master']['category_id'], lastweek=lastweek)
print(f'[*] Calling: {req}')
r = requests.get(req, cookies=cookies)
try:
    data = r.json()
    for schedule in data:
        schedule_id = schedule.get('id')
        schedule_start = dateparser.parse(schedule.get('start')).date()
        schedule_end   = dateparser.parse(schedule.get('end')).date()
        print(f'  [*] Schedule ID {schedule_id} runs from {schedule_start} to {schedule_end}')
        if schedule_start == target_sunday:
            print(f'  [*] Correct schedule found (ID {schedule_id})')
            schedule_found = True
            break
except json.decoder.JSONDecodeError:
    print("  [!] Did not get valid JSON")
    print("  Instead, got this:", r.text)
    pass


req = schedule_fetch_url.format(category_id=category_map['master']['category_id'], schedule_id=schedule_id)
print(f'  [*] Calling: {req}')
r = requests.get(req, cookies=cookies)
try:
    data = r.json()[0]
except (json.decoder.JSONDecodeError, IndexError):
    print("  [!] Did not get valid JSON")
    raise
    sys.exit()


# --------------------------------------------------------------------------
# Text measurement (to size columns from the displayed text)
# --------------------------------------------------------------------------
class TextMeasurer:
    """Returns text width in Excel column-width units (1 unit ~ one '0' of the
    default font, 7 px for Calibri 11 at 96 dpi)."""

    FONT_FILES = [
        "Carlito-Regular.ttf", "calibri.ttf", "Calibri.ttf",
        "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
        "C:/Windows/Fonts/calibri.ttf",
        "/Library/Fonts/Calibri.ttf",
        "/System/Library/Fonts/Supplemental/Calibri.ttf",
    ]

    def __init__(self, size_pt: float = FONT_SIZE):
        self.font = None
        try:
            from PIL import ImageFont
        except ImportError:
            return
        px = size_pt * 96 / 72
        for name in self.FONT_FILES:
            for size in (px, round(px)):  # older Pillow only accepts ints
                try:
                    self.font = ImageFont.truetype(name, size)
                    return
                except (OSError, TypeError, ValueError):
                    continue

    def width(self, text: str) -> float:
        if not text:
            return 0.0
        if self.font is not None:
            return self.font.getlength(text) / 7.0
        # Fallback estimate: Calibri averages a little under one '0' per char.
        return len(text) * 0.95


MEASURE = TextMeasurer()


# --------------------------------------------------------------------------
# Date helpers
# --------------------------------------------------------------------------
def parse_dt(value) -> datetime | None:
    """Parse the JSON's mixed date formats; returns a naive datetime or None."""
    if not value or str(value).startswith("0000"):
        return None
    text = str(value).strip().replace(" ", "T", 1)
    try:
        dt = datetime.fromisoformat(text)
    except ValueError:
        return None
    return dt.replace(tzinfo=None)


def fmt_date(dt: datetime) -> str:
    return f"{MONTH_ABBR[dt.month - 1]} {dt.day}, {dt.year}"


def fmt_event_time(dt: datetime) -> str:
    hour = dt.hour % 12 or 12
    return (f"{DAY_ABBR[dt.weekday()]}, {fmt_date(dt)} "
            f"{hour}:{dt.minute:02d} {'AM' if dt.hour < 12 else 'PM'}")


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------
@dataclass
class Slot:
    name: str    # "" for an open slot
    status: str  # display text: Accepted / Declined / Open / ...
    kind: str    # accepted | declined | open | other


@dataclass
class Position:
    team_key: str
    name: str
    slots: dict  # column key -> list[Slot]

    @property
    def row_count(self) -> int:
        return max((len(s) for s in self.slots.values()), default=0)


@dataclass
class Block:
    header: str                 # "Nov 8, 2026 - Nov 14, 2026"
    events: dict                # column key -> (event name, "Sun, Nov 8, 2026 9:30 AM")
    positions: list             # list[Position], only those with >= 1 row


def load_schedules(paths) -> list:
    schedules = []
    for path in paths:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            data = data.get("schedules", [data]) if "events" not in data else [data]
        schedules.extend(s for s in data if isinstance(s, dict) and "events" in s)
    schedules.sort(key=lambda s: parse_dt(s.get("start")) or datetime.max)
    return schedules


def event_sort_key(event: dict):
    name = (event.get("name") or "").strip().lower()
    order = [n.lower() for n in EVENT_ORDER]
    idx = order.index(name) if name in order else len(order)
    return (idx, parse_dt(event.get("start")) or datetime.max)


def make_slots(ep: dict, open_mode: str) -> list:
    slots = []
    for a in ep.get("assignments") or []:
        raw = str(a.get("status") or "").upper()
        indiv = (a.get("volunteer") or {}).get("individual") or {}
        name = indiv.get("name") or f"{indiv.get('first_name', '')} {indiv.get('last_name', '')}".strip()
        label = STATUS_LABELS.get(raw, raw.replace("_", " ").title())
        slots.append(Slot(name, label, STATUS_KIND.get(raw, "other")))

    if open_mode == "unfilled":   # quantity minus everyone listed (incl. declined)
        n_open = max(0, int(ep.get("quantity") or 0) - len(slots))
    else:                         # "needed": the JSON's own count of open slots
        n_open = max(0, int(ep.get("needed") or 0))
    slots.extend(Slot("", "Open", "open") for _ in range(n_open))
    return slots


def build_block(schedule: dict, open_mode: str) -> Block:
    events = sorted(schedule.get("events") or [], key=event_sort_key)

    # Column key = (event name, nth occurrence of that name in this schedule).
    seen, keyed = {}, []
    for ev in events:
        name = (ev.get("name") or "").strip()
        seen[name] = seen.get(name, 0) + 1
        keyed.append(((name, seen[name]), ev))

    positions: dict = {}
    col_info = {}
    for key, ev in keyed:
        start = parse_dt(ev.get("start"))
        col_info[key] = (key[0], fmt_event_time(start) if start else "")
        for et in ev.get("event_teams") or []:
            team = et.get("team") or {}
            team_key = str(et.get("team_id", team.get("id", team.get("name", ""))))
            for ep in et.get("event_positions") or []:
                pos = ep.get("position") or {}
                pkey = (team_key, ep.get("position_id", pos.get("id", pos.get("name"))))
                p = positions.setdefault(pkey, Position(team_key, pos.get("name", ""), {}))
                p.slots[key] = make_slots(ep, open_mode)

    start, end = parse_dt(schedule.get("start")), parse_dt(schedule.get("end"))
    if not (start and end):
        dts = [parse_dt(e.get("start")) for e in events if parse_dt(e.get("start"))]
        start, end = (min(dts), max(dts)) if dts else (None, None)
    header = f"{fmt_date(start)} - {fmt_date(end)}" if start and end else (schedule.get("name") or "")

    return Block(header, col_info, [p for p in positions.values() if p.row_count > 0])


# --------------------------------------------------------------------------
# Workbook writer
# --------------------------------------------------------------------------
RANK = {None: 0, "thin": 1, "medium": 2}


class Edges:
    """Accumulates cell borders; the heavier style wins when two meet."""

    def __init__(self):
        self.map = {}

    def add(self, row, col, side, style):
        cur = self.map.setdefault((row, col), {})
        if RANK[style] > RANK[cur.get(side)]:
            cur[side] = style

    def box(self, r1, c1, r2, c2, style):
        for c in range(c1, c2 + 1):
            self.add(r1, c, "top", style)
            self.add(r2, c, "bottom", style)
        for r in range(r1, r2 + 1):
            self.add(r, c1, "left", style)
            self.add(r, c2, "right", style)

    def border(self, row, col) -> Border:
        sides = self.map.get((row, col), {})
        return Border(**{s: Side(style=st, color="000000") for s, st in sides.items()})


def write_workbook(blocks: list, out_path: str, title: str = TITLE):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME[:31]
    ws.sheet_view.showGridLines = False

    # Global event columns, shared by all weeks so that they line up.
    col_keys, first_seen = [], {}
    for b in blocks:
        for key, (name, _) in b.events.items():
            if key not in first_seen:
                first_seen[key] = len(first_seen)
                col_keys.append(key)
    order = [n.lower() for n in EVENT_ORDER]
    col_keys.sort(key=lambda k: (order.index(k[0].lower()) if k[0].lower() in order else len(order),
                                 k[1], first_seen[k]))
    name_col = {k: 2 + 2 * i for i, k in enumerate(col_keys)}   # B, D, F, H
    last_col = 1 + 2 * len(col_keys)

    font = Font(name=FONT_NAME, size=FONT_SIZE)
    fill = {"blue": PatternFill("solid", fgColor=BLUE),
            "accepted": PatternFill("solid", fgColor=GREEN),
            "declined": PatternFill("solid", fgColor=RED)}
    edges = Edges()

    # Width bookkeeping ------------------------------------------------------
    need = {c: 0.0 for c in range(1, last_col + 1)}
    spans = []   # (first col, last col, width needed) for merged header cells

    def need_text(col, text):
        need[col] = max(need[col], MEASURE.width(text) + PAD_CHARS)

    # Title -------------------------------------------------------------------
    ws.cell(1, 1, title).font = Font(name=FONT_NAME, size=TITLE_SIZE, bold=True)
    ws.cell(1, 1).alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 34
    row = 3
    ws.row_dimensions[2].height = 12

    for bi, block in enumerate(blocks):
        if bi:
            row += BLOCK_GAP_ROWS
        top = row

        # Week header (merged across all event columns) -----------------------
        merges = [(row, 2, row, last_col)]
        c = ws.cell(row, 2, block.header)
        c.font, c.fill = font, fill["blue"]
        c.alignment = Alignment(horizontal="center", vertical="center")
        spans.append((2, last_col, MEASURE.width(block.header) + PAD_CHARS))
        for col in range(2, last_col + 1):
            ws.cell(row, col).fill = fill["blue"]

        # Event name + date/time header rows -----------------------------------
        for hr in (row + 1, row + 2):
            for col in range(2, last_col + 1):
                ws.cell(hr, col).fill = fill["blue"]
        for key in col_keys:
            c1 = name_col[key]
            ename, etime = block.events.get(key, ("", ""))
            for hr, text in ((row + 1, ename), (row + 2, etime)):
                cell = ws.cell(hr, c1, text or None)
                cell.font = font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                merges.append((hr, c1, hr, c1 + 1))
                spans.append((c1, c1 + 1, MEASURE.width(text) + PAD_CHARS))
                edges.add(hr, c1, "left", "thin")
                edges.add(hr, c1 + 1, "right", "thin")
        for hr in (row, row + 1, row + 2):          # thin rule under each header row
            for col in range(2, last_col + 1):
                edges.add(hr, col, "bottom", "thin")
        row += 3

        # Body rows ------------------------------------------------------------
        body_first = row
        team_runs = []   # (team_key, first row, last row)
        for pos in block.positions:
            for i in range(pos.row_count):
                a = ws.cell(row, 1, pos.name)
                a.font, a.fill = font, fill["blue"]
                a.alignment = Alignment(horizontal="left", vertical="center")
                need_text(1, pos.name)
                edges.box(row, 1, row, 1, "thin")

                for key in col_keys:
                    c1 = name_col[key]
                    edges.add(row, c1, "left", "thin")
                    edges.add(row, c1 + 1, "right", "thin")
                    slots = pos.slots.get(key, [])
                    if i >= len(slots):
                        continue
                    slot = slots[i]
                    n_cell, s_cell = ws.cell(row, c1), ws.cell(row, c1 + 1)
                    n_cell.value = slot.name or None
                    s_cell.value = slot.status
                    for cell, h in ((n_cell, "left"), (s_cell, "right")):
                        cell.font = font
                        cell.alignment = Alignment(horizontal=h, vertical="center")
                        if slot.kind in fill:
                            cell.fill = fill[slot.kind]
                        edges.add(row, cell.column, "top", "thin")
                        edges.add(row, cell.column, "bottom", "thin")
                    need_text(c1, slot.name)
                    need_text(c1 + 1, slot.status)

                if team_runs and team_runs[-1][0] == pos.team_key and team_runs[-1][2] == row - 1:
                    team_runs[-1] = (pos.team_key, team_runs[-1][1], row)
                else:
                    team_runs.append((pos.team_key, row, row))
                row += 1
        body_last = row - 1

        # Team groups in the label column get a heavier outline.
        for _, r1, r2 in team_runs:
            edges.box(r1, 1, r2, 1, "medium")
        # Heavy outline around the week's event area.
        edges.box(top, 2, max(body_last, top + 2), last_col, "medium")

        # Apply merges (anchor takes the outer edges of the whole range).
        for r1, c1, r2, c2 in merges:
            for col in range(c1, c2 + 1):
                for rr in range(r1, r2 + 1):
                    ws.cell(rr, col).border = edges.border(rr, col)
            anchor = ws.cell(r1, c1)
            anchor.border = Border(
                left=edges.border(r1, c1).left, top=edges.border(r1, c1).top,
                right=edges.border(r1, c2).right, bottom=edges.border(r2, c1).bottom)
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    # Borders for all remaining (non-merged) cells ---------------------------
    for (r, c) in edges.map:
        cell = ws.cell(r, c)
        if type(cell).__name__ != "MergedCell":
            cell.border = edges.border(r, c)

    # Column widths from displayed text --------------------------------------
    for key in col_keys:
        need[name_col[key] + 1] = max(need[name_col[key] + 1], MIN_STATUS_WIDTH)
    for c1, c2, w in sorted(spans, key=lambda s: s[1] - s[0]):   # narrow spans first
        have = sum(need[c] for c in range(c1, c2 + 1))
        if have < w:
            # Give the shortfall to the name (first) column of each event pair.
            target = c1 if c2 - c1 == 1 else c2 - 1
            need[target] += w - have
    for col, w in need.items():
        ws.column_dimensions[get_column_letter(col)].width = math.ceil(max(w, 8.43) * 100) / 100

    # Printing ---------------------------------------------------------------
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(out_path)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Schedule JSON -> Services Master Schedule XLSX")
    ap.add_argument("json_files", nargs="+", help="schedule JSON file(s); one block per schedule/week")
    ap.add_argument("-o", "--output", default="Services_Master_Schedule.xlsx")
    ap.add_argument("--title", default=TITLE)
    ap.add_argument("--open-mode", choices=["needed", "unfilled"], default="needed",
                    help="how many 'Open' rows to show per position: the JSON's `needed` "
                         "(default), or quantity minus everyone listed incl. declined")
    args = ap.parse_args(argv)

    schedules = load_schedules(args.json_files)
    if not schedules:
        print("No schedules found in input.", file=sys.stderr)
        return 1
    blocks = [build_block(s, args.open_mode) for s in schedules]
    write_workbook(blocks, args.output, args.title)
    print(f"Wrote {args.output}: {len(blocks)} week(s), "
          f"{sum(sum(p.row_count for p in b.positions) for b in blocks)} body rows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
