"""
Patches the user's manually-edited BI_1409_Freight_Cost_Report_1.xlsx in
place (preserving every row they deleted or hand-edited) rather than
regenerating the whole workbook. Three changes, each scoped as tightly as
requested:

1. "NL - Battery + UK": was falling back to DBS/Ship Cost Matrix when the
   customer wasn't in the MNT UK price list at all. Per the request, this
   population must be priced ONLY from MNT UK price list or Q3 prices
   AUGUST -- no other fallback. Re-priced: try MNT UK first, then Q3
   AUGUST by Shipment Number ("PL nr."); if neither matches, left
   unpriced and flagged (was previously silently priced via the matrix
   fallback, e.g. Krannich Solar M.E.P.E. Greece Battery line).

2. Any NL warehouse (3PLDBSNL/3PLDBSBRNL) line to Greece was classified
   as export (Ship Cost Matrix) because "Parent Region" is not EUROPE for
   Greece in the source data -- a data quirk, not a real routing fact:
   DBS Price list 2026 has a full GR sheet/zone table. Per the request
   (3PLDBSNL -> Krannich Solar M.E.P.E., Greece), every NL-to-Greece line
   is now re-priced via DBS Price list 2026 (zone + total pallets),
   taking priority even over the Battery+UK rule above (there was one
   Battery-family Greece line).

3. "Still duplicate rows by Shipment Number" -- investigated and NOT a
   bug: every case checked (e.g. SE Order# 658608, 27 apparently-identical
   "Optimizers, 1 pallet" Backlog lines; Shipment Number SH21526983012,
   3 apparently-identical "Battery, 20 pallet" Shipped lines) matches the
   source data's own distinct Line#/Part Number count exactly -- these
   are genuinely separate physical order lines that just happen to share
   every column this report shows (a Line# column was never included).
   Added a "Line#" column (from source, informational only) to make this
   self-evident without re-deriving any cost.

Only rows that (a) fall into one of the two fixes above and (b) still
carry the Cost this script's own prior run originally computed for that
exact line (matched by Shipment Number, SE Order#, Customer Name, Family
Type and pallets, in source order) are touched -- any row the user has
since zeroed out or hand-edited is left exactly as they left it, per the
explicit request. No row is ever added or removed.
"""
import math
import os
import sys
from collections import defaultdict, deque

import openpyxl
from openpyxl.styles import Font

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'scripts'))
import bi_1409_freight_cost_report as core

IN_FILE = os.path.join(BASE_DIR, 'data', 'BI_1409_Freight_Cost_Report_user_edit.xlsx')
# Stable, never-overwritten snapshot of the report as originally delivered --
# NOT output/BI_1409_Freight_Cost_Report.xlsx, which this script itself
# overwrites on every run (comparing against a path this script writes to
# would compare the user's file against its own prior patched output on any
# second run, silently treating every already-fixed row as "user-edited").
ORIGINAL_FILE = os.path.join(BASE_DIR, 'data', 'BI_1409_Freight_Cost_Report_original_baseline.xlsx')
OUT_FILE = os.path.join(BASE_DIR, 'output', 'BI_1409_Freight_Cost_Report.xlsx')
SOURCE_FILE = core.MAIN_FILE

COLS = ['Source', 'Shipment Number', 'SE Order#', 'Sending WHS Code', 'Dest WHS Code',
        'Forwarder', 'ShipMode', 'Customer Name', 'Destination country', 'ZIP',
        'Family Type', 'A/I', '# of Pallets per line-up', 'Population', 'Method',
        'Zone / Route', 'Cost', 'Currency', 'Cost (USD)', 'Note']
COL = {name: i + 1 for i, name in enumerate(COLS)}
LINE_COL = len(COLS) + 1  # new "Line#" column appended at the end


def row_key(get):
    return (get('Source'), get('Shipment Number'), get('SE Order#'), get('Customer Name'),
            get('Family Type'), get('# of Pallets per line-up'))


def build_original_queues(sheet_name):
    """{key: deque of original (cost, currency) in source row order}, so we
    can tell whether the user's current row still matches what this script
    itself originally computed for that exact line."""
    ws = openpyxl.load_workbook(ORIGINAL_FILE, data_only=True)[sheet_name]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    queues = defaultdict(deque)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in r):
            continue
        get = lambda h: r[idx[h]]
        queues[row_key(get)].append((r[idx['Cost']], r[idx['Currency']]))
    return queues


def build_line_number_index():
    """(Shipment Number, SE Order#, Customer Name, Family Type, pallets) ->
    deque of Line# in source row order, for the informational Line# column."""
    wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)
    ws = wb['BI 1409']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    queues = defaultdict(deque)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in r):
            continue
        key = (r[idx['Source']], r[idx['Shipment Number']], r[idx['SE Order#']], r[idx['Customer Name']],
               r[idx['Family Type']], r[idx['# of Pallets per line-up']])
        queues[key].append(r[idx['Line#']])
    return queues


def is_untouched(ws, r, orig_queue):
    if not orig_queue:
        return False  # no original record for this exact line -- don't guess, leave it alone
    get = lambda h: ws.cell(row=r, column=COL[h]).value
    cost = ws.cell(row=r, column=COL['Cost']).value
    currency = ws.cell(row=r, column=COL['Currency']).value
    orig_cost, orig_currency = orig_queue[0]
    match = (_close(cost, orig_cost) and currency == orig_currency)
    if match:
        orig_queue.popleft()
    return match


def _close(a, b):
    if a is None or b is None:
        return a == b
    return abs(a - b) < 0.005


def add_line_number_column(ws, line_queues):
    ws.cell(row=1, column=LINE_COL, value='Line# (source)').font = Font(name='Arial', bold=True)
    for r in range(2, ws.max_row + 1):
        get = lambda h: ws.cell(row=r, column=COL[h]).value
        key = row_key(get)
        q = line_queues.get(key)
        line_no = q.popleft() if q else None
        ws.cell(row=r, column=LINE_COL, value=line_no).font = Font(name='Arial')


def fix_battery_uk_strict(ws, mnt_uk_book, q3_book, orig_queues):
    """Re-price 'NL - Battery + UK' rows using ONLY MNT UK / Q3 AUGUST --
    remove the DBS/matrix fallback. Grouped by Shipment Number (or SE
    Order# for Backlog) within each source, same flat-rate-per-shipment
    allocation as everywhere else."""
    max_row = ws.max_row
    candidates = []
    for r in range(2, max_row + 1):
        get = lambda h: ws.cell(row=r, column=COL[h]).value
        if get('Population') != 'NL - Battery + UK':
            continue
        if get('Destination country') == 'Greece':
            continue  # Greece takes priority -- handled by fix_greece_to_dbs
        if not is_untouched(ws, r, orig_queues.get(row_key(get))):
            continue
        candidates.append(r)

    mnt_uk_groups = defaultdict(list)
    q3_groups = defaultdict(list)
    unmatched = []
    for r in candidates:
        get = lambda h: ws.cell(row=r, column=COL[h]).value
        sh = get('Shipment Number')
        gkey = sh if sh and sh != 'N/A' else get('SE Order#')
        customer = get('Customer Name')
        zip_code = get('ZIP')
        _, by_name = mnt_uk_book
        if customer in by_name:
            mnt_uk_groups[(gkey, customer, zip_code)].append(r)
        elif sh in q3_book:
            q3_groups[sh].append(r)
        else:
            unmatched.append(r)

    n_fixed = 0
    for (gkey, customer, zip_code), rows in mnt_uk_groups.items():
        price, note = core.mnt_uk_lookup(mnt_uk_book, customer, zip_code)
        _apply_group(ws, rows, price, 'EUR' if price is not None else None,
                     'MNT UK price list', 'MNT UK (Price Q3)', note,
                     f'MNT UK full-truck rate {price} EUR allocated across the shipment.' if price is not None
                     else None)
        n_fixed += len(rows)

    for sh, rows in q3_groups.items():
        info = q3_book[sh]
        total_pallets = sum((ws.cell(row=r, column=COL['# of Pallets per line-up']).value or 0) for r in rows)
        _apply_group(ws, rows, info['total_eur'], 'EUR', 'Q3 AUGUST (MNT, POD)', f'Q3 AUGUST (PL nr. {sh})', None,
                     f'Q3 AUGUST shipment-level rate {info["total_eur"]} EUR allocated by pallet share.')
        if info['duty_gbp']:
            for r in rows:
                pallets = ws.cell(row=r, column=COL['# of Pallets per line-up']).value or 0
                share = (pallets / total_pallets) if total_pallets else (1 / len(rows))
                note_cell = ws.cell(row=r, column=COL['Note'])
                note_cell.value = (note_cell.value + f' Duty {round(info["duty_gbp"] * share, 2)} GBP not '
                                    'included (different currency).')
        n_fixed += len(rows)

    for r in unmatched:
        get = lambda h: ws.cell(row=r, column=COL[h]).value
        ws.cell(row=r, column=COL['Method']).value = 'MNT UK / Q3 AUGUST (no match)'
        ws.cell(row=r, column=COL['Zone / Route']).value = None
        ws.cell(row=r, column=COL['Cost']).value = None
        ws.cell(row=r, column=COL['Currency']).value = None
        ws.cell(row=r, column=COL['Cost (USD)']).value = None
        ws.cell(row=r, column=COL['Note']).value = (
            f"Customer {get('Customer Name')!r} not in MNT UK price list, and Shipment Number "
            f"{get('Shipment Number')!r} not on file in Q3 AUGUST -- this population is priced only from "
            f"those two sources, confirm with WH contact.")
        n_fixed += 1

    print(f"  Re-priced {n_fixed} 'NL - Battery + UK' line(s) (MNT UK / Q3 AUGUST only, no fallback).")


def fix_greece_to_dbs(ws, dbs_book, orig_queues):
    """Any NL-warehouse line to Greece: re-price via DBS Price list 2026,
    grouped by Shipment Number (or SE Order#), overriding whatever
    population/method it was in before (Ship Cost Matrix export, or the
    Battery+UK fallback)."""
    max_row = ws.max_row
    candidates = []
    for r in range(2, max_row + 1):
        get = lambda h: ws.cell(row=r, column=COL[h]).value
        if get('Sending WHS Code') not in core.NL_WHS or get('Destination country') != 'Greece':
            continue
        if not is_untouched(ws, r, orig_queues.get(row_key(get))):
            continue
        candidates.append(r)

    groups = defaultdict(list)
    for r in candidates:
        get = lambda h: ws.cell(row=r, column=COL[h]).value
        sh = get('Shipment Number')
        gkey = sh if sh and sh != 'N/A' else get('SE Order#')
        groups[gkey].append(r)

    n_fixed = 0
    for gkey, rows in groups.items():
        sample_zip = ws.cell(row=rows[0], column=COL['ZIP']).value
        total_pallets = sum((ws.cell(row=r, column=COL['# of Pallets per line-up']).value or 0) for r in rows)
        zone, price, note = core.dbs_lookup(dbs_book, 'GR', sample_zip, total_pallets or 1)
        base_note = f'DBS Price list 2026, zone {zone}: {price} EUR for {total_pallets:g} total pallets.' \
            if price is not None else None
        _apply_group(ws, rows, price, 'EUR' if price is not None else None,
                     'DBS Price list 2026', zone, note, base_note)
        for r in rows:
            ws.cell(row=r, column=COL['Population']).value = 'NL - Domestic'
        n_fixed += len(rows)

    print(f"  Re-priced {n_fixed} NL-to-Greece line(s) via DBS Price list 2026 (was Ship Cost Matrix export).")


def _apply_group(ws, rows, flat_cost, currency, method_label, route, extra_note, base_note):
    total_pallets = sum((ws.cell(row=r, column=COL['# of Pallets per line-up']).value or 0) for r in rows)
    for r in rows:
        ws.cell(row=r, column=COL['Method']).value = method_label
        ws.cell(row=r, column=COL['Zone / Route']).value = route
        if flat_cost is None:
            ws.cell(row=r, column=COL['Cost']).value = None
            ws.cell(row=r, column=COL['Currency']).value = None
            ws.cell(row=r, column=COL['Cost (USD)']).value = None
            ws.cell(row=r, column=COL['Note']).value = extra_note or 'Not priced -- confirm with WH contact.'
            continue
        pallets = ws.cell(row=r, column=COL['# of Pallets per line-up']).value or 0
        share = (pallets / total_pallets) if total_pallets else (1 / len(rows))
        cost = round(flat_cost * share, 2)
        ws.cell(row=r, column=COL['Cost']).value = cost
        ws.cell(row=r, column=COL['Currency']).value = currency
        ws.cell(row=r, column=COL['Cost (USD)']).value = core.to_usd(cost, currency)
        note = base_note or ''
        if len(rows) > 1:
            note += f' Allocated pro-rata by pallet share across this shipment\'s {len(rows)} lines.'
        if extra_note:
            note = (note + ' ' + extra_note).strip()
        ws.cell(row=r, column=COL['Note']).value = note.strip()


def rebuild_summary(wb):
    """Static-value Summary sheet, rebuilt from the just-patched Shipped/Backlog
    sheets so it reflects both the user's own edits/deletions and this
    round's fixes (the original was written as static numbers, not formulas)."""
    if 'Summary' in wb.sheetnames:
        del wb['Summary']
    ws = wb.create_sheet('Summary')
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    headers = ['Population', 'Status', 'Lines', 'Priced', 'Unpriced', 'Total Cost (USD)']
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold

    by_pop_status = defaultdict(lambda: {'lines': 0, 'priced': 0, 'cost': 0.0})
    for sheet_name in ('Shipped', 'Backlog'):
        ws_data = wb[sheet_name]
        idx = {c.value: i for i, c in enumerate(ws_data[1])}
        for r in ws_data.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in r):
                continue
            key = (r[idx['Population']], sheet_name)
            agg = by_pop_status[key]
            agg['lines'] += 1
            if r[idx['Cost (USD)']] is not None:
                agg['priced'] += 1
                agg['cost'] += r[idx['Cost (USD)']]

    populations = sorted(set(k[0] for k in by_pop_status))
    grand = {'lines': 0, 'priced': 0, 'cost': 0.0}
    for pop in populations:
        for status in ('Shipped', 'Backlog'):
            agg = by_pop_status.get((pop, status))
            if not agg:
                continue
            ws.append([pop, status, agg['lines'], agg['priced'], agg['lines'] - agg['priced'], round(agg['cost'], 2)])
            for c in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=c).font = arial
            ws.cell(row=ws.max_row, column=6).number_format = '#,##0.00'
            grand['lines'] += agg['lines']
            grand['priced'] += agg['priced']
            grand['cost'] += agg['cost']

    ws.append(['Total', None, grand['lines'], grand['priced'], grand['lines'] - grand['priced'], round(grand['cost'], 2)])
    for c in range(1, len(headers) + 1):
        ws.cell(row=ws.max_row, column=c).font = bold
    ws.cell(row=ws.max_row, column=6).number_format = '#,##0.00'

    from openpyxl.utils import get_column_letter
    for c, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(22, len(h) + 4)


def update_readme(wb):
    ws = wb['README']
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    ws.append((None,))
    ws.append(('Patched per user review:',))
    ws.cell(row=ws.max_row, column=1).font = bold
    lines = [
        '1) "NL - Battery + UK" is now priced ONLY from MNT UK price list or Q3 prices AUGUST -- the DBS/Ship '
        'Cost Matrix fallback used when a customer wasn\'t in MNT UK at all has been removed; unmatched lines '
        'are left unpriced and flagged instead.',
        '2) Any NL warehouse (3PLDBSNL/3PLDBSBRNL) line to Greece is now priced via DBS Price list 2026 (zone '
        'GR+ZIP, total pallets), not Ship Cost Matrix -- Greece has a full DBS zone table; "Parent Region" '
        'marking it non-Europe in the source data was a data quirk, not a real routing fact. This takes '
        'priority even over the Battery+UK rule above (one Battery-family Greece line).',
        '3) "Duplicate rows by Shipment Number" investigated and confirmed NOT a bug: every case checked '
        '(e.g. SE Order# 658608, 27 apparently-identical Backlog lines; Shipment Number SH21526983012, 3 '
        'apparently-identical Shipped lines) matches the source data\'s own distinct Line# count exactly -- '
        'genuinely separate order lines this report never showed a Line# for. Added a "Line# (source)" '
        'column so this is self-evident without re-deriving any cost.',
        'Only rows still carrying the Cost this report originally computed for them were touched by fixes 1-2 '
        '-- every row the user zeroed out or hand-edited was left exactly as they left it, and no row was '
        'added or removed.',
    ]
    for text in lines:
        ws.append((text,))
        ws.cell(row=ws.max_row, column=1).font = arial


def main():
    print("Loading price sources...")
    dbs_book = core.parse_dbs_price_book(core.DBS_FILE)
    mnt_uk_book = core.parse_mnt_uk(core.MNT_UK_FILE)
    q3_book = core.parse_q3_august(core.Q3_AUG_FILE)

    wb = openpyxl.load_workbook(IN_FILE)

    for sheet_name in ('Shipped', 'Backlog'):
        print(f"{sheet_name}:")
        ws = wb[sheet_name]
        orig_queues = build_original_queues(sheet_name)

        fix_greece_to_dbs(ws, dbs_book, orig_queues)
        fix_battery_uk_strict(ws, mnt_uk_book, q3_book, orig_queues)
        add_line_number_column(ws, build_line_number_index())

    rebuild_summary(wb)
    update_readme(wb)

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)
    print("Saved", OUT_FILE)


if __name__ == '__main__':
    main()
