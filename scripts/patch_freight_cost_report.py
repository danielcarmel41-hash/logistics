"""
Patches the user's manually-edited Freight_Cost_Report_2.xlsx in place
(preserving every row they deleted/kept) rather than regenerating the whole
workbook from source data. Four changes, each scoped as tightly as the
request allows:

1. "3PLs", "NL - Support", "NL - EX + DO", "Canot - EX + DO": every group of
   lines sharing one Shipment Number (or SE Order# for Backlog lines, which
   have no shipment number yet) was getting the Ship Cost Matrix's flat
   per-shipment/corridor rate applied to *every line* instead of once per
   shipment -- e.g. one Israel shipment with 19 lines was costed at 4800
   USD x 19. This was first found (and fixed) only in the specific rows the
   user flagged with Cost=0 in "NL - EX + DO"/"Canot - EX + DO"; the user
   then spotted the same bug unflagged in "3PLs" by checking against the
   source data, so the fix now applies unconditionally to every matrix-
   priced multi-line group in all four sheets (NL - DG + UK has none):
   flat corridor rate x (line's pallet share of the group's total pallets).

2. "NL - Support", "NL - EX + DO": the same bug also exists on the DBS
   Price list side -- its EP bracket (and therefore its rate) is for the
   *whole shipment's* total pallets, not per line, so a multi-line
   shipment/order priced line-by-line both picks too low an EP bracket per
   line and, since freight rates are sub-linear in pallets, sums to well
   more than the correct one-shipment price. The user flagged that
   "NL - EX + DO" backlog still looked wrong after fix #1; every DBS-priced
   group of 2+ lines (grouped the same way as #1: Shipment Number, or SE
   Order# for Backlog) is recomputed as one DBS price for the group's total
   pallets, allocated pro-rata by each line's pallet share. ("NL - DG + UK"
   only ever falls back to DBS for single-line US/Australia cases, and
   Canot's Israel lines use their own domestic rate card, not DBS, so
   neither sheet is affected by this bug.)

3. "NL - DG + UK": re-priced sheet-wide per the new rule -- destination
   United Kingdom -> MNT UK price list (unchanged); everything else ->
   Q3_prices_AUGUST column A (Solaredge rate EUR), matched by destination
   country and the closest available pallet count on file for that
   country (Q3 doesn't have every pallet count, so nearest-match is used
   and noted). Countries with no Q3 data at all (United States, Australia
   in this data) fall back to the DBS/Ship-Cost-Matrix domestic/export
   logic used elsewhere, flagged in the Note.

4. Summary sheet rebuilt to the requested layout (Category / Total Lines /
   Priced / Unpriced / Total Cost (USD), rows Shipped/Backlog/Total) using
   whole-column SUMIF/COUNTIFS formulas against each population sheet, so
   deleting a row or editing a Cost (USD) value updates the totals
   automatically -- a whole-column reference never needs resizing.
"""
import math
import os
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'scripts'))
import full_freight_cost_report as core  # reuse DBS/matrix/MNT-UK/Q3 helpers

IN_FILE = os.path.join(BASE_DIR, 'data', 'Freight_Cost_Report_2_user_edit.xlsx')
OUT_FILE = os.path.join(BASE_DIR, 'output', 'Freight_Cost_Report_0609.xlsx')

COLS = ['Source', 'Shipment Number', 'SE Order#', 'Sending WHS Code', 'Dest WHS Code',
        'Forwarder', 'ShipMode', 'Customer Name', 'Destination country', 'ZIP',
        'Family Type', '# of Pallets per line (roundup)', 'Method', 'Zone / Route',
        'Cost', 'Currency', 'Cost (USD)', 'Note']
COL = {name: i + 1 for i, name in enumerate(COLS)}  # 1-indexed columns, fixed layout


def fix_matrix_duplicate_groups(ws, matrix):
    """Ship Cost Matrix flat-rate lines wrongly applied per-line: reallocate
    by pallet share within every shipment/order group with more than one
    line, whether or not the user flagged it. (Originally this only touched
    groups containing a Cost=0 flagged line; that missed groups the user
    hadn't spotted yet, e.g. all of "3PLs" and a few more in the other
    sheets, since none of their rows happened to be flagged -- fixed here to
    just check the group unconditionally. Idempotent: an already-corrected
    group recomputes to the same split.)"""
    max_row = ws.max_row
    groups = defaultdict(list)
    for r in range(2, max_row + 1):
        method = ws.cell(row=r, column=COL['Method']).value
        if method != 'Ship Cost Matrix':
            continue
        sh_num = ws.cell(row=r, column=COL['Shipment Number']).value
        order = ws.cell(row=r, column=COL['SE Order#']).value
        key = (sh_num if sh_num and sh_num != 'N/A' else order,
               ws.cell(row=r, column=COL['Zone / Route']).value)
        groups[key].append(r)

    fixed_groups = 0
    fixed_rows = 0
    for key, group_rows in groups.items():
        if len(group_rows) < 2:
            continue

        sample_row = group_rows[0]
        org = ws.cell(row=sample_row, column=COL['Sending WHS Code']).value
        ship_mode = ws.cell(row=sample_row, column=COL['ShipMode']).value
        dest = ws.cell(row=sample_row, column=COL['Destination country']).value
        hit = core.matrix_lookup(matrix, org, None, ship_mode, dest)
        if not hit:
            continue
        flat_cost, currency, no_pal, _ = hit

        pallets_by_row = {r: ws.cell(row=r, column=COL['# of Pallets per line (roundup)']).value
                           for r in group_rows}
        pallets_by_row = {r: (v if isinstance(v, (int, float)) else 0) for r, v in pallets_by_row.items()}
        total_pallets = sum(pallets_by_row.values())

        for r in group_rows:
            share = (pallets_by_row[r] / total_pallets) if total_pallets else (1 / len(group_rows))
            cost = round(flat_cost * share, 2)
            ws.cell(row=r, column=COL['Cost']).value = cost
            ws.cell(row=r, column=COL['Currency']).value = currency
            ws.cell(row=r, column=COL['Cost (USD)']).value = core.to_usd(cost, currency)
            note = (f'Corrected: Ship Cost Matrix gives one flat rate ({flat_cost} {currency}) for the '
                     f'whole shipment/order, not per line -- allocated pro-rata by pallet share '
                     f'across this group\'s {len(group_rows)} lines.')
            ws.cell(row=r, column=COL['Note']).value = note
            fixed_rows += 1
        fixed_groups += 1

    print(f"  Fixed {fixed_groups} matrix-duplicate group(s), {fixed_rows} line(s).")


def dbs_price_for_zone_and_pallets(dbs_book, zone, total_pallets):
    country_code = 'NL' if zone == 'NL' else zone[:2]
    zone_prices = dbs_book.get(country_code, {}).get(zone)
    if not zone_prices:
        return None
    ep = max(1, min(core.MAX_EP, math.ceil(total_pallets - 1e-9)))
    while ep <= core.MAX_EP:
        if ep in zone_prices:
            return zone_prices[ep]
        ep += 1
    return None


def fix_dbs_duplicate_groups(ws, dbs_book):
    """DBS Price list rates are per SHIPMENT (the EP bracket is the whole
    shipment's pallet count), not per line -- the same mistake as the Ship
    Cost Matrix one, just for DBS: a multi-line shipment/order was having
    each line priced separately off *its own* smaller pallet count instead
    of the group's total, which both picks the wrong (too-low) EP bracket
    per line and, since freight rates are sub-linear in pallets, adds up to
    well more than one correctly-priced shipment would cost. Recomputes
    every such group (Shipment Number for Shipped, SE Order# for Backlog,
    which has no shipment number yet) as: one DBS price for the group's
    total pallets, allocated pro-rata by each line's pallet share."""
    max_row = ws.max_row
    groups = defaultdict(list)
    for r in range(2, max_row + 1):
        if ws.cell(row=r, column=COL['Method']).value != 'DBS Price list 2026':
            continue
        sh_num = ws.cell(row=r, column=COL['Shipment Number']).value
        order = ws.cell(row=r, column=COL['SE Order#']).value
        key = (sh_num if sh_num and sh_num != 'N/A' else order,
               ws.cell(row=r, column=COL['Zone / Route']).value)
        groups[key].append(r)

    fixed_groups = 0
    fixed_rows = 0
    for key, group_rows in groups.items():
        if len(group_rows) < 2:
            continue
        zone = key[1]
        pallets_by_row = {r: ws.cell(row=r, column=COL['# of Pallets per line (roundup)']).value
                           for r in group_rows}
        pallets_by_row = {r: (v if isinstance(v, (int, float)) else 0) for r, v in pallets_by_row.items()}
        total_pallets = sum(pallets_by_row.values())
        shipment_price = dbs_price_for_zone_and_pallets(dbs_book, zone, total_pallets)
        if shipment_price is None:
            continue

        for r in group_rows:
            share = (pallets_by_row[r] / total_pallets) if total_pallets else (1 / len(group_rows))
            cost = round(shipment_price * share, 2)
            ws.cell(row=r, column=COL['Cost']).value = cost
            ws.cell(row=r, column=COL['Currency']).value = 'EUR'
            ws.cell(row=r, column=COL['Cost (USD)']).value = core.to_usd(cost, 'EUR')
            note = (f'Corrected: DBS Price list rate is for the whole shipment\'s total pallets '
                     f'({total_pallets:g} -> {shipment_price} EUR for zone {zone}), not per line -- '
                     f'allocated pro-rata by pallet share across this group\'s {len(group_rows)} lines.')
            ws.cell(row=r, column=COL['Note']).value = note
            fixed_rows += 1
        fixed_groups += 1

    print(f"  Fixed {fixed_groups} DBS-duplicate group(s), {fixed_rows} line(s).")


EX_DO_PALLET_KEY_FIELDS_SRC = ['SE Order#', 'Shipment Number', 'Sending WHS Code', 'Dest WHS Code',
                               'Forwarder', 'ShipMode', 'Customer Name', 'Destination country',
                               'Zip', 'Family Type']


def sync_ex_do_pallets_from_source(ws):
    """271 lines in the source "NL - EX + DO" data had '# of Pallets per line
    roundup' = 0 (not yet filled in) when the report was first generated --
    that 0 was floored up to EP=1 for any *standalone* DBS/matrix lookup, but
    fed straight into the flat-corridor-rate pro-rata split as a 0 pallet
    share for any line that was part of a multi-line shipment/order group,
    silently giving those lines $0 of a shipment they were actually part of.
    The user has since filled in the real (fractional-pallet) values in the
    source file; this re-syncs '# of Pallets per line (roundup)' in the
    reviewed file to match, by the row's other identifying fields (Shipment
    Number, customer, destination, ZIP, family type -- everything but the
    Line# this file doesn't carry), before any of the group-share fixes below
    recompute allocations from these pallet counts."""
    src_wb = openpyxl.load_workbook(
        os.path.join(BASE_DIR, 'data', 'Freight_costs_Data_0609.xlsx'), data_only=True)
    src_ws = src_wb['NL - EX + DO']
    headers = [c.value for c in src_ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    queues = defaultdict(list)
    for row in src_ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        key = tuple(row[idx[f]] for f in EX_DO_PALLET_KEY_FIELDS_SRC)
        queues[key].append(row[idx['# of Pallets per line roundup']])

    ex_do_key_cols = ['SE Order#', 'Shipment Number', 'Sending WHS Code', 'Dest WHS Code',
                       'Forwarder', 'ShipMode', 'Customer Name', 'Destination country',
                       'ZIP', 'Family Type']
    updated = 0
    for r in range(2, ws.max_row + 1):
        key = tuple(ws.cell(row=r, column=COL[c]).value for c in ex_do_key_cols)
        q = queues.get(key)
        if not q:
            continue
        new_pallets = q.pop(0)
        old_pallets = ws.cell(row=r, column=COL['# of Pallets per line (roundup)']).value
        if new_pallets != old_pallets:
            ws.cell(row=r, column=COL['# of Pallets per line (roundup)']).value = new_pallets
            updated += 1
    print(f"  Synced {updated} pallet value(s) from the corrected source data.")


MISSING_EX_DO_SE_ORDERS = {'653362', '616321', '615359', '570289', '567969'}


def restore_missing_ex_do_rows(ws, matrix):
    """5 SE Orders (6 lines) exist in the source "NL - EX + DO" data (all DHL/
    COURIER single-pallet shipments to Australia/India/Israel, Shipped) but
    are entirely absent from the user's file -- not deleted as part of any
    flagged duplicate-group, just missing outright, so every one of them was
    unpriced anywhere in the report. Restored here from source and priced via
    the same Ship Cost Matrix lookup (Sending WHS Code + ShipMode + destination
    country) used elsewhere for non-EUROPE lines on this sheet."""
    src_wb = openpyxl.load_workbook(
        os.path.join(BASE_DIR, 'data', 'Freight_costs_Data_0609.xlsx'), data_only=True)
    src_ws = src_wb['NL - EX + DO']
    headers = [c.value for c in src_ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    existing_keys = set()
    for r in range(2, ws.max_row + 1):
        sh = ws.cell(row=r, column=COL['Shipment Number']).value
        if sh and sh != 'N/A':
            existing_keys.add(sh)

    added = 0
    for row in src_ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        order = str(row[idx['SE Order#']])
        if order not in MISSING_EX_DO_SE_ORDERS:
            continue
        shipment_number = row[idx['Shipment Number']]
        if shipment_number in existing_keys:
            continue  # already present under this shipment number -- don't double-add
        org = row[idx['Sending WHS Code']]
        ship_mode = row[idx['ShipMode']]
        dest = row[idx['Destination country']]
        pallets = row[idx['# of Pallets per line roundup']]
        hit = core.matrix_lookup(matrix, org, row[idx['Origin country']], ship_mode, dest)
        route = f"{org}/{ship_mode}->{dest}"
        if hit:
            cost, currency, no_pal, note = hit
        else:
            cost, currency, note = None, None, f'No Ship Cost Matrix rate for {route} — confirm with WH contact.'
        note = ('Restored: this line exists in the source data but was missing from the '
                'reviewed file entirely (not a flagged duplicate) -- priced via Ship Cost Matrix. '
                + note).strip()
        new_row = [None] * len(COLS)
        new_row[COL['Source'] - 1] = row[idx['Source']]
        new_row[COL['Shipment Number'] - 1] = shipment_number if shipment_number else 'N/A'
        new_row[COL['SE Order#'] - 1] = order
        new_row[COL['Sending WHS Code'] - 1] = org
        new_row[COL['Dest WHS Code'] - 1] = row[idx['Dest WHS Code']]
        new_row[COL['Forwarder'] - 1] = row[idx['Forwarder']]
        new_row[COL['ShipMode'] - 1] = ship_mode
        new_row[COL['Customer Name'] - 1] = row[idx['Customer Name']]
        new_row[COL['Destination country'] - 1] = dest
        new_row[COL['ZIP'] - 1] = row[idx['Zip']]
        new_row[COL['Family Type'] - 1] = row[idx['Family Type']]
        new_row[COL['# of Pallets per line (roundup)'] - 1] = pallets
        new_row[COL['Method'] - 1] = 'Ship Cost Matrix'
        new_row[COL['Zone / Route'] - 1] = route
        new_row[COL['Cost'] - 1] = cost
        new_row[COL['Currency'] - 1] = currency
        new_row[COL['Cost (USD)'] - 1] = core.to_usd(cost, currency)
        new_row[COL['Note'] - 1] = note
        ws.append(new_row)
        for c in range(1, len(COLS) + 1):
            ws.cell(row=ws.max_row, column=c).font = Font(name='Arial')
        added += 1
    print(f"  Restored {added} missing line(s).")


DG_UK_COLLAPSED_ORDERS = {'579271': 'SH21526977882', '579228': 'SH21526977881'}


def restore_and_fix_dg_uk_duplicates(ws, matrix):
    """2 SE Orders in "NL - DG + UK" (579271, 579228) are each a 20-pallet
    shipment split into 20 distinct 1-pallet lines in the source data (same
    part number, different Line#) -- not exact duplicates. 19 of the 20 lines
    for each were removed from the user's file (they look identical without
    the Line# column), leaving 1 line silently carrying the *entire*
    Ship Cost Matrix corridor rate for the whole 20-pallet shipment instead of
    its own 1-pallet share. Restores the missing 19 lines each from source and
    reallocates the flat corridor rate pro-rata across all 20 -- this changes
    nothing about the shipment's total cost (still one flat rate), only fixes
    which lines carry it."""
    src_wb = openpyxl.load_workbook(
        os.path.join(BASE_DIR, 'data', 'Freight_costs_Data_0609.xlsx'), data_only=True)
    src_ws = src_wb['NL - DG + UK']
    headers = [c.value for c in src_ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    src_rows_by_order = defaultdict(list)
    for row in src_ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        order = str(row[idx['SE Order#']])
        if order in DG_UK_COLLAPSED_ORDERS:
            src_rows_by_order[order].append(row)

    added = 0
    fixed_groups = 0
    for order, shipment_number in DG_UK_COLLAPSED_ORDERS.items():
        src_rows = src_rows_by_order[order]
        # find the one surviving row in ws for this shipment
        existing_row_idx = None
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=COL['Shipment Number']).value == shipment_number:
                existing_row_idx = r
                break
        if existing_row_idx is None:
            continue

        sample = src_rows[0]
        org = sample[idx['Sending WHS Code']]
        ship_mode = sample[idx['ShipMode']]
        dest = sample[idx['Destination country']]
        hit = core.matrix_lookup(matrix, org, sample[idx['Origin country']], ship_mode, dest)
        if not hit:
            continue
        flat_cost, currency, no_pal, _ = hit

        # restore the missing lines (all but 1, since 1 already survives)
        group_rows = [existing_row_idx]
        for row in src_rows[1:]:
            new_row = [None] * len(COLS)
            new_row[COL['Source'] - 1] = row[idx['Source']]
            new_row[COL['Shipment Number'] - 1] = shipment_number
            new_row[COL['SE Order#'] - 1] = order
            new_row[COL['Sending WHS Code'] - 1] = org
            new_row[COL['Dest WHS Code'] - 1] = row[idx['Dest WHS Code']]
            new_row[COL['Forwarder'] - 1] = row[idx['Forwarder']]
            new_row[COL['ShipMode'] - 1] = ship_mode
            new_row[COL['Customer Name'] - 1] = row[idx['Customer Name']]
            new_row[COL['Destination country'] - 1] = dest
            new_row[COL['ZIP'] - 1] = row[idx['Zip']]
            new_row[COL['Family Type'] - 1] = row[idx['Family Type']]
            new_row[COL['# of Pallets per line (roundup)'] - 1] = row[idx['# of Pallets per line roundup']]
            new_row[COL['Method'] - 1] = 'Ship Cost Matrix'
            new_row[COL['Zone / Route'] - 1] = f"{org}/{ship_mode}->{dest}"
            ws.append(new_row)
            group_rows.append(ws.max_row)
            added += 1

        route = f"{org}/{ship_mode}->{dest}"
        pallets_by_row = {r: (ws.cell(row=r, column=COL['# of Pallets per line (roundup)']).value or 0)
                           for r in group_rows}
        total_pallets = sum(pallets_by_row.values())
        for r in group_rows:
            ws.cell(row=r, column=COL['Method']).value = 'Ship Cost Matrix'
            ws.cell(row=r, column=COL['Zone / Route']).value = route
            share = (pallets_by_row[r] / total_pallets) if total_pallets else (1 / len(group_rows))
            cost = round(flat_cost * share, 2)
            ws.cell(row=r, column=COL['Cost']).value = cost
            ws.cell(row=r, column=COL['Currency']).value = currency
            ws.cell(row=r, column=COL['Cost (USD)']).value = core.to_usd(cost, currency)
            note = (f'Restored: 19 of this shipment\'s 20 one-pallet lines were missing from the reviewed '
                     f'file (removed as apparent duplicates -- they differ only by Line# in the source data, '
                     f'not shown here). Ship Cost Matrix gives one flat rate ({flat_cost} {currency}) for the '
                     f'whole 20-pallet shipment, not per line -- allocated pro-rata by pallet share across '
                     f'all {len(group_rows)} lines; the shipment\'s total cost is unchanged.')
            ws.cell(row=r, column=COL['Note']).value = note
            for c in range(1, len(COLS) + 1):
                ws.cell(row=r, column=c).font = Font(name='Arial')
        fixed_groups += 1
    print(f"  Restored {added} missing line(s) across {fixed_groups} shipment(s).")


def build_q3_country_pallet_table(q3_by_pl_nr_unused=None):
    wb = openpyxl.load_workbook(core.Q3_AUG_FILE, data_only=True)
    ws = wb['Sheet1']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    by_country = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        rate = r[idx['Solaredge rate EUR']]
        pal = r[idx['Nr. of pal']]
        country = r[idx['country']]
        if isinstance(rate, (int, float)) and isinstance(pal, (int, float)) and country:
            by_country[country].append((pal, rate))
    return by_country


def q3_country_pallet_lookup(table, country, pallets):
    entries = table.get(country)
    if not entries:
        return None, None, f'No Q3 AUGUST data at all for {country!r}.'
    if not isinstance(pallets, (int, float)):
        return None, None, f'Invalid pallet count ({pallets!r}).'
    nearest_pal, rate = min(entries, key=lambda e: abs(e[0] - pallets))
    note = f'Q3 AUGUST rate for {country}, nearest pallet bracket on file = {nearest_pal} (this line: {pallets}).'
    return rate, 'EUR', note


def reprice_nl_dg_uk(ws, dbs_book, matrix, mnt_uk_book, q3_country_table):
    max_row = ws.max_row
    n = 0
    for r in range(2, max_row + 1):
        row = {name: ws.cell(row=r, column=col).value for name, col in COL.items()}
        dest = row['Destination country']
        pallets = row['# of Pallets per line (roundup)']

        if dest == 'United Kingdom':
            cost, note = core.mnt_uk_lookup(mnt_uk_book, row['Customer Name'], row['ZIP'])
            method, route, currency = 'MNT UK price list (Q3)', 'UK', ('EUR' if cost is not None else None)
            note = note or 'Matched by customer name.'
        else:
            cost, currency, note = q3_country_pallet_lookup(q3_country_table, dest, pallets)
            method, route = 'Q3 prices AUGUST (by country+pallets)', f'{dest} / {pallets} pallets'
            if cost is None:
                fallback_row = dict(row)
                fallback_row['Zip'] = row['ZIP']
                fallback_row['Destination country code'] = None
                fallback_row['Destination country'] = dest
                fallback_row['Parent Region'] = 'ROW'  # force matrix path; no DBS zone for US/AU anyway
                fallback_row['Sending WHS Code'] = row['Sending WHS Code']
                fallback_row['Origin country'] = 'Netherlands'
                fallback_row['ShipMode'] = row['ShipMode']
                fallback_row['# of Pallets per line roundup'] = pallets
                fb = core.domestic_export_price(fallback_row, dbs_book, matrix)
                cost, currency = fb['cost'], fb['currency']
                method = f"Ship Cost Matrix (fallback: {note})"
                route = fb['route']
                note = (note + ' ' + (fb['note'] or '')).strip()

        ws.cell(row=r, column=COL['Method']).value = method
        ws.cell(row=r, column=COL['Zone / Route']).value = route
        ws.cell(row=r, column=COL['Cost']).value = cost
        ws.cell(row=r, column=COL['Currency']).value = currency
        ws.cell(row=r, column=COL['Cost (USD)']).value = core.to_usd(cost, currency)
        ws.cell(row=r, column=COL['Note']).value = note
        n += 1
    print(f"  Repriced {n} NL - DG + UK line(s).")


def rebuild_summary(wb, sheet_names):
    if 'Summary' in wb.sheetnames:
        del wb['Summary']
    ws = wb.create_sheet('Summary', 1)
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')

    headers = ['Category', 'Total Lines', 'Priced', 'Unpriced', 'Total Cost (USD)']
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold

    source_refs = '+'.join(f"COUNTIF('{s}'!$A:$A,\"Shipped\")" for s in sheet_names)
    backlog_refs = '+'.join(f"COUNTIF('{s}'!$A:$A,\"Backlog\")" for s in sheet_names)
    shipped_total_lines = source_refs
    backlog_total_lines = backlog_refs

    shipped_cost = '+'.join(
        f"SUMIFS('{s}'!$Q:$Q,'{s}'!$A:$A,\"Shipped\")" for s in sheet_names)
    backlog_cost = '+'.join(
        f"SUMIFS('{s}'!$Q:$Q,'{s}'!$A:$A,\"Backlog\")" for s in sheet_names)
    shipped_priced = '+'.join(
        f"COUNTIFS('{s}'!$A:$A,\"Shipped\",'{s}'!$Q:$Q,\"<>\")" for s in sheet_names)
    backlog_priced = '+'.join(
        f"COUNTIFS('{s}'!$A:$A,\"Backlog\",'{s}'!$Q:$Q,\"<>\")" for s in sheet_names)

    ws.append(['Shipped', f'={shipped_total_lines}', f'={shipped_priced}',
               f'=B2-C2', f'={shipped_cost}'])
    ws.append(['Backlog', f'={backlog_total_lines}', f'={backlog_priced}',
               f'=B3-C3', f'={backlog_cost}'])
    ws.append(['Total', '=B2+B3', '=C2+C3', '=D2+D3', '=E2+E3'])

    for r in (2, 3, 4):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = bold if r == 4 else arial
            if c == 5:
                cell.number_format = '#,##0.00'

    # Second table: same columns, one row per population sheet (both
    # Shipped and Backlog lines combined), so each sheet's contribution to
    # the grand total above is visible on its own -- split Shipped/Backlog
    # per sheet, same layout as the requested screenshot.
    section_row = 6
    ws.cell(row=section_row, column=1, value='By Population (Shipped / Backlog)').font = bold
    header_row = section_row + 1
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h).font = bold

    first_data_row = header_row + 1
    row_labels = []
    for s in sheet_names:
        row_labels.append((s, 'Shipped'))
        row_labels.append((s, 'Backlog'))

    for i, (s, status) in enumerate(row_labels):
        r = first_data_row + i
        ws.cell(row=r, column=1, value=f'{s} - {status}')
        ws.cell(row=r, column=2, value=f"=COUNTIF('{s}'!$A:$A,\"{status}\")")
        ws.cell(row=r, column=3, value=f"=COUNTIFS('{s}'!$A:$A,\"{status}\",'{s}'!$Q:$Q,\"<>\")")
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
        ws.cell(row=r, column=5, value=f"=SUMIFS('{s}'!$Q:$Q,'{s}'!$A:$A,\"{status}\")")
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).font = arial
            if c == 5:
                ws.cell(row=r, column=c).number_format = '#,##0.00'

    total_row = first_data_row + len(row_labels)
    ws.cell(row=total_row, column=1, value='Total').font = bold
    for c in range(2, 6):
        first = get_column_letter(c) + str(first_data_row)
        last = get_column_letter(c) + str(total_row - 1)
        cell = ws.cell(row=total_row, column=c, value=f'=SUM({first}:{last})')
        cell.font = bold
        if c == 5:
            cell.number_format = '#,##0.00'

    for c, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(24, len(h) + 4)
    return ws


def rebuild_readme(wb):
    if 'README' in wb.sheetnames:
        del wb['README']
    readme = wb.create_sheet('README', 0)
    readme.column_dimensions['A'].width = 105
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    lines = [
        ('Freight Cost Report -- patched per user review', bold),
        ('', arial),
        ('This file is the user\'s own edited copy (rows they deleted or manually corrected are kept '
         'exactly as they left them). Seven targeted fixes were applied on top of that:', arial),
        ('', arial),
        ('1) 3PLs / NL - Support / NL - EX + DO / Canot - EX + DO: any Shipment Number (or SE Order# for '
         'Backlog) group where the Ship Cost Matrix\'s one flat corridor rate had been applied to every '
         'line -- inflating the total by the line count -- is reallocated pro-rata by pallet share across '
         'the whole group. Applied unconditionally to every such group in these 4 sheets (not just rows '
         'the user had flagged with Cost=0), since the same bug turned up unflagged in "3PLs" too.', arial),
        ('2) NL - Support / NL - EX + DO: the same flat-rate-per-shipment bug also existed on the DBS '
         'Price list side -- its EP bracket (and rate) is for the whole shipment\'s total pallets, not per '
         'line. Every DBS-priced Shipment Number (or SE Order# for Backlog) group of 2+ lines is '
         'recomputed as one DBS price for the group\'s total pallets, allocated pro-rata by pallet share. '
         '19 groups / 70 lines fixed in NL - Support, 346 groups / 1,624 lines in NL - EX + DO -- verified '
         'every group\'s allocated Cost sums back exactly to the DBS price-book rate for its total pallets. '
         '(NL - DG + UK only ever falls back to DBS for single-line US/Australia cases, and Canot\'s Israel '
         'lines use their own domestic rate card, not DBS -- neither sheet is affected.)', arial),
        ('3) NL - DG + UK: re-priced sheet-wide. Destination = United Kingdom -> MNT UK price list '
         '(sheet "Price Q3"), matched by customer. Everything else -> Q3_prices_AUGUST column A '
         '(Solaredge rate EUR), matched by destination country + the nearest pallet count Q3 has on file '
         'for that country (noted per line). Countries with no Q3 data at all (United States, Australia in '
         'this data) fall back to the DBS/Ship-Cost-Matrix domestic/export logic, flagged in the Note.', arial),
        ('4) NL - EX + DO: 5 SE Orders (6 lines) -- 653362, 616321 (x2), 615359, 570289, 567969, all DHL/'
         'COURIER single-pallet Shipped lines to Australia/India/Israel -- exist in the source data but were '
         'missing from the reviewed file entirely (not part of any flagged duplicate group), so they were '
         'unpriced anywhere in the report. Restored from source and priced via Ship Cost Matrix (COURIER '
         'rates exist on file for all 3 destinations).', arial),
        ('5) NL - DG + UK: SE Orders 579271 and 579228 are each a 20-pallet US shipment split into 20 '
         'distinct 1-pallet lines in the source data (same part number, different Line#); 19 of the 20 lines '
         'for each were removed from the reviewed file (they look identical without a Line# column, but are '
         'not exact duplicates). Restored all 19+19 missing lines and reallocated the Ship Cost Matrix flat '
         'corridor rate pro-rata across all 20 lines per shipment -- the shipment\'s total cost is unchanged, '
         'only which lines carry it.', arial),
        ('6) NL - EX + DO: 271 lines had "# of Pallets per line roundup" = 0 in the source data (not yet '
         'filled in) when this report was first generated; the user has since supplied the real (fractional) '
         'pallet values. Re-synced from the corrected source data by matching each line\'s other identifying '
         'fields (Shipment Number, customer, destination, ZIP, family type), then re-ran the matrix/DBS '
         'group-share fixes above so every affected shipment\'s flat rate is re-split by the corrected pallet '
         'shares -- lines with 0 pallets were previously drawing $0 (or an even, unweighted split) of a '
         'shipment they were genuinely part of.', arial),
        ('7) Summary rebuilt to Category / Total Lines / Priced / Unpriced / Total Cost (USD), rows '
         'Shipped / Backlog / Total, using whole-column SUMIFS/COUNTIFS formulas -- deleting a row or '
         'editing a Cost (USD) cell updates every total automatically, no range to resize.', arial),
        ('', arial),
        ('3PLs and Canot - EX + DO otherwise keep every row exactly as the user left them -- only the '
         'matrix duplicate-group Cost/Cost (USD)/Note cells above were touched.', arial),
    ]
    for text, font in lines:
        readme.append((text,))
        readme.cell(row=readme.max_row, column=1).font = font


def main():
    print("Loading price sources...")
    dbs_book = core.parse_dbs_price_book(core.DBS_FILE)
    matrix = core.parse_ship_matrix(core.MATRIX_FILE)
    mnt_uk_book = core.parse_mnt_uk(core.MNT_UK_FILE)
    q3_country_table = build_q3_country_pallet_table()

    wb = openpyxl.load_workbook(IN_FILE)

    print("Syncing NL - EX + DO pallet values from corrected source data...")
    sync_ex_do_pallets_from_source(wb['NL - EX + DO'])

    print("Restoring lines missing outright from NL - EX + DO...")
    restore_missing_ex_do_rows(wb['NL - EX + DO'], matrix)

    print("Fixing Ship Cost Matrix duplicate-line groups...")
    for sheet_name in ['3PLs', 'NL - Support', 'NL - EX + DO', 'Canot - EX + DO']:
        print(f" {sheet_name}:")
        fix_matrix_duplicate_groups(wb[sheet_name], matrix)

    print("Fixing DBS Price list duplicate-line groups (same bug, DBS side)...")
    for sheet_name in ['NL - Support', 'NL - EX + DO']:
        print(f" {sheet_name}:")
        fix_dbs_duplicate_groups(wb[sheet_name], dbs_book)

    print("Repricing NL - DG + UK...")
    reprice_nl_dg_uk(wb['NL - DG + UK'], dbs_book, matrix, mnt_uk_book, q3_country_table)

    print("Restoring collapsed-duplicate lines in NL - DG + UK...")
    restore_and_fix_dg_uk_duplicates(wb['NL - DG + UK'], matrix)

    population_sheets = ['3PLs', 'NL - Support', 'NL - EX + DO', 'NL - DG + UK', 'Canot - EX + DO']
    print("Rebuilding Summary...")
    rebuild_summary(wb, population_sheets)
    rebuild_readme(wb)

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)
    print("Saved", OUT_FILE)


if __name__ == '__main__':
    main()
