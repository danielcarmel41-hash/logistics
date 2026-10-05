"""
Prices the new Q4 backlog pull (data/Backlog_Q4_Data.xlsx, 516 source lines,
single flat "Backlog Q4" sheet, same 94-column layout as data/BI_2809.xlsx)
and merges it into the user's current, hand-reviewed report
(data/BI_2809_Freight_1.1_v2.xlsx), touching ONLY the "Backlog" sheet --
every other sheet (README, Shipped, the user's own "Sum"/"Summary" tabs) is
carried over byte-for-byte untouched, per the explicit request.

Classification and pricing reuse the exact same rules as the main pipeline
(bi_2809_freight_cost_report.classify/price_all): Canot/DBS/Ship Cost
Matrix/BayWa IT overrides unchanged; MNT/Q3-sourced groups are priced from
the September USD price list (data/SEPTEMBER_USD.xlsx, column H) by exact
Shipment Number match first, falling back to the relaxed nearest-pallet/
same-country/same-customer estimate -- same as the last patch round, except
the sample pool is now drawn from the CURRENT REPORT's own confirmed (non-
estimated) MNT prices (across both its Shipped and Backlog sheets), so the
most recently confirmed quotes are used. Every row/group gets its Population
label prefixed "Q4- " (e.g. "Q4- NL - Domestic") to mark it as coming from
this Q4 backlog pull, per the request.

Duplicate-order check (the user explicitly asked to verify this): of the
163 distinct SE Order# in the Q4 pull, 13 already have a consolidated row in
the CURRENT Backlog sheet. Line-by-line comparison against the original
source pull (data/BI_2809.xlsx) shows this Q4 file is a refreshed snapshot
of those same orders' remaining backlog (pallet quantities shifted, some
lines already shipped and dropped out, in one case new lines appeared) --
not additional new lines on top. Per the user's explicit choice, those 13
existing Backlog rows are replaced by the freshly priced Q4 version (old
row removed, new one added) rather than being added alongside the stale
one, which would double count. The other 150 orders are net-new to the
Backlog sheet and are simply added. (A further 80 SE Order# also appear in
the Shipped sheet, by order number only -- checked line-by-line and
confirmed to be entirely different physical lines of the same order, no
Line#/Part Number overlap at all, i.e. the already-shipped portion versus
the still-pending portion; these are safe to add as new Backlog rows with
no conflict.)
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
import bi_2809_freight_cost_report as core
import patch_bi_2809_september_usd as sept

DATA = lambda name: os.path.join(BASE_DIR, 'data', name)
Q4_FILE = DATA('Backlog_Q4_Data.xlsx')
IN_FILE = DATA('BI_2809_Freight_1.1_v2.xlsx')
OUT_FILE = os.path.join(BASE_DIR, 'output', 'BI_2809_Freight_Cost_Report.xlsx')

# SE Order# already present in the current Backlog sheet, confirmed (by the
# user) to be stale snapshots of these same 13 Q4 orders -- replace, don't add.
REPLACE_ORDERS = {
    '661973', '663075', '676985', '788906', '811554', '821226', '825855',
    '837802', '838221', '840406', '857526', '863800', '865204',
}


def load_q4_rows():
    wb = openpyxl.load_workbook(Q4_FILE, data_only=True)
    ws = wb['Backlog Q4']
    headers = [c.value for c in ws[1]]
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(v is None for v in vals):
            continue
        rows.append(dict(zip(headers, vals)))
    return rows


def build_mnt_sample_pool(in_wb):
    """Confirmed (non-estimated) MNT/Q3-sourced USD prices already in the
    current report's Shipped + Backlog sheets, keyed by destination country."""
    samples = defaultdict(list)
    mnt_methods = {'MNT UK / Q3 AUGUST', 'Q3 AUGUST (MNT, POD)'}
    for sheet_name in ('Shipped', 'Backlog'):
        ws = in_wb[sheet_name]
        headers = [c.value for c in ws[1]]
        idx = {h: i for i, h in enumerate(headers)}
        for r in ws.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in r):
                continue
            if r[idx['Method']] not in mnt_methods:
                continue
            cost_usd = r[idx['Cost (USD)']]
            note = r[idx['Note']] or ''
            if cost_usd is None or 'Estimated' in note:
                continue
            country = r[idx['Destination country']]
            customer = r[idx['Customer Name']]
            pallets = r[idx['# of Pallets (total)']] or 1
            samples[country].append((customer, pallets, cost_usd))
    return samples


def price_group(method, group_rows, dbs_book, matrix, sept_book, mnt_samples):
    sample = group_rows[0]
    total_pallets = sum((core.num_pallets(r) or 0) for r in group_rows)
    rec = {'rows': group_rows, 'method': method, 'route': None, 'cost': None, 'currency': None, 'note': ''}

    if method == 'dbs':
        zone, price, note = core.dbs_lookup(dbs_book, sample.get('Destination country code'),
                                             sample.get('Zip'), total_pallets or 1)
        rec['route'] = zone
        if price is not None:
            rec['cost'] = price
            rec['currency'] = 'EUR'
            rec['note'] = note or f'DBS Price list 2026, zone {zone}: {price} EUR for {total_pallets:g} total pallets.'
        else:
            est_price, est_note = core.dbs_country_estimate(dbs_book, sample.get('Destination country code'),
                                                              total_pallets or 1)
            if est_price is not None:
                rec['cost'] = est_price
                rec['currency'] = 'EUR'
                rec['note'] = est_note
            else:
                rec['note'] = (note or '') + ' ' + est_note

    elif method in ('q3_mnt', 'battery_uk'):
        ship_num = sample.get('Shipment Number')
        sept_rate = sept_book.get(ship_num) if ship_num and ship_num != 'N/A' else None
        if sept_rate is not None:
            rec['route'] = f'September USD price list (PL nr. {ship_num})'
            rec['cost'] = sept_rate
            rec['currency'] = 'USD'
            rec['note'] = (f'Confirmed MNT rate {sept_rate} USD for the whole shipment, from the September USD '
                            f'price list (column H), matched by Shipment Number.')
        else:
            country = sample.get('Destination country')
            customer = sample.get('Customer Name')
            price, est_note = sept.q3_country_estimate(mnt_samples, country, customer, total_pallets or 1)
            if price is not None:
                rec['route'] = 'September USD price list (estimated)'
                rec['cost'] = price
                rec['currency'] = 'USD'
                rec['note'] = est_note
            else:
                rec['note'] = (f"MNT/Battery+UK-sourced group, no Shipment Number match -- {est_note}")

    elif method == 'matrix':
        org = sample.get('Sending WHS Code')
        hit = core.matrix_lookup(matrix, org, sample.get('Origin country'), sample.get('ShipMode'),
                                  sample.get('Destination country'))
        route = f"{org}/{sample.get('ShipMode')}->{sample.get('Destination country')}"
        rec['route'] = route
        if not hit:
            rec['note'] = f'No Ship Cost Matrix rate for {route} -- confirm with WH contact.'
        else:
            cost, currency, no_pal, mnote = hit
            rec['cost'] = cost
            rec['currency'] = currency
            note = f'Ship Cost Matrix flat rate {cost} {currency} for the whole shipment ({total_pallets:g} pallets).'
            if isinstance(no_pal, (int, float)) and total_pallets > no_pal:
                note += f' Shipment total exceeds this corridor\'s {no_pal:.0f}-pallet rate -- verify capacity.'
            if mnote:
                note += ' ' + mnote
            rec['note'] = note

    elif method == 'canot':
        price, currency, note = core.canot_israel_lookup(sample.get('City'), total_pallets or 1)
        rec['route'] = 'Canot domestic (IL)'
        rec['cost'] = price
        rec['currency'] = currency
        rec['note'] = note

    elif method == 'baywa_it':
        capped_total = max(1, math.ceil(total_pallets - 1e-9)) if total_pallets else 1
        price = core.BAYWA_IT_OVERRIDES.get(capped_total)
        rec['route'] = 'BayWa IT (manual)'
        rec['cost'] = price
        rec['currency'] = 'EUR' if price is not None else None
        if price is None:
            rec['note'] = f'No BayWa IT rate on file or interpolatable for {capped_total} total pallets.'
        elif capped_total in core.BAYWA_IT_INTERPOLATED:
            rec['note'] = (f'Estimated BayWa IT rate {price} EUR for {capped_total} total pallets -- interpolated '
                            f'linearly between the known 2-pallet (411 EUR) and 12-pallet (2100 EUR) rates.')
        else:
            rec['note'] = f'User-supplied BayWa IT rate {price} EUR for {capped_total} total pallets.'

    else:
        rec['note'] = f"Forwarder {sample.get('Forwarder')!r} is neither DBSCHENKER nor MNT -- price manually."

    return rec


def main():
    print("Loading price sources...")
    dbs_book = core.parse_dbs_price_book(core.DBS_FILE)
    matrix = core.parse_ship_matrix(core.MATRIX_FILE)
    sept_book = sept.parse_september_usd(sept.PRICE_FILE)

    print("Loading the Q4 backlog pull...")
    q4_rows = load_q4_rows()
    print(f"  {len(q4_rows)} source lines.")
    total_src_pallets = sum((core.num_pallets(r) or 0) for r in q4_rows)

    print("Loading the current report (base for the merge)...")
    in_wb_ro = openpyxl.load_workbook(IN_FILE, data_only=True)  # for sample-pool + reconciliation reads
    mnt_samples = build_mnt_sample_pool(in_wb_ro)

    print("Classifying and pricing the Q4 rows...")
    buckets = defaultdict(list)
    for row in q4_rows:
        population, method = core.classify(row)
        buckets[(method, core.group_key_for(row), population)].append(row)

    new_recs = []
    for (method, gkey, population), group_rows in buckets.items():
        rec = price_group(method, group_rows, dbs_book, matrix, sept_book, mnt_samples)
        rec['population'] = 'Q4- ' + population
        new_recs.append(rec)

    # ---- reconciliation: every source line accounted for exactly once ----
    covered_lines = sum(len(r['rows']) for r in new_recs)
    covered_pallets = sum(sum((core.num_pallets(x) or 0) for x in r['rows']) for r in new_recs)
    assert covered_lines == len(q4_rows), f"line count mismatch: {covered_lines} vs {len(q4_rows)}"
    assert covered_pallets == total_src_pallets, f"pallet sum mismatch: {covered_pallets} vs {total_src_pallets}"
    print(f"  Reconciled: {covered_lines} source lines / {covered_pallets} pallets -> "
          f"{len(new_recs)} consolidated groups (100% accounted for).")

    orders_in_recs = set()
    for r in new_recs:
        so = core._joined(r['rows'], 'SE Order#')
        orders_in_recs.add(so)
    replaced = orders_in_recs & REPLACE_ORDERS
    added = orders_in_recs - REPLACE_ORDERS
    print(f"  {len(replaced)} order(s) replace a stale existing Backlog row, {len(added)} are net-new.")
    missing_replace = REPLACE_ORDERS - orders_in_recs
    if missing_replace:
        print(f"  WARNING: expected to replace {missing_replace} but found no corresponding Q4 group.")

    print("Patching the Backlog sheet of the user's current report...")
    out_wb = openpyxl.load_workbook(IN_FILE)  # preserves styles/formulas; this is what gets saved
    ws = out_wb['Backlog']
    headers = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    # remove the 13 stale rows (match by exact SE Order# -- each is a single,
    # un-joined value in these backlog rows, confirmed above)
    rows_to_delete = []
    old_values_by_order = {}
    for r in range(2, ws.max_row + 1):
        so = ws.cell(row=r, column=col['SE Order#']).value
        if so is None:
            continue
        so = str(so)
        if so in REPLACE_ORDERS:
            old_values_by_order[so] = {
                'pallets': ws.cell(row=r, column=col['# of Pallets (total)']).value,
                'cost_usd': ws.cell(row=r, column=col['Cost (USD)']).value,
                'population': ws.cell(row=r, column=col['Population']).value,
            }
            rows_to_delete.append(r)
    for r in reversed(rows_to_delete):
        ws.delete_rows(r)
    print(f"  Removed {len(rows_to_delete)} stale row(s) for the replaced orders.")

    # append the new/replacement rows
    for rec in new_recs:
        values = [getter(rec) for _, getter in core.OUT_COLUMNS]
        ws.append(values)
        r = ws.max_row
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).font = Font(name='Arial')
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    out_wb.save(OUT_FILE)
    print("Saved", OUT_FILE)

    # ---- summary for the user ----
    total_new_usd = sum(core.to_usd(r['cost'], r['currency']) or 0 for r in new_recs if r['cost'] is not None)
    n_priced = sum(1 for r in new_recs if r['cost'] is not None)
    print(f"\nNew/updated Backlog rows: {len(new_recs)} ({n_priced} priced, {len(new_recs)-n_priced} unpriced)")
    print(f"Total cost of the Q4 batch: {total_new_usd:,.2f} USD")

    print("\nReplaced orders -- old vs new:")
    for so in sorted(replaced):
        old = old_values_by_order.get(so, {})
        new_rec = next(r for r in new_recs if core._joined(r['rows'], 'SE Order#') == so)
        new_pallets = sum((core.num_pallets(x) or 0) for x in new_rec['rows'])
        new_usd = core.to_usd(new_rec['cost'], new_rec['currency'])
        print(f"  {so}: pallets {old.get('pallets')} -> {new_pallets}, "
              f"cost(USD) {old.get('cost_usd')} -> {new_usd}, population {old.get('population')} -> {new_rec['population']}")


if __name__ == '__main__':
    main()
