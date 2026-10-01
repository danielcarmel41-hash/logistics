"""
Patches the user's latest hand-edited BI_2809_Freight_011026.xlsx in place
with an updated MNT price source (data/SEPTEMBER_USD.xlsx, "Solaredge rate
USD" column H, keyed by Shipment Number / "PL nr."), per two explicit
requests and nothing else:

1. VLOOKUP: every row whose pricing Method is MNT-sourced ("MNT UK / Q3
   AUGUST" or "Q3 AUGUST (MNT, POD)") and whose Shipment Number is found in
   the September USD price list gets its Cost/Currency/Cost (USD) replaced
   with that list's column H value (already in USD) -- this supersedes
   whatever was there before (blank, a prior exact match, or a prior
   relaxed estimate), since it is now a confirmed quote.

2. Review/estimate: every remaining MNT-sourced row whose Shipment Number
   is NOT in the September list gets a refreshed estimate -- nearest
   available pallet count for the same destination country, preferring
   the same customer name -- drawn from the pool of rows that DO now have
   a confirmed price (either a fresh VLOOKUP match from (1), or a
   pre-existing exact match that was never an estimate to begin with).
   This reuses the same estimate methodology as the prior round, just
   against a larger, fresher sample pool (the September file adds real
   coverage for Belgium, Romania and Finland that Q3 AUGUST never had).

Anything else is left completely untouched, per the explicit instruction:
non-MNT rows (DBS Price list 2026, Ship Cost Matrix, Canot rate card,
BayWa IT (manual)), the Manual Review sheet, and -- most importantly --
any row the user has zeroed out (Cost == 0) are never touched, whether or
not its Shipment Number happens to be in the September list.
"""
import os
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = lambda name: os.path.join(BASE_DIR, 'data', name)

IN_FILE = DATA('BI_2809_Freight_011026.xlsx')
PRICE_FILE = DATA('SEPTEMBER_USD.xlsx')
OUT_FILE = os.path.join(BASE_DIR, 'output', 'BI_2809_Freight_Cost_Report.xlsx')

MNT_METHODS = {'MNT UK / Q3 AUGUST', 'Q3 AUGUST (MNT, POD)'}


def parse_september_usd(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    book = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in r):
            continue
        pl = r[idx['PL nr.']]
        rate_usd = r[idx['Solaredge rate USD']]
        if isinstance(pl, str):
            # a handful of PL nr. values carry stray leading/trailing spaces
            # or non-breaking spaces (\xa0) that would otherwise silently
            # fail to match an exact Shipment Number.
            pl = pl.replace('\xa0', ' ').strip()
        if pl and isinstance(rate_usd, (int, float)):
            book[pl] = round(rate_usd, 2)
    return book


def q3_country_estimate(samples, country, customer, total_pallets):
    pool = samples.get(country)
    if not pool:
        return None, f"No confirmed MNT price on file for destination country {country!r} -- cannot estimate."
    same_cust = [s for s in pool if s[0] == customer]
    use = same_cust if same_cust else pool
    best_customer, best_pallets, best_rate = min(use, key=lambda s: abs(s[1] - (total_pallets or 1)))
    cust_note = f"same customer {customer!r}" if same_cust else "no other shipment for this customer on file"
    note = (f"Estimated MNT rate {best_rate} USD for {total_pallets:g} total pallets -- nearest available "
            f"pallet count ({best_pallets:g}) on file for destination country {country!r} ({cust_note}); no "
            f"exact Shipment Number match in the September USD price list, confirm with WH.")
    return best_rate, note


def main():
    print("Loading the September USD MNT price list...")
    price_book = parse_september_usd(PRICE_FILE)
    print(f"  {len(price_book)} shipment numbers on file.")

    print("Loading the user's latest edited report...")
    wb = openpyxl.load_workbook(IN_FILE)

    # ---- pass 1: collect every row + VLOOKUP-update matched ones --------
    all_mnt_rows = []  # (sheet_name, row_idx, col, get, set)
    for sheet_name in ('Shipped', 'Backlog'):
        ws = wb[sheet_name]
        headers = [c.value for c in ws[1]]
        col = {h: i + 1 for i, h in enumerate(headers)}
        for r in range(2, ws.max_row + 1):
            get = lambda h, r=r: ws.cell(row=r, column=col[h]).value
            if all(ws.cell(row=r, column=c).value is None for c in range(1, len(headers) + 1)):
                continue
            if get('Method') not in MNT_METHODS:
                continue
            all_mnt_rows.append((sheet_name, r, ws, col, get))

    n_vlookup = 0
    n_zero_skipped = 0
    for sheet_name, r, ws, col, get in all_mnt_rows:
        cost = get('Cost')
        if cost == 0:
            n_zero_skipped += 1
            continue
        sh = get('Shipment Number')
        if sh in price_book:
            price = price_book[sh]
            ws.cell(row=r, column=col['Cost']).value = price
            ws.cell(row=r, column=col['Currency']).value = 'USD'
            ws.cell(row=r, column=col['Cost (USD)']).value = price
            ws.cell(row=r, column=col['Zone / Route']).value = f'September USD price list (PL nr. {sh})'
            ws.cell(row=r, column=col['Note']).value = (
                f'Confirmed MNT rate {price} USD for the whole shipment, from the September USD price list '
                f'(column H), matched by Shipment Number.')
            n_vlookup += 1

    print(f"Request 1 (VLOOKUP by Shipment Number): updated {n_vlookup} row(s); "
          f"left {n_zero_skipped} user-zeroed row(s) untouched.")

    # ---- pass 2: build the sample pool from confirmed (non-estimated) prices
    samples = defaultdict(list)
    for sheet_name, r, ws, col, get in all_mnt_rows:
        cost = get('Cost')
        note = get('Note') or ''
        if cost in (None, 0) or 'Estimated' in note:
            continue
        country = get('Destination country')
        customer = get('Customer Name')
        pallets = get('# of Pallets (total)') or 1
        samples[country].append((customer, pallets, cost))

    # ---- pass 3: refresh the estimate for every still-unmatched row -----
    n_estimated = 0
    n_still_unresolved = 0
    for sheet_name, r, ws, col, get in all_mnt_rows:
        cost = get('Cost')
        note = get('Note') or ''
        if cost == 0:
            continue
        sh = get('Shipment Number')
        if sh in price_book:
            continue  # already handled by the VLOOKUP update above
        if cost is not None and 'Estimated' not in note:
            continue  # a genuine, non-estimated, non-matched price -- leave it alone
        country = get('Destination country')
        customer = get('Customer Name')
        pallets = get('# of Pallets (total)') or 1
        price, est_note = q3_country_estimate(samples, country, customer, pallets)
        if price is None:
            n_still_unresolved += 1
            continue
        ws.cell(row=r, column=col['Cost']).value = price
        ws.cell(row=r, column=col['Currency']).value = 'USD'
        ws.cell(row=r, column=col['Cost (USD)']).value = price
        ws.cell(row=r, column=col['Zone / Route']).value = 'September USD price list (estimated)'
        ws.cell(row=r, column=col['Note']).value = est_note
        n_estimated += 1

    print(f"Request 2 (country/customer/pallet estimate): (re-)estimated {n_estimated} row(s); "
          f"{n_still_unresolved} row(s) still have no confirmed or estimable MNT price on file.")

    rebuild_summary(wb)
    update_readme(wb, n_vlookup, n_estimated, n_still_unresolved)

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)
    print("Saved", OUT_FILE)


def rebuild_summary(wb):
    if 'Summary' in wb.sheetnames:
        del wb['Summary']
    ws = wb.create_sheet('Summary', 1)
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    headers = ['Population', 'Status', 'Lines (consolidated rows)', 'Priced', 'Unpriced', 'Total Cost (USD)']
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold

    by_pop_status = defaultdict(lambda: [0, 0, 0.0])
    for sheet_name in ('Shipped', 'Backlog'):
        ws_data = wb[sheet_name]
        idx = {c.value: i for i, c in enumerate(ws_data[1])}
        for r in ws_data.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in r):
                continue
            key = (r[idx['Population']], sheet_name)
            agg = by_pop_status[key]
            agg[0] += 1
            cost_usd = r[idx['Cost (USD)']]
            if cost_usd is not None:
                agg[1] += 1
                agg[2] += cost_usd

    populations = sorted(set(k[0] for k in by_pop_status))
    r_out = 2
    for pop in populations:
        for status in ('Shipped', 'Backlog'):
            agg = by_pop_status.get((pop, status))
            if not agg:
                continue
            ws.append([pop, status, agg[0], agg[1], agg[0] - agg[1], round(agg[2], 2)])
            for c in range(1, len(headers) + 1):
                ws.cell(row=r_out, column=c).font = arial
            ws.cell(row=r_out, column=6).number_format = '#,##0.00'
            r_out += 1

    manual_ws = wb['Manual Review']
    n_manual = manual_ws.max_row - 1 if manual_ws.max_row > 1 else 0
    ws.append(['Manual Review (not priced)', 'Mixed', n_manual, 0, n_manual, 0])
    for c in range(1, len(headers) + 1):
        ws.cell(row=r_out, column=c).font = arial
    r_out += 1

    total_row = r_out
    ws.cell(row=total_row, column=1, value='Total').font = bold
    ws.cell(row=total_row, column=2).font = bold
    for c in (3, 4, 5, 6):
        col_letter = get_column_letter(c)
        cell = ws.cell(row=total_row, column=c, value=f'=SUM({col_letter}2:{col_letter}{total_row - 1})')
        cell.font = bold
        if c == 6:
            cell.number_format = '#,##0.00'
    for c, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(24, len(h) + 4)


def update_readme(wb, n_vlookup, n_estimated, n_still_unresolved):
    ws = wb['README']
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    ws.append((None,))
    ws.append(('Patched with the September USD MNT price list:',))
    ws.cell(row=ws.max_row, column=1).font = bold
    lines = [
        f'1) Every MNT-sourced row ("MNT UK / Q3 AUGUST" or "Q3 AUGUST (MNT, POD)" Method) whose Shipment '
        f'Number is on file in the September USD price list (data/SEPTEMBER_USD.xlsx, column H "Solaredge rate '
        f'USD") had its Cost/Currency/Cost (USD) replaced with that confirmed quote -- {n_vlookup} row(s) '
        f'updated.',
        f'2) Every remaining MNT-sourced row with no Shipment Number match got a refreshed estimate: the '
        f'nearest available pallet count on file for the same destination country, preferring the same '
        f'customer name, drawn from the rows that now have a confirmed (non-estimated) MNT price -- '
        f'{n_estimated} row(s) estimated this round, {n_still_unresolved} still have no confirmed or '
        f'estimable MNT price on file at all for that country.',
        'Nothing else was touched: every DBS Price list 2026/Ship Cost Matrix/Canot rate card/BayWa IT row, '
        'the Manual Review sheet, and every row the user had zeroed out, were left exactly as they were.',
    ]
    for text in lines:
        ws.append((text,))
        ws.cell(row=ws.max_row, column=1).font = arial


if __name__ == '__main__':
    main()
