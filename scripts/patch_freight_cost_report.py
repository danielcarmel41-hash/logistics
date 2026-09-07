"""
Patches the user's manually-edited Freight_Cost_Report_2.xlsx in place
(preserving every row they deleted/kept) rather than regenerating the whole
workbook from source data. Three changes, each scoped as tightly as the
request allows:

1. "NL - EX + DO" and "Canot - EX + DO": every group of lines sharing one
   Shipment Number (or SE Order# for Backlog lines, which have no shipment
   number yet) that the user flagged with Cost = 0 was getting the Ship
   Cost Matrix's flat per-shipment/corridor rate applied to *every line*
   instead of once per shipment -- e.g. one Israel shipment with 19 lines
   was costed at 4800 USD x 19. Recomputes the whole group (all its lines,
   including any the user had NOT zeroed, since leaving that one line at
   the full flat rate while zeroing its siblings would make the group's
   total wrong) as: flat corridor rate x (line's pallet share of the
   group's total pallets). Groups with no flagged line are left untouched.

2. "NL - DG + UK": re-priced sheet-wide per the new rule -- destination
   United Kingdom -> MNT UK price list (unchanged); everything else ->
   Q3_prices_AUGUST column A (Solaredge rate EUR), matched by destination
   country and the closest available pallet count on file for that
   country (Q3 doesn't have every pallet count, so nearest-match is used
   and noted). Countries with no Q3 data at all (United States, Australia
   in this data) fall back to the DBS/Ship-Cost-Matrix domestic/export
   logic used elsewhere, flagged in the Note.

3. Summary sheet rebuilt to the requested layout (Category / Total Lines /
   Priced / Unpriced / Total Cost (USD), rows Shipped/Backlog/Total) using
   whole-column SUMIF/COUNTIFS formulas against each population sheet, so
   deleting a row or editing a Cost (USD) value updates the totals
   automatically -- a whole-column reference never needs resizing.
"""
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
    by pallet share within each shipment/order group that has a flagged (0)
    line."""
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
        has_zero = any(ws.cell(row=r, column=COL['Cost']).value == 0 for r in group_rows)
        if not has_zero or len(group_rows) < 2:
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
    # the grand total above is visible on its own.
    section_row = 6
    ws.cell(row=section_row, column=1, value='By Population').font = bold
    header_row = section_row + 1
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h).font = bold

    first_data_row = header_row + 1
    for i, s in enumerate(sheet_names):
        r = first_data_row + i
        ws.cell(row=r, column=1, value=s)
        ws.cell(row=r, column=2, value=f"=COUNTA('{s}'!$A$2:$A$100000)")
        ws.cell(row=r, column=3, value=f"=COUNTIF('{s}'!$Q$2:$Q$100000,\"<>\")")
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
        ws.cell(row=r, column=5, value=f"=SUM('{s}'!$Q:$Q)")
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).font = arial
            if c == 5:
                ws.cell(row=r, column=c).number_format = '#,##0.00'

    total_row = first_data_row + len(sheet_names)
    ws.cell(row=total_row, column=1, value='Total').font = bold
    for c, col_letter in zip(range(2, 6), ['B', 'C', 'D', 'E']):
        first = get_column_letter(c) + str(first_data_row)
        last = get_column_letter(c) + str(total_row - 1)
        cell = ws.cell(row=total_row, column=c, value=f'=SUM({first}:{last})')
        cell.font = bold
        if c == 5:
            cell.number_format = '#,##0.00'

    for c, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(20, len(h) + 4)
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
         'exactly as they left them). Three targeted fixes were applied on top of that:', arial),
        ('', arial),
        ('1) NL - EX + DO / Canot - EX + DO: any Shipment Number (or SE Order# for Backlog) group where '
         'the Ship Cost Matrix\'s one flat corridor rate had been applied to every line -- inflating the '
         'total by the line count -- and that the user flagged with Cost = 0, was reallocated pro-rata by '
         'pallet share across the whole group (including a line the user had not zeroed, where leaving it '
         'at the full flat rate would have made the group\'s total wrong).', arial),
        ('2) NL - DG + UK: re-priced sheet-wide. Destination = United Kingdom -> MNT UK price list '
         '(sheet "Price Q3"), matched by customer. Everything else -> Q3_prices_AUGUST column A '
         '(Solaredge rate EUR), matched by destination country + the nearest pallet count Q3 has on file '
         'for that country (noted per line). Countries with no Q3 data at all (United States, Australia in '
         'this data) fall back to the DBS/Ship-Cost-Matrix domestic/export logic, flagged in the Note.', arial),
        ('3) Summary rebuilt to Category / Total Lines / Priced / Unpriced / Total Cost (USD), rows '
         'Shipped / Backlog / Total, using whole-column SUMIFS/COUNTIFS formulas -- deleting a row or '
         'editing a Cost (USD) cell updates every total automatically, no range to resize.', arial),
        ('', arial),
        ('3PLs and NL - Support were left untouched (no flagged rows, no rule change requested for them).', arial),
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

    print("Fixing Ship Cost Matrix duplicate-line groups...")
    for sheet_name in ['NL - EX + DO', 'Canot - EX + DO']:
        print(f" {sheet_name}:")
        fix_matrix_duplicate_groups(wb[sheet_name], matrix)

    print("Repricing NL - DG + UK...")
    reprice_nl_dg_uk(wb['NL - DG + UK'], dbs_book, matrix, mnt_uk_book, q3_country_table)

    population_sheets = ['3PLs', 'NL - Support', 'NL - EX + DO', 'NL - DG + UK', 'Canot - EX + DO']
    print("Rebuilding Summary...")
    rebuild_summary(wb, population_sheets)
    rebuild_readme(wb)

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)
    print("Saved", OUT_FILE)


if __name__ == '__main__':
    main()
