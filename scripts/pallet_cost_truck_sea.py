"""
Simple, single-purpose version of the freight costing: for every line in
data/Freight_costs_3108.xlsx (sheet "DB B2B"), price it by pallet count using
the "# of Pallets per line roundup" column (already rounded up in the source
data — used as-is, no re-rounding).

Two sheets, split by "Ship Method":

  "TRUCK"  — every line whose Ship Method contains "TRUCK" (the LAND/truck
             lines: DBSCHENKER-LAND-TRUCK, DEFAULT-LAND-TRUCK, MNT-LAND-TRUCK,
             FC-LAND-TRUCK). Priced from data/DBS_Price_list_2026.xlsx:
               - each sheet in that workbook is one destination country
               - column G ("EP") holds the pallet-count bracket (1..33,
                 33 = "Full load")
               - the rate zone is derived from destination country + ZIP
                 (country code + first 2 digits of the ZIP; the price list
                 itself labels this "2 Digit zipcodes" on its NL sheet).
                 UK is the one exception: its zones are postcode-area
                 letters (e.g. "LE3 1BY" -> GBLE).
             Verified against the user's own example: Germany, 4 pallets,
             zone DE04 = 474 EUR.

  "SEA"    — every line whose Ship Method contains "SEA" (ICL-SEA-STANDARD,
             DEFAULT-SEA-STANDARD, MNT-SEA-STANDARD). Priced from
             data/Ship_Cost_Matrix.xlsx, matched by Sending Org Code
             (= Sending WHS Code) + Ship Mode ("SEA") + destination country.
             Rows with no matching corridor in the matrix are left blank for
             manual entry, as requested.

COURIER (UPS) and LAND-DIRECT lines are neither TRUCK nor SEA and are left
out of both sheets.
"""
import os
import re

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = lambda name: os.path.join(BASE_DIR, "data", name)
OUT_FILE = os.path.join(BASE_DIR, "output", "Pallet_Cost_by_Ship_Method.xlsx")

FREIGHT_FILE = DATA("Freight_costs_3108.xlsx")
DBS_FILE = DATA("DBS_Price_list_2026.xlsx")
MATRIX_FILE = DATA("Ship_Cost_Matrix.xlsx")

MAX_EP = 33


# --------------------------------------------------------------------------
# DBS Price list 2026
# --------------------------------------------------------------------------

def parse_dbs_price_book(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    book = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row = next(
            (r for r in range(1, 15)
             if any(ws.cell(row=r, column=c).value == 'EP' for c in range(1, ws.max_column + 1))),
            None,
        )
        if header_row is None:
            continue
        ep_col = None
        zone_cols = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            if v == 'EP':
                ep_col = c
            elif isinstance(v, str) and v.strip().upper().startswith(sheet_name.upper()):
                zone_cols[v.strip()] = c
            elif isinstance(v, str) and sheet_name.upper() == 'NL' and re.match(r'^\d+\s*-\s*\d+$', v.strip()):
                zone_cols['NL'] = c

        zones = {z: {} for z in zone_cols}
        data_start = header_row + 2
        for r in range(data_start, ws.max_row + 1):
            ep_val = ws.cell(row=r, column=ep_col).value
            if not isinstance(ep_val, (int, float)):
                continue
            ep_val = int(ep_val)
            for zone, col in zone_cols.items():
                price = ws.cell(row=r, column=col).value
                if isinstance(price, (int, float)):
                    zones[zone][ep_val] = price
        book[sheet_name] = zones
    return book


def dbs_zone_for(country_code, zip_code):
    country_code = (country_code or '').strip().upper()
    zip_code = str(zip_code or '').strip().upper()
    if country_code == 'NL':
        return 'NL'
    if country_code == 'GB':
        m = re.match(r'^[A-Z]+', zip_code.replace(' ', ''))
        return 'GB' + m.group() if m else None
    digits = ''.join(ch for ch in zip_code if ch.isdigit())
    if len(digits) < 2:
        return None
    return country_code + digits[:2]


def dbs_lookup(dbs_book, country_code, zip_code, pallets_roundup):
    zone = dbs_zone_for(country_code, zip_code)
    if not zone:
        return None, None, 'ZIP could not be parsed into a rate zone'
    zone_prices = dbs_book.get((country_code or '').strip().upper(), {}).get(zone)
    if not zone_prices:
        return zone, None, f'No DBS rate sheet/zone for country={country_code!r} zone={zone!r}'
    if not isinstance(pallets_roundup, (int, float)):
        return zone, None, f'Invalid pallet count on this line ({pallets_roundup!r})'
    ep = max(1, min(MAX_EP, int(pallets_roundup)))
    while ep <= MAX_EP:
        if ep in zone_prices:
            return zone, zone_prices[ep], None
        ep += 1
    return zone, None, f'Zone {zone} has no rate bracket for {pallets_roundup} pallets'


# --------------------------------------------------------------------------
# Ship Cost Matrix (SEA)
# --------------------------------------------------------------------------

def parse_ship_matrix(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Ship Cost Matrix']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    by_org = {}
    by_country = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        org = r[idx['Sending Org Code**']]
        country = r[idx['Sending Country Name']]
        mode = (r[idx['Ship Mode*']] or '').strip().upper()
        dest = r[idx['Dest Country Name*']]
        cost = r[idx['Cost*']]
        currency = r[idx['Currency']]
        no_pal = r[idx['No. Pallets']]
        if not isinstance(cost, (int, float)) or not dest or not mode:
            continue
        by_country.setdefault((country, mode, dest), (cost, currency, no_pal))
        if org:
            by_org[(org, mode, dest)] = (cost, currency, no_pal)
    return by_org, by_country


def matrix_lookup(matrix, org_code, sending_country, dest_country):
    by_org, by_country = matrix
    hit = by_org.get((org_code, 'SEA', dest_country))
    if hit:
        return hit
    return by_country.get((sending_country, 'SEA', dest_country))


# --------------------------------------------------------------------------
# Main data
# --------------------------------------------------------------------------

def load_freight_rows():
    wb = openpyxl.load_workbook(FREIGHT_FILE, data_only=True)
    ws = wb['DB B2B']
    headers = [c.value for c in ws[1]]
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(v is None for v in vals):
            continue
        rows.append(dict(zip(headers, vals)))
    return rows


TRUCK_COLUMNS = [
    ('Source', lambda row, extra: row.get('Source')),
    ('Sending WHS Code', lambda row, extra: row.get('Sending WHS Code')),
    ('ZIP', lambda row, extra: row.get('Zip')),
    ('Customer Name', lambda row, extra: row.get('Customer Name')),
    ('Destination country', lambda row, extra: row.get('Destination country')),
    ('# of Pallets per line', lambda row, extra: row.get('# of Pallets per line')),
    ('# of Pallets per line roundup', lambda row, extra: row.get('# of Pallets per line roundup')),
    ('Zone', lambda row, extra: extra['zone']),
    ('Cost (EUR)', lambda row, extra: extra['cost']),
    ('Note', lambda row, extra: extra['note']),
]

SEA_COLUMNS = [
    ('Source', lambda row, extra: row.get('Source')),
    ('Sending WHS Code', lambda row, extra: row.get('Sending WHS Code')),
    ('ZIP', lambda row, extra: row.get('Zip')),
    ('Customer Name', lambda row, extra: row.get('Customer Name')),
    ('Destination country', lambda row, extra: row.get('Destination country')),
    ('# of Pallets per line', lambda row, extra: row.get('# of Pallets per line')),
    ('# of Pallets per line roundup', lambda row, extra: row.get('# of Pallets per line roundup')),
    ('Matched Corridor', lambda row, extra: extra['zone']),
    ('Cost', lambda row, extra: extra['cost']),
    ('Currency', lambda row, extra: extra['currency']),
    ('Note', lambda row, extra: extra['note']),
]


def write_sheet(wb, title, rows_with_extra, columns):
    ws = wb.create_sheet(title)
    headers = [c[0] for c in columns]
    ws.append(headers)
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold
    for row, extra in rows_with_extra:
        ws.append([getter(row, extra) for _, getter in columns])
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).font = arial
    for c, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(14, min(40, len(header) + 4))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    ws.freeze_panes = 'A2'
    return ws


def main():
    dbs_book = parse_dbs_price_book(DBS_FILE)
    matrix = parse_ship_matrix(MATRIX_FILE)
    rows = load_freight_rows()

    truck_rows = [r for r in rows if 'TRUCK' in (r.get('Ship Method') or '').upper()]
    sea_rows = [r for r in rows if 'SEA' in (r.get('Ship Method') or '').upper()]
    print(f"TRUCK lines: {len(truck_rows)}, SEA lines: {len(sea_rows)}")

    truck_out = []
    for row in truck_rows:
        zone, cost, note = dbs_lookup(
            dbs_book,
            row.get('Destination country code'),
            row.get('Zip'),
            row.get('# of Pallets per line roundup'),
        )
        truck_out.append((row, {'zone': zone, 'cost': cost, 'note': note or ''}))

    sea_out = []
    for row in sea_rows:
        hit = matrix_lookup(matrix, row.get('Sending WHS Code'), row.get('Origin country'),
                             row.get('Destination country'))
        if hit:
            cost, currency, no_pal = hit
            pallets_roundup = row.get('# of Pallets per line roundup')
            note = ''
            if isinstance(no_pal, (int, float)) and isinstance(pallets_roundup, (int, float)) \
                    and pallets_roundup > no_pal:
                note = (f'This corridor\'s matrix rate covers up to {no_pal:.0f} pallets; '
                        f'this line has {pallets_roundup} — verify capacity.')
            extra = {'zone': f"{row.get('Sending WHS Code')} -> {row.get('Destination country')}",
                     'cost': cost, 'currency': currency, 'note': note}
        else:
            extra = {'zone': f"{row.get('Sending WHS Code')} -> {row.get('Destination country')}",
                     'cost': None, 'currency': None,
                     'note': 'No SEA rate found in Ship Cost Matrix for this corridor — fill in manually.'}
        sea_out.append((row, extra))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_sheet(wb, 'TRUCK - DBS Price List', truck_out, TRUCK_COLUMNS)
    write_sheet(wb, 'SEA - Ship Cost Matrix', sea_out, SEA_COLUMNS)

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)

    n_priced_truck = sum(1 for _, e in truck_out if e['cost'] is not None)
    n_priced_sea = sum(1 for _, e in sea_out if e['cost'] is not None)
    print(f"TRUCK priced: {n_priced_truck}/{len(truck_out)}")
    print(f"SEA priced: {n_priced_sea}/{len(sea_out)}")


if __name__ == '__main__':
    main()
