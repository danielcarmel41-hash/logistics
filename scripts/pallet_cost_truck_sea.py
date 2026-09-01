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

Two manual overrides sit ahead of the DBS lookup on the TRUCK sheet:
  - BayWa r.e. Solar Systems srl (Italy) lines whose ZIP is garbage text
    ("3c block" / "4c block") use the user-supplied EUR rates by pallet
    count (BAYWA_IT_OVERRIDES below) instead of the DBS zone lookup, which
    can never match a non-numeric ZIP.
  - 3PLCANOT -> Israel (domestic) lines use the user's "Canot Whs." NIS
    rate card (CANOT_*  below) instead of DBS, which has no Israel sheet.
    That card is by region (Main Land Gedera-Hadera / Gedera to south /
    Hadera and north); the region per delivery city is inferred from
    geography since it isn't in the source data — flagged in the Note
    column for the WH contact to confirm.

Both sheets also carry:
  - "SH#" (Shipment Number), populated only for Source = "Shipped" lines —
    Backlog lines have no shipment number yet in the source data ("N/A").
  - "Cost (USD)", converted from the Cost/Currency columns using the FX
    rates in FX_TO_USD below (EUR and ILS -> USD; USD is already USD).
    Rates are a same-day snapshot (mid-market, via web search on the date
    this script was run — see FX_RATES_ASOF/FX_RATES_SOURCE), not a
    contracted rate — treat the USD column as indicative and swap in the
    company's actual rate if one applies.
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

# Same-day mid-market FX snapshot (web search, see docstring). Not a
# contracted rate — swap in the company's actual rate if one applies.
FX_RATES_ASOF = '2026-09-01'
FX_RATES_SOURCE = 'xe.com / investing.com mid-market rates, retrieved 2026-09-01'
FX_TO_USD = {
    'EUR': 1.1618,
    'ILS': 0.3344,
    'USD': 1.0,
}


def to_usd(cost, currency):
    if cost is None or currency not in FX_TO_USD:
        return None
    return round(cost * FX_TO_USD[currency], 2)


# Manual override: user-supplied EUR rates for the recurring BayWa (Italy)
# lines whose ZIP field holds garbage text ("3c block" / "4c block") instead
# of a postal code, so the normal DBS zone lookup can never match them.
BAYWA_IT_CUSTOMER = 'BayWa r.e. Solar Systems srl'
BAYWA_IT_OVERRIDES = {1: 240, 2: 411, 12: 2100, 13: 1410}

# Manual override: 3PLCANOT -> Israel (domestic) trucking rates, from the
# user's "Canot Whs." rate card (NIS). A single pallet ships flat at 200 NIS
# regardless of region; 2-16 pallets take the 8T truck rate; above 16 the
# 12T truck rate (same NIS figure as 8T on this card). The truck rate itself
# depends which of the card's three regions the delivery city falls in.
# Region assignment below is by geography (not in the source data) and
# should be confirmed with the WH contact.
CANOT_SINGLE_PALLET_RATE_ILS = 200
CANOT_REGION_RATE_ILS = {
    'mainland': 1200,  # "Cannot Whs. - Main Land (Gedera-Hadera)"
    'south': 1300,     # "Cannot Whs. - Gedera to south"
    'north': 1400,     # "Cannot Whs. - Hadera and north"
}
CANOT_CITY_REGION = {
    'KIRYAT GAT': 'south',
    'ASHDOD': 'south',
    'KADIMA': 'mainland',
    'EIN HAEMEK': 'north',
    'BEIT HASHITTA': 'north',
    'INDUSTRIAL PARK KIDMAT GALIL': 'north',
}


def canot_israel_lookup(city, pallets_roundup):
    if not isinstance(pallets_roundup, (int, float)):
        return None, 'ILS', f'Invalid pallet count on this line ({pallets_roundup!r})'

    region = CANOT_CITY_REGION.get((city or '').strip().upper())
    region_note = (f'Region "{region}" inferred from city {city!r} (not in the source data) — confirm with WH.'
                   if region else f'City {city!r} not mapped to a Canot region — confirm with WH; not priced.')

    if pallets_roundup == 1:
        return CANOT_SINGLE_PALLET_RATE_ILS, 'ILS', f'Single-pallet flat rate. {region_note}'
    if not region:
        return None, 'ILS', region_note
    if pallets_roundup <= 16:
        return CANOT_REGION_RATE_ILS[region], 'ILS', f'8T truck rate, region={region}. {region_note}'
    return CANOT_REGION_RATE_ILS[region], 'ILS', f'12T truck rate, region={region}. {region_note}'


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


def sh_number(row):
    return row.get('Shipment Number') if row.get('Source') == 'Shipped' else None


TRUCK_COLUMNS = [
    ('Source', lambda row, extra: row.get('Source')),
    ('SH#', lambda row, extra: sh_number(row)),
    ('Sending WHS Code', lambda row, extra: row.get('Sending WHS Code')),
    ('ZIP', lambda row, extra: row.get('Zip')),
    ('Customer Name', lambda row, extra: row.get('Customer Name')),
    ('Destination country', lambda row, extra: row.get('Destination country')),
    ('# of Pallets per line', lambda row, extra: row.get('# of Pallets per line')),
    ('# of Pallets per line roundup', lambda row, extra: row.get('# of Pallets per line roundup')),
    ('Zone', lambda row, extra: extra['zone']),
    ('Cost', lambda row, extra: extra['cost']),
    ('Currency', lambda row, extra: extra['currency']),
    ('Cost (USD)', lambda row, extra: to_usd(extra['cost'], extra['currency'])),
    ('Note', lambda row, extra: extra['note']),
]

SEA_COLUMNS = [
    ('Source', lambda row, extra: row.get('Source')),
    ('SH#', lambda row, extra: sh_number(row)),
    ('Sending WHS Code', lambda row, extra: row.get('Sending WHS Code')),
    ('ZIP', lambda row, extra: row.get('Zip')),
    ('Customer Name', lambda row, extra: row.get('Customer Name')),
    ('Destination country', lambda row, extra: row.get('Destination country')),
    ('# of Pallets per line', lambda row, extra: row.get('# of Pallets per line')),
    ('# of Pallets per line roundup', lambda row, extra: row.get('# of Pallets per line roundup')),
    ('Matched Corridor', lambda row, extra: extra['zone']),
    ('Cost', lambda row, extra: extra['cost']),
    ('Currency', lambda row, extra: extra['currency']),
    ('Cost (USD)', lambda row, extra: to_usd(extra['cost'], extra['currency'])),
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
        pallets_roundup = row.get('# of Pallets per line roundup')
        whs = row.get('Sending WHS Code')
        customer = row.get('Customer Name')
        zip_code = row.get('Zip')

        if whs == '3PLCANOT' and row.get('Destination country') == 'Israel':
            cost, currency, note = canot_israel_lookup(row.get('City'), pallets_roundup)
            zone = 'Canot (domestic IL)'
        elif customer == BAYWA_IT_CUSTOMER and isinstance(zip_code, str) and 'block' in zip_code.lower():
            cost = BAYWA_IT_OVERRIDES.get(pallets_roundup)
            currency = 'EUR' if cost is not None else None
            zone = 'BayWa IT (manual)'
            note = ('User-supplied rate for invalid ZIP.' if cost is not None
                    else f'No user-supplied rate for {pallets_roundup} pallets — please provide.')
        else:
            zone, cost, note = dbs_lookup(dbs_book, row.get('Destination country code'), zip_code, pallets_roundup)
            currency = 'EUR' if cost is not None else None

        truck_out.append((row, {'zone': zone, 'cost': cost, 'currency': currency, 'note': note or ''}))

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

    readme = wb.create_sheet('README', 0)
    readme.column_dimensions['A'].width = 100
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    readme.append(('FX rates used for the "Cost (USD)" column',))
    readme['A1'].font = bold
    for cur, rate in FX_TO_USD.items():
        if cur == 'USD':
            continue
        readme.append((f'  1 {cur} = {rate} USD',))
        readme.cell(row=readme.max_row, column=1).font = arial
    readme.append((f'  As of: {FX_RATES_ASOF}  —  Source: {FX_RATES_SOURCE}',))
    readme.cell(row=readme.max_row, column=1).font = arial
    readme.append(('  This is a same-day market snapshot, not a contracted rate — '
                    'replace with the company\'s actual FX rate if one applies.',))
    readme.cell(row=readme.max_row, column=1).font = arial

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)

    n_priced_truck = sum(1 for _, e in truck_out if e['cost'] is not None)
    n_priced_sea = sum(1 for _, e in sea_out if e['cost'] is not None)
    print(f"TRUCK priced: {n_priced_truck}/{len(truck_out)}")
    print(f"SEA priced: {n_priced_sea}/{len(sea_out)}")


if __name__ == '__main__':
    main()
