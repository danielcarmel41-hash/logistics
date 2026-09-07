"""
Prices every line in data/Freight_costs_Data_0609.xlsx's five population
sheets, each per its own routing rule, and writes
output/Freight_Cost_Report_0609.xlsx with one output sheet per population
plus a Summary sheet totaling Cost (USD) by Shipped/Backlog.

Populations and their rules (as given):

  3PLs            Ship Cost Matrix only: Sending WHS Code + Ship Mode +
                  Destination country (no domestic/export split).

  NL - Support    Domestic (Parent Region = EUROPE) -> DBS Price list 2026,
                  by destination ZIP zone + pallet count (same logic
                  validated before: DE, 4 pallets, zone DE04 = 474 EUR).
                  Export (everything else) -> Ship Cost Matrix.

  NL - EX + DO    Forwarder = MNT -> Q3 prices AUGUST, matched by Shipment
                  Number = "PL nr." (col G), cost = col A (Solaredge rate
                  EUR); no match -> falls back to the domestic/export rule
                  below (same fallback chain used in the NL to SUP report).
                  Everyone else: domestic (Parent Region = EUROPE) -> DBS;
                  export -> Ship Cost Matrix.

  NL - DG + UK    Destination = United Kingdom -> MNT UK price list
                  (sheet "Price Q3"), matched by Customer Name (+ ZIP for
                  customers with more than one listed address). Forwarder
                  = MNT -> Q3 prices AUGUST (as above). This data pull
                  happens to have neither UK destinations nor unmatched-
                  MNT lines needing it, but the rule is applied anyway in
                  case future data does. Everything else in this sheet
                  (Netherlands/US/Australia lines, no UK/MNT match) falls
                  back to the same domestic/export split as the other NL
                  sheets, to stay consistent with "use the logic of
                  previous reports" -- flagged in the Note column since
                  the request didn't spell out this sheet's non-UK/MNT
                  lines explicitly.

  Canot - EX + DO Destination = Israel (domestic) -> the user's "Canot
                  Whs." NIS rate card by delivery-city region (Main Land
                  Gedera-Hadera / Gedera to south / Hadera and north):
                  200 ILS flat for 1 pallet, the 8T rate for 2-16, the 12T
                  rate above 16. Region per city is the same mapping
                  established earlier (Kiryat Gat/Ashdod -> south, Kadima
                  -> mainland, Ein HaEmek/Beit HaShitta/Kidmat Galil ->
                  north). Everything else (export) -> Ship Cost Matrix
                  (Sending Org Code 3PLCANOT).

Shared building blocks, all reused from earlier reports in this project:
  - DBS zone = destination country code + first 2 digits of ZIP (GB uses
    postcode-area letters; NL is a single zone); a handful of countries
    with a fixed ZIP digit count (CH/AT/BE/HU/SI/LU/NL) get a leading
    zero stripped if the ZIP is one digit too long (a padding artifact,
    not part of the code -- caught on the Swiss "04658"/Daniken case).
  - EP (pallet bracket) = "# of Pallets per line" rounded up to the next
    whole pallet, 1-33 (33 = Full load).
  - Ship Cost Matrix match key = Sending Org Code (= Sending WHS Code) +
    Ship Mode + Destination country; tries the ShipMode recorded on the
    line first, then falls back to "SEA" (many backlog/placeholder lines
    carry a default ShipMode that isn't the lane's real one). When a
    route has more than one rate on file, they are averaged (documented
    in the Note) rather than picking one arbitrarily.
  - BayWa r.e. Solar Systems srl (Italy) lines whose ZIP is garbage text
    ("3c/4c block") use the same user-supplied EUR rates by pallet count
    established earlier (1->240, 2->411, 12->2100, 13->1410).
  - FX to USD: 1 EUR = 1.1618, 1 ILS = 0.3344 (same snapshot used
    throughout this project, 2026-09-01, xe.com/investing.com -- not a
    contracted rate); Ship Cost Matrix costs are already USD.
"""
import math
import os
import re
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = lambda name: os.path.join(BASE_DIR, "data", name)
OUT_FILE = os.path.join(BASE_DIR, "output", "Freight_Cost_Report_0609.xlsx")

MAIN_FILE = DATA("Freight_costs_Data_0609.xlsx")
DBS_FILE = DATA("DBS_Price_list_2026.xlsx")
MATRIX_FILE = DATA("Ship_Cost_Matrix.xlsx")
MNT_UK_FILE = DATA("MNT_UK_price_list.xlsx")
Q3_AUG_FILE = DATA("Q3_prices_AUGUST.xlsx")

MAX_EP = 33
FX_TO_USD = {'EUR': 1.1618, 'ILS': 0.3344, 'USD': 1.0}

BAYWA_IT_CUSTOMER = 'BayWa r.e. Solar Systems srl'
BAYWA_IT_OVERRIDES = {1: 240, 2: 411, 12: 2100, 13: 1410}

CANOT_SINGLE_PALLET_RATE_ILS = 200
CANOT_REGION_RATE_ILS = {'mainland': 1200, 'south': 1300, 'north': 1400}
CANOT_CITY_REGION = {
    'KIRYAT GAT': 'south', 'ASHDOD': 'south', 'KADIMA': 'mainland',
    'EIN HAEMEK': 'north', 'BEIT HASHITTA': 'north',
    'INDUSTRIAL PARK KIDMAT GALIL': 'north',
}

EXPECTED_ZIP_DIGITS = {'CH': 4, 'AT': 4, 'BE': 4, 'HU': 4, 'SI': 4, 'LU': 4, 'NL': 4,
                       'DE': 5, 'IT': 5, 'FR': 5, 'PL': 5, 'CZ': 5, 'FI': 5}

# Known postal codes for "General Customer" placeholder lines with a blank ZIP,
# established the same way in the earlier NL to SUP report (city -> real ZIP).
CITY_ZIP_OVERRIDE = {'TAXENBACH': '5660', 'LJUBLJANA': '1000'}


def to_usd(cost, currency):
    if cost is None or currency not in FX_TO_USD:
        return None
    return round(cost * FX_TO_USD[currency], 2)


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
    expected = EXPECTED_ZIP_DIGITS.get(country_code)
    if expected and len(digits) == expected + 1 and digits[0] == '0':
        digits = digits[1:]
    if len(digits) < 2:
        return None
    return country_code + digits[:2]


def resolve_zip(zip_code, city):
    if zip_code:
        return zip_code, None
    override = CITY_ZIP_OVERRIDE.get((city or '').strip().upper())
    if override:
        return override, f'ZIP was blank; used known postal code {override} for city {city!r}.'
    return zip_code, None


def dbs_lookup(dbs_book, country_code, zip_code, pallets, city=None):
    zip_code, zip_note = resolve_zip(zip_code, city)

    def _with_zip_note(note):
        return f'{zip_note} {note}'.strip() if zip_note and note else (zip_note or note)

    zone = dbs_zone_for(country_code, zip_code)
    if not zone:
        return None, None, _with_zip_note('ZIP could not be parsed into a rate zone')
    zone_prices = dbs_book.get((country_code or '').strip().upper(), {}).get(zone)
    if not zone_prices:
        return zone, None, _with_zip_note(f'No DBS rate sheet/zone for country={country_code!r} zone={zone!r}')
    if not isinstance(pallets, (int, float)):
        return zone, None, _with_zip_note(f'Invalid pallet count on this line ({pallets!r})')
    ep = max(1, min(MAX_EP, math.ceil(pallets - 1e-9)))
    while ep <= MAX_EP:
        if ep in zone_prices:
            return zone, zone_prices[ep], zip_note
        ep += 1
    return zone, None, f'Zone {zone} has no rate bracket for {pallets} pallets'


# --------------------------------------------------------------------------
# Ship Cost Matrix
# --------------------------------------------------------------------------

def parse_ship_matrix(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Ship Cost Matrix']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    by_org = defaultdict(list)
    by_country = defaultdict(list)
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
        by_country[(country, mode, dest)].append((cost, currency, no_pal))
        if org:
            by_org[(org, mode, dest)].append((cost, currency, no_pal))
    return by_org, by_country


def _resolve_matrix_entries(entries):
    if not entries:
        return None
    costs = [e[0] for e in entries]
    currency = entries[0][1]
    no_pal = entries[0][2]
    avg_cost = round(sum(costs) / len(costs), 2)
    note = '' if len(set(costs)) <= 1 else f'Averaged {len(costs)} rates on file for this route ({costs}).'
    return avg_cost, currency, no_pal, note


def matrix_lookup(matrix, org_code, sending_country, ship_mode, dest_country):
    by_org, by_country = matrix
    mode = (ship_mode or '').strip().upper()
    for m in ([mode, 'SEA'] if mode != 'SEA' else ['SEA']):
        hit = _resolve_matrix_entries(by_org.get((org_code, m, dest_country)))
        if hit:
            return hit
    for m in ([mode, 'SEA'] if mode != 'SEA' else ['SEA']):
        hit = _resolve_matrix_entries(by_country.get((sending_country, m, dest_country)))
        if hit:
            return hit
    return None


def canot_israel_lookup(city, pallets):
    if not isinstance(pallets, (int, float)):
        return None, 'ILS', f'Invalid pallet count on this line ({pallets!r})'
    pallets = math.ceil(pallets - 1e-9)
    region = CANOT_CITY_REGION.get((city or '').strip().upper())
    region_note = (f'Region "{region}" inferred from city {city!r} (not in the source data) — confirm with WH.'
                   if region else f'City {city!r} not mapped to a Canot region — confirm with WH; not priced.')
    if pallets == 1:
        return CANOT_SINGLE_PALLET_RATE_ILS, 'ILS', f'Single-pallet flat rate. {region_note}'
    if not region:
        return None, 'ILS', region_note
    truck = '8T' if pallets <= 16 else '12T'
    return CANOT_REGION_RATE_ILS[region], 'ILS', f'{truck} truck rate, region={region}. {region_note}'


# --------------------------------------------------------------------------
# MNT UK price list (sheet "Price Q3")
# --------------------------------------------------------------------------

def parse_mnt_uk(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Price Q3']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    by_name_zip = {}
    by_name = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        name = r[idx['Customer Name']]
        addr = r[idx['Ship To Address']] or ''
        total = r[idx['Total price']]
        m = re.search(r'([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})', addr.upper())
        zip_code = m.group(1).replace(' ', '') if m else None
        if zip_code:
            by_name_zip[(name, zip_code)] = total
        by_name[name].append(total)
    return by_name_zip, by_name


def mnt_uk_lookup(mnt_book, customer_name, zip_code):
    by_name_zip, by_name = mnt_book
    zip_norm = (zip_code or '').upper().replace(' ', '')
    hit = by_name_zip.get((customer_name, zip_norm))
    if hit is not None:
        return hit, None
    rates = by_name.get(customer_name)
    if not rates:
        return None, f'Customer {customer_name!r} not found in MNT UK price list'
    if len(rates) == 1:
        return rates[0], None
    avg = round(sum(rates) / len(rates), 2)
    return avg, (f'Multiple MNT UK rates for {customer_name!r}; ZIP {zip_code!r} not an exact match — '
                 f'used average of {len(rates)} known rates, confirm with WH.')


# --------------------------------------------------------------------------
# Q3 prices AUGUST
# --------------------------------------------------------------------------

def parse_q3_august(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    by_pl_nr = {}
    by_country = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        rate = r[idx['Solaredge rate EUR']]
        if not isinstance(rate, (int, float)):
            continue
        pl_nr = r[idx['PL nr.']]
        if pl_nr:
            by_pl_nr[pl_nr] = rate
        country = r[idx['country']]
        if country:
            by_country[country].append(rate)
    country_avg = {c: round(sum(v) / len(v), 2) for c, v in by_country.items()}
    return by_pl_nr, country_avg


# --------------------------------------------------------------------------
# Main data
# --------------------------------------------------------------------------

def load_sheet_rows(sheet_name):
    wb = openpyxl.load_workbook(MAIN_FILE, data_only=True)
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(v is None for v in vals):
            continue
        rows.append(dict(zip(headers, vals)))
    return rows


def num_pallets(row):
    v = row.get('# of Pallets per line')
    return v if isinstance(v, (int, float)) else None


def baywa_override(row):
    zip_code = row.get('Zip')
    if row.get('Customer Name') == BAYWA_IT_CUSTOMER and isinstance(zip_code, str) and 'block' in zip_code.lower():
        raw_pallets = num_pallets(row)
        pallets = max(1, math.ceil(raw_pallets - 1e-9)) if raw_pallets is not None else None
        cost = BAYWA_IT_OVERRIDES.get(pallets)
        note = ('User-supplied rate for invalid ZIP.' if cost is not None
                else f'BayWa IT — no user-supplied rate for {pallets} pallets, please provide.')
        return {'method': 'BayWa IT (manual)', 'route': 'BayWa IT (manual)', 'cost': cost,
                'currency': 'EUR' if cost is not None else None, 'note': note}
    return None


def domestic_export_price(row, dbs_book, matrix):
    """Domestic (Parent Region = EUROPE) -> DBS; export (else) -> Ship Cost Matrix."""
    override = baywa_override(row)
    if override:
        return override

    pallets = num_pallets(row)
    if row.get('Parent Region') == 'EUROPE':
        zone, cost, note = dbs_lookup(dbs_book, row.get('Destination country code'), row.get('Zip'), pallets,
                                       city=row.get('City'))
        return {'method': 'DBS Price list 2026', 'route': zone, 'cost': cost,
                'currency': 'EUR' if cost is not None else None, 'note': note or ''}

    hit = matrix_lookup(matrix, row.get('Sending WHS Code'), row.get('Origin country'),
                         row.get('ShipMode'), row.get('Destination country'))
    route = f"{row.get('Sending WHS Code')}/{row.get('ShipMode')}->{row.get('Destination country')}"
    if not hit:
        return {'method': 'Ship Cost Matrix', 'route': route, 'cost': None, 'currency': None,
                'note': f'No Ship Cost Matrix rate for {route} — confirm with WH contact.'}
    cost, currency, no_pal, note = hit
    if isinstance(no_pal, (int, float)) and isinstance(pallets, (int, float)) and pallets > no_pal:
        note = (note + ' ' if note else '') + f'Line has {pallets} pallets, exceeds this corridor\'s {no_pal:.0f}-pallet rate — verify capacity.'
    return {'method': 'Ship Cost Matrix', 'route': route, 'cost': cost, 'currency': currency, 'note': note}


def price_3pls(row, matrix):
    hit = matrix_lookup(matrix, row.get('Sending WHS Code'), row.get('Origin country'),
                         row.get('ShipMode'), row.get('Destination country'))
    route = f"{row.get('Sending WHS Code')}/{row.get('ShipMode')}->{row.get('Destination country')}"
    if not hit:
        return {'method': 'Ship Cost Matrix', 'route': route, 'cost': None, 'currency': None,
                'note': f'No Ship Cost Matrix rate for {route} — confirm with WH contact.'}
    cost, currency, no_pal, note = hit
    return {'method': 'Ship Cost Matrix', 'route': route, 'cost': cost, 'currency': currency, 'note': note}


def price_nl_support(row, dbs_book, matrix):
    return domestic_export_price(row, dbs_book, matrix)


def price_mnt_q3(row, q3_by_pl_nr, q3_country_avg, fallback_fn):
    shipment_number = row.get('Shipment Number')
    rate = q3_by_pl_nr.get(shipment_number)
    if rate is not None:
        return {'method': 'Q3 prices AUGUST (MNT)', 'route': shipment_number, 'cost': rate,
                'currency': 'EUR', 'note': 'Matched by Shipment Number = PL nr.'}
    fallback = fallback_fn()
    fallback = dict(fallback)
    fallback['note'] = (f'MNT — no Q3 AUGUST shipment matched by Shipment Number; ' + fallback['note']).strip()
    return fallback


def price_nl_ex_do(row, dbs_book, matrix, q3_by_pl_nr, q3_country_avg):
    if row.get('Forwarder') == 'MNT':
        return price_mnt_q3(row, q3_by_pl_nr, q3_country_avg,
                             lambda: domestic_export_price(row, dbs_book, matrix))
    return domestic_export_price(row, dbs_book, matrix)


def price_nl_dg_uk(row, dbs_book, matrix, mnt_uk_book, q3_by_pl_nr, q3_country_avg):
    if row.get('Destination country') == 'United Kingdom':
        cost, note = mnt_uk_lookup(mnt_uk_book, row.get('Customer Name'), row.get('Zip'))
        return {'method': 'MNT UK price list (Q3)', 'route': 'UK', 'cost': cost,
                'currency': 'EUR' if cost is not None else None,
                'note': note or 'Matched by customer name.'}
    if row.get('Forwarder') == 'MNT':
        return price_mnt_q3(row, q3_by_pl_nr, q3_country_avg,
                             lambda: domestic_export_price(row, dbs_book, matrix))
    rec = domestic_export_price(row, dbs_book, matrix)
    rec = dict(rec)
    rec['note'] = ('Not UK destination / not MNT — no rule specified for this line in the request; '
                    'used the domestic(DBS)/export(matrix) fallback from the other NL sheets. '
                    + rec['note']).strip()
    return rec


def price_canot(row, matrix):
    if row.get('Destination country') == 'Israel':
        cost, currency, note = canot_israel_lookup(row.get('City'), num_pallets(row))
        return {'method': 'Canot rate card', 'route': 'Canot (domestic IL)', 'cost': cost,
                'currency': currency, 'note': note}
    hit = matrix_lookup(matrix, row.get('Sending WHS Code'), row.get('Origin country'),
                         row.get('ShipMode'), row.get('Destination country'))
    route = f"{row.get('Sending WHS Code')}/{row.get('ShipMode')}->{row.get('Destination country')}"
    if not hit:
        return {'method': 'Ship Cost Matrix', 'route': route, 'cost': None, 'currency': None,
                'note': f'No Ship Cost Matrix rate for {route} — confirm with WH contact.'}
    cost, currency, no_pal, note = hit
    return {'method': 'Ship Cost Matrix', 'route': route, 'cost': cost, 'currency': currency, 'note': note}


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------

OUT_COLUMNS = [
    ('Source', lambda row, rec: row.get('Source')),
    ('Shipment Number', lambda row, rec: row.get('Shipment Number')),
    ('SE Order#', lambda row, rec: row.get('SE Order#')),
    ('Sending WHS Code', lambda row, rec: row.get('Sending WHS Code')),
    ('Dest WHS Code', lambda row, rec: row.get('Dest WHS Code')),
    ('Forwarder', lambda row, rec: row.get('Forwarder')),
    ('ShipMode', lambda row, rec: row.get('ShipMode')),
    ('Customer Name', lambda row, rec: row.get('Customer Name')),
    ('Destination country', lambda row, rec: row.get('Destination country')),
    ('ZIP', lambda row, rec: row.get('Zip')),
    ('Family Type', lambda row, rec: row.get('Family Type')),
    ('# of Pallets per line', lambda row, rec: row.get('# of Pallets per line')),
    ('Method', lambda row, rec: rec['method']),
    ('Zone / Route', lambda row, rec: rec['route']),
    ('Cost', lambda row, rec: rec['cost']),
    ('Currency', lambda row, rec: rec['currency']),
    ('Cost (USD)', lambda row, rec: to_usd(rec['cost'], rec['currency'])),
    ('Note', lambda row, rec: rec['note']),
]


def write_sheet(wb, title, rows_with_rec):
    ws = wb.create_sheet(title)
    headers = [c[0] for c in OUT_COLUMNS]
    ws.append(headers)
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold
    for row, rec in rows_with_rec:
        ws.append([getter(row, rec) for _, getter in OUT_COLUMNS])
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).font = arial
    for c, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(14, min(32, len(h) + 4))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    ws.freeze_panes = 'A2'
    return ws, len(rows_with_rec)


def write_summary_sheet(wb, sheet_row_counts):
    ws = wb.create_sheet('Summary', 1)
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    headers = ['Population', 'Total Lines', 'Shipped (USD)', 'Backlog (USD)', 'Total (USD)']
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold

    source_col = 'A'  # Source
    cost_usd_col = 'Q'  # Cost (USD) - 17th column
    row_n = 2
    data_rows = []
    for sheet_name, n_rows in sheet_row_counts:
        rng_source = f"'{sheet_name}'!{source_col}2:{source_col}{n_rows + 1}"
        rng_cost = f"'{sheet_name}'!{cost_usd_col}2:{cost_usd_col}{n_rows + 1}"
        ws.append([sheet_name, n_rows,
                   f'=SUMIF({rng_source},"Shipped",{rng_cost})',
                   f'=SUMIF({rng_source},"Backlog",{rng_cost})',
                   f'=C{row_n}+D{row_n}'])
        data_rows.append(row_n)
        row_n += 1

    total_row = row_n
    ws.append(['Grand Total', f'=SUM(B2:B{total_row - 1})',
               f'=SUM(C2:C{total_row - 1})', f'=SUM(D2:D{total_row - 1})', f'=SUM(E2:E{total_row - 1})'])

    for r in range(2, total_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = bold if r == total_row else arial
            if c >= 3:
                cell.number_format = '#,##0.00'

    for c, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(16, len(h) + 4)
    return ws


def main():
    print("Loading price sources...")
    dbs_book = parse_dbs_price_book(DBS_FILE)
    matrix = parse_ship_matrix(MATRIX_FILE)
    mnt_uk_book = parse_mnt_uk(MNT_UK_FILE)
    q3_by_pl_nr, q3_country_avg = parse_q3_august(Q3_AUG_FILE)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheet_row_counts = []

    print("Pricing 3PLs...")
    rows = load_sheet_rows('3PLs')
    recs = [(r, price_3pls(r, matrix)) for r in rows]
    _, n = write_sheet(wb, '3PLs', recs)
    sheet_row_counts.append(('3PLs', n))
    print(f"  {n} rows, {sum(1 for _, rec in recs if rec['cost'] is not None)} priced")

    print("Pricing NL - Support...")
    rows = load_sheet_rows('NL - Support ')
    recs = [(r, price_nl_support(r, dbs_book, matrix)) for r in rows]
    _, n = write_sheet(wb, 'NL - Support', recs)
    sheet_row_counts.append(('NL - Support', n))
    print(f"  {n} rows, {sum(1 for _, rec in recs if rec['cost'] is not None)} priced")

    print("Pricing NL - EX + DO...")
    rows = load_sheet_rows('NL - EX + DO')
    recs = [(r, price_nl_ex_do(r, dbs_book, matrix, q3_by_pl_nr, q3_country_avg)) for r in rows]
    _, n = write_sheet(wb, 'NL - EX + DO', recs)
    sheet_row_counts.append(('NL - EX + DO', n))
    print(f"  {n} rows, {sum(1 for _, rec in recs if rec['cost'] is not None)} priced")

    print("Pricing NL - DG + UK...")
    rows = load_sheet_rows('NL - DG + UK')
    recs = [(r, price_nl_dg_uk(r, dbs_book, matrix, mnt_uk_book, q3_by_pl_nr, q3_country_avg)) for r in rows]
    _, n = write_sheet(wb, 'NL - DG + UK', recs)
    sheet_row_counts.append(('NL - DG + UK', n))
    print(f"  {n} rows, {sum(1 for _, rec in recs if rec['cost'] is not None)} priced")

    print("Pricing Canot - EX + DO...")
    rows = load_sheet_rows('Canot - EX + DO')
    recs = [(r, price_canot(r, matrix)) for r in rows]
    _, n = write_sheet(wb, 'Canot - EX + DO', recs)
    sheet_row_counts.append(('Canot - EX + DO', n))
    print(f"  {n} rows, {sum(1 for _, rec in recs if rec['cost'] is not None)} priced")

    readme = wb.create_sheet('README', 0)
    write_summary_sheet(wb, sheet_row_counts)
    readme.column_dimensions['A'].width = 105
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    lines = [
        ('Freight Cost Report — 5 populations, priced per their own rule', bold),
        ('', arial),
        ('3PLs: Ship Cost Matrix only (Sending WHS Code + Ship Mode + destination country).', arial),
        ('NL - Support: domestic (Parent Region = EUROPE) -> DBS Price list 2026; export -> Ship Cost Matrix.', arial),
        ('NL - EX + DO: Forwarder=MNT -> Q3 AUGUST by Shipment Number; else domestic->DBS, export->matrix.', arial),
        ('NL - DG + UK: Destination=UK -> MNT UK price list; Forwarder=MNT -> Q3 AUGUST; else domestic/export '
         'fallback (this data pull has neither UK nor an unmatched-MNT line, so this is a documented '
         'extension for lines the request did not explicitly cover — see each such line\'s Note).', arial),
        ('Canot - EX + DO: Destination=Israel -> Canot Whs. NIS rate card by city region; else -> Ship Cost Matrix.', arial),
        ('', arial),
        ('Ship Cost Matrix match: Sending WHS Code + Ship Mode + destination country, falling back to "SEA" '
         'mode when the literal ShipMode has no rate (many backlog lines carry a placeholder mode). Multiple '
         'rates on file for one route are averaged, noted in the Note column.', arial),
        ('BayWa r.e. Solar Systems srl (Italy) lines with a "3c/4c block" ZIP use the same user-supplied EUR '
         'rates by pallet count established earlier, not the DBS lookup.', arial),
        ('FX to USD: 1 EUR = 1.1618, 1 ILS = 0.3344 (2026-09-01 snapshot, xe.com/investing.com — not a '
         'contracted rate); Ship Cost Matrix figures are already USD.', arial),
        ('', arial),
        ('"Summary" totals Cost (USD) per population, split Shipped/Backlog, via SUMIF formulas.', bold),
        ('Any line with a blank Cost has a Note explaining why (no matrix corridor, no DBS coverage, '
         'invalid ZIP/pallet count, etc.) — these need a human check.', arial),
    ]
    for text, font in lines:
        readme.append((text,))
        readme.cell(row=readme.max_row, column=1).font = font

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)
    print("Saved", OUT_FILE)


if __name__ == '__main__':
    main()
