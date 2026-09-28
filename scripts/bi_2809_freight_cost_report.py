"""
Prices every line in data/BI_2809.xlsx (single flat "BI 2809" sheet) and
writes output/BI_2809_Freight_Cost_Report.xlsx (Shipped / Backlog /
Manual Review / Summary / README), consolidating same-shipment lines into
one row per (Shipment Number, pricing method) group.

Populations and routing rules (as given), checked in this order:

  Sending WHS Code = 3PLCANOT (Israel):
    - Destination = Israel (domestic)  -> Canot Whs. NIS rate card by
      delivery-city region (200 ILS flat / 1 pallet, 1200/1300/1400 ILS
      8T/12T for 2+ -- same rate card confirmed again against the
      forwarded price-list screenshot).
    - Everything else (export)         -> Ship Cost Matrix, Sending Org
      Code 3PLCANOT + Ship Mode (SEA) + destination country.

  Sending WHS Code in {3PLDBSNL, 3PLDBSBRNL} (Netherlands):
    1. Family Type = Battery, OR Destination country = United Kingdom
                                        -> priced ONLY from Q3 prices
       AUGUST (Solaredge rate EUR, column A, matched by Shipment Number
       when there is a real Logistic POD) or MNT UK price list (sheet
       "Price Q3", matched by Customer Name/ZIP) -- no DBS/matrix
       fallback (same strict rule as the last BI 1409 fix). Unmatched
       lines are flagged, not silently priced another way.
    2. Dest WHS Code ends "SUP" (support warehouse)
                                        -> same forwarder-based rule as
       #3 below ("as done before").
    3. Everything else, split by Forwarder:
         - DBSCHENKER                  -> DBS Price list 2026, by
           destination zone (country code + first 2 digits of ZIP, or
           GB postcode area) + total pallets.
         - MNT                          -> Q3 prices AUGUST (Solaredge
           rate EUR), matched by Shipment Number when there is a real
           Logistic POD; no match -> flagged, not priced via DBS.
         - Any other forwarder          -> NOT priced here -- collected
           in the "Manual Review" sheet for the user to price by hand,
           per the request.
       BayWa r.e. Solar Systems srl (Italy) lines whose ZIP is garbage
       text ("3c/4c block") use the fixed EUR-by-pallet override table
       instead of DBS, regardless of forwarder.
    4. ShipMode = SEA                  -> Ship Cost Matrix (export),
       instead of the forwarder split above.

  Any other Sending WHS Code (3PLDBSTW, 3PLEXPAU, 3PLAWLNL, 3PLMCNY,
  3PLMCCA, ...)                        -> Ship Cost Matrix.

Every rate source above gives ONE flat price per shipment/order, not per
line -- lines are grouped by Shipment Number (or SE Order# when there is
none yet) *within* each pricing method (a single Shipment Number can
split across methods, e.g. Battery and non-Battery lines in one truck),
looking up one price for the group's total whole-number pallets
("# of Pallets per line roundup", ceil()'d defensively).

Per the explicit request this round, the output does NOT keep one row
per source line: every group above is written as a single consolidated
row (summed pallets, the group's total cost, a "Lines" count of how many
source lines were merged, and every other column taken from the group
-- joined/listed when it genuinely varies within the group, e.g. Family
Type or SE Order#). Customer Name and Destination country are always
uniform within a Shipment Number group in this data pull (verified) so
are never ambiguous.

Date (source column W) and Shipment Number (source column U) are both
kept as their own output columns, per the request.

FX to USD: 1 EUR = 1.1392, 1 ILS = 0.32559 (2026-09-28 snapshot,
xe.com/tradingeconomics.com -- not a contracted rate); Ship Cost Matrix
costs are already USD.
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
OUT_FILE = os.path.join(BASE_DIR, "output", "BI_2809_Freight_Cost_Report.xlsx")

MAIN_FILE = DATA("BI_2809.xlsx")
DBS_FILE = DATA("DBS_Price_list_2026_v3.xlsx")
MATRIX_FILE = DATA("Ship_Cost_Matrix_v9.xlsx")
MNT_UK_FILE = DATA("MNT_UK_price_list_v3.xlsx")
Q3_AUG_FILE = DATA("Q3_prices_AUGUST_v3.xlsx")

MAX_EP = 33
FX_TO_USD = {'EUR': 1.1392, 'ILS': 0.32559, 'USD': 1.0}

NL_WHS = {'3PLDBSNL', '3PLDBSBRNL'}
CANOT_WHS = '3PLCANOT'
DBSCHENKER_FORWARDERS = {'DBSCHENKER'}
MNT_FORWARDERS = {'MNT'}

BAYWA_IT_CUSTOMER = 'BayWa r.e. Solar Systems srl'
BAYWA_IT_KNOWN_RATES = {1: 240, 2: 411, 12: 2100, 13: 1410}
# 4/6/8-pallet shipments have no rate on file -- interpolated linearly between
# the nearest known brackets below and above (2 pallets->411 EUR, 12 pallets->
# 2100 EUR), per the user's explicit request to derive these from the known
# per-shipment prices by pallet count. Flagged as estimated in the Note.
BAYWA_IT_INTERPOLATED = {}
_lo_p, _lo_v = 2, BAYWA_IT_KNOWN_RATES[2]
_hi_p, _hi_v = 12, BAYWA_IT_KNOWN_RATES[12]
_slope = (_hi_v - _lo_v) / (_hi_p - _lo_p)
for _p in (4, 6, 8):
    BAYWA_IT_INTERPOLATED[_p] = round(_lo_v + _slope * (_p - _lo_p), 2)
BAYWA_IT_OVERRIDES = {**BAYWA_IT_KNOWN_RATES, **BAYWA_IT_INTERPOLATED}

CANOT_SINGLE_PALLET_RATE_ILS = 200
CANOT_REGION_RATE_ILS = {'mainland': 1200, 'south': 1300, 'north': 1400}
CANOT_CITY_REGION = {
    'KIRYAT GAT': 'south', 'ASHDOD': 'south', 'KADIMA': 'mainland',
    'EIN HAEMEK': 'north', 'BEIT HASHITTA': 'north',
    'INDUSTRIAL PARK KIDMAT GALIL': 'north',
}

EXPECTED_ZIP_DIGITS = {'CH': 4, 'AT': 4, 'BE': 4, 'HU': 4, 'SI': 4, 'LU': 4, 'NL': 4,
                       'DE': 5, 'IT': 5, 'FR': 5, 'PL': 5, 'CZ': 5, 'FI': 5}


def to_usd(cost, currency):
    if cost is None or currency not in FX_TO_USD:
        return None
    return round(cost * FX_TO_USD[currency], 2)


# --------------------------------------------------------------------------
# Price sources
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


def dbs_price_for_pallets(zone_prices, total_pallets):
    ep = max(1, min(MAX_EP, math.ceil(total_pallets - 1e-9)))
    while ep <= MAX_EP:
        if ep in zone_prices:
            return zone_prices[ep]
        ep += 1
    return None


def dbs_lookup(dbs_book, country_code, zip_code, total_pallets):
    zone = dbs_zone_for(country_code, zip_code)
    if not zone:
        return None, None, 'ZIP could not be parsed into a rate zone -- confirm with WH contact.'
    zone_prices = dbs_book.get((country_code or '').strip().upper(), {}).get(zone)
    if not zone_prices:
        return zone, None, f'No DBS rate sheet/zone for country={country_code!r} zone={zone!r}.'
    price = dbs_price_for_pallets(zone_prices, total_pallets)
    if price is None:
        return zone, None, f'Zone {zone} has no rate bracket for {total_pallets:g} pallets.'
    return zone, price, None


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


def canot_israel_lookup(city, total_pallets):
    pallets = max(1, math.ceil(total_pallets - 1e-9))
    region = CANOT_CITY_REGION.get((city or '').strip().upper())
    region_note = (f'Region "{region}" inferred from city {city!r} (not in the source data).'
                   if region else f'City {city!r} not mapped to a Canot region -- confirm with WH; not priced.')
    if pallets == 1:
        return CANOT_SINGLE_PALLET_RATE_ILS, 'ILS', f'Single-pallet flat rate. {region_note}'
    if not region:
        return None, 'ILS', region_note
    truck = '8T' if pallets <= 16 else '12T'
    return CANOT_REGION_RATE_ILS[region], 'ILS', f'{truck} truck rate, region={region}. {region_note}'


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
        if not isinstance(total, (int, float)):
            continue
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
        return None, f'Customer {customer_name!r} not found in MNT UK price list.'
    if len(rates) == 1:
        return rates[0], None
    avg = round(sum(rates) / len(rates), 2)
    return avg, (f'Multiple MNT UK rates for {customer_name!r}; ZIP {zip_code!r} not an exact match -- '
                 f'used average of {len(rates)} known rates, confirm with WH.')


def parse_q3_august(path):
    """Column A ("Solaredge rate EUR") only, per this round's request."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    book = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        pl = r[idx['PL nr.']]
        rate = r[idx['Solaredge rate EUR']]
        if pl and isinstance(rate, (int, float)):
            book[pl] = rate
    return book


def has_pod(row):
    pod = row.get('Logistic POD')
    return isinstance(pod, (int, float)) or hasattr(pod, 'year')


# --------------------------------------------------------------------------
# Load source data
# --------------------------------------------------------------------------

def load_rows():
    wb = openpyxl.load_workbook(MAIN_FILE, data_only=True)
    ws = wb['BI 2809']
    headers = [c.value for c in ws[1]]
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(v is None for v in vals):
            continue
        rows.append(dict(zip(headers, vals)))
    return rows


def num_pallets(row):
    v = row.get('# of Pallets per line roundup')
    if not isinstance(v, (int, float)):
        return None
    return max(1, math.ceil(v - 1e-9))


def group_key_for(row):
    sh = row.get('Shipment Number')
    if sh and sh != 'N/A':
        return ('SH', sh)
    return ('ORD', row.get('SE Order#'))


def is_support(row):
    dest_whs = row.get('Dest WHS Code')
    return isinstance(dest_whs, str) and dest_whs.upper().endswith('SUP')


# --------------------------------------------------------------------------
# Classification (population + pricing method), per row
# --------------------------------------------------------------------------

def classify(row):
    whs = row.get('Sending WHS Code')
    dest = row.get('Destination country')
    family = row.get('Family Type')
    forwarder = row.get('Forwarder')
    ship_mode = (row.get('ShipMode') or '').strip().upper()

    if whs == CANOT_WHS:
        if dest == 'Israel':
            return 'Canot - Domestic', 'canot'
        return 'Canot - Export', 'matrix'

    if whs in NL_WHS:
        if row.get('Customer Name') == BAYWA_IT_CUSTOMER and isinstance(row.get('Zip'), str) \
                and 'block' in row['Zip'].lower():
            return 'NL - BayWa IT (manual)', 'baywa_it'
        if family == 'Battery' or dest == 'United Kingdom':
            return 'NL - Battery + UK', 'battery_uk'
        if is_support(row):
            if forwarder in DBSCHENKER_FORWARDERS:
                return 'NL - Support', 'dbs'
            if forwarder in MNT_FORWARDERS:
                return 'NL - Support', 'q3_mnt'
            return 'NL - Support (manual)', 'manual'
        if ship_mode == 'SEA':
            return 'NL - Export (SEA)', 'matrix'
        if forwarder in DBSCHENKER_FORWARDERS:
            return 'NL - Domestic', 'dbs'
        if forwarder in MNT_FORWARDERS:
            return 'NL - Domestic', 'q3_mnt'
        return 'NL - Domestic (manual)', 'manual'

    return 'Other WHS - Export', 'matrix'


# --------------------------------------------------------------------------
# Pricing (grouped by shipment + method, flat rate allocated to the group
# as a whole -- the group becomes exactly one consolidated output row)
# --------------------------------------------------------------------------

def price_all(rows, dbs_book, matrix, mnt_uk_book, q3_book):
    recs = []
    for row in rows:
        population, method = classify(row)
        recs.append({'rows': [row], 'population': population, 'method': method,
                     'route': None, 'cost': None, 'currency': None, 'note': ''})

    buckets = defaultdict(list)
    manual_recs = []
    for rec in recs:
        if rec['method'] == 'manual':
            manual_recs.append(rec)
        else:
            buckets[(rec['method'], group_key_for(rec['rows'][0]))].append(rec)

    priced_recs = []
    for (method, gkey), group in buckets.items():
        group_rows = [r['rows'][0] for r in group]
        consolidated = {'rows': group_rows, 'population': group[0]['population'], 'method': method,
                        'route': None, 'cost': None, 'currency': None, 'note': ''}
        total_pallets = sum((num_pallets(r) or 0) for r in group_rows)
        sample = group_rows[0]

        if method == 'dbs':
            zone, price, note = dbs_lookup(dbs_book, sample.get('Destination country code'),
                                            sample.get('Zip'), total_pallets or 1)
            consolidated['route'] = zone
            consolidated['cost'] = price
            consolidated['currency'] = 'EUR' if price is not None else None
            consolidated['note'] = note or f'DBS Price list 2026, zone {zone}: {price} EUR for {total_pallets:g} total pallets.'

        elif method == 'q3_mnt':
            ship_num = sample.get('Shipment Number')
            is_pod = has_pod(sample)
            rate = q3_book.get(ship_num) if is_pod else None
            if rate is not None:
                consolidated['route'] = f'Q3 AUGUST (PL nr. {ship_num})'
                consolidated['cost'] = rate
                consolidated['currency'] = 'EUR'
                consolidated['note'] = f'Q3 AUGUST rate {rate} EUR for the whole shipment.'
            else:
                consolidated['note'] = (
                    f"MNT forwarder, but {'no real Logistic POD yet' if not is_pod else f'Shipment Number {ship_num!r} not on file in Q3 AUGUST'} "
                    "-- not priced via DBS this round, confirm with WH contact.")

        elif method == 'matrix':
            org = sample.get('Sending WHS Code')
            hit = matrix_lookup(matrix, org, sample.get('Origin country'), sample.get('ShipMode'),
                                sample.get('Destination country'))
            route = f"{org}/{sample.get('ShipMode')}->{sample.get('Destination country')}"
            consolidated['route'] = route
            if not hit:
                consolidated['note'] = f'No Ship Cost Matrix rate for {route} -- confirm with WH contact.'
            else:
                cost, currency, no_pal, mnote = hit
                consolidated['cost'] = cost
                consolidated['currency'] = currency
                note = f'Ship Cost Matrix flat rate {cost} {currency} for the whole shipment ({total_pallets:g} pallets).'
                if isinstance(no_pal, (int, float)) and total_pallets > no_pal:
                    note += f' Shipment total exceeds this corridor\'s {no_pal:.0f}-pallet rate -- verify capacity.'
                if mnote:
                    note += ' ' + mnote
                consolidated['note'] = note

        elif method == 'canot':
            price, currency, note = canot_israel_lookup(sample.get('City'), total_pallets or 1)
            consolidated['route'] = 'Canot domestic (IL)'
            consolidated['cost'] = price
            consolidated['currency'] = currency
            consolidated['note'] = note

        elif method == 'battery_uk':
            ship_num = sample.get('Shipment Number')
            customer = sample.get('Customer Name')
            zip_code = sample.get('Zip')
            _, by_name = mnt_uk_book
            if has_pod(sample) and ship_num in q3_book:
                consolidated['route'] = f'Q3 AUGUST (PL nr. {ship_num})'
                consolidated['cost'] = q3_book[ship_num]
                consolidated['currency'] = 'EUR'
                consolidated['note'] = f'Q3 AUGUST rate {q3_book[ship_num]} EUR for the whole shipment.'
            elif customer in by_name:
                price, note = mnt_uk_lookup(mnt_uk_book, customer, zip_code)
                consolidated['route'] = 'MNT UK (Price Q3)'
                consolidated['cost'] = price
                consolidated['currency'] = 'EUR' if price is not None else None
                consolidated['note'] = note or f'MNT UK full-truck rate {price} EUR for the whole shipment.'
            else:
                consolidated['note'] = (f"Customer {customer!r} not in MNT UK price list, and no POD/Q3 AUGUST "
                                          f"match for Shipment Number {ship_num!r} -- this population is priced "
                                          "only from those two sources, confirm with WH contact.")

        elif method == 'baywa_it':
            capped_total = max(1, math.ceil(total_pallets - 1e-9)) if total_pallets else 1
            price = BAYWA_IT_OVERRIDES.get(capped_total)
            consolidated['route'] = 'BayWa IT (manual)'
            consolidated['cost'] = price
            consolidated['currency'] = 'EUR' if price is not None else None
            if price is None:
                consolidated['note'] = (f'No BayWa IT rate on file or interpolatable for {capped_total} total '
                                         f'pallets -- confirm with WH contact.')
            elif capped_total in BAYWA_IT_INTERPOLATED:
                consolidated['note'] = (f'Estimated BayWa IT rate {price} EUR for {capped_total} total pallets '
                                         f'-- interpolated linearly between the known 2-pallet (411 EUR) and '
                                         f'12-pallet (2100 EUR) rates; not a directly quoted price, confirm with WH.')
            else:
                consolidated['note'] = f'User-supplied BayWa IT rate {price} EUR for {capped_total} total pallets.'

        priced_recs.append(consolidated)

    return priced_recs, manual_recs


# --------------------------------------------------------------------------
# Output workbook
# --------------------------------------------------------------------------

def _joined(rows, field):
    vals = [r.get(field) for r in rows]
    distinct = sorted(set(v for v in vals if v is not None), key=str)
    if len(distinct) <= 1:
        return distinct[0] if distinct else None
    return '; '.join(str(v) for v in distinct)


def _earliest_date(rows):
    dates = [r.get('Date') for r in rows if r.get('Date') is not None]
    return min(dates) if dates else None


OUT_COLUMNS = [
    ('Source', lambda rec: _joined(rec['rows'], 'Source')),
    ('Date', lambda rec: _earliest_date(rec['rows'])),
    ('Shipment Number', lambda rec: _joined(rec['rows'], 'Shipment Number')),
    ('SE Order#', lambda rec: _joined(rec['rows'], 'SE Order#')),
    ('Sending WHS Code', lambda rec: _joined(rec['rows'], 'Sending WHS Code')),
    ('Dest WHS Code', lambda rec: _joined(rec['rows'], 'Dest WHS Code')),
    ('Forwarder', lambda rec: _joined(rec['rows'], 'Forwarder')),
    ('ShipMode', lambda rec: _joined(rec['rows'], 'ShipMode')),
    ('Customer Name', lambda rec: _joined(rec['rows'], 'Customer Name')),
    ('Destination country', lambda rec: _joined(rec['rows'], 'Destination country')),
    ('ZIP', lambda rec: _joined(rec['rows'], 'Zip')),
    ('Family Type', lambda rec: _joined(rec['rows'], 'Family Type')),
    ('# of Pallets (total)', lambda rec: sum((num_pallets(r) or 0) for r in rec['rows'])),
    ('Lines', lambda rec: len(rec['rows'])),
    ('Population', lambda rec: rec['population']),
    ('Method', lambda rec: METHOD_LABELS.get(rec['method'], rec['method'])),
    ('Zone / Route', lambda rec: rec['route']),
    ('Cost', lambda rec: rec['cost']),
    ('Currency', lambda rec: rec['currency']),
    ('Cost (USD)', lambda rec: to_usd(rec['cost'], rec['currency'])),
    ('Note', lambda rec: rec['note']),
]

MANUAL_COLUMNS = [
    ('Source', lambda rec: rec['row'].get('Source')),
    ('Date', lambda rec: rec['row'].get('Date')),
    ('Shipment Number', lambda rec: rec['row'].get('Shipment Number')),
    ('SE Order#', lambda rec: rec['row'].get('SE Order#')),
    ('Sending WHS Code', lambda rec: rec['row'].get('Sending WHS Code')),
    ('Dest WHS Code', lambda rec: rec['row'].get('Dest WHS Code')),
    ('Forwarder', lambda rec: rec['row'].get('Forwarder')),
    ('ShipMode', lambda rec: rec['row'].get('ShipMode')),
    ('Customer Name', lambda rec: rec['row'].get('Customer Name')),
    ('Destination country', lambda rec: rec['row'].get('Destination country')),
    ('ZIP', lambda rec: rec['row'].get('Zip')),
    ('Family Type', lambda rec: rec['row'].get('Family Type')),
    ('# of Pallets', lambda rec: num_pallets(rec['row'])),
    ('Population', lambda rec: rec['population']),
    ('Cost', lambda rec: None),
    ('Currency', lambda rec: None),
    ('Cost (USD)', lambda rec: None),
    ('Note', lambda rec: f"Forwarder {rec['row'].get('Forwarder')!r} is neither DBSCHENKER nor MNT -- price manually."),
]

METHOD_LABELS = {
    'dbs': 'DBS Price list 2026', 'matrix': 'Ship Cost Matrix', 'canot': 'Canot rate card',
    'battery_uk': 'MNT UK / Q3 AUGUST', 'baywa_it': 'BayWa IT (manual)', 'q3_mnt': 'Q3 AUGUST (MNT, POD)',
}


def write_sheet(wb, title, records, columns):
    ws = wb.create_sheet(title)
    headers = [c[0] for c in columns]
    ws.append(headers)
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold
    for rec in records:
        ws.append([getter(rec) for _, getter in columns])
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).font = arial
    for c, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(14, min(40, len(header) + 4))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    ws.freeze_panes = 'A2'
    return ws


def write_summary(wb, shipped_recs, backlog_recs, manual_recs):
    ws = wb.create_sheet('Summary', 1)
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    headers = ['Population', 'Status', 'Lines (consolidated rows)', 'Priced', 'Unpriced', 'Total Cost (USD)']
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold

    populations = sorted(set(r['population'] for r in shipped_recs + backlog_recs))
    r_out = 2
    for pop in populations:
        for status, recs in (('Shipped', shipped_recs), ('Backlog', backlog_recs)):
            subset = [r for r in recs if r['population'] == pop]
            if not subset:
                continue
            priced = [r for r in subset if r['cost'] is not None]
            total_usd = sum(to_usd(r['cost'], r['currency']) or 0 for r in priced)
            ws.append([pop, status, len(subset), len(priced), len(subset) - len(priced), round(total_usd, 2)])
            for c in range(1, len(headers) + 1):
                ws.cell(row=r_out, column=c).font = arial
            ws.cell(row=r_out, column=6).number_format = '#,##0.00'
            r_out += 1

    ws.append(['Manual Review (not priced)', 'Mixed', len(manual_recs), 0, len(manual_recs), 0])
    for c in range(1, len(headers) + 1):
        ws.cell(row=r_out, column=c).font = arial
    r_out += 1

    total_row = r_out
    ws.cell(row=total_row, column=1, value='Total').font = bold
    ws.cell(row=total_row, column=2).font = bold
    for c in (3, 4, 5, 6):
        col = get_column_letter(c)
        cell = ws.cell(row=total_row, column=c, value=f'=SUM({col}2:{col}{total_row - 1})')
        cell.font = bold
        if c == 6:
            cell.number_format = '#,##0.00'
    for c, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(24, len(h) + 4)


def write_readme(wb, stats):
    readme = wb.create_sheet('README', 0)
    readme.column_dimensions['A'].width = 105
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    lines = [
        ('BI 2809 Freight Cost Report', bold),
        ('', arial),
        (f"Source: data/BI_2809.xlsx, single flat \"BI 2809\" sheet ({stats['total_source_rows']} lines: "
         f"{stats['shipped_source_rows']} Shipped, {stats['backlog_source_rows']} Backlog).", arial),
        ('', arial),
        ('Populations and rules:', bold),
        ('  Canot (3PLCANOT): Destination=Israel -> Canot Whs. NIS rate card by city region. Everything else -> '
         'Ship Cost Matrix (SEA).', arial),
        ('  Netherlands (3PLDBSNL/3PLDBSBRNL), checked in order: (1) Family Type=Battery or Destination=UK -> '
         'priced ONLY from Q3 AUGUST (POD + Shipment Number match) or MNT UK price list (Customer/ZIP match) -- '
         'no DBS/matrix fallback; (2) Dest WHS Code ends "SUP" -> same forwarder rule as (3); (3) Forwarder='
         'DBSCHENKER -> DBS Price list 2026 (zone+pallets); Forwarder=MNT -> Q3 AUGUST (POD + Shipment Number '
         'match), no match -> flagged, not DBS; any other forwarder -> "Manual Review" sheet, not priced here; '
         '(4) ShipMode=SEA -> Ship Cost Matrix instead of the forwarder split. BayWa r.e. Solar Systems srl '
         '(Italy) "3c/4c block" ZIP lines use the fixed EUR-by-pallet override table regardless of forwarder '
         '(1->240, 2->411, 12->2100, 13->1410 EUR, confirmed against the forwarded price-list screenshot; '
         '4/6/8 pallets have no rate on file and are linearly interpolated between the 2- and 12-pallet rates, '
         'flagged as estimated in the Note -- confirm with WH if a real quote exists).', arial),
        ('  Any other Sending WHS Code -> Ship Cost Matrix.', arial),
        ('', arial),
        ('Consolidation: every rate source above gives ONE flat price per shipment/order, not per line. Per this '
         'round\'s request, the Shipped/Backlog sheets do NOT list one row per source line -- lines are grouped '
         'by Shipment Number (or SE Order# when there is none yet) within each pricing method, and each group is '
         'written as ONE consolidated row: pallets summed, the group\'s full cost (not split), a "Lines" column '
         'showing how many source lines were merged, and any field that genuinely varies within the group (Family '
         'Type, SE Order#) joined with "; ". Customer Name and Destination country are always uniform within a '
         'Shipment Number group in this data pull, so are never ambiguous. '
         f"{stats['total_source_rows']} source lines consolidated into {stats['total_output_rows']} rows.", arial),
        ('', arial),
        ('"Manual Review" sheet: NL lines whose Forwarder is neither DBSCHENKER nor MNT -- not priced by this '
         'report at all, per the request, for the user to price by hand. Kept as one row per source line (not '
         'consolidated), since no cost computation groups them.', arial),
        ('', arial),
        ('FX to USD: 1 EUR = 1.1392, 1 ILS = 0.32559 (2026-09-28 snapshot, xe.com/tradingeconomics.com -- not a '
         'contracted rate); Ship Cost Matrix costs are already USD.', arial),
        ('', arial),
        ('Rows with a blank Cost need a human check (see their Note) -- no DBS zone/coverage, no Ship Cost Matrix '
         'corridor, no MNT UK/Q3 AUGUST match, no BayWa override rate for that shipment\'s total pallets, or an '
         'MNT line with no POD yet.', bold),
    ]
    for text, font in lines:
        readme.append((text,))
        readme.cell(row=readme.max_row, column=1).font = font


def main():
    print("Loading price sources...")
    dbs_book = parse_dbs_price_book(DBS_FILE)
    matrix = parse_ship_matrix(MATRIX_FILE)
    mnt_uk_book = parse_mnt_uk(MNT_UK_FILE)
    q3_book = parse_q3_august(Q3_AUG_FILE)

    print("Loading BI 2809 data...")
    rows = load_rows()
    print(f"Loaded {len(rows)} rows.")
    shipped_source = sum(1 for r in rows if r.get('Source') == 'Shipped')
    backlog_source = sum(1 for r in rows if r.get('Source') == 'Backlog')

    print("Pricing and consolidating...")
    priced_recs, manual_recs = price_all(rows, dbs_book, matrix, mnt_uk_book, q3_book)

    shipped_recs = [r for r in priced_recs if _joined(r['rows'], 'Source') == 'Shipped']
    backlog_recs = [r for r in priced_recs if _joined(r['rows'], 'Source') == 'Backlog']
    shipped_manual = [r for r in manual_recs if r['rows'][0].get('Source') == 'Shipped']
    backlog_manual = [r for r in manual_recs if r['rows'][0].get('Source') == 'Backlog']

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_sheet(wb, 'Shipped', shipped_recs, OUT_COLUMNS)
    write_sheet(wb, 'Backlog', backlog_recs, OUT_COLUMNS)
    manual_flat = [{'row': r['rows'][0], 'population': r['population']} for r in manual_recs]
    write_sheet(wb, 'Manual Review', manual_flat, MANUAL_COLUMNS)
    write_summary(wb, shipped_recs, backlog_recs, manual_recs)

    stats = {
        'total_source_rows': len(rows), 'shipped_source_rows': shipped_source,
        'backlog_source_rows': backlog_source,
        'total_output_rows': len(shipped_recs) + len(backlog_recs) + len(manual_recs),
    }
    write_readme(wb, stats)

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)
    print("Saved", OUT_FILE)

    all_priced = shipped_recs + backlog_recs
    n_priced = sum(1 for r in all_priced if r['cost'] is not None)
    total_usd = sum(to_usd(r['cost'], r['currency']) or 0 for r in all_priced if r['cost'] is not None)
    print(f"Source lines: {len(rows)} -> consolidated rows: {len(all_priced)} (+{len(manual_recs)} manual)")
    print(f"Priced: {n_priced}, unpriced: {len(all_priced) - n_priced}")
    print(f"Grand total: {total_usd:,.2f} USD")


if __name__ == '__main__':
    main()
