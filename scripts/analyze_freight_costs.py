"""
Costs every Backlog and Shipped line in data/Freight_costs_3108.xlsx (sheet
"DB B2B") against four different rate sources, per the routing rules below,
and writes output/Freight_Cost_Analysis.xlsx.

============================== BACKLOG rows ==================================
("Source" column = "Backlog", 793 rows)

Routing (checked in this order for every line):
 1. A/I == "UPS"                          -> out of scope (UPS courier tariff,
                                              not covered by any provided file).
 2. Family Type == "Battery"
    OR Destination country == "United Kingdom"
                                          -> data/MNT_UK_price_list.xlsx,
                                             sheet "Price Q3", matched by
                                             Customer Name (+ ship-to ZIP when
                                             the customer has more than one
                                             listed address).
 3. Sending WHS Code in {3PLDBSNL, 3PLDBSBRNL} (the two NL warehouses)
                                          -> data/DBS_Price_list_2026.xlsx,
                                             by destination zone (country code
                                             + first 2 digits of ZIP, or the
                                             UK postcode area) and pallet
                                             count, same logic validated in
                                             the DB_NL analysis (DE, 4 pallets,
                                             zone DE04 = 474 EUR).
 4. Any other sending warehouse (3PLCANOT / 3PLDBSTW)
                                          -> data/Ship_Cost_Matrix.xlsx,
                                             matched by Sending Org Code +
                                             Ship Mode (as recorded on the
                                             line) + destination country. The
                                             matrix gives one flat rate per
                                             corridor (a truck/container
                                             rate), so it is computed once per
                                             SE Order# + destination and
                                             allocated across that order's
                                             lines pro-rata by pallet share.

Three views are written for Backlog:
 - "Backlog - Full"                 all 793 lines, whichever rule matched.
 - "Backlog - NL to EU (Customers)" rule-3 subset: NL warehouses, Parent
                                     Region = EUROPE, A/I != UPS (per the
                                     explicit instruction to leave UPS out).
 - "Backlog - Battery & UK (MNT)"   rule-2 subset.

============================== SHIPPED rows ===================================
("Source" column = "Shipped", 1676 rows)

Routing:
 1. Forwarder == "MNT" AND has a real Logistic POD date
                                          -> data/Q3_prices_AUGUST.xlsx,
                                             matched by Shipment Number = "PL
                                             nr." (column G). Cost = Solaredge
                                             rate + Waiting fee + Cancellation
                                             fee + Customs Clearance Fee (all
                                             EUR); Duty is in GBP and reported
                                             separately since it is a
                                             different currency. Computed once
                                             per Shipment Number and allocated
                                             across that shipment's lines
                                             pro-rata by pallet share.
 2. Everything else (not MNT+POD, or MNT+POD but no matching PL nr in the
    August file)                         -> data/DBS_Price_list_2026.xlsx, by
                                             ZIP + pallet count (same as
                                             Backlog rule 3). Destinations
                                             outside the DBS country sheets
                                             (Israel, Thailand, Japan) have no
                                             match and are flagged.

One view is written for Shipped: "Shipped - Full".

Every row that could not be priced, or that needed a judgement call (an
Edmundson UK address not in the MNT price list, an order that exceeds the
matched Ship Cost Matrix corridor's pallet capacity, a MNT/POD shipment
without an August rate, a destination the DBS book doesn't cover), is kept
in the output with cost blank and a Note explaining why -- these are exactly
the lines to check with the warehouse contact.
"""
import os
import re
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = lambda name: os.path.join(BASE_DIR, "data", name)
OUT_FILE = os.path.join(BASE_DIR, "output", "Freight_Cost_Analysis.xlsx")

FREIGHT_FILE = DATA("Freight_costs_3108.xlsx")
DBS_FILE = DATA("DBS_Price_list_2026.xlsx")
MATRIX_FILE = DATA("Ship_Cost_Matrix.xlsx")
MNT_UK_FILE = DATA("MNT_UK_price_list.xlsx")
Q3_AUG_FILE = DATA("Q3_prices_AUGUST.xlsx")

MAX_EP = 33  # DBS: last pallet bracket = "Full load"
NL_WHS = {"3PLDBSNL", "3PLDBSBRNL"}


# --------------------------------------------------------------------------
# 1. DBS Price list 2026 (Backlog rule 3 / Shipped rule 2)
# --------------------------------------------------------------------------

def parse_dbs_price_book(path):
    """{country_code: {zone_code: {pallet_count: price_eur}}}"""
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


def dbs_price_for_pallets(zone_prices, pallets):
    if not isinstance(pallets, (int, float)):
        return None
    ep = max(1, min(MAX_EP, int(round(pallets + 0.4999))))
    while ep <= MAX_EP:
        if ep in zone_prices:
            return zone_prices[ep]
        ep += 1
    return None


def dbs_lookup(dbs_book, country_code, zip_code, pallets):
    zone = dbs_zone_for(country_code, zip_code)
    if not zone:
        return None, None, 'ZIP could not be parsed into a rate zone'
    zone_prices = dbs_book.get((country_code or '').strip().upper(), {}).get(zone)
    if not zone_prices:
        return zone, None, f'No DBS rate sheet/zone for country={country_code!r} zone={zone!r}'
    if not isinstance(pallets, (int, float)):
        return zone, None, f'Invalid pallet count on this line ({pallets!r}) — cannot look up a rate bracket'
    price = dbs_price_for_pallets(zone_prices, pallets)
    if price is None:
        return zone, None, f'Zone {zone} has no rate bracket for {pallets} pallets'
    return zone, price, None


# --------------------------------------------------------------------------
# 2. Ship Cost Matrix (Backlog rule 4)
# --------------------------------------------------------------------------

def parse_ship_matrix(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Ship Cost Matrix']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    by_org = {}   # (org_code, ship_mode, dest_country) -> (cost, currency, no_pallets)
    by_country = {}  # (sending_country, ship_mode, dest_country) -> same, first match wins
    for r in rows:
        org = r[idx['Sending Org Code**']]
        country = r[idx['Sending Country Name']]
        mode = (r[idx['Ship Mode*']] or '').strip().upper()
        dest = r[idx['Dest Country Name*']]
        cost = r[idx['Cost*']]
        currency = r[idx['Currency']]
        no_pal = r[idx['No. Pallets']]
        if not isinstance(cost, (int, float)) or not dest or not mode:
            continue
        key_country = (country, mode, dest)
        by_country.setdefault(key_country, (cost, currency, no_pal))
        if org:
            key_org = (org, mode, dest)
            by_org[key_org] = (cost, currency, no_pal)
    return by_org, by_country


def matrix_lookup(matrix, org_code, sending_country, ship_mode, dest_country):
    by_org, by_country = matrix
    mode = (ship_mode or '').strip().upper()
    hit = by_org.get((org_code, mode, dest_country))
    if hit:
        return hit
    hit = by_country.get((sending_country, mode, dest_country))
    if hit:
        return hit
    return None


# --------------------------------------------------------------------------
# 3. MNT UK price list, sheet "Price Q3" (Backlog rule 2)
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
    avg = sum(rates) / len(rates)
    return avg, (f'Ship-to ZIP {zip_code!r} not among {customer_name!r}\'s listed addresses; '
                 f'used average of {len(rates)} known rates for this customer — confirm with WH.')


# --------------------------------------------------------------------------
# 4. Q3 prices AUGUST (Shipped rule 1)
# --------------------------------------------------------------------------

def parse_q3_august(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1']
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    book = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        pl = r[idx['PL nr.']]
        if not pl:
            continue
        rate = r[idx['Solaredge rate EUR']] or 0
        waiting = r[idx['Waiting fee EUR']] or 0
        cancel = r[idx['Cancellation fee EUR']] or 0
        customs = r[idx['Customs Clerance Fee EUR']] or 0
        duty_gbp = r[idx['Duty GBP']]
        nr_pal = r[idx['Nr. of pal']]
        total_eur = rate + waiting + cancel + customs
        book[pl] = {'total_eur': total_eur, 'duty_gbp': duty_gbp, 'nr_pal': nr_pal}
    return book


# --------------------------------------------------------------------------
# Main data load
# --------------------------------------------------------------------------

def load_freight_rows():
    wb = openpyxl.load_workbook(FREIGHT_FILE, data_only=True)
    ws = wb['DB B2B']
    headers = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers)}
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(v is None for v in vals):
            continue
        rows.append(dict(zip(headers, vals)))
    return rows


def has_pod(row):
    pod = row.get('Logistic POD')
    return isinstance(pod, (int, float)) or hasattr(pod, 'year')


def num_pallets(row):
    """Numeric pallet count for a line, or None for bad data (#N/A, #DIV/0!, ...)."""
    v = row.get('# of Pallets per line')
    return v if isinstance(v, (int, float)) else None


# --------------------------------------------------------------------------
# Backlog pricing
# --------------------------------------------------------------------------

def price_backlog(rows, dbs_book, matrix, mnt_book):
    out = []
    # First pass: everything except the Ship Cost Matrix (order-allocated) rows.
    matrix_bucket = []  # (row, group_key)

    for row in rows:
        whs = row.get('Sending WHS Code')
        ai = (row.get('A/I') or '').strip().upper()
        family_type = row.get('Family Type')
        dest_country = row.get('Destination country')
        pallets = row.get('# of Pallets per line')

        rec = {
            'row': row,
            'method': None,
            'zone_or_key': None,
            'cost': None,
            'currency': None,
            'note': '',
        }

        if ai == 'UPS':
            rec['method'] = 'UPS (out of scope)'
            rec['note'] = 'Priced via UPS courier tariff — not covered by any provided price file.'
        elif family_type == 'Battery' or dest_country == 'United Kingdom':
            cost, note = mnt_uk_lookup(mnt_book, row.get('Customer Name'), row.get('Zip'))
            rec['method'] = 'MNT UK price list (Q3)'
            rec['cost'] = cost
            rec['currency'] = 'EUR' if cost is not None else None
            rec['note'] = note or ''
            if cost is None:
                rec['note'] = ('Family Type=Battery / destination=UK but no MNT UK rate matched — '
                                'confirm with WH contact. ' + (note or ''))
        elif whs in NL_WHS:
            zone, price, note = dbs_lookup(dbs_book, row.get('Destination country code'), row.get('Zip'), pallets)
            rec['method'] = 'DBS Price list 2026'
            rec['zone_or_key'] = zone
            rec['cost'] = price
            rec['currency'] = 'EUR' if price is not None else None
            rec['note'] = note or ''
        else:
            rec['method'] = 'Ship Cost Matrix (order-allocated)'
            group_key = (row.get('SE Order#'), dest_country)
            matrix_bucket.append((rec, group_key))

        out.append(rec)

    # Second pass: allocate Ship Cost Matrix flat rate per SE Order# + destination.
    groups = defaultdict(list)
    for rec, key in matrix_bucket:
        groups[key].append(rec)

    for (order_no, dest_country), recs in groups.items():
        sample_row = recs[0]['row']
        org_code = sample_row.get('Sending WHS Code')
        sending_country = sample_row.get('Origin country')
        ship_mode = sample_row.get('ShipMode')
        hit = matrix_lookup(matrix, org_code, sending_country, ship_mode, dest_country)

        total_pallets = sum((num_pallets(r['row']) or 0) for r in recs)
        if not hit:
            for r in recs:
                r['note'] = (f'No Ship Cost Matrix rate for {org_code}/{ship_mode}->{dest_country} '
                             '— confirm with WH contact.')
            continue

        cost, currency, no_pal = hit
        r_pal_note = ''
        if isinstance(no_pal, (int, float)) and total_pallets > no_pal:
            r_pal_note = (f' Order total {total_pallets:.1f} pallets exceeds this corridor\'s '
                          f'{no_pal:.0f}-pallet capacity — may need >1 truck, confirm with WH.')
        for r in recs:
            line_pallets = num_pallets(r['row'])
            if line_pallets is None:
                r['note'] = 'Invalid pallet count on this line (formula error in source) — cannot allocate cost.'
                continue
            share = (line_pallets / total_pallets) if total_pallets else (1 / len(recs))
            r['cost'] = round(cost * share, 2)
            r['currency'] = currency
            r['zone_or_key'] = f'{org_code}/{ship_mode}->{dest_country}'
            r['note'] = (f'Order-level flat rate {cost} {currency} allocated pro-rata by pallet share.'
                        + r_pal_note)

    return out


# --------------------------------------------------------------------------
# Shipped pricing
# --------------------------------------------------------------------------

def price_shipped(rows, dbs_book, q3_book):
    out = []
    q3_bucket = []  # (rec, shipment_number)

    for row in rows:
        rec = {'row': row, 'method': None, 'zone_or_key': None, 'cost': None,
               'currency': None, 'duty_gbp': None, 'note': ''}

        is_mnt_pod = row.get('Forwarder') == 'MNT' and has_pod(row)
        ship_num = row.get('Shipment Number')

        if is_mnt_pod and ship_num in q3_book:
            rec['method'] = 'Q3 prices AUGUST (MNT, POD)'
            q3_bucket.append((rec, ship_num))
        else:
            zone, price, note = dbs_lookup(dbs_book, row.get('Destination country code'), row.get('Zip'),
                                            row.get('# of Pallets per line'))
            rec['method'] = ('DBS Price list 2026' if not is_mnt_pod
                              else 'DBS Price list 2026 (MNT/POD, no Aug. rate found)')
            rec['zone_or_key'] = zone
            rec['cost'] = price
            rec['currency'] = 'EUR' if price is not None else None
            rec['note'] = note or ''

        out.append(rec)

    groups = defaultdict(list)
    for rec, ship_num in q3_bucket:
        groups[ship_num].append(rec)

    for ship_num, recs in groups.items():
        info = q3_book[ship_num]
        total_pallets = sum((num_pallets(r['row']) or 0) for r in recs)
        for r in recs:
            line_pallets = num_pallets(r['row'])
            if line_pallets is None:
                r['note'] = 'Invalid pallet count on this line (formula error in source) — cannot allocate cost.'
                continue
            share = (line_pallets / total_pallets) if total_pallets else (1 / len(recs))
            r['cost'] = round(info['total_eur'] * share, 2)
            r['currency'] = 'EUR'
            r['zone_or_key'] = ship_num
            note = f'Shipment-level rate (PL nr. {ship_num}) allocated pro-rata by pallet share.'
            if info['duty_gbp']:
                note += f' Duty {round(info["duty_gbp"] * share, 2)} GBP not included (different currency).'
            r['note'] = note

    return out


# --------------------------------------------------------------------------
# Output workbook
# --------------------------------------------------------------------------

OUT_COLUMNS = [
    ('Source', lambda rec: rec['row'].get('Source')),
    ('SE Order#', lambda rec: rec['row'].get('SE Order#')),
    ('Shipment Number', lambda rec: rec['row'].get('Shipment Number')),
    ('Sending WHS Code', lambda rec: rec['row'].get('Sending WHS Code')),
    ('Origin country', lambda rec: rec['row'].get('Origin country')),
    ('Customer Name', lambda rec: rec['row'].get('Customer Name')),
    ('Destination country', lambda rec: rec['row'].get('Destination country')),
    ('Zip', lambda rec: rec['row'].get('Zip')),
    ('A/I', lambda rec: rec['row'].get('A/I')),
    ('ShipMode', lambda rec: rec['row'].get('ShipMode')),
    ('Forwarder', lambda rec: rec['row'].get('Forwarder')),
    ('Family Type', lambda rec: rec['row'].get('Family Type')),
    ('# of Pallets per line', lambda rec: rec['row'].get('# of Pallets per line')),
    ('Pricing Method', lambda rec: rec['method']),
    ('Zone / Match Key', lambda rec: rec['zone_or_key']),
    ('Cost', lambda rec: rec['cost']),
    ('Currency', lambda rec: rec['currency']),
    ('Note', lambda rec: rec['note']),
]


def write_sheet(wb, title, records):
    ws = wb.create_sheet(title)
    headers = [c[0] for c in OUT_COLUMNS]
    ws.append(headers)
    bold = Font(name='Arial', bold=True)
    arial = Font(name='Arial')
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold
    for rec in records:
        ws.append([getter(rec) for _, getter in OUT_COLUMNS])
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).font = arial
    for c, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(14, min(40, len(header) + 4))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    ws.freeze_panes = 'A2'
    return ws


def write_readme(wb, stats):
    ws = wb.create_sheet('README', 0)
    ws.column_dimensions['A'].width = 100
    bold = Font(name='Arial', bold=True, size=13)
    arial = Font(name='Arial')
    lines = [
        ("Freight Cost Analysis — Backlog & Shipped", bold),
        ("", arial),
        ("Sheets:", Font(name='Arial', bold=True)),
        ("  Backlog - Full                 all Backlog lines, whichever rate source matched", arial),
        ("  Backlog - NL to EU (Customers)  NL warehouses -> Europe, A/I <> UPS", arial),
        ("  Backlog - Battery & UK (MNT)    Family Type = Battery, or destination = United Kingdom", arial),
        ("  Shipped - Full                  all Shipped lines, whichever rate source matched", arial),
        ("", arial),
        ("Routing rules applied (see script docstring for full detail):", Font(name='Arial', bold=True)),
        ("  Backlog: UPS (A/I) -> out of scope  >  Battery/UK -> MNT UK price list Q3  >  "
         "NL warehouses -> DBS Price list 2026 (zone+pallets)  >  other warehouses -> "
         "Ship Cost Matrix (flat rate per SE Order#+destination, allocated by pallet share)", arial),
        ("  Shipped: MNT forwarder + real POD + Shipment Number found as 'PL nr.' in Q3 AUGUST -> "
         "Q3 AUGUST rate (allocated by pallet share)  >  everything else -> DBS Price list 2026 (ZIP+pallets)", arial),
        ("", arial),
        ("Currency: DBS and MNT UK and Q3 AUGUST rates are EUR. Ship Cost Matrix rates are mostly USD "
         "(see the Currency column per row) — no FX rate was supplied, so USD figures were NOT converted.", arial),
        ("", arial),
        ("Rows with a blank Cost need a human check (see their Note) — mostly: no DBS coverage for a "
         "destination outside Europe, no Ship Cost Matrix corridor for that lane, no MNT UK customer "
         "match, or a MNT/POD shipment with no matching rate in the August file.", Font(name='Arial', bold=True)),
        ("", arial),
    ]
    for text, font in lines:
        cell = ws.cell(row=ws.max_row + 1 if ws.max_row > 1 else 1, column=1, value=text)
        cell.font = font
        cell.alignment = Alignment(wrap_text=False)

    ws.append(("", ))
    ws.append(("Row counts", ))
    ws.cell(row=ws.max_row, column=1).font = Font(name='Arial', bold=True)
    for label, value in stats:
        ws.append((label, value))
        ws.cell(row=ws.max_row, column=1).font = arial
        ws.cell(row=ws.max_row, column=2).font = arial


def main():
    print("Loading price sources...")
    dbs_book = parse_dbs_price_book(DBS_FILE)
    matrix = parse_ship_matrix(MATRIX_FILE)
    mnt_book = parse_mnt_uk(MNT_UK_FILE)
    q3_book = parse_q3_august(Q3_AUG_FILE)

    print("Loading freight data...")
    all_rows = load_freight_rows()
    backlog_rows = [r for r in all_rows if r.get('Source') == 'Backlog']
    shipped_rows = [r for r in all_rows if r.get('Source') == 'Shipped']
    print(f"Backlog rows: {len(backlog_rows)}, Shipped rows: {len(shipped_rows)}")

    backlog_recs = price_backlog(backlog_rows, dbs_book, matrix, mnt_book)
    shipped_recs = price_shipped(shipped_rows, dbs_book, q3_book)

    nl_europe = [r for r in backlog_recs
                 if r['row'].get('Sending WHS Code') in NL_WHS
                 and r['row'].get('Parent Region') == 'EUROPE'
                 and (r['row'].get('A/I') or '').strip().upper() != 'UPS']

    battery_uk = [r for r in backlog_recs
                  if r['row'].get('Family Type') == 'Battery'
                  or r['row'].get('Destination country') == 'United Kingdom']

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    write_sheet(wb, 'Backlog - Full', backlog_recs)
    write_sheet(wb, 'Backlog - NL to EU (Customers)', nl_europe)
    write_sheet(wb, 'Backlog - Battery & UK (MNT)', battery_uk)
    write_sheet(wb, 'Shipped - Full', shipped_recs)

    def n_priced(recs):
        return sum(1 for r in recs if r['cost'] is not None)

    stats = [
        ('Backlog - Full: total rows', len(backlog_recs)),
        ('Backlog - Full: priced', n_priced(backlog_recs)),
        ('Backlog - Full: unpriced / flagged', len(backlog_recs) - n_priced(backlog_recs)),
        ('Backlog - NL to EU (Customers): rows', len(nl_europe)),
        ('Backlog - NL to EU (Customers): priced', n_priced(nl_europe)),
        ('Backlog - Battery & UK (MNT): rows', len(battery_uk)),
        ('Backlog - Battery & UK (MNT): priced', n_priced(battery_uk)),
        ('Shipped - Full: total rows', len(shipped_recs)),
        ('Shipped - Full: priced', n_priced(shipped_recs)),
        ('Shipped - Full: unpriced / flagged', len(shipped_recs) - n_priced(shipped_recs)),
    ]
    write_readme(wb, stats)

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    wb.save(OUT_FILE)

    for label, value in stats:
        print(f"{label}: {value}")


if __name__ == '__main__':
    main()
