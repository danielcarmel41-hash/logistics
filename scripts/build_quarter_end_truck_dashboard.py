import re, datetime as dt, warnings
import openpyxl
from openpyxl.utils import get_column_letter as L
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
warnings.filterwarnings('ignore')

UP='/root/.claude/uploads/d8c8fad7-28f6-530c-b158-96560ec8d3e0/'
BL_SRC=UP+'7ce49fae-Backlog_-_NL_251126.xlsx'
REP_SRC=UP+'cd001e7e-REP_364_3PL_Communication_of_the_Shipment_Lines.xlsx'
OUT='/home/user/logistics/output/Quarter_End_Truck_Planning_Dashboard.xlsx'

# ---------- brand / style ----------
BLUE='00082E'; AZURE='2E8FD6'; SKY='78BAE6'; GRAY='696969'; INK='232323'
LIGHT='EEF3FA'; WHITE='FFFFFF'; LINE='D5DCE8'
FONT='Segoe UI'
def f(sz=10,b=False,c=INK,i=False): return Font(name=FONT,size=sz,bold=b,color=c,italic=i)
def fill(c): return PatternFill('solid',start_color=c,end_color=c)
thin=Side(style='thin',color=LINE)
BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
INPUT_FILL=fill('FFF8D6'); INPUT_FONT=Font(name=FONT,size=10,color='0000FF')
RED_F=fill('F7B6B6'); ORG_F=fill('FFD59A'); GRN_F=fill('BFE3C8')

wb=openpyxl.Workbook()
wsD=wb.active; wsD.title='Executive Dashboard'
wsW=wb.create_sheet('Weekly Truck Forecast')
wsC=wb.create_sheet('Country Analysis')
wsF=wb.create_sheet('Forwarder Allocation')
wsT=wb.create_sheet('Shipment Tracking')
wsS=wb.create_sheet('Assumptions & Settings')
wsH=wb.create_sheet('WH Closures')
wsK=wb.create_sheet('Data Checks')
wsR=wb.create_sheet('SE_REP')
wsB=wb.create_sheet('Backlog_NL')
for w in wb.worksheets: w.sheet_view.showGridLines=False
wsR.sheet_view.showGridLines=True; wsB.sheet_view.showGridLines=True

def name(n,ref): wb.defined_names[n]=DefinedName(n,attr_text=ref)

# ---------- raw sheets (copied unchanged) ----------
def copy_raw(src_path,sheet,dst):
    s=openpyxl.load_workbook(src_path)[sheet]
    for row in s.iter_rows():
        for c in row:
            if c.value is None: continue
            d=dst.cell(row=c.row,column=c.column,value=c.value)
            if isinstance(c.value,(dt.datetime,dt.date)): d.number_format='yyyy-mm-dd hh:mm' if isinstance(c.value,dt.datetime) and (c.value.hour or c.value.minute or c.value.second) else 'yyyy-mm-dd'
    for c in dst[1]: c.font=Font(name=FONT,bold=True,color=WHITE,size=9); c.fill=fill(BLUE)
    dst.freeze_panes='A2'
    hdr=[c.value for c in dst[1]]
    return hdr
rep_h=copy_raw(REP_SRC,'Sheet2',wsR)
bl_h=copy_raw(BL_SRC,'Backlog - NL',wsB)
wsR.sheet_properties.tabColor=GRAY; wsB.sheet_properties.tabColor=GRAY
RC=lambda h: L(rep_h.index(h)+1)
BC=lambda h: L(bl_h.index(h)+1)
assert BC('# of Pallets per line - roundup')=='BZ', BC('# of Pallets per line - roundup')
assert RC('Scheduled Ship Date')=='F' and RC('Total Amount after Discount in USD')=='Z'
MAXR=1200; R0=6; R1=R0+MAXR-1          # tracking rows 6..1205 <-> SE_REP rows 2..1201
name('BL_Ord',f"Backlog_NL!${BC('SE Order#')}$2:${BC('SE Order#')}$3000")
name('BL_Line',f"Backlog_NL!${BC('Line#')}$2:${BC('Line#')}$3000")
name('BL_Pal',f"Backlog_NL!$BZ$2:$BZ$3000")
name('BL_TT',f"Backlog_NL!${BC('Transit Days')}$2:${BC('Transit Days')}$3000")
name('BL_Fwd',f"Backlog_NL!${BC('Forwarder')}$2:${BC('Forwarder')}$3000")

# ---------- Settings ----------
S=wsS; SN="'Assumptions & Settings'!"
S.column_dimensions['A'].width=46; S.column_dimensions['B'].width=22; S.column_dimensions['C'].width=62
for col in 'DEFGHI': S.column_dimensions[col].width=22
S['A1']='Assumptions & settings'; S['A1'].font=f(18,True,BLUE)
S['A2']='Yellow cells with blue text are inputs - change them and every sheet recalculates. All other cells are formulas or fixed lists.'; S['A2'].font=f(9,c=GRAY)
def hdr(ws,cell,text,w=None):
    ws[cell]=text; ws[cell].font=f(10,True,WHITE); ws[cell].fill=fill(BLUE); ws[cell].alignment=Alignment(vertical='center')
def sect(ws,row,text,span=3):
    for c in range(1,span+1): ws.cell(row=row,column=c).fill=fill(BLUE)
    ws.cell(row=row,column=1,value=text).font=f(10,True,WHITE)
sect(S,3,'1. Parameters')
params=[
 (4,'Report date override (leave blank = today)',None,'Optional. Enter a date to run the plan "as of" another day.','d',None),
 (5,'Report date used by the workbook','=IF(B4="",TODAY(),INT(B4))','Formula. Drives overdue / approaching flags.','d','ReportDate'),
 (6,'Quarter end',dt.datetime(2026,12,31),'Quarter-end shipping cut-off.','d','QEnd'),
 (7,'Truck capacity - standard (pallets / truck)',33,'GIVEN BY USER: 33 pallets per truck.','n','CapStd'),
 (8,'Truck capacity - DG / CSS Battery (pallets / truck)',9,'GIVEN BY USER: CSS Battery = 9 pallets per truck. DG and Non-DG are never loaded together.','n','CapDG'),
 (9,'TT counting (Business / Calendar)','Business','ASSUMPTION: transit days are working days (Mon-Fri, excl. holidays below).','t','TTMode'),
 (10,'"Approaching" window (days from report date)',7,'ASSUMPTION: pickup due within this many calendar days = "approaching".','n','ApproachDays'),
 (11,'High pallet volume threshold (pallets per line)',20,'ASSUMPTION: a line at/above this needs extra truck capacity.','n','HighPallet'),
 (12,'Long TT: TT as % of remaining delivery window',0.5,'ASSUMPTION: flag when TT uses at least this share of the days left before the customer closes.','p','LongTTPct'),
 (13,'Tight buffer before customer WH closure (days)',2,'ASSUMPTION: arrival within this many days of closure = tight.','n','TightBuf'),
 (14,'MNT dedicated truck per customer + delivery address? (Yes / No)','Yes','CONFIRMED BY USER: MNT runs direct trucks per customer and per delivery address. DBS trucks stay pooled per country / day / DG.','t','MNTperCust'),
 (15,'Batteries to OTHER countries also to MNT? (Yes / No)','Yes','CONFIRMED BY USER: all CSS Batteries go with MNT, whatever the country. Set to No to limit MNT batteries to the country list below.','t','BattOtherMNT'),
 (16,'First week shown in forecast (a Monday)',dt.datetime(2026,11,16),'Forecast shows 12 weeks from this Monday.','d','FcstStart'),
]
for r,lab,val,note,typ,nm in params:
    S.cell(row=r,column=1,value=lab).font=f(10)
    c=S.cell(row=r,column=2,value=val); c.font=f(10); c.border=BOX; c.alignment=Alignment(horizontal='right')
    S.cell(row=r,column=3,value=note).font=f(9,c=GRAY)
    if typ=='d': c.number_format='dd-mmm-yyyy'
    if typ=='p': c.number_format='0%'
    if r!=5: c.fill=INPUT_FILL; c.font=INPUT_FONT
    if nm: name(nm,f"{SN}$B${r}")
name('RepOverride',f"{SN}$B$4")
dv=DataValidation(type='list',formula1='"Business,Calendar"',allow_blank=False); S.add_data_validation(dv); dv.add('B9')
dv=DataValidation(type='list',formula1='"Yes,No"',allow_blank=False); S.add_data_validation(dv); dv.add('B14'); dv.add('B15')

sect(S,19,'2. Rule lists (editable)',9)
heads={'A':'MNT battery countries','C':'MNT - ALL shipments countries','E':'DG product families (Family Type)','G':'Holidays (no pickup / delivery)'}
for col,t in heads.items(): hdr(S,f'{col}20',t)
for col in 'BDFH': S[f'{col}20'].fill=fill(BLUE)
for i,cn in enumerate(['Netherlands','Germany','Poland','United Kingdom']): S[f'A{21+i}']=cn
S['C21']='United Kingdom'
S['E21']='Battery'
for i,d in enumerate([dt.datetime(2026,12,25),dt.datetime(2026,12,26),dt.datetime(2027,1,1)]):
    S[f'G{21+i}']=d; S[f'G{21+i}'].number_format='dd-mmm-yyyy'
for col,n in (('A',6),('C',4),('E',5),('G',10)):
    for i in range(n):
        c=S[f'{col}{21+i}']; c.fill=INPUT_FILL; c.font=INPUT_FONT; c.border=BOX
name('MNTBatt',f"{SN}$A$21:$A$26"); name('MNTAll',f"{SN}$C$21:$C$24")
name('DGFam',f"{SN}$E$21:$E$25"); name('Holidays',f"{SN}$G$21:$G$30")
S['A32']='Rule applied: MNT = all UK shipments + all DG (CSS Battery) shipments (B15 = Yes; if No, only to the countries in the first list). DBS = everything else.'; S['A32'].font=f(9,c=GRAY)
S['A33']='DG assumption: the source files have no DG flag, so Family Type = "Battery" (CSS Battery) is treated as DG; every other family is Non-DG.'; S['A33'].font=f(9,c=GRAY)

sect(S,35,'3. Destination countries and dedicated-truck TT (MNT)',9)
hdr(S,'A36','Country'); hdr(S,'B36','TT dedicated (days)'); hdr(S,'C36','Source / comment'); hdr(S,'I36','Dropdown list')
countries=['Austria','Belgium','France','Germany','Hungary','Ireland','Italy','Japan','Luxembourg','Netherlands','Poland','Romania','Switzerland','United Kingdom']
ttd={'Austria':3,'Belgium':2,'France':4,'Germany':2,'Ireland':5,'Italy':4,'Netherlands':2,'Poland':4,'Switzerland':4,'United Kingdom':4}
CL0=37; CL1=56
S['I37']='All'
for i in range(20):
    r=CL0+i
    a=S.cell(row=r,column=1,value=countries[i] if i<len(countries) else None); a.fill=INPUT_FILL; a.font=INPUT_FONT; a.border=BOX
    b=S.cell(row=r,column=2,value=ttd.get(countries[i]) if i<len(countries) else None); b.fill=INPUT_FILL; b.font=INPUT_FONT; b.border=BOX
    if i<len(countries):
        S.cell(row=r,column=3,value='TT dedicated per user table (Backlog file, sheet "Sheet4")' if countries[i] in ttd else 'No TT dedicated in source - not invented. Needed only if a line to this country is assigned to MNT.').font=f(9,c=GRAY)
    S.cell(row=r+0,column=9,value=f'=IF(A{r}="","",A{r})').font=f(9,c=GRAY)
S['I37'].value='All'
for i in range(20): S.cell(row=CL0+i+1,column=9,value=f'=IF(A{CL0+i}="","",A{CL0+i})')
hdr(S,'K36','Sending warehouses'); hdr(S,'L36','Dropdown list'); S.column_dimensions['K'].width=18; S.column_dimensions['L'].width=14
for i,wn in enumerate(['3PLDBSNL','3PLDBSBRNL','3PLAWLNL']): S.cell(row=37+i,column=11,value=wn)
S['L37']='All'
for i in range(8):
    c=S.cell(row=37+i,column=11); c.fill=INPUT_FILL; c.font=INPUT_FONT; c.border=BOX
    S.cell(row=38+i,column=12,value=f'=IF(K{37+i}="","",K{37+i})')
name('WhsList',f"{SN}$K$37:$K$44")
name('TTCountry',f"{SN}$A${CL0}:$A${CL1}"); name('TTDays',f"{SN}$B${CL0}:$B${CL1}"); name('CountryList',f"{SN}$A${CL0}:$A${CL1}")
S['A58']='Transit days: DBS lines use the "Transit Days" column of Backlog_NL. MNT direct trucks use the TT-dedicated column above (kept separate as requested).'; S['A58'].font=f(9,c=GRAY)

sect(S,60,'4. CONFIRMED truck bookings (input) - everything else in the workbook is an ESTIMATE',9)
for col,t in zip('ABCDEF',['Booking date (any day in the pickup week)','Country','Forwarder (DBS/MNT)','DG / Non-DG','Trucks confirmed','Booking reference / carrier']): hdr(S,f'{col}61',t)
BK0=62; BK1=111
for r in range(BK0,BK1+1):
    for col in 'ABCDEF':
        c=S[f'{col}{r}']; c.fill=INPUT_FILL; c.font=INPUT_FONT; c.border=BOX
    S[f'A{r}'].number_format='dd-mmm-yyyy'
S['A112']='No confirmed bookings were provided in the source files, so confirmed = 0 and all trucks are shown as estimates. Enter bookings here to track "still to book".'; S['A112'].font=f(9,c=GRAY)
name('BkDate',f"{SN}$A${BK0}:$A${BK1}"); name('BkTrucks',f"{SN}$E${BK0}:$E${BK1}")
name('BkCountry',f"{SN}$B${BK0}:$B${BK1}"); name('BkFwd',f"{SN}$C${BK0}:$C${BK1}"); name('BkDG',f"{SN}$D${BK0}:$D${BK1}")

sect(S,114,'5. Notes and interpretations to confirm',9)
notes=[
 'Source model: SE REP (REP_364) is the master list - one row per order line. Pallets come from Backlog_NL column BZ "# of Pallets per line - roundup", joined on SE Order# + Line#. Lines not found there take pallets = ROUNDUP(quantity / units per pallet) from SE REP itself (same logic as BZ; column BN shows the source). Lines with neither remain without pallets and are flagged red.',
 'Pallets are the sum of per-line round-ups (as instructed). Because every line is rounded up on its own, pallet and truck totals can be higher than the physical count.',
 'Forwarder rule (confirmed): all UK -> MNT, all CSS Battery (DG) -> MNT in every country, everything else -> DBS. MNT trucks are direct, one per customer + delivery address; DBS trucks are pooled per country / day / DG.',
 'Planning date: Scheduled Ship Date (col F of SE REP) is the shipping-document close date. Its 00:59:59 time stamps are end-of-previous-day (they match the Backlog "Date" column), so the workbook uses date minus 1 hour. Weekend / holiday dates move back to the previous working day.',
 'Pickup cannot be brought forward: pickup date = document close date. When pickup + TT lands inside a customer warehouse closure, the line is flagged RED ("TT not synced") and the workbook shows the latest pickup that would still meet the closure, so the team can ask for the document close date to be changed.',
 'Customer closures come from "Close WH - Christmas 2026" (sheet Summary 2026). Customers can have several windows; the workbook checks every window, so a customer that reopens between two closures is handled. "Final WH close in Q" = start of the latest window that begins on/before quarter end.',
 'Trucks = ROUNDUP(pallets / capacity) per group of Country + Forwarder + DG/Non-DG + pickup date. Only LAND lines are trucked; SEA/AIR lines are listed but need no truck.',
 'Risk colours: RED = missing pallets/TT/date, pickup date already passed, wrong forwarder, or TT not synced with document close + customer WH closure. ORANGE = pickup within the approaching window, high pallet volume, long TT vs window, tight buffer, missing value / closure data, non-standard forwarder. GREEN = none of these.',
 'Weekly Booking Plan tab: one row per delivery stop (pickup date + sending warehouse + country + customer + address + cargo type + forwarder), sorted by date, filterable by week / warehouse / country / forwarder. Trucks are loaded per sending warehouse.',
 'Capacity: the Tracking sheet is pre-built for 1,200 order lines (SE_REP rows 2-1201) and Backlog_NL rows 2-3000. Paste a refreshed SE REP export over SE_REP and a refreshed Backlog over Backlog_NL (same columns) and everything updates.',
]
for i,t in enumerate(notes):
    c=S.cell(row=115+i,column=1,value=f'{i+1}. {t}'); c.font=f(9); c.alignment=Alignment(wrap_text=True,vertical='top')
    S.merge_cells(start_row=115+i,start_column=1,end_row=115+i,end_column=7); S.row_dimensions[115+i].height=40
S.freeze_panes='A3'
S.sheet_properties.tabColor=AZURE

# ---------- WH closures ----------
H=wsH
H['A1']='Customer warehouse closures (Christmas 2026)'; H['A1'].font=f(18,True,BLUE)
H['A2']='Source: "Close WH - Christmas 2026.xlsx" (Summary 2026, customer comments), parsed into date windows. Edit / add rows (names must match SE REP "Customer Name"). Windows must be in chronological order; closure end = last closed day.'; H['A2'].font=f(9,c=GRAY)
D=lambda d,m,y=2026: dt.datetime(y,m,d)
def w(sd,sm,ed,em): return (D(sd,sm,2026 if sm>6 else 2027),D(ed,em,2026 if em>6 else 2027))
clos={
'BayWa r.e. Solar Energy Systems GmbH':('Germany','WH close: 18/12-7/1',[w(18,12,7,1)],''),
'BayWa r.e. Solar Systems SAS':('France','WH close: 11/11, 23-31/12, 1-3/1',[w(23,12,3,1)],'11/11 window already past - excluded.'),
'BayWa r.e. Solar Systems srl':('Italy','WH close: 08/12, 16/12-6/1',[w(8,12,8,12),w(16,12,6,1)],''),
'BayWa r.e.Solar Systems Sarl':('Belgium','WH close: 18/12-3/1',[w(18,12,3,1)],''),
'CCL Components Ltd':('United Kingdom','WH close: 24/12-4/01',[w(24,12,4,1)],''),
'CCL Solar BV':('Netherlands','WH close: 24/12 (midday) - 3/1',[w(24,12,3,1)],'Midday closure treated as a full closed day (conservative).'),
'Coenergia S.r.l':('Italy','WH close: 23/12 - 6/1',[w(23,12,6,1)],''),
'ESAVING SRL':('Italy','WH close: 23/12-6/1',[w(23,12,6,1)],''),
'Ecostal Yomatec':('France','WH close: 18-31/12, 1-4/1',[w(18,12,4,1)],'Merged contiguous ranges.'),
'Edmundson Electrical Ltd':('United Kingdom','WH close: 24/12/26 - 04/01/27',[w(24,12,4,1)],''),
'Emil Loeffelhardt GmbH & Co. KG':('Germany','WH close: 24/12-1/1',[w(24,12,1,1)],'Some SE REP lines say 24/12-3/1 - Close WH file used.'),
'Energy3000 Solar GmbH':('Austria','Muellendorf: WH close 11/12-06/1 | Ennsdorf: WH close 21/12-06/1',[w(21,12,6,1)],'Ennsdorf site (all current lines). Other site: see Müllendorf row (key customer|city).'),
'Energy3000 Solar GmbH|Müllendorf':('Austria','Muellendorf: WH close 11/12-06/1',[w(11,12,6,1)],'Site key = customer|Ship-To city. Reopens 7/1. Used only if a line ships to a Ship-To city named Müllendorf.'),
'KENO Sp. z o.o.':('Poland','WH close: 29-31/12, 1-11/1',[w(29,12,11,1)],''),
'Koninklijke Oosterberg B.V.':('Netherlands','WH close Apeldoorn: 18(midday), 24(midday), 25, 31(midday), 01/01',[w(18,12,18,12),w(24,12,25,12),w(31,12,1,1)],'Midday closures treated as full days.'),
'Krannich Solar AG':('Switzerland','WH close: 18/12-4/1',[w(18,12,4,1)],''),
'Krannich Solar GmbH & Co. KG':('Germany','WH close: 18/12-11/1',[w(18,12,11,1)],''),
'Krannich Solar S.A.S.':('France','WH close: 17-31/12, 1-4/1',[w(17,12,4,1)],''),
'KuehnSolar GmbH':('Germany','WH close: 21/12-11/1',[w(21,12,11,1)],''),
'Memodo GmbH':('Germany','WH close: 21/12-3/01',[w(21,12,3,1)],'SE REP also lists 24-26/12; 31/12-1/1 for some lines - Close WH file used.'),
'Natec Sunergy BV':('Netherlands','WH close: 14/12-4/1',[w(14,12,4,1)],''),
'Oekoteam Solar Photovoltaikverbund GmbH':('Austria','WH close 17/12-17/1',[w(17,12,17,1)],''),
'PMGREEN S.P.A.':('Italy','WH close: 24/12-6/1',[w(24,12,6,1)],''),
'REXEL FRANCE':('France','WH close: 25/12, 1/1',[w(25,12,25,12),w(1,1,1,1)],''),
'Rexel Belgium':('Belgium','WH close: 25-27/12, 1-3/1',[w(25,12,27,12),w(1,1,3,1)],''),
'Rexel Nederland B.V':('Netherlands','WH close: 24/12-3/1',[w(24,12,3,1)],''),
'SAM HYDRO-SOLAR':('France','WH close: 24-27/12, 31/12, 1-3/1',[w(24,12,27,12),w(31,12,3,1)],''),
'SONEPAR ITALIA S.P.A.':('Italy','WH close: 18/12-6/1',[w(18,12,6,1)],''),
'Segen':('United Kingdom','WH close: 24(midday)/12-04/01',[w(24,12,4,1)],'Midday closure treated as a full closed day.'),
'Solarmarkt GmbH':('Switzerland','WH close: 19/12-3/1',[w(19,12,3,1)],''),
'Solen Energy Group':('United Kingdom','WH close: 19/12-3/1',[w(19,12,3,1)],''),
'Solen Energy IE':('Ireland','WH close: 19/12-3/1',[w(19,12,3,1)],'SE REP says 19/12-4/1 - Close WH file used.'),
'Solexis SA':('Switzerland','WH close: 23/12-3/1',[w(23,12,3,1)],''),
'VP Solar srl':('Italy','WH close: 16/12-11/1',[w(16,12,11,1)],''),
'Vaillant GmbH':('Germany','WH close: 23/12-27/12; 30/12-01/01',[w(23,12,27,12),w(30,12,1,1)],''),
'Fankhauser Solar AG':('Switzerland','WH close: 23/12-27/12; 31/12-03/01',[w(23,12,27,12),w(31,12,3,1)],''),
}
for i,t in enumerate(['Customer (SE REP name) or customer|Ship-To city','Country','Source','Source comment','Start 1','Start 2','Start 3','Start 4','End 1','End 2','End 3','End 4','Note']):
    c=H.cell(row=4,column=i+1,value=t); c.font=f(10,True,WHITE); c.fill=fill(BLUE)
for i,(cu,(co,txt,wins,note)) in enumerate(clos.items()):
    r=5+i
    H.cell(row=r,column=1,value=cu); H.cell(row=r,column=2,value=co); H.cell(row=r,column=3,value='Close WH file'); H.cell(row=r,column=4,value=txt)
    for k,(s_,e_) in enumerate(wins):
        H.cell(row=r,column=5+k,value=s_); H.cell(row=r,column=9+k,value=e_)
    H.cell(row=r,column=13,value=note)
for r in range(5,105):
    for c in range(1,14):
        cell=H.cell(row=r,column=c); cell.font=INPUT_FONT if c<=12 else f(9,c=GRAY); cell.border=BOX
        if 5<=c<=12: cell.number_format='dd-mmm-yy'
for col,wd in zip('ABCDEFGHIJKLM',[40,16,13,52,11,11,11,11,11,11,11,11,70]): H.column_dimensions[col].width=wd
H.freeze_panes='B5'; H.sheet_properties.tabColor=AZURE
name('WH_Cust',"'WH Closures'!$A$5:$A$104"); name('WH_Start',"'WH Closures'!$E$5:$H$104"); name('WH_End',"'WH Closures'!$I$5:$L$104")

# ---------- Shipment Tracking ----------
T=wsT
cols=[]  # (key, header, width, numfmt, formula)
def add(key,header,width,fmt,formula): cols.append((key,header,width,fmt,formula))
R=lambda h: f"SE_REP!${RC(h)}{{rr}}"
add('A','Line ID',16,None,'=IF([AF]=0,"",'+R('Line Number')+')')
add('B','Order #',11,None,'=IF([AF]=0,"",'+R('Order Number')+'&"")')
add('C','Customer',30,None,'=IF([AF]=0,"",'+R('Customer Name')+'&"")')
add('D','Country',15,None,'=IF([AF]=0,"",'+R('Country')+'&"")')
add('E','Family type',17,None,'=IF([AF]=0,"",'+R('Family Type')+'&"")')
add('F','Mode',7,None,'=IF([AF]=0,"",'+R('Mode Of Transport')+'&"")')
add('G','Pallets',8,'0','=IF([AF]=0,"",IF([AI]>0,INDEX(BL_Pal,[AI]),IFERROR(ROUNDUP('+R('Quantity')+'/VALUE('+R('Qty of pallet on Sea pallet')+'),0),"")))')
add('H','Shipment value (USD)',14,'#,##0','=IF([AF]=0,"",IF(ISNUMBER('+R('Total Amount after Discount in USD')+'),'+R('Total Amount after Discount in USD')+',""))')
add('I','Doc close date',12,'dd-mmm-yy','=IF([AF]=0,"",IF(ISNUMBER('+R('Scheduled Ship Date')+'),WORKDAY(INT('+R('Scheduled Ship Date')+'-1/24)+1,-1,Holidays),""))')
add('J','DG / Non-DG',9,None,'=IF([AF]=0,"",IF(AND([E]<>"",COUNTIF(DGFam,[E])>0),"DG","Non-DG"))')
add('K','Assigned forwarder (rule)',11,None,'=IF([AF]=0,"",IF(AND([D]<>"",COUNTIF(MNTAll,[D])>0),"MNT",IF([J]="DG",IF(OR(AND([D]<>"",COUNTIF(MNTBatt,[D])>0),BattOtherMNT="Yes"),"MNT","DBS"),"DBS")))')
add('L','Current forwarder (in system)',12,None,'=IF([AF]=0,"",IF(AND('+R('FWD Name')+'<>"DEFAULT",'+R('FWD Name')+'<>""),SUBSTITUTE(UPPER('+R('FWD Name')+'),"DBSCHENKER","DBS"),IF([AI]>0,IF(AND(INDEX(BL_Fwd,[AI])<>"DEFAULT",INDEX(BL_Fwd,[AI])<>""),SUBSTITUTE(UPPER(INDEX(BL_Fwd,[AI])&""),"DBSCHENKER","DBS"),"Unassigned"),"Unassigned")))')
add('M','Forwarder check',17,None,'=IF([AF]=0,"",IF([L]="Unassigned","Not yet assigned",IF([L]=[K],"OK",IF(OR([L]="DBS",[L]="MNT"),"WRONG FORWARDER","Non-standard forwarder"))))')
add('N','TT (days)',7,'0','=IF([AF]=0,"",IF([K]="MNT",IFERROR(1/(1/INDEX(TTDays,MATCH([D],TTCountry,0))),""),IF([AI]=0,"",IFERROR(1/(1/VALUE(INDEX(BL_TT,[AI]))),""))))')
add('O','TT basis',16,None,'=IF([AF]=0,"",IF([K]="MNT","MNT dedicated TT","DBS transit days"))')
add('P','Final WH close in Q',12,'dd-mmm-yy','=IF(OR([AF]=0,[AK]=0),"",IFERROR(1/(1/SUMPRODUCT(MAX(INDEX(WH_Start,[AJ],0)*(INDEX(WH_Start,[AJ],0)<=QEnd)))),""))')
add('Q','Binding WH closure start',12,'dd-mmm-yy','=IF(OR([AF]=0,[AL]=0),"",IF([AL]>[AK],"",INDEX(WH_Start,[AJ],[AL])))')
add('R','Arrival if shipped at doc close',12,'dd-mmm-yy','=IF(OR([AF]=0,[I]="",[N]=""),"",IF(TTMode="Business",WORKDAY([I],[N],Holidays),[I]+[N]))')
add('S','Pickup date (= doc close)',12,'dd-mmm-yy','=IF(OR([AF]=0,[I]=""),"",[I])')
add('T','Days to pickup',8,'0','=IF([S]="","",[S]-ReportDate)')
add('U','Days arrival vs last open day (neg = in closure)',12,'0','=IF(OR([Q]="",[R]=""),"",[Q]-1-[R])')
add('V','Truck capacity',8,'0','=IF([AF]=0,"",IF([J]="DG",CapDG,CapStd))')
add('W','Truck group key',22,None,'=IF(OR([AF]=0,[F]<>"LAND",NOT(ISNUMBER([G])),[S]=""),"",IF([G]<=0,"",[BE]&"|"&[D]&"|"&[K]&"|"&[J]&"|"&[S]&IF(AND(MNTperCust="Yes",[K]="MNT"),"|"&LEFT([C],40)&"|"&[BA],"")))')
add('X','Group pallets',8,'0','=IF([W]="","",SUMIFS([G:],[W:],[W]))')
add('Y','Group trucks (est.)',9,'0','=IF([W]="","",ROUNDUP([X]/[V],0))')
add('Z','Trucks counted once',9,'0','=IF([W]="","",IF(COUNTIF($W$'+str(R0)+':[W],[W])=1,[Y],0))')
add('AA','Truck allocation (share)',10,'0.00','=IF([W]="","",[Y]*[G]/[X])')
add('AB','Pickup week (Mon)',11,'dd-mmm-yy','=IF([S]="","",[S]-WEEKDAY([S],2)+1)')
add('AC','Planning status',24,None,'=IF([AF]=0,"",IF(OR([G]="",[N]="",[I]=""),"Incomplete data",IF([F]<>"LAND","Non-truck mode ("&[F]&")",IF([S]<ReportDate,"Overdue - ship now",IF([BC]=1,"TT not synced with WH closure",IF([S]-ReportDate<=ApproachDays,"Pickup due soon","Planned"))))))')
add('AD','Risk reasons',60,None,'=IF([AF]=0,"",IF([G]="","No pallet data; ","")&IF([N]="","No TT; ","")&IF([I]="","No doc close date; ","")&IF([E]="","No family type; ","")&IF([H]="","No value; ","")&IF([AK]=0,"No WH closure data; ","")&IF([M]="WRONG FORWARDER","Wrong forwarder (rule: "&[K]&"); ","")&IF([M]="Non-standard forwarder","Non-standard forwarder; ","")&IF([AM]=1,"Pickup date passed; ","")&IF([BC]=1,"TT NOT SYNCED: pickup "&TEXT([I],"dd-mmm")&" + "&[N]&"d TT = arrival "&TEXT([R],"dd-mmm")&", inside customer WH closure from "&TEXT([Q],"dd-mmm")&" (latest pickup to fit: "&TEXT([BD],"dd-mmm")&"); ","")&IF([AN]=1,"Pickup within "&ApproachDays&" days; ","")&IF([AO]=1,"High pallet volume; ","")&IF([AP]=1,"Long TT vs delivery window; ","")&IF([AQ]=1,"Tight WH buffer; ",""))')
add('AE','Risk indicator',10,None,'=IF([AF]=0,"",IF(OR([G]="",[N]="",[I]="",[AM]=1,[BC]=1,[M]="WRONG FORWARDER"),"Red",IF(OR([AN]=1,[AO]=1,[AP]=1,[AQ]=1,[H]="",[AK]=0,[E]="",[M]="Non-standard forwarder"),"Orange","Green")))')
# helpers
add('AF','Active (1/0)',7,'0','=IF('+R('Line Number')+'="",0,1)')
add('AG','Line suffix',9,None,'=IF([AF]=0,"",SUBSTITUTE([A],[B]&"_",""))')
add('AH','Backlog row (exact)',9,'0','=IF([AF]=0,"",SUMPRODUCT((BL_Ord&""=[B])*(BL_Line&""=[AG])*(ROW(BL_Ord)-1)))')
add('AI','Backlog row used',9,'0','=IF([AF]=0,0,IF([AH]>0,[AH],IF(RIGHT([AG],2)="-1",SUMPRODUCT((BL_Ord&""=[B])*(BL_Line&""=LEFT([AG],LEN([AG])-2))*(ROW(BL_Ord)-1)),0)))')
add('AJ','WH sheet row',8,'0','=IF([AF]=0,"",IFERROR(MATCH([C]&"|"&[BB],WH_Cust,0),IFERROR(MATCH([C],WH_Cust,0),0)))')
add('AK','# closure windows',8,'0','=IF(OR([AF]=0,[AJ]=0),0,COUNT(INDEX(WH_End,[AJ],0)))')
add('AL','Binding window #',8,'0','=IF(OR([R]="",[AK]=0),0,COUNTIF(INDEX(WH_End,[AJ],0),"<"&[R])+1)')
add('AM','Overdue flag',7,'0','=IF([AF]=0,0,IF(AND(ISNUMBER([S]),[S]<ReportDate),1,0))')
add('AN','Approaching flag',7,'0','=IF([AF]=0,0,IF(AND(ISNUMBER([S]),[S]>=ReportDate,[S]-ReportDate<=ApproachDays),1,0))')
add('AO','High pallet flag',7,'0','=IF([AF]=0,0,IF(AND(ISNUMBER([G]),[G]>=HighPallet),1,0))')
add('AP','Long TT flag',7,'0','=IF([AF]=0,0,IF(OR([Q]="",[N]=""),0,IF(ReportDate<[Q]-1,IF(IF(TTMode="Business",[N]*7/5,[N])>=LongTTPct*([Q]-1-ReportDate),1,0),0)))')
add('AQ','Tight buffer flag',7,'0','=IF([AF]=0,0,IF(AND([U]<>"",[U]>=0,[U]<=TightBuf),1,0))')
add('AR','In dashboard filter',8,'0','=IF([AF]=0,0,IF(AND(OR(SelCountry="All",[D]=SelCountry),OR(SelFwd="All",[K]=SelFwd),OR(SelDG="All",[J]=SelDG),OR(SelFrom="",AND(ISNUMBER([S]),[S]>=SelFrom)),OR(SelTo="",AND(ISNUMBER([S]),[S]<=SelTo))),1,0))')
add('AS','Row id',6,'0','=ROW()')
add('AT','Sort key',9,'0','=IF([S]="",99999,[S])')
add('AU','Urgent rank',7,'0','=IF(AND([AR]=1,[AE]="Red"),COUNTIFS([AR:],1,[AE:],"Red",[AT:],"<"&[AT])+COUNTIFS([AR:],1,[AE:],"Red",[AT:],[AT],[AS:],"<"&[AS])+1,"")')
add('AV','Fwd exception rank',7,'0','=IF(OR([M]="WRONG FORWARDER",[M]="Non-standard forwarder"),COUNTIFS([M:],"WRONG FORWARDER",[AS:],"<"&[AS])+COUNTIFS([M:],"Non-standard forwarder",[AS:],"<"&[AS])+1,"")')
add('AW','Missing pallet flag',7,'0','=IF(AND([AF]=1,[G]=""),1,0)')
add('AX','Missing pallet rank',7,'0','=IF([AW]=1,COUNTIFS([AW:],1,[AS:],"<"&[AS])+1,"")')
add('AY','No-closure customer (first line)',8,'0','=IF(AND([AF]=1,[AK]=0),IF(COUNTIFS($C$'+str(R0)+':[C],[C],$AK$'+str(R0)+':[AK],0)=1,1,0),0)')
add('AZ','No-closure customer rank',7,'0','=IF([AY]=1,COUNTIFS([AY:],1,[AS:],"<"&[AS])+1,"")')
add('BA','Ship-to zip + city',16,None,'=IF([AF]=0,"",LEFT('+R('Ship To Zip')+'&" "&'+R('Ship To City')+',40))')
add('BB','Ship-to city',14,None,'=IF([AF]=0,"",'+R('Ship To City')+'&"")')
add('BC','TT-WH sync conflict flag',8,'0','=IF([AF]=0,0,IF(AND([Q]<>"",[R]<>""),IF([R]>=[Q],1,0),0))')
add('BE','Sending warehouse',12,None,'=IF([AF]=0,"",'+R('Sending Whs')+'&"")')
add('BF','Delivery address',46,None,'=IF([AF]=0,"",TRIM('+R('Ship To Address')+'&""))')
add('BG','Stop key',20,None,'=IF(OR([AF]=0,[S]=""),"",TEXT([S],"yyyymmdd")&"|"&[BE]&"|"&[D]&"|"&LEFT([C],30)&"|"&LEFT([BF],30)&"|"&[J]&"|"&[K])')
add('BH','Stop first line',7,'0','=IF([BG]="",0,IF(COUNTIF($BG$'+str(R0)+':[BG],[BG])=1,1,0))')
add('BI','Risk rank (3/2/1)',7,'0','=IF([AE]="Red",3,IF([AE]="Orange",2,IF([AE]="Green",1,0)))')
add('BJ','In booking-plan filter',7,'0','=IF(OR([AF]=0,[S]=""),0,IF(AND(OR(PlanWeek="All",[AB]=PlanWeek),OR(PlanWhs="All",[BE]=PlanWhs),OR(PlanCountry="All",[D]=PlanCountry),OR(PlanFwd="All",[K]=PlanFwd)),1,0))')
add('BK','Stop rank',7,'0','=IF(AND([BH]=1,[BJ]=1),COUNTIFS([BH:],1,[BJ:],1,[BG:],"<"&[BG])+1,"")')
add('BL','Binding WH closure end',11,'dd-mmm-yy','=IF(OR([AF]=0,[Q]=""),"",INDEX(WH_End,[AJ],[AL]))')
add('BN','Pallet source',22,None,'=IF([AF]=0,"",IF([AI]>0,"Backlog col BZ",IF([G]<>"","SE REP qty / units per pallet","MISSING")))')
add('BM','Stop+risk key',20,None,'=IF([BG]="","",[BG]&"|"&[BI])')
add('BD','Latest pickup that meets WH closure',12,'dd-mmm-yy','=IF(OR([AF]=0,[Q]="",[N]=""),"",WORKDAY(IF(TTMode="Business",WORKDAY(WORKDAY([Q],-1,Holidays),-[N],Holidays),WORKDAY([Q],-1,Holidays)-[N])+1,-1,Holidays))')
from openpyxl.utils import column_index_from_string as CI
keys=[c[0] for c in cols]
# order columns by letter
cols.sort(key=lambda c: CI(c[0]))
def tr(formula,r):
    rr=r-4
    def sub(m):
        k=m.group(1)
        if k.endswith(':'): return f"${k[:-1]}${R0}:${k[:-1]}${R1}"
        return f"${k}{r}"
    out=re.sub(r'\[([A-Z]+:?)\]',sub,formula)
    return out.replace('{rr}',str(rr))
# title
T['A1']='Shipment tracking - order-line level'; T['A1'].font=f(18,True,BLUE)
T['A2']='One row per SE REP order line. Pallets from Backlog_NL (col BZ). Columns A-AE and BE-BF are the planning view; the other columns are helpers (grouped - expand with [+]).'; T['A2'].font=f(9,c=GRAY)
T['A3']='="Report date: "&TEXT(ReportDate,"dd-mmm-yyyy")&"   |   Lines: "&SUM(AF'+str(R0)+':AF'+str(R1)+')&"   |   RED: "&COUNTIF(AE'+str(R0)+':AE'+str(R1)+',"Red")&"   ORANGE: "&COUNTIF(AE'+str(R0)+':AE'+str(R1)+',"Orange")&"   GREEN: "&COUNTIF(AE'+str(R0)+':AE'+str(R1)+',"Green")'
T['A3'].font=f(10,True,BLUE)
HR=5
for key,header,width,fmt,formula in cols:
    ci=CI(key); c=T.cell(row=HR,column=ci,value=header)
    c.font=f(9,True,WHITE); c.fill=fill(BLUE if ci<=CI('AE') else GRAY); c.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center'); c.border=BOX
    T.column_dimensions[key].width=width
    for r in range(R0,R1+1):
        cell=T.cell(row=r,column=ci,value=tr(formula,r)); cell.font=f(9)
        if fmt: cell.number_format=fmt
T.row_dimensions[HR].height=42
T.freeze_panes=f'D{R0}'
T.auto_filter.ref=f"A{HR}:{L(CI('AE'))}{R1}"
T.column_dimensions.group('AF','BD',hidden=True,outline_level=1)
T.column_dimensions.group('BG','BM',hidden=True,outline_level=1)
# conditional formats
rng=f"AE{R0}:AE{R1}"
T.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"Red"'],fill=RED_F,font=Font(name=FONT,bold=True,color='7A0000')))
T.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"Orange"'],fill=ORG_F,font=Font(name=FONT,bold=True,color='7A4A00')))
T.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"Green"'],fill=GRN_F,font=Font(name=FONT,bold=True,color='1D5A33')))
T.conditional_formatting.add(f"AC{R0}:AC{R1}",FormulaRule(formula=[f'OR(LEFT($AC{R0},7)="Overdue",$AC{R0}="Incomplete data",LEFT($AC{R0},10)="TT not syn")'],fill=RED_F))
T.conditional_formatting.add(f"AC{R0}:AC{R1}",FormulaRule(formula=[f'$AC{R0}="Pickup due soon"'],fill=ORG_F))
T.conditional_formatting.add(f"AC{R0}:AC{R1}",FormulaRule(formula=[f'$AC{R0}="Planned"'],fill=GRN_F))
T.conditional_formatting.add(f"M{R0}:M{R1}",FormulaRule(formula=[f'$M{R0}="WRONG FORWARDER"'],fill=RED_F))
T.conditional_formatting.add(f"M{R0}:M{R1}",FormulaRule(formula=[f'$M{R0}="Non-standard forwarder"'],fill=ORG_F))
T.conditional_formatting.add(f"G{R0}:G{R1}",FormulaRule(formula=[f'AND($AF{R0}=1,$G{R0}="")'],fill=RED_F))
T.conditional_formatting.add(f"T{R0}:T{R1}",FormulaRule(formula=[f'AND(ISNUMBER($T{R0}),$T{R0}<0)'],fill=RED_F))
T.conditional_formatting.add(f"T{R0}:T{R1}",FormulaRule(formula=[f'AND(ISNUMBER($T{R0}),$T{R0}>=0,$T{R0}<=ApproachDays)'],fill=ORG_F))
T.sheet_properties.tabColor=BLUE
def TR(k): return f"'Shipment Tracking'!${k}${R0}:${k}${R1}"

# ---------- Dashboard ----------
D_=wsD
for i,wd in enumerate([2]+[13]*12+[2]): D_.column_dimensions[L(i+1)].width=wd
for r in range(1,3):
    for c in range(1,15): D_.cell(row=r,column=c).fill=fill(BLUE)
D_['B1']='Quarter-end truck planning & shipment tracking'; D_['B1'].font=Font(name=FONT,size=20,bold=True,color=WHITE)
D_['B2']='="Executive dashboard  |  report date "&TEXT(ReportDate,"dd-mmm-yyyy")&"  |  quarter end "&TEXT(QEnd,"dd-mmm-yyyy")&"  |  truck capacity "&CapStd&" pallets (DG / CSS Battery "&CapDG&")"'
D_['B2'].font=Font(name=FONT,size=10,color=SKY)
D_.row_dimensions[1].height=34; D_.row_dimensions[2].height=20
# overline
for c in range(2,3): D_.cell(row=3,column=c).fill=fill('FF3232')
D_.row_dimensions[3].height=4
labels=[('C','Country'),('E','Forwarder'),('G','DG / Non-DG'),('I','Pickup date from'),('K','Pickup date to')]
D_['B4']='FILTERS'; D_['B4'].font=f(9,True,GRAY)
for col,t in labels:
    D_[f'{col}4']=t; D_[f'{col}4'].font=f(9,True,GRAY)
    D_.merge_cells(f'{col}5:{L(CI(col)+1)}5')
    c=D_[f'{col}5']; c.fill=INPUT_FILL; c.font=Font(name=FONT,size=11,bold=True,color='0000FF'); c.border=BOX; c.alignment=Alignment(horizontal='center')
    D_[f'{L(CI(col)+1)}5'].border=BOX; D_[f'{L(CI(col)+1)}5'].fill=INPUT_FILL
D_['C5']='All'; D_['E5']='All'; D_['G5']='All'
for col in 'IK': D_[f'{col}5'].number_format='dd-mmm-yyyy'
D_['M4']='Clear dates = all'; D_['M4'].font=f(8,c=GRAY)
D_.row_dimensions[5].height=22
d1=DataValidation(type='list',formula1=f"={SN}$I$37:$I$57",allow_blank=False); d2=DataValidation(type='list',formula1='"All,DBS,MNT"',allow_blank=False); d3=DataValidation(type='list',formula1='"All,DG,Non-DG"',allow_blank=False)
d4=DataValidation(type='date',operator='greaterThan',formula1='36526',allow_blank=True)
for d in (d1,d2,d3,d4): D_.add_data_validation(d)
d1.add('C5'); d2.add('E5'); d3.add('G5'); d4.add('I5'); d4.add('K5')
for n,cell in (('SelCountry','C5'),('SelFwd','E5'),('SelDG','G5'),('SelFrom','I5'),('SelTo','K5')): name(n,f"'Executive Dashboard'!${cell[0]}${cell[1:]}")
INF=TR('AR')
def card(r,col,label,formula,fmt,sub,accent=AZURE):
    c2=L(CI(col)+1)
    D_.merge_cells(f'{col}{r}:{c2}{r}'); D_.merge_cells(f'{col}{r+1}:{c2}{r+2}'); D_.merge_cells(f'{col}{r+3}:{c2}{r+3}')
    for rr_ in range(r,r+4):
        for cc in (col,c2):
            D_[f'{cc}{rr_}'].fill=fill(WHITE)
    for cc in (col,c2): D_[f'{cc}{r}'].border=Border(top=Side(style='thick',color=accent))
    D_[f'{col}{r}']=label; D_[f'{col}{r}'].font=f(9,True,GRAY); D_[f'{col}{r}'].alignment=Alignment(horizontal='left',indent=1)
    v=D_[f'{col}{r+1}']; v.value=formula; v.font=Font(name=FONT,size=22,bold=True,color=BLUE); v.number_format=fmt; v.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    s=D_[f'{col}{r+3}']; s.value=sub; s.font=f(8,c=GRAY); s.alignment=Alignment(horizontal='left',indent=1)
cnt=lambda extra='': f'COUNTIFS({INF},1{extra})'
card(7,'B','SHIPMENTS (ORDER LINES)',f'=COUNTIFS({INF},1)','#,##0','="from "&SUMPRODUCT(('+TR('AF')+'=1)/COUNTIF('+TR('B')+','+TR('B')+'&""))&" orders (all data)"')
card(7,'D','TOTAL PALLETS',f'=SUMIFS({TR("G")},{INF},1)','#,##0','="DG "&TEXT(SUMIFS('+TR("G")+','+INF+',1,'+TR("J")+',"DG"),"#,##0")&"  |  Non-DG "&TEXT(SUMIFS('+TR("G")+','+INF+',1,'+TR("J")+',"Non-DG"),"#,##0")')
card(7,'F','SHIPMENT VALUE (USD)',f'=SUMIFS({TR("H")},{INF},1)','$#,##0.0,,"M"','="avg "&TEXT(IFERROR(SUMIFS('+TR("H")+','+INF+',1)/COUNTIFS('+INF+',1,'+TR("H")+',">0"),0),"$#,##0")&" per line"')
card(7,'H','EST. TRUCKS REQUIRED',f'=SUMIFS({TR("Z")},{INF},1)','#,##0','="ESTIMATE  |  confirmed booked: "&TEXT(SUMPRODUCT(--(BkTrucks<>""),BkTrucks),"0")',accent='FF3232')
card(7,'J','TRUCKS - DBS',f'=SUMIFS({TR("Z")},{INF},1,{TR("K")},"DBS")','#,##0','="DBS pallets "&TEXT(SUMIFS('+TR("G")+','+INF+',1,'+TR("K")+',"DBS"),"#,##0")')
card(7,'L','TRUCKS - MNT',f'=SUMIFS({TR("Z")},{INF},1,{TR("K")},"MNT")','#,##0','="MNT pallets "&TEXT(SUMIFS('+TR("G")+','+INF+',1,'+TR("K")+',"MNT"),"#,##0")')
bkw=lambda: ('SUMPRODUCT((BkDate<>"")*(IF(SelFrom="",0,SelFrom-6)<=BkDate)*(BkDate<=IF(SelTo="",99999,SelTo))*(BkTrucks))')
card(12,'B','APPROACHING PICKUP',f'=COUNTIFS({INF},1,{TR("AN")},1)','#,##0','="pickup within "&ApproachDays&" days of report date"',accent='F59E0B')
card(12,'D','TT NOT SYNCED (WH CLOSE)',f'=COUNTIFS({INF},1,{TR("BC")},1)','#,##0','="arrival in closure | date passed: "&COUNTIFS('+INF+',1,'+TR("AM")+',1)',accent='FF3232')
card(12,'F','CRITICAL (RED)',f'=COUNTIFS({INF},1,{TR("AE")},"Red")','#,##0','="pallets "&TEXT(SUMIFS('+TR("G")+','+INF+',1,'+TR("AE")+',"Red"),"#,##0")',accent='FF3232')
card(12,'H','ATTENTION (ORANGE)',f'=COUNTIFS({INF},1,{TR("AE")},"Orange")','#,##0','="pallets "&TEXT(SUMIFS('+TR("G")+','+INF+',1,'+TR("AE")+',"Orange"),"#,##0")',accent='F59E0B')
card(12,'J','ON TRACK (GREEN)',f'=COUNTIFS({INF},1,{TR("AE")},"Green")','#,##0','="pallets "&TEXT(SUMIFS('+TR("G")+','+INF+',1,'+TR("AE")+',"Green"),"#,##0")',accent='22A06B')
card(12,'L','TRUCKS STILL TO BOOK',f'=MAX(0,H8-SUMPRODUCT(--(BkTrucks<>""),BkTrucks))','#,##0','="estimated minus confirmed bookings"',accent=BLUE)
for r in (8,13): D_.row_dimensions[r].height=20; D_.row_dimensions[r+1].height=20
D_['B17']='Charts follow the filters above. Estimated trucks = ROUNDUP(pallets / capacity) per country + forwarder + DG/Non-DG + pickup day. Confirmed bookings are entered on the Assumptions & Settings tab.'; D_['B17'].font=f(8,c=GRAY,i=False)
# chart data tables (rows 70+)
DT=72
sect_row=DT-2
D_.cell(row=sect_row,column=2,value='Chart data (filtered)').font=f(10,True,BLUE)
# pallets by country raw + sorted
hdrs=['Country','Pallets','Rank','','Rank','Country (sorted)','Pallets (sorted)']
for i,t in enumerate(['Country','Pallets','Rank']): 
    c=D_.cell(row=DT,column=2+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
for i,t in enumerate(['#','Country (sorted)','Pallets']):
    c=D_.cell(row=DT,column=6+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
for i in range(20):
    r=DT+1+i
    D_[f'B{r}']=f'=IF({SN}A{CL0+i}="","",{SN}A{CL0+i})'
    D_[f'C{r}']=f'=IF(B{r}="","",SUMIFS({TR("G")},{INF},1,{TR("D")},B{r}))'
    D_[f'D{r}']=f'=IF(B{r}="","",COUNTIF($C${DT+1}:$C${DT+20},">"&C{r})+COUNTIFS($C${DT+1}:C{r},C{r}))'
    D_[f'F{r}']=i+1
    D_[f'G{r}']=f'=IFERROR(INDEX($B${DT+1}:$B${DT+20},MATCH(F{r},$D${DT+1}:$D${DT+20},0)),"")'
    D_[f'H{r}']=f'=IFERROR(INDEX($C${DT+1}:$C${DT+20},MATCH(F{r},$D${DT+1}:$D${DT+20},0)),"")'
    for col in 'BCDFGH': D_[f'{col}{r}'].font=f(9)
# DG/NonDG, fwd, risk, weekly
def small(r0,c0,title_cols,rows):
    for i,t in enumerate(title_cols):
        c=D_.cell(row=r0,column=c0+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
    for j,row in enumerate(rows):
        for i,v in enumerate(row):
            c=D_.cell(row=r0+1+j,column=c0+i,value=v); c.font=f(9)
small(DT,10,['DG split','Pallets','Trucks'],[
 ['DG',f'=SUMIFS({TR("G")},{INF},1,{TR("J")},"DG")',f'=SUMIFS({TR("Z")},{INF},1,{TR("J")},"DG")'],
 ['Non-DG',f'=SUMIFS({TR("G")},{INF},1,{TR("J")},"Non-DG")',f'=SUMIFS({TR("Z")},{INF},1,{TR("J")},"Non-DG")']])
small(DT+4,10,['Forwarder','Shipments','Pallets','Trucks'],[
 [fw,f'=COUNTIFS({INF},1,{TR("K")},"{fw}")',f'=SUMIFS({TR("G")},{INF},1,{TR("K")},"{fw}")',f'=SUMIFS({TR("Z")},{INF},1,{TR("K")},"{fw}")'] for fw in ('DBS','MNT')])
small(DT+8,10,['Risk','Shipments'],[[k,f'=COUNTIFS({INF},1,{TR("AE")},"{k}")'] for k in ('Red','Orange','Green')])
WK0=DT+13
small(WK0,10,['Pickup week','DBS trucks','MNT trucks','Pallets'],[
 [f'=FcstStart+7*{i}',f'=SUMIFS({TR("Z")},{INF},1,{TR("AB")},J{WK0+1+i},{TR("K")},"DBS")',f'=SUMIFS({TR("Z")},{INF},1,{TR("AB")},J{WK0+1+i},{TR("K")},"MNT")',f'=SUMIFS({TR("G")},{INF},1,{TR("AB")},J{WK0+1+i})'] for i in range(12)])
for i in range(12): D_[f'J{WK0+1+i}'].number_format='dd-mmm'
# charts
def style_bar(ch,title,w,h):
    ch.title=title; ch.width=w; ch.height=h; ch.roundedCorners=False
    if ch.legend is not None: ch.legend.overlay=False
    return ch
ch1=BarChart(); ch1.type='bar'; style_bar(ch1,'Pallets by destination country',16.8,9)
ch1.add_data(Reference(D_,min_col=8,min_row=DT,max_row=DT+14),titles_from_data=True); ch1.set_categories(Reference(D_,min_col=7,min_row=DT+1,max_row=DT+14))
ch1.y_axis.delete=False; ch1.x_axis.delete=False; ch1.x_axis.scaling.orientation='maxMin'; ch1.legend=None
ch1.series[0].graphicalProperties.solidFill=AZURE; ch1.dataLabels=DataLabelList(); ch1.dataLabels.showVal=True
for a in ('showSerName','showCatName','showLegendKey','showPercent'): setattr(ch1.dataLabels,a,False)
ch1.y_axis.majorGridlines=None
D_.add_chart(ch1,'B19')
ch2=BarChart(); ch2.type='col'; ch2.grouping='stacked'; ch2.overlap=100; style_bar(ch2,'Estimated trucks per pickup week',16.8,9)
ch2.add_data(Reference(D_,min_col=11,max_col=12,min_row=WK0,max_row=WK0+12),titles_from_data=True); ch2.set_categories(Reference(D_,min_col=10,min_row=WK0+1,max_row=WK0+12))
ch2.series[0].graphicalProperties.solidFill=BLUE; ch2.series[1].graphicalProperties.solidFill=SKY
ch2.x_axis.number_format='dd-mmm'; ch2.x_axis.delete=False; ch2.y_axis.delete=False; ch2.legend.position='b'
D_.add_chart(ch2,'H19')
def doughnut(title,cats,vals,anchor,colors):
    ch=DoughnutChart(); ch.title=title; ch.width=11; ch.height=8.5; ch.roundedCorners=False; ch.holeSize=55
    ch.add_data(vals,titles_from_data=True); ch.set_categories(cats)
    for i,c in enumerate(colors):
        pt=DataPoint(idx=i); pt.graphicalProperties.solidFill=c; ch.series[0].dPt.append(pt)
    ch.dataLabels=DataLabelList(); ch.dataLabels.showPercent=True
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph,ParagraphProperties,CharacterProperties
    cp=CharacterProperties(sz=1000,b=True,solidFill='FFFFFF')
    ch.dataLabels.txPr=RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp),endParaRPr=cp)])
    for a in ('showVal','showSerName','showCatName','showLegendKey'): setattr(ch.dataLabels,a,False)
    ch.legend.position='b'; ch.legend.overlay=False
    D_.add_chart(ch,anchor)
doughnut('DG vs Non-DG pallets',Reference(D_,min_col=10,min_row=DT+1,max_row=DT+2),Reference(D_,min_col=11,min_row=DT,max_row=DT+2),'B38',[AZURE,BLUE])
doughnut('DBS vs MNT shipments',Reference(D_,min_col=10,min_row=DT+5,max_row=DT+6),Reference(D_,min_col=11,min_row=DT+4,max_row=DT+6),'F38',[BLUE,SKY])
ch5=BarChart(); ch5.type='col'; style_bar(ch5,'Risk indicator',11,8.5)
ch5.add_data(Reference(D_,min_col=11,min_row=DT+8,max_row=DT+11),titles_from_data=True); ch5.set_categories(Reference(D_,min_col=10,min_row=DT+9,max_row=DT+11)); ch5.legend=None
ch5.y_axis.delete=False; ch5.x_axis.delete=False; ch5.y_axis.majorGridlines=None
ch5.series[0].graphicalProperties.solidFill=AZURE
for i,c in enumerate(['E05252','F59E0B','22A06B']):
    pt=DataPoint(idx=i); pt.graphicalProperties.solidFill=c; ch5.series[0].dPt.append(pt)
ch5.dataLabels=DataLabelList(); ch5.dataLabels.showVal=True
for a in ('showSerName','showCatName','showLegendKey','showPercent'): setattr(ch5.dataLabels,a,False)
D_.add_chart(ch5,'J38')
# top urgent table
UR=56
D_[f'B{UR-1}']='Top 10 critical shipments (RED, follows filters, earliest pickup first) - see Risk reasons for TT / WH-closure conflicts'; D_[f'B{UR-1}'].font=f(11,True,BLUE)
uh=['#','Line ID','Customer','Country','Pallets','DG','Forwarder','Pickup by','Days','Risk reasons']
ucol=[2,3,4,6,7,8,9,10,11,12]
for t,cn in zip(uh,ucol):
    c=D_.cell(row=UR,column=cn,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
for cn in range(2,14): D_.cell(row=UR,column=cn).fill=fill(BLUE)
D_.merge_cells(start_row=UR,start_column=4,end_row=UR,end_column=5)
D_.merge_cells(start_row=UR,start_column=12,end_row=UR,end_column=13)
for i in range(10):
    r=UR+1+i
    D_.merge_cells(start_row=r,start_column=4,end_row=r,end_column=5); D_.merge_cells(start_row=r,start_column=12,end_row=r,end_column=13)
    m=f'MATCH($B{r},{TR("AU")},0)'
    D_[f'B{r}']=i+1
    for cn,k in zip([3,4,6,7,8,9,10,11,12],['A','C','D','G','J','K','S','T','AD']):
        D_.cell(row=r,column=cn,value=f'=IFERROR(INDEX({TR(k)},{m}),"")')
    D_[f'J{r}'].number_format='dd-mmm'; 
    for cn in range(2,14):
        c=D_.cell(row=r,column=cn); c.font=f(9); c.border=Border(bottom=thin)
    D_[f'L{r}'].font=f(8,c=GRAY)
D_[f'B{UR+11}']='="Showing up to 10 of "&COUNTIFS('+INF+',1,'+TR("AE")+',"Red")&" red lines. Full list: Shipment Tracking, filter column AE = Red."'; D_[f'B{UR+11}'].font=f(8,c=GRAY)
D_.sheet_properties.tabColor=BLUE
D_.freeze_panes='A6'
for r in range(DT-2,DT+30): pass

# ---------- Weekly forecast ----------
W=wsW
for i,wd in enumerate([13,13,8]+[11]*20): W.column_dimensions[L(i+1)].width=wd
W['A1']='Weekly truck forecast'; W['A1'].font=f(18,True,BLUE)
W['A2']='Grouped by PICKUP week (= document close date; pickup cannot be brought forward). Trucks = ROUNDUP(pallets / capacity) per country + forwarder + DG + pickup day, then summed. All values are ESTIMATES until entered as confirmed bookings (Settings, section 4).'
W['A2'].font=f(9,c=GRAY)
heads=['Week start (Mon)','Week end (Sun)','ISO wk','Lines','Pallets','DG pallets','Non-DG pallets','Trucks DBS','Trucks MNT','Trucks DG','Trucks Non-DG','TOTAL est. trucks','Confirmed trucks','Still to book','Note']
for i,t in enumerate(heads):
    c=W.cell(row=4,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE); c.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
for k in range(4):
    c=W.cell(row=4,column=16+k,value=f'=IF(INDEX(WhsList,{k+1})="","","Trucks "&INDEX(WhsList,{k+1}))'); c.font=f(9,True,WHITE); c.fill=fill(AZURE); c.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
    for r_ in range(5,17):
        W.cell(row=r_,column=16+k,value=f'=IF({L(16+k)}$4="","",SUMIFS({TR("Z")},{TR("AB")},$A{r_},{TR("BE")},INDEX(WhsList,{k+1})))')
    W.cell(row=17,column=16+k,value=f'=IF({L(16+k)}$4="","",SUMIFS({TR("Z")},{TR("BE")},INDEX(WhsList,{k+1}))-SUM({L(16+k)}5:{L(16+k)}16))')
    W.cell(row=18,column=16+k,value=f'=IF({L(16+k)}$4="","",SUM({L(16+k)}5:{L(16+k)}17))')
    for r_ in range(5,19):
        cc=W.cell(row=r_,column=16+k); cc.font=f(10,b=(r_==18)); cc.border=BOX; cc.number_format='0;-0;"-"'; cc.alignment=Alignment(horizontal='right')
        if r_==18: cc.fill=fill(LIGHT)
W.row_dimensions[4].height=32
for i in range(12):
    r=5+i
    W[f'A{r}']=f'=FcstStart+7*{i}'; W[f'B{r}']=f'=A{r}+6'; W[f'C{r}']=f'=WEEKNUM(A{r},21)'
    W[f'D{r}']=f'=COUNTIFS({TR("AB")},A{r},{TR("AF")},1)'
    W[f'E{r}']=f'=SUMIFS({TR("G")},{TR("AB")},A{r})'
    W[f'F{r}']=f'=SUMIFS({TR("G")},{TR("AB")},A{r},{TR("J")},"DG")'
    W[f'G{r}']=f'=SUMIFS({TR("G")},{TR("AB")},A{r},{TR("J")},"Non-DG")'
    W[f'H{r}']=f'=SUMIFS({TR("Z")},{TR("AB")},A{r},{TR("K")},"DBS")'
    W[f'I{r}']=f'=SUMIFS({TR("Z")},{TR("AB")},A{r},{TR("K")},"MNT")'
    W[f'J{r}']=f'=SUMIFS({TR("Z")},{TR("AB")},A{r},{TR("J")},"DG")'
    W[f'K{r}']=f'=SUMIFS({TR("Z")},{TR("AB")},A{r},{TR("J")},"Non-DG")'
    W[f'L{r}']=f'=H{r}+I{r}'
    W[f'M{r}']=f'=SUMPRODUCT((BkDate<>"")*(BkDate-WEEKDAY(BkDate,2)+1=A{r})*BkTrucks)'
    W[f'N{r}']=f'=MAX(0,L{r}-M{r})'
    W[f'O{r}']=f'=IF(B{r}<ReportDate,"Past week",IF(A{r}<=ReportDate,"Current week",""))'
    W[f'A{r}'].number_format='dd-mmm-yy'; W[f'B{r}'].number_format='dd-mmm-yy'
r=17
W[f'A{r}']='Outside range / undated'
for col,fm in (('D',f'=COUNTIFS({TR("AF")},1)-SUM(D5:D16)'),('E',f'=SUMIFS({TR("G")},{TR("AF")},1)-SUM(E5:E16)'),
               ('F',f'=SUMIFS({TR("G")},{TR("J")},"DG")-SUM(F5:F16)'),('G',f'=SUMIFS({TR("G")},{TR("J")},"Non-DG")-SUM(G5:G16)'),
               ('H',f'=SUMIFS({TR("Z")},{TR("K")},"DBS")-SUM(H5:H16)'),('I',f'=SUMIFS({TR("Z")},{TR("K")},"MNT")-SUM(I5:I16)'),
               ('J',f'=SUMIFS({TR("Z")},{TR("J")},"DG")-SUM(J5:J16)'),('K',f'=SUMIFS({TR("Z")},{TR("J")},"Non-DG")-SUM(K5:K16)'),
               ('L','=H17+I17'),('M','=SUMPRODUCT(--(BkTrucks<>""),BkTrucks)-SUM(M5:M16)'),('N','=MAX(0,L17-M17)')):
    W[f'{col}{r}']=fm
W['O17']='Lines with no pickup week, or a week beyond the 12 shown (sea/air lines have no trucks)'
r=18; W[f'A{r}']='TOTAL'
for col in 'DEFGHIJKLMN': W[f'{col}{r}']=f'=SUM({col}5:{col}17)'
for r in range(5,19):
    for c in range(1,16):
        cell=W.cell(row=r,column=c); cell.font=f(10,b=(r==18 or c==12)); cell.border=BOX
        if 4<=c<=14: cell.number_format='#,##0'; cell.alignment=Alignment(horizontal='right')
    if r in (17,): 
        for c in range(1,16): W.cell(row=r,column=c).font=f(9,c=GRAY)
for c in range(1,16): W.cell(row=18,column=c).fill=fill(LIGHT)
for c in range(1,2): W.cell(row=17,column=c).alignment=Alignment(wrap_text=False)
W['O17'].font=f(8,c=GRAY)
W.conditional_formatting.add('N5:N16',CellIsRule(operator='greaterThan',formula=['0'],fill=ORG_F))
W.conditional_formatting.add('A5:O16',FormulaRule(formula=['$O5="Past week"'],font=Font(name=FONT,color='9A1B1B')))
W.conditional_formatting.add('A5:O16',FormulaRule(formula=['$O5="Current week"'],fill=fill('E4F0FB')))
W['A20']='Estimated trucks by pickup week x destination country (four blocks: forwarder x DG / Non-DG)'; W['A20'].font=f(12,True,BLUE)
blocks=[('DBS','Non-DG'),('DBS','DG'),('MNT','Non-DG'),('MNT','DG')]
b0=22
for bi,(fw,dg) in enumerate(blocks):
    top=b0+bi*17
    W[f'A{top}']=f'{fw} - {dg}  (trucks)'; W[f'A{top}'].font=f(10,True,WHITE)
    for c in range(1,24): W.cell(row=top,column=c).fill=fill(AZURE if fw=='DBS' else BLUE)
    W.cell(row=top+1,column=1,value='Week start'); W.cell(row=top+1,column=2,value='Total')
    for i in range(20): W.cell(row=top+1,column=3+i,value=f'=IF({SN}A{CL0+i}="","",{SN}A{CL0+i})')
    for c in range(1,23):
        cell=W.cell(row=top+1,column=c); cell.font=f(9,True,BLUE); cell.fill=fill(LIGHT); cell.alignment=Alignment(horizontal='center',wrap_text=True)
    W.row_dimensions[top+1].height=28
    for i in range(12):
        r=top+2+i
        W[f'A{r}']=f'=$A${5+i}'; W[f'A{r}'].number_format='dd-mmm-yy'
        W[f'B{r}']=f'=SUMIFS({TR("Z")},{TR("AB")},$A{r},{TR("K")},"{fw}",{TR("J")},"{dg}")'
        for k in range(20):
            cl=L(3+k)
            W[f'{cl}{r}']=f'=IF({cl}${top+1}="","",SUMIFS({TR("Z")},{TR("AB")},$A{r},{TR("D")},{cl}${top+1},{TR("K")},"{fw}",{TR("J")},"{dg}"))'
    r=top+14; W[f'A{r}']='Total'
    for k in range(21):
        cl=L(2+k); W[f'{cl}{r}']=f'=IF({cl}${top+1}="","",SUM({cl}{top+2}:{cl}{top+13}))' if k>0 else f'=SUM(B{top+2}:B{top+13})'
    for rr_ in range(top+2,top+15):
        for c in range(1,23):
            cell=W.cell(row=rr_,column=c); cell.font=f(9,b=(rr_==top+14 or c==2)); cell.border=BOX
            if c>=2: cell.number_format='0;-0;"-"'; cell.alignment=Alignment(horizontal='center')
        if rr_==top+14:
            for c in range(1,23): W.cell(row=rr_,column=c).fill=fill(LIGHT)
W.freeze_panes='B5'; W.sheet_properties.tabColor=AZURE

# ---------- Country analysis ----------
C=wsC
heads=['Country','Shipments (lines)','Pallets','Shipment value (USD)','Avg TT (days)','DG pallets','Non-DG pallets','Est. trucks - total','of which DBS','of which MNT','of which DG','of which Non-DG','Assigned forwarder (rule)','Lines missing pallets']
widths=[18,12,10,16,10,10,12,12,10,10,10,11,22,12]
C['A1']='Country analysis'; C['A1'].font=f(18,True,BLUE)
C['A2']='All shipments in the SE REP master list (not affected by dashboard filters). Country list is maintained on the Assumptions & Settings tab.'; C['A2'].font=f(9,c=GRAY)
for i,(t,wd) in enumerate(zip(heads,widths)):
    c=C.cell(row=4,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE); c.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center'); C.column_dimensions[L(1+i)].width=wd
C.row_dimensions[4].height=34
for i in range(20):
    r=5+i
    C[f'A{r}']=f'=IF({SN}A{CL0+i}="","",{SN}A{CL0+i})'
    def cf(expr): return f'=IF($A{r}="","",{expr})'
    C[f'B{r}']=cf(f'COUNTIFS({TR("D")},$A{r},{TR("AF")},1)')
    C[f'C{r}']=cf(f'SUMIFS({TR("G")},{TR("D")},$A{r})')
    C[f'D{r}']=cf(f'SUMIFS({TR("H")},{TR("D")},$A{r})')
    C[f'E{r}']=cf(f'IFERROR(AVERAGEIFS({TR("N")},{TR("D")},$A{r}),"n/a")')
    C[f'F{r}']=cf(f'SUMIFS({TR("G")},{TR("D")},$A{r},{TR("J")},"DG")')
    C[f'G{r}']=cf(f'SUMIFS({TR("G")},{TR("D")},$A{r},{TR("J")},"Non-DG")')
    C[f'H{r}']=cf(f'SUMIFS({TR("Z")},{TR("D")},$A{r})')
    C[f'I{r}']=cf(f'SUMIFS({TR("Z")},{TR("D")},$A{r},{TR("K")},"DBS")')
    C[f'J{r}']=cf(f'SUMIFS({TR("Z")},{TR("D")},$A{r},{TR("K")},"MNT")')
    C[f'K{r}']=cf(f'SUMIFS({TR("Z")},{TR("D")},$A{r},{TR("J")},"DG")')
    C[f'L{r}']=cf(f'SUMIFS({TR("Z")},{TR("D")},$A{r},{TR("J")},"Non-DG")')
    C[f'M{r}']=cf(f'IF(AND(COUNTIFS({TR("D")},$A{r},{TR("K")},"DBS")>0,COUNTIFS({TR("D")},$A{r},{TR("K")},"MNT")>0),"DBS + MNT (DG battery)",IF(COUNTIFS({TR("D")},$A{r},{TR("K")},"MNT")>0,"MNT",IF(COUNTIFS({TR("D")},$A{r},{TR("K")},"DBS")>0,"DBS","-")))')
    C[f'N{r}']=cf(f'COUNTIFS({TR("D")},$A{r},{TR("AW")},1)')
r=25; C[f'A{r}']='TOTAL'
for col in 'BCDFGHIJKLN': C[f'{col}{r}']=f'=SUM({col}5:{col}24)'
C['E25']=f'=IFERROR(AVERAGE({TR("N")}),"n/a")'
for r in range(5,26):
    for c in range(1,15):
        cell=C.cell(row=r,column=c); cell.font=f(10,b=(r==25)); cell.border=BOX
        if c in (2,3,6,7,8,9,10,11,12,14): cell.number_format='#,##0;-#,##0;"-"'
        if c==4: cell.number_format='$#,##0'
        if c==5: cell.number_format='0.0'
        if r==25: cell.fill=fill(LIGHT)
C.conditional_formatting.add('N5:N24',CellIsRule(operator='greaterThan',formula=['0'],fill=RED_F))
C.conditional_formatting.add('H5:H24',FormulaRule(formula=['AND(ISNUMBER($H5),$H5>=10)'],fill=ORG_F))
ch=BarChart(); ch.type='col'; ch.grouping='stacked'; ch.overlap=100; ch.title='Pallets by country: DG vs Non-DG'; ch.width=24; ch.height=9; ch.roundedCorners=False
ch.add_data(Reference(C,min_col=6,max_col=7,min_row=4,max_row=18),titles_from_data=True); ch.set_categories(Reference(C,min_col=1,min_row=5,max_row=18))
ch.series[0].graphicalProperties.solidFill=BLUE; ch.series[1].graphicalProperties.solidFill=SKY
ch.x_axis.delete=False; ch.y_axis.delete=False; ch.legend.position='b'; ch.legend.overlay=False
C.add_chart(ch,'A28')
C.freeze_panes='B5'; C.sheet_properties.tabColor=AZURE

# ---------- Forwarder allocation ----------
F=wsF
for i,wd in enumerate([26,14,12,16,12,12,12,12,12,40]): F.column_dimensions[L(i+1)].width=wd
F['A1']='Forwarder allocation'; F['A1'].font=f(18,True,BLUE)
F['A2']='Assigned forwarder = rule result (UK -> MNT; DG / CSS Battery to NL, DE, PL, UK -> MNT; all else DBS). "Current" = forwarder already set in SE REP / Backlog (DEFAULT = not assigned yet).'; F['A2'].font=f(9,c=GRAY)
heads=['Forwarder','Shipments (lines)','Pallets','Shipment value (USD)','Est. trucks','Trucks DG','Trucks Non-DG','DG pallets','Non-DG pallets','Share of pallets']
for i,t in enumerate(heads):
    c=F.cell(row=4,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE); c.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
F.row_dimensions[4].height=30
for i,fw in enumerate(['DBS','MNT']):
    r=5+i
    F[f'A{r}']=fw
    F[f'B{r}']=f'=COUNTIFS({TR("K")},"{fw}",{TR("AF")},1)'
    F[f'C{r}']=f'=SUMIFS({TR("G")},{TR("K")},"{fw}")'
    F[f'D{r}']=f'=SUMIFS({TR("H")},{TR("K")},"{fw}")'
    F[f'E{r}']=f'=SUMIFS({TR("Z")},{TR("K")},"{fw}")'
    F[f'F{r}']=f'=SUMIFS({TR("Z")},{TR("K")},"{fw}",{TR("J")},"DG")'
    F[f'G{r}']=f'=SUMIFS({TR("Z")},{TR("K")},"{fw}",{TR("J")},"Non-DG")'
    F[f'H{r}']=f'=SUMIFS({TR("G")},{TR("K")},"{fw}",{TR("J")},"DG")'
    F[f'I{r}']=f'=SUMIFS({TR("G")},{TR("K")},"{fw}",{TR("J")},"Non-DG")'
    F[f'J{r}']=f'=IFERROR(C{r}/C$7,0)'
F['A7']='TOTAL'
for col in 'BCDEFGHI': F[f'{col}7']=f'=SUM({col}5:{col}6)'
F['J7']='=SUM(J5:J6)'
for r in range(5,8):
    for c in range(1,11):
        cell=F.cell(row=r,column=c); cell.font=f(10,b=(r==7)); cell.border=BOX
        if 2<=c<=9: cell.number_format='#,##0'
        if c==4: cell.number_format='$#,##0'
        if c==10: cell.number_format='0%'
        if r==7: cell.fill=fill(LIGHT)
F['A9']='Rule compliance (status of the forwarder already in the system vs the rule)'; F['A9'].font=f(12,True,BLUE)
for i,t in enumerate(['Check result','Lines','Pallets','Meaning']):
    c=F.cell(row=10,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
chk=[('OK','Current forwarder equals the rule'),('WRONG FORWARDER','Current forwarder is DBS/MNT but breaks the rule - re-assign'),('Non-standard forwarder','Forwarder is neither DBS nor MNT (e.g. UPS) - review'),('Not yet assigned','System shows DEFAULT - use the rule result as the recommendation')]
for i,(k,m) in enumerate(chk):
    r=11+i; F[f'A{r}']=k; F[f'B{r}']=f'=COUNTIFS({TR("M")},A{r})'; F[f'C{r}']=f'=SUMIFS({TR("G")},{TR("M")},A{r})'; F[f'D{r}']=m
    for c in range(1,5): F.cell(row=r,column=c).font=f(10); F.cell(row=r,column=c).border=BOX
    F[f'B{r}'].number_format='#,##0'; F[f'C{r}'].number_format='#,##0'
F.conditional_formatting.add('A12:B12',FormulaRule(formula=['$B$12>0'],fill=RED_F))
F.conditional_formatting.add('A13:B13',FormulaRule(formula=['$B$13>0'],fill=ORG_F))
F['A16']='Non-compliant lines (WRONG / non-standard forwarder) - first 20'; F['A16'].font=f(12,True,BLUE)
for i,t in enumerate(['Line ID','Customer','','Country','Family','Pallets','Current','Rule','Check']):
    c=F.cell(row=17,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
for i in range(20):
    r=18+i; F[f'J{r}']=i+1; F[f'J{r}'].font=f(8,c='FFFFFF')
    m=f'MATCH($J{r},{TR("AV")},0)'
    for col,k in zip('ABDEFGHI',['A','C','D','E','G','L','K','M']): F[f'{col}{r}']=f'=IFERROR(INDEX({TR(k)},{m}),"")'
    for c in range(1,10): F.cell(row=r,column=c).font=f(9); F.cell(row=r,column=c).border=Border(bottom=thin)
F['A39']='If this list is empty, no line already assigned in the system conflicts with the rule.'; F['A39'].font=f(8,c=GRAY)
F['A41']='Forwarder x country: estimated trucks'; F['A41'].font=f(12,True,BLUE)
for i,t in enumerate(['Country','DBS trucks','MNT trucks','DBS pallets','MNT pallets']):
    c=F.cell(row=42,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
for i in range(20):
    r=43+i
    F[f'A{r}']=f'=IF({SN}A{CL0+i}="","",{SN}A{CL0+i})'
    for col,(fw,k) in zip('BCDE',[('DBS','Z'),('MNT','Z'),('DBS','G'),('MNT','G')]):
        F[f'{col}{r}']=f'=IF($A{r}="","",SUMIFS({TR(k)},{TR("D")},$A{r},{TR("K")},"{fw}"))'
    for c in range(1,6): F.cell(row=r,column=c).font=f(9); F.cell(row=r,column=c).border=Border(bottom=thin)
    for c in range(2,6): F.cell(row=r,column=c).number_format='0;-0;"-"'
F.sheet_properties.tabColor=AZURE

# ---------- Data checks ----------
K=wsK
K.column_dimensions['A'].width=58; K.column_dimensions['B'].width=14; K.column_dimensions['C'].width=70
K['A1']='Data checks - missing or inconsistent source data'; K['A1'].font=f(18,True,BLUE)
K['A2']='Live checks on the SE REP and Backlog files. Nothing here is estimated or filled in - gaps are listed so they can be completed in the source.'; K['A2'].font=f(9,c=GRAY)
for i,t in enumerate(['Check','Count','What to do']):
    c=K.cell(row=4,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
checks=[
 ('Order lines in SE REP (master)',f'=SUM({TR("AF")})','Master list - one row per order line.'),
 ('Distinct orders',f'=SUMPRODUCT(({TR("AF")}=1)/COUNTIF({TR("B")},{TR("B")}&""))',''),
 ('Lines in Backlog_NL',"=COUNTA(Backlog_NL!$A$2:$A$3000)",'Backlog has more lines than SE REP.'),
 ('Backlog lines with no matching SE REP line',f'=B7-SUMPRODUCT(--({TR("AI")}>0))','Not planned - outside the SE REP master list.'),
 ('SE REP lines with NO pallet data at all (no Backlog match and no units-per-pallet in SE REP)',f'=SUM({TR("AW")})','Add these lines to the Backlog file (col BZ) - listed below. Shown RED, excluded from truck counts.'),
 ('  - pallets matched via fallback (Line "x-1" vs "x")',f'=SUMPRODUCT(--({TR("AH")}=0),--({TR("AI")}>0))','Matched on order + base line number.'),
 ('Lines without shipment value (SE REP col Z blank)',f'=SUMPRODUCT(({TR("AF")}=1)*({TR("H")}=""))','Value missing in the source.'),
 ('Lines without TT',f'=SUMPRODUCT(({TR("AF")}=1)*({TR("N")}=""))','TT missing in Backlog Transit Days (or no dedicated TT for an MNT country).'),
 ('Lines without Family Type',f'=SUMPRODUCT(({TR("AF")}=1)*({TR("E")}=""))','DG status cannot be derived - treated as Non-DG and flagged.'),
 ('Lines for customers with NO closure data',f'=SUMPRODUCT(--({TR("AK")}=0),--({TR("AF")}=1))','Add the customer to the WH Closures tab. Customers listed below.'),
 ('Lines with non-truck mode (SEA / AIR)',f'=SUMPRODUCT(({TR("AF")}=1)*({TR("F")}<>"LAND"))','Listed, but no truck is calculated.'),
 ('Lines with a country missing from the Settings country list',f'=SUMPRODUCT(({TR("AF")}=1)*(COUNTIF(CountryList,{TR("D")})=0))','Add the country on the Settings tab (Country Analysis / charts read that list).'),
 ('Lines with WRONG forwarder vs rule',f'=COUNTIF({TR("M")},"WRONG FORWARDER")','See Forwarder Allocation tab.'),
 ('Lines with non-standard forwarder (e.g. UPS)',f'=COUNTIF({TR("M")},"Non-standard forwarder")',''),
 ('Lines not yet assigned (system = DEFAULT)',f'=COUNTIF({TR("M")},"Not yet assigned")','Informational - the rule result is the recommendation.'),
 ('Check: total value in SE REP minus value in tracking',f'=SUM(SE_REP!$Z$2:$Z$1201)-SUM({TR("H")})','Should be 0.'),
 ('Check: lines beyond tracking capacity (1,200)',"=MAX(0,COUNTA(SE_REP!$E$2:$E$5000)-1200)",'Should be 0.'),
 ('Pallets derived from SE REP (qty / units per pallet) - no Backlog match',f'=COUNTIF({TR("BN")},"SE REP*")','Same logic as Backlog col BZ: ROUNDUP(quantity / units per pallet). Matches Backlog on 495 of 497 lines that exist in both. Please confirm.'),
]
for i,(a,b,c_) in enumerate(checks):
    r=5+i; K[f'A{r}']=a; K[f'B{r}']=b; K[f'C{r}']=c_
    K[f'A{r}'].font=f(10); K[f'B{r}'].font=f(10,True); K[f'C{r}'].font=f(9,c=GRAY); K[f'B{r}'].number_format='#,##0'
    for c in range(1,4): K.cell(row=r,column=c).border=BOX
K.conditional_formatting.add('B9',CellIsRule(operator='greaterThan',formula=['0'],fill=RED_F))
K.conditional_formatting.add('B11:B14',CellIsRule(operator='greaterThan',formula=['0'],fill=ORG_F))
K.conditional_formatting.add('B17',CellIsRule(operator='greaterThan',formula=['0'],fill=RED_F))
K['A24']='Lines with no pallet data'; K['A24'].font=f(12,True,BLUE)
for i,t in enumerate(['Line ID','Country','Customer / family']):
    c=K.cell(row=25,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
for i in range(40):
    r=26+i; K[f'E{r}']=i+1; K[f'E{r}'].font=f(8,c='FFFFFF')
    m=f'MATCH($E{r},{TR("AX")},0)'
    K[f'A{r}']=f'=IFERROR(INDEX({TR("A")},{m}),"")'; K[f'B{r}']=f'=IFERROR(INDEX({TR("D")},{m}),"")'
    K[f'C{r}']=f'=IFERROR(INDEX({TR("C")},{m})&"  |  "&INDEX({TR("E")},{m}),"")'
    for c in range(1,4): K.cell(row=r,column=c).font=f(9); K.cell(row=r,column=c).border=Border(bottom=thin)
K['A67']='Customers with no closure data (first line shown)'; K['A67'].font=f(12,True,BLUE)
for i,t in enumerate(['Customer','Country','']):
    c=K.cell(row=68,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE)
for i in range(15):
    r=69+i; K[f'E{r}']=i+1; K[f'E{r}'].font=f(8,c='FFFFFF')
    m=f'MATCH($E{r},{TR("AZ")},0)'
    K[f'A{r}']=f'=IFERROR(INDEX({TR("C")},{m}),"")'; K[f'B{r}']=f'=IFERROR(INDEX({TR("D")},{m}),"")'
    for c in range(1,3): K.cell(row=r,column=c).font=f(9); K.cell(row=r,column=c).border=Border(bottom=thin)
K.sheet_properties.tabColor=AZURE


# ---------- Weekly Booking Plan ----------
P=wb.create_sheet('Weekly Booking Plan',1); P.sheet_view.showGridLines=False; P.sheet_properties.tabColor=BLUE
for i,wd in enumerate([5,11,11,13,15,32,46,9,9,6,8,6,12,11,11,11,9,8,9,8,9,52]): P.column_dimensions[L(i+1)].width=wd
P.column_dimensions['AA'].hidden=True; P.column_dimensions['Z'].hidden=True
for r_ in (1,2):
    for c_ in range(1,23): P.cell(row=r_,column=c_).fill=fill(BLUE)
P['B1']='Weekly booking plan - trucks to book'; P['B1'].font=Font(name=FONT,size=20,bold=True,color=WHITE)
P['B2']='="One row per delivery stop: pickup date x sending warehouse x destination x customer x address x cargo type x forwarder.  Report date "&TEXT(ReportDate,"dd-mmm-yyyy")&"  |  ESTIMATED trucks, capacity "&CapStd&" pallets (DG "&CapDG&")"'; P['B2'].font=Font(name=FONT,size=10,color=SKY)
P.row_dimensions[1].height=34
for c_ in range(2,3): P.cell(row=3,column=c_).fill=fill('FF3232')
P.row_dimensions[3].height=4
# helper dropdown lists
P['Z4']='All'
for i in range(12): P.cell(row=5+i,column=26,value=f"='Weekly Truck Forecast'!A{5+i}").number_format='dd-mmm-yy'
sel=[('B','Week (pickup week start)',2),('D','Sending warehouse',1),('E','Country',1),('F','Forwarder',1)]
for col,lab,span in sel:
    P[f'{col}4']=lab; P[f'{col}4'].font=f(9,True,GRAY)
P.merge_cells('B5:C5')
for cell in ('B5','D5','E5','F5','C5'):
    c=P[cell]; c.fill=INPUT_FILL; c.font=Font(name=FONT,size=11,bold=True,color='0000FF'); c.border=BOX; c.alignment=Alignment(horizontal='center')
    P[cell].number_format='dd-mmm-yyyy'
for cell in ('B5','D5','E5','F5'): P[cell]='All'
P.row_dimensions[5].height=22
dvw=DataValidation(type='list',formula1="='Weekly Booking Plan'!$Z$4:$Z$16",allow_blank=False); dvh=DataValidation(type='list',formula1=f"={SN}$L$37:$L$45",allow_blank=False)
dvc=DataValidation(type='list',formula1=f"={SN}$I$37:$I$57",allow_blank=False); dvf=DataValidation(type='list',formula1='"All,DBS,MNT"',allow_blank=False)
for d in (dvw,dvh,dvc,dvf): P.add_data_validation(d)
dvw.add('B5'); dvh.add('D5'); dvc.add('E5'); dvf.add('F5')
for n,cell in (('PlanWeek','B'),('PlanWhs','D'),('PlanCountry','E'),('PlanFwd','F')): name(n,f"'Weekly Booking Plan'!${cell}$5")
PF=TR('BJ')
def kp(col,lab,formula,fmt):
    P[f'{col}7']=lab; P[f'{col}7'].font=f(9,True,GRAY)
    c=P[f'{col}8']; c.value=formula; c.font=Font(name=FONT,size=20,bold=True,color=BLUE); c.number_format=fmt; c.alignment=Alignment(horizontal='left')
    P[f'{col}7'].border=Border(top=Side(style='thick',color=AZURE))
kp('B','TRUCKS TO BOOK (EST.)',f'=SUMIFS({TR("Z")},{PF},1)','#,##0')
kp('D','DBS TRUCKS',f'=SUMIFS({TR("Z")},{PF},1,{TR("K")},"DBS")','#,##0')
kp('E','MNT TRUCKS',f'=SUMIFS({TR("Z")},{PF},1,{TR("K")},"MNT")','#,##0')
kp('F','PALLETS  |  VALUE (USD)',f'=TEXT(SUMIFS({TR("G")},{PF},1),"#,##0")&" pallets  |  $"&TEXT(SUMIFS({TR("H")},{PF},1),"#,##0")','@')
kp('G','DELIVERY STOPS  |  ORDER LINES',f'=COUNTIFS({TR("BK")},">0")&" stops  |  "&COUNTIFS({PF},1)&" lines"','@')
P['F8'].font=Font(name=FONT,size=11,bold=True,color=BLUE); P['G8'].font=Font(name=FONT,size=11,bold=True,color=BLUE); P['G8'].alignment=Alignment(horizontal='left',indent=1)
P.row_dimensions[8].height=28
P['B10']='Trucks to book per stop = whole trucks counted once per truck load. DBS loads are pooled per warehouse + country + pickup day + DG, so several stops can share one truck (see "Truck share" and "Trucks in shared load"). MNT trucks are direct: one per customer + address + day. A red row = review before booking (reasons on the right).'
P['B10'].font=f(8,c=GRAY)
PH=12; PR0=13; PCAP=600
heads=['#','Week start','Pickup date (doc close)','Sending warehouse','Destination country','Customer','Delivery address','Cargo type','Forwarder','Lines','Pallets','TT (days)','Shipment value (USD)','Customer WH closes from','WH reopens after','Arrival if shipped at pickup date','TRUCKS TO BOOK','Truck share','Trucks in shared load','Truck capacity','Risk','Risk reasons (worst line of the stop)']
for i,t in enumerate(heads):
    c=P.cell(row=PH,column=1+i,value=t); c.font=f(9,True,WHITE); c.fill=fill(BLUE if t!='TRUCKS TO BOOK' else 'FF3232'); c.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center'); c.border=BOX
P.row_dimensions[PH].height=40
for i in range(PCAP):
    r_=PR0+i
    P[f'A{r_}']=i+1
    P[f'AA{r_}']=f'=IFERROR(MATCH($A{r_},{TR("BK")},0),"")'
    ix=lambda k: f'=IF($AA{r_}="","",INDEX({TR(k)},$AA{r_}))'
    key=f'INDEX({TR("BG")},$AA{r_})'
    P[f'B{r_}']=ix('AB'); P[f'C{r_}']=ix('S'); P[f'D{r_}']=ix('BE'); P[f'E{r_}']=ix('D'); P[f'F{r_}']=ix('C'); P[f'G{r_}']=ix('BF'); P[f'H{r_}']=ix('J'); P[f'I{r_}']=ix('K')
    P[f'J{r_}']=f'=IF($AA{r_}="","",COUNTIFS({TR("BG")},{key}))'
    P[f'K{r_}']=f'=IF($AA{r_}="","",SUMIFS({TR("G")},{TR("BG")},{key}))'
    P[f'L{r_}']=ix('N')
    P[f'M{r_}']=f'=IF($AA{r_}="","",SUMIFS({TR("H")},{TR("BG")},{key}))'
    P[f'N{r_}']=ix('Q'); P[f'O{r_}']=ix('BL'); P[f'P{r_}']=ix('R')
    P[f'Q{r_}']=f'=IF($AA{r_}="","",SUMIFS({TR("Z")},{TR("BG")},{key}))'
    P[f'R{r_}']=f'=IF($AA{r_}="","",SUMIFS({TR("AA")},{TR("BG")},{key}))'
    P[f'S{r_}']=ix('Y'); P[f'T{r_}']=ix('V')
    worst=f'SUMPRODUCT(MAX(({TR("BG")}={key})*{TR("BI")}))'
    P[f'U{r_}']=f'=IF($AA{r_}="","",IF({worst}=3,"Red",IF({worst}=2,"Orange","Green")))'
    P[f'V{r_}']=f'=IF($AA{r_}="","",IFERROR(INDEX({TR("AD")},MATCH({key}&"|"&{worst},{TR("BM")},0)),""))'
    for c_ in range(1,23):
        cell=P.cell(row=r_,column=c_); cell.font=f(9); cell.border=Border(bottom=thin)
    for col in 'BCNOP': P[f'{col}{r_}'].number_format='dd-mmm-yy'
    P[f'M{r_}'].number_format='#,##0'; P[f'R{r_}'].number_format='0.00'; P[f'Q{r_}'].font=f(10,True,BLUE)
    for col in 'ABCDEHIJKLQRSTU': P[f'{col}{r_}'].alignment=Alignment(horizontal='center')
    P[f'A{r_}'].font=f(8,c=GRAY)
rg=f'U{PR0}:U{PR0+PCAP-1}'
P.conditional_formatting.add(rg,CellIsRule(operator='equal',formula=['"Red"'],fill=RED_F,font=Font(name=FONT,bold=True,color='7A0000')))
P.conditional_formatting.add(rg,CellIsRule(operator='equal',formula=['"Orange"'],fill=ORG_F,font=Font(name=FONT,bold=True,color='7A4A00')))
P.conditional_formatting.add(rg,CellIsRule(operator='equal',formula=['"Green"'],fill=GRN_F,font=Font(name=FONT,bold=True,color='1D5A33')))
P.conditional_formatting.add(f'I{PR0}:I{PR0+PCAP-1}',CellIsRule(operator='equal',formula=['"MNT"'],font=Font(name=FONT,bold=True,color=AZURE)))
P.conditional_formatting.add(f'A{PR0+1}:V{PR0+PCAP-1}',FormulaRule(formula=[f'AND($B{PR0+1}<>"",$B{PR0+1}<>$B{PR0})'],border=Border(top=Side(style='medium',color=BLUE))))
P.conditional_formatting.add(f'N{PR0}:N{PR0+PCAP-1}',FormulaRule(formula=[f'AND(ISNUMBER($N{PR0}),ISNUMBER($P{PR0}),$P{PR0}>=$N{PR0})'],fill=RED_F))
P.freeze_panes=f'H{PR0}'
P.auto_filter.ref=f'A{PH}:V{PR0+PCAP-1}'
P.page_setup.orientation='landscape'; P.page_setup.fitToWidth=1; P.page_setup.fitToHeight=0
P.sheet_properties.pageSetUpPr=openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)

# page setup
for w_ in (wsD,wsW,wsC,wsF,wsT,wsS,wsK):
    w_.page_setup.orientation='landscape'; w_.page_setup.fitToWidth=1; w_.page_setup.fitToHeight=0
    w_.sheet_properties.pageSetUpPr=openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
wb.calculation.fullCalcOnLoad=True
wb.save(OUT); print('saved',OUT)
