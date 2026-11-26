"""Builds a self-contained, filterable HTML page from the recalculated dashboard workbook."""
import openpyxl, json, datetime as dt, warnings
warnings.filterwarnings('ignore')
SRC='output/Quarter_End_Truck_Planning_Dashboard.xlsx'; OUT='output/Quarter_End_Truck_Planning.html'
wb=openpyxl.load_workbook(SRC,data_only=True)
def iso(v): return v.strftime('%Y-%m-%d') if isinstance(v,(dt.datetime,dt.date)) else None
def num(v):
    return float(v) if isinstance(v,(int,float)) else None
P=wb['Weekly Booking Plan']; stops=[]
for r in P.iter_rows(min_row=13,max_row=412,max_col=22,values_only=True):
    if r[1] in (None,''): continue
    stops.append(dict(week=iso(r[1]),pickup=iso(r[2]),whs=r[3],country=r[4],customer=r[5],address=r[6],cargo=r[7],fwd=r[8],lines=num(r[9]),pallets=num(r[10]),tt=num(r[11]),value=num(r[12]),closeFrom=iso(r[13]),reopen=iso(r[14]),arrival=iso(r[15]),trucks=num(r[16]),share=num(r[17]),shared=num(r[18]),cap=num(r[19]),risk=r[20],reasons=r[21] or ''))
T=wb['Shipment Tracking']; hdr=[c.value for c in T[5]]; ix={h:i for i,h in enumerate(hdr)}
ISS=[('No pallet','Missing pallets'),('NOT SYNCED','TT not synced with WH closure'),('Wrong forwarder','Wrong forwarder'),('Pickup date passed','Pickup date passed'),('No TT','Missing TT'),('No family','Missing family type'),('No value','Missing value'),('No WH closure','No WH closure data')]
lines=[]
for r in T.iter_rows(min_row=6,max_row=805,max_col=len(hdr),values_only=True):
    if r[ix['Line ID']] in (None,''): continue
    g=lambda h: r[ix[h]]
    rs=g('Risk reasons') or ''
    lines.append(dict(id=g('Line ID'),order=g('Order #'),customer=g('Customer'),address=g('Delivery address'),country=g('Country'),whs=g('Sending warehouse'),family=g('Family type'),cargo=g('DG / Non-DG'),mode=g('Mode'),pallets=num(g('Pallets')),palSrc=g('Pallet source'),tt=num(g('TT (days)')),value=num(g('Shipment value (USD)')),docClose=iso(g('Doc close date')),closeFrom=iso(g('Binding WH closure start')),reopen=iso(g('Binding WH closure end')),arrival=iso(g('Arrival if shipped at doc close')),week=iso(g('Pickup week (Mon)')),fwd=g('Assigned forwarder (rule)'),curFwd=g('Current forwarder (in system)'),fwdCheck=g('Forwarder check'),status=g('Planning status'),risk=g('Risk indicator'),reasons=rs,issues=[v for k,v in ISS if k in rs]))
meta=dict(generated=dt.date.today().isoformat(),reportDate=iso(wb['Assumptions & Settings']['B5'].value),capStd=wb['Assumptions & Settings']['B7'].value,capDG=wb['Assumptions & Settings']['B8'].value)
data=json.dumps(dict(meta=meta,stops=stops,lines=lines),ensure_ascii=False,separators=(',',':'))
html=open('scripts/html_template.html',encoding='utf8').read().replace('/*DATA*/null',data)
open(OUT,'w',encoding='utf8').write(html); print(OUT,len(stops),len(lines),len(html)//1024,'KB')
