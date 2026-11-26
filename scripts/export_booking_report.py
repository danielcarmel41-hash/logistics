"""Static (values-only) Excel report generated from the recalculated dashboard workbook."""
import openpyxl, pandas as pd, datetime as dt, warnings
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
warnings.filterwarnings('ignore')
SRC='output/Quarter_End_Truck_Planning_Dashboard.xlsx'; OUT='output/Quarter_End_Truck_Booking_Report.xlsx'
wb=openpyxl.load_workbook(SRC,data_only=True)
def frame(ws,hr,r1,ncol,key=0):
    rows=[[c.value for c in r] for r in ws.iter_rows(min_row=hr,max_row=r1,max_col=ncol)]
    df=pd.DataFrame(rows[1:],columns=rows[0]); return df[df.iloc[:,key].notna()&(df.iloc[:,key]!='')]
plan=frame(wb['Weekly Booking Plan'],12,611,22,1).drop(columns=['#'])
t=wb['Shipment Tracking']; hdr=[c.value for c in t[5]]
trk=pd.DataFrame([[c.value for c in r] for r in t.iter_rows(min_row=6,max_row=1205,max_col=len(hdr))],columns=hdr)
trk=trk[trk['Line ID'].notna()&(trk['Line ID']!='')]
lines=trk[['Line ID','Order #','Customer','Delivery address','Country','Sending warehouse','Family type','DG / Non-DG','Mode','Pallets','Pallet source','TT (days)','TT basis','Shipment value (USD)','Doc close date','Pickup date (= doc close)','Binding WH closure start','Binding WH closure end','Arrival if shipped at doc close','Pickup week (Mon)','Assigned forwarder (rule)','Current forwarder (in system)','Forwarder check','Group trucks (est.)','Truck allocation (share)','Planning status','Risk indicator','Risk reasons']].copy()
lines=lines.rename(columns={'Binding WH closure end':'WH reopens after','Binding WH closure start':'Customer WH closes from'})
ex=lines[(lines['Risk reasons'].astype(str).str.contains('NOT SYNCED|No pallet|No value|No WH closure|Wrong forwarder|Pickup date passed|No TT|No family'))].copy()
def reason(s):
    for k,v in [('No pallet','Missing pallets'),('NOT SYNCED','TT not synced with WH closure'),('Wrong forwarder','Wrong forwarder'),('Pickup date passed','Pickup date passed'),('No TT','Missing TT'),('No family','Missing family type'),('No value','Missing value'),('No WH closure','No WH closure data')]:
        if k in s: return v
    return ''
ex.insert(0,'Main issue',ex['Risk reasons'].map(lambda s: '; '.join(v for k,v in [('No pallet','Missing pallets'),('NOT SYNCED','TT not synced with WH closure'),('Wrong forwarder','Wrong forwarder'),('Pickup date passed','Pickup date passed'),('No TT','Missing TT'),('No family','Missing family type'),('No value','Missing value'),('No WH closure','No WH closure data')] if k in str(s))))
# weekly summary
w=wb['Weekly Truck Forecast']; wk=frame(w,4,16,19,0)
wk=wk.loc[:, [c for c in wk.columns if c and c!='Note']]
BLUE='00082E'; AZ='2E8FD6'
hf=Font(name='Segoe UI',bold=True,color='FFFFFF',size=10); hfill=PatternFill('solid',start_color=BLUE,end_color=BLUE)
thin=Side(style='thin',color='D5DCE8'); box=Border(left=thin,right=thin,top=thin,bottom=thin)
fills={'Red':'F7B6B6','Orange':'FFD59A','Green':'BFE3C8'}
out=openpyxl.Workbook(); S=out.active; S.title='Summary'
def put(ws,df,r0,widths=None,datecols=(),risk=None):
    for j,c in enumerate(df.columns):
        x=ws.cell(row=r0,column=j+1,value=c); x.font=hf; x.fill=hfill; x.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center'); x.border=box
    ws.row_dimensions[r0].height=34
    for i,row in enumerate(df.itertuples(index=False)):
        for j,v in enumerate(row):
            if isinstance(v,pd.Timestamp): v=v.to_pydatetime() if not pd.isna(v) else None
            if v is not None and isinstance(v,float) and pd.isna(v): v=None
            c=ws.cell(row=r0+1+i,column=j+1,value=v); c.font=Font(name='Segoe UI',size=9); c.border=box
            if isinstance(v,dt.datetime): c.number_format='dd-mmm-yy'
            elif isinstance(v,float): c.number_format='#,##0.##'
    if risk is not None:
        j=list(df.columns).index(risk)+1
        for i in range(len(df)):
            c=ws.cell(row=r0+1+i,column=j)
            if c.value in fills: c.fill=PatternFill('solid',start_color=fills[c.value],end_color=fills[c.value]); c.font=Font(name='Segoe UI',size=9,bold=True)
    if widths:
        for j,wd in enumerate(widths): ws.column_dimensions[L(j+1)].width=wd
S['A1']='Quarter-end truck booking report'; S['A1'].font=Font(name='Segoe UI',size=18,bold=True,color=BLUE)
S['A2']=f'Values-only snapshot generated {dt.date.today():%d-%b-%Y} from Quarter_End_Truck_Planning_Dashboard.xlsx. Truck numbers are ESTIMATES (no confirmed bookings provided). Truck capacity 33 pallets, DG / CSS Battery 9.'; S['A2'].font=Font(name='Segoe UI',size=9,color='696969')
tot=lambda k: lines[k].apply(pd.to_numeric,errors='coerce').sum()
kp=[('Order lines',len(lines)),('Pallets',tot('Pallets')),('Shipment value (USD)',tot('Shipment value (USD)')),('Estimated trucks',plan['TRUCKS TO BOOK'].sum()),
    ('Trucks DBS',plan.loc[plan['Forwarder']=='DBS','TRUCKS TO BOOK'].sum()),('Trucks MNT',plan.loc[plan['Forwarder']=='MNT','TRUCKS TO BOOK'].sum()),
    ('Lines Red',(lines['Risk indicator']=='Red').sum()),('Lines Orange',(lines['Risk indicator']=='Orange').sum()),('Lines Green',(lines['Risk indicator']=='Green').sum()),
    ('Delivery stops',len(plan))]
for i,(k,v) in enumerate(kp):
    S.cell(row=4,column=1+i,value=k).font=Font(name='Segoe UI',size=9,bold=True,color='696969')
    c=S.cell(row=5,column=1+i,value=float(v)); c.font=Font(name='Segoe UI',size=14,bold=True,color=BLUE); c.number_format='#,##0'
S['A7']='Trucks to book per week'; S['A7'].font=Font(name='Segoe UI',size=12,bold=True,color=BLUE)
put(S,wk,8,[16,16,8,9,10,11,13,11,11,11,13,14,14,12,14,14,14,14])
r=9+len(wk)+2
S.cell(row=r,column=1,value='Trucks by destination country').font=Font(name='Segoe UI',size=12,bold=True,color=BLUE)
pc=plan.groupby('Destination country').agg(Stops=('Lines','size'),Pallets=('Pallets','sum'),Value=('Shipment value (USD)','sum'),Trucks=('TRUCKS TO BOOK','sum')).reset_index()
dg=plan.pivot_table(index='Destination country',columns='Forwarder',values='TRUCKS TO BOOK',aggfunc='sum',fill_value=0).reset_index()
pc=pc.merge(dg,on='Destination country',how='left').rename(columns={'DBS':'Trucks DBS','MNT':'Trucks MNT','Value':'Value (USD)'})
put(S,pc,r+1)
S.freeze_panes='A4'
B=out.create_sheet('Booking Plan'); put(B,plan,1,[11,12,13,14,32,46,9,9,6,8,6,13,13,13,13,9,8,9,8,9,55],risk='Risk')
B.freeze_panes='G2'; B.auto_filter.ref=f'A1:{L(plan.shape[1])}{len(plan)+1}'
Ln=out.create_sheet('Shipment Lines'); put(Ln,lines,1,[16,10,30,44,14,13,17,9,7,8,22,7,16,13,12,12,13,13,13,12,10,12,17,9,9,26,9,60],risk='Risk indicator')
Ln.freeze_panes='D2'; Ln.auto_filter.ref=f'A1:{L(lines.shape[1])}{len(lines)+1}'
E=out.create_sheet('Exceptions'); put(E,ex,1,[34,16,10,30,44,14,13,17,9,7,8,22,7,16,13,12,12,13,13,13,12,10,12,17,9,9,26,9,60],risk='Risk indicator')
E.freeze_panes='E2'; E.auto_filter.ref=f'A1:{L(ex.shape[1])}{len(ex)+1}'
for ws in out.worksheets: ws.sheet_view.showGridLines=False if ws.title=='Summary' else True; ws.sheet_properties.tabColor=BLUE if ws.title=='Summary' else AZ
out.save(OUT); print(OUT,len(plan),len(lines),len(ex))
