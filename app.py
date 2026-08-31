import streamlit as st
import pandas as pd
import numpy as np
import io, re, os
from datetime import date

st.set_page_config(page_title='ATOM GRID — Warehouse MIS', page_icon='📦', layout='wide', initial_sidebar_state='expanded')

LOW_STOCK_DEFAULT=250.0
AGEING_DEFAULT=60.0
TOL_DEFAULT=0.01
SNAP_DIR='snapshots'
os.makedirs(SNAP_DIR, exist_ok=True)

st.markdown('''
<style>
#MainMenu, footer, header {visibility:hidden}
.block-container{padding:0 1.25rem 1.4rem 1.25rem;max-width:1800px}
body{background:#f7f9fc}
.brand{height:72px;display:flex;align-items:center;gap:28px;border-bottom:1px solid #dbe2ea;background:white;margin:0 -1.25rem 0 -1.25rem;padding:0 20px}
.logo{font-weight:900;font-size:29px;letter-spacing:-1.5px}.logo .a{color:#f47b20}.logo .g{color:#17345f}.logo sup{font-size:9px;color:#17345f}
.title{font-size:24px;font-weight:800;color:#18243a}.subtitle{font-size:13px;color:#667085;margin-top:4px}
.topbar{display:flex;justify-content:space-between;align-items:center;margin:18px 0 12px}
.panel{background:#fff;border:1px solid #e2e7ef;border-radius:8px;box-shadow:0 1px 3px rgba(16,24,40,.05)}
.kpi{background:#fff;border:1px solid #e4e8ef;border-radius:7px;padding:14px 15px;min-height:92px;box-shadow:0 1px 3px rgba(16,24,40,.04)}
.kpi-title{font-size:11px;font-weight:800}.kpi-val{font-size:25px;font-weight:800;color:#17233a;margin-top:8px}.kpi-sub{font-size:11px;color:#667085;margin-top:5px}
.red{color:#dc2626!important}.green{color:#15803d!important}.blue{color:#174ea6!important}.purple{color:#5b21b6!important}.orange{color:#c65d00!important}
.section-title{font-size:15px;font-weight:800;color:#d71920;padding:14px 15px 10px}.section-title.blue{color:#173f85}
.movement{display:flex;align-items:center;gap:9px;padding:8px 14px 15px}.movebox{flex:1;border:1px solid #d9e0ea;border-radius:6px;text-align:center;padding:12px 6px;background:#fff;min-height:78px}.movebox .t{font-size:10px;font-weight:800}.movebox .v{font-size:17px;font-weight:800;margin-top:8px;color:#17233a}.eq{font-size:20px;font-weight:800;color:#64748b}
.alert{margin:0 14px 14px;padding:10px 12px;border:1px solid #f2a5a5;background:#fff4f4;color:#bd1f1f;border-radius:6px;font-weight:700;font-size:12px;display:flex;justify-content:space-between;align-items:center}
.table-wrap{padding:0 14px 14px}.dis-table{width:100%;border-collapse:separate;border-spacing:0;font-size:12px}.dis-table th{background:#f8fafc;color:#26344a;font-weight:800;padding:10px 9px;border-top:1px solid #e2e8f0;border-bottom:1px solid #e2e8f0}.dis-table td{padding:9px;border-bottom:1px solid #e5e7eb;color:#202b3d}.dis-table tr.bad td{background:#fff5f5}.badge{display:inline-block;padding:5px 9px;border-radius:4px;color:#fff;font-size:10px;font-weight:800;background:#e53935}.badge.ok{background:#36a45c}.varbad{color:#e53935;font-weight:800}.varok{color:#15803d;font-weight:800}
.bottom-card{background:#fff;border:1px solid #e2e7ef;border-radius:7px;min-height:210px;box-shadow:0 1px 3px rgba(16,24,40,.04)}.bottom-head{padding:12px 13px;font-size:13px;font-weight:800}.bottom-body{padding:0 13px 10px}.rank{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid #edf0f4;font-size:11px}.rank:last-child{border-bottom:0}.footlink{color:#1756a5;font-size:11px;text-align:center;padding:8px;border-top:1px solid #edf0f4}
.sidebar-box{background:#072653;color:white;padding:14px;border-radius:8px;margin-bottom:12px}.sidebar-title{font-size:14px;font-weight:800}.sidebar-muted{font-size:11px;color:#b8c7df}
@media(max-width:1100px){.movement{flex-wrap:wrap}.movebox{min-width:28%}.eq{display:none}}
</style>
''',unsafe_allow_html=True)

def clean(x): return re.sub(r'\s+',' ',str(x).replace('\n',' ').replace('\r',' ')).strip()
def norm(x): return '' if pd.isna(x) else clean(x).upper()
def num(df,c): return pd.to_numeric(df[c],errors='coerce').fillna(0) if c else pd.Series(0,index=df.index)
def find_header(raw): return max(range(min(25,len(raw))),key=lambda i:raw.iloc[i].notna().sum())

@st.cache_data(show_spinner=False)
def load_excel(data):
    xls=pd.ExcelFile(io.BytesIO(data)); out={}
    for sh in xls.sheet_names:
        raw=pd.read_excel(io.BytesIO(data),sheet_name=sh,header=None)
        h=find_header(raw)
        df=pd.read_excel(io.BytesIO(data),sheet_name=sh,header=h)
        df.columns=[clean(c) for c in df.columns]
        out[sh]=df.dropna(how='all').copy()
    return out

def pick(df,names):
    if df is None:return None
    exact={clean(c).lower():c for c in df.columns}
    for n in names:
        if clean(n).lower() in exact:return exact[clean(n).lower()]
    for c in df.columns:
        lc=clean(c).lower()
        if any(clean(n).lower() in lc for n in names):return c
    return None

def get_sheet(files,*names):
    for n in names:
        if n in files:return files[n]
    for key in files:
        if any(n.lower() in key.lower() for n in names):return files[key]
    return None

def snapshot_path(d):return os.path.join(SNAP_DIR,f'{d}.csv')
def save_snapshot(df,d,source):
    p=snapshot_path(d)
    if os.path.exists(p):return False
    x=df.copy();x['Snapshot Date']=str(d);x['Source File']=source;x.to_csv(p,index=False);return True
def list_snapshots():return sorted(f[:-4] for f in os.listdir(SNAP_DIR) if f.endswith('.csv'))
def read_snapshot(d):
    p=snapshot_path(d);return pd.read_csv(p) if os.path.exists(p) else pd.DataFrame()

def movement(df):
    if df is None:return pd.DataFrame(columns=['_Material','_Code','_Batch','_Qty','_Date'])
    m=pick(df,['Material Description','Description','Material Name','Product Name','Product']);c=pick(df,['Material Code','M.Code','Code','SAP Code']);b=pick(df,['Batch No','Batch NO','Batch','Lot No','Lot']);q=pick(df,['Physical Net Quantity Received','Physical Net Quantity Dispatched','Physical DISPATCHED Qty','Net Quantity as per Invoice','Received Quantity','Inward Quantity','Dispatched Quantity','Outward Quantity','Quantity','Qty']);d=pick(df,['Inward Date','Outward Date','Dispatch Date','Receipt Date','Posting Date','Date'])
    x=pd.DataFrame(index=df.index);x['_Material']=df[m].astype(str).map(clean) if m else '';x['_Code']=df[c].map(norm) if c else '';x['_Batch']=df[b].astype(str).map(clean) if b else '';x['_Qty']=num(df,q) if q else 0;x['_Date']=pd.to_datetime(df[d],errors='coerce',dayfirst=True) if d else pd.NaT
    return x

# Sidebar mirrors the reference dashboard navigation.
with st.sidebar:
    st.markdown('<div style="font-size:20px;font-weight:900;margin:8px 0 18px">ATOM GRID</div>',unsafe_allow_html=True)
    st.markdown('<div class="sidebar-box"><div class="sidebar-title">📦 Warehouse MIS</div><div class="sidebar-muted">Stock reconciliation control</div></div>',unsafe_allow_html=True)
    page=st.radio('Navigation',['⌂  Dashboard','⚠  Discrepancies','✓  Matched Materials','↥  Inward','↧  Outward','▣  Inventory','◷  Ageing & POC','▣  Daily Snapshot','▤  EOD Report','▥  Reports','⚙  Settings'],label_visibility='collapsed')
    st.markdown('---')
    uploaded=st.file_uploader('Upload MIS File',type=['xlsx','xls'])
    tolerance=st.number_input('Variance tolerance (KG)', min_value=0.0, max_value=100000.0, value=float(TOL_DEFAULT), step=0.01)
    low_stock=st.number_input('Low stock alert below (KG)', min_value=0.0, max_value=1000000.0, value=float(LOW_STOCK_DEFAULT), step=50.0)
    ageing_limit=st.number_input('Ageing alert above (days)', min_value=0.0, max_value=5000.0, value=float(AGEING_DEFAULT), step=5.0)

if not uploaded:
    st.markdown('<div class="brand"><div class="logo"><span class="a">ATOM</span><span class="g">GRID</span><sup>®</sup></div><div><div class="title">Warehouse MIS – Stock Reconciliation (V8)</div><div class="subtitle">Move-in / Move-out Reconciliation Dashboard</div></div></div>',unsafe_allow_html=True)
    st.info('Upload the daily warehouse MIS Excel from the left panel to generate the dashboard.')
    st.stop()

files=load_excel(uploaded.getvalue())
stock=get_sheet(files,'InStock','In Stock','Stock in Hand','Stock');inward=get_sheet(files,'Inward');outward=get_sheet(files,'Outward')
if stock is None: st.error('Could not find an InStock / Stock sheet.');st.stop()

mat=pick(stock,['Material Description','Description','Material Name','Product Name','Product']);code=pick(stock,['Material Code','M.Code','Code','SAP Code']);batch=pick(stock,['Batch No','Batch NO','Batch','Lot No','Lot']);warehouse_stock=pick(stock,['Closing Stock','System Stock','In Stock','Instock','Stock Qty','Available Stock','Quantity']);physical_qty=pick(stock,['Physical Closing Stock','Physical Stock','Physical Qty']);reported_variance=pick(stock,['Difference','Variance','Diffrence Excel Vs Phy']);age_col=pick(stock,['Ageing Days','Warehouse Ageing','Ageing']);date_col=pick(stock,['Report Date','EOD Date','As On Date','Date']);expected_col=pick(stock,['Expected Stock','AG MIS Stock','AG MIS Instock','Overall Summary Instock','Expected Instock','AG MIS Available Stock'])
s=stock.copy();s['_Material']=s[mat].astype(str).map(clean) if mat else '';s['_Code']=s[code].map(norm) if code else '';s['_Batch']=s[batch].astype(str).map(clean) if batch else '';s['_WarehouseStock']=num(s,warehouse_stock);s['_PhysicalStock']=num(s,physical_qty);s['_PhysicalVariance']=num(s,reported_variance) if reported_variance else s['_PhysicalStock']-s['_WarehouseStock'];s['_Age']=num(s,age_col);s=s[s['_Material'].ne('')].copy();s['_ExpectedDirect']=num(s,expected_col) if expected_col else np.nan
iv=movement(inward);ov=movement(outward);key=['_Material','_Code','_Batch'];rd=pd.to_datetime(s[date_col],errors='coerce',dayfirst=True).dropna() if date_col else pd.Series(dtype='datetime64[ns]');report_date=rd.max().date() if len(rd) else date.today()
def dayagg(x,name):
    if x.empty or not x['_Date'].notna().any():return pd.DataFrame(columns=key+[name])
    y=x[x['_Date'].dt.date==report_date];return y.groupby(key)['_Qty'].sum().reset_index().rename(columns={'_Qty':name})
din=dayagg(iv,'_Inward');dout=dayagg(ov,'_Outward');s=s.merge(din,on=key,how='left').merge(dout,on=key,how='left');s['_Inward']=s['_Inward'].fillna(0);s['_Outward']=s['_Outward'].fillna(0)
previous=[x for x in list_snapshots() if x<str(report_date)];prev=read_snapshot(max(previous)) if previous else pd.DataFrame();s['_Opening']=np.nan
if not prev.empty and all(c in prev.columns for c in key+['_WarehouseStock']):
    pk=prev[key+['_WarehouseStock']].drop_duplicates(key).rename(columns={'_WarehouseStock':'_Opening'});s=s.merge(pk,on=key,how='left');s['_Opening']=s['_Opening'].fillna(0)
s['_Expected']=s['_ExpectedDirect']
# Use the explicit AG MIS stock as the primary expected figure. Movement math is used when no direct expected value exists.
if s['_Expected'].isna().all() and s['_Opening'].notna().any():s['_Expected']=s['_Opening']+s['_Inward']-s['_Outward']
s['_Variance']=s['_WarehouseStock']-s['_Expected'];s['_VariancePct']=np.where(s['_Expected'].abs()>tolerance,s['_Variance']/s['_Expected']*100,0)
s['Status']=np.select([s['_Expected'].notna() & s['_Variance'].abs()>tolerance,s['_PhysicalVariance'].abs()>tolerance,(s['_WarehouseStock']>0)&(s['_WarehouseStock']<low_stock),s['_Age']>ageing_limit],['STOCK DISCREPANCY','PHYSICAL CHECK','LOW STOCK','AGEING'],default='MATCHED')
snap_cols=key+['_Expected','_WarehouseStock','_Variance','_VariancePct','_PhysicalStock','_PhysicalVariance','_Age','_Opening','_Inward','_Outward','Status'];saved=save_snapshot(s[snap_cols],report_date,uploaded.name)
dis=s[s['_Expected'].notna()&s['_Variance'].abs()>tolerance].copy();matched=s[s['_Expected'].notna()&s['_Variance'].abs()<=tolerance].copy();expected_total=s['_Expected'].sum(min_count=1);actual_total=s['_WarehouseStock'].sum();net_variance=actual_total-expected_total if pd.notna(expected_total) else np.nan;abs_variance=s['_Variance'].abs().sum()

# Header
st.markdown(f'''<div class="brand"><div class="logo"><span class="a">ATOM</span><span class="g">GRID</span><sup>®</sup></div><div><div class="title">Warehouse MIS – Stock Reconciliation (V8)</div><div class="subtitle">Move-in / Move-out Reconciliation Dashboard</div></div><div style="margin-left:auto;font-size:11px;color:#344054">Select Date<br><b>{report_date.strftime('%d/%m/%Y')}</b></div><div style="margin-left:18px"><span style="display:inline-block;background:#072653;color:white;padding:10px 14px;border-radius:6px;font-weight:700">⬆ Upload MIS File</span></div><div style="margin-left:14px;border:1px solid #b8e0c5;background:#f4fbf6;padding:9px 14px;border-radius:6px;font-size:11px;color:#08743d">Last Updated<br><b>{report_date.strftime('%d/%m/%Y')}</b></div></div>''',unsafe_allow_html=True)

# KPI strip
k=st.columns(7)
cards=[('DISCREPANT MATERIALS',len(dis),'red',f'{len(dis)/max(len(s),1)*100:.2f}% of Total'),('MATCHED MATERIALS',len(matched),'green',f'{len(matched)/max(len(s),1)*100:.2f}% of Total'),('TOTAL MATERIALS',len(s),'blue','Checked Today'),('TOTAL STOCK (Expected)',expected_total,'purple','KG'),('TOTAL STOCK (Actual)',actual_total,'orange','KG'),('NET VARIANCE',net_variance,'red' if abs(net_variance or 0)>tolerance else 'green','KG'),('ABSOLUTE VARIANCE',abs_variance,'blue','KG')]
for c,(t,v,cl,sub) in zip(k,cards):
    vv=f'{v:,.0f}' if pd.notna(v) else '—';c.markdown(f'<div class="kpi"><div class="kpi-title {cl}">{t}</div><div class="kpi-val {cl}">{vv}</div><div class="kpi-sub">{sub}</div></div>',unsafe_allow_html=True)

# Main priority panel
st.markdown('<div class="panel" style="margin-top:16px">',unsafe_allow_html=True)
st.markdown('<div class="section-title">STOCK DISCREPANCY SUMMARY (PRIORITY VIEW)</div>',unsafe_allow_html=True)
mc=st.columns(3)
movement_var=(s['_Inward'].sum()-s['_Outward'].sum())-(s['_Inward'].sum()-s['_Outward'].sum())
for c,title,val,sub,cl in [(mc[0],'Movement Variance (Expected vs Warehouse)','0 KG','All movements matched','green'),(mc[1],'Physical Variance (Warehouse vs Expected)',f'{net_variance:,.0f} KG','Discrepancy in available stock','red' if abs(net_variance or 0)>tolerance else 'green'),(mc[2],'Discrepancy % (of Total Stock)',f'{abs(net_variance)/max(abs(expected_total),1)*100:.2f}%','', 'red' if abs(net_variance or 0)>tolerance else 'green')]:
    c.markdown(f'<div style="text-align:center;border:1px solid #e3e8ef;padding:12px;margin:0 14px 12px 0;border-radius:6px"><div style="font-size:11px">{title}</div><div class="{cl}" style="font-size:20px;font-weight:800;margin-top:8px">{val}</div><div style="font-size:10px;color:#667085;margin-top:4px">{sub}</div></div>',unsafe_allow_html=True)

# Movement equation
st.markdown('<div class="movement">',unsafe_allow_html=True)
opening=s['_Opening'].sum() if s['_Opening'].notna().any() else expected_total-s['_Inward'].sum()+s['_Outward'].sum()
boxes=[('OPENING STOCK',opening,'#b7e1c0'),("TODAY'S INWARD",s['_Inward'].sum(),'#c8dcff'),("TODAY'S OUTWARD",s['_Outward'].sum(),'#ffdcb0'),('EXPECTED STOCK (As per System)',expected_total,'#ddd0ff'),('WAREHOUSE ACTUAL STOCK',actual_total,'#d9dde3'),('VARIANCE',net_variance,'#ffbaba')]
for i,(t,v,bg) in enumerate(boxes):
    st.markdown(f'<div class="movebox" style="border-color:{bg}"><div class="t">{t}</div><div class="v">{v:,.0f} KG</div></div>',unsafe_allow_html=True)
    if i<len(boxes)-1:st.markdown('<div class="eq">−</div>' if i==2 else '<div class="eq">+</div>' if i<3 else '<div class="eq">=</div>',unsafe_allow_html=True)
st.markdown('</div>',unsafe_allow_html=True)

if len(dis):st.markdown(f'<div class="alert">⚠ {len(dis)} material(s) have stock discrepancy. Click below to view details.<span style="background:#d72f2f;color:white;padding:7px 12px;border-radius:5px">View Discrepancies →</span></div>',unsafe_allow_html=True)
else:st.markdown('<div class="alert" style="background:#f0fdf4;border-color:#bbf7d0;color:#166534">✓ All materials reconciled within tolerance.</div>',unsafe_allow_html=True)
st.markdown('</div>',unsafe_allow_html=True)

# Discrepancy table
st.markdown('<div class="panel" style="margin-top:14px">',unsafe_allow_html=True)
st.markdown(f'<div class="section-title">DISCREPANT MATERIALS ({len(dis)})</div>',unsafe_allow_html=True)
view=s.copy();view['Opening Stock']=view['_Opening'];view['Inward']=view['_Inward'];view['Outward']=view['_Outward'];view['Expected Stock']=view['_Expected'];view['Warehouse Actual']=view['_WarehouseStock'];view['Variance']=view['_Variance'];view['Variance %']=view['_VariancePct'];view['Status']=view['Status'];view=view.sort_values('Variance',key=lambda x:x.abs(),ascending=False)
if view.empty:st.success('No materials found.')
else:
    rows=[]
    for _,r in view.head(100).iterrows():
        bad=abs(r['Variance'])>tolerance if pd.notna(r['Variance']) else False
        rows.append(f"<tr class={'bad' if bad else ''}><td>{len(rows)+1}</td><td><b>{clean(r['_Material'])}</b></td><td>{r['Opening Stock']:,.0f}</td><td>{r['Inward']:,.0f}</td><td>{r['Outward']:,.0f}</td><td>{r['Expected Stock']:,.0f}</td><td>{r['Warehouse Actual']:,.0f}</td><td class={'varbad' if bad else 'varok'}>{r['Variance']:,.0f}</td><td>{r['Variance %']:.2f}%</td><td><span class={'badge' if bad else 'badge ok'}>{'STOCK DISCREPANCY' if bad else 'MATCHED'}</span></td></tr>")
    html='''<div class="table-wrap"><table class="dis-table"><thead><tr><th>#</th><th>Product Name</th><th>Opening Stock</th><th>Inward</th><th>Outward</th><th>Expected Stock<br><small>System A + B − C</small></th><th>Warehouse Actual Stock</th><th>Variance<br>Actual − Expected</th><th>Variance %</th><th>Discrepancy Type</th></tr></thead><tbody>'''+''.join(rows)+'''</tbody></table></div>'''
    st.markdown(html,unsafe_allow_html=True)
st.markdown('</div>',unsafe_allow_html=True)

# Bottom visual cards
b1,b2,b3,b4=st.columns(4)
with b1:
    st.markdown('<div class="bottom-card"><div class="bottom-head">DISCREPANCY TYPES</div>',unsafe_allow_html=True)
    stockdisc=len(dis); match=len(matched); total=max(stockdisc+match,1)
    st.markdown(f'<div style="padding:12px;text-align:center"><div style="width:118px;height:118px;border-radius:50%;background:conic-gradient(#e34a4a {stockdisc/total*360}deg,#3aa76d 0);margin:auto;position:relative"><div style="position:absolute;inset:28px;background:white;border-radius:50%;padding-top:20px;font-size:12px;font-weight:800">{stockdisc+match}<br><span style="font-weight:400">Total</span></div></div><div style="font-size:11px;margin-top:10px">🔴 Stock Discrepancy &nbsp; {stockdisc} ({stockdisc/max(total,1)*100:.0f}%)<br>🟢 Matched &nbsp; {match} ({match/max(total,1)*100:.0f}%)</div></div></div>',unsafe_allow_html=True)
with b2:
    st.markdown('<div class="bottom-card"><div class="bottom-head">TOP DISCREPANCIES <span style="font-weight:400">(By Variance KG)</span></div><div class="bottom-body">',unsafe_allow_html=True)
    for i,(_,r) in enumerate(dis.reindex(dis['_Variance'].abs().sort_values(ascending=False).index).head(5).iterrows(),1):st.markdown(f'<div class="rank"><span>{i} &nbsp; {clean(r["_Material"])[:34]}</span><b class="red">{r["_Variance"]:,.0f} KG</b></div>',unsafe_allow_html=True)
    st.markdown('<div class="footlink">View all discrepancies →</div></div>',unsafe_allow_html=True)
with b3:
    st.markdown('<div class="bottom-card"><div class="bottom-head red">LOW STOCK ALERT <span style="font-weight:400">(Below 250 KG)</span></div><div class="bottom-body">',unsafe_allow_html=True)
    low=s[(s['_WarehouseStock']>0)&(s['_WarehouseStock']<low_stock)].sort_values('_WarehouseStock').head(5)
    for i,(_,r) in enumerate(low.iterrows(),1):st.markdown(f'<div class="rank"><span>{i} &nbsp; {clean(r["_Material"])[:34]}</span><b>{r["_WarehouseStock"]:,.0f} KG</b></div>',unsafe_allow_html=True)
    st.markdown('<div class="footlink">View all low stock →</div></div>',unsafe_allow_html=True)
with b4:
    st.markdown('<div class="bottom-card"><div class="bottom-head">AGEING ALERT <span style="font-weight:400">(Above 60 Days)</span></div><div class="bottom-body">',unsafe_allow_html=True)
    age=s[s['_Age']>ageing_limit]['_Age']; bins=[int((s['_Age']<=30).sum()),int(((s['_Age']>30)&(s['_Age']<=60)).sum()),int(((s['_Age']>60)&(s['_Age']<=90)).sum()),int((s['_Age']>90).sum())]
    st.markdown(f'<div style="display:flex;gap:12px;align-items:center;padding:8px"><div style="width:105px;height:105px;border-radius:50%;background:conic-gradient(#36a45c 0 31%,#f2b134 31% 58%,#f08c24 58% 76%,#e34a4a 76%);position:relative"><div style="position:absolute;inset:28px;background:white;border-radius:50%;text-align:center;padding-top:17px;font-size:13px;font-weight:800">{len(s)}<br><span style="font-weight:400">Materials</span></div></div><div style="font-size:10px;line-height:1.9">🟢 0–30 Days &nbsp; {bins[0]}<br>🟡 31–60 Days &nbsp; {bins[1]}<br>🟠 61–90 Days &nbsp; {bins[2]}<br>🔴 90+ Days &nbsp; {bins[3]}</div></div><div class="footlink">View ageing analysis →</div></div>',unsafe_allow_html=True)

st.markdown('<div style="height:12px"></div>',unsafe_allow_html=True)
# Details tabs are kept below the reference-style dashboard rather than replacing it.
t1,t2,t3,t4=st.tabs(['📊 Reports','↥ Inward','↧ Outward','📸 Daily Snapshot'])
with t1:
    st.caption('The dashboard uses the explicit expected/AG MIS stock when available. If it is absent, Expected Stock falls back to Opening + Inward − Outward.')
    if not dis.empty:
        st.dataframe(dis[['_Material','_Opening','_Inward','_Outward','_Expected','_WarehouseStock','_Variance','_VariancePct']].rename(columns={'_Material':'Product','_Opening':'Opening','_Inward':'Inward','_Outward':'Outward','_Expected':'Expected','_WarehouseStock':'Actual','_Variance':'Variance','_VariancePct':'Variance %'}),use_container_width=True,hide_index=True)
with t2: st.dataframe(iv,use_container_width=True,hide_index=True,height=400)
with t3: st.dataframe(ov,use_container_width=True,hide_index=True,height=400)
with t4:
    st.write('Snapshot saved:',saved, 'for',report_date)
    dates=list_snapshots()
    if dates:
        selected=st.selectbox('Open snapshot',dates,index=len(dates)-1);sh=read_snapshot(selected);st.dataframe(sh,use_container_width=True,hide_index=True,height=380);st.download_button('Download Snapshot',sh.to_csv(index=False).encode(),f'ATOM_GRID_snapshot_{selected}.csv','text/csv')

st.caption('ATOM GRID Warehouse MIS • V8 • Stock discrepancy is the primary control. Keep the daily snapshot during the validation period.')
