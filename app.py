import streamlit as st
import pandas as pd
import numpy as np
import io, re, os, json
from datetime import date

st.set_page_config(page_title="ATOM GRID | V6 Reconciliation", page_icon="📦", layout="wide")

LOW_STOCK=250.0
VAR_TOL=0.01
SNAP_DIR="snapshots"
os.makedirs(SNAP_DIR,exist_ok=True)

st.markdown("""<style>
.block-container{padding-top:1rem;max-width:1500px}
[data-testid="stMetric"]{border:1px solid #e5e7eb;border-radius:14px;padding:12px}
.bad{background:#fff0f0;border:1px solid #f3c2c0;border-radius:12px;padding:12px 16px;font-weight:600}
.good{background:#eaf8f0;border:1px solid #bfe8d0;border-radius:12px;padding:12px 16px;font-weight:600}
</style>""",unsafe_allow_html=True)

def norm(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+"," ",str(x).strip()).upper()
def clean(x): return re.sub(r"\s+"," ",str(x).replace("\n"," ").replace("\r"," ")).strip()
def num(df,c):
    return pd.to_numeric(df[c],errors="coerce").fillna(0) if c else pd.Series(0,index=df.index)
def find_header(raw):
    n=min(20,len(raw)); return max(range(n),key=lambda i:raw.iloc[i].notna().sum())
@st.cache_data(show_spinner=False)
def load_excel(data):
    xls=pd.ExcelFile(io.BytesIO(data)); out={}
    for sh in xls.sheet_names:
        raw=pd.read_excel(io.BytesIO(data),sheet_name=sh,header=None)
        h=find_header(raw)
        df=pd.read_excel(io.BytesIO(data),sheet_name=sh,header=h)
        df.columns=[clean(c) for c in df.columns]
        out[sh]=df.dropna(how="all").copy()
    return out
def pick(df,names):
    if df is None:return None
    ex={clean(c).lower():c for c in df.columns}
    for n in names:
        if clean(n).lower() in ex:return ex[clean(n).lower()]
    for c in df.columns:
        if any(clean(n).lower() in clean(c).lower() for n in names):return c
    return None
def snapfile(d): return os.path.join(SNAP_DIR,f"{d}.csv")
def snapshots():
    return sorted([f[:-4] for f in os.listdir(SNAP_DIR) if f.endswith(".csv")])
def read_snap(d):
    return pd.read_csv(snapfile(d)) if os.path.exists(snapfile(d)) else pd.DataFrame()
def save_snap(df,d,source):
    p=snapfile(d)
    if os.path.exists(p): return False
    x=df.copy(); x["Snapshot Date"]=str(d); x["Source File"]=source
    x.to_csv(p,index=False); return True

st.sidebar.title("📦 ATOM GRID MIS")
uploaded=st.sidebar.file_uploader("Upload today's warehouse MIS",type=["xlsx","xls"])
tol=st.sidebar.number_input("Discrepancy tolerance (kg)",0.0,100000.0,VAR_TOL,0.01)
low=st.sidebar.number_input("Low stock threshold (kg)",0.0,1000000.0,LOW_STOCK,50.0)

if not uploaded:
    st.title("📦 ATOM GRID — Warehouse Reconciliation V6")
    st.markdown("### Stock discrepancy is the primary control.")
    st.info("Upload the warehouse Excel. V6 checks physical variance and, when dated movement data plus a previous snapshot are available, calculates Opening + Inward − Outward = Expected Stock.")
    st.stop()

files=load_excel(uploaded.getvalue())
def sheet(*names):
    for n in names:
        if n in files:return files[n]
stock=sheet("Stock in Hand","Stock In Hand","Stock")
inward=sheet("Inward")
outward=sheet("Outward")
if stock is None:
    st.error("Stock in Hand sheet not found."); st.stop()

# Stock fields
sd=pick(stock,["Material Description","Description","Material Name","Product Name"])
sc=pick(stock,["Material Code","M.Code","Code"])
sb=pick(stock,["Batch No","Batch NO","Batch","Lot No","Lot"])
ss=pick(stock,["Closing Stock","Closing Stock kg","System Stock"])
sp=pick(stock,["Physical Closing Stock","Physical Stock kg","Physical Qty"])
svr=pick(stock,["Diffrence Excel Vs Phy","Difference Excel Vs Phy","Physical Difference","Variance"])
sa=pick(stock,["Warehouse Ageing","Ageing Days","Ageing"])
sdt=pick(stock,["Report Date","EOD Date","As On Date","Date"])

# Movement fields
def movement(df,kind):
    if df is None:return pd.DataFrame(columns=["_Material","_Code","_Batch","_Qty","_Date"])
    d=pick(df,["Material Description","Description","Material Name","Product Name"])
    c=pick(df,["Material Code","M.Code","Code"])
    b=pick(df,["Batch No","Batch NO","Batch","Lot No","Lot"])
    q=pick(df,["Physical Net Quantity Received","Physical Net Quantity Dispatched","Physical DISPATCHED Qty","Net Quantity as per Invoice","Received Quantity","Inward Quantity","Dispatched Quantity","Outward Quantity","Quantity","Qty"])
    dt=pick(df,["Inward Date","Outward Date","Dispatch Date","Receipt Date","Posting Date","Date"])
    x=pd.DataFrame(index=df.index)
    x["_Material"]=df[d].astype(str).str.strip() if d else ""
    x["_Code"]=df[c].map(norm) if c else ""
    x["_Batch"]=df[b].astype(str).str.strip() if b else ""
    x["_Qty"]=num(df,q) if q else 0
    x["_Date"]=pd.to_datetime(df[dt],errors="coerce",dayfirst=True) if dt else pd.NaT
    return x
iv=movement(inward,"in")
ov=movement(outward,"out")

calc=stock.copy()
calc["_Material"]=calc[sd].astype(str).str.strip() if sd else ""
calc["_Code"]=calc[sc].map(norm) if sc else ""
calc["_Batch"]=calc[sb].astype(str).str.strip() if sb else ""
calc["_System"]=num(calc,ss)
calc["_Physical"]=num(calc,sp)
calc["_PhysicalVariance"]=num(calc,svr) if svr else calc["_Physical"]-calc["_System"]
calc["_Age"]=num(calc,sa)
calc=calc[calc["_Material"].str.strip().ne("")].copy()

rd=pd.to_datetime(calc[sdt],errors="coerce",dayfirst=True).dropna() if sdt else pd.Series(dtype="datetime64[ns]")
report_date=rd.max().date() if len(rd) else date.today()
key=["_Material","_Code","_Batch"]

def dayagg(x,dt,qty):
    if x.empty:return pd.DataFrame(columns=key+[qty])
    y=x[x["_Date"].notna() & (x["_Date"].dt.date==report_date)]
    return y.groupby(key)["_Qty"].sum().reset_index().rename(columns={"_Qty":qty})
din=dayagg(iv,None,"_Inward")
dout=dayagg(ov,None,"_Outward")
calc=calc.merge(din,on=key,how="left").merge(dout,on=key,how="left")
calc["_Inward"]=calc["_Inward"].fillna(0); calc["_Outward"]=calc["_Outward"].fillna(0)

# Previous snapshot supplies opening stock. Only movement data with actual dates is used.
snaps=snapshots(); prev_dates=[d for d in snaps if d<str(report_date)]
prev=read_snap(max(prev_dates)) if prev_dates else pd.DataFrame()
movement_ready=bool(not prev.empty and (not iv.empty and iv["_Date"].notna().any()) and (not ov.empty and ov["_Date"].notna().any()))
if movement_ready:
    pk=prev[key+["_System"]].drop_duplicates(key).rename(columns={"_System":"_Opening"})
    calc=calc.merge(pk,on=key,how="left")
    calc["_Opening"]=calc["_Opening"].fillna(0)
    calc["_Expected"] = calc["_Opening"] + calc["_Inward"] - calc["_Outward"]
    calc["_MovementVariance"]=calc["_Expected"]-calc["_System"]
else:
    calc["_Opening"]=np.nan; calc["_Expected"]=np.nan; calc["_MovementVariance"]=np.nan

calc["_PhysicalVariance"]=calc["_PhysicalVariance"].fillna(calc["_Physical"]-calc["_System"])
calc["Status"]=np.select([
    calc["_PhysicalVariance"].abs()>tol,
    movement_ready & (calc["_MovementVariance"].abs()>tol),
    (calc["_System"]>0)&(calc["_System"]<low),
    calc["_Age"]>=90],
    ["🔴 PHYSICAL DISCREPANCY","🟣 MOVEMENT DISCREPANCY","🟠 LOW STOCK","🟡 AGEING"],default="🟢 MATCHED")

# Save immutable daily snapshot for validation.
snapcols=key+["_System","_Physical","_PhysicalVariance","_Age","_Opening","_Inward","_Outward","_Expected","_MovementVariance","Status"]
saved=save_snap(calc[snapcols],report_date,uploaded.name)
snaps=snapshots()

# Header / KPIs
st.title("📦 ATOM GRID — Stock Reconciliation")
st.caption(f"EOD {report_date.strftime('%d-%b-%Y')} • {uploaded.name}")
phys=calc[calc["_PhysicalVariance"].abs()>tol]
mov=calc[movement_ready & (calc["_MovementVariance"].abs()>tol)]
matched=calc[(calc["_PhysicalVariance"].abs()<=tol) & (~movement_ready | (calc["_MovementVariance"].abs()<=tol))]
net=calc["_PhysicalVariance"].sum(); absolute=calc["_PhysicalVariance"].abs().sum()

if len(phys)==0:
    st.markdown('<div class="good">✅ Physical stock reconciled within tolerance.</div>',unsafe_allow_html=True)
else:
    st.markdown(f'<div class="bad">🔴 {len(phys):,} physical discrepancy line(s) require investigation.</div>',unsafe_allow_html=True)

k=st.columns(6)
for c,l,v in [(k[0],"🔴 Physical Discrepancies",len(phys)),(k[1],"🟣 Movement Issues",len(mov)),(k[2],"🟢 Matched",len(matched)),(k[3],"Net Variance KG",net),(k[4],"Absolute Variance KG",absolute),(k[5],"Stock Checked KG",calc["_System"].sum())]:
    c.metric(l,f"{v:,.2f}" if isinstance(v,(float,np.floating)) else f"{v:,}")

tabs=st.tabs(["🔴 Discrepancy Control","📸 Daily Snapshot","🔄 Movement Reconciliation","📦 Inventory","⏳ Ageing","📄 EOD Report"])

with tabs[0]:
    st.subheader("Stock discrepancy — priority view")
    if phys.empty: st.success("No physical discrepancies above tolerance.")
    else:
        g=phys.groupby("_Material",as_index=False).agg(System_Stock_KG=("_System","sum"),Physical_Stock_KG=("_Physical","sum"),Variance_KG=("_PhysicalVariance","sum"))
        g["Variance_%"]=np.where(g.System_Stock_KG!=0,g.Variance_KG/g.System_Stock_KG*100,0)
        g=g.sort_values("Variance_KG",key=lambda x:x.abs(),ascending=False)
        a,b=st.columns([1.1,1])
        with a: st.bar_chart(g.head(15).set_index("_Material")["Variance_KG"],height=360)
        with b: st.dataframe(g.head(15),use_container_width=True,hide_index=True,height=360)
        detail=phys.copy()
        detail["Direction"]=np.where(detail["_PhysicalVariance"]>0,"Physical > System","Physical < System")
        cols=[c for c in [sd,sc,sb,ss,sp,svr,sa] if c]
        st.dataframe(detail[cols+["Direction"]],use_container_width=True,hide_index=True,height=430)
        st.download_button("Download physical discrepancies",detail[cols+["Direction"]].to_csv(index=False).encode(),"ATOM_GRID_physical_discrepancies.csv","text/csv")

with tabs[1]:
    st.subheader("📸 Daily reconciliation snapshots")
    st.write("The dashboard saves one snapshot per EOD date and never overwrites an existing snapshot. Use this for the first 7–14 days of manual validation.")
    if saved: st.success(f"Snapshot saved for {report_date}.")
    elif str(report_date) in snaps: st.info(f"Snapshot for {report_date} already exists and was not overwritten.")
    st.write("Saved dates:",", ".join(snaps) if snaps else "None")
    if snaps:
        pickdate=st.selectbox("Open snapshot",snaps,index=len(snaps)-1)
        sh=read_snap(pickdate)
        st.dataframe(sh,use_container_width=True,hide_index=True,height=450)
        st.download_button("Download snapshot",sh.to_csv(index=False).encode(),f"ATOM_GRID_snapshot_{pickdate}.csv","text/csv")
    st.caption("Compare the selected snapshot with the warehouse's manually prepared EOD report. Keep both during the validation period.")

with tabs[2]:
    st.subheader("🔄 Movement-aware stock reconciliation")
    if not movement_ready:
        st.warning("Movement reconciliation is waiting for: (1) a previous saved snapshot and (2) dated Inward and Outward records. Once available, the dashboard will calculate Opening + Inward − Outward and compare it with system stock.")
    else:
        mv=calc.copy()
        mv["Opening Stock"]=mv["_Opening"]; mv["Today's Inward"]=mv["_Inward"]; mv["Today's Outward"]=mv["_Outward"]
        mv["Expected Stock"]=mv["_Expected"]; mv["Warehouse/System Stock"]=mv["_System"]; mv["Variance"]=mv["_MovementVariance"]
        st.dataframe(mv[key+["Opening Stock","Today's Inward","Today's Outward","Expected Stock","Warehouse/System Stock","Variance","Status"]].sort_values("Variance",key=lambda x:x.abs(),ascending=False),use_container_width=True,hide_index=True,height=500)
    c=st.columns(3)
    c[0].metric("Today's Inward KG",f"{calc['_Inward'].sum():,.2f}")
    c[1].metric("Today's Outward KG",f"{calc['_Outward'].sum():,.2f}")
    c[2].metric("Net Movement KG",f"{calc['_Inward'].sum()-calc['_Outward'].sum():,.2f}")

with tabs[3]:
    st.subheader("📦 Inventory")
    search=st.text_input("Search material")
    f=calc.copy()
    if search:f=f[f["_Material"].str.contains(search,case=False,na=False)]
    st.dataframe(f[key+["_System","_Physical","_PhysicalVariance","_Age","Status"]],use_container_width=True,hide_index=True,height=520)

with tabs[4]:
    st.subheader("⏳ Ageing")
    b=pd.cut(calc["_Age"],[-1,30,60,90,np.inf],labels=["0–30","31–60","61–90","90+"])
    ag=calc.assign(Bucket=b).groupby("Bucket",observed=False).agg(Lines=("_Material","count"),Stock_KG=("_System","sum")).reset_index()
    st.dataframe(ag,use_container_width=True,hide_index=True)
    st.bar_chart(ag.set_index("Bucket")["Stock_KG"])

with tabs[5]:
    st.subheader(f"📄 Inventory Report EOD Till {report_date.strftime('%d-%b-%Y')}")
    e=calc.groupby("_Material",as_index=False).agg(Inbound=("_Inward","sum"),Outbound=("_Outward","sum"),Instock=("_System","sum"),Variance=("_PhysicalVariance","sum"))
    e=e.rename(columns={"_Material":"Product Name"})
    st.dataframe(e,use_container_width=True,hide_index=True)
    st.download_button("Download EOD report",e.to_csv(index=False).encode(),f"ATOM_GRID_EOD_{report_date}.csv","text/csv")

st.divider()
st.caption("ATOM GRID Warehouse MIS V6 • No authentication • Stock discrepancy first • Movement-aware • Daily snapshots")
