
import streamlit as st
import pandas as pd
import numpy as np
import io, re, hashlib
from datetime import datetime

st.set_page_config(page_title="ATOM GRID | Warehouse MIS", page_icon="📦", layout="wide")

# ----------------------------
# Theme / helpers
# ----------------------------
st.markdown("""
<style>
.block-container {padding-top:1.1rem; padding-bottom:2rem;}
h1 {margin-bottom:0.1rem;}
[data-testid="stMetric"] {border:1px solid #d9e1f2;border-radius:12px;padding:10px 14px;}
.small {font-size:.82rem;color:#667085;}
</style>
""", unsafe_allow_html=True)

LOW_STOCK = 250
AGEING = 90
VAR_TOL = 0.01

def norm(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+", " ", str(x).strip()).upper()

def header_row(raw):
    n=min(15,len(raw))
    counts=[raw.iloc[i].notna().sum() for i in range(n)]
    return int(max(range(n), key=lambda i: counts[i]))

@st.cache_data(show_spinner=False)
def load_excel(file_bytes):
    xls=pd.ExcelFile(io.BytesIO(file_bytes))
    out={}
    for sh in xls.sheet_names:
        raw=pd.read_excel(io.BytesIO(file_bytes),sheet_name=sh,header=None)
        hr=header_row(raw)
        df=pd.read_excel(io.BytesIO(file_bytes),sheet_name=sh,header=hr)
        df.columns=[re.sub(r"\s+"," ",str(c).replace("\n"," ").replace("\r"," ")).strip() for c in df.columns]
        df=df.dropna(how="all").copy()
        out[sh]=df
    return out

def pick(df, candidates):
    if df is None: return None
    lookup={str(c).strip().lower():c for c in df.columns}
    for c in candidates:
        if c.lower() in lookup: return lookup[c.lower()]
    # fuzzy fallback
    for col in df.columns:
        lc=str(col).lower()
        if any(c.lower() in lc for c in candidates):
            return col
    return None

def numeric(df,col):
    if df is None or col is None: return pd.Series(0,index=df.index)
    return pd.to_numeric(df[col],errors="coerce").fillna(0)

def date_col(df, candidates):
    c=pick(df,candidates)
    if c is None: return None
    return pd.to_datetime(df[c],errors="coerce",dayfirst=True)

def row_hash(df):
    if df is None or len(df)==0: return pd.Series(dtype=str)
    return df.astype(str).fillna("").agg("|".join,axis=1).map(lambda s: hashlib.sha1(s.encode()).hexdigest())

# ----------------------------
# Optional company login
# ----------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated=False

try:
    oidc_user=st.user
    if getattr(oidc_user,"is_logged_in",False):
        email=str(getattr(oidc_user,"email","")).lower().strip()
        if email.endswith("@atomgrid.in"):
            st.session_state.authenticated=True
            st.session_state.user_email=email
        else:
            st.error("Access restricted to authorised @atomgrid.in accounts.")
            if st.button("Sign out"):
                st.logout()
            st.stop()
except Exception:
    # Local testing mode: no OIDC secrets required.
    pass

if not st.session_state.authenticated:
    st.title("📦 ATOM GRID — Warehouse MIS")
    st.write("Live warehouse visibility from the latest MIS Excel.")
    st.info("For the published version, sign in with your authorised @atomgrid.in Google account.")
    try:
        if st.button("Sign in with Google", type="primary"):
            st.login()
    except Exception:
        st.caption("Local testing mode is active. The dashboard will be available without login.")
    # Local mode continues; production OIDC can be enforced by setting AUTH_REQUIRED=true in secrets.
    try:
        if st.secrets.get("AUTH_REQUIRED", False):
            st.stop()
    except Exception:
        pass

# ----------------------------
# Sidebar
# ----------------------------
st.sidebar.title("ATOM GRID MIS")
uploaded=st.sidebar.file_uploader("Upload latest warehouse MIS",type=["xlsx","xls"])
master_file=st.sidebar.file_uploader("Upload approved Material Master (optional)",type=["xlsx","xls","csv"])
st.sidebar.divider()
st.sidebar.caption("Controls")
low_stock=float(st.sidebar.number_input("Low stock threshold (kg)",value=250.0,min_value=0.0))
ageing_threshold=float(st.sidebar.number_input("Ageing threshold (days)",value=90.0,min_value=0.0))
variance_tol=float(st.sidebar.number_input("Physical variance tolerance (kg)",value=0.01,min_value=0.0))

if not uploaded:
    st.title("📦 ATOM GRID — Warehouse MIS")
    st.subheader("Upload the latest warehouse MIS to begin")
    st.markdown("""
    **What you get**
    - Current stock and physical reconciliation
    - Inward / outward visibility
    - Material and batch drill-down
    - Ageing buckets
    - Low-stock alerts
    - New / unmapped material control
    - Pending orders, vehicles and cancelled invoices
    - Downloadable exception reports
    """)
    st.stop()

files=load_excel(uploaded.getvalue())
st.title("📦 ATOM GRID — Warehouse MIS")
st.caption(f"**{uploaded.name}**  •  {len(files)} sheets loaded  •  Updated in this browser session")

# Source sheets with name fallbacks
def sheet(name, fallbacks=[]):
    if name in files: return files[name]
    for x in fallbacks:
        if x in files: return files[x]
    return None

inward=sheet("Inward")
outward=sheet("Outward")
stock=sheet("Stock in Hand",["Stock in hand","Stock"])
pending=sheet("Pending Order Report",["Pending Orders"])
vehicle=sheet("Vehicle Indent")
cancelled=sheet("Cancelled Invoice Report",["Cancelled Invoices"])

if stock is None:
    st.error("The uploaded MIS does not contain a 'Stock in Hand' sheet.")
    st.stop()

# Material master
master=None
if master_file:
    if master_file.name.lower().endswith(".csv"):
        master=pd.read_csv(master_file)
    else:
        mf=load_excel(master_file.getvalue())
        master=mf.get("Material_Master") or next(iter(mf.values()))
    master.columns=[re.sub(r"\s+"," ",str(c).strip()) for c in master.columns]

# ----------------------------
# Normalize stock
# ----------------------------
c_desc=pick(stock,["Material Description","Description","Material Name"])
c_code=pick(stock,["Material Code","M.Code","Code"])
c_batch=pick(stock,["Batch No","Batch NO","Batch"])
c_close=pick(stock,["Closing Stock","Closing Stock kg"])
c_phys=pick(stock,["Physical Closing Stock","Physical Stock kg"])
c_diff=pick(stock,["Diffrence Excel Vs Phy","Difference Excel Vs Phy","Physical Difference kg","Variance"])
c_age=pick(stock,["Warehouse Ageing","Ageing Days","Ageing"])
c_gr=pick(stock,["GR No","GR NO"])
c_loc=pick(stock,["Storage Location","Location"])
c_date=pick(stock,["Date","Stock Date","GR Date"])

sv=stock.copy()
sv["_Material"]=sv[c_desc].astype(str).str.strip() if c_desc else ""
sv["_Code"]=sv[c_code].map(norm) if c_code else ""
sv["_Batch"]=sv[c_batch].map(norm) if c_batch else ""
sv["_Closing"]=numeric(sv,c_close)
sv["_Physical"]=numeric(sv,c_phys)
sv["_Variance"]=numeric(sv,c_diff)
sv["_Age"]=numeric(sv,c_age)
sv["_Date"]=pd.to_datetime(sv[c_date],errors="coerce",dayfirst=True) if c_date else pd.NaT

# ----------------------------
# Sidebar filters
# ----------------------------
materials=sorted([x for x in sv["_Material"].unique() if str(x).strip()])
sel_m=st.sidebar.multiselect("Material",materials)
if sel_m: sv=sv[sv["_Material"].isin(sel_m)].copy()

if c_loc:
    locs=sorted([x for x in sv[c_loc].dropna().astype(str).unique() if x.strip()])
    sel_l=st.sidebar.multiselect("Storage Location",locs)
    if sel_l: sv=sv[sv[c_loc].astype(str).isin(sel_l)].copy()

# ----------------------------
# KPI row
# ----------------------------
mismatch=int((sv["_Variance"].abs()>variance_tol).sum())
low=int(((sv["_Closing"]>0)&(sv["_Closing"]<low_stock)).sum())
old=int((sv["_Age"]>=ageing_threshold).sum())

in_qty=0; out_qty=0
if inward is not None:
    iq=pick(inward,["Physical Net Quantity Received","Net Quantity as per Invoice","Received Quantity","Quantity"])
    in_qty=float(numeric(inward,iq).sum()) if iq else 0
if outward is not None:
    oq=pick(outward,["Physical Net Quantity Dispatched","Physical DISPATCHED Qty","Net Quantity as per Invoice","Dispatched Quantity","Quantity"])
    out_qty=float(numeric(outward,oq).sum()) if oq else 0

k=st.columns(8)
for col,(label,val) in zip(k,[
    ("System Stock (kg)",sv["_Closing"].sum()),
    ("Physical Stock (kg)",sv["_Physical"].sum()),
    ("Variance (kg)",sv["_Variance"].sum()),
    ("Stock Lines",len(sv)),
    ("Low Stock",low),
    (f"{int(ageing_threshold)}+ Days",old),
    ("Inward Qty",in_qty),
    ("Outward Qty",out_qty)
]):
    col.metric(label, f"{val:,.2f}" if isinstance(val,(float,np.floating)) else f"{val:,}")

# ----------------------------
# Tabs
# ----------------------------
tabs=st.tabs(["Overview","Stock","Inward","Outward","Ageing","Exceptions","New Materials","Pending Orders","Vehicles","Cancelled Invoices"])

with tabs[0]:
    left,right=st.columns(2)
    with left:
        st.subheader("Stock by Material")
        sm=sv.groupby("_Material")["_Closing"].sum().sort_values(ascending=False).head(20)
        st.bar_chart(sm)
    with right:
        st.subheader("Ageing distribution")
        b=pd.cut(sv["_Age"],[-1,30,60,90,np.inf],labels=["0–30","31–60","61–90","90+"])
        ab=sv.assign(Bucket=b).groupby("Bucket",observed=False)["_Closing"].sum()
        st.bar_chart(ab)
    st.subheader("Quick status")
    status=pd.Series({
        "Physical mismatches": mismatch,
        "Low stock lines": low,
        f"{int(ageing_threshold)}+ ageing lines": old
    })
    st.dataframe(status.rename("Count").to_frame(),use_container_width=True)
    st.caption("Use the tabs above for detailed drill-down and downloads.")

with tabs[1]:
    f=sv.copy()
    f["Status"]=np.select([
        f["_Variance"].abs()>variance_tol,
        (f["_Closing"]>0)&(f["_Closing"]<low_stock),
        f["_Age"]>=ageing_threshold],
        ["PHYSICAL MISMATCH","LOW STOCK","AGEING"],default="OK")
    cols=[c for c in [c_gr,c_code,c_desc,c_batch,c_loc,c_close,c_phys,c_diff,c_age,c_date] if c]
    st.dataframe(f[cols+["Status"]],use_container_width=True,hide_index=True)
    st.download_button("Download filtered stock",f.to_csv(index=False).encode(),"ATOM_GRID_current_stock.csv","text/csv")

with tabs[2]:
    if inward is not None:
        st.dataframe(inward,use_container_width=True,hide_index=True)
        st.download_button("Download Inward CSV",inward.to_csv(index=False).encode(),"ATOM_GRID_inward.csv","text/csv")
    else: st.warning("Inward sheet not found.")

with tabs[3]:
    if outward is not None:
        st.dataframe(outward,use_container_width=True,hide_index=True)
        st.download_button("Download Outward CSV",outward.to_csv(index=False).encode(),"ATOM_GRID_outward.csv","text/csv")
    else: st.warning("Outward sheet not found.")

with tabs[4]:
    aa=sv.copy()
    aa["Ageing Bucket"]=pd.cut(aa["_Age"],[-1,30,60,90,np.inf],labels=["0–30","31–60","61–90","90+"])
    summary=aa.groupby("Ageing Bucket",observed=False).agg(Stock_Lines=("_Closing","count"),Stock_KG=("_Closing","sum")).reset_index()
    st.dataframe(summary,use_container_width=True,hide_index=True)
    st.bar_chart(summary.set_index("Ageing Bucket")["Stock_KG"])
    olddf=aa[aa["_Age"]>=ageing_threshold]
    st.subheader(f"Ageing ≥ {int(ageing_threshold)} days")
    st.dataframe(olddf[[c for c in [c_gr,c_desc,c_batch,c_close,c_age] if c]],use_container_width=True,hide_index=True)

with tabs[5]:
    rows=[]
    for _,r in sv.iterrows():
        if abs(r["_Variance"])>variance_tol:
            rows.append(["PHYSICAL MISMATCH",r.get(c_gr,""),r["_Material"],r.get(c_batch,""),r["_Variance"],"Verify physical stock."])
        if 0<r["_Closing"]<low_stock:
            rows.append(["LOW STOCK",r.get(c_gr,""),r["_Material"],r.get(c_batch,""),r["_Closing"],"Review requirement / consumption."])
        if r["_Age"]>=ageing_threshold:
            rows.append(["AGEING",r.get(c_gr,""),r["_Material"],r.get(c_batch,""),r["_Age"],"Prioritise consumption / ageing review."])
    ex=pd.DataFrame(rows,columns=["Exception","GR No","Material","Batch","Value","Action"])
    if ex.empty: st.success("No configured stock exceptions.")
    else:
        st.dataframe(ex,use_container_width=True,hide_index=True)
        st.download_button("Download exceptions",ex.to_csv(index=False).encode(),"ATOM_GRID_exceptions.csv","text/csv")

with tabs[6]:
    st.subheader("New / Unmapped Materials")
    if master is None:
        st.warning("Upload the approved Material Master from the sidebar to enable controlled mapping.")
        starter=sv[[c for c in [c_desc,c_code] if c]].drop_duplicates().copy()
        if c_desc: starter=starter.rename(columns={c_desc:"Warehouse Material Name"})
        if c_code: starter=starter.rename(columns={c_code:"Material Code"})
        starter["AG Catalogue Name"]=""
        starter["Status"]="ACTIVE"
        st.download_button("Download starter Material Master",starter.to_csv(index=False).encode(),"ATOM_GRID_Material_Master_Template.csv","text/csv")
    else:
        mh={str(c).strip().lower():c for c in master.columns}
        mn=mh.get("warehouse material name") or mh.get("material") or mh.get("description")
        mc=mh.get("material code") or mh.get("m.code") or mh.get("code")
        if not mn:
            st.error("Material Master must contain 'Warehouse Material Name'.")
        else:
            known=set(master[mn].dropna().map(norm))
            codes=set(master[mc].dropna().map(norm)) if mc else set()
            nm=sv.copy()
            nm["Mapping Status"]=np.where((nm["_Material"].map(norm).isin(known)) | (nm["_Code"].isin(codes)),"MAPPED","NEW / NOT MAPPED")
            new=nm[nm["Mapping Status"]=="NEW / NOT MAPPED"]
            st.metric("New / unmapped stock lines",len(new))
            show=[c for c in [c_gr,c_code,c_desc,c_batch,c_close,c_age] if c]+["Mapping Status"]
            st.dataframe(new[show],use_container_width=True,hide_index=True)
            st.info("Unknown materials are intentionally not auto-mapped. Add the approved mapping to your Material Master and upload it again.")

with tabs[7]:
    if pending is not None: st.dataframe(pending,use_container_width=True,hide_index=True)
    else: st.info("Pending Order Report not found.")

with tabs[8]:
    if vehicle is not None: st.dataframe(vehicle,use_container_width=True,hide_index=True)
    else: st.info("Vehicle Indent not found.")

with tabs[9]:
    if cancelled is not None: st.dataframe(cancelled,use_container_width=True,hide_index=True)
    else: st.info("Cancelled Invoice Report not found.")

st.divider()
st.caption("ATOM GRID Warehouse MIS • Upload-driven live visibility • Recommended production access: authorised @atomgrid.in accounts")
