import streamlit as st
import pandas as pd
import numpy as np
import io, re, os

st.set_page_config(page_title="ATOM GRID | Warehouse MIS", page_icon="📦", layout="wide")

# ---------- GOOGLE OIDC ACCESS ----------
# Google sign-in is required. Only @atomgrid.in accounts are permitted.
if not st.user.is_logged_in:
    st.title("📦 ATOM GRID")
    st.subheader("Warehouse MIS")
    st.write("Sign in with your ATOM GRID Google account to continue.")
    if st.button("Sign in with Google", type="primary", use_container_width=True):
        st.login("google")
    st.stop()

user_email = (getattr(st.user, "email", "") or "").strip().lower()
if not user_email.endswith("@atomgrid.in"):
    st.error("Access denied. This dashboard is restricted to @atomgrid.in accounts.")
    st.write(f"Signed-in account: {user_email or 'Unknown'}")
    if st.button("Sign out"):
        st.logout()
    st.stop()


LOW_STOCK=250
AGEING=90
VAR_TOL=0.01

def norm(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+"," ",str(x).strip()).upper()

def find_header(raw):
    n=min(15,len(raw))
    counts=[raw.iloc[i].notna().sum() for i in range(n)]
    return max(range(n),key=lambda i:counts[i])

@st.cache_data(show_spinner=False)
def load_excel(data):
    xls=pd.ExcelFile(io.BytesIO(data))
    out={}
    for sh in xls.sheet_names:
        raw=pd.read_excel(io.BytesIO(data),sheet_name=sh,header=None)
        hr=find_header(raw)
        df=pd.read_excel(io.BytesIO(data),sheet_name=sh,header=hr)
        df.columns=[re.sub(r"\s+"," ",str(c).replace("\n"," ").replace("\r"," ")).strip() for c in df.columns]
        out[sh]=df.dropna(how="all").copy()
    return out

def pick(df,names):
    if df is None:return None
    exact={str(c).strip().lower():c for c in df.columns}
    for n in names:
        if n.lower() in exact:return exact[n.lower()]
    for c in df.columns:
        if any(n.lower() in str(c).lower() for n in names):return c
    return None

def num(df,col):
    return pd.to_numeric(df[col],errors="coerce").fillna(0) if col else pd.Series(0,index=df.index)

# ---------- DASHBOARD ACCESS ----------
# No authentication: anyone with the Render URL can open and use the dashboard.

# ---------- SIDEBAR ----------
st.sidebar.title("ATOM GRID MIS")
st.sidebar.caption(f"Signed in: {user_email}")
if st.sidebar.button("Sign out"):
    st.logout()
st.sidebar.caption("User: "+st.session_state.get("user",""))
if st.sidebar.button("Logout"):
    st.session_state.logged_in=False
    st.rerun()

uploaded=st.sidebar.file_uploader("Upload latest warehouse MIS",type=["xlsx","xls"])
master_file=st.sidebar.file_uploader("Approved Material Master (optional)",type=["xlsx","xls","csv"])
st.sidebar.divider()
low_stock=st.sidebar.number_input("Low stock threshold (kg)",value=250.0,min_value=0.0)
ageing=st.sidebar.number_input("Ageing threshold (days)",value=90.0,min_value=0.0)
variance_tol=st.sidebar.number_input("Variance tolerance (kg)",value=0.01,min_value=0.0)

if not uploaded:
    st.title("📦 ATOM GRID — Warehouse MIS")
    st.write("Upload the latest warehouse MIS Excel from the sidebar.")
    st.info("The dashboard will process the file and provide stock, movement and exception visibility.")
    st.stop()

files=load_excel(uploaded.getvalue())

def sheet(name,*fallbacks):
    for n in (name,)+fallbacks:
        if n in files:return files[n]
    return None

inward=sheet("Inward")
outward=sheet("Outward")
stock=sheet("Stock in Hand","Stock in hand","Stock")
pending=sheet("Pending Order Report","Pending Orders")
vehicle=sheet("Vehicle Indent")
cancelled=sheet("Cancelled Invoice Report","Cancelled Invoices")

if stock is None:
    st.error("Stock in Hand sheet was not found.")
    st.stop()

# ---------- STOCK ----------
desc=pick(stock,["Material Description","Description","Material Name"])
code=pick(stock,["Material Code","M.Code","Code"])
batch=pick(stock,["Batch No","Batch NO","Batch"])
closing=pick(stock,["Closing Stock","Closing Stock kg"])
physical=pick(stock,["Physical Closing Stock","Physical Stock kg"])
difference=pick(stock,["Diffrence Excel Vs Phy","Difference Excel Vs Phy","Physical Difference","Variance"])
age_col=pick(stock,["Warehouse Ageing","Ageing Days","Ageing"])
gr=pick(stock,["GR No","GR NO"])
location=pick(stock,["Storage Location","Location"])

sv=stock.copy()
sv["_Material"]=sv[desc].astype(str).str.strip() if desc else ""
sv["_Closing"]=num(sv,closing)
sv["_Physical"]=num(sv,physical)
sv["_Variance"]=num(sv,difference)
sv["_Age"]=num(sv,age_col)
sv["_Code"]=sv[code].map(norm) if code else ""

# filters
mats=sorted([x for x in sv["_Material"].unique() if str(x).strip()])
selected=st.sidebar.multiselect("Material",mats)
if selected:sv=sv[sv["_Material"].isin(selected)].copy()

if location:
    locs=sorted([x for x in sv[location].dropna().astype(str).unique() if x.strip()])
    selected_l=st.sidebar.multiselect("Storage Location",locs)
    if selected_l:sv=sv[sv[location].astype(str).isin(selected_l)].copy()

mismatch=int((sv["_Variance"].abs()>variance_tol).sum())
low=int(((sv["_Closing"]>0)&(sv["_Closing"]<low_stock)).sum())
old=int((sv["_Age"]>=ageing).sum())

in_qty=out_qty=0
if inward is not None:
    c=pick(inward,["Physical Net Quantity Received","Net Quantity as per Invoice","Received Quantity","Quantity"])
    if c:in_qty=num(inward,c).sum()
if outward is not None:
    c=pick(outward,["Physical Net Quantity Dispatched","Physical DISPATCHED Qty","Net Quantity as per Invoice","Dispatched Quantity","Quantity"])
    if c:out_qty=num(outward,c).sum()

st.title("📦 ATOM GRID — Warehouse MIS")
st.caption(f"{uploaded.name} • {len(files)} sheets loaded")

k=st.columns(8)
for col,(label,value) in zip(k,[
    ("System Stock (kg)",sv["_Closing"].sum()),
    ("Physical Stock (kg)",sv["_Physical"].sum()),
    ("Variance (kg)",sv["_Variance"].sum()),
    ("Stock Lines",len(sv)),
    ("Low Stock",low),
    (f"{int(ageing)}+ Days",old),
    ("Inward Qty",in_qty),
    ("Outward Qty",out_qty)]):
    col.metric(label,f"{value:,.2f}" if isinstance(value,(float,np.floating)) else f"{value:,}")

tabs=st.tabs(["Overview","Stock","Inward","Outward","Ageing","Exceptions","New Materials","Pending Orders","Vehicles","Cancelled Invoices"])

with tabs[0]:
    a,b=st.columns(2)
    with a:
        st.subheader("Top materials by stock")
        st.bar_chart(sv.groupby("_Material")["_Closing"].sum().sort_values(ascending=False).head(20))
    with b:
        st.subheader("Ageing by stock")
        bucket=pd.cut(sv["_Age"],[-1,30,60,90,np.inf],labels=["0–30","31–60","61–90","90+"])
        st.bar_chart(sv.assign(Bucket=bucket).groupby("Bucket",observed=False)["_Closing"].sum())
    st.dataframe(pd.Series({"Physical mismatches":mismatch,"Low stock lines":low,f"{int(ageing)}+ ageing lines":old}).rename("Count").to_frame(),use_container_width=True)

with tabs[1]:
    f=sv.copy()
    f["Status"]=np.select([
        f["_Variance"].abs()>variance_tol,
        (f["_Closing"]>0)&(f["_Closing"]<low_stock),
        f["_Age"]>=ageing],
        ["PHYSICAL MISMATCH","LOW STOCK","AGEING"],default="OK")
    cols=[c for c in [gr,code,desc,batch,location,closing,physical,difference,age_col] if c]
    st.dataframe(f[cols+["Status"]],use_container_width=True,hide_index=True)
    st.download_button("Download filtered stock",f.to_csv(index=False).encode(),"ATOM_GRID_stock.csv","text/csv")

with tabs[2]:
    if inward is not None:
        st.dataframe(inward,use_container_width=True,hide_index=True)
        st.download_button("Download Inward",inward.to_csv(index=False).encode(),"ATOM_GRID_inward.csv","text/csv")
    else:st.warning("Inward sheet not found.")

with tabs[3]:
    if outward is not None:
        st.dataframe(outward,use_container_width=True,hide_index=True)
        st.download_button("Download Outward",outward.to_csv(index=False).encode(),"ATOM_GRID_outward.csv","text/csv")
    else:st.warning("Outward sheet not found.")

with tabs[4]:
    aa=sv.copy()
    aa["Ageing Bucket"]=pd.cut(aa["_Age"],[-1,30,60,90,np.inf],labels=["0–30","31–60","61–90","90+"])
    summary=aa.groupby("Ageing Bucket",observed=False).agg(Stock_Lines=("_Closing","count"),Stock_KG=("_Closing","sum")).reset_index()
    st.dataframe(summary,use_container_width=True,hide_index=True)
    st.bar_chart(summary.set_index("Ageing Bucket")["Stock_KG"])
    st.subheader(f"Stock aged {int(ageing)}+ days")
    st.dataframe(aa[aa["_Age"]>=ageing][[c for c in [gr,desc,batch,closing,age_col] if c]],use_container_width=True,hide_index=True)

with tabs[5]:
    rows=[]
    for _,r in sv.iterrows():
        if abs(r["_Variance"])>variance_tol:rows.append(["PHYSICAL MISMATCH",r.get(gr,""),r["_Material"],r.get(batch,""),r["_Variance"],"Verify physical stock"])
        if 0<r["_Closing"]<low_stock:rows.append(["LOW STOCK",r.get(gr,""),r["_Material"],r.get(batch,""),r["_Closing"],"Review requirement"])
        if r["_Age"]>=ageing:rows.append(["AGEING",r.get(gr,""),r["_Material"],r.get(batch,""),r["_Age"],"Prioritise ageing review"])
    ex=pd.DataFrame(rows,columns=["Exception","GR No","Material","Batch","Value","Action"])
    if ex.empty:st.success("No configured stock exceptions.")
    else:
        st.dataframe(ex,use_container_width=True,hide_index=True)
        st.download_button("Download exceptions",ex.to_csv(index=False).encode(),"ATOM_GRID_exceptions.csv","text/csv")

with tabs[6]:
    st.subheader("New / Unmapped Materials")
    master=None
    if master_file:
        if master_file.name.lower().endswith(".csv"):master=pd.read_csv(master_file)
        else:
            mf=load_excel(master_file.getvalue())
            master=mf.get("Material_Master") or next(iter(mf.values()))
    if master is None:
        st.warning("Upload your approved Material Master to activate controlled material mapping.")
        template=sv[[c for c in [desc,code] if c]].drop_duplicates()
        if desc:template=template.rename(columns={desc:"Warehouse Material Name"})
        if code:template=template.rename(columns={code:"Material Code"})
        template["AG Catalogue Name"]="";template["Status"]="ACTIVE"
        st.download_button("Download Material Master Template",template.to_csv(index=False).encode(),"ATOM_GRID_Material_Master_Template.csv","text/csv")
    else:
        master.columns=[str(c).strip() for c in master.columns]
        ml={str(c).lower():c for c in master.columns}
        mn=ml.get("warehouse material name") or ml.get("material") or ml.get("description")
        mc=ml.get("material code") or ml.get("m.code") or ml.get("code")
        if not mn:st.error("Material Master needs a 'Warehouse Material Name' column.")
        else:
            known=set(master[mn].dropna().map(norm))
            codes=set(master[mc].dropna().map(norm)) if mc else set()
            check=sv.copy()
            check["Mapping Status"]=np.where(check["_Material"].map(norm).isin(known)|check["_Code"].isin(codes),"MAPPED","NEW / NOT MAPPED")
            new=check[check["Mapping Status"]=="NEW / NOT MAPPED"]
            st.metric("New / unmapped stock lines",len(new))
            show=[c for c in [gr,code,desc,batch,closing,age_col] if c]+["Mapping Status"]
            st.dataframe(new[show],use_container_width=True,hide_index=True)
            st.caption("Unknown materials are never silently auto-mapped.")

with tabs[7]:
    if pending is not None:st.dataframe(pending,use_container_width=True,hide_index=True)
    else:st.info("Pending Order Report not found.")
with tabs[8]:
    if vehicle is not None:st.dataframe(vehicle,use_container_width=True,hide_index=True)
    else:st.info("Vehicle Indent not found.")
with tabs[9]:
    if cancelled is not None:st.dataframe(cancelled,use_container_width=True,hide_index=True)
    else:st.info("Cancelled Invoice Report not found.")

st.divider()
st.caption("ATOM GRID Warehouse MIS • Google sign-in • @atomgrid.in only • Upload-driven dashboard")
