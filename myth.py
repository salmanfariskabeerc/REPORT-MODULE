import streamlit as st
import pandas as pd, re
from io import BytesIO
from collections import Counter
from datetime import datetime
import zipfile, xml.etree.ElementTree as ET
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Al Madina Reports", page_icon="📊",
                   layout="centered", initial_sidebar_state="collapsed")
PASSWORD = "123123123"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
#MainMenu,footer,header,[data-testid="stSidebar"]{display:none!important;}
.main .block-container{padding:2rem 1.5rem 4rem;max-width:820px;margin:0 auto;}
/* Login */
.login-wrap{min-height:80vh;display:flex;align-items:center;justify-content:center;}
.lcard{background:white;border-radius:24px;padding:48px 44px;max-width:400px;width:100%;
  box-shadow:0 24px 80px rgba(0,0,0,.12);border:1px solid #f1f5f9;}
.llogo{width:60px;height:60px;border-radius:16px;background:linear-gradient(135deg,#f97316,#dc2626);
  display:flex;align-items:center;justify-content:center;font-size:26px;margin:0 auto 18px;}
.ltitle{font-size:22px;font-weight:800;color:#0f172a;text-align:center;margin-bottom:4px;letter-spacing:-.02em;}
.lsub{font-size:12px;color:#94a3b8;text-align:center;margin-bottom:28px;}
/* Brand */
.brand{background:linear-gradient(135deg,#0f172a,#1e293b);border-radius:16px;
  padding:20px 24px;margin-bottom:20px;display:flex;align-items:center;gap:14px;}
.b-icon{width:46px;height:46px;border-radius:12px;background:linear-gradient(135deg,#f97316,#dc2626);
  display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0;}
.b-title{font-size:17px;font-weight:700;color:white;margin:0;}
.b-sub{font-size:11px;color:#475569;margin-top:2px;}
/* Upload box */
.ubox{background:white;border:2px dashed #cbd5e1;border-radius:16px;padding:40px 28px;
  text-align:center;margin-bottom:18px;transition:border-color .2s;}
.u-icon{font-size:38px;margin-bottom:12px;}
.u-title{font-size:16px;font-weight:700;color:#0f172a;margin-bottom:5px;}
.u-sub{font-size:12px;color:#94a3b8;}
/* Success */
.sbar{background:#f0fdf4;border:1px solid #86efac;border-radius:12px;padding:13px 17px;
  margin-bottom:16px;display:flex;align-items:center;gap:10px;}
/* Chips */
.chips{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:18px;}
.chip{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;
  padding:5px 11px;font-size:11px;color:#475569;}
.chip b{color:#0f172a;}
/* Cards */
.rcard{background:white;border:1px solid #e2e8f0;border-radius:16px;padding:22px;margin-bottom:12px;}
.rcard-title{font-size:14px;font-weight:700;color:#0f172a;margin-bottom:3px;letter-spacing:-.01em;}
.rcard-sub{font-size:11px;color:#94a3b8;margin-bottom:16px;line-height:1.5;}
.rcard-feat{list-style:none;padding:0;margin:0 0 16px;}
.rcard-feat li{font-size:11px;color:#475569;padding:3px 0;display:flex;align-items:center;gap:6px;}
.rcard-feat li::before{content:"✓";color:#10b981;font-weight:800;font-size:12px;}
/* Outlet strip */
.outlet-strip{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;
  padding:10px 14px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center;}
.os-name{font-size:12px;font-weight:700;color:#0f172a;}
.os-meta{font-size:11px;color:#64748b;}
.os-badge{font-size:10px;font-weight:700;padding:3px 9px;border-radius:10px;}
/* Buttons */
.stButton>button{background:linear-gradient(135deg,#f97316,#dc2626)!important;color:white!important;
  border:none!important;border-radius:10px!important;font-weight:700!important;
  font-size:13px!important;padding:12px 20px!important;width:100%!important;letter-spacing:.01em!important;}
.stDownloadButton>button{background:linear-gradient(135deg,#0f172a,#1e293b)!important;color:white!important;
  border:none!important;border-radius:10px!important;font-weight:700!important;
  font-size:13px!important;padding:12px 20px!important;width:100%!important;}
div[data-testid="stTextInput"] input{border-radius:10px!important;border:1.5px solid #e2e8f0!important;
  padding:12px 14px!important;font-size:14px!important;}
div[data-testid="stTextInput"] input:focus{border-color:#f97316!important;
  box-shadow:0 0 0 3px rgba(249,115,22,.15)!important;}
.stMultiSelect [data-baseweb="select"]{border-radius:10px!important;}
.stExpander{border:1px solid #e2e8f0!important;border-radius:12px!important;}
</style>""", unsafe_allow_html=True)

# ─── Constants ─────────────────────────────────────────────────
NUMERIC_COLS=["Subtotal","Packaging charges","Minimum order value fee","Vendor Refunds",
    "Tax Charge","Online Payment Fee","Discount Funded by you","Voucher Funded by you",
    "Commission","Operational Charges","Ads Fee","Marketing Fees","Avoidable cancellation fee",
    "Estimated earnings","Cash amount already collected by you","Amount owed back to Talabat",
    "Payout Amount","Total Discount","Total Voucher","Tax Amount"]
DATE_COLS=["Order received at","Accepted at","Ready to pick up at","Rider near pickup at",
    "In delivery at","Delivered at","Cancelled at"]
CLRS={"Liwan":"#f97316","Dubai Investment Park":"#3b82f6","Oud Metha":"#8b5cf6",
      "Naif":"#ef4444","Al Muteena":"#10b981","Al Hamriya":"#f59e0b"}

# ─── Loader ────────────────────────────────────────────────────
def _parse_stdlib(file_bytes):
    ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
        names=[i.filename for i in z.infolist()]
        sheet=z.read("xl/worksheets/sheet1.xml")
        shared=[]
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter("{"+ns+"}si"):
                shared.append("".join(t.text or "" for t in si.iter("{"+ns+"}t")))
    def ci(ref):
        col="".join(filter(str.isalpha,ref)); idx=0
        for ch in col: idx=idx*26+(ord(ch.upper())-64)
        return idx-1
    rows_out=[]; max_col=0
    for row_el in ET.fromstring(sheet).iter("{"+ns+"}row"):
        cells={}
        for cell in row_el.iter("{"+ns+"}c"):
            ref=cell.get("r","A1"); col_i=ci(ref); t=cell.get("t","n")
            v_el=cell.find("{"+ns+"}v"); is_el=cell.find("{"+ns+"}is")
            if is_el is not None:
                t_el=is_el.find("{"+ns+"}t"); val=t_el.text if t_el is not None else ""
            elif v_el is not None:
                rv=v_el.text or ""
                if t=="s": val=shared[int(rv)] if shared else rv
                else:
                    try: val=float(rv) if "." in rv else int(rv)
                    except: val=rv
            else: val=None
            cells[col_i]=val; max_col=max(max_col,col_i)
        rows_out.append([cells.get(i) for i in range(max_col+1)])
    return pd.DataFrame(rows_out) if rows_out else pd.DataFrame()

def get_area(name):
    n=str(name).lower()
    if "liwan" in n: return "Liwan"
    if "dip" in n or "investment park" in n: return "Dubai Investment Park"
    if "oud metha" in n: return "Oud Metha"
    if "naif" in n: return "Naif"
    if "muteena" in n: return "Al Muteena"
    if "hamriya" in n: return "Al Hamriya"
    if "Al Madina Hypermarket, Dubai Investments Park 1" in n: return "fida almadina"
    return "Other"

@st.cache_data(show_spinner=False)
def load_data(file_bytes):
    df_raw=_parse_stdlib(file_bytes)
    hr=None
    for i in range(min(5,len(df_raw))):
        if "Order status" in [str(v) for v in df_raw.iloc[i].values]:
            hr=i; break
    if hr is None: hr=0
    df=df_raw.iloc[hr+1:].copy()
    df.columns=[str(c) for c in df_raw.iloc[hr].tolist()]
    df=df.dropna(how="all").reset_index(drop=True)
    for col in NUMERIC_COLS:
        if col in df.columns: df[col]=pd.to_numeric(df[col],errors="coerce").fillna(0)
    for col in DATE_COLS:
        if col in df.columns: df[col]=pd.to_datetime(df[col],errors="coerce")
    if "Restaurant name" not in df.columns:
        raise RuntimeError("Not a valid Talabat order report.")
    df["_area"]=df["Restaurant name"].apply(get_area)
    df["_date"]=df["Order received at"].dt.date
    df["_hour"]=df["Order received at"].dt.hour
    df["_dow"]=df["Order received at"].dt.day_name()
    return df

def sdiv(a,b,pct=False):
    if b==0: return 0
    return (a/b*100) if pct else a/b

def calc_quick(df_sub):
    d=df_sub[df_sub["Order status"]=="Delivered"]
    c=df_sub[df_sub["Order status"]=="Cancelled"]
    total=len(df_sub); gmv=d["Subtotal"].sum()
    diff=(d["Delivered at"]-d["Order received at"]).dt.total_seconds()/60; diff=diff[diff>0]
    return {"total":total,"delivered":len(d),"cancelled":len(c),
            "can_rate":sdiv(len(c),total,pct=True),
            "del_rate":sdiv(len(d),total,pct=True),
            "gmv":gmv,"payout":d["Payout Amount"].sum(),"lost_gmv":c["Subtotal"].sum(),
            "del_time":round(diff.mean(),1) if len(diff)>0 else None}

def fp(v): return "{:.1f}%".format(v)
def fm(v): return "{:.0f} min".format(v) if v else "N/A"
def fa(v): return "AED {:,.0f}".format(v)


# ─── Inline the full report builder ────────────────────────────

import pandas as pd, re
from collections import Counter
from datetime import datetime
from io import BytesIO
import zipfile, xml.etree.ElementTree as ET

REPORT_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Inter',Arial,sans-serif;color:#1e293b;background:#dde3ea;font-size:12px;line-height:1.55;}

/* ── COVER ── */
.cover{background:linear-gradient(160deg,#0f172a 0%,#1e293b 40%,#0c1a35 100%);
  display:flex;flex-direction:column;justify-content:center;
  align-items:center;text-align:center;padding:60px 40px;page-break-after:always;
  position:relative;overflow:hidden;margin-bottom:12px;}
.cover::before{content:'';position:absolute;top:-40%;right:-20%;width:600px;height:600px;
  border-radius:50%;background:radial-gradient(circle,rgba(249,115,22,.12),transparent 70%);}
.cover::after{content:'';position:absolute;bottom:-30%;left:-10%;width:500px;height:500px;
  border-radius:50%;background:radial-gradient(circle,rgba(59,130,246,.1),transparent 70%);}
.cover-badge{background:rgba(249,115,22,.15);border:1px solid rgba(249,115,22,.4);
  color:#fb923c;padding:6px 20px;border-radius:20px;font-size:10px;font-weight:700;
  letter-spacing:.12em;text-transform:uppercase;margin-bottom:24px;display:inline-block;position:relative;z-index:1;}
.cover-logo{width:80px;height:80px;border-radius:20px;
  background:linear-gradient(135deg,#f97316,#dc2626);
  display:flex;align-items:center;justify-content:center;
  font-size:36px;margin:0 auto 24px;position:relative;z-index:1;
  box-shadow:0 20px 60px rgba(249,115,22,.4);}
.cover h1{font-size:40px;font-weight:800;color:white;margin-bottom:8px;
  letter-spacing:-.03em;position:relative;z-index:1;}
.cover h2{font-size:18px;font-weight:400;color:#64748b;margin-bottom:40px;position:relative;z-index:1;}
.cover-stats{display:flex;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);
  border-radius:16px;overflow:hidden;margin-bottom:36px;position:relative;z-index:1;backdrop-filter:blur(10px);}
.cs{padding:20px 26px;text-align:center;border-right:1px solid rgba(255,255,255,.07);}
.cs:last-child{border-right:none;}
.cs-l{font-size:9px;color:#475569;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px;}
.cs-v{font-size:18px;font-weight:700;color:white;}
.cover-pills{display:flex;gap:8px;flex-wrap:wrap;justify-content:center;position:relative;z-index:1;}
.cpill{border-radius:8px;padding:7px 16px;font-size:11px;font-weight:600;border:1px solid rgba(255,255,255,.1);}

/* ── PAGES ── */
.page{background:white;padding:36px 40px 56px;
  page-break-after:always;position:relative;
  margin-bottom:12px;box-shadow:0 2px 12px rgba(0,0,0,.08);}
.page:last-child{page-break-after:auto;}
.ph{display:flex;justify-content:space-between;align-items:flex-start;
  margin-bottom:22px;padding-bottom:14px;border-bottom:2px solid #f8fafc;}
.ph-title{font-size:19px;font-weight:800;color:#0f172a;margin-bottom:2px;letter-spacing:-.02em;}
.ph-sub{font-size:10px;color:#94a3b8;font-weight:400;}
.ph-badge{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;
  padding:5px 14px;font-size:10px;font-weight:600;color:#64748b;}

/* ── SECTION TITLES ── */
.st{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;
  color:#94a3b8;margin:20px 0 10px;display:flex;align-items:center;gap:8px;}
.st::after{content:'';flex:1;height:1px;background:#f1f5f9;}
.st.first{margin-top:0;}

/* ── KPI CARDS ── */
.kg{display:grid;gap:10px;margin-bottom:4px;}
.kg3{grid-template-columns:repeat(3,1fr);}
.kg4{grid-template-columns:repeat(4,1fr);}
.kg5{grid-template-columns:repeat(5,1fr);}
.kg6{grid-template-columns:repeat(6,1fr);}
.kc{background:#f8fafc;border-radius:10px;padding:12px 14px;border-left:3px solid #e2e8f0;position:relative;}
.kc.or{border-left-color:#f97316;} .kc.gr{border-left-color:#10b981;}
.kc.re{border-left-color:#ef4444;} .kc.bl{border-left-color:#3b82f6;}
.kc.pu{border-left-color:#8b5cf6;} .kc.am{border-left-color:#f59e0b;}
.kc.cy{border-left-color:#06b6d4;} .kc.pi{border-left-color:#ec4899;}
.k-l{font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#94a3b8;margin-bottom:5px;}
.k-v{font-size:20px;font-weight:700;color:#0f172a;line-height:1.1;}
.k-s{font-size:9px;color:#64748b;margin-top:3px;}
.k-trend{position:absolute;top:10px;right:12px;font-size:9px;font-weight:600;padding:2px 6px;border-radius:6px;}
.k-trend.up{background:#d1fae5;color:#065f46;}
.k-trend.dn{background:#fee2e2;color:#991b1b;}

/* ── TABLES ── */
table{width:100%;border-collapse:collapse;font-size:10.5px;margin-bottom:4px;}
thead th{padding:8px 10px;font-size:8.5px;font-weight:700;text-transform:uppercase;
  letter-spacing:.07em;color:white;text-align:left;white-space:nowrap;}
tbody td{padding:8px 10px;border-bottom:1px solid #f8fafc;color:#374151;vertical-align:middle;}
tbody tr:last-child td{border-bottom:none;}
tbody tr:nth-child(even) td{background:#fafbfc;}
td.r,th.r{text-align:right;}td.c,th.c{text-align:center;}
td.b{font-weight:700;color:#0f172a;}td.gr{color:#065f46;font-weight:600;}
td.re{color:#991b1b;font-weight:600;}td.am{color:#92400e;font-weight:600;}
.rank-medal{font-size:14px;}

/* ── BADGES ── */
.bdg{padding:2px 8px;border-radius:10px;font-size:9.5px;font-weight:600;white-space:nowrap;display:inline-block;}
.bdg-gr{background:#d1fae5;color:#065f46;}
.bdg-re{background:#fee2e2;color:#991b1b;}
.bdg-am{background:#fef3c7;color:#92400e;}
.bdg-bl{background:#dbeafe;color:#1e40af;}
.bdg-pu{background:#ede9fe;color:#5b21b6;}
.bdg-cy{background:#cffafe;color:#0e7490;}

/* ── BARS ── */
.br{display:flex;align-items:center;gap:8px;margin-bottom:7px;}
.br-l{font-size:10px;color:#475569;width:155px;flex-shrink:0;text-align:right;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.br-t{flex:1;height:9px;background:#f1f5f9;border-radius:5px;overflow:hidden;}
.br-f{height:100%;border-radius:5px;}
.br-v{font-size:10px;font-weight:600;color:#0f172a;min-width:54px;}

/* ── MINI BARS (inline in table) ── */
.mbar{display:inline-flex;align-items:center;gap:5px;width:100%;}
.mbar-t{flex:1;height:6px;background:#f1f5f9;border-radius:3px;overflow:hidden;}
.mbar-f{height:100%;border-radius:3px;}

/* ── OUTLET HEADER CARD ── */
.ohdr{border-radius:14px;padding:18px 22px;margin-bottom:16px;color:white;position:relative;overflow:hidden;}
.ohdr::before{content:'';position:absolute;top:-30%;right:-10%;width:200px;height:200px;
  border-radius:50%;background:rgba(255,255,255,.07);}
.ohdr h2{font-size:20px;font-weight:800;margin-bottom:2px;position:relative;}
.ohdr .store{font-size:10px;opacity:.6;margin-bottom:14px;position:relative;}
.ohdr .metrics{display:flex;gap:22px;position:relative;}
.mc-l{font-size:8px;text-transform:uppercase;letter-spacing:.07em;opacity:.55;margin-bottom:3px;}
.mc-v{font-size:17px;font-weight:700;}

/* ── ALERTS ── */
.al{border-radius:8px;padding:9px 12px;margin-bottom:8px;font-size:10.5px;
  display:flex;gap:9px;align-items:flex-start;line-height:1.4;}
.al-icon{font-size:13px;flex-shrink:0;margin-top:1px;}
.al-r{background:#fff1f2;border-left:3px solid #ef4444;color:#7f1d1d;}
.al-a{background:#fffbeb;border-left:3px solid #f59e0b;color:#78350f;}
.al-g{background:#f0fdf4;border-left:3px solid #10b981;color:#14532d;}
.al-b{background:#eff6ff;border-left:3px solid #3b82f6;color:#1e3a8a;}
.al-pu{background:#faf5ff;border-left:3px solid #8b5cf6;color:#4c1d95;}

/* ── FLOW DIAGRAM ── */
.flow{display:flex;align-items:stretch;gap:0;margin:10px 0 16px;}
.fstep{flex:1;position:relative;}
.fstep:not(:last-child)::after{content:'▶';position:absolute;right:-8px;top:50%;
  transform:translateY(-50%);color:#cbd5e1;font-size:10px;z-index:2;}
.fbox{background:#f8fafc;border-radius:10px;padding:10px 8px;text-align:center;
  border:1px solid #f1f5f9;height:100%;display:flex;flex-direction:column;justify-content:center;}
.fbox.highlight{border-color:#f97316;background:#fff7ed;}
.f-l{font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px;}
.f-v{font-size:17px;font-weight:700;color:#0f172a;}
.f-u{font-size:8px;color:#94a3b8;}

/* ── HEATMAP ── */
.heat-wrap{margin:6px 0 14px;}
.heat-row{display:flex;align-items:center;gap:4px;margin-bottom:3px;}
.heat-label{font-size:9px;color:#94a3b8;width:28px;text-align:right;flex-shrink:0;}
.heat-cell{height:16px;border-radius:3px;flex:1;position:relative;}

/* ── SCORE RING ── */
.score-ring{display:inline-flex;align-items:center;justify-content:center;
  width:52px;height:52px;border-radius:50%;font-size:14px;font-weight:800;
  border:3px solid;}
.ring-gr{color:#10b981;border-color:#10b981;background:#f0fdf4;}
.ring-am{color:#f59e0b;border-color:#f59e0b;background:#fffbeb;}
.ring-re{color:#ef4444;border-color:#ef4444;background:#fff1f2;}

/* ── INSIGHT BOXES ── */
.insight-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:10px 0;}
.insight{background:#f8fafc;border-radius:10px;padding:12px 14px;border-top:3px solid #e2e8f0;}
.insight.or{border-top-color:#f97316;} .insight.re{border-top-color:#ef4444;}
.insight.gr{border-top-color:#10b981;} .insight.bl{border-top-color:#3b82f6;}
.ins-t{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;
  color:#94a3b8;margin-bottom:5px;}
.ins-v{font-size:13px;font-weight:700;color:#0f172a;margin-bottom:3px;}
.ins-s{font-size:10px;color:#64748b;}

/* ── TWO/THREE COL ── */
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:18px;}
.three-col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;}
.col-6040{display:grid;grid-template-columns:60fr 40fr;gap:18px;}
.col-4060{display:grid;grid-template-columns:40fr 60fr;gap:18px;}

/* ── PAGE FOOTER ── */
.pf{position:absolute;bottom:16px;left:36px;right:36px;display:flex;justify-content:space-between;
  font-size:9px;color:#cbd5e1;border-top:1px solid #f8fafc;padding-top:8px;}

@media print{body{background:white;}.page,.cover{box-shadow:none;margin-bottom:0;}}
@media screen{.page,.cover{max-width:960px;margin-left:auto;margin-right:auto;}}
</style>
"""

def build_full_report(df, areas, title="Full Performance Report"):
    delivered = df[df["Order status"]=="Delivered"]
    cancelled = df[df["Order status"]=="Cancelled"]
    date_min = df["Order received at"].min()
    date_max = df["Order received at"].max()
    ds = ""
    if pd.notna(date_min) and pd.notna(date_max):
        ds = date_min.strftime("%d %b %Y")+" – "+date_max.strftime("%d %b %Y")

    CLRS = {"Liwan":"#f97316","Dubai Investment Park":"#3b82f6","Oud Metha":"#8b5cf6",
            "Naif":"#ef4444","Al Muteena":"#10b981","Al Hamriya":"#f59e0b"}

    def sdiv(a,b,pct=False):
        if b==0: return 0
        return (a/b*100) if pct else a/b

    def fp(v): return "{:.1f}%".format(v)
    def fm(v): return "{:.0f}m".format(v) if v else "N/A"
    def fa(v, dec=0): return "AED {:,.{}f}".format(v,dec)

    def calc_om(df_sub):
        d=df_sub[df_sub["Order status"]=="Delivered"]
        c=df_sub[df_sub["Order status"]=="Cancelled"]
        total=len(df_sub); gmv=d["Subtotal"].sum()
        diff=(d["Delivered at"]-d["Order received at"]).dt.total_seconds()/60; diff=diff[diff>0]
        prep=(d["Ready to pick up at"]-d["Accepted at"]).dt.total_seconds()/60; prep=prep[prep>0]
        lm=(d["Delivered at"]-d["In delivery at"]).dt.total_seconds()/60; lm=lm[lm>0]
        rider=(d["Rider near pickup at"]-d["Ready to pick up at"]).dt.total_seconds()/60; rider=rider[rider>0]
        return {"total":total,"delivered":len(d),"cancelled":len(c),
            "can_rate":sdiv(len(c),total,pct=True),"del_rate":sdiv(len(d),total,pct=True),
            "gmv":gmv,"payout":d["Payout Amount"].sum(),"commission":d["Commission"].sum(),
            "op_charges":d["Operational Charges"].sum(),"online_fee":d["Online Payment Fee"].sum(),
            "avg_order":sdiv(gmv,len(d)),
            "del_time":round(diff.mean(),1) if len(diff)>0 else None,
            "del_p25":round(diff.quantile(.25),1) if len(diff)>0 else None,
            "del_p75":round(diff.quantile(.75),1) if len(diff)>0 else None,
            "del_p90":round(diff.quantile(.9),1) if len(diff)>0 else None,
            "del_max":round(diff.max(),1) if len(diff)>0 else None,
            "prep_time":round(prep.mean(),1) if len(prep)>0 else None,
            "last_mile":round(lm.mean(),1) if len(lm)>0 else None,
            "rider_wait":round(rider.mean(),1) if len(rider)>0 else None,
            "pro_pct":sdiv(len(d[d["Is Pro Order"]=="Y"]),len(d),pct=True),
            "online_pct":sdiv(len(d[d["Payment type"]=="Online"]),len(d),pct=True),
            "cash_pct":sdiv(len(d[d["Payment type"]=="Cash"]),len(d),pct=True),
            "complaints":len(d[d["Has Complaint?"]=="Y"]),
            "complaint_rate":sdiv(len(d[d["Has Complaint?"]=="Y"]),len(d),pct=True),
            "lost_gmv":c["Subtotal"].sum()}

    def extract_items(series, top=15):
        ctr=Counter()
        for row in series.dropna():
            for item in str(row).split(","):
                item=item.strip(); parts=item.split(" ",1)
                if len(parts)==2 and parts[0].isdigit(): qty,name=int(parts[0]),parts[1].strip()
                else: qty,name=1,item
                skip={"Sheet","Sheets","Pieces","square feet","Square Feet","Sachets","Sticks","Pellets","Diapers"}
                if name and len(name)>6 and not re.match(r'^[\d\s\.,gGmMkKlLx%]+$',name) and name not in skip:
                    ctr[name]+=qty
        return ctr.most_common(top)

    # Compute outlet data
    outlet_data=[(a,df[df["_area"]==a]["Restaurant name"].iloc[0],calc_om(df[df["_area"]==a]))
                 for a in sorted(areas) if len(df[df["_area"]==a])>0]
    outlet_data.sort(key=lambda x:x[2]["gmv"],reverse=True)
    m=calc_om(df[df["_area"].isin(areas)])
    sub_can=cancelled[cancelled["_area"].isin(areas)]
    del_in=delivered[delivered["_area"].isin(areas)]

    # ── HTML helpers ───────────────────────────────────────────
    def kc(label,val,sub="",cls="or",trend=None):
        t=""
        if trend: t='<span class="k-trend {}">{}</span>'.format("up" if "↑" in trend or "+" in trend else "dn",trend)
        return '<div class="kc {}">{}<div class="k-l">{}</div><div class="k-v">{}</div>{}</div>'.format(
            cls,t,label,val,'<div class="k-s">{}</div>'.format(sub) if sub else "")

    def bdg(text,cls="bl"):
        return '<span class="bdg bdg-{}">{}</span>'.format(cls,text)

    def can_cls(r):
        return "re" if r>20 else ("am" if r>10 else "gr")

    def del_cls(v):
        if v is None: return "am"
        return "gr" if v<=35 else ("re" if v>45 else "am")

    def tbl(headers,rows,color="#0f172a",r_cols=None,c_cols=None):
        r_cols=r_cols or set(); c_cols=c_cols or set()
        ths="".join('<th class="{}" style="background:{}">{}</th>'.format(
            "r" if i in r_cols else("c" if i in c_cols else""),color,h) for i,h in enumerate(headers))
        body=""
        for ri,row in enumerate(rows):
            tds="".join('<td class="{}">{}</td>'.format(
                "r" if ci in r_cols else("c" if ci in c_cols else""),cell) for ci,cell in enumerate(row))
            body+="<tr>{}</tr>".format(tds)
        return "<table><thead><tr>{}</tr></thead><tbody>{}</tbody></table>".format(ths,body)

    def bar(items,color="#f97316",show_val=True):
        if not items: return ""
        mx=max(v for _,v in items) or 1
        out=""
        for name,val in items:
            pct=min(val/mx*100,100)
            out+=('<div class="br"><div class="br-l">{}</div>'
                  '<div class="br-t"><div class="br-f" style="width:{:.1f}%;background:{};"></div></div>'
                  '<div class="br-v">{}</div></div>').format(name[:26],pct,color,"{:,}".format(val) if show_val else fp(val))
        return out

    def inline_bar(val,max_val,color="#f97316"):
        pct=min(val/max_val*100,100) if max_val>0 else 0
        return '<div class="mbar"><div class="mbar-t"><div class="mbar-f" style="width:{:.0f}%;background:{};"></div></div></div>'.format(pct,color)

    def sec(title,first=False):
        return '<div class="st{}">{}</div>'.format(" first" if first else "",title)

    def ph(title,sub):
        return '<div class="ph"><div><div class="ph-title">{}</div><div class="ph-sub">{}</div></div><div class="ph-badge">{}</div></div>'.format(title,sub,ds)

    def pf(note=""):
        return '<div class="pf"><span>Al Madina Hypermarket &nbsp;·&nbsp; Confidential</span><span>{} &nbsp;|&nbsp; Generated {}</span></div>'.format(note,datetime.now().strftime("%d %b %Y %H:%M"))

    def al(msg,cls="b",icon="ℹ️"):
        return '<div class="al al-{}"><span class="al-icon">{}</span><span>{}</span></div>'.format(cls,icon,msg)

    # ═══════════════════════════════════════════════════════════
    H = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>"
    H += "<title>Al Madina {} – {}</title>".format(title,ds)
    H += REPORT_CSS + "</head><body>"

    # ── COVER ──────────────────────────────────────────────────
    H+='<div class="cover">'
    H+='<div class="cover-badge">Al Madina Hypermarket · Confidential</div>'
    H+='<div class="cover-logo">🏪</div>'
    H+='<h1>{}</h1>'.format(title)
    H+='<h2>Data-Driven Performance Intelligence Report</h2>'
    H+='<div class="cover-stats">'
    for l,v in [("Report Period",ds),("Outlets Analysed",str(len(outlet_data))),
                ("Total Orders",str(m["total"])),("Gross Revenue",fa(m["gmv"])),
                ("Payout",fa(m["payout"])),("Cancel Rate",fp(m["can_rate"]))]:
        H+='<div class="cs"><div class="cs-l">{}</div><div class="cs-v">{}</div></div>'.format(l,v)
    H+='</div>'
    H+='<div class="cover-pills">'
    for area,_,o in outlet_data:
        c=CLRS.get(area,"#64748b")
        H+='<div class="cpill" style="background:{}18;color:{};border-color:{}35;">{} &nbsp; {} &nbsp; {:.1f}% cancel</div>'.format(c,c,c,area,fa(o["gmv"]),o["can_rate"])
    H+='</div></div>'

    # ═══════════════════════════════════════════════════════════
    # PAGE 1 — EXECUTIVE DASHBOARD
    # ═══════════════════════════════════════════════════════════
    H+='<div class="page">'
    H+=ph("Executive Dashboard","Group Overview — {} Outlets — {}".format(len(outlet_data),ds))
    H+=sec("REVENUE & ORDER PERFORMANCE",True)
    H+='<div class="kg kg5">'
    H+=kc("Total Orders",m["total"],"{} del / {} can".format(m["delivered"],m["cancelled"]),"or")
    H+=kc("Gross Revenue",fa(m["gmv"]),"Total order value","bl")
    H+=kc("Total Payout",fa(m["payout"]),"{:.1f}% of GMV".format(sdiv(m["payout"],m["gmv"],pct=True)),"gr")
    H+=kc("Cancellation Rate",fp(m["can_rate"]),"AED {:,.0f} lost".format(m["lost_gmv"]),"re")
    H+=kc("Avg Order Value",fa(m["avg_order"],1),"per delivered order","pu")
    H+='</div>'
    H+='<div class="kg kg5">'
    H+=kc("Commission",fa(m["commission"]),"{:.1f}% of GMV".format(sdiv(m["commission"],m["gmv"],pct=True)),"am")
    H+=kc("Op Charges",fa(m["op_charges"]),"{:.1f}% of GMV".format(sdiv(m["op_charges"],m["gmv"],pct=True)),"am")
    H+=kc("Avg Delivery",fm(m["del_time"]),"P50 across all outlets","cy")
    H+=kc("Pro Orders %",fp(m["pro_pct"]),"of delivered orders","pu")
    H+=kc("Online Payment",fp(m["online_pct"]),"vs {:.0f}% cash".format(m["cash_pct"]),"gr")
    H+='</div>'

    H+=sec("OUTLET LEAGUE TABLE — RANKED BY GMV")
    medals=["🥇","🥈","🥉"]
    max_gmv=max(o["gmv"] for _,_,o in outlet_data) or 1
    lr=[]
    for rank,(area,fname,o) in enumerate(outlet_data,1):
        medal=medals[rank-1] if rank<=3 else "#{:02d}".format(rank)
        gmv_bar=inline_bar(o["gmv"],max_gmv,CLRS.get(area,"#f97316"))
        can_color=can_cls(o["can_rate"]); del_color=del_cls(o["del_time"])
        lr.append([medal,area,str(o["total"]),str(o["delivered"]),
            bdg(fp(o["can_rate"]),can_color),
            gmv_bar+"<b> "+fa(o["gmv"])+"</b>",fa(o["payout"]),
            fa(o["avg_order"],1),fm(o["del_time"]),
            bdg(fp(o["complaint_rate"]),"re" if o["complaint_rate"]>2 else "gr")])
    H+=tbl(["","Outlet","Orders","Delivered","Cancel%","GMV (AED)","Payout","Avg Order","Del Time","Complaints"],
           lr,"#0f172a",r_cols={2,3,6,7},c_cols={0,4,8,9})

    H+=sec("DAILY ORDER TREND — ALL OUTLETS")
    dg=df[df["_area"].isin(areas)].groupby("_date").agg(
        total=("Order ID","count"),
        delivered=("Order status",lambda x:(x=="Delivered").sum()),
        cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
        gmv=("Subtotal","sum")).reset_index()
    dp=del_in.groupby(del_in["_date"])["Payout Amount"].sum()
    dg["payout"]=dp.reindex(dg["_date"]).fillna(0).values
    dg["can_pct"]=dg.apply(lambda r:sdiv(r["cancelled"],r["total"],pct=True),axis=1)
    dg["gmv_bar"]=dg["gmv"].apply(lambda v:inline_bar(v,dg["gmv"].max(),"#3b82f6"))
    dr=[]
    for _,r in dg.iterrows():
        cb=bdg(fp(r["can_pct"]),can_cls(r["can_pct"]))
        dr.append([str(r["_date"]),int(r["total"]),int(r["delivered"]),int(r["cancelled"]),cb,
                   r["gmv_bar"]+"<b> "+fa(r["gmv"])+"</b>",fa(r["payout"])])
    H+=tbl(["Date","Total","Delivered","Cancelled","Cancel Rate","GMV (AED)","Payout (AED)"],
           dr,"#1e293b",r_cols={1,2,3,6},c_cols={4})
    H+=pf("Page 1 – Executive Dashboard")
    H+='</div>'

    # ═══════════════════════════════════════════════════════════
    # PAGE 2 — SALES & FINANCIAL DEEP DIVE
    # ═══════════════════════════════════════════════════════════
    H+='<div class="page">'
    H+=ph("Sales & Financial Analysis","Revenue Breakdown, Costs, Payment Mix & Profitability")

    H+='<div class="col-6040">'
    H+='<div>'
    H+=sec("COST STRUCTURE & PAYOUT WATERFALL",True)
    total_cost=m["commission"]+m["op_charges"]+m["online_fee"]
    net_margin=sdiv(m["payout"],m["gmv"],pct=True)
    fr=[]
    for l,v in [("Subtotal (GMV)",m["gmv"]),("Commission",m["commission"]),
                ("Operational Charges",m["op_charges"]),("Online Payment Fee",m["online_fee"])]:
        pct=sdiv(v,m["gmv"],pct=True)
        bar_w=int(pct*0.8)
        fr.append(["<b>{}</b>".format(l) if "GMV" in l or "Payout" in l else l,
                   "<b>{}</b>".format(fa(v)) if "GMV" in l else fa(v),
                   "<b>{}</b>".format(fp(pct)) if "GMV" in l else fp(pct),
                   fa(sdiv(v,m["delivered"],1),1)])
    fr.append(["────────","────","────","────"])
    fr.append(["<b>Total Deductions</b>","<b>{}</b>".format(fa(total_cost)),"<b>{}</b>".format(fp(sdiv(total_cost,m["gmv"],pct=True))),fa(sdiv(total_cost,m["delivered"],1),1)])
    fr.append(["<b>Net Payout</b>","<b>{}</b>".format(fa(m["payout"])),"<b>{:.1f}% margin</b>".format(net_margin),fa(sdiv(m["payout"],m["delivered"],1),1)])
    H+=tbl(["Item","Total (AED)","% of GMV","Per Order"],fr,"#1e293b",r_cols={1,2,3})

    H+=sec("OUTLET PROFITABILITY COMPARISON")
    pr=[]
    for area,_,o in outlet_data:
        margin=sdiv(o["payout"],o["gmv"],pct=True)
        cost_ratio=sdiv(o["commission"]+o["op_charges"],o["gmv"],pct=True)
        pr.append([area,fa(o["gmv"]),fa(o["payout"]),
                   bdg("{:.0f}%".format(margin),"gr" if margin>70 else("am" if margin>60 else"re")),
                   bdg("{:.0f}%".format(cost_ratio),"re" if cost_ratio>20 else("am" if cost_ratio>15 else"gr"))])
    H+=tbl(["Outlet","GMV","Payout","Margin","Cost Ratio"],pr,"#1e293b",r_cols={1,2},c_cols={3,4})
    H+='</div>'
    H+='<div>'
    H+=sec("PAYMENT MIX",True)
    for ptype,color,bdg_c in [("Online","#3b82f6","bl"),("Cash","#10b981","gr")]:
        sdf=del_in[del_in["Payment type"]==ptype]
        H+=('<div class="kc {}" style="margin-bottom:8px;">'
            '<div class="k-l">{} Payments</div>'
            '<div class="k-v">{}</div>'
            '<div class="k-s">{} orders · {} GMV · {:.0f}% share</div></div>').format(
            bdg_c.replace("bl","bl").replace("gr","gr"),ptype,
            fp(sdiv(len(sdf),m["delivered"],pct=True)),
            len(sdf),fa(sdf["Subtotal"].sum()),sdiv(len(sdf),m["delivered"],pct=True))

    H+=sec("PRO vs STANDARD ORDERS")
    pro=del_in[del_in["Is Pro Order"]=="Y"]; std=del_in[del_in["Is Pro Order"]=="N"]
    H+=tbl(["Type","Orders","GMV","Avg Order","Share"],
        [["Pro Orders",str(len(pro)),fa(pro["Subtotal"].sum()),fa(sdiv(pro["Subtotal"].sum(),max(len(pro),1),1),1),bdg(fp(sdiv(len(pro),m["delivered"],pct=True)),"pu")],
         ["Standard",str(len(std)),fa(std["Subtotal"].sum()),fa(sdiv(std["Subtotal"].sum(),max(len(std),1),1),1),bdg(fp(sdiv(len(std),m["delivered"],pct=True)),"bl")]],
        "#8b5cf6",r_cols={1,2,3},c_cols={4})

    H+=sec("GMV BY OUTLET")
    H+=bar([(a,round(o["gmv"])) for a,_,o in outlet_data],"#f97316")

    H+=sec("AVG ORDER VALUE BY OUTLET")
    H+=bar([(a,round(o["avg_order"],1)) for a,_,o in outlet_data],"#3b82f6",show_val=True)

    H+=sec("DAY-OF-WEEK PATTERN")
    dow_order=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    dw=df[df["_area"].isin(areas)].groupby("_dow")["Order ID"].count().reindex(dow_order).fillna(0)
    mx_dw=dw.max()
    H+=bar([(d,int(c)) for d,c in dw.items()],"#8b5cf6")
    H+='</div></div>'
    H+=pf("Page 2 – Sales & Financial Analysis")
    H+='</div>'

    # ═══════════════════════════════════════════════════════════
    # PAGE 3 — CANCELLATION INTELLIGENCE
    # ═══════════════════════════════════════════════════════════
    H+='<div class="page">'
    H+=ph("Cancellation Intelligence","Root Cause Analysis, Item-Level Failures & Recovery Roadmap")

    H+=sec("CANCELLATION SCORECARD BY OUTLET",True)
    max_can_rate=max(o["can_rate"] for _,_,o in outlet_data) or 1
    csr=[]
    for area,_,o in outlet_data:
        sc=sub_can[sub_can["_area"]==area]
        ow=sc["Cancellation owner"].value_counts()
        item_na=int(sc["Cancellation reason"].str.contains("Item not available",na=False).sum())
        fraud=int(sc["Cancellation reason"].str.contains("Fraudulent",na=False).sum())
        other_r=len(sc)-item_na-fraud
        # Severity score
        sev=min(int(o["can_rate"]*2),10)
        ring_cls="ring-re" if o["can_rate"]>20 else("ring-am" if o["can_rate"]>10 else"ring-gr")
        score_html='<div class="score-ring {}">{}/10</div>'.format(ring_cls,sev)
        csr.append([area,score_html,str(o["cancelled"]),
            bdg(fp(o["can_rate"]),can_cls(o["can_rate"])),
            str(int(ow.get("Vendor",0))),str(int(ow.get("Customer",0))),str(int(ow.get("Rider",0))),
            str(item_na),str(fraud),str(other_r),
            fa(o["lost_gmv"])])
    H+=tbl(["Outlet","Severity","#Can","Rate","Vendor","Customer","Rider","ItemN/A","Fraud","Other","Lost GMV"],
           csr,"#ef4444",r_cols={2,4,5,6,7,8,9,10},c_cols={1,3})

    H+='<div class="two-col" style="margin-top:12px;">'
    H+='<div>'
    H+=sec("CANCELLATION REASON BREAKDOWN")
    rg=sub_can.groupby(["Cancellation reason","Cancellation owner"]).agg(
        Count=("Order ID","count"),Lost=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False)
    rr=[]
    for _,r in rg.iterrows():
        rr.append([str(r["Cancellation reason"])[:32],str(r["Cancellation owner"]).strip(),
                   str(int(r["Count"])),bdg(fp(sdiv(r["Count"],len(sub_can),pct=True)),"re" if r["Count"]>5 else"am"),
                   fa(r["Lost"])])
    H+=tbl(["Reason","Owner","Count","% Share","Lost GMV"],rr,"#dc2626",r_cols={2,4},c_cols={3})

    H+=sec("CANCELLATION ALERTS & ACTION ITEMS")
    for area,_,o in outlet_data:
        sc=sub_can[sub_can["_area"]==area]
        top_r=sc["Cancellation reason"].value_counts()
        if o["can_rate"]>40:
            H+=al("<b>{}</b>: CRITICAL {:.1f}% cancel rate — {} orders lost. Revenue impact: {}. Escalate immediately — review inventory management and order acceptance process.".format(area,o["can_rate"],o["cancelled"],fa(o["lost_gmv"])),"r","🚨")
        elif o["can_rate"]>10:
            top=str(top_r.index[0])[:40] if len(top_r)>0 else "unknown"
            H+=al("<b>{}</b>: Elevated {:.1f}% cancel rate. Primary driver: <em>{}</em>. Lost revenue: {}. Review stock availability for high-demand items.".format(area,o["can_rate"],top,fa(o["lost_gmv"])),"a","⚠️")
        else:
            H+=al("<b>{}</b>: Healthy {:.1f}% cancel rate. Maintain current inventory and acceptance practices.".format(area,o["can_rate"]),"g","✅")
    H+='</div>'

    H+='<div>'
    H+=sec("CANCELLED ITEMS — STOCK GAP ANALYSIS")
    cancelled_items=[]
    for row in sub_can["Order Items"].dropna():
        for item in str(row).split(","):
            item=item.strip(); parts=item.split(" ",1)
            if len(parts)==2 and parts[0].isdigit(): qty,name=int(parts[0]),parts[1].strip()
            else: qty,name=1,item
            if name and len(name)>5: cancelled_items.append(name)
    ctr_can=Counter(cancelled_items)
    top_can=[(n,c) for n,c in ctr_can.most_common(15) if len(n)>6 and not re.match(r'^[\d\s\.,gGmMkKlLx]+$',n)]
    if top_can:
        cir=[[str(i+1),"<b>"+name[:30]+"</b>",str(qty),
              bdg("HIGH","re") if qty>=3 else bdg("MEDIUM","am") if qty>=2 else bdg("LOW","bl")]
             for i,(name,qty) in enumerate(top_can[:12])]
        H+=tbl(["#","Item Name","Orders Lost","Priority"],cir,"#ef4444",r_cols={2},c_cols={0,3})
        H+=al("These items are causing cancellations due to stock unavailability. Prioritize restocking top items to recover lost revenue.","pu","📦")
    else:
        H+='<p style="color:#94a3b8;font-size:11px;padding:8px 0;">No item-level data available in cancelled orders.</p>'

    H+=sec("DAILY CANCELLATION TREND")
    dc=df[df["_area"].isin(areas)].groupby("_date").agg(
        total=("Order ID","count"),
        cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
        lost=("Subtotal","sum")).reset_index()
    dc["can_pct"]=dc.apply(lambda r:sdiv(r["cancelled"],r["total"],pct=True),axis=1)
    dc_rows=[]
    for _,r in dc.iterrows():
        dc_rows.append([str(r["_date"]),int(r["total"]),int(r["cancelled"]),
                        bdg(fp(r["can_pct"]),can_cls(r["can_pct"])),fa(r["lost"])])
    H+=tbl(["Date","Orders","Cancelled","Rate","Lost GMV"],dc_rows,"#dc2626",r_cols={1,2,4},c_cols={3})
    H+='</div></div>'
    H+=pf("Page 3 – Cancellation Intelligence")
    H+='</div>'

    # ═══════════════════════════════════════════════════════════
    # PAGE 4 — OPERATIONAL EFFICIENCY
    # ═══════════════════════════════════════════════════════════
    H+='<div class="page">'
    H+=ph("Operational Efficiency","Delivery Timing, Stage Analysis, Bottleneck Identification")

    H+=sec("ORDER JOURNEY — AVERAGE TIME PER STAGE (GROUP)",True)
    all_del=del_in
    stage_data=[]
    for lbl,s,e in [("Accept","Order received at","Accepted at"),
                    ("Prepare","Accepted at","Ready to pick up at"),
                    ("Rider Wait","Ready to pick up at","Rider near pickup at"),
                    ("Rider→Transit","Rider near pickup at","In delivery at"),
                    ("Last Mile","In delivery at","Delivered at")]:
        if s in all_del.columns and e in all_del.columns:
            diff=(all_del[e]-all_del[s]).dt.total_seconds()/60; diff=diff[diff>0]
            stage_data.append((lbl,round(diff.mean(),1) if len(diff)>0 else 0))
    total_flow=sum(v for _,v in stage_data)
    H+='<div class="flow">'
    for lbl,val in stage_data:
        pct_of_total=sdiv(val,total_flow,pct=True)
        is_bottleneck = lbl in ["Prepare","Last Mile"] and val>12
        H+='<div class="fstep"><div class="fbox {}">{}<div class="f-l">{}</div><div class="f-v">{:.0f}</div><div class="f-u">min ({:.0f}%)</div></div></div>'.format(
            "highlight" if is_bottleneck else "",
            '<div style="font-size:8px;color:#f97316;font-weight:700;margin-bottom:3px;">⚠ BOTTLENECK</div>' if is_bottleneck else "",
            lbl,val,pct_of_total)
    H+='</div>'

    H+=sec("DELIVERY TIME BREAKDOWN BY OUTLET — WITH PERCENTILES")
    timing_rows=[]
    for area,_,o in outlet_data:
        sd_a=del_in[del_in["_area"]==area]
        tot=(sd_a["Delivered at"]-sd_a["Order received at"]).dt.total_seconds()/60; tot=tot[tot>0]
        prep=(sd_a["Ready to pick up at"]-sd_a["Accepted at"]).dt.total_seconds()/60; prep=prep[prep>0]
        lm=(sd_a["Delivered at"]-sd_a["In delivery at"]).dt.total_seconds()/60; lm=lm[lm>0]
        rider=(sd_a["Rider near pickup at"]-sd_a["Ready to pick up at"]).dt.total_seconds()/60; rider=rider[rider>0]
        avg_tot=round(tot.mean(),1) if len(tot)>0 else None
        p75=round(tot.quantile(.75),1) if len(tot)>0 else None
        p90=round(tot.quantile(.9),1) if len(tot)>0 else None
        timing_rows.append([area,
            fm(round(prep.mean(),1) if len(prep)>0 else None),
            fm(round(rider.mean(),1) if len(rider)>0 else None),
            fm(round(lm.mean(),1) if len(lm)>0 else None),
            "<b>{}</b>".format(fm(avg_tot)),
            fm(p75),fm(p90),
            bdg("Fast","gr") if avg_tot and avg_tot<=35 else(bdg("Slow","re") if avg_tot and avg_tot>45 else bdg("OK","am")),
            str(len(tot))])
    H+=tbl(["Outlet","Prep","Rider Wait","Last Mile","Avg Total","P75","P90","Grade","n"],
           timing_rows,"#1e293b",r_cols={8},c_cols={4,5,6,7})

    H+=sec("HOURLY DEMAND HEATMAP — OUTLET COMPARISON")
    hour_labels=["{:02d}".format(h) for h in range(8,24)]
    H+='<table style="width:100%;font-size:9px;border-collapse:collapse;margin-bottom:12px;">'
    H+='<thead><tr><th style="background:#1e293b;color:white;padding:5px 6px;text-align:left;width:100px;">Outlet</th>'
    for h in hour_labels:
        H+='<th style="background:#1e293b;color:white;padding:5px 3px;text-align:center;">{}</th>'.format(h)
    H+='<th style="background:#1e293b;color:white;padding:5px 6px;text-align:center;">Peak</th>'
    H+='</tr></thead><tbody>'
    dfa=df[df["_area"].isin(areas)]
    for area,_,o in outlet_data:
        s=dfa[dfa["_area"]==area]
        hourly=s.groupby("_hour")["Order ID"].count()
        mx=hourly.max() if len(hourly)>0 else 1
        color=CLRS.get(area,"#64748b")
        H+='<tr><td style="padding:5px 6px;font-weight:600;color:#0f172a;border-bottom:1px solid #f8fafc;">{}</td>'.format(area)
        for h in range(8,24):
            cnt=int(hourly.get(h,0))
            if cnt==0:
                H+='<td style="padding:3px;border-bottom:1px solid #f8fafc;"><div style="background:#f8fafc;border-radius:3px;height:18px;width:100%;"></div></td>'
            else:
                intensity=cnt/mx
                alpha=0.15+intensity*0.75
                H+='<td style="padding:3px;border-bottom:1px solid #f8fafc;" title="{}:00 — {} orders"><div style="background:{};opacity:{:.2f};border-radius:3px;height:18px;width:100%;display:flex;align-items:center;justify-content:center;font-size:8px;font-weight:700;color:white;opacity:1;background:rgba({},{},{},{:.2f});">{}</div></td>'.format(
                    h,cnt,color,alpha,
                    int(color[1:3],16),int(color[3:5],16),int(color[5:7],16),alpha,cnt)
        peak_h=hourly.idxmax() if len(hourly)>0 else 0
        H+='<td style="text-align:center;padding:5px;border-bottom:1px solid #f8fafc;"><b>{:02d}:00</b></td>'.format(int(peak_h))
        H+='</tr>'
    H+='</tbody></table>'

    H+='<div class="two-col">'
    H+='<div>'
    H+=sec("DAY-OF-WEEK ORDER DISTRIBUTION")
    dow_order2=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
    full_dow=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    dw2=dfa.groupby("_dow")["Order ID"].count().reindex(full_dow).fillna(0)
    mx2=dw2.max()
    for day_full,day_short in zip(full_dow,dow_order2):
        cnt=int(dw2.get(day_full,0))
        pct=cnt/mx2*100 if mx2>0 else 0
        H+=('<div class="br"><div class="br-l">{}</div>'
            '<div class="br-t"><div class="br-f" style="width:{:.0f}%;background:#8b5cf6;"></div></div>'
            '<div class="br-v">{}</div></div>').format(day_full[:3],pct,cnt)
    H+='</div>'
    H+='<div>'
    H+=sec("TIMING PERFORMANCE ALERTS")
    for area,_,o in outlet_data:
        sd_a=del_in[del_in["_area"]==area]
        tot=(sd_a["Delivered at"]-sd_a["Order received at"]).dt.total_seconds()/60; tot=tot[tot>0]
        if o["del_time"] is None: continue
        if o["del_time"]>45:
            H+=al("<b>{}</b>: Slow avg delivery {:.0f}m. P90={:.0f}m — 10% of orders take over {:.0f}m. Review rider allocation & prep workflow.".format(area,o["del_time"],round(tot.quantile(.9),0),round(tot.quantile(.9),0)),"r","🔴")
        elif o["del_time"]>35:
            H+=al("<b>{}</b>: Acceptable avg delivery {:.0f}m. Optimise prep stage to hit sub-35m target.".format(area,o["del_time"]),"a","🟡")
        else:
            H+=al("<b>{}</b>: Excellent avg delivery {:.0f}m. Benchmark met — maintain this standard.".format(area,o["del_time"]),"g","🟢")

    H+=sec("OPERATIONAL COST EFFICIENCY")
    oc=[]
    for area,_,o in outlet_data:
        total_cost=o["commission"]+o["op_charges"]+o["online_fee"]
        cost_per_order=sdiv(total_cost,o["delivered"],1)
        oc.append([area,fa(o["commission"]),fa(o["op_charges"]),fa(o["online_fee"]),
                   "<b>{}</b>".format(fa(total_cost)),fa(cost_per_order,1)])
    H+=tbl(["Outlet","Commission","Op Charges","Online Fee","Total Cost","Cost/Order"],
           oc,"#f97316",r_cols={1,2,3,4,5})
    H+='</div></div>'
    H+=pf("Page 4 – Operational Efficiency")
    H+='</div>'

    # ═══════════════════════════════════════════════════════════
    # PAGE 5 — PRODUCT & CUSTOMER INTELLIGENCE
    # ═══════════════════════════════════════════════════════════
    H+='<div class="page">'
    H+=ph("Product & Customer Intelligence","Top Items, Basket Analysis, Customer Behaviour & Engagement")

    H+='<div class="two-col">'
    H+='<div>'
    H+=sec("TOP SOLD ITEMS — GROUP (QTY ORDERED)",True)
    top_del=extract_items(del_in["Order Items"],top=20)
    if top_del:
        tr=[[str(i+1),"<b>"+name[:28]+"</b>" if i<3 else name[:28],str(qty),
             bdg("⭐ STAR","am") if i<3 else bdg("TOP 20","bl")]
            for i,(name,qty) in enumerate(top_del[:15])]
        H+=tbl(["#","Item Name","Qty Sold","Rank"],tr,"#10b981",r_cols={2},c_cols={0,3})
    H+=sec("BASKET SIZE ANALYSIS")
    def count_items(s):
        if pd.isna(s): return 0
        return len([x for x in str(s).split(',') if x.strip()])
    del_copy=del_in.copy(); del_copy["item_count"]=del_copy["Order Items"].apply(count_items)
    bsr=[]
    for area,_,o in outlet_data:
        sub=del_copy[del_copy["_area"]==area]
        avg_items=sub["item_count"].mean()
        max_items=sub["item_count"].max()
        bsr.append([area,"{:.1f}".format(avg_items),str(int(max_items)),
                    bdg("Large Basket","pu") if avg_items>7 else(bdg("Medium","bl") if avg_items>4 else bdg("Small","am"))])
    bsr.append(["GROUP AVG","{:.1f}".format(del_copy["item_count"].mean()),str(int(del_copy["item_count"].max())),""])
    H+=tbl(["Outlet","Avg Items/Order","Max Items","Basket Size"],bsr,"#8b5cf6",r_cols={1,2},c_cols={3})
    H+='</div>'
    H+='<div>'
    H+=sec("CANCELLED ITEMS — STOCK GAP ANALYSIS",True)
    top_can_items=[]
    for row in sub_can["Order Items"].dropna():
        for item in str(row).split(","):
            item=item.strip(); parts=item.split(" ",1)
            if len(parts)==2 and parts[0].isdigit(): name=parts[1].strip()
            else: name=item
            if name and len(name)>6 and not re.match(r'^[\d\s\.,gGmMkKlLx]+$',name):
                top_can_items.append(name)
    ctr_can2=Counter(top_can_items)
    if ctr_can2:
        can_items_display=[(n,c) for n,c in ctr_can2.most_common(12) if len(n)>5][:12]
        cir2=[]
        for i,(name,qty) in enumerate(can_items_display):
            cir2.append([str(i+1),name[:28],str(qty),
                         bdg("URGENT","re") if qty>=3 else(bdg("ACTION","am") if qty>=2 else bdg("MONITOR","bl"))])
        H+=tbl(["#","Item","Times Cancelled","Action"],cir2,"#ef4444",r_cols={2},c_cols={0,3})
        H+=al("Stock these items immediately to convert cancellations to revenue. Each unit represents a lost order.","pu","💡")

    H+=sec("COMPLAINT ANALYSIS — CUSTOMER EXPERIENCE")
    comp_all=del_in[del_in["Has Complaint?"]=="Y"]
    if len(comp_all)>0:
        for _,row in comp_all.iterrows():
            H+=('<div style="background:#fff1f2;border-radius:8px;padding:10px 12px;margin-bottom:7px;border-left:3px solid #ef4444;">'
                '<div style="font-size:10px;font-weight:700;color:#991b1b;margin-bottom:3px;">⚠ Complaint — {}</div>'
                '<div style="font-size:10px;color:#7f1d1d;margin-bottom:3px;"><b>Reason:</b> {}</div>'
                '<div style="font-size:9px;color:#94a3b8;">Order: {} · Value: {}</div></div>').format(
                row["_area"],row.get("Complaint Reason","Unknown"),
                row.get("Order ID","N/A"),fa(row["Subtotal"]))
    else:
        H+=al("No complaints recorded in this period. Excellent customer experience!","g","✅")

    H+=sec("CUSTOMER PAYMENT BEHAVIOUR BY OUTLET")
    cpb=[]
    for area,_,o in outlet_data:
        sdo=del_in[del_in["_area"]==area]
        on=len(sdo[sdo["Payment type"]=="Online"])
        cash=len(sdo[sdo["Payment type"]=="Cash"])
        pro=len(sdo[sdo["Is Pro Order"]=="Y"])
        cpb.append([area,
                    "{} ({:.0f}%)".format(on,sdiv(on,len(sdo),pct=True)),
                    "{} ({:.0f}%)".format(cash,sdiv(cash,len(sdo),pct=True)),
                    "{} ({:.0f}%)".format(pro,sdiv(pro,len(sdo),pct=True)),
                    bdg("High Pro","pu") if o["pro_pct"]>60 else(bdg("Mid Pro","bl") if o["pro_pct"]>40 else bdg("Low Pro","am"))])
    H+=tbl(["Outlet","Online Orders","Cash Orders","Pro Orders","Segment"],cpb,"#06b6d4",r_cols={1,2,3},c_cols={4})
    H+='</div></div>'
    H+=pf("Page 5 – Product & Customer Intelligence")
    H+='</div>'

    # ═══════════════════════════════════════════════════════════
    # PER-OUTLET PAGES (one each)
    # ═══════════════════════════════════════════════════════════
    for area,fname,o in outlet_data:
        color=CLRS.get(area,"#f97316")
        sub_df=df[df["_area"]==area]; sub_d=sub_df[sub_df["Order status"]=="Delivered"]
        sub_c=sub_df[sub_df["Order status"]=="Cancelled"]
        can_sc=can_cls(o["can_rate"]); del_sc=del_cls(o["del_time"])

        H+='<div class="page">'
        H+='<div class="ohdr" style="background:linear-gradient(135deg,{},{}cc);">'.format(color,color)
        H+='<h2>{}</h2>'.format(area)
        H+='<div class="store">{}</div>'.format(fname)
        H+='<div class="metrics">'
        for l,v in [("GMV",fa(o["gmv"])),("Orders",str(o["total"])),("Delivered",str(o["delivered"])),
                    ("Cancel%",fp(o["can_rate"])),("Avg Delivery",fm(o["del_time"])),
                    ("Payout",fa(o["payout"])),("Avg Order",fa(o["avg_order"],1))]:
            H+='<div><div class="mc-l">{}</div><div class="mc-v">{}</div></div>'.format(l,v)
        H+='</div></div>'

        H+='<div class="kg kg4">'
        H+=kc("Gross Revenue",fa(o["gmv"]),"{} delivered orders".format(o["delivered"]),"bl")
        H+=kc("Total Payout",fa(o["payout"]),"{:.1f}% margin".format(sdiv(o["payout"],o["gmv"],pct=True)),"gr")
        H+=kc("Cancel Rate",fp(o["can_rate"]),"AED {:,.0f} revenue lost".format(o["lost_gmv"]),"re" if o["can_rate"]>20 else("am" if o["can_rate"]>10 else"gr"))
        H+=kc("Avg Delivery",fm(o["del_time"]),"prep: {} · last mile: {}".format(fm(o["prep_time"]),fm(o["last_mile"])),"cy")
        H+='</div>'

        H+='<div class="col-6040">'
        H+='<div>'
        H+=sec("DAILY PERFORMANCE — SALES & CANCELLATIONS")
        ds3=sub_df.groupby("_date").agg(
            total=("Order ID","count"),
            delivered=("Order status",lambda x:(x=="Delivered").sum()),
            cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
            gmv=("Subtotal","sum")).reset_index()
        dp3=sub_d.groupby(sub_d["_date"])["Payout Amount"].sum()
        ds3["payout"]=dp3.reindex(ds3["_date"]).fillna(0).values
        ds3["can_pct"]=ds3.apply(lambda r:sdiv(r["cancelled"],r["total"],pct=True),axis=1)
        max_gmv3=ds3["gmv"].max() or 1
        daily_out=[]
        for _,r in ds3.iterrows():
            cb=bdg(fp(r["can_pct"]),can_cls(r["can_pct"]))
            gmv_bar=inline_bar(r["gmv"],max_gmv3,color)
            daily_out.append([str(r["_date"]),int(r["total"]),int(r["delivered"]),int(r["cancelled"]),cb,
                              gmv_bar+"<b> "+fa(r["gmv"])+"</b>",fa(r["payout"])])
        H+=tbl(["Date","Total","Del","Can","Rate","GMV","Payout"],daily_out,color,r_cols={1,2,3,6},c_cols={4})

        H+=sec("DELIVERY TIMING DETAIL")
        tot2=(sub_d["Delivered at"]-sub_d["Order received at"]).dt.total_seconds()/60; tot2=tot2[tot2>0]
        prep2=(sub_d["Ready to pick up at"]-sub_d["Accepted at"]).dt.total_seconds()/60; prep2=prep2[prep2>0]
        lm2=(sub_d["Delivered at"]-sub_d["In delivery at"]).dt.total_seconds()/60; lm2=lm2[lm2>0]
        rider2=(sub_d["Rider near pickup at"]-sub_d["Ready to pick up at"]).dt.total_seconds()/60; rider2=rider2[rider2>0]
        timing_detail=[]
        for stage,diff in [("Prep Time",prep2),("Rider Wait",rider2),("Last Mile",lm2),("TOTAL",tot2)]:
            if len(diff)>0:
                timing_detail.append(["<b>{}</b>".format(stage) if "TOTAL" in stage else stage,
                    "<b>{:.0f}m</b>".format(diff.mean()) if "TOTAL" in stage else "{:.0f}m".format(diff.mean()),
                    "{:.0f}m".format(diff.quantile(.25)),"{:.0f}m".format(diff.quantile(.75)),
                    "{:.0f}m".format(diff.quantile(.9)),"{:.0f}m".format(diff.max())])
        H+=tbl(["Stage","Avg","P25","P75","P90","Max"],timing_detail,color)
        H+='</div>'
        H+='<div>'
        H+=sec("FINANCIAL BREAKDOWN")
        fin_out=[]
        for l,v in [("Subtotal (GMV)",o["gmv"]),("Commission",o["commission"]),
                    ("Op Charges",o["op_charges"]),("Online Fee",o["online_fee"]),("Total Payout",o["payout"])]:
            fin_out.append(["<b>{}</b>".format(l) if "Payout" in l else l,
                           "<b>{}</b>".format(fa(v)) if "Payout" in l else fa(v),
                           fp(sdiv(v,o["gmv"],pct=True)),fa(sdiv(v,max(o["delivered"],1),1),1)])
        H+=tbl(["Item","Total","% GMV","Per Order"],fin_out,color,r_cols={1,2,3})

        if len(sub_c)>0:
            H+=sec("CANCELLATION DETAIL")
            cg=sub_c.groupby(["Cancellation owner","Cancellation reason"]).agg(
                Count=("Order ID","count"),Lost=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False)
            can_out=[[str(r["Cancellation owner"]).strip(),str(r["Cancellation reason"])[:28],
                      int(r["Count"]),fa(r["Lost"])] for _,r in cg.iterrows()]
            H+=tbl(["Owner","Reason","#","Lost GMV"],can_out,"#ef4444",r_cols={2,3})

            out_can_items=[]
            for row in sub_c["Order Items"].dropna():
                for item in str(row).split(","):
                    item=item.strip(); parts=item.split(" ",1)
                    n=parts[1].strip() if len(parts)==2 and parts[0].isdigit() else item
                    if n and len(n)>6: out_can_items.append(n)
            ctr3=Counter(out_can_items)
            top3=[(n,c) for n,c in ctr3.most_common(8) if len(n)>5]
            if top3:
                H+=sec("CANCELLED ITEMS — RESTOCK REQUIRED")
                H+=tbl(["Item","Lost Orders"],[[name[:28],str(qty)] for name,qty in top3],"#ef4444",r_cols={1})

        H+=sec("HOURLY ORDERS")
        h2=sub_df.groupby("_hour")["Order ID"].count().reset_index()
        mx_h=h2["Order ID"].max() or 1
        for _,r in h2.iterrows():
            pct=r["Order ID"]/mx_h*100
            H+=('<div class="br"><div class="br-l">{:02d}:00–{:02d}:00</div>'
                '<div class="br-t"><div class="br-f" style="width:{:.0f}%;background:{};"></div></div>'
                '<div class="br-v">{} orders</div></div>').format(int(r["_hour"]),int(r["_hour"])+1,pct,color,int(r["Order ID"]))
        H+='</div></div>'
        H+=pf("Outlet Report – {}".format(area))
        H+='</div>'

    # ═══════════════════════════════════════════════════════════
    # FINAL PAGE — STRATEGIC RECOMMENDATIONS
    # ═══════════════════════════════════════════════════════════
    H+='<div class="page">'
    H+=ph("Strategic Recommendations","Data-Driven Actions to Boost Performance & Revenue")

    H+=sec("PRIORITY ACTION MATRIX",True)
    H+='<div class="insight-grid">'

    # High cancel outlets
    high_can=[(a,o) for a,_,o in outlet_data if o["can_rate"]>15]
    if high_can:
        names=", ".join(a for a,_ in high_can)
        H+=('<div class="insight re"><div class="ins-t">🚨 Critical — Cancellation Crisis</div>'
            '<div class="ins-v">Outlets: {}</div>'
            '<div class="ins-s">Implement real-time stock sync. Set up auto-pause for out-of-stock items. Target: reduce cancel rate below 10% within 2 weeks.</div></div>').format(names)

    # Slow delivery
    slow_del=[(a,o) for a,_,o in outlet_data if o["del_time"] and o["del_time"]>40]
    if slow_del:
        names2=", ".join("{} ({:.0f}m)".format(a,o["del_time"]) for a,o in slow_del)
        H+=('<div class="insight or"><div class="ins-t">⏱ Delivery Optimisation</div>'
            '<div class="ins-v">{}</div>'
            '<div class="ins-s">Review kitchen prep workflow. Consider pre-staging high-demand items during peak hours (11:00–12:00 and 20:00–21:00). Target: &lt;35 min average.</div></div>').format(names2)

    # Revenue recovery


    H+='<div class="insight-grid">'
    H+=('<div class="insight re"><div class="ins-t">🚨 Revenue Lost to Cancellations</div>'
        '<div class="ins-v">AED {:,.0f} this period</div>'
        '<div class="ins-s">Restocking top 5 cancelled items could recover up to 70% of lost revenue. Implement daily stock audit at opening.</div></div>').format(m["lost_gmv"])
    H+=('<div class="insight or"><div class="ins-t">⏱ Delivery Speed Improvement</div>'
        '<div class="ins-v">Group Avg: {}</div>'
        '<div class="ins-s">Reducing avg delivery by 5 minutes improves customer satisfaction score significantly. Focus on prep stage which accounts for the largest bottleneck.</div></div>').format(fm(m["del_time"]))
    H+=('<div class="insight gr"><div class="ins-t">📈 Pro Order Growth</div>'
        '<div class="ins-v">{:.0f}% Current Pro Rate</div>'
        '<div class="ins-s">Outlets with higher Pro order % show better basket values. Run targeted Pro subscription campaigns in low-Pro outlets to increase avg order value.</div></div>').format(m["pro_pct"])
    H+=('<div class="insight bl"><div class="ins-t">📦 Peak Hour Staffing</div>'
        '<div class="ins-v">Peak: 11:00 &amp; 20:00–21:00</div>'
        '<div class="ins-s">Ensure maximum staff coverage during peak hours. Pre-pack common items before rush. These 2 windows generate ~30% of daily orders.</div></div>')
    H+='</div>'

    H+=sec("KPI BENCHMARKS & TARGETS")
    benchmarks=[
        ["Cancellation Rate","Current: {}".format(fp(m["can_rate"])),"Target: &lt;8%",
         bdg("CRITICAL","re") if m["can_rate"]>20 else(bdg("AT RISK","am") if m["can_rate"]>10 else bdg("ON TARGET","gr"))],
        ["Avg Delivery Time","Current: {}".format(fm(m["del_time"])),"Target: &lt;35 min",
         bdg("SLOW","re") if m["del_time"] and m["del_time"]>45 else(bdg("OK","am") if m["del_time"] and m["del_time"]>35 else bdg("GOOD","gr"))],
        ["Avg Order Value","Current: {}".format(fa(m["avg_order"],1)),"Target: AED 75+",
         bdg("BELOW","am") if m["avg_order"]<75 else bdg("ON TARGET","gr")],
        ["Complaint Rate","Current: {}".format(fp(m["complaint_rate"])),"Target: &lt;1%",
         bdg("GOOD","gr") if m["complaint_rate"]<1 else bdg("WATCH","am")],
        ["Pro Order %","Current: {}".format(fp(m["pro_pct"])),"Target: &gt;65%",
         bdg("GROWING","am") if m["pro_pct"]<65 else bdg("STRONG","gr")],
        ["Payout Margin","Current: {:.1f}%".format(sdiv(m["payout"],m["gmv"],pct=True)),"Target: &gt;75%",
         bdg("WATCH","am") if sdiv(m["payout"],m["gmv"],pct=True)<75 else bdg("HEALTHY","gr")],
    ]
    H+=tbl(["KPI","Current Value","Target","Status"],benchmarks,"#0f172a",c_cols={3})

    H+=sec("30-DAY IMPROVEMENT ROADMAP")
    H+=tbl(["Priority","Action","Owner","Expected Impact","Timeline"],
        [["🔴 P1","Implement daily stock audit & auto-pause OOS items","Ops Manager","Reduce cancel rate by 50%","Week 1"],
         ["🔴 P1","Restock top 5 cancelled items (Takis, Frangosul Chicken)","Procurement","Recover AED {:,.0f}+/period".format(int(m["lost_gmv"]*0.7)),"Week 1"],
         ["🟡 P2","Review prep workflow at slow outlets","Branch Managers","–5 min avg delivery","Week 2"],
         ["🟡 P2","Increase staffing during 11:00–12:00 & 20:00–21:00","HR / Ops","+15% peak hour throughput","Week 2"],
         ["🟢 P3","Launch Pro subscription push campaign","Marketing","Increase Pro % to 65%+","Week 3–4"],
         ["🟢 P3","Set up customer callback for wrong-item complaints","CX Team","Reduce complaint churn","Week 3–4"]],
        "#0f172a",c_cols={0,3,4})

    H+=pf("Page {} – Strategic Recommendations".format(len(outlet_data)+6))
    H+='</div>'
    H+="</body></html>"
    return H.encode("utf-8")

# ─── PASSWORD GATE ────────────────────────────────────────────
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col = st.columns([1,2,1])[1]
    with col:
        st.markdown("""<div style="background:white;border-radius:24px;padding:44px 40px;
            box-shadow:0 24px 80px rgba(0,0,0,.1);border:1px solid #f1f5f9;text-align:center;">
            <div style="width:60px;height:60px;border-radius:16px;background:linear-gradient(135deg,#f97316,#dc2626);
                display:flex;align-items:center;justify-content:center;font-size:26px;margin:0 auto 18px;">📊</div>
            <div style="font-size:22px;font-weight:800;color:#0f172a;margin-bottom:4px;letter-spacing:-.02em;">Al Madina Reports</div>
            <div style="font-size:12px;color:#94a3b8;margin-bottom:28px;">Confidential reporting platform</div>
        </div>""", unsafe_allow_html=True)
        st.markdown("<div style='margin-top:-10px;'>", unsafe_allow_html=True)
        pwd = st.text_input("", type="password", placeholder="Enter password…", label_visibility="collapsed", key="pwd_input")
        if st.button("Sign In →", use_container_width=True, key="signin"):
            if pwd == PASSWORD:
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")
        st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# ─── MAIN APP ─────────────────────────────────────────────────
hcol, scol = st.columns([7,1])
with hcol:
    st.markdown("""<div class="brand">
        <div class="b-icon">📊</div>
        <div>
            <div class="b-title">Al Madina Report Generator</div>
            <div class="b-sub">Upload order export → select outlets → download full performance report</div>
        </div></div>""", unsafe_allow_html=True)
with scol:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔒 Out", key="signout", help="Sign out"):
        st.session_state.auth = False; st.rerun()

uploaded = st.file_uploader("", type=["xlsx","xls"], label_visibility="collapsed",
    help="Upload your Talabat order details .xlsx export")

if uploaded is None:
    st.markdown("""<div class="ubox">
        <div class="u-icon">📂</div>
        <div class="u-title">Drop your Talabat order export here</div>
        <div class="u-sub">Supports .xlsx &nbsp;·&nbsp; All outlets auto-detected &nbsp;·&nbsp; Zero dependencies</div>
    </div>""", unsafe_allow_html=True)
    st.markdown("""<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">
        <div style="background:white;border:1px solid #e2e8f0;border-radius:12px;padding:16px;">
            <div style="font-size:20px;margin-bottom:8px;">📈</div>
            <div style="font-size:12px;font-weight:700;color:#0f172a;margin-bottom:4px;">6 Report Sections</div>
            <div style="font-size:11px;color:#94a3b8;">Executive dashboard, sales, cancellations, operations, products, strategy</div>
        </div>
        <div style="background:white;border:1px solid #e2e8f0;border-radius:12px;padding:16px;">
            <div style="font-size:20px;margin-bottom:8px;">🏪</div>
            <div style="font-size:12px;font-weight:700;color:#0f172a;margin-bottom:4px;">Per-Outlet Pages</div>
            <div style="font-size:11px;color:#94a3b8;">Individual deep-dive for each outlet with full timing, items, cancellations</div>
        </div>
        <div style="background:white;border:1px solid #e2e8f0;border-radius:12px;padding:16px;">
            <div style="font-size:20px;margin-bottom:8px;">🎯</div>
            <div style="font-size:12px;font-weight:700;color:#0f172a;margin-bottom:4px;">Action Roadmap</div>
            <div style="font-size:11px;color:#94a3b8;">30-day improvement plan with priority actions and expected impact</div>
        </div>
    </div>""", unsafe_allow_html=True)
    st.stop()

with st.spinner("Reading file..."):
    try:
        df = load_data(uploaded.read())
    except Exception as e:
        st.error("Could not read file: {}".format(e)); st.stop()

if "Order status" not in df.columns:
    st.error("File doesn't look like a Talabat order report."); st.stop()

outlets_all = [a for a in sorted(df["_area"].unique()) if a != "Other"]
date_min = df["Order received at"].min()
date_max = df["Order received at"].max()
ds_str = ""
if pd.notna(date_min) and pd.notna(date_max):
    ds_str = "{} – {}".format(date_min.strftime("%d %b %Y"), date_max.strftime("%d %b %Y"))
m_g = calc_quick(df[df["_area"].isin(outlets_all)])

st.markdown("""<div class="sbar">
    <span style="font-size:20px;">✅</span>
    <div>
        <div style="font-size:13px;font-weight:700;color:#14532d;">{orders} orders across {n} outlets loaded — {period}</div>
        <div style="font-size:11px;color:#166534;margin-top:1px;">Ready to generate your report</div>
    </div></div>""".format(orders=m_g["total"],n=len(outlets_all),period=ds_str),
    unsafe_allow_html=True)

st.markdown("""<div class="chips">
    <div class="chip">💰 <b>{gmv}</b> GMV</div>
    <div class="chip">✅ <b>{d}</b> delivered</div>
    <div class="chip">❌ <b>{c}</b> cancelled (<b>{cr}</b>)</div>
    <div class="chip">💸 <b>{p}</b> payout</div>
    <div class="chip">🔴 <b>{l}</b> lost to cancels</div>
    <div class="chip">⏱ <b>{t}</b> avg delivery</div>
</div>""".format(gmv=fa(m_g["gmv"]),d=m_g["delivered"],c=m_g["cancelled"],
    cr=fp(m_g["can_rate"]),p=fa(m_g["payout"]),l=fa(m_g["lost_gmv"]),t=fm(m_g["del_time"])),
    unsafe_allow_html=True)

# ── GROUP REPORT ──────────────────────────────────────────────
st.markdown('<div class="rcard">', unsafe_allow_html=True)
st.markdown('<div class="rcard-title">🏢 Full Group Report</div>', unsafe_allow_html=True)
st.markdown('<div class="rcard-sub">Comprehensive multi-outlet performance report with all analytics sections. Download as HTML and open in Chrome → Ctrl+P → Save as PDF to get a print-ready PDF.</div>', unsafe_allow_html=True)

sel = st.multiselect("Select outlets to include", options=outlets_all, default=outlets_all, key="grp_sel")

if sel:
    for area in sel:
        s = df[df["_area"]==area]
        if len(s)==0: continue
        o = calc_quick(s)
        color = CLRS.get(area,"#64748b")
        can_color = "#ef4444" if o["can_rate"]>20 else ("#f59e0b" if o["can_rate"]>10 else "#10b981")
        st.markdown(
            '<div class="outlet-strip">'
            '<div><div class="os-name">{}</div>'
            '<div class="os-meta">{} orders &nbsp;·&nbsp; {} &nbsp;·&nbsp; {} payout</div></div>'
            '<div><span class="os-badge" style="background:{}20;color:{};">{} cancel</span></div>'
            '</div>'.format(area,o["total"],fa(o["gmv"]),fa(o["payout"]),
                           can_color,can_color,fp(o["can_rate"])),
            unsafe_allow_html=True)

    st.markdown('<ul class="rcard-feat">', unsafe_allow_html=True)
    for feat in ["Executive Dashboard — KPIs, league table, daily trend",
                 "Sales & Financial — costs, payout waterfall, payment mix, profitability",
                 "Cancellation Intelligence — root cause, item analysis, alerts, daily trend",
                 "Operational Efficiency — timing flow, heatmap, hourly demand, DOW pattern",
                 "Product & Customer Intelligence — top items, basket analysis, complaints, behaviour",
                 "Per-Outlet Deep Dives — full detail for each selected outlet",
                 "Strategic Recommendations — 30-day action roadmap with KPI targets"]:
        st.markdown('<li>{}</li>'.format(feat), unsafe_allow_html=True)
    st.markdown('</ul>', unsafe_allow_html=True)

    if st.button("📄 Generate Full Group Report", use_container_width=True, type="primary", key="gen_grp"):
        with st.spinner("Building comprehensive report... ({} outlets, all sections)".format(len(sel))):
            try:
                html = build_full_report(df, sel, "Full Performance Report — Al Madina Hypermarket")
                fname = "AlMadina_Full_Report_{}.html".format(datetime.now().strftime("%Y%m%d_%H%M"))
                st.download_button("💾 Download Full Report", data=html, file_name=fname,
                                   mime="text/html", use_container_width=True, key="dl_grp")
                st.info("**To get PDF:** Open in Chrome → Ctrl+P (or Cmd+P) → Destination: Save as PDF → Layout: Landscape → Print", icon="💡")
            except Exception as e:
                st.error("Report failed: {}".format(e))
                import traceback; st.code(traceback.format_exc())
elif not sel:
    st.warning("Please select at least one outlet.")
st.markdown('</div>', unsafe_allow_html=True)

# ── INDIVIDUAL OUTLET REPORTS ─────────────────────────────────
st.markdown('<div class="rcard">', unsafe_allow_html=True)
st.markdown('<div class="rcard-title">🏪 Individual Outlet Reports</div>', unsafe_allow_html=True)
st.markdown('<div class="rcard-sub">Dedicated report for a single outlet — includes all sections focused on that outlet. Share directly with branch managers or area supervisors.</div>', unsafe_allow_html=True)

for area in outlets_all:
    s = df[df["_area"]==area]
    if len(s)==0: continue
    o = calc_quick(s)
    color = CLRS.get(area,"#64748b")
    can_color = "#ef4444" if o["can_rate"]>20 else ("#f59e0b" if o["can_rate"]>10 else "#10b981")

    with st.expander("📍 {} — {} · {} orders · {} cancel rate".format(area, fa(o["gmv"]), o["total"], fp(o["can_rate"]))):
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("GMV", fa(o["gmv"]))
        c2.metric("Delivered", str(o["delivered"]), fp(o["del_rate"]))
        c3.metric("Cancelled", str(o["cancelled"]),
                  "⚠ " + fp(o["can_rate"]) if o["can_rate"]>10 else fp(o["can_rate"]))
        c4.metric("Avg Delivery", fm(o["del_time"]))

        if st.button("📄 Generate {} Report".format(area), use_container_width=True,
                     type="primary", key="gen_out_"+area):
            with st.spinner("Building {} report...".format(area)):
                try:
                    html = build_full_report(df, [area], "{} Outlet Report — Al Madina Hypermarket".format(area))
                    fname = "AlMadina_{}_{}.html".format(area.replace(" ","_"), datetime.now().strftime("%Y%m%d_%H%M"))
                    st.download_button("💾 Download {} Report".format(area), data=html,
                        file_name=fname, mime="text/html", use_container_width=True,
                        key="dl_out_"+area)
                    st.info("Open in Chrome → Ctrl+P → Save as PDF", icon="💡")
                except Exception as e:
                    st.error("Failed: {}".format(e))
                    import traceback; st.code(traceback.format_exc())

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div style="text-align:center;font-size:10px;color:#cbd5e1;padding:20px 0;">Al Madina Hypermarket &nbsp;·&nbsp; Confidential Reporting Platform &nbsp;·&nbsp; Zero external dependencies</div>', unsafe_allow_html=True)
