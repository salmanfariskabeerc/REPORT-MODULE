import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import zipfile, re, xml.etree.ElementTree as ET
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Al Madina Reports", page_icon="📊",
                   layout="centered", initial_sidebar_state="collapsed")

PASSWORD = "123123123"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
#MainMenu,footer,header,[data-testid="stSidebar"]{display:none!important;}
.main .block-container{padding:2rem 1.5rem 4rem;max-width:860px;margin:0 auto;}
.brand-bar{background:linear-gradient(135deg,#0f1117,#1e293b);border-radius:16px;padding:22px 26px;margin-bottom:24px;display:flex;align-items:center;gap:14px;}
.brand-icon{width:46px;height:46px;border-radius:11px;background:linear-gradient(135deg,#f97316,#dc2626);display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0;}
.brand-title{font-size:18px;font-weight:700;color:white;margin:0;}
.brand-sub{font-size:11px;color:#64748b;margin-top:2px;}
.upload-box{background:white;border:2px dashed #e2e8f0;border-radius:14px;padding:36px 28px;text-align:center;margin-bottom:20px;}
.upload-box-icon{font-size:36px;margin-bottom:10px;}
.upload-box-title{font-size:16px;font-weight:600;color:#0f172a;margin-bottom:5px;}
.upload-box-sub{font-size:12px;color:#94a3b8;}
.success-bar{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;padding:13px 17px;margin-bottom:18px;display:flex;align-items:center;gap:10px;}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px;}
.chip{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:5px 11px;font-size:11px;color:#475569;}
.chip b{color:#0f172a;}
.card{background:white;border:1px solid #e2e8f0;border-radius:14px;padding:22px;margin-bottom:14px;}
.card-title{font-size:14px;font-weight:600;color:#0f172a;margin-bottom:3px;}
.card-sub{font-size:11px;color:#94a3b8;margin-bottom:16px;}
.stButton>button{background:linear-gradient(135deg,#f97316,#dc2626)!important;color:white!important;border:none!important;border-radius:10px!important;font-weight:600!important;font-size:13px!important;padding:11px 20px!important;width:100%!important;}
.stDownloadButton>button{background:linear-gradient(135deg,#0f172a,#1e293b)!important;color:white!important;border:none!important;border-radius:10px!important;font-weight:600!important;font-size:13px!important;padding:11px 20px!important;width:100%!important;}
div[data-testid="stTextInput"] input{border-radius:10px!important;border:1.5px solid #e2e8f0!important;padding:11px 13px!important;font-size:14px!important;}
div[data-testid="stTextInput"] input:focus{border-color:#f97316!important;box-shadow:0 0 0 3px rgba(249,115,22,.15)!important;}
.login-card{background:white;border-radius:20px;padding:44px 40px;max-width:400px;margin:10vh auto 0;box-shadow:0 20px 60px rgba(0,0,0,.1);border:1px solid #f1f5f9;}
.login-logo{width:54px;height:54px;border-radius:13px;background:linear-gradient(135deg,#f97316,#dc2626);display:flex;align-items:center;justify-content:center;font-size:22px;margin:0 auto 14px;}
.login-title{font-size:21px;font-weight:700;color:#0f172a;text-align:center;margin-bottom:3px;}
.login-sub{font-size:12px;color:#94a3b8;text-align:center;margin-bottom:26px;}
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────
NUMERIC_COLS = ["Subtotal","Packaging charges","Minimum order value fee","Vendor Refunds",
    "Tax Charge","Online Payment Fee","Discount Funded by you","Voucher Funded by you",
    "Commission","Operational Charges","Ads Fee","Marketing Fees","Avoidable cancellation fee",
    "Estimated earnings","Cash amount already collected by you","Amount owed back to Talabat",
    "Payout Amount","Total Discount","Total Voucher","Tax Amount"]
DATE_COLS = ["Order received at","Accepted at","Ready to pick up at","Rider near pickup at",
    "In delivery at","Delivered at","Cancelled at"]
CLRS = {"Liwan":"#f97316","Dubai Investment Park":"#3b82f6","Oud Metha":"#8b5cf6",
        "Naif":"#ef4444","Al Muteena":"#10b981","Al Hamriya":"#f59e0b"}

# ── Helpers ───────────────────────────────────────────────────
def get_area(name):
    n = str(name).lower()
    if "liwan" in n: return "Liwan"
    if "dip" in n or "investment park" in n: return "Dubai Investment Park"
    if "oud metha" in n: return "Oud Metha"
    if "naif" in n: return "Naif"
    if "muteena" in n: return "Al Muteena"
    if "hamriya" in n: return "Al Hamriya"
    return "Other"

def sdiv(a, b, pct=False):
    if b == 0: return 0
    return (a / b * 100) if pct else a / b

def calc_om(df_sub):
    d = df_sub[df_sub["Order status"] == "Delivered"]
    c = df_sub[df_sub["Order status"] == "Cancelled"]
    total = len(df_sub); gmv = d["Subtotal"].sum()
    diff  = (d["Delivered at"]-d["Order received at"]).dt.total_seconds()/60; diff=diff[diff>0]
    prep  = (d["Ready to pick up at"]-d["Accepted at"]).dt.total_seconds()/60; prep=prep[prep>0]
    lm    = (d["Delivered at"]-d["In delivery at"]).dt.total_seconds()/60; lm=lm[lm>0]
    return {"total":total,"delivered":len(d),"cancelled":len(c),
        "can_rate":sdiv(len(c),total,pct=True),"del_rate":sdiv(len(d),total,pct=True),
        "gmv":gmv,"payout":d["Payout Amount"].sum(),"commission":d["Commission"].sum(),
        "op_charges":d["Operational Charges"].sum(),"online_fee":d["Online Payment Fee"].sum(),
        "avg_order":sdiv(gmv,len(d)),
        "del_time":round(diff.mean(),1) if len(diff)>0 else None,
        "prep_time":round(prep.mean(),1) if len(prep)>0 else None,
        "last_mile":round(lm.mean(),1) if len(lm)>0 else None,
        "pro_pct":sdiv(len(d[d["Is Pro Order"]=="Y"]),len(d),pct=True),
        "online_pct":sdiv(len(d[d["Payment type"]=="Online"]),len(d),pct=True),
        "complaints":len(d[d["Has Complaint?"]=="Y"]),
        "complaint_rate":sdiv(len(d[d["Has Complaint?"]=="Y"]),len(d),pct=True),
        "lost_gmv":c["Subtotal"].sum()}

def fp(v): return "{:.1f}%".format(v)
def fm(v): return "{:.0f} min".format(v) if v else "N/A"

def extract_items(series, top=15):
    from collections import Counter
    ctr = Counter()
    for row in series.dropna():
        for item in str(row).split(","):
            item = item.strip()
            parts = item.split(" ", 1)
            if len(parts)==2 and parts[0].isdigit():
                qty, name = int(parts[0]), parts[1].strip()
            else:
                qty, name = 1, item
            if name and len(name)>4 and not re.match(r'^[\d\s\.\,gGmMkKlL]+$', name):
                ctr[name] += qty
    return ctr.most_common(top)

# ── Data Loader ────────────────────────────────────────────────
def _parse_stdlib(file_bytes):
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
        names = [i.filename for i in z.infolist()]
        sheet = z.read("xl/worksheets/sheet1.xml")
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter("{"+ns+"}si"):
                shared.append("".join(t.text or "" for t in si.iter("{"+ns+"}t")))
    def ci(ref):
        col = "".join(filter(str.isalpha, ref)); idx = 0
        for ch in col: idx = idx*26+(ord(ch.upper())-64)
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

@st.cache_data(show_spinner=False)
def load_data(file_bytes):
    df_raw = _parse_stdlib(file_bytes)
    hr = None
    for i in range(min(5,len(df_raw))):
        if "Order status" in [str(v) for v in df_raw.iloc[i].values]:
            hr=i; break
    if hr is None: hr=0
    df = df_raw.iloc[hr+1:].copy()
    df.columns = [str(c) for c in df_raw.iloc[hr].tolist()]
    df = df.dropna(how="all").reset_index(drop=True)
    for col in NUMERIC_COLS:
        if col in df.columns: df[col]=pd.to_numeric(df[col],errors="coerce").fillna(0)
    for col in DATE_COLS:
        if col in df.columns: df[col]=pd.to_datetime(df[col],errors="coerce")
    if "Restaurant name" not in df.columns:
        raise RuntimeError("Not a valid Talabat report.")
    df["_area"]=df["Restaurant name"].apply(get_area)
    df["_date"]=df["Order received at"].dt.date
    df["_hour"]=df["Order received at"].dt.hour
    df["_dow"]=df["Order received at"].dt.day_name()
    return df

# ── Report CSS ─────────────────────────────────────────────────
REPORT_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Inter',Arial,sans-serif;color:#1e293b;background:#f1f5f9;font-size:12px;line-height:1.55;}
.cover{background:linear-gradient(135deg,#0f172a 0%,#1e293b 60%,#0f172a 100%);color:white;
       min-height:100vh;display:flex;flex-direction:column;justify-content:center;
       align-items:center;text-align:center;padding:60px 40px;page-break-after:always;}
.cover-badge{background:rgba(249,115,22,.18);border:1px solid rgba(249,115,22,.5);
             color:#fb923c;padding:6px 20px;border-radius:20px;font-size:10px;font-weight:700;
             letter-spacing:.1em;text-transform:uppercase;margin-bottom:28px;display:inline-block;}
.cover-logo{width:76px;height:76px;border-radius:18px;
            background:linear-gradient(135deg,#f97316,#dc2626);
            display:flex;align-items:center;justify-content:center;
            font-size:34px;margin:0 auto 26px;}
.cover h1{font-size:36px;font-weight:700;color:white;margin-bottom:8px;letter-spacing:-.02em;}
.cover h2{font-size:20px;font-weight:400;color:#94a3b8;margin-bottom:36px;}
.cover-stats{display:flex;gap:0;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);border-radius:14px;overflow:hidden;margin-bottom:36px;}
.cover-stat{padding:18px 28px;text-align:center;border-right:1px solid rgba(255,255,255,.08);}
.cover-stat:last-child{border-right:none;}
.cs-label{font-size:9px;color:#64748b;text-transform:uppercase;letter-spacing:.1em;margin-bottom:5px;}
.cs-value{font-size:20px;font-weight:700;color:white;}
.cover-outlets{display:flex;gap:8px;flex-wrap:wrap;justify-content:center;}
.cpill{border-radius:8px;padding:7px 16px;font-size:11px;font-weight:600;border:1px solid rgba(255,255,255,.1);}

.page{background:white;margin:0;padding:32px 36px;min-height:100vh;page-break-after:always;position:relative;}
.page:last-child{page-break-after:auto;}
.page-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:24px;padding-bottom:14px;border-bottom:2px solid #f8fafc;}
.ph-title{font-size:18px;font-weight:700;color:#0f172a;margin-bottom:2px;}
.ph-sub{font-size:10px;color:#94a3b8;}
.ph-badge{background:#f8fafc;border-radius:8px;padding:5px 13px;font-size:10px;font-weight:600;color:#64748b;border:1px solid #e2e8f0;}

.sec-title{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:#94a3b8;
           margin:20px 0 10px;display:flex;align-items:center;gap:8px;}
.sec-title::after{content:'';flex:1;height:1px;background:#f1f5f9;}

.kgrid{display:grid;gap:10px;margin-bottom:6px;}
.kg3{grid-template-columns:repeat(3,1fr);}
.kg4{grid-template-columns:repeat(4,1fr);}
.kcard{background:#f8fafc;border-radius:10px;padding:12px 14px;border-left:3px solid #e2e8f0;}
.kcard.or{border-left-color:#f97316;} .kcard.gr{border-left-color:#10b981;}
.kcard.re{border-left-color:#ef4444;} .kcard.bl{border-left-color:#3b82f6;}
.kcard.pu{border-left-color:#8b5cf6;} .kcard.am{border-left-color:#f59e0b;}
.k-label{font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#94a3b8;margin-bottom:5px;}
.k-value{font-size:20px;font-weight:700;color:#0f172a;line-height:1.1;}
.k-sub{font-size:9px;color:#64748b;margin-top:3px;}

table{width:100%;border-collapse:collapse;font-size:10.5px;margin-bottom:4px;}
thead th{padding:8px 10px;font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:white;text-align:left;}
tbody td{padding:8px 10px;border-bottom:1px solid #f8fafc;color:#374151;vertical-align:middle;}
tbody tr:last-child td{border-bottom:none;}
tbody tr:nth-child(even) td{background:#fafbfc;}
td.r,th.r{text-align:right;}
td.c,th.c{text-align:center;}
td.b{font-weight:600;color:#0f172a;}

.badge{padding:2px 8px;border-radius:10px;font-size:9.5px;font-weight:600;white-space:nowrap;}
.bg{background:#d1fae5;color:#065f46;}
.br{background:#fee2e2;color:#991b1b;}
.ba{background:#fef3c7;color:#92400e;}
.bb{background:#dbeafe;color:#1e40af;}

.bar-row{display:flex;align-items:center;gap:8px;margin-bottom:7px;}
.bar-lbl{font-size:10px;color:#475569;width:150px;flex-shrink:0;text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.bar-track{flex:1;height:10px;background:#f1f5f9;border-radius:5px;overflow:hidden;}
.bar-fill{height:100%;border-radius:5px;}
.bar-val{font-size:10px;font-weight:600;color:#0f172a;min-width:60px;}

.flow{display:flex;align-items:stretch;gap:0;margin:10px 0 16px;}
.fstep{flex:1;position:relative;}
.fstep:not(:last-child)::after{content:'▶';position:absolute;right:-8px;top:50%;transform:translateY(-50%);color:#cbd5e1;font-size:10px;z-index:1;}
.fbox{background:#f8fafc;border-radius:10px;padding:10px 8px;text-align:center;border:1px solid #f1f5f9;height:100%;}
.f-lbl{font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px;}
.f-val{font-size:17px;font-weight:700;color:#0f172a;}
.f-unit{font-size:8px;color:#94a3b8;}

.two-col{display:grid;grid-template-columns:1fr 1fr;gap:18px;}
.three-col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;}

.outlet-hdr{border-radius:12px;padding:16px 20px;margin-bottom:16px;color:white;}
.oh-name{font-size:18px;font-weight:700;margin-bottom:2px;}
.oh-store{font-size:10px;opacity:.65;margin-bottom:12px;}
.oh-chips{display:flex;gap:20px;}
.oh-chip-label{font-size:8px;text-transform:uppercase;letter-spacing:.07em;opacity:.55;margin-bottom:3px;}
.oh-chip-val{font-size:16px;font-weight:700;}

.alert{border-radius:8px;padding:9px 12px;margin-bottom:8px;font-size:10px;display:flex;gap:8px;align-items:flex-start;}
.al-r{background:#fff1f2;border-left:3px solid #ef4444;color:#7f1d1d;}
.al-a{background:#fffbeb;border-left:3px solid #f59e0b;color:#78350f;}
.al-g{background:#f0fdf4;border-left:3px solid #10b981;color:#14532d;}
.al-b{background:#eff6ff;border-left:3px solid #3b82f6;color:#1e3a8a;}

.page-footer{position:absolute;bottom:20px;left:36px;right:36px;display:flex;justify-content:space-between;font-size:9px;color:#cbd5e1;border-top:1px solid #f8fafc;padding-top:10px;}

@media print{
  body{background:white;}
  .page,.cover{box-shadow:none;}
}
</style>
"""

# ── Report Builder ─────────────────────────────────────────────
def build_report(df, areas, label="Group Report"):
    delivered = df[df["Order status"]=="Delivered"]
    cancelled = df[df["Order status"]=="Cancelled"]
    date_min = df["Order received at"].min()
    date_max = df["Order received at"].max()
    ds = ""
    if pd.notna(date_min) and pd.notna(date_max):
        ds = date_min.strftime("%d %b %Y")+" – "+date_max.strftime("%d %b %Y")

    outlet_data = [(a, df[df["_area"]==a]["Restaurant name"].iloc[0], calc_om(df[df["_area"]==a]))
                   for a in sorted(areas) if len(df[df["_area"]==a])>0]
    outlet_data.sort(key=lambda x: x[2]["gmv"], reverse=True)
    m = calc_om(df[df["_area"].isin(areas)])

    # ── HTML helpers ─────────────────────────────────────────
    def kc(label, val, sub="", cls="or"):
        return ('<div class="kcard {c}"><div class="k-label">{l}</div>'
                '<div class="k-value">{v}</div>{s}</div>').format(
            c=cls, l=label, v=val,
            s='<div class="k-sub">{}</div>'.format(sub) if sub else "")

    def bdg(text, cls="bb"):
        return '<span class="badge {}">{}</span>'.format(cls, text)

    def can_cls(rate):
        return "br" if rate>20 else ("ba" if rate>10 else "bg")

    def del_cls(mins):
        if mins is None: return "ba"
        return "bg" if mins<=35 else ("br" if mins>42 else "ba")

    def tbl(headers, rows, color="#0f172a", r_cols=None, c_cols=None):
        r_cols = r_cols or set(); c_cols = c_cols or set()
        ths = "".join('<th class="{}" style="background:{}">{}</th>'.format(
            "r" if i in r_cols else ("c" if i in c_cols else ""), color, h)
            for i,h in enumerate(headers))
        body = ""
        for ri, row in enumerate(rows):
            tds = "".join('<td class="{}">{}</td>'.format(
                "r" if ci in r_cols else ("c" if ci in c_cols else ""), cell)
                for ci,cell in enumerate(row))
            body += "<tr>{}</tr>".format(tds)
        return "<table><thead><tr>{}</tr></thead><tbody>{}</tbody></table>".format(ths,body)

    def bar_chart(items, color="#f97316"):
        if not items: return ""
        mx = max(v for _,v in items) or 1
        out = ""
        for name, val in items:
            pct = min(val/mx*100, 100)
            out += ('<div class="bar-row">'
                    '<div class="bar-lbl">{}</div>'
                    '<div class="bar-track"><div class="bar-fill" style="width:{:.1f}%;background:{};"></div></div>'
                    '<div class="bar-val">{:,}</div></div>').format(name[:28], pct, color, val)
        return out

    def ph(title, sub):
        return ('<div class="page-header">'
                '<div><div class="ph-title">{}</div><div class="ph-sub">{}</div></div>'
                '<div class="ph-badge">{}</div></div>').format(title, sub, ds)

    def sec(title):
        return '<div class="sec-title">{}</div>'.format(title)

    def pf(note=""):
        return '<div class="page-footer"><span>Al Madina Hypermarket &nbsp;·&nbsp; {}</span><span>{} &nbsp;|&nbsp; Generated {}</span></div>'.format(
            label, note, datetime.now().strftime("%d %b %Y %H:%M"))

    # ── BUILD ─────────────────────────────────────────────────
    H = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>"
    H += "<title>Al Madina {} – {}</title>".format(label, ds)
    H += REPORT_CSS + "</head><body>"

    # COVER
    H += '<div class="cover">'
    H += '<div class="cover-badge">Al Madina Hypermarket</div>'
    H += '<div class="cover-logo">🏪</div>'
    H += '<h1>{}</h1>'.format(label)
    H += '<h2>Performance &amp; Analytics Report</h2>'
    H += '<div class="cover-stats">'
    for lbl, val in [("Period",ds),("Outlets",str(len(outlet_data))),
                      ("Total Orders",str(m["total"])),
                      ("Gross Revenue","AED {:,.0f}".format(m["gmv"])),
                      ("Payout","AED {:,.0f}".format(m["payout"]))]:
        H += '<div class="cover-stat"><div class="cs-label">{}</div><div class="cs-value">{}</div></div>'.format(lbl,val)
    H += '</div>'
    H += '<div class="cover-outlets">'
    for area,_,o in outlet_data:
        c = CLRS.get(area,"#64748b")
        H += '<div class="cpill" style="background:{}20;color:{};border-color:{}40;">{} &nbsp; AED {:,.0f}</div>'.format(c,c,c,area,o["gmv"])
    H += '</div></div>'

    # PAGE 1 — EXECUTIVE SUMMARY
    H += '<div class="page">'
    H += ph("Executive Summary","{} Outlets — {}".format(len(outlet_data),ds))
    H += sec("KEY METRICS")
    H += '<div class="kgrid kg3">'
    H += kc("Total Orders",m["total"],"{} delivered / {} cancelled".format(m["delivered"],m["cancelled"]),"or")
    H += kc("Gross Revenue (GMV)","AED {:,.0f}".format(m["gmv"]),"Total order value","bl")
    H += kc("Total Payout","AED {:,.0f}".format(m["payout"]),"After all deductions","gr")
    H += '</div><div class="kgrid kg3">'
    H += kc("Cancellation Rate",fp(m["can_rate"]),"AED {:,.0f} revenue lost".format(m["lost_gmv"]),"re")
    H += kc("Avg Order Value","AED {:,.1f}".format(m["avg_order"]),"Per delivered order","pu")
    H += kc("Commission","AED {:,.0f}".format(m["commission"]),"{:.1f}% of GMV".format(sdiv(m["commission"],m["gmv"],pct=True)),"am")
    H += '</div><div class="kgrid kg3">'
    H += kc("Avg Delivery Time",fm(m["del_time"]),"Order received → delivered","bl")
    H += kc("Pro Orders",fp(m["pro_pct"]),"of all delivered orders","or")
    H += kc("Online Payments",fp(m["online_pct"]),"vs {:.0f}% cash".format(100-m["online_pct"]),"gr")
    H += '</div>'

    H += sec("OUTLET LEAGUE TABLE — RANKED BY GMV")
    medal = ["🥇","🥈","🥉"]
    league_rows = []
    for rank,(area,_,o) in enumerate(outlet_data,1):
        m_icon = medal[rank-1] if rank<=3 else "#{}".format(rank)
        league_rows.append([m_icon, area,
            str(o["total"]), str(o["delivered"]),
            bdg(fp(o["can_rate"]),can_cls(o["can_rate"])),
            "AED {:,.0f}".format(o["gmv"]),
            "AED {:,.0f}".format(o["payout"]),
            "AED {:,.1f}".format(o["avg_order"]),
            fm(o["del_time"]),
            bdg("Fast","bg") if o["del_time"] and o["del_time"]<=35 else (bdg("Slow","br") if o["del_time"] and o["del_time"]>42 else bdg("OK","ba"))])
    H += tbl(["","Outlet","Orders","Del","Cancel %","GMV","Payout","Avg Order","Del Time","Speed"],
             league_rows, "#0f172a", r_cols={2,3,5,6,7}, c_cols={0,8,9})

    H += sec("GMV BY OUTLET")
    H += bar_chart([(a,round(o["gmv"])) for a,_,o in outlet_data])
    H += pf("Page 1 – Executive Summary")
    H += '</div>'

    # PAGE 2 — FINANCIAL & DAILY TREND
    H += '<div class="page">'
    H += ph("Financial Analysis","Revenue, Costs &amp; Payouts")
    H += sec("DAILY ORDER &amp; REVENUE TREND")
    dg = df[df["_area"].isin(areas)].groupby("_date").agg(
        total=("Order ID","count"),
        delivered=("Order status",lambda x:(x=="Delivered").sum()),
        cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
        gmv=("Subtotal","sum")).reset_index()
    dp = delivered[delivered["_area"].isin(areas)].groupby(delivered["_date"])["Payout Amount"].sum()
    dg["payout"] = dp.reindex(dg["_date"]).fillna(0).values
    dg["can_pct"] = dg.apply(lambda r:sdiv(r["cancelled"],r["total"],pct=True),axis=1)
    daily_rows = []
    for _,r in dg.iterrows():
        cb = bdg(fp(r["can_pct"]),can_cls(r["can_pct"]))
        daily_rows.append([str(r["_date"]),int(r["total"]),int(r["delivered"]),int(r["cancelled"]),
                           cb,"AED {:,.0f}".format(r["gmv"]),"AED {:,.0f}".format(r["payout"])])
    H += tbl(["Date","Total","Delivered","Cancelled","Cancel Rate","GMV (AED)","Payout (AED)"],
             daily_rows, "#1e293b", r_cols={1,2,3,5,6}, c_cols={4})

    H += '<div class="two-col" style="margin-top:12px;">'
    H += '<div>'
    H += sec("COST BREAKDOWN")
    total_cost = m["commission"]+m["op_charges"]+m["online_fee"]
    fin_rows = []
    for lbl,val in [("Commission",m["commission"]),("Op Charges",m["op_charges"]),("Online Fee",m["online_fee"])]:
        fin_rows.append([lbl,"AED {:,.0f}".format(val),"AED {:,.1f}".format(sdiv(val,m["delivered"])),fp(sdiv(val,m["gmv"],pct=True))])
    fin_rows.append(["<b>Total Deductions</b>","<b>AED {:,.0f}</b>".format(total_cost),"<b>AED {:,.1f}</b>".format(sdiv(total_cost,m["delivered"])),
                     "<b>{}</b>".format(fp(sdiv(total_cost,m["gmv"],pct=True)))])
    fin_rows.append(["<b>Net Payout</b>","<b>AED {:,.0f}</b>".format(m["payout"]),"<b>AED {:,.1f}</b>".format(sdiv(m["payout"],m["delivered"])),
                     "<b>{}</b>".format(fp(sdiv(m["payout"],m["gmv"],pct=True)))])
    H += tbl(["Item","Total","Per Order","% GMV"],fin_rows,"#1e293b",r_cols={1,2,3})
    H += '</div>'
    H += '<div>'
    H += sec("PAYMENT MIX")
    del_in_areas = delivered[delivered["_area"].isin(areas)]
    for ptype in ["Online","Cash"]:
        sdf = del_in_areas[del_in_areas["Payment type"]==ptype]
        fin_rows_p = [[ptype,str(len(sdf)),"AED {:,.0f}".format(sdf["Subtotal"].sum()),fp(sdiv(len(sdf),m["delivered"],pct=True))]]
    H += tbl(["Type","Orders","GMV","Share"],
        [["Online",str(len(del_in_areas[del_in_areas["Payment type"]=="Online"])),"AED {:,.0f}".format(del_in_areas[del_in_areas["Payment type"]=="Online"]["Subtotal"].sum()),fp(sdiv(len(del_in_areas[del_in_areas["Payment type"]=="Online"]),m["delivered"],pct=True))],
         ["Cash",str(len(del_in_areas[del_in_areas["Payment type"]=="Cash"])),"AED {:,.0f}".format(del_in_areas[del_in_areas["Payment type"]=="Cash"]["Subtotal"].sum()),fp(sdiv(len(del_in_areas[del_in_areas["Payment type"]=="Cash"]),m["delivered"],pct=True))]],
        "#3b82f6",r_cols={1,2},c_cols={3})
    H += sec("PRO vs STANDARD")
    pro = del_in_areas[del_in_areas["Is Pro Order"]=="Y"]
    std = del_in_areas[del_in_areas["Is Pro Order"]=="N"]
    H += tbl(["Type","Orders","GMV","Share"],
        [["Pro Orders",str(len(pro)),"AED {:,.0f}".format(pro["Subtotal"].sum()),fp(sdiv(len(pro),m["delivered"],pct=True))],
         ["Standard",str(len(std)),"AED {:,.0f}".format(std["Subtotal"].sum()),fp(sdiv(len(std),m["delivered"],pct=True))]],
        "#8b5cf6",r_cols={1,2},c_cols={3})
    H += sec("DAY OF WEEK")
    dow_order=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    dw = df[df["_area"].isin(areas)].groupby("_dow")["Order ID"].count().reindex(dow_order).fillna(0)
    mx_dw = dw.max()
    H += "".join('<div class="bar-row"><div class="bar-lbl">{}</div><div class="bar-track"><div class="bar-fill" style="width:{:.0f}%;background:#3b82f6;"></div></div><div class="bar-val">{}</div></div>'.format(
        day, min(cnt/mx_dw*100,100) if mx_dw>0 else 0, int(cnt)) for day,cnt in dw.items())
    H += '</div></div>'
    H += pf("Page 2 – Financial Analysis")
    H += '</div>'

    # PAGE 3 — CANCELLATION ANALYSIS
    H += '<div class="page">'
    H += ph("Cancellation Analysis","Root Cause, Item-Level &amp; Actions Required")
    sub_can = cancelled[cancelled["_area"].isin(areas)]

    H += sec("CANCELLATION SCORECARD BY OUTLET")
    can_rows = []
    for area,_,o in outlet_data:
        sc = sub_can[sub_can["_area"]==area]
        ow = sc["Cancellation owner"].value_counts()
        can_rows.append([area,str(o["cancelled"]),bdg(fp(o["can_rate"]),can_cls(o["can_rate"])),
            str(int(ow.get("Vendor",0))),str(int(ow.get("Customer",0))),str(int(ow.get("Rider",0))),
            str(int(sc["Cancellation reason"].str.contains("Item not available",na=False).sum())),
            str(int(sc["Cancellation reason"].str.contains("Fraudulent",na=False).sum())),
            "AED {:,.0f}".format(sc["Subtotal"].sum())])
    H += tbl(["Outlet","#","Rate","Vendor","Customer","Rider","Item N/A","Fraud","Lost GMV"],
             can_rows,"#ef4444",r_cols={1,3,4,5,6,7,8},c_cols={2})

    H += '<div class="two-col">'
    H += '<div>'
    H += sec("CANCELLATION BY REASON")
    rg = sub_can.groupby("Cancellation reason").agg(Count=("Order ID","count"),Lost=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False)
    H += tbl(["Reason","Count","Lost GMV"],
        [[str(r["Cancellation reason"])[:35],str(int(r["Count"])),"AED {:,.0f}".format(r["Lost"])] for _,r in rg.iterrows()],
        "#dc2626",r_cols={1,2})

    H += sec("ALERTS")
    for area,_,o in outlet_data:
        sc = sub_can[sub_can["_area"]==area]
        top_r = sc["Cancellation reason"].value_counts()
        if o["can_rate"]>40:
            H += '<div class="alert al-r">🚨 <div><b>{}</b>: CRITICAL {:.1f}% cancel rate. {} orders cancelled. Immediate action needed.</div></div>'.format(area,o["can_rate"],o["cancelled"])
        elif o["can_rate"]>10:
            reason_str = str(top_r.index[0])[:30] if len(top_r)>0 else "unknown"
            H += '<div class="alert al-a">⚠️ <div><b>{}</b>: Elevated {:.1f}% cancel rate. Top cause: <em>{}</em>.</div></div>'.format(area,o["can_rate"],reason_str)
        else:
            H += '<div class="alert al-g">✅ <div><b>{}</b>: Good {:.1f}% cancellation rate.</div></div>'.format(area,o["can_rate"])
    H += '</div>'

    H += '<div>'
    H += sec("CANCELLED ITEMS (FROM CANCELLED ORDERS)")
    cancelled_items = extract_items(sub_can["Order Items"], top=15)
    if cancelled_items:
        H += tbl(["#","Item Name","Qty in Cancelled Orders"],
            [[str(i+1),name[:32],str(qty)] for i,(name,qty) in enumerate(cancelled_items)],
            "#ef4444",r_cols={2},c_cols={0})
    else:
        H += '<p style="color:#94a3b8;font-size:11px;padding:8px 0;">No item data in cancelled orders.</p>'
    H += '</div></div>'
    H += pf("Page 3 – Cancellation Analysis")
    H += '</div>'

    # PAGE 4 — OPERATIONS
    H += '<div class="page">'
    H += ph("Operations Analysis","Delivery Timing, Hourly Patterns &amp; Efficiency")

    all_del = delivered[delivered["_area"].isin(areas)]
    H += sec("ORDER JOURNEY — AVG TIME PER STAGE (MINUTES)")
    H += '<div class="flow">'
    for lbl, sc_col, ec_col in [
        ("Order →<br>Accepted","Order received at","Accepted at"),
        ("Accepted →<br>Ready","Accepted at","Ready to pick up at"),
        ("Ready →<br>Rider","Ready to pick up at","Rider near pickup at"),
        ("Rider →<br>Transit","Rider near pickup at","In delivery at"),
        ("Transit →<br>Delivered","In delivery at","Delivered at")]:
        if sc_col in all_del.columns and ec_col in all_del.columns:
            diff=(all_del[ec_col]-all_del[sc_col]).dt.total_seconds()/60; diff=diff[diff>0]
            H += '<div class="fstep"><div class="fbox"><div class="f-lbl">{}</div><div class="f-val">{:.0f}</div><div class="f-unit">min</div></div></div>'.format(lbl,diff.mean() if len(diff)>0 else 0)
    H += '</div>'

    H += sec("TIMING BY OUTLET")
    timing_rows=[]
    for area,_,o in outlet_data:
        sd_a=delivered[delivered["_area"]==area]
        tot=(sd_a["Delivered at"]-sd_a["Order received at"]).dt.total_seconds()/60; tot=tot[tot>0]
        prep2=(sd_a["Ready to pick up at"]-sd_a["Accepted at"]).dt.total_seconds()/60; prep2=prep2[prep2>0]
        lm2=(sd_a["Delivered at"]-sd_a["In delivery at"]).dt.total_seconds()/60; lm2=lm2[lm2>0]
        avg_tot=round(tot.mean(),1) if len(tot)>0 else None
        timing_rows.append([area,
            fm(round(prep2.mean(),1) if len(prep2)>0 else None),
            fm(round(lm2.mean(),1) if len(lm2)>0 else None),
            "<b>{}</b>".format(fm(avg_tot)),
            bdg("Fast","bg") if avg_tot and avg_tot<=35 else (bdg("Slow","br") if avg_tot and avg_tot>42 else bdg("OK","ba")),
            str(len(tot))])
    H += tbl(["Outlet","Avg Prep","Avg Last Mile","Total Avg","Speed","Sample"],
             timing_rows,"#1e293b",r_cols={5},c_cols={4})

    H += '<div class="two-col" style="margin-top:12px;">'
    H += '<div>'
    H += sec("HOURLY ORDER DISTRIBUTION (ALL OUTLETS)")
    dfa = df[df["_area"].isin(areas)]
    hourly = dfa.groupby("_hour")["Order ID"].count().reset_index()
    mx_h = hourly["Order ID"].max()
    H += "".join('<div class="bar-row"><div class="bar-lbl">{:02d}:00</div><div class="bar-track"><div class="bar-fill" style="width:{:.0f}%;background:#f97316;"></div></div><div class="bar-val">{}</div></div>'.format(
        int(r["_hour"]),min(r["Order ID"]/mx_h*100,100),int(r["Order ID"])) for _,r in hourly.iterrows())
    H += '</div>'
    H += '<div>'
    H += sec("COMPLAINT ANALYSIS")
    comp_rows=[]
    for area,_,o in outlet_data:
        sd_a=delivered[delivered["_area"]==area]
        comp=sd_a[sd_a["Has Complaint?"]=="Y"]
        reasons=", ".join(comp["Complaint Reason"].dropna().unique()) if len(comp)>0 else "—"
        comp_rows.append([area,str(len(comp)),bdg(fp(o["complaint_rate"]),"br" if o["complaint_rate"]>2 else "bg"),reasons[:28]])
    H += tbl(["Outlet","#","Rate","Reasons"],comp_rows,"#8b5cf6",r_cols={1},c_cols={2})

    H += sec("TOP DELIVERED ITEMS (GROUP)")
    top_del = extract_items(delivered[delivered["_area"].isin(areas)]["Order Items"], top=10)
    if top_del:
        H += tbl(["#","Item","Qty Ordered"],
            [[str(i+1),name[:32],str(qty)] for i,(name,qty) in enumerate(top_del)],
            "#10b981",r_cols={2},c_cols={0})
    H += '</div></div>'
    H += pf("Page 4 – Operations")
    H += '</div>'

    # PER-OUTLET PAGES
    for area, fname, o in outlet_data:
        color = CLRS.get(area,"#f97316")
        sub_df=df[df["_area"]==area]; sub_d=sub_df[sub_df["Order status"]=="Delivered"]
        sub_c=sub_df[sub_df["Order status"]=="Cancelled"]

        H += '<div class="page">'
        H += '<div class="outlet-hdr" style="background:linear-gradient(135deg,{},{}cc);">'.format(color,color)
        H += '<div class="oh-name">{}</div>'.format(area)
        H += '<div class="oh-store">{}</div>'.format(fname)
        H += '<div class="oh-chips">'
        for lbl,val in [("GMV","AED {:,.0f}".format(o["gmv"])),("Orders",str(o["total"])),
                         ("Delivered",str(o["delivered"])),("Cancel Rate",fp(o["can_rate"])),
                         ("Avg Del",fm(o["del_time"])),("Payout","AED {:,.0f}".format(o["payout"]))]:
            H += '<div><div class="oh-chip-label">{}</div><div class="oh-chip-val">{}</div></div>'.format(lbl,val)
        H += '</div></div>'

        H += '<div class="kgrid kg4">'
        H += kc("Gross Revenue","AED {:,.0f}".format(o["gmv"]),"","bl")
        H += kc("Total Payout","AED {:,.0f}".format(o["payout"]),"","gr")
        H += kc("Commission","AED {:,.0f}".format(o["commission"]),fp(sdiv(o["commission"],o["gmv"],pct=True))+" of GMV","am")
        H += kc("Avg Order","AED {:,.1f}".format(o["avg_order"]),"per order","or")
        H += '</div>'

        H += '<div class="two-col">'
        H += '<div>'
        H += sec("DAILY PERFORMANCE")
        ds2=sub_df.groupby("_date").agg(
            total=("Order ID","count"),
            delivered=("Order status",lambda x:(x=="Delivered").sum()),
            cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
            gmv=("Subtotal","sum")).reset_index()
        ds2["can_pct"]=ds2.apply(lambda r:sdiv(r["cancelled"],r["total"],pct=True),axis=1)
        daily_out=[]
        for _,r in ds2.iterrows():
            cb=bdg(fp(r["can_pct"]),can_cls(r["can_pct"]))
            daily_out.append([str(r["_date"]),int(r["total"]),int(r["delivered"]),int(r["cancelled"]),cb,"AED {:,.0f}".format(r["gmv"])])
        H += tbl(["Date","Total","Del","Can","Rate","GMV"],daily_out,color,r_cols={1,2,3,5},c_cols={4})

        H += sec("FINANCIAL")
        fin_out=[]
        for lbl,val in [("GMV",o["gmv"]),("Commission",o["commission"]),
                        ("Op Charges",o["op_charges"]),("Online Fee",o["online_fee"]),("Payout",o["payout"])]:
            fin_out.append(["<b>{}</b>".format(lbl) if "Payout" in lbl else lbl,
                           "AED {:,.0f}".format(val),fp(sdiv(val,o["gmv"],pct=True))])
        H += tbl(["Item","Total","% GMV"],fin_out,color,r_cols={1,2})
        H += '</div>'

        H += '<div>'
        H += sec("TIMING")
        tot3=(sub_d["Delivered at"]-sub_d["Order received at"]).dt.total_seconds()/60; tot3=tot3[tot3>0]
        p3=(sub_d["Ready to pick up at"]-sub_d["Accepted at"]).dt.total_seconds()/60; p3=p3[p3>0]
        l3=(sub_d["Delivered at"]-sub_d["In delivery at"]).dt.total_seconds()/60; l3=l3[l3>0]
        H += tbl(["Stage","Avg","Min","Max"],
            [["Prep",fm(round(p3.mean(),1) if len(p3)>0 else None),fm(round(p3.min(),1) if len(p3)>0 else None),fm(round(p3.max(),1) if len(p3)>0 else None)],
             ["Last Mile",fm(round(l3.mean(),1) if len(l3)>0 else None),fm(round(l3.min(),1) if len(l3)>0 else None),fm(round(l3.max(),1) if len(l3)>0 else None)],
             ["<b>Total</b>","<b>{}</b>".format(fm(round(tot3.mean(),1) if len(tot3)>0 else None)),
              "<b>{}</b>".format(fm(round(tot3.min(),1) if len(tot3)>0 else None)),
              "<b>{}</b>".format(fm(round(tot3.max(),1) if len(tot3)>0 else None))]],
            color)

        if len(sub_c)>0:
            H += sec("CANCELLATIONS")
            cg=sub_c.groupby(["Cancellation owner","Cancellation reason"]).agg(
                Count=("Order ID","count"),Lost=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False)
            H += tbl(["Owner","Reason","#","Lost GMV"],
                [[str(r["Cancellation owner"]).strip(),str(r["Cancellation reason"])[:28],int(r["Count"]),"AED {:,.0f}".format(r["Lost"])] for _,r in cg.iterrows()],
                "#ef4444",r_cols={2,3})

            cancelled_items_out = extract_items(sub_c["Order Items"],top=8)
            if cancelled_items_out:
                H += sec("CANCELLED ITEMS")
                H += tbl(["Item","Qty"],[[name[:30],str(qty)] for name,qty in cancelled_items_out],"#ef4444",r_cols={1})

        H += sec("HOURLY ORDERS")
        h3=sub_df.groupby("_hour")["Order ID"].count().reset_index()
        mx3=h3["Order ID"].max()
        H += "".join('<div class="bar-row"><div class="bar-lbl">{:02d}:00</div><div class="bar-track"><div class="bar-fill" style="width:{:.0f}%;background:{};"></div></div><div class="bar-val">{}</div></div>'.format(
            int(r["_hour"]),min(r["Order ID"]/mx3*100,100) if mx3>0 else 0,color,int(r["Order ID"])) for _,r in h3.iterrows())
        H += '</div></div>'

        H += pf("Outlet – {}".format(area))
        H += '</div>'

    H += "</body></html>"
    return H.encode("utf-8")

# ─────────────────────────────────────────────────────────────
#  PASSWORD GATE
# ─────────────────────────────────────────────────────────────
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.markdown("""<div class="login-card">
        <div class="login-logo">📊</div>
        <div class="login-title">Al Madina Reports</div>
        <div class="login-sub">Enter your password to access the reporting platform</div>
    </div>""", unsafe_allow_html=True)
    col = st.columns([1,3,1])[1]
    with col:
        pwd = st.text_input("", type="password", placeholder="Enter password…", label_visibility="collapsed")
        if st.button("Sign In", use_container_width=True):
            if pwd == PASSWORD:
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    st.stop()

# ─────────────────────────────────────────────────────────────
#  MAIN APP
# ─────────────────────────────────────────────────────────────
col_brand, col_signout = st.columns([6,1])
with col_brand:
    st.markdown("""<div class="brand-bar">
        <div class="brand-icon">📊</div>
        <div><div class="brand-title">Al Madina Report Generator</div>
        <div class="brand-sub">Upload your Talabat order export — select outlets — download your report</div>
        </div></div>""", unsafe_allow_html=True)
with col_signout:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔒 Sign out"):
        st.session_state.auth = False; st.rerun()

uploaded = st.file_uploader("", type=["xlsx","xls"], label_visibility="collapsed")

if uploaded is None:
    st.markdown("""<div class="upload-box">
        <div class="upload-box-icon">📂</div>
        <div class="upload-box-title">Drop your Talabat order export here</div>
        <div class="upload-box-sub">Supports .xlsx exports &nbsp;·&nbsp; Outlets auto-detected &nbsp;·&nbsp; No install required</div>
    </div>""", unsafe_allow_html=True)
    st.markdown("""<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
        <div style="background:white;border:1px solid #e2e8f0;border-radius:12px;padding:18px;">
            <div style="font-size:20px;margin-bottom:8px;">🏢</div>
            <div style="font-size:13px;font-weight:600;color:#0f172a;margin-bottom:4px;">Management Report</div>
            <div style="font-size:11px;color:#94a3b8;">KPIs, league table, daily trend, financials — for senior management sharing</div>
        </div>
        <div style="background:white;border:1px solid #e2e8f0;border-radius:12px;padding:18px;">
            <div style="font-size:20px;margin-bottom:8px;">🏪</div>
            <div style="font-size:13px;font-weight:600;color:#0f172a;margin-bottom:4px;">Per-Outlet Report</div>
            <div style="font-size:11px;color:#94a3b8;">Individual outlet with cancellations, items, timing, hourly — for branch managers</div>
        </div>
    </div>""", unsafe_allow_html=True)
    st.stop()

with st.spinner("Reading file..."):
    try:
        df = load_data(uploaded.read())
    except Exception as e:
        st.error("Could not read file: {}".format(e)); st.stop()

if "Order status" not in df.columns:
    st.error("File doesn't look like a Talabat order report. Please check."); st.stop()

outlets_all = [a for a in sorted(df["_area"].unique()) if a != "Other"]
date_min = df["Order received at"].min()
date_max = df["Order received at"].max()
ds_str = ""
if pd.notna(date_min) and pd.notna(date_max):
    ds_str = "{} – {}".format(date_min.strftime("%d %b %Y"), date_max.strftime("%d %b %Y"))
m_group = calc_om(df[df["_area"].isin(outlets_all)])

st.markdown("""<div class="success-bar">
    <span style="font-size:18px;">✅</span>
    <div>
        <div style="font-size:13px;font-weight:600;color:#14532d;">File loaded — {orders} orders across {outlets} outlets &nbsp;·&nbsp; {period}</div>
        <div style="font-size:11px;color:#166534;margin-top:2px;">Ready to generate reports</div>
    </div></div>""".format(orders=m_group["total"],outlets=len(outlets_all),period=ds_str),
    unsafe_allow_html=True)

st.markdown("""<div class="chips">
    <div class="chip">💰 GMV <b>AED {gmv:,.0f}</b></div>
    <div class="chip">✅ <b>{d}</b> delivered</div>
    <div class="chip">❌ <b>{c}</b> cancelled (<b>{cr}</b>)</div>
    <div class="chip">💸 Payout <b>AED {p:,.0f}</b></div>
    <div class="chip">⏱ Avg delivery <b>{t}</b></div>
    <div class="chip">🔥 Cancel revenue lost <b>AED {l:,.0f}</b></div>
</div>""".format(gmv=m_group["gmv"],d=m_group["delivered"],c=m_group["cancelled"],
    cr=fp(m_group["can_rate"]),p=m_group["payout"],
    t=fm(m_group["del_time"]),l=m_group["lost_gmv"]),
    unsafe_allow_html=True)

# ── GROUP REPORTS ─────────────────────────────────────────────
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="card-title">🏢 Group Report</div>', unsafe_allow_html=True)
st.markdown('<div class="card-sub">Full report covering selected outlets — share with management or the whole team. Includes executive summary, financials, cancellation analysis, operations, and per-outlet pages.</div>', unsafe_allow_html=True)

sel = st.multiselect("Select outlets to include in group report",
    options=outlets_all, default=outlets_all)

if sel:
    for area in sel:
        o = calc_om(df[df["_area"]==area])
        color = CLRS.get(area,"#64748b")
        can_color = "#ef4444" if o["can_rate"]>20 else ("#f59e0b" if o["can_rate"]>10 else "#10b981")
        st.markdown('<div style="display:flex;align-items:center;gap:10px;padding:7px 10px;background:#f8fafc;border-radius:8px;margin-bottom:6px;border-left:3px solid {};">'.format(color), unsafe_allow_html=True)
        st.markdown('<span style="font-size:12px;font-weight:600;color:#0f172a;flex:1;">{}</span><span style="font-size:11px;color:#64748b;">AED {:,.0f} GMV &nbsp;·&nbsp; {} orders &nbsp;·&nbsp; <span style="color:{};">{:.1f}% cancel</span></span>'.format(area,o["gmv"],o["total"],can_color,o["can_rate"]), unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    if st.button("📄 Generate Group Report (HTML → Print as PDF)", use_container_width=True, type="primary"):
        with st.spinner("Building report... this may take a moment"):
            report = build_report(df, sel, "Group Performance Report – Al Madina Hypermarket")
        fname = "AlMadina_Group_Report_{}.html".format(datetime.now().strftime("%Y%m%d_%H%M"))
        st.download_button("💾 Download Report", data=report, file_name=fname,
                           mime="text/html", use_container_width=True)
        st.info("**How to get PDF:** Open downloaded file in Chrome or Safari → Press **Ctrl+P** (Windows) or **Cmd+P** (Mac) → Set destination to **Save as PDF** → Print", icon="💡")
else:
    st.warning("Please select at least one outlet.")
st.markdown('</div>', unsafe_allow_html=True)

# ── INDIVIDUAL OUTLET REPORTS ─────────────────────────────────
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="card-title">🏪 Individual Outlet Reports</div>', unsafe_allow_html=True)
st.markdown('<div class="card-sub">Dedicated report per outlet — share directly with branch managers. Includes outlet KPIs, daily performance, cancellations, cancelled items, timing, and hourly pattern.</div>', unsafe_allow_html=True)

for area in outlets_all:
    s = df[df["_area"]==area]
    if len(s)==0: continue
    o = calc_om(s)
    color = CLRS.get(area,"#64748b")
    can_color = "#ef4444" if o["can_rate"]>20 else ("#f59e0b" if o["can_rate"]>10 else "#10b981")

    with st.expander("📍  {}  &nbsp;—&nbsp;  AED {:,.0f}  ·  {} orders  ·  {:.1f}% cancel".format(
            area, o["gmv"], o["total"], o["can_rate"])):
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("GMV","AED {:,.0f}".format(o["gmv"]))
        c2.metric("Delivered",str(o["delivered"]),fp(o["del_rate"]))
        c3.metric("Cancelled",str(o["cancelled"]),fp(o["can_rate"]))
        c4.metric("Avg Delivery",fm(o["del_time"]))

        if st.button("📄 Generate {} Report".format(area),
                     use_container_width=True, type="primary", key="gen_"+area):
            with st.spinner("Building {} report...".format(area)):
                report = build_report(df,[area],"{} Outlet Report – Al Madina Hypermarket".format(area))
            fname = "AlMadina_{}_{}.html".format(area.replace(" ","_"),datetime.now().strftime("%Y%m%d_%H%M"))
            st.download_button("💾 Download {} Report".format(area), data=report,
                file_name=fname, mime="text/html", use_container_width=True, key="dl_"+area)
            st.info("Open in Chrome → Ctrl+P / Cmd+P → Save as PDF", icon="💡")

st.markdown('</div>', unsafe_allow_html=True)
