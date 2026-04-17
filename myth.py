import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import zipfile, re, xml.etree.ElementTree as ET
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Al Madina | Performance Dashboard",
                   page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
#MainMenu,footer,header{visibility:hidden;}
.stDeployButton{display:none;}
[data-testid="stSidebar"]{background:#0f1117;border-right:1px solid #1e2130;}
[data-testid="stSidebar"] *{color:#e2e8f0 !important;}
[data-testid="stSidebar"] hr{border-color:#1e2130;}
.main .block-container{padding:0 2rem 2rem;max-width:1400px;}
.top-bar{background:linear-gradient(135deg,#0f1117 0%,#1a1f2e 100%);border-bottom:1px solid #1e2130;padding:16px 28px;margin:-1rem -2rem 2rem;display:flex;align-items:center;justify-content:space-between;}
.logo-circle{width:38px;height:38px;border-radius:10px;background:linear-gradient(135deg,#f97316,#dc2626);display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:700;color:white;}
.top-bar-brand{display:flex;align-items:center;gap:12px;}
.top-bar-brand h1{font-size:17px;font-weight:600;color:#f1f5f9;margin:0;}
.top-bar-brand span{font-size:11px;color:#64748b;}
.section-label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#94a3b8;margin:1.5rem 0 .6rem;padding-bottom:.5rem;border-bottom:1px solid #e2e8f0;}
.kpi-card{background:white;border:1px solid #e2e8f0;border-radius:12px;padding:14px 18px;position:relative;overflow:hidden;}
.kpi-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.kpi-card.orange::before{background:linear-gradient(90deg,#f97316,#fb923c);}
.kpi-card.green::before{background:linear-gradient(90deg,#10b981,#34d399);}
.kpi-card.red::before{background:linear-gradient(90deg,#ef4444,#f87171);}
.kpi-card.blue::before{background:linear-gradient(90deg,#3b82f6,#60a5fa);}
.kpi-card.purple::before{background:linear-gradient(90deg,#8b5cf6,#a78bfa);}
.kpi-card.amber::before{background:linear-gradient(90deg,#f59e0b,#fbbf24);}
.kpi-label{font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;}
.kpi-value{font-size:24px;font-weight:600;color:#0f172a;line-height:1;}
.kpi-sub{font-size:11px;color:#64748b;margin-top:3px;}
.outlet-card{background:white;border:1px solid #e2e8f0;border-radius:14px;padding:18px;margin-bottom:10px;}
.outlet-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;}
.outlet-name{font-size:14px;font-weight:600;color:#0f172a;}
.outlet-area{font-size:11px;color:#94a3b8;margin-top:2px;}
.outlet-badge{font-size:11px;font-weight:600;padding:3px 10px;border-radius:20px;white-space:nowrap;display:inline-block;}
.badge-good{background:#d1fae5;color:#065f46;}
.badge-warn{background:#fef3c7;color:#92400e;}
.badge-bad{background:#fee2e2;color:#991b1b;}
.prog-wrap{margin:3px 0 7px;}
.prog-label{display:flex;justify-content:space-between;font-size:11px;color:#64748b;margin-bottom:3px;}
.prog-track{height:5px;background:#f1f5f9;border-radius:3px;overflow:hidden;}
.prog-fill{height:100%;border-radius:3px;}
.compare-row{display:flex;gap:6px;flex-wrap:wrap;margin:6px 0;}
.compare-chip{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:5px 10px;font-size:11px;color:#475569;display:flex;gap:5px;align-items:center;}
.chip-val{font-weight:600;color:#0f172a;}
.alert-box{border-radius:10px;padding:11px 14px;margin:7px 0;font-size:13px;display:flex;gap:10px;align-items:flex-start;}
.alert-red{background:#fff1f2;border:1px solid #fecdd3;color:#9f1239;}
.alert-amber{background:#fffbeb;border:1px solid #fed7aa;color:#92400e;}
.alert-green{background:#f0fdf4;border:1px solid #bbf7d0;color:#14532d;}
.alert-blue{background:#eff6ff;border:1px solid #bfdbfe;color:#1e3a8a;}
.alert-icon{font-size:15px;flex-shrink:0;}
.stTabs [data-baseweb="tab-list"]{gap:4px;border-bottom:2px solid #e2e8f0;}
.stTabs [data-baseweb="tab"]{background:transparent;border:none;border-radius:8px 8px 0 0;padding:8px 16px;font-size:13px;font-weight:500;color:#64748b;}
.stTabs [aria-selected="true"]{background:white;color:#f97316 !important;border-bottom:2px solid #f97316;}
.dl-card{background:white;border:1px solid #e2e8f0;border-radius:16px;padding:24px;height:100%;}
.dl-card-header{display:flex;align-items:center;gap:12px;margin-bottom:14px;}
.dl-icon{width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;}
.dl-icon.orange{background:#fff7ed;}
.dl-icon.blue{background:#eff6ff;}
.dl-icon.green{background:#f0fdf4;}
.dl-title{font-size:15px;font-weight:600;color:#0f172a;}
.dl-desc{font-size:12px;color:#64748b;margin-bottom:14px;line-height:1.6;}
.dl-features{list-style:none;padding:0;margin:0 0 16px;}
.dl-features li{font-size:12px;color:#64748b;padding:3px 0;display:flex;align-items:center;gap:6px;}
.dl-features li::before{content:"✓";color:#10b981;font-weight:700;}
.outlet-select-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px;}
.outlet-select-chip{background:white;border:2px solid #e2e8f0;border-radius:12px;padding:12px 14px;cursor:pointer;transition:all .15s;}
.outlet-select-chip.selected{border-color:#f97316;background:#fff7ed;}
.outlet-select-chip-name{font-size:13px;font-weight:600;color:#0f172a;}
.outlet-select-chip-gmv{font-size:11px;color:#94a3b8;margin-top:2px;}
</style>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════
NUMERIC_COLS = ["Subtotal","Packaging charges","Minimum order value fee","Vendor Refunds",
    "plugins.reports.order_details_report.customer_fee_total","Tax Charge","Online Payment Fee",
    "Discount Funded by you","Voucher Funded by you","Commission","Operational Charges",
    "Ads Fee","Marketing Fees","Avoidable cancellation fee","Estimated earnings",
    "Cash amount already collected by you","Amount owed back to Talabat","Payout Amount",
    "Talabat-Funded Discount","Talabat-Funded Voucher","Total Discount","Total Voucher","Tax Amount"]
DATE_COLS = ["Order received at","Accepted at","Estimated ready to pick up time",
    "Ready to pick up at","Rider near pickup at","In delivery at",
    "Estimated delivery time","Delivered at","Cancelled at"]
OUTLET_COLORS = {"Liwan":"#f97316","Dubai Investment Park":"#3b82f6","Oud Metha":"#8b5cf6",
    "Naif":"#ef4444","Al Muteena":"#10b981","Al Hamriya":"#f59e0b"}

# ══════════════════════════════════════════════════════════════
#  STDLIB XLSX WRITER  (zero external deps)
# ══════════════════════════════════════════════════════════════
def _write_xlsx(sheets_data):
    def col_letter(n):
        s = ""
        while n >= 0:
            s = chr(65 + n % 26) + s; n = n // 26 - 1
        return s
    def esc(v):
        return str(v).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")
    st_list = []; st_map = {}
    def ss(val):
        v = "" if val is None else str(val)
        if v not in st_map: st_map[v] = len(st_list); st_list.append(v)
        return st_map[v]
    for sd in sheets_data:
        for h in sd["headers"]: ss(h)
        for row in sd["rows"]:
            for cell in row: ss(cell)

    STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="5">'
        '<font><sz val="9"/><name val="Calibri"/></font>'
        '<font><b/><sz val="9"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>'
        '<font><b/><sz val="13"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>'
        '<font><b/><sz val="9"/><name val="Calibri"/><color rgb="FFF97316"/></font>'
        '<font><sz val="8"/><name val="Calibri"/><color rgb="FF94A3B8"/></font>'
        '</fonts>'
        '<fills count="11">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF0F1117"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF97316"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF10B981"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFEF4444"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF3B82F6"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF8B5CF6"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF59E0B"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFE4E6"/></patternFill></fill>'
        '</fills>'
        '<borders count="2">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color rgb="FFE2E8F0"/></left>'
        '<right style="thin"><color rgb="FFE2E8F0"/></right>'
        '<top style="thin"><color rgb="FFE2E8F0"/></top>'
        '<bottom style="thin"><color rgb="FFE2E8F0"/></bottom><diagonal/></border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="12">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"><alignment vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0"><alignment vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0"><alignment vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="1" fillId="6" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="1" fillId="7" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="1" fillId="8" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="1" fillId="9" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="10" borderId="1" xfId="0"><alignment vertical="center"/></xf>'
        '</cellXfs>'
        '</styleSheet>')
    # XF map by color name
    HDR_XF = {"orange":1,"dark":2,"green":5,"red":6,"blue":7,"purple":8,"amber":9}

    def build_sheet(sd):
        headers   = sd["headers"]; rows = sd["rows"]
        col_widths= sd.get("col_widths",[])
        hxf       = HDR_XF.get(sd.get("hdr_color","orange"),1)
        sections  = {s[0]:s for s in sd.get("sections",[])}
        lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
        if col_widths:
            lines.append("<cols>")
            for ci,w in enumerate(col_widths):
                lines.append('<col min="{}" max="{}" width="{}" customWidth="1"/>'.format(ci+1,ci+1,w))
            lines.append("</cols>")
        lines.append("<sheetData>")
        er = 1
        # title row if present
        if sd.get("title"):
            t = sd["title"]
            txf = HDR_XF.get(sd.get("title_color","dark"),2)
            ncols = max(len(headers), 8)
            lines.append('<row r="{}" ht="28" customHeight="1">'.format(er))
            lines.append('<c r="{}{}\" t="s" s="{}"><v>{}</v></c>'.format(col_letter(0),er,txf,ss(t)))
            for ci in range(1,ncols):
                lines.append('<c r="{}{}" t="s" s="{}"><v>{}</v></c>'.format(col_letter(ci),er,txf,ss("")))
            lines.append("</row>"); er += 1
            # subtitle
            if sd.get("subtitle"):
                lines.append('<row r="{}" ht="14" customHeight="1">'.format(er))
                lines.append('<c r="{}{}" t="s" s="10"><v>{}</v></c>'.format(col_letter(0),er,ss(sd["subtitle"])))
                for ci in range(1,ncols):
                    lines.append('<c r="{}{}" t="s" s="10"><v>{}</v></c>'.format(col_letter(ci),er,ss("")))
                lines.append("</row>"); er += 2
        # header row
        lines.append('<row r="{}" ht="22" customHeight="1">'.format(er))
        for ci,h in enumerate(headers):
            lines.append('<c r="{}{}" t="s" s="{}"><v>{}</v></c>'.format(col_letter(ci),er,hxf,ss(h)))
        lines.append("</row>"); er += 1
        # data rows
        for ri,row in enumerate(rows):
            if ri in sections:
                _,stitle,scolor = sections[ri]
                sxf = HDR_XF.get(scolor,1)
                lines.append('<row r="{}" ht="20" customHeight="1">'.format(er))
                for ci in range(max(len(headers),8)):
                    lines.append('<c r="{}{}" t="s" s="{}"><v>{}</v></c>'.format(
                        col_letter(ci),er,sxf,ss("  "+stitle if ci==0 else "")))
                lines.append("</row>"); er += 1
            bg = 3 if ri%2==0 else 0
            lines.append('<row r="{}">'.format(er))
            for ci,cell in enumerate(row):
                lines.append('<c r="{}{}" t="s" s="{}"><v>{}</v></c>'.format(
                    col_letter(ci),er,bg,ss(cell)))
            lines.append("</row>"); er += 1
        lines.append("</sheetData></worksheet>")
        return "\n".join(lines)

    n = len(sheets_data)
    overrides = "".join('<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i+1) for i in range(n))
    CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          +overrides+
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
          '</Types>')
    RELS_ROOT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')
    sheets_el = "".join('<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(esc(sd["name"]),i+1,i+1) for i,sd in enumerate(sheets_data))
    WORKBOOK = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets>'+sheets_el+'</sheets></workbook>')
    rels = "".join('<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>'.format(i+1,i+1) for i in range(n))
    rels += '<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'.format(n+1)
    rels += '<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'.format(n+2)
    WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rels+'</Relationships>')
    SHARED = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{c}" uniqueCount="{c}">'.format(c=len(st_list))
              +"".join('<si><t xml:space="preserve">{}</t></si>'.format(esc(s)) for s in st_list)
              +'</sst>')
    buf = BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",CT)
        z.writestr("_rels/.rels",RELS_ROOT)
        z.writestr("xl/workbook.xml",WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels",WB_RELS)
        z.writestr("xl/styles.xml",STYLES)
        z.writestr("xl/sharedStrings.xml",SHARED)
        for i,sd in enumerate(sheets_data):
            z.writestr("xl/worksheets/sheet{}.xml".format(i+1),build_sheet(sd))
    buf.seek(0); return buf.read()

# ══════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════
def get_area(name):
    n = str(name).lower()
    if "liwan" in n: return "Liwan"
    if "dip" in n or "investment park" in n: return "Dubai Investment Park"
    if "oud metha" in n: return "Oud Metha"
    if "naif" in n: return "Naif"
    if "muteena" in n: return "Al Muteena"
    if "hamriya" in n: return "Al Hamriya"
    return "Other"

def safe_div(a, b, pct=False):
    if b == 0: return 0
    return (a/b*100) if pct else a/b

def outlet_metrics(df_sub):
    d = df_sub[df_sub["Order status"] == "Delivered"]
    c = df_sub[df_sub["Order status"] == "Cancelled"]
    total = len(df_sub); gmv = d["Subtotal"].sum()
    diff  = (d["Delivered at"]-d["Order received at"]).dt.total_seconds()/60; diff=diff[diff>0]
    prep  = (d["Ready to pick up at"]-d["Accepted at"]).dt.total_seconds()/60; prep=prep[prep>0]
    lm    = (d["Delivered at"]-d["In delivery at"]).dt.total_seconds()/60; lm=lm[lm>0]
    return {"total":total,"delivered":len(d),"cancelled":len(c),
        "can_rate":safe_div(len(c),total,pct=True),"del_rate":safe_div(len(d),total,pct=True),
        "gmv":gmv,"payout":d["Payout Amount"].sum(),"commission":d["Commission"].sum(),
        "op_charges":d["Operational Charges"].sum(),"online_fee":d["Online Payment Fee"].sum(),
        "avg_order":safe_div(gmv,len(d)),"del_time":round(diff.mean(),1) if len(diff)>0 else None,
        "prep_time":round(prep.mean(),1) if len(prep)>0 else None,
        "last_mile":round(lm.mean(),1) if len(lm)>0 else None,
        "pro_pct":safe_div(len(d[d["Is Pro Order"]=="Y"]),len(d),pct=True),
        "online_pct":safe_div(len(d[d["Payment type"]=="Online"]),len(d),pct=True),
        "complaints":len(d[d["Has Complaint?"]=="Y"]),
        "complaint_rate":safe_div(len(d[d["Has Complaint?"]=="Y"]),len(d),pct=True),
        "lost_gmv":c["Subtotal"].sum()}

def can_badge(r):
    if r<=8: return "badge-good","Low"
    if r<=20: return "badge-warn","Medium"
    return "badge-bad","High"

def del_badge(m):
    if m is None: return "badge-warn","N/A"
    if m<=35: return "badge-good","Fast"
    if m<=45: return "badge-warn","Average"
    return "badge-bad","Slow"

def kpi(label, value, sub="", color="orange"):
    st.markdown(('<div class="kpi-card {c}"><div class="kpi-label">{l}</div>'
                 '<div class="kpi-value">{v}</div>{s}</div>').format(
                     c=color,l=label,v=value,
                     s='<div class="kpi-sub">'+sub+'</div>' if sub else ""),
                unsafe_allow_html=True)

def progress_bar(label, value, max_val, color="#f97316", suffix=""):
    pct = min(safe_div(value,max_val,pct=True),100)
    st.markdown(('<div class="prog-wrap"><div class="prog-label"><span>{l}</span><span>{v}{s}</span></div>'
                 '<div class="prog-track"><div class="prog-fill" style="width:{p:.1f}%;background:{c};"></div></div></div>').format(
                     l=label,v=value,s=suffix,p=pct,c=color),unsafe_allow_html=True)

def alert(msg, type_="blue", icon="ℹ"):
    st.markdown('<div class="alert-box alert-{t}"><span class="alert-icon">{i}</span><span>{m}</span></div>'.format(
        t=type_,i=icon,m=msg),unsafe_allow_html=True)

def fmt(v, prefix="AED ", dec=0):
    if v is None: return "N/A"
    return "{}{:,.{}f}".format(prefix, v, dec)

def _style_df(df, heat_col=None, heat_col2=None, reverse=False):
    def bar(val, col_vals, color):
        try:
            mn,mx = col_vals.min(),col_vals.max()
            pct = 50 if mx==mn else (val-mn)/(mx-mn)*100
            if reverse: pct = 100-pct
            return "background:linear-gradient(90deg,{c} {p:.0f}%,transparent {p:.0f}%);color:#0f172a;".format(c=color,p=pct)
        except: return ""
    styles = pd.DataFrame("",index=df.index,columns=df.columns)
    if heat_col and heat_col in df.columns:
        cv = pd.to_numeric(df[heat_col],errors="coerce").fillna(0)
        styles[heat_col] = cv.apply(lambda v: bar(v,cv,"rgba(249,115,22,0.2)"))
    if heat_col2 and heat_col2 in df.columns:
        cv2 = pd.to_numeric(df[heat_col2],errors="coerce").fillna(0)
        styles[heat_col2] = cv2.apply(lambda v: bar(v,cv2,"rgba(239,68,68,0.2)"))
    return df.style.apply(lambda _: styles, axis=None)

# ══════════════════════════════════════════════════════════════
#  DATA LOADER
# ══════════════════════════════════════════════════════════════
def _parse_xlsx_stdlib(file_bytes):
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
        names = [i.filename for i in z.infolist()]
        sheet_xml = z.read("xl/worksheets/sheet1.xml")
        shared = []
        if "xl/sharedStrings.xml" in names:
            ss_tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_tree.iter("{"+ns+"}si"):
                shared.append("".join(t.text or "" for t in si.iter("{"+ns+"}t")))
    def col_idx(ref):
        col = "".join(filter(str.isalpha,ref)); idx=0
        for c in col: idx=idx*26+(ord(c.upper())-ord("A")+1)
        return idx-1
    tree = ET.fromstring(sheet_xml)
    rows_out=[]; max_col=0
    for row_el in tree.iter("{"+ns+"}row"):
        cells={}
        for cell in row_el.iter("{"+ns+"}c"):
            ref=cell.get("r","A1"); ci=col_idx(ref); t=cell.get("t","n")
            v_el=cell.find("{"+ns+"}v"); is_el=cell.find("{"+ns+"}is")
            if is_el is not None:
                t_el=is_el.find("{"+ns+"}t"); val=t_el.text if t_el is not None else ""
            elif v_el is not None:
                rv=v_el.text or ""
                if t=="s": val=shared[int(rv)] if shared else rv
                elif t=="b": val=bool(int(rv))
                else:
                    try: val=float(rv) if "." in rv else int(rv)
                    except: val=rv
            else: val=None
            cells[ci]=val; max_col=max(max_col,ci)
        rows_out.append([cells.get(i) for i in range(max_col+1)])
    return pd.DataFrame(rows_out) if rows_out else pd.DataFrame()

def _fix_and_read_openpyxl(file_bytes):
    VALID_VERT={b"top",b"center",b"bottom",b"justify",b"distributed"}
    VALID_HORIZ={b"general",b"left",b"center",b"right",b"fill",b"justify",b"centerContinuous",b"distributed"}
    def fix_attr(xml,attr,valid,fallback):
        pat=re.compile(rb"("+attr+rb'=")([^"]*?)(")')
        return pat.sub(lambda m: m.group(0) if m.group(2) in valid else m.group(1)+fallback+m.group(3),xml)
    buf_out=BytesIO()
    with zipfile.ZipFile(BytesIO(file_bytes),"r") as zin, zipfile.ZipFile(buf_out,"w",zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename=="xl/styles.xml":
                data=fix_attr(data,b"vertical",VALID_VERT,b"bottom")
                data=fix_attr(data,b"horizontal",VALID_HORIZ,b"general")
            zout.writestr(item.filename,data)
    buf_out.seek(0)
    return pd.read_excel(buf_out,engine="openpyxl",header=0)

def _read_excel_robust(file_bytes):
    # 1. stdlib XML (always works)
    try:
        with zipfile.ZipFile(BytesIO(file_bytes)) as z:
            if "xl/worksheets/sheet1.xml" in [i.filename for i in z.infolist()]:
                df=_parse_xlsx_stdlib(file_bytes)
                if len(df)>1: return df
    except Exception: pass
    # 2. calamine
    try:
        import python_calamine
        return pd.read_excel(BytesIO(file_bytes),engine="calamine",header=0)
    except Exception: pass
    # 3. openpyxl repaired
    try: return _fix_and_read_openpyxl(file_bytes)
    except Exception: pass
    # 4. openpyxl raw
    try: return pd.read_excel(BytesIO(file_bytes),engine="openpyxl",header=0)
    except Exception: pass
    # 5. xlrd
    try: return pd.read_excel(BytesIO(file_bytes),engine="xlrd",header=0)
    except Exception as e:
        raise RuntimeError("Could not read file. Please re-save as .xlsx from Excel. Detail: "+str(e))

@st.cache_data
def load_data(file_bytes):
    df_raw = _read_excel_robust(file_bytes)
    header_row = None
    for i in range(min(5,len(df_raw))):
        if "Order status" in [str(v) for v in df_raw.iloc[i].values]:
            header_row = i; break
    if header_row is None: header_row = 0
    df = df_raw.iloc[header_row+1:].copy()
    df.columns = [str(c) for c in df_raw.iloc[header_row].tolist()]
    df.reset_index(drop=True,inplace=True)
    df = df.dropna(how="all").reset_index(drop=True)
    for col in NUMERIC_COLS:
        if col in df.columns: df[col]=pd.to_numeric(df[col],errors="coerce").fillna(0)
    for col in DATE_COLS:
        if col in df.columns: df[col]=pd.to_datetime(df[col],errors="coerce")
    if "Restaurant name" not in df.columns:
        raise RuntimeError("Column 'Restaurant name' not found. Is this a Talabat order report?")
    df["_area"]=df["Restaurant name"].apply(get_area)
    df["_date"]=df["Order received at"].dt.date
    df["_hour"]=df["Order received at"].dt.hour
    df["_dow"]=df["Order received at"].dt.day_name()
    return df

# ══════════════════════════════════════════════════════════════
#  REPORT BUILDERS
# ══════════════════════════════════════════════════════════════
def _outlet_sheets(df, areas, include_raw=True):
    """Build list of sheet dicts for _write_xlsx covering all requested areas."""
    delivered = df[df["Order status"]=="Delivered"]
    cancelled = df[df["Order status"]=="Cancelled"]
    date_min  = df["Order received at"].min()
    date_max  = df["Order received at"].max()
    ds = ""
    if pd.notna(date_min) and pd.notna(date_max):
        ds = date_min.strftime("%d %b %Y")+" - "+date_max.strftime("%d %b %Y")

    outlet_data = []
    for area in sorted(areas):
        s = df[df["_area"]==area]
        if len(s)>0: outlet_data.append((area,s["Restaurant name"].iloc[0],outlet_metrics(s)))
    outlet_data.sort(key=lambda x:x[2]["gmv"],reverse=True)

    COLOR_MAP = {"Liwan":"orange","Dubai Investment Park":"blue","Oud Metha":"purple",
                 "Naif":"red","Al Muteena":"green","Al Hamriya":"amber"}

    def f(v,dec=0): return "{:,.{}f}".format(v,dec) if v is not None else "N/A"
    def fp(v): return "{:.1f}%".format(v)
    def fm(v): return "{:.0f} min".format(v) if v else "N/A"

    sheets = []

    # ── Group Summary ──────────────────────────────────────────
    if len(areas) > 1:
        m = outlet_metrics(df[df["_area"].isin(areas)])
        # League table
        lt_rows = []
        for rank,(area,fname,om) in enumerate(outlet_data,1):
            lt_rows.append([str(rank),area,fname,str(om["total"]),str(om["delivered"]),
                str(om["cancelled"]),fp(om["can_rate"]),
                f(om["gmv"]),f(om["payout"]),f(om["commission"]),f(om["avg_order"],1),
                fm(om["del_time"]),str(om["complaints"]),fp(om["pro_pct"])])
        # Daily
        dg=df[df["_area"].isin(areas)].groupby("_date").agg(
            total=("Order ID","count"),
            delivered=("Order status",lambda x:(x=="Delivered").sum()),
            cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
            gmv=("Subtotal","sum")).reset_index()
        dg["can_pct"]=dg.apply(lambda r:safe_div(r["cancelled"],r["total"],pct=True),axis=1)
        daily_rows=[[str(r["_date"]),str(int(r["total"])),str(int(r["delivered"])),
            str(int(r["cancelled"])),fp(r["can_pct"]),f(r["gmv"])] for _,r in dg.iterrows()]
        # KPI summary rows
        kpi_rows=[
            ["Total Orders",str(m["total"]),"Delivered",str(m["delivered"])+" ("+fp(m["del_rate"])+")",
             "Cancelled",str(m["cancelled"])+" ("+fp(m["can_rate"])+")"],
            ["Gross Revenue","AED "+f(m["gmv"]),"Total Payout","AED "+f(m["payout"]),
             "Commission","AED "+f(m["commission"])],
            ["Avg Order","AED "+f(m["avg_order"],1),"Avg Delivery",fm(m["del_time"]),
             "Complaints",str(m["complaints"])+" ("+fp(m["complaint_rate"])+")"],
            ["Online Pay %",fp(m["online_pct"]),"Pro Orders %",fp(m["pro_pct"]),
             "Lost to Cancels","AED "+f(m["lost_gmv"])],
        ]
        all_rows = (kpi_rows + [["","","","","",""]] + lt_rows +
                    [["","","","","",""]] + daily_rows)
        sections_map = [
            (0,"GROUP KEY METRICS","dark"),
            (len(kpi_rows)+1,"OUTLET LEAGUE TABLE (sorted by GMV)","orange"),
            (len(kpi_rows)+1+len(lt_rows)+1,"DAILY GROUP TREND","blue"),
        ]
        sheets.append({"name":"Group Summary",
            "title":"AL MADINA HYPERMARKET  -  GROUP PERFORMANCE REPORT",
            "subtitle":"Period: "+ds+"   |   "+str(len(outlet_data))+" Outlets   |   Generated: "+datetime.now().strftime("%d %b %Y %H:%M"),
            "title_color":"dark",
            "headers":["Metric / Rank","Value / Outlet","Metric / Area","Value / Orders",
                       "Metric / Del","Value / Can"],
            "rows":all_rows,"col_widths":[24,20,22,12,12,12],
            "hdr_color":"dark","sections":sections_map})

    # ── Cancellation Analysis ──────────────────────────────────
    sub_can = cancelled[cancelled["_area"].isin(areas)]
    can_rows=[]
    for area,fname,om in outlet_data:
        sc=sub_can[sub_can["_area"]==area]
        ow=sc["Cancellation owner"].value_counts()
        can_rows.append([area,str(len(sc)),fp(om["can_rate"]),
            str(int(ow.get("Vendor",0))),str(int(ow.get("Customer",0))),str(int(ow.get("Rider",0))),
            str(int(sc["Cancellation reason"].str.contains("Item not available",na=False).sum())),
            str(int(sc["Cancellation reason"].str.contains("Fraudulent",na=False).sum())),
            "AED "+f(sc["Subtotal"].sum())])
    reason_rows=[]
    for area,_,_ in outlet_data:
        sc=sub_can[sub_can["_area"]==area]
        if len(sc)==0: continue
        for _, r in sc.groupby(["Cancellation reason","Cancellation owner"]).agg(
                Count=("Order ID","count"),Lost=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False).iterrows():
            reason_rows.append([area,str(r["Cancellation owner"]).strip(),
                str(r["Cancellation reason"]),str(int(r["Count"])),"AED "+f(r["Lost"])])
    # Item-level cancellation
    from collections import Counter
    item_can_by_outlet = {}
    for area,_,_ in outlet_data:
        sc=sub_can[sub_can["_area"]==area]
        ctr=Counter()
        for items_str in sc["Order Items"].dropna():
            for item in str(items_str).split(","):
                item=item.strip()
                parts=item.split(" ",1)
                qty=int(parts[0]) if len(parts)==2 and parts[0].isdigit() else 1
                name=parts[1].strip() if len(parts)==2 and parts[0].isdigit() else item
                if name and len(name)>3 and not name.replace(".","").replace("g","").replace("ml","").isdigit():
                    ctr[name]+=qty
        item_can_by_outlet[area]=ctr
    item_rows=[]
    for area,_,_ in outlet_data:
        for item,qty in item_can_by_outlet[area].most_common(10):
            item_rows.append([area,item,str(qty)])
    all_can_rows = (can_rows+[["","","","","","","","",""]]+
                    reason_rows+[["","","","",""]]+item_rows)
    can_secs = [
        (0,"OUTLET CANCELLATION SCORECARD","red"),
        (len(can_rows)+1,"CANCELLATION BY REASON & OWNER","red"),
        (len(can_rows)+1+len(reason_rows)+1,"ITEMS CANCELLED (from cancelled orders)","red"),
    ]
    sheets.append({"name":"Cancellations",
        "title":"CANCELLATION ANALYSIS","title_color":"red",
        "headers":["Outlet","Cancelled","Rate %","Vendor","Customer","Rider",
                   "Item N/A","Fraudulent","Lost GMV"],
        "rows":all_can_rows,"col_widths":[22,10,10,10,10,10,12,12,16],
        "hdr_color":"red","sections":can_secs})

    # ── Per-outlet sheets ──────────────────────────────────────
    for area,fname,om in outlet_data:
        color = COLOR_MAP.get(area,"orange")
        sub_df=df[df["_area"]==area]; sub_d=sub_df[sub_df["Order status"]=="Delivered"]
        sub_c=sub_df[sub_df["Order status"]=="Cancelled"]

        # Daily
        ds2=sub_df.groupby("_date").agg(total=("Order ID","count"),
            delivered=("Order status",lambda x:(x=="Delivered").sum()),
            cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
            gmv=("Subtotal","sum")).reset_index()
        dp2=sub_d.groupby(sub_d["_date"])["Payout Amount"].sum().reset_index()
        dp2.columns=["_date","payout"]; ds2=ds2.merge(dp2,on="_date",how="left").fillna(0)
        ds2["can_pct"]=ds2.apply(lambda r:safe_div(r["cancelled"],r["total"],pct=True),axis=1)
        daily_r=[[str(r["_date"]),str(int(r["total"])),str(int(r["delivered"])),
            str(int(r["cancelled"])),fp(r["can_pct"]),f(r["gmv"]),"AED "+f(r["payout"])]
            for _,r in ds2.iterrows()]
        # Hourly
        hourly=sub_df.groupby("_hour")["Order ID"].count().reset_index()
        hourly_r=[["{:02d}:00".format(int(r["_hour"])),str(int(r["Order ID"]))] for _,r in hourly.iterrows()]
        # Financial
        fin_items=[("GMV",om["gmv"]),("Commission",om["commission"]),
            ("Op Charges",om["op_charges"]),("Online Fee",om["online_fee"]),
            ("Total Payout",om["payout"])]
        fin_r=[[l,"AED "+f(v,2),"AED "+f(safe_div(v,om["delivered"]),2),fp(safe_div(v,om["gmv"],pct=True))]
               for l,v in fin_items]
        # Cancellations
        can_r2=[]
        if len(sub_c)>0:
            cd=sub_c.groupby(["Cancellation owner","Cancellation reason"]).agg(
                Count=("Order ID","count"),Lost=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False)
            can_r2=[[str(r["Cancellation owner"]).strip(),str(r["Cancellation reason"]),
                str(int(r["Count"])),"AED "+f(r["Lost"])] for _,r in cd.iterrows()]
        # Item-level cancel
        item_r2=[[item,str(qty)] for item,qty in item_can_by_outlet.get(area,Counter()).most_common(15)]
        # Timing
        timing_pairs=[
            ("Received -> Accepted","Order received at","Accepted at"),
            ("Accepted -> Ready","Accepted at","Ready to pick up at"),
            ("Rider Wait","Ready to pick up at","Rider near pickup at"),
            ("In Delivery -> Delivered","In delivery at","Delivered at"),
            ("TOTAL: Received -> Delivered","Order received at","Delivered at"),
        ]
        timing_r=[]
        for label,sc_col,ec_col in timing_pairs:
            if sc_col in sub_d.columns and ec_col in sub_d.columns:
                diff=(sub_d[ec_col]-sub_d[sc_col]).dt.total_seconds()/60; diff=diff[diff>0]
                if len(diff)>0:
                    timing_r.append([label,"{:.1f}".format(diff.mean()),
                        "{:.1f}".format(diff.min()),"{:.1f}".format(diff.max()),
                        "{:.1f}".format(diff.median())])
        # Payment
        pay_r=[]
        for ptype in ["Online","Cash"]:
            sd_p=sub_d[sub_d["Payment type"]==ptype]
            if len(sd_p)>0:
                pay_r.append([ptype,str(len(sd_p)),f(sd_p["Subtotal"].sum()),
                    fp(safe_div(len(sd_p),len(sub_d),pct=True))])

        # Assemble all rows
        kpi_r=[
            ["Total Orders",str(om["total"]),"Delivered",str(om["delivered"])+" ("+fp(om["del_rate"])+")",
             "Cancelled",str(om["cancelled"])+" ("+fp(om["can_rate"])+")"],
            ["GMV","AED "+f(om["gmv"]),"Payout","AED "+f(om["payout"]),"Commission","AED "+f(om["commission"])],
            ["Avg Order","AED "+f(om["avg_order"],1),"Avg Del",fm(om["del_time"]),"Avg Prep",fm(om["prep_time"])],
            ["Pro Orders %",fp(om["pro_pct"]),"Online Pay %",fp(om["online_pct"]),
             "Complaints",str(om["complaints"])+" ("+fp(om["complaint_rate"])+")"],
        ]
        blank6=["","","","","",""]
        blank7=["","","","","","",""]
        all_rows=(kpi_r+[blank6]+daily_r+[blank7]+
            (can_r2 if can_r2 else [])+
            ([blank6] if can_r2 else [])+
            item_r2+[["",""]]+
            timing_r+[blank6]+fin_r+[blank6]+hourly_r+[pay_r[0] if pay_r else []]+
            (pay_r[1:] if len(pay_r)>1 else []))

        offset = len(kpi_r)+1
        secs = [(0,"KEY METRICS",color)]
        secs.append((offset,"DAILY PERFORMANCE",color))
        offset2 = offset+len(daily_r)+1
        if can_r2:
            secs.append((offset2,"CANCELLATIONS BY REASON",color))
            offset2 += len(can_r2)+1
        if item_r2:
            secs.append((offset2,"CANCELLED ITEMS","red"))
            offset2 += len(item_r2)+1
        secs.append((offset2,"DELIVERY TIMING (minutes)",color))
        secs.append((offset2+len(timing_r)+1,"FINANCIAL BREAKDOWN",color))
        secs.append((offset2+len(timing_r)+1+len(fin_r)+1,"HOURLY ORDERS",color))

        sheets.append({"name":area[:28],
            "title":"OUTLET REPORT  -  "+area.upper(),
            "subtitle":fname+"  |  "+ds,
            "title_color":color,
            "headers":["Metric / Item","Value","Metric / Item","Value",
                       "Metric / Item","Value / Extra"],
            "rows":all_rows,"col_widths":[28,18,28,18,22,18],
            "hdr_color":color,"sections":secs})

    # ── Raw Data ───────────────────────────────────────────────
    if include_raw:
        rc=[c for c in df.columns if not c.startswith("_")]
        rdf=df[df["_area"].isin(areas)][rc]
        raw_rows=[[str(v) if pd.notna(v) else "" for v in row] for row in rdf.itertuples(index=False)]
        sheets.append({"name":"Raw Data","headers":list(rdf.columns),
            "rows":raw_rows,"hdr_color":"dark",
            "col_widths":[16]*len(rdf.columns)})
    return sheets

def build_group_excel(df, areas):
    sheets = _outlet_sheets(df, areas, include_raw=True)
    return _write_xlsx(sheets)

def build_outlet_excel(df, area):
    sheets = _outlet_sheets(df, [area], include_raw=False)
    return _write_xlsx(sheets)

# ── HTML Report ────────────────────────────────────────────────
def build_html_report(df, areas, title="Group Report"):
    delivered=df[df["Order status"]=="Delivered"]; cancelled=df[df["Order status"]=="Cancelled"]
    date_min=df["Order received at"].min(); date_max=df["Order received at"].max()
    outlet_data=[(a,outlet_metrics(df[df["_area"]==a])) for a in sorted(areas) if len(df[df["_area"]==a])>0]
    outlet_data.sort(key=lambda x:x[1]["gmv"],reverse=True)
    CLRS={"Liwan":"#F97316","Dubai Investment Park":"#3B82F6","Oud Metha":"#8B5CF6",
          "Naif":"#EF4444","Al Muteena":"#10B981","Al Hamriya":"#F59E0B"}
    ds=""
    if pd.notna(date_min) and pd.notna(date_max):
        ds=date_min.strftime("%d %b %Y")+" - "+date_max.strftime("%d %b %Y")

    CSS="""<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:Arial,sans-serif;color:#1a1a2e;background:#f8fafc;padding:20px;}
.page{background:white;border-radius:12px;padding:24px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,.06);}
.hdr{color:white;padding:18px 22px;border-radius:10px;margin-bottom:18px;}
h1{font-size:18px;font-weight:700;}
h2{font-size:13px;font-weight:700;color:white;background:#1a1a2e;padding:6px 12px;border-radius:6px;margin:14px 0 8px;}
.sec{padding:6px 12px;border-radius:6px;font-size:12px;font-weight:700;margin:14px 0 8px;color:white;}
table{width:100%;border-collapse:collapse;margin-bottom:14px;font-size:11px;}
th{padding:7px 9px;text-align:left;color:white;}
td{padding:5px 9px;border-bottom:1px solid #f1f5f9;}
tr:nth-child(even) td{background:#f8fafc;}
.kgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px;}
.kcard{border:1px solid #e2e8f0;border-radius:8px;padding:11px;border-top:3px solid #f97316;}
.klabel{font-size:9px;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em;}
.kvalue{font-size:19px;font-weight:700;color:#0f172a;margin-top:3px;}
.ksub{font-size:10px;color:#64748b;margin-top:2px;}
.alert{border-radius:8px;padding:9px 12px;margin:6px 0;font-size:11px;display:flex;gap:8px;}
.alert-r{background:#fff1f2;border:1px solid #fecdd3;color:#9f1239;}
.alert-a{background:#fffbeb;border:1px solid #fed7aa;color:#92400e;}
.alert-g{background:#f0fdf4;border:1px solid #bbf7d0;color:#14532d;}
@media print{body{background:white;padding:0;}.page{box-shadow:none;page-break-after:always;}}
</style>"""

    def tbl(headers,rows,color="#F97316"):
        ths="".join("<th style='background:{}'>".format(color)+str(h)+"</th>" for h in headers)
        trs=""
        for i,r in enumerate(rows):
            bg="#f8fafc" if i%2==0 else "white"
            tds="".join("<td style='background:{}'>".format(bg)+str(v)+"</td>" for v in r)
            trs+="<tr>"+tds+"</tr>"
        return "<table><thead><tr>"+ths+"</tr></thead><tbody>"+trs+"</tbody></table>"

    def kcard(label,value,sub="",color="#F97316"):
        return ('<div class="kcard" style="border-top-color:'+color+';">'
                '<div class="klabel">'+str(label)+'</div>'
                '<div class="kvalue">'+str(value)+'</div>'
                +('<div class="ksub">'+str(sub)+'</div>' if sub else "")+'</div>')

    def sec(text,color="#F97316"):
        return '<div class="sec" style="background:{};">'.format(color)+text+"</div>"

    html="<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'><title>"+title+"</title>"+CSS+"</head><body>"

    m_all=outlet_metrics(df[df["_area"].isin(areas)])
    html+='<div class="page">'
    html+=('<div class="hdr" style="background:#1a1a2e;">'
           '<div style="font-size:11px;color:#f97316;font-weight:700;letter-spacing:.08em;margin-bottom:4px;">AL MADINA HYPERMARKET</div>'
           '<h1 style="color:white;">'+title+'</h1>'
           '<div style="font-size:10px;color:#94a3b8;margin-top:5px;">Period: '+ds+' &nbsp;|&nbsp; '+str(len(outlet_data))+' Outlet(s) &nbsp;|&nbsp; Generated: '+datetime.now().strftime("%d %b %Y %H:%M")+'</div>'
           '</div>')
    html+=sec("GROUP KEY METRICS","#1a1a2e")
    html+='<div class="kgrid">'
    html+=kcard("Total Orders",m_all["total"])
    html+=kcard("Delivered",m_all["delivered"],"{:.1f}% rate".format(m_all["del_rate"]),"#10b981")
    html+=kcard("Cancelled",m_all["cancelled"],"{:.1f}% rate".format(m_all["can_rate"]),"#ef4444")
    html+=kcard("Gross Revenue","AED {:,.0f}".format(m_all["gmv"]),"","#3b82f6")
    html+=kcard("Total Payout","AED {:,.0f}".format(m_all["payout"]),"","#8b5cf6")
    html+=kcard("Avg Order","AED {:,.1f}".format(m_all["avg_order"]),"","#f97316")
    html+=kcard("Commission","AED {:,.0f}".format(m_all["commission"]),"{:.1f}% of GMV".format(safe_div(m_all["commission"],m_all["gmv"],pct=True)))
    html+=kcard("Avg Delivery",("{:.0f} min".format(m_all["del_time"])) if m_all["del_time"] else "N/A")
    html+=kcard("Lost to Cancels","AED {:,.0f}".format(m_all["lost_gmv"]),"","#ef4444")
    html+='</div>'

    if len(outlet_data)>1:
        html+=sec("OUTLET LEAGUE TABLE")
        html+=tbl(["#","Outlet","Orders","Delivered","Cancel %","GMV (AED)","Payout","Avg Order","Del Time"],
            [[str(rank),area,str(om["total"]),str(om["delivered"]),
              "{:.1f}%".format(om["can_rate"]),"AED {:,.0f}".format(om["gmv"]),
              "AED {:,.0f}".format(om["payout"]),"AED {:,.1f}".format(om["avg_order"]),
              ("{:.0f} min".format(om["del_time"])) if om["del_time"] else "N/A"]
             for rank,(area,om) in enumerate(outlet_data,1)])

    # Cancellation analysis section
    html+=sec("CANCELLATION ANALYSIS","#EF4444")
    sub_can=cancelled[cancelled["_area"].isin(areas)]
    can_table_rows=[]
    for area,om in outlet_data:
        sc=sub_can[sub_can["_area"]==area]
        ow=sc["Cancellation owner"].value_counts()
        can_table_rows.append([area,str(len(sc)),"{:.1f}%".format(om["can_rate"]),
            str(int(ow.get("Vendor",0))),str(int(ow.get("Customer",0))),
            str(int(sc["Cancellation reason"].str.contains("Item not available",na=False).sum())),
            "AED {:,.0f}".format(sc["Subtotal"].sum())])
    html+=tbl(["Outlet","Cancelled","Rate","Vendor","Customer","Item N/A","Lost GMV"],
              can_table_rows,"#EF4444")

    # Alerts
    for area,om in outlet_data:
        if om["can_rate"]>40:
            html+='<div class="alert alert-r">🚨 <b>'+area+'</b>: CRITICAL '+"{:.1f}%".format(om["can_rate"])+" cancellation rate - immediate action needed.</div>"
        elif om["can_rate"]>10:
            html+='<div class="alert alert-a">⚠️ <b>'+area+'</b>: Elevated '+"{:.1f}%".format(om["can_rate"])+" cancellation rate.</div>"
        else:
            html+='<div class="alert alert-g">✅ <b>'+area+'</b>: Good '+"{:.1f}%".format(om["can_rate"])+" cancellation rate.</div>"

    html+='</div>'  # end group page

    # Per-outlet pages
    for area,om in outlet_data:
        color=CLRS.get(area,"#F97316")
        sub_df=df[df["_area"]==area]; sub_d=sub_df[sub_df["Order status"]=="Delivered"]
        sub_c=sub_df[sub_df["Order status"]=="Cancelled"]
        html+='<div class="page">'
        html+='<div class="hdr" style="background:'+color+';">'
        html+='<h1>'+area+'</h1>'
        html+='<div style="font-size:10px;color:rgba(255,255,255,.7);margin-top:4px;">'+sub_df["Restaurant name"].iloc[0]+'</div>'
        html+='</div>'
        html+='<div class="kgrid">'
        html+=kcard("Orders",om["total"],"",color)
        html+=kcard("Delivered",om["delivered"],"{:.1f}%".format(om["del_rate"]),"#10b981")
        html+=kcard("Cancelled",om["cancelled"],"{:.1f}%".format(om["can_rate"]),"#ef4444")
        html+=kcard("GMV","AED {:,.0f}".format(om["gmv"]),"",color)
        html+=kcard("Payout","AED {:,.0f}".format(om["payout"]),"","#8b5cf6")
        html+=kcard("Avg Order","AED {:,.1f}".format(om["avg_order"]),"",color)
        html+='</div>'
        html+=sec("Daily Performance",color)
        ds2=sub_df.groupby("_date").agg(total=("Order ID","count"),
            delivered=("Order status",lambda x:(x=="Delivered").sum()),
            cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
            gmv=("Subtotal","sum")).reset_index()
        ds2["can_pct"]=ds2.apply(lambda r:safe_div(r["cancelled"],r["total"],pct=True),axis=1)
        html+=tbl(["Date","Total","Delivered","Cancelled","Cancel %","GMV (AED)"],
            [[str(r["_date"]),int(r["total"]),int(r["delivered"]),int(r["cancelled"]),
              "{:.1f}%".format(r["can_pct"]),"AED {:,.0f}".format(r["gmv"])]
             for _,r in ds2.iterrows()],color)
        if len(sub_c)>0:
            html+=sec("Cancellation Detail","#EF4444")
            cg=sub_c.groupby(["Cancellation owner","Cancellation reason"]).agg(
                Count=("Order ID","count"),Lost=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False)
            html+=tbl(["Owner","Reason","Count","Lost GMV"],
                [[str(r["Cancellation owner"]).strip(),str(r["Cancellation reason"]),
                  int(r["Count"]),"AED {:,.0f}".format(r["Lost"])] for _,r in cg.iterrows()],"#EF4444")
            # cancelled items
            from collections import Counter
            ctr=Counter()
            for items_str in sub_c["Order Items"].dropna():
                for item in str(items_str).split(","):
                    item=item.strip(); parts=item.split(" ",1)
                    name=parts[1].strip() if len(parts)==2 and parts[0].isdigit() else item
                    if name and len(name)>3 and not name.replace(".","").replace("g","").replace("ml","").isdigit():
                        ctr[name]+=int(parts[0]) if len(parts)==2 and parts[0].isdigit() else 1
            if ctr:
                html+=sec("Cancelled Items (from cancelled orders)","#EF4444")
                html+=tbl(["Item","Qty in Cancelled Orders"],list(ctr.most_common(10)),"#EF4444")
        html+=sec("Financial Breakdown",color)
        fin=[("GMV",om["gmv"]),("Commission",om["commission"]),
             ("Op Charges",om["op_charges"]),("Online Fee",om["online_fee"]),("Payout",om["payout"])]
        html+=tbl(["Item","Total (AED)","Per Order","% of GMV"],
            [[l,"AED {:,.2f}".format(v),"AED {:,.2f}".format(safe_div(v,om["delivered"])),
              "{:.1f}%".format(safe_div(v,om["gmv"],pct=True))] for l,v in fin],color)
        html+='</div>'

    html+='</body></html>'
    return html.encode("utf-8")

# ══════════════════════════════════════════════════════════════
#  MAIN APP
# ══════════════════════════════════════════════════════════════
def main():
    with st.sidebar:
        st.markdown("""<div style='text-align:center;padding:16px 0 10px;'>
            <div style='font-size:24px;margin-bottom:5px;'>📊</div>
            <div style='font-size:13px;font-weight:600;color:#f1f5f9;'>Al Madina Dashboard</div>
            <div style='font-size:10px;color:#475569;margin-top:2px;'>Performance Intelligence</div>
        </div>""",unsafe_allow_html=True)
        st.divider()
        uploaded=st.file_uploader("Upload Order Export (.xlsx)",type=["xlsx","xls"],
                                   help="Upload Talabat order details export")
        st.divider()

    if uploaded is None:
        st.markdown("""<div class="top-bar">
            <div class="top-bar-brand">
                <div class="logo-circle">AM</div>
                <div><h1>Al Madina Performance Dashboard</h1>
                <span>Multi-Outlet Intelligence Platform</span></div>
            </div></div>""",unsafe_allow_html=True)
        c1,c2,c3=st.columns(3)
        for col,icon,title,desc,color in [
            (c1,"🏪","Multi-Outlet Analysis","Compare all branches side-by-side with league tables and benchmarks.","blue"),
            (c2,"📈","Full Analytics","Sales, cancellations, items, delivery timing, operations — every metric.","orange"),
            (c3,"📥","Smart Downloads","Excel per outlet or group-wide, plus HTML reports. Zero failed installs.","green"),
        ]:
            with col:
                st.markdown('<div class="kpi-card {c}" style="padding:22px;"><div style="font-size:26px;margin-bottom:8px;">{i}</div><div style="font-size:14px;font-weight:600;color:#0f172a;margin-bottom:6px;">{t}</div><div style="font-size:12px;color:#64748b;">{d}</div></div>'.format(c=color,i=icon,t=title,d=desc),unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)
        alert("Upload your Talabat order details export (.xlsx) in the sidebar to launch.","blue","👈")
        return

    with st.spinner("Loading data..."):
        df=load_data(uploaded.read())
    if "Order status" not in df.columns:
        alert("Could not detect required columns. Check file format.","red","❌"); return

    outlets_all=sorted(df["_area"].unique())

    with st.sidebar:
        st.markdown('<div style="font-size:10px;color:#94a3b8;text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;">Date Range</div>',unsafe_allow_html=True)
        date_range=st.date_input("",value=(df["_date"].min(),df["_date"].max()),
            min_value=df["_date"].min(),max_value=df["_date"].max(),label_visibility="collapsed")
        st.markdown('<div style="font-size:10px;color:#94a3b8;text-transform:uppercase;letter-spacing:.08em;margin:10px 0 6px;">Outlets</div>',unsafe_allow_html=True)
        selected=st.multiselect("",options=outlets_all,default=outlets_all,label_visibility="collapsed")
        st.divider()
        st.markdown('<div style="font-size:10px;color:#475569;text-align:center;padding:6px;">Al Madina Hypermarket<br>v3.0</div>',unsafe_allow_html=True)

    if len(date_range)==2:
        df=df[(df["_date"]>=date_range[0])&(df["_date"]<=date_range[1])]
    if selected:
        df=df[df["_area"].isin(selected)]

    delivered=df[df["Order status"]=="Delivered"]; cancelled=df[df["Order status"]=="Cancelled"]
    date_min=df["Order received at"].min(); date_max=df["Order received at"].max()

    st.markdown("""<div class="top-bar">
        <div class="top-bar-brand">
            <div class="logo-circle">AM</div>
            <div><h1>Al Madina Performance Dashboard</h1>
            <span>{n} outlet{pl} &nbsp;·&nbsp; {d1} – {d2}</span></div>
        </div>
        <div style="font-size:11px;color:#475569;">Generated {now}</div>
    </div>""".format(
        n=len(selected),pl="s" if len(selected)!=1 else "",
        d1=date_min.strftime("%d %b") if pd.notna(date_min) else "?",
        d2=date_max.strftime("%d %b %Y") if pd.notna(date_max) else "?",
        now=datetime.now().strftime("%d %b %Y %H:%M")),unsafe_allow_html=True)

    tab1,tab2,tab3,tab4,tab5,tab6=st.tabs([
        "🏢 Group","🏪 Outlet","💰 Sales","❌ Cancellations","⚙️ Operations","📥 Downloads"])

    # ── TAB 1: GROUP ───────────────────────────────────────────
    with tab1:
        m=outlet_metrics(df)
        st.markdown('<div class="section-label">Group KPIs</div>',unsafe_allow_html=True)
        c1,c2,c3,c4,c5,c6=st.columns(6)
        with c1: kpi("Total Orders",m["total"],"","orange")
        with c2: kpi("Delivered",m["delivered"],"{:.1f}% rate".format(m["del_rate"]),"green")
        with c3: kpi("Cancelled",m["cancelled"],"{:.1f}% rate".format(m["can_rate"]),"red")
        with c4: kpi("Gross Revenue","AED {:,.0f}".format(m["gmv"]),"","blue")
        with c5: kpi("Total Payout","AED {:,.0f}".format(m["payout"]),"","purple")
        with c6: kpi("Avg Order","AED {:,.1f}".format(m["avg_order"]),"","amber")
        c1,c2,c3,c4,c5,c6=st.columns(6)
        with c1: kpi("Commission","AED {:,.0f}".format(m["commission"]),"{:.1f}% of GMV".format(safe_div(m["commission"],m["gmv"],pct=True)),"orange")
        with c2: kpi("Op. Charges","AED {:,.0f}".format(m["op_charges"]),"","blue")
        with c3: kpi("Avg Delivery",("{:.0f} min".format(m["del_time"])) if m["del_time"] else "N/A","","green")
        with c4: kpi("Pro Orders","{:.0f}%".format(m["pro_pct"]),"","purple")
        with c5: kpi("Online Pay","{:.0f}%".format(m["online_pct"]),"","amber")
        with c6: kpi("Complaints",str(m["complaints"]),"{:.1f}%".format(m["complaint_rate"]),"red")

        st.markdown('<div class="section-label">Outlet League Table — Ranked by GMV</div>',unsafe_allow_html=True)
        outlet_data=[(a,outlet_metrics(df[df["_area"]==a])) for a in selected if len(df[df["_area"]==a])>0]
        outlet_data.sort(key=lambda x:x[1]["gmv"],reverse=True)
        max_gmv=max((o[1]["gmv"] for o in outlet_data),default=1)

        for rank,(area,om) in enumerate(outlet_data,1):
            color=OUTLET_COLORS.get(area,"#64748b")
            can_cls,can_lbl=can_badge(om["can_rate"]); del_cls,del_lbl=del_badge(om["del_time"])
            c1,c2,c3=st.columns([2,3,2])
            with c1:
                store_id=str(df[df["_area"]==area]["Store ID"].iloc[0]) if len(df[df["_area"]==area])>0 else "N/A"
                st.markdown(('<div class="outlet-card"><div class="outlet-header"><div>'
                    '<div class="outlet-name"><span style="font-size:18px;font-weight:700;color:{c};margin-right:6px;">#{r}</span>{a}</div>'
                    '<div class="outlet-area">Store {sid}</div></div>'
                    '<div><span class="outlet-badge {cc}" style="display:block;margin-bottom:3px;">{cl} cancel</span>'
                    '<span class="outlet-badge {dc}">{dl} delivery</span></div></div>'
                    '<div style="font-size:22px;font-weight:600;color:#0f172a;">AED {gmv:,.0f}</div>'
                    '<div style="font-size:12px;color:#94a3b8;margin-top:2px;">{tot} orders &nbsp;·&nbsp; {d} delivered</div>'
                    '</div>').format(c=color,r=rank,a=area,sid=store_id,
                        cc=can_cls,cl=can_lbl,dc=del_cls,dl=del_lbl,
                        gmv=om["gmv"],tot=om["total"],d=om["delivered"]),
                    unsafe_allow_html=True)
            with c2:
                st.markdown('<div style="margin-top:8px;">',unsafe_allow_html=True)
                progress_bar("Orders delivered",om["delivered"],om["total"],color,"/"+str(om["total"]))
                progress_bar("GMV vs top outlet",round(om["gmv"]),round(max_gmv),color," AED")
                progress_bar("Pro order share",round(om["pro_pct"]),100,color,"%")
                progress_bar("Cancel rate",round(om["can_rate"],1),60,"#ef4444" if om["can_rate"]>20 else "#f97316","%")
                st.markdown('</div>',unsafe_allow_html=True)
            with c3:
                st.markdown(('<div style="margin-top:8px;">'
                    '<div class="compare-row"><div class="compare-chip">Payout <span class="chip-val">AED {p:,.0f}</span></div>'
                    '<div class="compare-chip">Avg <span class="chip-val">AED {a:,.0f}</span></div></div>'
                    '<div class="compare-row"><div class="compare-chip">Commission <span class="chip-val">AED {c:,.0f}</span></div></div>'
                    '<div class="compare-row"><div class="compare-chip">Del <span class="chip-val">{d}</span></div>'
                    '<div class="compare-chip">Cancel <span class="chip-val">{cr:.1f}%</span></div></div>'
                    '<div class="compare-row"><div class="compare-chip">Online <span class="chip-val">{o:.0f}%</span></div>'
                    '<div class="compare-chip">Pro <span class="chip-val">{pr:.0f}%</span></div></div>'
                    '</div>').format(p=om["payout"],a=om["avg_order"],c=om["commission"],
                        d=("{:.0f}m".format(om["del_time"])) if om["del_time"] else "N/A",
                        cr=om["can_rate"],o=om["online_pct"],pr=om["pro_pct"]),
                    unsafe_allow_html=True)

        st.markdown('<div class="section-label">Daily Trend</div>',unsafe_allow_html=True)
        dg=df.groupby("_date").agg(total=("Order ID","count"),
            delivered=("Order status",lambda x:(x=="Delivered").sum()),
            cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
            gmv=("Subtotal","sum")).reset_index()
        dg.columns=["Date","Total","Delivered","Cancelled","GMV (AED)"]
        dg["Cancel %"]=dg.apply(lambda r:round(safe_div(r["Cancelled"],r["Total"],pct=True),1),axis=1)
        dg["GMV (AED)"]=dg["GMV (AED)"].round(0)
        st.dataframe(_style_df(dg,heat_col="GMV (AED)",heat_col2="Cancel %"),
                     use_container_width=True,hide_index=True)

        st.markdown('<div class="section-label">Key Insights</div>',unsafe_allow_html=True)
        if outlet_data:
            top=outlet_data[0]
            alert("<b>{}</b> leads with AED {:,.0f} GMV — {:.0f}% of group total.".format(
                top[0],top[1]["gmv"],safe_div(top[1]["gmv"],m["gmv"],pct=True)),"green","🏆")
            hcan=max(outlet_data,key=lambda x:x[1]["can_rate"])
            if hcan[1]["can_rate"]>15:
                alert("<b>{}</b> has critical cancellation rate <b>{:.1f}%</b>. Lost AED {:,.0f}.".format(
                    hcan[0],hcan[1]["can_rate"],hcan[1]["lost_gmv"]),"red","🚨")
            slow=["{} ({:.0f}m)".format(a,o["del_time"]) for a,o in outlet_data if o["del_time"] and o["del_time"]>42]
            if slow: alert("Slow delivery outlets (>42 min): <b>{}</b>".format(", ".join(slow)),"amber","⏱")
            if m["can_rate"]>10:
                alert("Group cancel rate <b>{:.1f}%</b>. Total lost: <b>AED {:,.0f}</b>.".format(
                    m["can_rate"],m["lost_gmv"]),"amber","💸")

    # ── TAB 2: OUTLET ─────────────────────────────────────────
    with tab2:
        sel_area=st.selectbox("Select outlet",options=[a for a in outlets_all if a in selected])
        sub_df=df[df["_area"]==sel_area]; sub_d=sub_df[sub_df["Order status"]=="Delivered"]
        sub_c=sub_df[sub_df["Order status"]=="Cancelled"]; om=outlet_metrics(sub_df)
        color=OUTLET_COLORS.get(sel_area,"#64748b")
        can_cls,can_lbl=can_badge(om["can_rate"]); del_cls,del_lbl=del_badge(om["del_time"])
        fname=sub_df["Restaurant name"].iloc[0] if len(sub_df)>0 else sel_area
        st.markdown(('<div style="background:white;border:1px solid #e2e8f0;border-radius:14px;'
                     'padding:16px 22px;margin-bottom:16px;border-left:4px solid {c};">'
                     '<div style="font-size:17px;font-weight:600;color:#0f172a;">{a}</div>'
                     '<div style="font-size:11px;color:#94a3b8;margin-top:2px;">{f}</div>'
                     '<div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap;">'
                     '<span class="outlet-badge {cc}">Cancel: {cl} ({cr:.1f}%)</span>'
                     '<span class="outlet-badge {dc}">Delivery: {dl} ({dm})</span>'
                     '</div></div>').format(c=color,a=sel_area,f=fname,
                         cc=can_cls,cl=can_lbl,cr=om["can_rate"],dc=del_cls,dl=del_lbl,
                         dm=("{:.0f} min".format(om["del_time"])) if om["del_time"] else "N/A"),
                    unsafe_allow_html=True)
        c1,c2,c3,c4,c5,c6=st.columns(6)
        with c1: kpi("Orders",om["total"],"","orange")
        with c2: kpi("Delivered",om["delivered"],"{:.1f}%".format(om["del_rate"]),"green")
        with c3: kpi("Cancelled",om["cancelled"],"{:.1f}%".format(om["can_rate"]),"red")
        with c4: kpi("GMV","AED {:,.0f}".format(om["gmv"]),"","blue")
        with c5: kpi("Payout","AED {:,.0f}".format(om["payout"]),"","purple")
        with c6: kpi("Avg Order","AED {:,.1f}".format(om["avg_order"]),"","amber")
        c1,c2,c3,c4=st.columns(4)
        with c1: kpi("Avg Delivery",("{:.0f} min".format(om["del_time"])) if om["del_time"] else "N/A","","orange")
        with c2: kpi("Avg Prep",("{:.0f} min".format(om["prep_time"])) if om["prep_time"] else "N/A","","blue")
        with c3: kpi("Last Mile",("{:.0f} min".format(om["last_mile"])) if om["last_mile"] else "N/A","","green")
        with c4: kpi("Complaints",str(om["complaints"]),"{:.1f}%".format(om["complaint_rate"]),"red")

        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="section-label">Daily Performance</div>',unsafe_allow_html=True)
            ds3=sub_df.groupby("_date").agg(total=("Order ID","count"),
                delivered=("Order status",lambda x:(x=="Delivered").sum()),
                cancelled=("Order status",lambda x:(x=="Cancelled").sum()),
                gmv=("Subtotal","sum"),payout=("Payout Amount","sum")).reset_index()
            ds3.columns=["Date","Total","Delivered","Cancelled","GMV (AED)","Payout (AED)"]
            ds3["Cancel %"]=ds3.apply(lambda r:round(safe_div(r["Cancelled"],r["Total"],pct=True),1),axis=1)
            ds3["GMV (AED)"]=ds3["GMV (AED)"].round(0); ds3["Payout (AED)"]=ds3["Payout (AED)"].round(0)
            st.dataframe(_style_df(ds3,heat_col="GMV (AED)"),use_container_width=True,hide_index=True)
        with c2:
            st.markdown('<div class="section-label">Financial Breakdown</div>',unsafe_allow_html=True)
            fin={"GMV":om["gmv"],"Commission":om["commission"],"Op Charges":om["op_charges"],
                 "Online Fee":om["online_fee"],"Total Payout":om["payout"]}
            fin_df=pd.DataFrame([{"Item":k,"Total":round(v,2),
                "Per Order":round(safe_div(v,om["delivered"]),2),
                "% of GMV":round(safe_div(v,om["gmv"],pct=True),1)} for k,v in fin.items()])
            st.dataframe(fin_df,use_container_width=True,hide_index=True)

        if len(sub_c)>0:
            st.markdown('<div class="section-label">Cancellation Detail</div>',unsafe_allow_html=True)
            c1,c2=st.columns(2)
            with c1:
                cd=sub_c.groupby(["Cancellation owner","Cancellation reason"]).agg(
                    Count=("Order ID","count"),Lost_GMV=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False)
                cd.columns=["Owner","Reason","Count","Lost GMV (AED)"]; cd["Lost GMV (AED)"]=cd["Lost GMV (AED)"].round(2)
                st.dataframe(cd,use_container_width=True,hide_index=True)
            with c2:
                from collections import Counter
                ctr=Counter()
                for items_str in sub_c["Order Items"].dropna():
                    for item in str(items_str).split(","):
                        item=item.strip(); parts=item.split(" ",1)
                        name=parts[1].strip() if len(parts)==2 and parts[0].isdigit() else item
                        if name and len(name)>3 and not name.replace(".","").replace("g","").replace("ml","").isdigit():
                            ctr[name]+=int(parts[0]) if len(parts)==2 and parts[0].isdigit() else 1
                if ctr:
                    items_df=pd.DataFrame(ctr.most_common(15),columns=["Cancelled Item","Qty"])
                    st.caption("Items from cancelled orders")
                    st.dataframe(items_df,use_container_width=True,hide_index=True)
        else:
            alert("No cancellations for {} in this period.".format(sel_area),"green","✅")

        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="section-label">Hourly Pattern</div>',unsafe_allow_html=True)
            h=sub_df.groupby("_hour")["Order ID"].count().reset_index()
            h.columns=["Hour","Orders"]; h["Hour"]=h["Hour"].apply(lambda x:"{:02d}:00".format(x))
            st.dataframe(_style_df(h,heat_col="Orders"),use_container_width=True,hide_index=True)
        with c2:
            st.markdown('<div class="section-label">Day of Week</div>',unsafe_allow_html=True)
            dow_order=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
            dw=sub_df.groupby("_dow")["Order ID"].count().reindex(dow_order).reset_index()
            dw.columns=["Day","Orders"]; dw["Orders"]=dw["Orders"].fillna(0).astype(int)
            st.dataframe(_style_df(dw,heat_col="Orders"),use_container_width=True,hide_index=True)

    # ── TAB 3: SALES ──────────────────────────────────────────
    with tab3:
        st.markdown('<div class="section-label">Revenue by Outlet</div>',unsafe_allow_html=True)
        rev=[]
        for area in selected:
            sd=delivered[delivered["_area"]==area]; s=df[df["_area"]==area]
            if len(s)==0: continue
            rev.append({"Outlet":area,"Orders":len(s),"Delivered":len(sd),
                "GMV (AED)":round(sd["Subtotal"].sum(),0),"Payout (AED)":round(sd["Payout Amount"].sum(),0),
                "Commission (AED)":round(sd["Commission"].sum(),0),
                "Avg Order":round(safe_div(sd["Subtotal"].sum(),len(sd)),1),
                "Pro %":round(safe_div(len(sd[sd["Is Pro Order"]=="Y"]),len(sd),pct=True),1),
                "Online %":round(safe_div(len(sd[sd["Payment type"]=="Online"]),len(sd),pct=True),1)})
        rev_df=pd.DataFrame(rev).sort_values("GMV (AED)",ascending=False)
        st.dataframe(_style_df(rev_df,heat_col="GMV (AED)"),use_container_width=True,hide_index=True)
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="section-label">GMV by Outlet by Day</div>',unsafe_allow_html=True)
            gp=delivered.groupby(["_date","_area"])["Subtotal"].sum().unstack(fill_value=0).round(0)
            st.dataframe(gp,use_container_width=True)
        with c2:
            st.markdown('<div class="section-label">Day-of-Week Pattern</div>',unsafe_allow_html=True)
            dow_order=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
            dw2=df.groupby("_dow").agg(Orders=("Order ID","count"),GMV=("Subtotal","sum")).reindex(dow_order).reset_index()
            dw2.columns=["Day","Total Orders","GMV (AED)"]; dw2["GMV (AED)"]=dw2["GMV (AED)"].round(0)
            st.dataframe(_style_df(dw2,heat_col="Total Orders"),use_container_width=True,hide_index=True)
        st.markdown('<div class="section-label">Payment Mix by Outlet</div>',unsafe_allow_html=True)
        pay=[]
        for area in selected:
            sd=delivered[delivered["_area"]==area]
            if len(sd)==0: continue
            on=sd[sd["Payment type"]=="Online"]; ca=sd[sd["Payment type"]=="Cash"]
            pay.append({"Outlet":area,"Online Orders":len(on),"Cash Orders":len(ca),
                "Online GMV":round(on["Subtotal"].sum(),0),"Cash GMV":round(ca["Subtotal"].sum(),0),
                "Online %":round(safe_div(len(on),len(sd),pct=True),1)})
        st.dataframe(pd.DataFrame(pay),use_container_width=True,hide_index=True)

    # ── TAB 4: CANCELLATIONS ──────────────────────────────────
    with tab4:
        st.markdown('<div class="section-label">Cancellation Scorecard</div>',unsafe_allow_html=True)
        can=[]
        for area in selected:
            s=df[df["_area"]==area]; sc=s[s["Order status"]=="Cancelled"]
            if len(s)==0: continue
            ow=sc["Cancellation owner"].value_counts()
            can.append({"Outlet":area,"Total":len(s),"Cancelled":len(sc),
                "Cancel Rate %":round(safe_div(len(sc),len(s),pct=True),1),
                "Vendor":int(ow.get("Vendor",0)),"Customer":int(ow.get("Customer",0)),
                "Rider":int(ow.get("Rider",0)),
                "Item N/A":int(sc["Cancellation reason"].str.contains("Item not available",na=False).sum()),
                "Fraudulent":int(sc["Cancellation reason"].str.contains("Fraudulent",na=False).sum()),
                "Lost GMV (AED)":round(sc["Subtotal"].sum(),0)})
        can_df=pd.DataFrame(can).sort_values("Cancel Rate %",ascending=False)
        st.dataframe(_style_df(can_df,heat_col="Lost GMV (AED)",heat_col2="Cancel Rate %"),
                     use_container_width=True,hide_index=True)
        st.markdown('<div class="section-label">Outlet Alerts</div>',unsafe_allow_html=True)
        for _,r in can_df.iterrows():
            if r["Cancel Rate %"]>40:
                alert("<b>{}</b>: CRITICAL — {:.1f}% cancel rate. {} cancelled / {} orders. Lost AED {:,.0f}.".format(
                    r["Outlet"],r["Cancel Rate %"],r["Cancelled"],r["Total"],r["Lost GMV (AED)"]),"red","🚨")
            elif r["Cancel Rate %"]>10:
                primary="item availability" if r["Item N/A"]>r["Fraudulent"] else "fraudulent orders"
                alert("<b>{}</b>: Elevated {:.1f}% cancel rate. Main cause: {}.".format(
                    r["Outlet"],r["Cancel Rate %"],primary),"amber","⚠️")
            else:
                alert("<b>{}</b>: Good {:.1f}% cancel rate.".format(r["Outlet"],r["Cancel Rate %"]),"green","✅")

        st.markdown('<div class="section-label">Cancelled Items Analysis</div>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        from collections import Counter
        with c1:
            st.caption("Top cancelled items across all outlets")
            all_item_ctr=Counter()
            for items_str in cancelled["Order Items"].dropna():
                for item in str(items_str).split(","):
                    item=item.strip(); parts=item.split(" ",1)
                    name=parts[1].strip() if len(parts)==2 and parts[0].isdigit() else item
                    if name and len(name)>3 and not name.replace(".","").replace("g","").replace("ml","").isdigit():
                        all_item_ctr[name]+=int(parts[0]) if len(parts)==2 and parts[0].isdigit() else 1
            if all_item_ctr:
                items_df=pd.DataFrame(all_item_ctr.most_common(20),columns=["Cancelled Item","Qty in Cancelled Orders"])
                st.dataframe(items_df,use_container_width=True,hide_index=True)
            else:
                st.info("No item data in cancelled orders.")
        with c2:
            st.caption("Cancel reasons breakdown")
            if len(cancelled)>0:
                r_all=cancelled.groupby(["_area","Cancellation reason"]).agg(
                    Count=("Order ID","count"),Lost=("Subtotal","sum")).reset_index().sort_values("Count",ascending=False)
                r_all.columns=["Outlet","Reason","Count","Lost GMV (AED)"]; r_all["Lost GMV (AED)"]=r_all["Lost GMV (AED)"].round(2)
                st.dataframe(r_all,use_container_width=True,hide_index=True)

    # ── TAB 5: OPERATIONS ────────────────────────────────────
    with tab5:
        st.markdown('<div class="section-label">Delivery Timing Comparison</div>',unsafe_allow_html=True)
        tm=[]
        for area in selected:
            sd=delivered[delivered["_area"]==area]
            if len(sd)==0: continue
            tot=(sd["Delivered at"]-sd["Order received at"]).dt.total_seconds()/60; tot=tot[tot>0]
            prep=(sd["Ready to pick up at"]-sd["Accepted at"]).dt.total_seconds()/60; prep=prep[prep>0]
            lm=(sd["Delivered at"]-sd["In delivery at"]).dt.total_seconds()/60; lm=lm[lm>0]
            tm.append({"Outlet":area,"Avg Total (min)":round(tot.mean(),1) if len(tot)>0 else None,
                "Avg Prep (min)":round(prep.mean(),1) if len(prep)>0 else None,
                "Avg Last Mile (min)":round(lm.mean(),1) if len(lm)>0 else None,
                "Max (min)":round(tot.max(),1) if len(tot)>0 else None,
                "Deliveries":len(tot)})
        tm_df=pd.DataFrame(tm).sort_values("Avg Total (min)")
        st.dataframe(_style_df(tm_df,heat_col="Avg Total (min)",reverse=True),use_container_width=True,hide_index=True)
        st.markdown('<div class="section-label">Timing Alerts</div>',unsafe_allow_html=True)
        for _,r in tm_df.iterrows():
            v=r["Avg Total (min)"]
            if v is None: continue
            if v<=33: alert("<b>{}</b>: Excellent — {} min avg".format(r["Outlet"],v),"green","🚀")
            elif v<=42: alert("<b>{}</b>: Acceptable — {} min avg. Benchmark is 35 min.".format(r["Outlet"],v),"blue","🕐")
            else: alert("<b>{}</b>: Slow — {} min avg. Prep: {}m, Last mile: {}m.".format(
                r["Outlet"],v,r["Avg Prep (min)"],r["Avg Last Mile (min)"]),"red","⚠️")
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="section-label">Operational Costs</div>',unsafe_allow_html=True)
            op=[]
            for area in selected:
                sd=delivered[delivered["_area"]==area]
                if len(sd)==0: continue
                tc=sd["Commission"].sum()+sd["Operational Charges"].sum()+sd["Online Payment Fee"].sum()
                op.append({"Outlet":area,"Commission":round(sd["Commission"].sum(),0),
                    "Op Charges":round(sd["Operational Charges"].sum(),0),
                    "Online Fee":round(sd["Online Payment Fee"].sum(),0),
                    "Total Cost":round(tc,0),"Cost/Order":round(safe_div(tc,len(sd)),1)})
            st.dataframe(pd.DataFrame(op).sort_values("Total Cost",ascending=False),use_container_width=True,hide_index=True)
        with c2:
            st.markdown('<div class="section-label">Hourly Demand by Outlet</div>',unsafe_allow_html=True)
            hp=df.groupby(["_hour","_area"])["Order ID"].count().unstack(fill_value=0)
            hp.index=["{:02d}:00".format(h) for h in hp.index]
            st.dataframe(hp,use_container_width=True)

    # ── TAB 6: DOWNLOADS ─────────────────────────────────────
    with tab6:
        st.markdown('<div class="section-label">Download Reports</div>',unsafe_allow_html=True)

        # Outlet data for chips
        outlet_data2=[(a,outlet_metrics(df[df["_area"]==a])) for a in selected if len(df[df["_area"]==a])>0]
        outlet_data2.sort(key=lambda x:x[1]["gmv"],reverse=True)
        m2=outlet_metrics(df)

        # ── Group Downloads ────────────────────────────────
        st.markdown("""<div style="background:white;border:1px solid #e2e8f0;border-radius:16px;padding:22px;margin-bottom:16px;">
            <div style="display:flex;align-items:center;gap:12px;margin-bottom:14px;">
                <div class="dl-icon orange" style="background:#fff7ed;">🏢</div>
                <div><div class="dl-title">Group Report — All {} Outlets</div>
                <div style="font-size:11px;color:#94a3b8;margin-top:2px;">{} orders &nbsp;·&nbsp; AED {:,.0f} GMV &nbsp;·&nbsp; {:.1f}% cancel rate</div>
                </div></div>""".format(len(outlet_data2),m2["total"],m2["gmv"],m2["can_rate"]),
            unsafe_allow_html=True)
        gc1,gc2=st.columns(2)
        with gc1:
            if st.button("📊 Generate Group Excel Report",use_container_width=True,type="primary",key="grp_xls"):
                with st.spinner("Building group Excel..."):
                    try:
                        xls=build_group_excel(df,selected)
                        st.download_button("💾 Download Group Excel",data=xls,
                            file_name="AlMadina_Group_"+datetime.now().strftime("%Y%m%d_%H%M")+".xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True)
                        st.success("✅ Excel built — includes Group Summary, Cancellations sheet, per-outlet sheets, Raw Data")
                    except Exception as e:
                        st.error("Excel failed: "+str(e))
        with gc2:
            if st.button("📄 Generate Group HTML Report",use_container_width=True,key="grp_html"):
                with st.spinner("Building report..."):
                    try:
                        html=build_html_report(df,selected,"Group Performance Report")
                        st.download_button("💾 Download Group HTML",data=html,
                            file_name="AlMadina_Group_"+datetime.now().strftime("%Y%m%d_%H%M")+".html",
                            mime="text/html",use_container_width=True)
                        st.info("💡 Open in browser → File → Print → Save as PDF",icon="💡")
                    except Exception as e:
                        st.error("HTML failed: "+str(e))
        st.markdown("</div>",unsafe_allow_html=True)

        # ── Per-Outlet Downloads ───────────────────────────
        st.markdown("""<div style="background:white;border:1px solid #e2e8f0;border-radius:16px;padding:22px;">
            <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;">
                <div class="dl-icon blue" style="background:#eff6ff;">🏪</div>
                <div><div class="dl-title">Individual Outlet Reports</div>
                <div style="font-size:11px;color:#94a3b8;margin-top:2px;">Download a dedicated report for each outlet to share with branch managers</div>
                </div></div>""",unsafe_allow_html=True)

        for area,om in outlet_data2:
            color=OUTLET_COLORS.get(area,"#64748b")
            can_cls,can_lbl=can_badge(om["can_rate"])
            exp=st.expander("📍 {} — AED {:,.0f} GMV &nbsp;·&nbsp; {} orders &nbsp;·&nbsp; {:.1f}% cancel".format(
                area,om["gmv"],om["total"],om["can_rate"]))
            with exp:
                oc1,oc2,oc3,oc4=st.columns(4)
                with oc1: kpi("GMV","AED {:,.0f}".format(om["gmv"]),"","orange")
                with oc2: kpi("Orders",str(om["total"]),"{} delivered".format(om["delivered"]),"blue")
                with oc3: kpi("Cancel","{:.1f}%".format(om["can_rate"]),can_lbl,"red" if om["can_rate"]>20 else "amber")
                with oc4: kpi("Del Time",("{:.0f} min".format(om["del_time"])) if om["del_time"] else "N/A","","green")
                st.markdown("<br>",unsafe_allow_html=True)
                bc1,bc2=st.columns(2)
                with bc1:
                    if st.button("📊 Excel — {}".format(area),use_container_width=True,type="primary",key="xls_"+area):
                        with st.spinner("Building Excel for {}...".format(area)):
                            try:
                                xls=build_outlet_excel(df,area)
                                st.download_button("💾 Download {} Excel".format(area),data=xls,
                                    file_name="AlMadina_"+area.replace(" ","_")+"_"+datetime.now().strftime("%Y%m%d_%H%M")+".xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,key="dl_xls_"+area)
                                st.success("✅ Includes: KPIs, daily performance, cancellations, cancelled items, timing, financials, hourly pattern")
                            except Exception as e:
                                st.error("Excel failed: "+str(e))
                with bc2:
                    if st.button("📄 HTML — {}".format(area),use_container_width=True,key="html_"+area):
                        with st.spinner("Building HTML for {}...".format(area)):
                            try:
                                html=build_html_report(df,[area],area+" Outlet Report")
                                st.download_button("💾 Download {} HTML".format(area),data=html,
                                    file_name="AlMadina_"+area.replace(" ","_")+"_"+datetime.now().strftime("%Y%m%d_%H%M")+".html",
                                    mime="text/html",use_container_width=True,key="dl_html_"+area)
                                st.info("💡 Open in browser → Print → Save as PDF",icon="💡")
                            except Exception as e:
                                st.error("HTML failed: "+str(e))
        st.markdown("</div>",unsafe_allow_html=True)
        st.divider()
        alert("Reports use <b>zero external dependencies</b> — Excel built with stdlib zipfile, HTML with pure Python. Nothing to install.","blue","ℹ️")

if __name__=="__main__":
    main()
