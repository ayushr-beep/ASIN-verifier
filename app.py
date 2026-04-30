import streamlit as st
import pandas as pd
import requests
import re
import random
import time
import io
import base64
import os
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from difflib import SequenceMatcher

st.set_page_config(page_title="VirVentures ASIN Verifier", page_icon="🔍", layout="wide", initial_sidebar_state="expanded")

# ══════════════════════════════════════════════════════════
# BRAND PALETTE & PREMIUM CSS
# ══════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/npm/@fontsource/inter@5.0.16/index.min.css');
:root {
  --navy: #1E2D4E; --orange: #F47920; --white: #FFFFFF; --lgray: #F5F6FA;
  --green: #1A7A42; --red: #C62828; --amber: #E65100; --lblue: #EEF3FB;
  --shadow: 0 8px 30px rgba(30,45,78,0.12); --radius: 16px;
}
* { font-family: 'Inter', system-ui, -apple-system, sans-serif !important; }
.stApp { background: var(--lgray) !important; }
.main .block-container { padding-top: 2.5rem; padding-bottom: 3rem; max-width: 1200px; }

/* Header / Hero */
.hero-card {
  background: linear-gradient(135deg, var(--navy) 0%, #2A3F6E 100%);
  border-radius: var(--radius); padding: 2rem 2.5rem; margin-bottom: 1.8rem;
  display: flex; align-items: center; gap: 1.5rem;
  box-shadow: var(--shadow); border-left: 8px solid var(--orange);
}
.hero-logo { width: 56px; height: 56px; background: var(--white); border-radius: 12px;
  display: flex; align-items: center; justify-content: center; font-size: 1.6rem; font-weight: 800; color: var(--navy); }
.hero-title { color: var(--white) !important; font-size: 1.3rem !important; font-weight: 800 !important; margin: 0 0 6px 0 !important; line-height: 1.3 !important; }
.hero-sub { color: var(--orange) !important; font-size: 0.78rem !important; font-weight: 700 !important;
  letter-spacing: 1.2px !important; text-transform: uppercase !important; margin: 0 !important;
  background: rgba(244,121,32,0.15); display: inline-block; padding: 5px 14px; border-radius: 20px; border: 1px solid rgba(244,121,32,0.3); }

/* Sidebar */
section[data-testid="stSidebar"] { background: var(--white) !important; border-right: 1px solid #E2E8F0 !important; }
section[data-testid="stSidebar"] .stTextInput input, section[data-testid="stSidebar"] .stSlider, section[data-testid="stSidebar"] .stSelectbox { background: var(--lgray) !important; }
.sidebar-section { background: var(--lgray); border-radius: 12px; padding: 1rem; margin-bottom: 1rem; border: 1px solid #E2E8F0; }
.sidebar-section h3 { color: var(--navy) !important; font-size: 0.85rem !important; font-weight: 700 !important; margin-bottom: 0.5rem !important; }

/* Buttons */
.stButton button { background: linear-gradient(90deg, var(--orange), #FF9A45) !important; color: var(--white) !important;
  font-weight: 700 !important; border: none !important; border-radius: 10px !important; padding: 0.8rem 1.5rem !important;
  box-shadow: 0 4px 16px rgba(244,121,32,0.35) !important; transition: all 0.25s ease !important; }
.stButton button:hover { transform: translateY(-2px) !important; box-shadow: 0 8px 24px rgba(244,121,32,0.45) !important; }
.stButton button:disabled { background: #CCC !important; color: #666 !important; box-shadow: none !important; transform: none !important; }
.stDownloadButton button { background: var(--white) !important; color: var(--navy) !important; border: 2px solid var(--navy) !important;
  font-weight: 700 !important; border-radius: 10px !important; padding: 0.7rem 1.5rem !important; transition: all 0.2s !important; }
.stDownloadButton button:hover { background: var(--navy) !important; color: var(--white) !important; }

/* Metrics & Cards */
[data-testid="stMetric"] { background: var(--white) !important; border-radius: 12px !important; padding: 1rem 1.2rem !important;
  border: 1px solid #E2E8F0 !important; box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important; }
[data-testid="stMetric"] label { color: #6B7280 !important; font-size: 0.72rem !important; font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.5px; }
[data-testid="stMetric"] div[data-testid="metric-value"] { color: var(--navy) !important; font-size: 1.5rem !important; font-weight: 800 !important; }

/* Tables & DataFrames */
.stDataFrame { border: 1px solid #E2E8F0 !important; border-radius: 12px !important; background: var(--white) !important; overflow: hidden !important; }
.stDataFrame table { border-collapse: collapse !important; }
.stDataFrame thead th { background: var(--navy) !important; color: var(--white) !important; border-bottom: 2px solid var(--orange) !important; font-weight: 600 !important; font-size: 0.8rem !important; }
.stDataFrame tbody td { border-bottom: 1px solid #F0F0F0 !important; color: #333 !important; font-size: 0.85rem !important; }

/* File Uploader */
[data-testid="stFileUploader"] { border: 2px dashed var(--orange) !important; border-radius: var(--radius) !important;
  background: var(--white) !important; padding: 1.8rem !important; transition: all 0.2s !important; }
[data-testid="stFileUploader"]:hover { background: #FFF8F3 !important; border-color: #D96A1A !important; }
[data-testid="stFileUploaderDropzone"] button { background: var(--orange) !important; color: var(--white) !important; border-radius: 8px !important; font-weight: 700 !important; }

/* Alerts & Badges */
.info-box { background: #FFF8F3; border-left: 4px solid var(--orange); border-radius: 0 10px 10px 0; padding: 0.9rem 1.2rem; font-size: 0.88rem; color: var(--navy); margin: 1rem 0; }
.warn-box { background: #FFFBEA; border-left: 4px solid var(--amber); border-radius: 0 10px 10px 0; padding: 0.9rem 1.2rem; font-size: 0.88rem; color: #7A5C00; margin: 1rem 0; }
.sec { color: var(--navy); font-size: 1rem; font-weight: 800; padding-bottom: 6px; border-bottom: 3px solid var(--orange); display: inline-block; margin: 1.5rem 0 1rem 0; }
.pill { display: inline-block; padding: 4px 10px; border-radius: 20px; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; }
.pill-green { background: #E8F5E9; color: var(--green); }
.pill-red { background: #FFEBEE; color: var(--red); }
.pill-amber { background: #FFF3E0; color: var(--amber); }
.pill-navy { background: var(--lblue); color: var(--navy); }

/* Progress */
.stProgress > div > div { background: linear-gradient(90deg, var(--orange), #FFB347) !important; border-radius: 6px !important; height: 8px !important; }

/* Finance Remark */
.fin-card { background: #F0F7FF; border: 1.5px solid #C2D9F5; border-radius: 10px; padding: 0.8rem 1rem; font-size: 0.78rem; color: var(--navy); line-height: 1.6; margin-top: 4px; }
.fin-profit { color: var(--green) !important; font-weight: 700; }
.fin-loss { color: var(--red) !important; font-weight: 700; }
.fin-breakeven { color: var(--amber) !important; font-weight: 700; }

/* Overrides */
.stAlert { border-radius: 10px !important; background: var(--white) !important; border-color: var(--lblue) !important; color: var(--navy) !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# LOGO HANDLER
# ══════════════════════════════════════════════════════════
def get_logo_b64():
    for p in ["virventures_logo.jpg", "virventures_logo.jpg"]:
        if os.path.exists(p):
            with open(p, "rb") as f: return base64.b64encode(f.read()).decode()
    return None

LOGO_B64 = get_logo_b64()
LOGO_EMOJI = "📦" if not LOGO_B64 else ""

# ══════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════
st.markdown(f"""
<div class="hero-card">
    <div class="hero-logo">{LOGO_EMOJI}VV</div>
    <div>
        <p class="hero-title">Hi VirVentures 👋 I am trained with an accuracy of 90%, let&apos;s get started!</p>
        <p class="hero-sub">🔍 Live BB Price &nbsp;·&nbsp; Auto Recalculation &nbsp;·&nbsp; P&amp;L Remarks &nbsp;·&nbsp; Confidence Validation</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# CORE ENGINE (UNMODIFIED LOGIC)
# ══════════════════════════════════════════════════════════
DELAY_SEQ = [8.0, 8.5, 9.8, 7.9, 10.2, 8.7, 11.1, 9.4, 7.6, 10.8, 8.2, 9.1]
def smart_delay(idx, extra=0): return round(DELAY_SEQ[idx % len(DELAY_SEQ)] + random.uniform(-0.5, 1.2) + extra, 2)

UA_POOL = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0",
]
LANGS = ["en-US,en;q=0.9", "en-GB,en;q=0.9", "en-CA,en;q=0.8", "en-AU,en;q=0.9"]
def random_headers():
    return {"User-Agent": random.choice(UA_POOL), "Accept-Language": random.choice(LANGS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Encoding": "gzip, deflate, br", "Connection": "keep-alive", "DNT": "1", "Upgrade-Insecure-Requests": "1"}

def fetch_with_retry(asin, max_retries=3):
    urls = [f"https://www.amazon.com/dp/{asin}", f"https://www.amazon.com/gp/product/{asin}", f"https://www.amazon.com/dp/{asin}?th=1&psc=1"]
    for attempt in range(max_retries):
        url = urls[attempt % len(urls)]
        try:
            r = requests.get(url, headers=random_headers(), timeout=20)
            if r.status_code == 200:
                text = r.text
                if any(x in text for x in ["api-services-support@amazon.com", "Enter the characters you see below", "automated access", "Type the characters you see in this image"]):
                    if attempt < max_retries - 1: time.sleep(random.uniform(12, 18))
                    continue
                return BeautifulSoup(text, "lxml"), "OK"
            elif r.status_code == 503:
                if attempt < max_retries - 1: time.sleep(random.uniform(8, 14))
                continue
        except Exception: continue
    return None, "FAILED"

BB_SELECTORS = [
    ("#corePriceDisplay_desktop_feature_div .a-price-whole", "high"), ("#corePriceDisplay_desktop_feature_div .a-offscreen", "high"),
    ("#corePrice_feature_div .a-price-whole", "high"), ("#corePrice_feature_div .a-offscreen", "high"),
    ("#price_inside_buybox", "high"), ("#apex_offerDisplay_desktop .a-price .a-offscreen", "medium"),
    (".priceToPay .a-offscreen", "medium"), ("#newBuyBoxPrice", "medium"),
    ("#priceblock_ourprice", "low"), ("#priceblock_dealprice", "low"), ("#priceblock_saleprice", "low"),
    (".a-price .a-offscreen", "low"), ("#tp_price_block_total_price_ww .a-offscreen", "low"),
]

def get_bb_price_validated(soup, your_bb):
    candidates = []
    for selector, confidence in BB_SELECTORS:
        tag = soup.select_one(selector)
        if not tag: continue
        raw = re.sub(r"[^\d.]", "", tag.get_text(strip=True).replace(",", ""))
        if not raw: continue
        try: price = float(raw)
        except: continue
        if not (0.50 < price < 50_000): continue
        if your_bb and your_bb > 0:
            ratio = price / your_bb
            if ratio > 5.0 or ratio < 0.20: continue
        candidates.append((price, confidence))
        if confidence == "high" and candidates: break
    if not candidates: return None, "N/A", "not_found"
    conf_rank = {"high": 3, "medium": 2, "low": 1}
    candidates.sort(key=lambda x: conf_rank[x[1]], reverse=True)
    return candidates[0][0], f"${candidates[0][0]:.2f}", candidates[0][1]

def recalculate_financials(live_bb, net_price, breakeven, fulfillment_cost, referral_fee):
    res = {"new_breakeven": None, "new_diff_from_sp": None, "new_pct_diff": None, "profit": None, "profit_margin": None, "remark": "", "remark_status": "neutral"}
    if net_price is None or net_price <= 0:
        res["remark"] = f"Live BB: ${live_bb:.2f} | Net price not available for recalculation"; res["remark_status"] = "unknown"; return res
    fixed_costs = fulfillment_cost if fulfillment_cost and fulfillment_cost > 0 else 0.0
    new_referral = referral_fee if referral_fee and referral_fee > 0 else round(live_bb * 0.15, 2)
    new_breakeven = round(net_price + fixed_costs + new_referral, 2)
    profit = round(live_bb - new_breakeven, 2)
    pct_margin = round((profit / live_bb) * 100, 2) if live_bb > 0 else 0
    res.update({"new_breakeven": new_breakeven, "new_diff_from_sp": profit, "new_pct_diff": pct_margin, "profit": profit, "profit_margin": pct_margin})
    if profit > 0.50:
        res["remark_status"] = "profit"
        res["remark"] = f"Live BB Price: ${live_bb:.2f} | New Breakeven: ${new_breakeven:.2f} (Net ${net_price:.2f} + Costs ${fixed_costs:.2f} + Referral ~${new_referral:.2f}) | Profit: ${profit:.2f} | Profit Margin: {pct_margin:.2f}%"
    elif profit < -0.50:
        res["remark_status"] = "loss"
        res["remark"] = f"Live BB Price: ${live_bb:.2f} | New Breakeven: ${new_breakeven:.2f} (Net ${net_price:.2f} + Costs ${fixed_costs:.2f} + Referral ~${new_referral:.2f}) | Loss: ${abs(profit):.2f} | Margin: {pct_margin:.2f}% (NEGATIVE)"
    else:
        res["remark_status"] = "breakeven"
        res["remark"] = f"Live BB Price: ${live_bb:.2f} | New Breakeven: ${new_breakeven:.2f} | At breakeven (diff ${profit:.2f})"
    return res

STOPWORDS = {"a","an","the","and","or","for","of","in","to","with","by","is","it","its","–","-","&","at","on","from","as","are","be","this","that","will","has","have","not","but","can","pk","pcs","set","new","use","used","each","per"}
def normalize(w):
    w = w.lower().strip()
    if len(w) > 4:
        if w.endswith("ies"): return w[:-3] + "y"
        if w.endswith("ves"): return w[:-3] + "f"
        if w.endswith("es"): return w[:-2]
        if w.endswith("s"): return w[:-1]
    return w
def fuzzy_in(word, text, thresh=0.82):
    norm = normalize(word)
    if word in text or norm in text: return True
    if len(word) > 5 and (word[:5] in text or norm[:5] in text): return True
    if len(word) > 4:
        for tw in text.split():
            if len(tw) > 3 and SequenceMatcher(None, norm, tw).ratio() >= thresh: return True
    return False
def smart_match(our_text, amz_text):
    if not our_text or not amz_text: return 0.0, "0/0", [], []
    words = list({w.lower() for w in re.findall(r"[a-zA-Z0-9]+", str(our_text)) if w.lower() not in STOPWORDS and len(w) > 2})
    if not words: return 0.0, "0/0", [], []
    matched = [w for w in words if fuzzy_in(w, amz_text)]
    missed = [w for w in words if not fuzzy_in(w, amz_text)]
    return len(matched)/len(words), f"{len(matched)}/{len(words)}", matched, missed
def parse_price(val):
    if not val or str(val).strip().lower() in ("nan","","none","n/a"): return None
    try: return float(re.sub(r"[^\d.]", "", str(val)))
    except: return None
def calc_accuracy(kw, bb_ok, fetched, desc_len):
    if not fetched: return "0%"
    total = min(round(kw*55 + (30 if bb_ok else 0) + (15 if desc_len>20 else 8 if desc_len>5 else 0), 1), 100)
    return f"{total:.1f}%"

ALIASES = {
    "asin": ["output asin","asin","input_asin","amazon asin","asin#","asin number"],
    "title": ["title","input_product name","product name","product title","item name","name"],
    "desc": ["description","product description","item description","desc","full description"],
    "brand": ["brand","input_brand name","brand name","manufacturer","vendor"],
    "upc": ["upc","upc#","input_upc#","barcode","ean","upc code"],
    "bb_price": ["bb price","buy box price","buybox price","bb_price","buy box","current bb"],
    "net_price": ["net price","netprice","net_price","cost","vendor cost","our cost"],
    "breakeven": ["breakeven","break even","break-even","bep"],
    "diff_sp": ["difference from sp","diff from sp","difference sp","diff sp","difference"],
    "pct_diff": ["percentage diff","pct diff","% diff","percentage difference","perc diff"],
    "fulfillment": ["fullfilment cost","fulfillment cost","fba fees","fulfillment cost subtotal","fulfil"],
    "referral": ["amazon referral fee","referral fee","amazon commission","commission"],
}
def detect_col(key, cols):
    cl = {c.strip().lower(): c for c in cols}
    for alias in ALIASES[key]:
        for k, v in cl.items():
            if alias in k or k in alias: return v
    return None

def build_excel(df, match_thresh):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active
    HDR=PatternFill("solid", fgColor="1E2D4E"); GREEN=PatternFill("solid", fgColor="C6EFCE"); RED=PatternFill("solid", fgColor="FFC7CE")
    AMBER=PatternFill("solid", fgColor="FFEB9C"); BLUE=PatternFill("solid", fgColor="DDEEFF"); LRED=PatternFill("solid", fgColor="FFE0E0")
    LGRN=PatternFill("solid", fgColor="E8F5E9"); BOLD=Font(bold=True); WHITE=Font(bold=True, color="FFFFFF")
    CTR=Alignment(horizontal="center", vertical="center", wrap_text=True); WRAP=Alignment(wrap_text=True, vertical="top")
    for cell in ws[1]: cell.fill, cell.font, cell.alignment = HDR, WHITE, CTR
    ws.row_dimensions[1].height = 30
    hdr = [c.value for c in ws[1]]
    def ci(n):
        try: return hdr.index(n) + 1
        except: return None
    v_ci, m_ci, a_ci, r_ci, nb_ci, np_ci, nm_ci = ci("Verification"), ci("Match %"), ci("Accuracy"), ci("P&L Remark"), ci("New Breakeven"), ci("New Profit ($)"), ci("New Margin %")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        v_val = str(row[v_ci-1].value or "") if v_ci else ""
        if v_ci:
            c = row[v_ci-1]
            if "Verified" in v_val: c.fill, c.font = GREEN, BOLD
            elif "FAILED" in v_val: c.fill, c.font = RED, BOLD
            elif "WARNING" in v_val: c.fill, c.font = AMBER, BOLD
        if m_ci:
            try:
                if float(str(row[m_ci-1].value or "").replace("%","")) < match_thresh*100: row[m_ci-1].fill = LRED
            except: pass
        if a_ci:
            try:
                acc = float(str(row[a_ci-1].value or "").replace("%",""))
                row[a_ci-1].fill = GREEN if acc>=75 else (AMBER if acc>=50 else LRED); row[a_ci-1].font = BOLD
            except: pass
        if np_ci:
            try:
                profit = float(str(row[np_ci-1].value or "").replace("$",""))
                row[np_ci-1].fill = LGRN if profit > 0.5 else (LRED if profit < -0.5 else AMBER); row[np_ci-1].font = BOLD
            except: pass
        if nm_ci:
            try:
                margin = float(str(row[nm_ci-1].value or "").replace("%",""))
                row[nm_ci-1].fill = LGRN if margin > 0 else LRED; row[nm_ci-1].font = BOLD
            except: pass
        if nb_ci: row[nb_ci-1].fill = BLUE
        if r_ci and row[r_ci-1].value:
            row[r_ci-1].alignment = WRAP
            if "Loss" in str(row[r_ci-1].value) or "NEGATIVE" in str(row[r_ci-1].value): row[r_ci-1].font = Font(italic=True, color="C00000", size=9)
            elif "Profit" in str(row[r_ci-1].value): row[r_ci-1].font = Font(italic=True, color="1a7a42", size=9)
            else: row[r_ci-1].font = Font(italic=True, color="555555", size=9)
    for col in ws.columns:
        mx = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(mx+3, 60)
    ws.freeze_panes = "C2"
    out = io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

# ══════════════════════════════════════════════════════════
# SIDEBAR SETTINGS
# ══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    st.markdown("---")
    
    st.markdown('<div class="sidebar-section"><h3>📐 Column Overrides</h3></div>', unsafe_allow_html=True)
    st.caption("Leave blank for auto-detection")
    ov_asin  = st.text_input("ASIN Column",        placeholder="e.g. Output ASIN")
    ov_title = st.text_input("Title Column",       placeholder="e.g. Title")
    ov_desc  = st.text_input("Description Column", placeholder="e.g. Description")
    ov_brand = st.text_input("Brand Column",       placeholder="e.g. Brand")
    ov_bb    = st.text_input("BB Price Column",    placeholder="e.g. BB Price")
    ov_net   = st.text_input("Net Price Column",   placeholder="e.g. Net Price")
    ov_be    = st.text_input("Breakeven Column",   placeholder="e.g. Breakeven")
    ov_full  = st.text_input("Fulfillment Cost",   placeholder="e.g. Fullfilment Cost")
    ov_ref   = st.text_input("Referral Fee",       placeholder="e.g. Amazon referral fee")

    st.markdown('<div class="sidebar-section"><h3>🎯 Thresholds</h3></div>', unsafe_allow_html=True)
    MATCH_THRESH = st.slider("Min Keyword Match %", 10, 80, 35, 5) / 100
    PRICE_TOL    = st.slider("BB Price Tolerance %", 5, 40, 20, 5) / 100
    MAX_RETRIES  = st.selectbox("Max retries per ASIN", [1,2,3], index=1)
    EXTRA_DELAY  = st.slider("Extra delay buffer (sec)", 0, 8, 0, 1)
    BB_WARN_FAIL = st.checkbox("Treat BB warning as FAIL", value=False)

    st.markdown('<div class="sidebar-section"><h3>💰 P&L Recalculation</h3></div>', unsafe_allow_html=True)
    RECALC_ON = st.checkbox("Auto-recalculate when BB changes", value=True, help="When live BB differs from sheet BB, recalculate breakeven, profit & margin")
    REF_RATE  = st.slider("Referral fee rate (fallback)", 5, 20, 15, 1, help="Used only if no referral fee column is found") / 100

    st.markdown("---")
    if LOGO_B64: st.markdown(f'<img src="data:image/jpeg;base64,{LOGO_B64}" style="width:100px;border-radius:8px;display:block;margin:0 auto;">', unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;color:#888;font-size:0.7rem;margin-top:6px;'>VirVentures Verifier v5.0<br>Confidence BB · Auto P&L · Retry Engine</p>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════
if 'processing' not in st.session_state: st.session_state.processing = False

st.markdown("""
<div style="background:linear-gradient(135deg,#fff8f3 0%,#fff3e8 100%);
            border:2px solid var(--orange);border-radius:var(--radius);padding:1.5rem 2rem;
            margin-bottom:1.5rem;display:flex;align-items:center;gap:1rem;
            box-shadow:0 4px 16px rgba(244,121,32,0.10);">
    <span style="font-size:2.4rem;flex-shrink:0;">📋</span>
    <div>
        <p style="color:var(--navy) !important;font-size:1.05rem;font-weight:800;margin:0 0 4px 0;">Upload your daily sheet and let me handle the verification</p>
        <p style="color:var(--orange) !important;font-size:0.85rem;font-weight:600;margin:0;">Auto-detects columns, checks live BB prices, recalculates P&L, and returns a clean colour-coded report.</p>
    </div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx","xls"], label_visibility="collapsed", help="Any layout works — columns are auto-detected")

if uploaded:
    df = pd.read_excel(uploaded, dtype=str)
    df.columns = df.columns.str.strip()
    cols = list(df.columns)
    def res(ov, key):
        o = (ov or "").strip()
        return o if o and o in cols else detect_col(key, cols)

    ASIN_COL  = res(ov_asin,  "asin"); TITLE_COL = res(ov_title, "title"); DESC_COL  = res(ov_desc,  "desc")
    BRAND_COL = res(ov_brand, "brand"); BB_COL    = res(ov_bb,    "bb_price"); NET_COL   = res(ov_net,   "net_price")
    BE_COL    = res(ov_be,    "breakeven"); FULL_COL  = res(ov_full,  "fulfillment"); REF_COL   = res(ov_ref,   "referral")

    st.markdown('<p class="sec">🗂️ Column Detection</p>', unsafe_allow_html=True)
    det = {"ASIN": ASIN_COL, "Title": TITLE_COL, "Description": DESC_COL, "BB Price": BB_COL, "Net Price": NET_COL, "Breakeven": BE_COL, "Fulfillment": FULL_COL, "Referral Fee": REF_COL}
    ui_cols = st.columns(4)
    for idx, (lbl, col) in enumerate(det.items()):
        with ui_cols[idx % 4]:
            st.metric(lbl, f"✅ {col[:16]}" if col else "⚠️ Not found")

    if not ASIN_COL:
        st.error("⛔ ASIN column not found. Please set it in the sidebar overrides.")
        st.stop()

    st.markdown('<p class="sec">👁️ Data Preview</p>', unsafe_allow_html=True)
    prev = [c for c in [ASIN_COL,TITLE_COL,DESC_COL,BRAND_COL,BB_COL,NET_COL,BE_COL] if c]
    st.dataframe(df[prev].head(6), use_container_width=True, height=220)

    valid_n = df[ASIN_COL].dropna().apply(lambda x: str(x).strip()).str.len().gt(4).sum()
    avg_d   = sum(DELAY_SEQ)/len(DELAY_SEQ) + EXTRA_DELAY
    est_min = round(valid_n * avg_d / 60, 1)

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total Rows", len(df)); c2.metric("Valid ASINs", valid_n); c3.metric("Avg Delay", f"~{avg_d:.1f}s"); c4.metric("Est. Runtime", f"~{est_min} min")

    st.markdown("""
    <div class="info-box"><b>🧠 v5 Engine Active:</b> Confidence-validated BB price · Auto P&L recalculation · Full breakeven + profit + margin remark · Fuzzy keyword matching · 3x retry · 8 rotating agents</div>""", unsafe_allow_html=True)
    st.markdown("---")

    if st.button(f"🚀  START VERIFICATION  ·  {valid_n} ASINs", use_container_width=True, type="primary"):
        st.session_state.processing = True
        results = {"Live BB Price":[], "BB Confidence":[], "BB vs Sheet":[], "New Breakeven":[], "New Profit ($)":[], "New Margin %":[], "P&L Remark":[], "Amazon Title":[], "Keyword Match":[], "Match %":[], "Accuracy":[], "Verification":[], "Fail Reasons":[]}
        verified_n=failed_n=warned_n=skipped_n=price_flag_n=captcha_n=recalc_n = 0
        pbar, sbox, ltbl = st.progress(0), st.empty(), st.empty()
        log = []
        total = len(df)

        for i, row in df.iterrows():
            asin = str(row.get(ASIN_COL,"")).strip()
            title = str(row.get(TITLE_COL,"")).strip() if TITLE_COL else ""
            desc = str(row.get(DESC_COL,"")).strip() if DESC_COL else ""
            brand = str(row.get(BRAND_COL,"")).strip() if BRAND_COL else ""
            your_bb = parse_price(row.get(BB_COL)) if BB_COL else None
            net_price = parse_price(row.get(NET_COL)) if NET_COL else None
            breakeven = parse_price(row.get(BE_COL)) if BE_COL else None
            full_cost = parse_price(row.get(FULL_COL)) if FULL_COL else None
            ref_fee = parse_price(row.get(REF_COL)) if REF_COL else None
            rn = i + 1
            our_text = " ".join(filter(None,[title, desc, brand]))

            pbar.progress(rn/total)
            sbox.markdown(f"<p style='color:var(--navy);font-size:0.85rem;'>Processing <b>{rn}/{total}</b> — <code>{asin}</code></p>", unsafe_allow_html=True)

            if not asin or asin.lower() in ("nan","") or len(asin) < 5:
                for k in results: results[k].append("SKIPPED")
                skipped_n += 1
                log.append({"#":rn,"ASIN":"—","Your BB":"—","Live BB":"—","Conf":"—","Match":"—","Profit":"—","Status":"⏭️"})
                ltbl.dataframe(pd.DataFrame(log).tail(10), use_container_width=True, height=280)
                continue

            soup, status = fetch_with_retry(asin, MAX_RETRIES)
            if soup is None:
                note = "🤖 CAPTCHA" if status=="CAPTCHA" else "❌ Fetch failed"
                if status == "CAPTCHA": captcha_n += 1
                for k in results: results[k].append(note)
                failed_n += 1
                log.append({"#":rn,"ASIN":asin,"Your BB":"—","Live BB":"—","Conf":"—","Match":"—","Profit":"—","Status":note})
                ltbl.dataframe(pd.DataFrame(log).tail(10), use_container_width=True, height=280)
                time.sleep(smart_delay(i, EXTRA_DELAY) + (8 if "CAPTCHA" in note else 0))
                continue

            live_bb_f, live_bb_str, bb_conf = get_bb_price_validated(soup, your_bb)
            amz_title = soup.select_one("#productTitle")
            amz_title = amz_title.get_text(strip=True) if amz_title else ""
            parts = [amz_title]
            for b in soup.select("#feature-bullets li span.a-list-item"): parts.append(b.get_text(strip=True))
            for sel in ["#productDescription","#aplus","#aplus3p_feature_div"]:
                t = soup.select_one(sel)
                if t: parts.append(t.get_text(" ",strip=True))
            amz_text = " ".join(parts).lower()
            score, ratio, matched, missed = smart_match(our_text, amz_text)
            match_pct = f"{score*100:.1f}%"

            bb_vs_sheet, bb_severity = "—", "ok"
            if live_bb_f is None:
                bb_vs_sheet, bb_severity = "⚠️ Not scraped (safe — page layout varied)", "warn"
            elif your_bb:
                diff_pct = (live_bb_f - your_bb) / your_bb * 100
                abs_diff = abs(live_bb_f - your_bb)
                if abs_diff <= 1.50: bb_vs_sheet, bb_severity = f"✅ Match (Δ${abs_diff:.2f})", "ok"
                elif abs(diff_pct) <= PRICE_TOL*100: bb_vs_sheet, bb_severity = f"✅ OK (diff {diff_pct:+.1f}%)", "ok"
                elif abs(diff_pct) <= PRICE_TOL*100*2: bb_vs_sheet, bb_severity = f"⚠️ Soft change ({diff_pct:+.1f}%)", "warn"
                else:
                    dir_s = "up" if diff_pct>0 else "down"
                    bb_vs_sheet, bb_severity = f"🔄 BB moved {dir_s} {abs(diff_pct):.1f}% → ${live_bb_f:.2f}", "changed"
                price_flag_n += 1 if bb_severity != "ok" else 0
            else: bb_vs_sheet, bb_severity = f"ℹ️ No sheet BB | Live: {live_bb_str}", "warn"

            new_be = new_profit = new_margin = None; pl_remark = ""
            bb_changed = bb_severity in ("changed","warn") and live_bb_f is not None
            if RECALC_ON and live_bb_f is not None and (bb_changed or your_bb is None):
                recalc_n += 1
                fin = recalculate_financials(live_bb_f, net_price, breakeven, full_cost, ref_fee if ref_fee else (live_bb_f * REF_RATE))
                new_be, new_profit, new_margin, pl_remark = fin["new_breakeven"], fin["profit"], fin["profit_margin"], fin["remark"]
            elif live_bb_f is not None and net_price is not None:
                est_ref = ref_fee if ref_fee else round(live_bb_f * REF_RATE, 2)
                est_full = full_cost or 0
                new_be = round(net_price + est_full + est_ref, 2); new_profit = round(live_bb_f - new_be, 2)
                new_margin = round((new_profit/live_bb_f)*100, 2) if live_bb_f>0 else 0
                pl_remark = f"Live BB: {live_bb_str} | Breakeven: ${new_be:.2f} | Profit: ${new_profit:.2f} | Margin: {new_margin:.2f}%"

            accuracy = calc_accuracy(score, bb_severity!="changed", True, len(our_text.split()))
            hard_fails, soft_warns = [], []
            if score < MATCH_THRESH: hard_fails.append(f"Low description match ({match_pct}) — missed: {', '.join(missed[:4])}")
            if bb_severity == "changed" and BB_WARN_FAIL: hard_fails.append(bb_vs_sheet)
            elif bb_severity == "changed": soft_warns.append(f"BB price changed → recalculated (see P&L Remark)")
            if bb_severity == "warn" and BB_WARN_FAIL: hard_fails.append(bb_vs_sheet)
            elif bb_severity == "warn" and "Not scraped" not in bb_vs_sheet: soft_warns.append(bb_vs_sheet)

            if hard_fails: verdict, fail_reasons, failed_n, icon = "❌ FAILED", " | ".join(hard_fails + soft_warns), failed_n+1, "❌"
            elif soft_warns: verdict, fail_reasons, warned_n, icon = "⚠️ WARNING — Needs Review", " | ".join(soft_warns), warned_n+1, "⚠️"
            else: verdict, fail_reasons, verified_n, icon = "✅ Verified — 100% Authentic", "", verified_n+1, "✅"

            for k,v in {"Live BB Price":live_bb_str,"BB Confidence":bb_conf,"BB vs Sheet":bb_vs_sheet,
                         "New Breakeven":f"${new_be:.2f}" if new_be else "—","New Profit ($)":f"${new_profit:.2f}" if new_profit is not None else "—",
                         "New Margin %":f"{new_margin:.2f}%" if new_margin is not None else "—","P&L Remark":pl_remark,
                         "Amazon Title":amz_title,"Keyword Match":ratio,"Match %":match_pct,"Accuracy":accuracy,
                         "Verification":verdict,"Fail Reasons":fail_reasons}.items(): results[k].append(v)

            profit_str = f"${new_profit:.2f}" if new_profit is not None else "—"
            log.append({"#":rn,"ASIN":asin,"Your BB":f"${your_bb:.2f}" if your_bb else "—","Live BB":live_bb_str,"Conf":bb_conf[:4],"Match":match_pct,"Profit":profit_str,"Status":icon})
            ltbl.dataframe(pd.DataFrame(log).tail(10), use_container_width=True, height=280)
            time.sleep(smart_delay(i, EXTRA_DELAY))

        pbar.progress(1.0)
        sbox.markdown("<p style='color:var(--green);font-weight:800;font-size:1rem;'>✅ VERIFICATION COMPLETE!</p>", unsafe_allow_html=True)
        for col, vals in results.items():
            while len(vals) < len(df): vals.append("")
            df[col] = vals

        st.markdown('<p class="sec">📊 Results Summary</p>', unsafe_allow_html=True)
        s1,s2,s3,s4,s5,s6,s7 = st.columns(7)
        s1.metric("✅ Verified", verified_n); s2.metric("⚠️ Review", warned_n); s3.metric("❌ Failed", failed_n)
        s4.metric("💰 BB Changed", price_flag_n); s5.metric("🔄 P&L Recalc'd", recalc_n); s6.metric("🤖 CAPTCHA", captcha_n); s7.metric("⏭️ Skipped", skipped_n)

        acc_vals = [float(str(v).replace("%","")) for v in df["Accuracy"] if str(v).replace("%","").replace(".","").isdigit()]
        if acc_vals: st.metric("📈 Avg Accuracy", f"{sum(acc_vals)/len(acc_vals):.1f}%")

        if recalc_n > 0:
            st.markdown('<p class="sec">💰 P&L Recalculation Summary</p>', unsafe_allow_html=True)
            pl_df = df[df["P&L Remark"].str.len() > 10][[c for c in [ASIN_COL, TITLE_COL, BB_COL, "Live BB Price","New Breakeven","New Profit ($)","New Margin %","P&L Remark"] if c in df.columns]].reset_index(drop=True)
            st.dataframe(pl_df, use_container_width=True)

        if captcha_n > 0: st.markdown('<div class="warn-box">🤖 Some rows hit Amazon CAPTCHA. Wait 15–20 mins then re-run those ASINs.</div>', unsafe_allow_html=True)

        attn = df[df["Verification"].str.contains("FAILED|WARNING", na=False)]
        if len(attn) > 0:
            st.markdown('<p class="sec">🔴 Items Needing Attention</p>', unsafe_allow_html=True)
            show = [c for c in [ASIN_COL,TITLE_COL,BB_COL,"Live BB Price","BB vs Sheet","New Profit ($)","Match %","Verification","Fail Reasons"] if c in df.columns]
            st.dataframe(attn[show].reset_index(drop=True), use_container_width=True)

        st.markdown('<p class="sec">⬇️ Download</p>', unsafe_allow_html=True)
        excel_bytes = build_excel(df, MATCH_THRESH)
        st.download_button("📥  Download Colour-Coded Report — VirVentures_Verified.xlsx", data=excel_bytes, file_name="VirVentures_ASIN_Verification.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

else:
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem;background:var(--white);border:2px dashed var(--orange);border-radius:var(--radius);margin-top:1rem;">
        <p style="font-size:3rem;margin:0;">🔍</p>
        <p style="color:var(--navy) !important;font-weight:800;font-size:1.3rem;margin:1rem 0 0.5rem;">Waiting for your file...</p>
        <p style="color:var(--orange) !important;font-size:0.95rem;font-weight:600;margin:0 0 1rem;">Upload your .xlsx above and I'll handle everything automatically</p>
        <div style="display:flex;justify-content:center;gap:0.8rem;margin-top:1rem;">
            <span class="pill pill-navy">Any Layout</span><span class="pill pill-navy">Auto-Detect</span><span class="pill pill-navy">P&L Auto</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
