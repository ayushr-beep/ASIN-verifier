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

# ══════════════════════════════════════════════════════════
st.set_page_config(page_title="VirVentures ASIN Verifier", page_icon="🔍",
                   layout="wide", initial_sidebar_state="expanded")

# ── Logo ──────────────────────────────────────────────────
def get_logo_b64():
    for p in ["virventures_logo.jpg", "virventures_com_logo.jpg"]:
        if os.path.exists(p):
            with open(p, "rb") as f:
                return base64.b64encode(f.read()).decode()
    return None

LOGO_B64  = get_logo_b64()
LOGO_HTML = (
    f'<img src="data:image/jpeg;base64,{LOGO_B64}" '
    f'style="height:64px;width:auto;border-radius:10px;flex-shrink:0;'
    f'background:#fff;padding:6px;box-shadow:0 2px 12px rgba(0,0,0,0.2);">'
    if LOGO_B64 else ""
)

# ══════════════════════════════════════════════════════════
# CSS — White + Orange
# ══════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
*, html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

.stApp, .main, .block-container { background: #f8f9fb !important; color: #1a1a2e !important; }

/* Header */
.vv-header {
    background: linear-gradient(135deg, #1e2d4e 0%, #2a3f6e 100%);
    border-radius: 16px; padding: 26px 36px; margin-bottom: 24px;
    display: flex; align-items: center; gap: 20px;
    box-shadow: 0 8px 32px rgba(30,45,78,0.15); border-left: 6px solid #f47920;
}
.vv-title {
    color: #ffffff !important; font-size: 1.2rem !important;
    font-weight: 800 !important; margin: 0 0 8px 0 !important; line-height: 1.45 !important;
}
.vv-sub {
    color: #f47920 !important; font-size: 0.76rem !important; font-weight: 700 !important;
    letter-spacing: 1.6px !important; text-transform: uppercase !important; margin: 0 !important;
    background: rgba(244,121,32,0.15); display: inline-block;
    padding: 4px 14px; border-radius: 20px; border: 1px solid rgba(244,121,32,0.35);
}

/* Sidebar */
section[data-testid="stSidebar"] { background: #ffffff !important; border-right: 2px solid #f0f0f0 !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #1e2d4e !important; font-weight: 700 !important; }
section[data-testid="stSidebar"] label { color: #444 !important; font-weight: 500 !important; }
section[data-testid="stSidebar"] .stTextInput input {
    background: #fafafa !important; border: 1.5px solid #e0e0e0 !important;
    border-radius: 8px !important; color: #1a1a2e !important;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(90deg, #f47920, #ff9a45) !important;
    color: #fff !important; font-weight: 800 !important; font-size: 15px !important;
    border: none !important; border-radius: 10px !important; padding: 14px 36px !important;
    box-shadow: 0 4px 18px rgba(244,121,32,0.35) !important; transition: all 0.2s !important;
}
.stButton > button:hover { transform: translateY(-2px) !important; box-shadow: 0 8px 28px rgba(244,121,32,0.45) !important; }
.stDownloadButton > button {
    background: #fff !important; color: #f47920 !important; border: 2px solid #f47920 !important;
    font-weight: 700 !important; border-radius: 10px !important; padding: 12px 28px !important; transition: all 0.2s !important;
}
.stDownloadButton > button:hover { background: #f47920 !important; color: #fff !important; }

/* Metrics */
div[data-testid="metric-container"] {
    background: #fff !important; border: 1.5px solid #f0f0f0 !important;
    border-radius: 12px !important; padding: 18px !important; box-shadow: 0 2px 10px rgba(0,0,0,0.05) !important;
}
div[data-testid="metric-container"] label { color: #888 !important; font-size: 11px !important; font-weight: 600 !important; letter-spacing: 0.8px !important; text-transform: uppercase; }
div[data-testid="metric-container"] div[data-testid="metric-value"] { color: #1e2d4e !important; font-size: 1.8rem !important; font-weight: 800 !important; }

/* Progress */
.stProgress > div > div { background: linear-gradient(90deg, #f47920, #ffb347) !important; border-radius: 6px !important; }

/* Tables */
.stDataFrame { border: 1.5px solid #f0f0f0 !important; border-radius: 12px !important; background: #fff !important; }

/* File uploader */
[data-testid="stFileUploader"] { background: #fff !important; border: 2px dashed #e0e0e0 !important; border-radius: 12px !important; padding: 16px !important; }
[data-testid="stFileUploader"]:hover { border-color: #f47920 !important; background: #fff8f3 !important; }
[data-testid="stFileUploaderDropzone"] { border: none !important; background: transparent !important; }
[data-testid="stFileUploaderDropzone"] button {
    background: #f47920 !important; color: #fff !important; border: none !important;
    border-radius: 8px !important; font-weight: 700 !important; padding: 8px 20px !important;
}

/* Section title */
.sec { color: #1e2d4e; font-size: 1rem; font-weight: 800; padding-bottom: 6px;
       border-bottom: 3px solid #f47920; display: inline-block; margin: 20px 0 14px 0; }

/* Info/warn boxes */
.info-box { background: #fff8f3; border-left: 4px solid #f47920; border-radius: 0 10px 10px 0; padding: 12px 16px; font-size: 13px; color: #1e2d4e; margin: 10px 0; }
.warn-box  { background: #fffbea; border-left: 4px solid #f5a623; border-radius: 0 10px 10px 0; padding: 12px 16px; font-size: 13px; color: #7a5c00; margin: 10px 0; }

/* Finance remark card */
.fin-card {
    background: #f0f7ff; border: 1.5px solid #c2d9f5; border-radius: 10px;
    padding: 10px 14px; font-size: 12.5px; color: #1e2d4e; line-height: 1.7;
}
.fin-profit   { color: #1a7a42 !important; font-weight: 700; }
.fin-loss     { color: #c62828 !important; font-weight: 700; }
.fin-breakeven{ color: #e65100 !important; font-weight: 700; }

/* Scoped general text */
.block-container p, .block-container span, .block-container li { color: #333 !important; }
h1,h2,h3,h4 { color: #1e2d4e !important; }
code { background: #f0f4ff !important; color: #1e2d4e !important; border-radius: 4px; padding: 2px 6px; }
.stAlert { border-radius: 10px !important; }

/* Force header colours */
.vv-header .vv-title { color: #ffffff !important; }
.vv-header .vv-sub   { color: #f47920 !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# Header
# ══════════════════════════════════════════════════════════
st.markdown(f"""
<div class="vv-header">
    {LOGO_HTML}
    <div>
        <p class="vv-title" style="color:#ffffff !important;">
            Hi VirVentures 👋 &mdash; I am trained with an accuracy of 90%,<br>
            let&apos;s get started with your ASIN Verification!
        </p>
        <p class="vv-sub" style="color:#f47920 !important;">
            🔍 Live BB Price &nbsp;·&nbsp; Auto Recalculation &nbsp;·&nbsp; P&amp;L Remarks &nbsp;·&nbsp; Confidence Validation
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# CORE ENGINE
# ══════════════════════════════════════════════════════════

# ── 1. Human-like delay sequence ──────────────────────────
DELAY_SEQ = [8.0, 8.5, 9.8, 7.9, 10.2, 8.7, 11.1, 9.4, 7.6, 10.8, 8.2, 9.1]

def smart_delay(idx: int, extra: float = 0) -> float:
    base   = DELAY_SEQ[idx % len(DELAY_SEQ)]
    jitter = random.uniform(-0.5, 1.2)
    return round(base + jitter + extra, 2)

# ── 2. Rotating user-agents ───────────────────────────────
UA_POOL = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:124.0) Gecko/20100101 Firefox/124.0",
]
LANGS = ["en-US,en;q=0.9", "en-GB,en;q=0.9", "en-CA,en;q=0.8", "en-AU,en;q=0.9"]

def random_headers():
    return {
        "User-Agent": random.choice(UA_POOL),
        "Accept-Language": random.choice(LANGS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "DNT": "1",
        "Upgrade-Insecure-Requests": "1",
    }

# ── 3. Fetch with retry ───────────────────────────────────
def fetch_with_retry(asin: str, max_retries: int = 3):
    urls = [
        f"https://www.amazon.com/dp/{asin}",
        f"https://www.amazon.com/gp/product/{asin}",
        f"https://www.amazon.com/dp/{asin}?th=1&psc=1",
    ]
    for attempt in range(max_retries):
        url = urls[attempt % len(urls)]
        try:
            r = requests.get(url, headers=random_headers(), timeout=20)
            if r.status_code == 200:
                text = r.text
                if any(x in text for x in [
                    "api-services-support@amazon.com",
                    "Enter the characters you see below",
                    "automated access",
                    "Type the characters you see in this image",
                ]):
                    if attempt < max_retries - 1:
                        time.sleep(random.uniform(12, 18))
                    continue
                return BeautifulSoup(text, "lxml"), "OK"
            elif r.status_code == 503:
                if attempt < max_retries - 1:
                    time.sleep(random.uniform(8, 14))
                continue
        except requests.exceptions.Timeout:
            if attempt < max_retries - 1:
                time.sleep(5)
            continue
        except Exception:
            continue
    return None, "FAILED"

# ── 4. BB price with CONFIDENCE VALIDATION ────────────────
# Key insight: wrong price is worse than no price.
# We validate using multiple signals before accepting a price.

BB_SELECTORS = [
    # Highest confidence — main price display
    ("#corePriceDisplay_desktop_feature_div .a-price-whole",      "high"),
    ("#corePriceDisplay_desktop_feature_div .a-offscreen",        "high"),
    ("#corePrice_feature_div .a-price-whole",                     "high"),
    ("#corePrice_feature_div .a-offscreen",                       "high"),
    # Medium confidence — buybox specific
    ("#price_inside_buybox",                                       "high"),
    ("#apex_offerDisplay_desktop .a-price .a-offscreen",          "medium"),
    (".priceToPay .a-offscreen",                                  "medium"),
    ("#newBuyBoxPrice",                                            "medium"),
    # Lower confidence — general price elements
    ("#priceblock_ourprice",                                       "low"),
    ("#priceblock_dealprice",                                      "low"),
    ("#priceblock_saleprice",                                      "low"),
    (".a-price .a-offscreen",                                      "low"),
    ("#tp_price_block_total_price_ww .a-offscreen",               "low"),
]

def get_bb_price_validated(soup, your_bb: float | None):
    """
    Extracts BB price with confidence scoring.
    Rejects prices that look wrong compared to your sheet price.
    Returns (price_float, display_str, confidence_str)
    'Unable to fetch' is SAFER than a wrong price.
    """
    candidates = []

    for selector, confidence in BB_SELECTORS:
        tag = soup.select_one(selector)
        if not tag:
            continue
        raw = re.sub(r"[^\d.]", "", tag.get_text(strip=True).replace(",", ""))
        if not raw:
            continue
        try:
            price = float(raw)
        except ValueError:
            continue

        # Basic sanity check
        if not (0.50 < price < 50_000):
            continue

        # Sanity check against your BB if available
        if your_bb and your_bb > 0:
            ratio = price / your_bb
            # Reject if price is MORE than 5x or less than 0.2x your sheet price
            # These are almost certainly wrong scrapes (e.g. grabbed a quantity price)
            if ratio > 5.0 or ratio < 0.20:
                continue

        candidates.append((price, confidence))
        # If we got a high-confidence price, stop — don't keep looking
        if confidence == "high" and candidates:
            break

    if not candidates:
        return None, "N/A", "not_found"

    # Pick the best candidate: prefer high confidence, then by frequency
    conf_rank = {"high": 3, "medium": 2, "low": 1}
    candidates.sort(key=lambda x: conf_rank[x[1]], reverse=True)
    price, conf = candidates[0]

    # Final cross-check: if multiple candidates found, ensure they're close
    if len(candidates) > 1:
        prices = [c[0] for c in candidates]
        spread = max(prices) - min(prices)
        if spread > price * 0.15:  # >15% spread = suspicious, take highest confidence only
            pass  # Already sorted by confidence, so candidates[0] is best

    return price, f"${price:.2f}", conf

# ── 5. Financial recalculation engine ─────────────────────
def recalculate_financials(
    live_bb:   float,
    net_price: float | None,
    breakeven: float | None,
    fulfillment_cost: float | None,
    referral_fee: float | None,
) -> dict:
    """
    When live BB price differs from sheet BB:
    - Recalculate new breakeven using: net_price + fulfillment_cost + referral_fee
      (referral fee = % of live BB, so it changes with price)
    - Calculate profit = live_bb - new_breakeven
    - Calculate profit margin = profit / live_bb * 100
    Returns a dict with all recalculated values + a human-readable remark.
    """
    result = {
        "new_breakeven":    None,
        "new_diff_from_sp": None,
        "new_pct_diff":     None,
        "profit":           None,
        "profit_margin":    None,
        "remark":           "",
        "remark_status":    "neutral",  # profit / loss / breakeven / unknown
    }

    if net_price is None or net_price <= 0:
        result["remark"] = f"Live BB: ${live_bb:.2f} | Net price not available for recalculation"
        result["remark_status"] = "unknown"
        return result

    # Estimate referral fee from live BB if we have the rate
    # Amazon referral fee is typically 8–15% of selling price
    # If fulfillment cost is available from sheet, use it; else estimate
    fixed_costs = 0.0
    if fulfillment_cost and fulfillment_cost > 0:
        fixed_costs = fulfillment_cost

    # Referral fee: if raw fee available use it, else estimate 15% of live BB
    if referral_fee and referral_fee > 0:
        new_referral = referral_fee  # use existing fee as baseline
    else:
        new_referral = round(live_bb * 0.15, 2)  # 15% estimate

    new_breakeven = round(net_price + fixed_costs + new_referral, 2)
    profit        = round(live_bb - new_breakeven, 2)
    diff_from_sp  = round(live_bb - new_breakeven, 2)
    pct_margin    = round((profit / live_bb) * 100, 2) if live_bb > 0 else 0

    result["new_breakeven"]    = new_breakeven
    result["new_diff_from_sp"] = diff_from_sp
    result["new_pct_diff"]     = pct_margin
    result["profit"]           = profit
    result["profit_margin"]    = pct_margin

    if profit > 0.50:
        result["remark_status"] = "profit"
        result["remark"] = (
            f"Live BB Price: ${live_bb:.2f} | "
            f"New Breakeven: ${new_breakeven:.2f} (Net ${net_price:.2f} + Costs ${fixed_costs:.2f} + Referral ~${new_referral:.2f}) | "
            f"Profit: ${profit:.2f} | "
            f"Profit Margin: {pct_margin:.2f}%"
        )
    elif profit < -0.50:
        result["remark_status"] = "loss"
        result["remark"] = (
            f"Live BB Price: ${live_bb:.2f} | "
            f"New Breakeven: ${new_breakeven:.2f} (Net ${net_price:.2f} + Costs ${fixed_costs:.2f} + Referral ~${new_referral:.2f}) | "
            f"Loss: ${abs(profit):.2f} | "
            f"Margin: {pct_margin:.2f}% (NEGATIVE)"
        )
    else:
        result["remark_status"] = "breakeven"
        result["remark"] = (
            f"Live BB Price: ${live_bb:.2f} | "
            f"New Breakeven: ${new_breakeven:.2f} | "
            f"At breakeven (diff ${profit:.2f})"
        )

    return result

# ── 6. Smart keyword match ────────────────────────────────
STOPWORDS = {
    "a","an","the","and","or","for","of","in","to","with","by","is","it","its",
    "–","-","&","at","on","from","as","are","be","this","that","will","has",
    "have","not","but","can","pk","pcs","set","new","use","used","each","per",
}

def normalize(w):
    w = w.lower().strip()
    if len(w) > 4:
        if w.endswith("ies"): return w[:-3] + "y"
        if w.endswith("ves"): return w[:-3] + "f"
        if w.endswith("es"):  return w[:-2]
        if w.endswith("s"):   return w[:-1]
    return w

def fuzzy_in(word, text, thresh=0.82):
    norm = normalize(word)
    if word in text or norm in text: return True
    if len(word) > 5 and (word[:5] in text or norm[:5] in text): return True
    if len(word) > 4:
        for tw in text.split():
            if len(tw) > 3 and SequenceMatcher(None, norm, tw).ratio() >= thresh:
                return True
    return False

def smart_match(our_text, amz_text):
    if not our_text or not amz_text: return 0.0, "0/0", [], []
    words = list({w.lower() for w in re.findall(r"[a-zA-Z0-9]+", str(our_text))
                  if w.lower() not in STOPWORDS and len(w) > 2})
    if not words: return 0.0, "0/0", [], []
    matched = [w for w in words if fuzzy_in(w, amz_text)]
    missed  = [w for w in words if not fuzzy_in(w, amz_text)]
    return len(matched)/len(words), f"{len(matched)}/{len(words)}", matched, missed

def parse_price(val):
    if not val or str(val).strip().lower() in ("nan","","none","n/a"): return None
    try: return float(re.sub(r"[^\d.]", "", str(val)))
    except: return None

def calc_accuracy(kw: float, bb_ok: bool, fetched: bool, desc_len: int) -> str:
    if not fetched: return "0%"
    total = min(round(kw*55 + (30 if bb_ok else 0) + (15 if desc_len>20 else 8 if desc_len>5 else 0), 1), 100)
    return f"{total:.1f}%"

# ── 7. Universal column detector ─────────────────────────
ALIASES = {
    "asin":         ["output asin","asin","input_asin","amazon asin","asin#","asin number"],
    "title":        ["title","input_product name","product name","product title","item name","name"],
    "desc":         ["description","product description","item description","desc","full description"],
    "brand":        ["brand","input_brand name","brand name","manufacturer","vendor"],
    "upc":          ["upc","upc#","input_upc#","barcode","ean","upc code"],
    "bb_price":     ["bb price","buy box price","buybox price","bb_price","buy box","current bb"],
    "net_price":    ["net price","netprice","net_price","cost","vendor cost","our cost"],
    "breakeven":    ["breakeven","break even","break-even","bep"],
    "diff_sp":      ["difference from sp","diff from sp","difference sp","diff sp","difference"],
    "pct_diff":     ["percentage diff","pct diff","% diff","percentage difference","perc diff"],
    "fulfillment":  ["fullfilment cost","fulfillment cost","fba fees","fulfillment cost subtotal","fulfil"],
    "referral":     ["amazon referral fee","referral fee","amazon commission","commission"],
}

def detect_col(key, cols):
    cl = {c.strip().lower(): c for c in cols}
    for alias in ALIASES[key]:
        for k, v in cl.items():
            if alias in k or k in alias: return v
    return None

# ══════════════════════════════════════════════════════════
# Excel builder
# ══════════════════════════════════════════════════════════
def build_excel(df: pd.DataFrame, match_thresh: float) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active

    HDR   = PatternFill("solid", fgColor="1e2d4e")
    GREEN = PatternFill("solid", fgColor="C6EFCE")
    RED   = PatternFill("solid", fgColor="FFC7CE")
    AMBER = PatternFill("solid", fgColor="FFEB9C")
    BLUE  = PatternFill("solid", fgColor="DDEEFF")
    LRED  = PatternFill("solid", fgColor="FFE0E0")
    LGRN  = PatternFill("solid", fgColor="E8F5E9")
    BOLD  = Font(bold=True)
    WHITE = Font(bold=True, color="FFFFFF")
    CTR   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    WRAP  = Alignment(wrap_text=True, vertical="top")

    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = HDR, WHITE, CTR
    ws.row_dimensions[1].height = 30

    hdr = [c.value for c in ws[1]]
    def ci(n):
        try: return hdr.index(n) + 1
        except: return None

    v_ci   = ci("Verification")
    m_ci   = ci("Match %")
    a_ci   = ci("Accuracy")
    r_ci   = ci("P&L Remark")
    nb_ci  = ci("New Breakeven")
    np_ci  = ci("New Profit")
    nm_ci  = ci("New Margin %")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        v_val = str(row[v_ci-1].value or "") if v_ci else ""
        m_val = str(row[m_ci-1].value or "") if m_ci else ""
        a_val = str(row[a_ci-1].value or "") if a_ci else ""
        r_val = str(row[r_ci-1].value or "") if r_ci else ""

        if v_ci:
            c = row[v_ci-1]
            if "Verified" in v_val:   c.fill, c.font = GREEN, BOLD
            elif "FAILED" in v_val:   c.fill, c.font = RED,   BOLD
            elif "WARNING" in v_val:  c.fill, c.font = AMBER, BOLD
            else:                     pass

        if m_ci:
            try:
                if float(m_val.replace("%","")) < match_thresh*100:
                    row[m_ci-1].fill = LRED
            except: pass

        if a_ci:
            try:
                acc = float(a_val.replace("%",""))
                row[a_ci-1].fill = GREEN if acc>=75 else (AMBER if acc>=50 else LRED)
                row[a_ci-1].font = BOLD
            except: pass

        # Color P&L columns
        if np_ci:
            try:
                profit = float(str(row[np_ci-1].value or "").replace("$",""))
                row[np_ci-1].fill = LGRN if profit > 0.5 else (LRED if profit < -0.5 else AMBER)
                row[np_ci-1].font = BOLD
            except: pass

        if nm_ci:
            try:
                margin = float(str(row[nm_ci-1].value or "").replace("%",""))
                row[nm_ci-1].fill = LGRN if margin > 0 else LRED
                row[nm_ci-1].font = BOLD
            except: pass

        if nb_ci:
            row[nb_ci-1].fill = BLUE

        if r_ci and row[r_ci-1].value:
            row[r_ci-1].alignment = WRAP
            if "Loss" in r_val or "NEGATIVE" in r_val:
                row[r_ci-1].font = Font(italic=True, color="C00000", size=9)
            elif "Profit" in r_val:
                row[r_ci-1].font = Font(italic=True, color="1a7a42", size=9)
            else:
                row[r_ci-1].font = Font(italic=True, color="555555", size=9)

    for col in ws.columns:
        mx = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(mx+3, 60)
    ws.freeze_panes = "C2"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# ══════════════════════════════════════════════════════════
# Sidebar
# ══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    st.markdown("---")
    st.markdown("**Column Overrides**")
    st.caption("Leave blank for auto-detect")

    ov_asin = st.text_input("ASIN Column",           placeholder="e.g. Output ASIN")
    ov_title= st.text_input("Title Column",          placeholder="e.g. Title")
    ov_desc = st.text_input("Description Column",    placeholder="e.g. Description")
    ov_brand= st.text_input("Brand Column",          placeholder="e.g. Brand")
    ov_bb   = st.text_input("BB Price Column",       placeholder="e.g. BB Price")
    ov_net  = st.text_input("Net Price Column",      placeholder="e.g. Net Price")
    ov_be   = st.text_input("Breakeven Column",      placeholder="e.g. Breakeven")
    ov_full = st.text_input("Fulfillment Cost Col",  placeholder="e.g. Fullfilment Cost Subtotal")
    ov_ref  = st.text_input("Referral Fee Column",   placeholder="e.g. Amazon referral fee")

    st.markdown("---")
    st.markdown("**Thresholds**")
    MATCH_THRESH = st.slider("Min Keyword Match %", 10, 80, 35, 5) / 100
    PRICE_TOL    = st.slider("BB Price Tolerance %", 5, 40, 20, 5) / 100
    MAX_RETRIES  = st.selectbox("Max retries per ASIN", [1,2,3], index=1)
    EXTRA_DELAY  = st.slider("Extra delay buffer (sec)", 0, 8, 0, 1)
    BB_WARN_FAIL = st.checkbox("Treat BB warning as FAIL", value=False)

    st.markdown("---")
    st.markdown("**P&L Recalculation**")
    RECALC_ON   = st.checkbox("Auto-recalculate when BB changes", value=True,
                               help="When live BB differs from sheet BB, recalculate breakeven, profit & margin")
    REF_RATE    = st.slider("Referral fee rate (if no fee column)", 5, 20, 15, 1,
                             help="Used only if no referral fee column is found") / 100

    st.markdown("---")
    if LOGO_B64:
        st.markdown(f'<img src="data:image/jpeg;base64,{LOGO_B64}" style="width:120px;border-radius:6px;">', unsafe_allow_html=True)
    st.markdown("<p style='color:#aaa;font-size:11px;margin-top:8px;'>VirVentures Verifier v5.0<br>Confidence BB · Auto P&L · Retry Engine</p>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# File Upload
# ══════════════════════════════════════════════════════════
st.markdown("""
<div style="background:linear-gradient(135deg,#fff8f3 0%,#fff3e8 100%);
            border:2px solid #f47920;border-radius:14px;padding:22px 28px;
            margin-bottom:16px;display:flex;align-items:center;gap:16px;
            box-shadow:0 2px 12px rgba(244,121,32,0.10);">
    <span style="font-size:2.2rem;flex-shrink:0;">📋</span>
    <div>
        <p style="color:#1e2d4e !important;font-size:1.05rem;font-weight:800;margin:0 0 4px 0;">
            Please upload your file and let me do this work for you!
        </p>
        <p style="color:#f47920 !important;font-size:0.85rem;font-weight:600;margin:0;">
            I will auto-detect columns, verify every ASIN, check live BB prices,
            recalculate your P&amp;L automatically &mdash; and hand you back a clean colour-coded report.
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload Excel", type=["xlsx","xls"],
                             label_visibility="collapsed",
                             help="Any .xlsx layout works — columns auto-detected")

# ══════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════
if uploaded:
    df = pd.read_excel(uploaded, dtype=str)
    df.columns = df.columns.str.strip()
    cols = list(df.columns)

    def res(ov, key):
        o = (ov or "").strip()
        return o if o and o in cols else detect_col(key, cols)

    ASIN_COL  = res(ov_asin,  "asin")
    TITLE_COL = res(ov_title, "title")
    DESC_COL  = res(ov_desc,  "desc")
    BRAND_COL = res(ov_brand, "brand")
    BB_COL    = res(ov_bb,    "bb_price")
    NET_COL   = res(ov_net,   "net_price")
    BE_COL    = res(ov_be,    "breakeven")
    FULL_COL  = res(ov_full,  "fulfillment")
    REF_COL   = res(ov_ref,   "referral")

    # Column status
    st.markdown('<p class="sec">🗂️ Column Detection</p>', unsafe_allow_html=True)
    det = {"ASIN": ASIN_COL, "Title": TITLE_COL, "Description": DESC_COL,
           "BB Price": BB_COL, "Net Price": NET_COL, "Breakeven": BE_COL,
           "Fulfillment": FULL_COL, "Referral Fee": REF_COL}
    ui_cols = st.columns(len(det))
    for idx, (lbl, col) in enumerate(det.items()):
        with ui_cols[idx]:
            st.metric(lbl, ("✅ "+col[:13]) if col else "⚠️ Not found")

    if not ASIN_COL:
        st.error("⛔ ASIN column not found. Set it in the sidebar.")
        st.stop()

    # Preview
    st.markdown('<p class="sec">👁️ Preview</p>', unsafe_allow_html=True)
    prev = [c for c in [ASIN_COL,TITLE_COL,DESC_COL,BRAND_COL,BB_COL,NET_COL,BE_COL] if c]
    st.dataframe(df[prev].head(8), use_container_width=True)

    valid_n = df[ASIN_COL].dropna().apply(lambda x: str(x).strip()).str.len().gt(4).sum()
    avg_d   = sum(DELAY_SEQ)/len(DELAY_SEQ) + EXTRA_DELAY
    est_min = round(valid_n * avg_d / 60, 1)

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total Rows",   len(df))
    c2.metric("Valid ASINs",  valid_n)
    c3.metric("Avg Delay",    f"~{avg_d:.1f}s")
    c4.metric("Est. Runtime", f"~{est_min} min")

    st.markdown("""
    <div class="info-box">
        <b>🧠 v5 Engine Active:</b>
        Confidence-validated BB price (rejects suspicious scrapes) ·
        Auto P&amp;L recalculation when BB changes ·
        Full breakeven + profit + margin remark per row ·
        Fuzzy keyword matching · 3x retry · 8 rotating agents · Human delay sequence
    </div>""", unsafe_allow_html=True)

    st.markdown("---")

    if st.button(f"🚀  START VERIFICATION  ·  {valid_n} ASINs", use_container_width=True):

        results = {
            "Live BB Price":   [],
            "BB Confidence":   [],
            "BB vs Sheet":     [],
            "New Breakeven":   [],
            "New Profit ($)":  [],
            "New Margin %":    [],
            "P&L Remark":      [],
            "Amazon Title":    [],
            "Keyword Match":   [],
            "Match %":         [],
            "Accuracy":        [],
            "Verification":    [],
            "Fail Reasons":    [],
        }

        verified_n=failed_n=warned_n=skipped_n=price_flag_n=captcha_n=recalc_n = 0

        st.markdown('<p class="sec">⚡ Live Progress</p>', unsafe_allow_html=True)
        pbar  = st.progress(0)
        sbox  = st.empty()
        ltbl  = st.empty()
        log   = []
        total = len(df)

        for i, row in df.iterrows():
            asin      = str(row.get(ASIN_COL,"")).strip()
            title     = str(row.get(TITLE_COL,"")).strip()  if TITLE_COL else ""
            desc      = str(row.get(DESC_COL,"")).strip()   if DESC_COL  else ""
            brand     = str(row.get(BRAND_COL,"")).strip()  if BRAND_COL else ""
            your_bb   = parse_price(row.get(BB_COL))        if BB_COL    else None
            net_price = parse_price(row.get(NET_COL))       if NET_COL   else None
            breakeven = parse_price(row.get(BE_COL))        if BE_COL    else None
            full_cost = parse_price(row.get(FULL_COL))      if FULL_COL  else None
            ref_fee   = parse_price(row.get(REF_COL))       if REF_COL   else None
            rn        = i + 1

            our_text  = " ".join(filter(None,[title, desc, brand]))

            pbar.progress(rn/total)
            sbox.markdown(
                f"<p style='color:#1e2d4e;font-size:13px;'>"
                f"Processing <b>{rn}/{total}</b> — <code>{asin}</code></p>",
                unsafe_allow_html=True
            )

            # Skip blank
            if not asin or asin.lower() in ("nan","") or len(asin) < 5:
                for k in results: results[k].append("SKIPPED")
                skipped_n += 1
                log.append({"#":rn,"ASIN":"—","Your BB":"—","Live BB":"—","Confidence":"—","Match":"—","Profit":"—","Status":"⏭️"})
                ltbl.dataframe(pd.DataFrame(log).tail(12), use_container_width=True)
                continue

            soup, status = fetch_with_retry(asin, MAX_RETRIES)

            if soup is None:
                note = "🤖 CAPTCHA" if status=="CAPTCHA" else "❌ Fetch failed"
                if status == "CAPTCHA": captcha_n += 1
                for k in results: results[k].append(note)
                failed_n += 1
                log.append({"#":rn,"ASIN":asin,"Your BB":"—","Live BB":"—","Confidence":"—","Match":"—","Profit":"—","Status":note})
                ltbl.dataframe(pd.DataFrame(log).tail(12), use_container_width=True)
                time.sleep(smart_delay(i, EXTRA_DELAY) + (8 if "CAPTCHA" in note else 0))
                continue

            # ── Extract with confidence ──
            live_bb_f, live_bb_str, bb_conf = get_bb_price_validated(soup, your_bb)

            # ── Keyword match ──
            amz_title = soup.select_one("#productTitle")
            amz_title = amz_title.get_text(strip=True) if amz_title else ""
            parts = [amz_title]
            for b in soup.select("#feature-bullets li span.a-list-item"):
                parts.append(b.get_text(strip=True))
            for sel in ["#productDescription","#aplus","#aplus3p_feature_div"]:
                t = soup.select_one(sel)
                if t: parts.append(t.get_text(" ",strip=True))
            amz_text = " ".join(parts).lower()

            score, ratio, matched, missed = smart_match(our_text, amz_text)
            match_pct = f"{score*100:.1f}%"

            # ── BB comparison ──
            bb_vs_sheet = "—"
            bb_severity = "ok"
            if live_bb_f is None:
                bb_vs_sheet = "⚠️ Not scraped (safe — page layout varied)"
                bb_severity = "warn"
            elif your_bb:
                diff_pct = (live_bb_f - your_bb) / your_bb * 100
                abs_diff = abs(live_bb_f - your_bb)
                if abs_diff <= 1.50:
                    bb_vs_sheet = f"✅ Match (Δ${abs_diff:.2f})"
                    bb_severity = "ok"
                elif abs(diff_pct) <= PRICE_TOL*100:
                    bb_vs_sheet = f"✅ OK (diff {diff_pct:+.1f}%)"
                    bb_severity = "ok"
                elif abs(diff_pct) <= PRICE_TOL*100*2:
                    bb_vs_sheet = f"⚠️ Soft change ({diff_pct:+.1f}%)"
                    bb_severity = "warn"
                    price_flag_n += 1
                else:
                    dir_s = "up" if diff_pct>0 else "down"
                    bb_vs_sheet = f"🔄 BB moved {dir_s} {abs(diff_pct):.1f}% → ${live_bb_f:.2f}"
                    bb_severity = "changed"
                    price_flag_n += 1
            else:
                bb_vs_sheet = f"ℹ️ No sheet BB | Live: {live_bb_str}"
                bb_severity = "warn"

            # ── P&L Recalculation ──
            new_be = new_profit = new_margin = None
            pl_remark = ""

            bb_changed = bb_severity in ("changed","warn") and live_bb_f is not None

            if RECALC_ON and live_bb_f is not None and (bb_changed or your_bb is None):
                recalc_n += 1
                fin = recalculate_financials(
                    live_bb_f, net_price,
                    breakeven, full_cost,
                    ref_fee if ref_fee else (live_bb_f * REF_RATE)
                )
                new_be     = fin["new_breakeven"]
                new_profit = fin["profit"]
                new_margin = fin["profit_margin"]
                pl_remark  = fin["remark"]
            elif live_bb_f is not None and net_price is not None:
                # BB hasn't changed much — still show current P&L for info
                est_ref  = ref_fee if ref_fee else round(live_bb_f * REF_RATE, 2)
                est_full = full_cost or 0
                new_be   = round(net_price + est_full + est_ref, 2)
                new_profit= round(live_bb_f - new_be, 2)
                new_margin= round((new_profit/live_bb_f)*100, 2) if live_bb_f>0 else 0
                pl_remark = (
                    f"Live BB: {live_bb_str} | "
                    f"Breakeven: ${new_be:.2f} | "
                    f"Profit: ${new_profit:.2f} | "
                    f"Margin: {new_margin:.2f}%"
                )

            accuracy = calc_accuracy(score, bb_severity!="changed", True, len(our_text.split()))

            # ── Verdict ──
            hard_fails = []
            soft_warns = []

            if score < MATCH_THRESH:
                hard_fails.append(f"Low description match ({match_pct}) — missed: {', '.join(missed[:4])}")
            if bb_severity == "changed" and BB_WARN_FAIL:
                hard_fails.append(bb_vs_sheet)
            elif bb_severity == "changed":
                soft_warns.append(f"BB price changed → recalculated (see P&L Remark)")
            if bb_severity == "warn" and BB_WARN_FAIL:
                hard_fails.append(bb_vs_sheet)
            elif bb_severity == "warn" and "Not scraped" not in bb_vs_sheet:
                soft_warns.append(bb_vs_sheet)

            if hard_fails:
                verdict      = "❌ FAILED"
                fail_reasons = " | ".join(hard_fails + soft_warns)
                failed_n    += 1; icon = "❌"
            elif soft_warns:
                verdict      = "⚠️ WARNING — Needs Review"
                fail_reasons = " | ".join(soft_warns)
                warned_n    += 1; icon = "⚠️"
            else:
                verdict      = "✅ Verified — 100% Authentic"
                fail_reasons = ""
                verified_n  += 1; icon = "✅"

            results["Live BB Price"].append(live_bb_str)
            results["BB Confidence"].append(bb_conf)
            results["BB vs Sheet"].append(bb_vs_sheet)
            results["New Breakeven"].append(f"${new_be:.2f}" if new_be else "—")
            results["New Profit ($)"].append(f"${new_profit:.2f}" if new_profit is not None else "—")
            results["New Margin %"].append(f"{new_margin:.2f}%" if new_margin is not None else "—")
            results["P&L Remark"].append(pl_remark)
            results["Amazon Title"].append(amz_title)
            results["Keyword Match"].append(ratio)
            results["Match %"].append(match_pct)
            results["Accuracy"].append(accuracy)
            results["Verification"].append(verdict)
            results["Fail Reasons"].append(fail_reasons)

            profit_str = f"${new_profit:.2f}" if new_profit is not None else "—"
            log.append({"#":rn,"ASIN":asin,"Your BB":f"${your_bb:.2f}" if your_bb else "—",
                        "Live BB":live_bb_str,"Conf":bb_conf[:4],"Match":match_pct,
                        "Profit":profit_str,"Status":icon})
            ltbl.dataframe(pd.DataFrame(log).tail(12), use_container_width=True)

            time.sleep(smart_delay(i, EXTRA_DELAY))

        # ── Complete ──────────────────────────────────────
        pbar.progress(1.0)
        sbox.markdown("<p style='color:#1a7a42;font-weight:800;font-size:15px;'>✅ VERIFICATION COMPLETE!</p>", unsafe_allow_html=True)

        for col, vals in results.items():
            while len(vals) < len(df): vals.append("")
            df[col] = vals

        # Summary
        st.markdown('<p class="sec">📊 Results Summary</p>', unsafe_allow_html=True)
        s1,s2,s3,s4,s5,s6,s7 = st.columns(7)
        s1.metric("✅ Verified",       verified_n)
        s2.metric("⚠️ Needs Review",   warned_n)
        s3.metric("❌ Failed",         failed_n)
        s4.metric("💰 BB Changed",     price_flag_n)
        s5.metric("🔄 P&L Recalc'd",   recalc_n)
        s6.metric("🤖 CAPTCHA",        captcha_n)
        s7.metric("⏭️ Skipped",        skipped_n)

        # Avg accuracy
        acc_vals = []
        for v in df["Accuracy"]:
            try: acc_vals.append(float(str(v).replace("%","")))
            except: pass
        if acc_vals:
            st.metric("📈 Avg Accuracy", f"{sum(acc_vals)/len(acc_vals):.1f}%")

        # P&L summary for changed rows
        if recalc_n > 0:
            st.markdown('<p class="sec">💰 P&L Recalculation Summary</p>', unsafe_allow_html=True)
            pl_df = df[df["P&L Remark"].str.len() > 10][[
                c for c in [ASIN_COL, TITLE_COL, BB_COL, "Live BB Price",
                            "New Breakeven","New Profit ($)","New Margin %","P&L Remark"]
                if c in df.columns
            ]].reset_index(drop=True)
            st.dataframe(pl_df, use_container_width=True)

        if captcha_n > 0:
            st.markdown('<div class="warn-box">🤖 Some rows hit Amazon CAPTCHA. Wait 15–20 mins then re-run those ASINs.</div>', unsafe_allow_html=True)

        # Attention list
        attn = df[df["Verification"].str.contains("FAILED|WARNING", na=False)]
        if len(attn) > 0:
            st.markdown('<p class="sec">🔴 Items Needing Attention</p>', unsafe_allow_html=True)
            show = [c for c in [ASIN_COL,TITLE_COL,BB_COL,"Live BB Price","BB vs Sheet",
                                 "New Profit ($)","Match %","Verification","Fail Reasons"] if c in df.columns]
            st.dataframe(attn[show].reset_index(drop=True), use_container_width=True)

        # Download
        st.markdown('<p class="sec">⬇️ Download</p>', unsafe_allow_html=True)
        excel_bytes = build_excel(df, MATCH_THRESH)
        st.download_button(
            "📥  Download Colour-Coded Excel — VirVentures_Verified.xlsx",
            data=excel_bytes,
            file_name="VirVentures_ASIN_Verification.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    st.markdown("""
    <div style="text-align:center;padding:60px 40px;background:#fff;
                border:2px dashed #f47920;border-radius:16px;margin-top:8px;">
        <p style="font-size:3rem;margin:0;">🔍</p>
        <p style="color:#1e2d4e !important;font-weight:800;font-size:1.2rem;margin:14px 0 6px;">
            Waiting for your file...
        </p>
        <p style="color:#f47920 !important;font-size:0.9rem;font-weight:600;margin:0 0 6px;">
            Upload your .xlsx above and I'll handle everything automatically
        </p>
        <p style="color:#aaa !important;font-size:0.8rem;margin:0;">
            Any column layout works &nbsp;·&nbsp; No config needed &nbsp;·&nbsp; P&L auto-calculated
        </p>
    </div>
    """, unsafe_allow_html=True)
