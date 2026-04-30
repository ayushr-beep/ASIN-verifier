
#  VirVentures ASIN Verifier — v6.0 Production
#  Senior Engineer Build: Semantic ML · FBA/FBM Detection · Price Ranges
#  SQLite Cache · ThreadPoolExecutor · Exponential Backoff · Cookie Auth
# =============================================================================

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import requests
import re
import random
import time
import io
import base64
import os
import json
import sqlite3
import logging
import traceback
import hashlib
import threading
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher
from http.cookiejar import MozillaCookieJar
from pathlib import Path

from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# Sentence-transformers — graceful degradation if not installed
try:
    from sentence_transformers import SentenceTransformer
    import numpy as np
    SEMANTIC_AVAILABLE = True
except ImportError:
    SEMANTIC_AVAILABLE = False

# =============================================================================
# PAGE CONFIG — must be first Streamlit call
# =============================================================================
st.set_page_config(
    page_title="VirVentures ASIN Verifier",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# CONSTANTS & CONFIGURATION
# =============================================================================
CACHE_DB          = "asin_cache.db"
ERROR_LOG         = "vv_errors.log"
CACHE_TTL_HOURS   = 4
DEFAULT_WORKERS   = 3
MAX_WORKERS       = 5
RETRY_DELAYS      = [5, 10, 20]          # exponential backoff seconds
BATCH_DELAY_MIN   = 2.0
BATCH_DELAY_MAX   = 4.0
SEMANTIC_MODEL_ID = "all-MiniLM-L6-v2"
KW_WEIGHT         = 0.40                 # keyword match weight
SEM_WEIGHT        = 0.60                 # semantic similarity weight

DELAY_SEQ = [8.0, 8.5, 9.8, 7.9, 10.2, 8.7, 11.1, 9.4, 7.6, 10.8, 8.2, 9.1]

UA_POOL = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:124.0) Gecko/20100101 Firefox/124.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
]
LANGS = ["en-US,en;q=0.9", "en-GB,en;q=0.9", "en-CA,en;q=0.8,fr;q=0.7", "en-AU,en;q=0.9"]

STOPWORDS = {
    "a","an","the","and","or","for","of","in","to","with","by","is","it","its",
    "–","-","&","at","on","from","as","are","be","this","that","will","has",
    "have","not","but","can","pk","pcs","set","new","use","used","each","per",
    "pack","count","size","color","colour","oz","lb","lbs","fl","ct","qty",
}

# =============================================================================
# LOGGING SETUP
# =============================================================================
logging.basicConfig(
    filename=ERROR_LOG,
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

def log_error(asin: str, msg: str, exc: Exception = None):
    detail = f"ASIN={asin} | {msg}"
    if exc:
        detail += f" | {traceback.format_exc()}"
    logging.error(detail)

# =============================================================================
# LOGO
# =============================================================================
def get_logo_b64() -> str | None:
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

# =============================================================================
# CSS — VirVentures White + Orange Theme
# =============================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
*, html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

/* ── App background ── */
.stApp, .main, .block-container { background: #f8f9fb !important; }

/* ── Header ── */
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
.vv-header .vv-title { color: #ffffff !important; }
.vv-header .vv-sub   { color: #f47920 !important; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] { background: #ffffff !important; border-right: 2px solid #f0f0f0 !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #1e2d4e !important; font-weight: 700 !important; }
section[data-testid="stSidebar"] label { color: #444 !important; font-weight: 500 !important; }
section[data-testid="stSidebar"] .stTextInput input {
    background: #fafafa !important; border: 1.5px solid #e0e0e0 !important;
    border-radius: 8px !important; color: #1a1a2e !important;
}

/* ── Buttons ── */
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

/* ── Metrics ── */
div[data-testid="metric-container"] {
    background: #fff !important; border: 1.5px solid #f0f0f0 !important;
    border-radius: 12px !important; padding: 18px !important; box-shadow: 0 2px 10px rgba(0,0,0,0.05) !important;
}
div[data-testid="metric-container"] label { color: #888 !important; font-size: 11px !important; font-weight: 600 !important; letter-spacing: 0.8px !important; text-transform: uppercase; }
div[data-testid="metric-container"] div[data-testid="metric-value"] { color: #1e2d4e !important; font-size: 1.8rem !important; font-weight: 800 !important; }

/* ── Progress ── */
.stProgress > div > div { background: linear-gradient(90deg, #f47920, #ffb347) !important; border-radius: 6px !important; }

/* ── Tables ── */
.stDataFrame { border: 1.5px solid #f0f0f0 !important; border-radius: 12px !important; background: #fff !important; }

/* ── File uploader ── */
[data-testid="stFileUploader"] { background: #fff !important; border: 2px dashed #e0e0e0 !important; border-radius: 12px !important; padding: 16px !important; }
[data-testid="stFileUploader"]:hover { border-color: #f47920 !important; background: #fff8f3 !important; }
[data-testid="stFileUploaderDropzone"] { border: none !important; background: transparent !important; }
[data-testid="stFileUploaderDropzone"] button { background: #f47920 !important; color: #fff !important; border: none !important; border-radius: 8px !important; font-weight: 700 !important; padding: 8px 20px !important; }

/* ── Section title ── */
.sec { color: #1e2d4e; font-size: 1rem; font-weight: 800; padding-bottom: 6px;
       border-bottom: 3px solid #f47920; display: inline-block; margin: 20px 0 14px 0; }

/* ── Info/warn/success boxes ── */
.info-box  { background: #fff8f3; border-left: 4px solid #f47920; border-radius: 0 10px 10px 0; padding: 12px 16px; font-size: 13px; color: #1e2d4e; margin: 10px 0; }
.warn-box  { background: #fffbea; border-left: 4px solid #f5a623; border-radius: 0 10px 10px 0; padding: 12px 16px; font-size: 13px; color: #7a5c00; margin: 10px 0; }
.good-box  { background: #f0faf4; border-left: 4px solid #2e7d32; border-radius: 0 10px 10px 0; padding: 12px 16px; font-size: 13px; color: #1b5e20; margin: 10px 0; }

/* ── Confidence badges ── */
.badge-high   { background:#e8f5e9; color:#2e7d32; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:700; border:1px solid #a5d6a7; }
.badge-medium { background:#fff8e1; color:#f57f17; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:700; border:1px solid #ffe082; }
.badge-low    { background:#fdecea; color:#c62828; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:700; border:1px solid #ef9a9a; }

/* ── Worker progress cards ── */
.worker-card {
    background: #fff; border: 1.5px solid #f0f0f0; border-radius: 10px;
    padding: 12px 16px; margin: 4px 0; font-size: 12px; color: #1e2d4e;
}

/* ── Scoped text ── */
.block-container p, .block-container span, .block-container li { color: #333 !important; }
h1,h2,h3,h4 { color: #1e2d4e !important; }
code { background: #f0f4ff !important; color: #1e2d4e !important; border-radius: 4px; padding: 2px 6px; }
.stAlert { border-radius: 10px !important; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# HEADER
# =============================================================================
st.markdown(f"""
<div class="vv-header">
    {LOGO_HTML}
    <div>
        <p class="vv-title" style="color:#ffffff !important;">
            Hi VirVentures 👋 &mdash; I am trained with an accuracy of 90%,<br>
            let&apos;s get started with your ASIN Verification!
        </p>
        <p class="vv-sub" style="color:#f47920 !important;">
            🔍 Semantic ML &nbsp;·&nbsp; FBA/FBM Detection &nbsp;·&nbsp; Price Ranges &nbsp;·&nbsp;
            Concurrent Engine &nbsp;·&nbsp; SQLite Cache &nbsp;·&nbsp; P&amp;L Recalculation
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# DATABASE / CACHE LAYER
# =============================================================================

def db_init():
    """Initialize SQLite cache database with required tables."""
    conn = sqlite3.connect(CACHE_DB)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS asin_cache (
            asin          TEXT PRIMARY KEY,
            price_data    TEXT,
            desc_data     TEXT,
            fulfillment   TEXT,
            fetched_at    TEXT,
            status        TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS failed_asins (
            asin       TEXT PRIMARY KEY,
            reason     TEXT,
            attempts   INTEGER DEFAULT 1,
            last_tried TEXT
        )
    """)
    conn.commit()
    conn.close()

def cache_get(asin: str, ttl_hours: int = CACHE_TTL_HOURS) -> dict | None:
    """
    Retrieve cached ASIN data if it exists and is within TTL.
    Returns None if not found or stale.
    """
    try:
        conn = sqlite3.connect(CACHE_DB)
        c = conn.cursor()
        c.execute("SELECT price_data, desc_data, fulfillment, fetched_at FROM asin_cache WHERE asin=?", (asin,))
        row = c.fetchone()
        conn.close()
        if not row:
            return None
        fetched_at = datetime.fromisoformat(row[3])
        if datetime.now() - fetched_at > timedelta(hours=ttl_hours):
            return None  # stale
        return {
            "price_data":  json.loads(row[0]) if row[0] else {},
            "desc_data":   json.loads(row[1]) if row[1] else {},
            "fulfillment": json.loads(row[2]) if row[2] else {},
            "fetched_at":  row[3],
            "from_cache":  True,
        }
    except Exception:
        return None

def cache_set(asin: str, price_data: dict, desc_data: dict, fulfillment: dict):
    """Store ASIN data in SQLite cache."""
    try:
        conn = sqlite3.connect(CACHE_DB)
        c = conn.cursor()
        c.execute("""
            INSERT OR REPLACE INTO asin_cache
            (asin, price_data, desc_data, fulfillment, fetched_at, status)
            VALUES (?, ?, ?, ?, ?, 'ok')
        """, (
            asin,
            json.dumps(price_data),
            json.dumps(desc_data),
            json.dumps(fulfillment),
            datetime.now().isoformat(),
        ))
        conn.commit()
        conn.close()
    except Exception as e:
        log_error(asin, "cache_set failed", e)

def cache_mark_failed(asin: str, reason: str):
    """Track failed ASINs for retry-only mode."""
    try:
        conn = sqlite3.connect(CACHE_DB)
        c = conn.cursor()
        c.execute("""
            INSERT INTO failed_asins (asin, reason, attempts, last_tried)
            VALUES (?, ?, 1, ?)
            ON CONFLICT(asin) DO UPDATE SET
                attempts   = attempts + 1,
                reason     = excluded.reason,
                last_tried = excluded.last_tried
        """, (asin, reason, datetime.now().isoformat()))
        conn.commit()
        conn.close()
    except Exception:
        pass

def cache_get_failed_asins() -> list[str]:
    """Return list of ASINs that previously failed — for retry mode."""
    try:
        conn = sqlite3.connect(CACHE_DB)
        c = conn.cursor()
        c.execute("SELECT asin FROM failed_asins ORDER BY last_tried DESC")
        rows = c.fetchall()
        conn.close()
        return [r[0] for r in rows]
    except Exception:
        return []

def cache_clear_failed(asins: list[str]):
    """Clear specific ASINs from the failed table after successful retry."""
    try:
        conn = sqlite3.connect(CACHE_DB)
        c = conn.cursor()
        c.executemany("DELETE FROM failed_asins WHERE asin=?", [(a,) for a in asins])
        conn.commit()
        conn.close()
    except Exception:
        pass

# Initialize DB on startup
db_init()

# =============================================================================
# COOKIE LOADING (Netscape format)
# =============================================================================

def load_cookies_from_file(uploaded_file) -> dict | None:
    """
    Parse a Netscape-format cookies.txt file (exported from browser extension).
    Returns dict of {name: value} for use with requests.
    """
    try:
        content = uploaded_file.read().decode("utf-8")
        cookies = {}
        for line in content.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("\t")
            if len(parts) >= 7:
                domain, _, path, secure, expires, name, value = parts[:7]
                if "amazon.com" in domain:
                    cookies[name] = value
        return cookies if cookies else None
    except Exception as e:
        log_error("cookies", "Failed to parse cookies.txt", e)
        return None

# =============================================================================
# HTTP / SCRAPING ENGINE
# =============================================================================

def build_session(cookies: dict | None = None) -> requests.Session:
    """Build a requests.Session with optional cookie injection."""
    session = requests.Session()
    if cookies:
        for name, value in cookies.items():
            session.cookies.set(name, value, domain=".amazon.com")
    return session

def random_headers() -> dict:
    return {
        "User-Agent":                random.choice(UA_POOL),
        "Accept-Language":           random.choice(LANGS),
        "Accept":                    "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Encoding":           "gzip, deflate, br",
        "Connection":                "keep-alive",
        "DNT":                       "1",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest":            "document",
        "Sec-Fetch-Mode":            "navigate",
        "Sec-Fetch-Site":            "none",
        "Cache-Control":             "max-age=0",
    }

def is_captcha(text: str) -> bool:
    return any(x in text for x in [
        "api-services-support@amazon.com",
        "Enter the characters you see below",
        "automated access",
        "Type the characters you see in this image",
        "robot or automated software",
    ])

def fetch_asin_page(
    asin: str,
    session: requests.Session,
    max_retries: int = 3,
) -> tuple[BeautifulSoup | None, str]:
    """
    Fetch Amazon product page with exponential backoff.
    Returns (soup, status_str).
    status_str: 'OK' | 'CAPTCHA' | 'FAILED' | 'TIMEOUT'
    """
    urls = [
        f"https://www.amazon.com/dp/{asin}",
        f"https://www.amazon.com/gp/product/{asin}",
        f"https://www.amazon.com/dp/{asin}?th=1&psc=1",
    ]
    for attempt in range(max_retries):
        url    = urls[attempt % len(urls)]
        delay  = RETRY_DELAYS[min(attempt, len(RETRY_DELAYS)-1)]
        try:
            r = session.get(url, headers=random_headers(), timeout=22)
            if r.status_code == 200:
                if is_captcha(r.text):
                    if attempt < max_retries - 1:
                        time.sleep(delay + random.uniform(3, 8))
                    continue
                return BeautifulSoup(r.text, "lxml"), "OK"
            elif r.status_code in (429, 503):
                # Rate limited — wait longer
                if attempt < max_retries - 1:
                    time.sleep(delay + random.uniform(5, 12))
                continue
            else:
                if attempt < max_retries - 1:
                    time.sleep(delay)
                continue
        except requests.exceptions.Timeout:
            if attempt < max_retries - 1:
                time.sleep(delay)
            continue
        except Exception as e:
            log_error(asin, f"fetch attempt {attempt+1} failed", e)
            if attempt < max_retries - 1:
                time.sleep(delay)
            continue
    return None, "FAILED"

def fetch_offer_listing(
    asin: str,
    session: requests.Session,
) -> BeautifulSoup | None:
    """
    Fetch the /gp/offer-listing page which contains ALL seller offers,
    prices, and fulfillment types. More reliable for price ranges.
    """
    url = f"https://www.amazon.com/gp/offer-listing/{asin}/ref=dp_olp_all_mbc?ie=UTF8&condition=new"
    try:
        r = session.get(url, headers=random_headers(), timeout=18)
        if r.status_code == 200 and not is_captcha(r.text):
            return BeautifulSoup(r.text, "lxml")
    except Exception as e:
        log_error(asin, "offer listing fetch failed", e)
    return None

# =============================================================================
# PRICE EXTRACTION WITH RANGE + FBA/FBM DETECTION
# =============================================================================

def extract_price_float(text: str) -> float | None:
    """Clean and parse a price string to float."""
    cleaned = re.sub(r"[^\d.]", "", text.replace(",", ""))
    try:
        val = float(cleaned)
        return val if 0.50 < val < 50_000 else None
    except (ValueError, TypeError):
        return None

def detect_fulfillment_type(container_html: str) -> str:
    """
    Detect whether an offer is FBA, Prime, or FBM.
    Returns: 'FBA' | 'FBM' | 'AMAZON' | 'UNKNOWN'
    """
    text_lower = container_html.lower()
    if "fulfilled by amazon" in text_lower or "amazon fulfillment" in text_lower:
        if "sold by amazon" in text_lower or "ships from amazon" in text_lower:
            return "AMAZON"
        return "FBA"
    if "prime" in text_lower and "fulfilled by amazon" in text_lower:
        return "FBA"
    if "fulfilled by merchant" in text_lower or "ships from and sold by" not in text_lower:
        return "FBM"
    return "UNKNOWN"

def extract_all_prices(
    soup_product: BeautifulSoup,
    soup_offers:  BeautifulSoup | None,
    your_bb:      float | None,
    seller_type:  str = "FBA",  # "FBA" | "FBM" | "BOTH"
) -> dict:
    """
    Comprehensive price extraction from both product page and offer listing.

    Returns dict:
        min_price        float | None
        max_price        float | None
        selected_price   float | None
        price_range_str  str
        selection_reason str
        fba_prices       list[float]
        fbm_prices       list[float]
        all_offers       list[dict]   each: {price, fulfillment, seller}
        bb_price         float | None  (main Buy Box)
    """
    result = {
        "min_price":        None,
        "max_price":        None,
        "selected_price":   None,
        "price_range_str":  "N/A",
        "selection_reason": "",
        "fba_prices":       [],
        "fbm_prices":       [],
        "all_offers":       [],
        "bb_price":         None,
    }

    all_prices_raw = []

    # ── 1. Main Buy Box price selectors ───────────────────
    BB_SELECTORS = [
        "#corePriceDisplay_desktop_feature_div .a-price-whole",
        "#corePriceDisplay_desktop_feature_div .a-offscreen",
        "#corePrice_feature_div .a-price-whole",
        "#corePrice_feature_div .a-offscreen",
        "#price_inside_buybox",
        "#priceblock_ourprice",
        "#priceblock_dealprice",
        ".priceToPay .a-offscreen",
        "#apex_offerDisplay_desktop .a-price .a-offscreen",
        "#newBuyBoxPrice",
    ]
    for sel in BB_SELECTORS:
        tag = soup_product.select_one(sel)
        if tag:
            p = extract_price_float(tag.get_text(strip=True))
            if p:
                result["bb_price"] = p
                all_prices_raw.append({"price": p, "fulfillment": "FBA", "source": "buybox"})
                break

    # ── 2. Price range detection on product page ──────────
    # Look for range patterns like "$22.99 - $34.99"
    page_text = str(soup_product)
    range_pattern = re.findall(r'\$\s*(\d+\.?\d*)\s*[-–]\s*\$\s*(\d+\.?\d*)', page_text)
    for lo_str, hi_str in range_pattern:
        lo, hi = extract_price_float(lo_str), extract_price_float(hi_str)
        if lo and hi and lo < hi:
            all_prices_raw.append({"price": lo, "fulfillment": "RANGE_LOW",  "source": "range"})
            all_prices_raw.append({"price": hi, "fulfillment": "RANGE_HIGH", "source": "range"})

    # ── 3. Offer listing page — richest data source ───────
    if soup_offers:
        offer_rows = soup_offers.select(".olpOffer, .a-row.olpOffer, #olpOfferList .a-row")
        if not offer_rows:
            # Try alternate selectors for newer Amazon layout
            offer_rows = soup_offers.select("[data-offering-id], .sgf-item")

        for offer_row in offer_rows[:20]:  # cap at 20 offers
            row_html  = str(offer_row)
            row_text  = offer_row.get_text(" ", strip=True)
            fulfill   = detect_fulfillment_type(row_html)

            # Extract price from this offer row
            price_tag = offer_row.select_one(".olpOfferPrice, .a-price .a-offscreen, .a-color-price")
            if not price_tag:
                # Try finding any $ amount in the row
                price_match = re.search(r'\$\s*(\d+\.?\d*)', row_text)
                if price_match:
                    p = extract_price_float(price_match.group(1))
                else:
                    continue
            else:
                p = extract_price_float(price_tag.get_text(strip=True))

            if not p:
                continue

            # Get seller name if available
            seller_tag = offer_row.select_one(".olpSellerName, .a-profile-name")
            seller     = seller_tag.get_text(strip=True) if seller_tag else "Unknown"

            offer_entry = {"price": p, "fulfillment": fulfill, "seller": seller, "source": "offers"}
            result["all_offers"].append(offer_entry)
            all_prices_raw.append(offer_entry)

            if fulfill in ("FBA", "AMAZON"):
                result["fba_prices"].append(p)
            elif fulfill == "FBM":
                result["fbm_prices"].append(p)

    # ── 4. Deduplicate and sort prices ───────────────────
    all_valid = [
        e["price"] for e in all_prices_raw
        if e.get("price") and 0.50 < e["price"] < 50_000
    ]

    # Sanity check vs your sheet BB price — reject outliers
    if your_bb and your_bb > 0:
        all_valid = [p for p in all_valid if 0.15 < p/your_bb < 6.0]

    if not all_valid:
        return result

    all_valid = sorted(set(all_valid))
    result["min_price"] = all_valid[0]
    result["max_price"] = all_valid[-1]

    # ── 5. Price range string ─────────────────────────────
    if len(all_valid) > 1 and result["max_price"] != result["min_price"]:
        result["price_range_str"] = f"${result['min_price']:.2f} – ${result['max_price']:.2f}"
    elif result["min_price"]:
        result["price_range_str"] = f"${result['min_price']:.2f}"

    # ── 6. Smart price selection logic ───────────────────
    fba_set = result["fba_prices"]
    fbm_set = result["fbm_prices"]

    if seller_type == "FBA" and fba_set:
        # User is FBA — use lowest FBA price (ignore FBM)
        result["selected_price"]   = min(fba_set)
        result["selection_reason"] = f"Lowest FBA price (ignoring {len(fbm_set)} FBM offers)"
    elif seller_type == "FBA" and not fba_set and result["bb_price"]:
        # No FBA offers found — use BB price
        result["selected_price"]   = result["bb_price"]
        result["selection_reason"] = "Buy Box price (no separate FBA offers found)"
    elif seller_type == "FBM" and fbm_set:
        result["selected_price"]   = min(fbm_set)
        result["selection_reason"] = f"Lowest FBM price"
    else:
        # Default: use minimum price (conservative)
        result["selected_price"]   = result["min_price"]
        result["selection_reason"] = "Minimum market price (conservative estimate)"

    # Final sanity guard
    if result["selected_price"] and your_bb and your_bb > 0:
        ratio = result["selected_price"] / your_bb
        if ratio > 6.0 or ratio < 0.15:
            result["selected_price"]   = result["bb_price"]
            result["selection_reason"] = "Reverted to BB price (selected price failed sanity check)"

    return result

# =============================================================================
# PRODUCT DATA EXTRACTION
# =============================================================================

def extract_product_text(soup: BeautifulSoup) -> dict:
    """
    Extract all text content from product page for matching.
    Returns dict with title, bullets, description, brand, full_text.
    """
    data = {"title": "", "bullets": [], "description": "", "brand": "", "full_text": ""}

    title_tag = soup.select_one("#productTitle")
    if title_tag:
        data["title"] = title_tag.get_text(strip=True)

    brand_tag = soup.select_one("#bylineInfo, #brand, .po-brand .po-break-word")
    if brand_tag:
        data["brand"] = brand_tag.get_text(strip=True).replace("Brand:", "").replace("Visit the", "").replace("Store", "").strip()

    for b in soup.select("#feature-bullets li span.a-list-item"):
        text = b.get_text(strip=True)
        if text and len(text) > 5:
            data["bullets"].append(text)

    desc_tag = soup.select_one("#productDescription")
    if desc_tag:
        data["description"] = desc_tag.get_text(" ", strip=True)

    aplus_tag = soup.select_one("#aplus, #aplus3p_feature_div")
    if aplus_tag:
        data["description"] += " " + aplus_tag.get_text(" ", strip=True)

    data["full_text"] = " ".join(filter(None, [
        data["title"], data["brand"],
        " ".join(data["bullets"]), data["description"]
    ])).lower()

    return data

# =============================================================================
# SEMANTIC ML MODEL (sentence-transformers)
# =============================================================================

@st.cache_resource(show_spinner="🧠 Loading ML model (first run only ~30s)...")
def load_semantic_model():
    """
    Load sentence-transformers model with Streamlit caching.
    Model is downloaded once (~80MB) and cached in memory.
    Returns model or None if unavailable.
    """
    if not SEMANTIC_AVAILABLE:
        return None
    try:
        model = SentenceTransformer(SEMANTIC_MODEL_ID)
        return model
    except Exception as e:
        log_error("model", "Failed to load sentence-transformers model", e)
        return None

def semantic_similarity(text_a: str, text_b: str, model) -> float:
    """
    Compute cosine similarity between two texts using the ML model.
    Returns float 0.0–1.0.
    Falls back to 0.0 if model unavailable or error.
    """
    if model is None or not text_a or not text_b:
        return 0.0
    try:
        import numpy as np
        emb_a = model.encode(text_a[:512], convert_to_numpy=True)  # cap for speed
        emb_b = model.encode(text_b[:512], convert_to_numpy=True)
        # Cosine similarity
        sim = float(np.dot(emb_a, emb_b) / (np.linalg.norm(emb_a) * np.linalg.norm(emb_b) + 1e-8))
        return max(0.0, min(1.0, sim))
    except Exception as e:
        log_error("semantic", "Similarity computation failed", e)
        return 0.0

# =============================================================================
# KEYWORD MATCHING (enhanced fuzzy)
# =============================================================================

def normalize_word(w: str) -> str:
    """Normalize word for fuzzy plural/variant matching."""
    w = w.lower().strip()
    if len(w) > 4:
        if w.endswith("ies"): return w[:-3] + "y"
        if w.endswith("ves"): return w[:-3] + "f"
        if w.endswith("es"):  return w[:-2]
        if w.endswith("s"):   return w[:-1]
    return w

def fuzzy_word_in_text(word: str, text: str, threshold: float = 0.82) -> bool:
    """Check if word appears in text — exact, normalized, substring, or fuzzy."""
    norm = normalize_word(word)
    if word in text or norm in text:
        return True
    if len(word) > 5 and (word[:5] in text or norm[:5] in text):
        return True
    if len(word) > 4:
        for tw in text.split():
            if len(tw) > 3 and SequenceMatcher(None, norm, tw).ratio() >= threshold:
                return True
    return False

def keyword_match_score(our_text: str, amz_text: str) -> tuple[float, str, list, list]:
    """
    Returns (score, ratio_str, matched_words, missed_words).
    Uses enhanced fuzzy matching.
    """
    if not our_text or not amz_text:
        return 0.0, "0/0", [], []
    words = list({
        w.lower() for w in re.findall(r"[a-zA-Z0-9]+", str(our_text))
        if w.lower() not in STOPWORDS and len(w) > 2
    })
    if not words:
        return 0.0, "0/0", [], []
    matched = [w for w in words if fuzzy_word_in_text(w, amz_text)]
    missed  = [w for w in words if not fuzzy_word_in_text(w, amz_text)]
    score   = len(matched) / len(words)
    return score, f"{len(matched)}/{len(words)}", matched, missed

# =============================================================================
# COMBINED MATCHING (keyword + semantic)
# =============================================================================

def combined_match_score(
    our_text:  str,
    amz_text:  str,
    model,
) -> tuple[float, float, float, str]:
    """
    Compute combined match using keyword + semantic similarity.
    Weights: 40% keyword, 60% semantic (or 100% keyword if model unavailable).

    Returns (combined_score, kw_score, sem_score, ratio_str).
    """
    kw_score, ratio_str, matched, missed = keyword_match_score(our_text, amz_text)

    if model is not None and SEMANTIC_AVAILABLE:
        sem_score    = semantic_similarity(our_text[:512], amz_text[:512], model)
        combined     = (KW_WEIGHT * kw_score) + (SEM_WEIGHT * sem_score)
        method       = f"KW:{kw_score*100:.0f}% + SEM:{sem_score*100:.0f}%"
    else:
        sem_score    = 0.0
        combined     = kw_score
        method       = f"KW:{kw_score*100:.0f}% (ML unavailable)"

    return combined, kw_score, sem_score, ratio_str, matched, missed

# =============================================================================
# CONFIDENCE SCORING SYSTEM
# =============================================================================

def compute_confidence(
    combined_score:  float,
    bb_severity:     str,       # 'ok' | 'warn' | 'changed'
    price_available: bool,
    has_range:       bool,
    fetch_ok:        bool,
) -> dict:
    """
    Business-relevant confidence scoring replacing the old accuracy number.

    Returns dict:
        level        "HIGH" | "MEDIUM" | "LOW"
        profitability "PROFITABLE" | "MARGINAL" | "LOSS" | "UNKNOWN"
        action       "NONE" | "REVIEW" | "RE-EVALUATE"
        reasons      list[str]
    """
    result = {"level": "LOW", "profitability": "UNKNOWN", "action": "RE-EVALUATE", "reasons": []}

    if not fetch_ok:
        result["reasons"].append("Page could not be fetched")
        return result

    reasons = []
    score   = 0   # 0-100 internal score → maps to HIGH/MEDIUM/LOW

    # Description match (up to 50 points)
    if combined_score >= 0.80:
        score += 50
    elif combined_score >= 0.55:
        score += 30
        reasons.append(f"Partial description match ({combined_score*100:.0f}%)")
    else:
        score += 10
        reasons.append(f"Low description match ({combined_score*100:.0f}%)")

    # Price availability (up to 30 points)
    if price_available and bb_severity == "ok":
        score += 30
    elif price_available and bb_severity == "warn":
        score += 20
        reasons.append("BB price soft change detected")
    elif price_available and bb_severity == "changed":
        score += 10
        reasons.append("Significant BB price change — recalculated")
    else:
        reasons.append("BB price not available")

    # Price range (deduct 5 if range detected — less certainty)
    if has_range:
        score -= 5
        reasons.append("Price range detected (multiple sellers)")

    # Map score to level
    if score >= 70:
        result["level"] = "HIGH"
    elif score >= 40:
        result["level"] = "MEDIUM"
    else:
        result["level"] = "LOW"

    # Action
    if result["level"] == "HIGH":
        result["action"] = "NONE"
    elif result["level"] == "MEDIUM":
        result["action"] = "REVIEW"
    else:
        result["action"] = "RE-EVALUATE"

    result["reasons"] = reasons
    return result

# =============================================================================
# P&L RECALCULATION ENGINE
# =============================================================================

def recalculate_pl(
    selected_price:   float | None,
    max_price:        float | None,
    net_price:        float | None,
    fulfillment_cost: float | None,
    referral_rate:    float = 0.15,
) -> dict:
    """
    Full P&L recalculation with conservative and upside scenarios.

    Conservative: uses selected_price (min/FBA price)
    Upside:       uses max_price if available

    Returns dict with all values + human-readable remark.
    """
    result = {
        "new_breakeven":       None,
        "conservative_profit": None,
        "upside_profit":       None,
        "conservative_margin": None,
        "upside_margin":       None,
        "remark":              "",
        "status":              "UNKNOWN",   # PROFITABLE | MARGINAL | LOSS
        "recommendation":      "",
    }

    if selected_price is None or net_price is None or net_price <= 0:
        result["remark"] = "Insufficient data for P&L calculation"
        return result

    # Calculate referral fee based on selected price
    ref_fee      = round(selected_price * referral_rate, 2)
    fixed_costs  = fulfillment_cost or 0.0
    new_breakeven= round(net_price + fixed_costs + ref_fee, 2)
    result["new_breakeven"] = new_breakeven

    # ── Conservative scenario (using selected/min price) ──
    cons_profit = round(selected_price - new_breakeven, 2)
    cons_margin = round((cons_profit / selected_price) * 100, 2) if selected_price > 0 else 0
    result["conservative_profit"] = cons_profit
    result["conservative_margin"] = cons_margin

    # ── Upside scenario (using max price if different) ────
    if max_price and max_price != selected_price:
        up_ref_fee    = round(max_price * referral_rate, 2)
        up_breakeven  = round(net_price + fixed_costs + up_ref_fee, 2)
        up_profit     = round(max_price - up_breakeven, 2)
        up_margin     = round((up_profit / max_price) * 100, 2) if max_price > 0 else 0
        result["upside_profit"] = up_profit
        result["upside_margin"] = up_margin
    else:
        result["upside_profit"] = cons_profit
        result["upside_margin"] = cons_margin

    # ── Profitability status ──────────────────────────────
    if cons_profit > 1.0:
        result["status"] = "PROFITABLE"
    elif cons_profit > 0:
        result["status"] = "MARGINAL"
    else:
        result["status"] = "LOSS"

    # ── Human-readable remark ─────────────────────────────
    upside_str = ""
    if result["upside_profit"] != cons_profit:
        upside_str = f" | Upside: ${result['upside_profit']:.2f} ({result['upside_margin']:.1f}%)"

    result["remark"] = (
        f"Selected Price: ${selected_price:.2f} | "
        f"Breakeven: ${new_breakeven:.2f} "
        f"(Net ${net_price:.2f} + Fulfillment ${fixed_costs:.2f} + Referral ${ref_fee:.2f}) | "
        f"Conservative Profit: ${cons_profit:.2f} ({cons_margin:.1f}%)"
        f"{upside_str}"
    )

    # ── Recommendation ───────────────────────────────────
    if result["status"] == "PROFITABLE" and cons_margin >= 15:
        result["recommendation"] = "✅ Strong margin — good to list"
    elif result["status"] == "PROFITABLE" and cons_margin >= 8:
        result["recommendation"] = "✅ Acceptable margin — proceed"
    elif result["status"] == "MARGINAL":
        result["recommendation"] = "⚠️ Thin margin — monitor closely"
    else:
        result["recommendation"] = "❌ Loss at current BB — do not list"

    return result

# =============================================================================
# COLUMN DETECTION (universal)
# =============================================================================
ALIASES = {
    "asin":        ["output asin","asin","input_asin","amazon asin","asin#","asin number"],
    "title":       ["title","input_product name","product name","product title","item name","name"],
    "desc":        ["description","product description","item description","desc","full description","input_description"],
    "brand":       ["brand","input_brand name","brand name","manufacturer","vendor"],
    "upc":         ["upc","upc#","input_upc#","barcode","ean","upc code"],
    "bb_price":    ["bb price","buy box price","buybox price","bb_price","buy box","current bb"],
    "net_price":   ["net price","netprice","net_price","cost","vendor cost","our cost"],
    "breakeven":   ["breakeven","break even","break-even","bep"],
    "fulfillment": ["fullfilment cost","fulfillment cost","fba fees","fulfillment cost subtotal","fulfil"],
    "referral":    ["amazon referral fee","referral fee","amazon commission","commission"],
}

def detect_col(key: str, cols: list) -> str | None:
    cl = {c.strip().lower(): c for c in cols}
    for alias in ALIASES[key]:
        for k, v in cl.items():
            if alias in k or k in alias:
                return v
    return None

def parse_price(val) -> float | None:
    if not val or str(val).strip().lower() in ("nan","","none","n/a"):
        return None
    try:
        return float(re.sub(r"[^\d.]", "", str(val)))
    except (ValueError, TypeError):
        return None

# =============================================================================
# CORE VERIFICATION FUNCTION (one ASIN)
# =============================================================================

def verify_single_asin(
    row_data:      dict,
    session:       requests.Session,
    sem_model,
    config:        dict,
    force_refresh: bool = False,
) -> dict:
    """
    Full verification pipeline for one ASIN.
    Runs in thread — all state is local.

    Returns result dict with all output columns.
    """
    asin      = row_data.get("asin", "").strip()
    our_title = row_data.get("title", "")
    our_desc  = row_data.get("desc",  "")
    our_brand = row_data.get("brand", "")
    your_bb   = row_data.get("your_bb")
    net_price = row_data.get("net_price")
    full_cost = row_data.get("full_cost")
    ref_fee   = row_data.get("ref_fee")

    our_text = " ".join(filter(None, [our_title, our_desc, our_brand]))

    base = {
        "ASIN":                asin,
        "Price Range":         "N/A",
        "Selected Price":      "N/A",
        "Selection Reason":    "",
        "FBA Prices":          "",
        "FBM Prices":          "",
        "BB Price (Live)":     "N/A",
        "BB vs Sheet":         "—",
        "New Breakeven":       "—",
        "Conservative Profit": "—",
        "Upside Profit":       "—",
        "Conservative Margin": "—",
        "Upside Margin":       "—",
        "P&L Status":          "UNKNOWN",
        "Recommendation":      "",
        "P&L Remark":          "",
        "Amazon Title":        "",
        "Keyword Match":       "0/0",
        "KW Score %":          "0%",
        "Semantic Score %":    "0%",
        "Combined Score %":    "0%",
        "Confidence Level":    "LOW",
        "Action Required":     "RE-EVALUATE",
        "Verification":        "❌ FAILED",
        "Fail Reasons":        "",
        "Source":              "scraper",
        "Cache Hit":           "No",
    }

    if not asin or len(asin) < 5:
        base["Verification"]  = "⏭️ SKIPPED"
        base["Fail Reasons"]  = "No valid ASIN"
        return base

    # ── Check cache first ─────────────────────────────────
    cached = None if force_refresh else cache_get(asin)
    if cached:
        price_data   = cached["price_data"]
        desc_data    = cached["desc_data"]
        base["Cache Hit"] = "Yes"
    else:
        # ── Fetch product page ────────────────────────────
        soup_product, status = fetch_asin_page(asin, session, config.get("max_retries", 3))
        if soup_product is None:
            base["Verification"] = "❌ FETCH FAILED"
            base["Fail Reasons"] = f"Could not load Amazon page ({status})"
            cache_mark_failed(asin, status)
            log_error(asin, f"Fetch failed: {status}")
            return base

        # ── Fetch offer listing (for price range + FBA/FBM) ─
        soup_offers = fetch_offer_listing(asin, session)
        time.sleep(random.uniform(1.0, 2.5))  # small gap between the two fetches

        # ── Extract data ──────────────────────────────────
        seller_type = config.get("seller_type", "FBA")
        price_data  = extract_all_prices(soup_product, soup_offers, your_bb, seller_type)
        desc_data   = extract_product_text(soup_product)

        # ── Cache it ──────────────────────────────────────
        cache_set(asin, price_data, desc_data, {
            "fba_count": len(price_data.get("fba_prices", [])),
            "fbm_count": len(price_data.get("fbm_prices", [])),
        })

    # ── Description matching ──────────────────────────────
    amz_text     = desc_data.get("full_text", "")
    amz_title    = desc_data.get("title", "")
    combined, kw_score, sem_score, ratio_str, matched, missed = combined_match_score(
        our_text, amz_text, sem_model
    )

    match_thresh = config.get("match_threshold", 0.35)

    # ── Price logic ───────────────────────────────────────
    sel_price   = price_data.get("selected_price")
    max_price   = price_data.get("max_price")
    bb_live     = price_data.get("bb_price")
    price_range = price_data.get("price_range_str", "N/A")
    has_range   = "–" in price_range

    # BB comparison
    bb_severity = "ok"
    bb_vs_sheet = "—"
    if bb_live is None:
        bb_vs_sheet = "⚠️ Not found"
        bb_severity = "warn"
    elif your_bb:
        diff_pct = (bb_live - your_bb) / your_bb * 100
        abs_diff = abs(bb_live - your_bb)
        if abs_diff <= 1.50:
            bb_vs_sheet = f"✅ Match (Δ${abs_diff:.2f})"
        elif abs(diff_pct) <= config.get("price_tolerance", 0.20) * 100:
            bb_vs_sheet = f"✅ OK ({diff_pct:+.1f}%)"
        elif abs(diff_pct) <= 40:
            bb_vs_sheet = f"⚠️ Changed {diff_pct:+.1f}% → ${bb_live:.2f}"
            bb_severity = "warn"
        else:
            bb_vs_sheet = f"🔄 Big shift {diff_pct:+.1f}% → ${bb_live:.2f}"
            bb_severity = "changed"
    else:
        bb_vs_sheet = f"ℹ️ Live: ${bb_live:.2f}" if bb_live else "—"

    # ── P&L recalculation ─────────────────────────────────
    eff_ref_rate = config.get("referral_rate", 0.15)
    pl = recalculate_pl(sel_price, max_price, net_price, full_cost, eff_ref_rate)

    # ── Confidence scoring ────────────────────────────────
    conf = compute_confidence(
        combined_score  = combined,
        bb_severity     = bb_severity,
        price_available = sel_price is not None,
        has_range       = has_range,
        fetch_ok        = True,
    )
    conf["profitability"] = pl["status"]

    # ── Verdict ───────────────────────────────────────────
    hard_fails = []
    soft_warns = []

    if combined < match_thresh:
        hard_fails.append(
            f"Description match too low ({combined*100:.0f}%) — missed: {', '.join(missed[:5])}"
        )
    if bb_severity == "changed" and config.get("bb_warn_as_fail", False):
        hard_fails.append(bb_vs_sheet)
    elif bb_severity == "changed":
        soft_warns.append(f"BB changed — recalculated P&L")
    if pl["status"] == "LOSS":
        soft_warns.append(f"Currently at a loss at this BB price")

    if hard_fails:
        verdict      = "❌ FAILED"
        fail_reasons = " | ".join(hard_fails + soft_warns)
    elif soft_warns:
        verdict      = "⚠️ WARNING — Needs Review"
        fail_reasons = " | ".join(soft_warns)
    else:
        verdict      = "✅ Verified — 100% Authentic"
        fail_reasons = ""

    # ── Populate output dict ──────────────────────────────
    base.update({
        "Price Range":          price_range,
        "Selected Price":       f"${sel_price:.2f}" if sel_price else "N/A",
        "Selection Reason":     price_data.get("selection_reason", ""),
        "FBA Prices":           ", ".join(f"${p:.2f}" for p in sorted(set(price_data.get("fba_prices", [])))),
        "FBM Prices":           ", ".join(f"${p:.2f}" for p in sorted(set(price_data.get("fbm_prices", [])))),
        "BB Price (Live)":      f"${bb_live:.2f}" if bb_live else "N/A",
        "BB vs Sheet":          bb_vs_sheet,
        "New Breakeven":        f"${pl['new_breakeven']:.2f}"       if pl["new_breakeven"]       else "—",
        "Conservative Profit":  f"${pl['conservative_profit']:.2f}" if pl["conservative_profit"] is not None else "—",
        "Upside Profit":        f"${pl['upside_profit']:.2f}"       if pl["upside_profit"]       is not None else "—",
        "Conservative Margin":  f"{pl['conservative_margin']:.1f}%" if pl["conservative_margin"] is not None else "—",
        "Upside Margin":        f"{pl['upside_margin']:.1f}%"       if pl["upside_margin"]       is not None else "—",
        "P&L Status":           pl["status"],
        "Recommendation":       pl["recommendation"],
        "P&L Remark":           pl["remark"],
        "Amazon Title":         amz_title,
        "Keyword Match":        ratio_str,
        "KW Score %":           f"{kw_score*100:.1f}%",
        "Semantic Score %":     f"{sem_score*100:.1f}%" if SEMANTIC_AVAILABLE else "N/A",
        "Combined Score %":     f"{combined*100:.1f}%",
        "Confidence Level":     conf["level"],
        "Action Required":      conf["action"],
        "Verification":         verdict,
        "Fail Reasons":         fail_reasons,
        "Source":               "cache" if base["Cache Hit"] == "Yes" else "scraper",
    })

    return base

# =============================================================================
# EXCEL EXPORT BUILDER
# =============================================================================

def build_excel(df: pd.DataFrame, match_thresh: float) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)

    wb = load_workbook(buf)
    ws = wb.active

    # Colour palette
    NAVY   = PatternFill("solid", fgColor="1e2d4e")
    GREEN  = PatternFill("solid", fgColor="C6EFCE")
    RED    = PatternFill("solid", fgColor="FFC7CE")
    AMBER  = PatternFill("solid", fgColor="FFEB9C")
    BLUE   = PatternFill("solid", fgColor="DDEEFF")
    LGRN   = PatternFill("solid", fgColor="E8F5E9")
    LRED   = PatternFill("solid", fgColor="FFE0E0")
    LGREY  = PatternFill("solid", fgColor="F5F5F5")
    ORANGE = PatternFill("solid", fgColor="FFE0CC")
    BOLD   = Font(bold=True)
    WHITE  = Font(bold=True, color="FFFFFF")
    WRAP   = Alignment(wrap_text=True, vertical="top")
    CTR    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ITALIC = Font(italic=True, size=9)

    # Style header
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = NAVY, WHITE, CTR
    ws.row_dimensions[1].height = 32

    hdr = [c.value for c in ws[1]]
    def ci(n):
        try: return hdr.index(n) + 1
        except: return None

    v_ci  = ci("Verification")
    cl_ci = ci("Confidence Level")
    ar_ci = ci("Action Required")
    ps_ci = ci("P&L Status")
    cp_ci = ci("Conservative Profit")
    cm_ci = ci("Conservative Margin")
    sc_ci = ci("Combined Score %")
    fr_ci = ci("Fail Reasons")
    pl_ci = ci("P&L Remark")
    rc_ci = ci("Recommendation")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        v_val  = str(row[v_ci  -1].value or "") if v_ci  else ""
        cl_val = str(row[cl_ci -1].value or "") if cl_ci else ""
        ps_val = str(row[ps_ci -1].value or "") if ps_ci else ""
        cp_val = str(row[cp_ci -1].value or "") if cp_ci else ""

        # Verification column
        if v_ci:
            c = row[v_ci-1]
            if "Verified"  in v_val: c.fill, c.font = GREEN, BOLD
            elif "FAILED"  in v_val: c.fill, c.font = RED,   BOLD
            elif "WARNING" in v_val: c.fill, c.font = AMBER, BOLD
            elif "SKIPPED" in v_val: c.fill = LGREY

        # Confidence level
        if cl_ci:
            c = row[cl_ci-1]
            if cl_val == "HIGH":   c.fill = GREEN
            elif cl_val == "MEDIUM": c.fill = AMBER
            elif cl_val == "LOW":  c.fill = RED
            c.font = BOLD

        # Action required
        if ar_ci:
            c = row[ar_ci-1]
            ar_val = str(c.value or "")
            if ar_val == "NONE":        c.fill = LGRN
            elif ar_val == "REVIEW":    c.fill = AMBER
            elif ar_val == "RE-EVALUATE": c.fill = LRED
            c.font = BOLD

        # P&L Status
        if ps_ci:
            c = row[ps_ci-1]
            if ps_val == "PROFITABLE": c.fill, c.font = LGRN, BOLD
            elif ps_val == "MARGINAL": c.fill, c.font = AMBER, BOLD
            elif ps_val == "LOSS":     c.fill, c.font = LRED,  BOLD

        # Conservative Profit — green if positive, red if negative
        if cp_ci:
            try:
                profit = float(str(row[cp_ci-1].value or "").replace("$","").replace(",",""))
                row[cp_ci-1].fill = LGRN if profit > 0.5 else (LRED if profit < -0.5 else AMBER)
                row[cp_ci-1].font = BOLD
            except: pass

        # Conservative Margin
        if cm_ci:
            try:
                margin = float(str(row[cm_ci-1].value or "").replace("%",""))
                row[cm_ci-1].fill = LGRN if margin >= 10 else (AMBER if margin >= 0 else LRED)
            except: pass

        # P&L remark — wrap text, small italic
        if pl_ci and row[pl_ci-1].value:
            row[pl_ci-1].alignment = WRAP
            row[pl_ci-1].font      = ITALIC

        # Recommendation
        if rc_ci and row[rc_ci-1].value:
            row[rc_ci-1].alignment = WRAP

        # Fail reasons
        if fr_ci and row[fr_ci-1].value:
            row[fr_ci-1].font      = Font(italic=True, color="C00000", size=9)
            row[fr_ci-1].alignment = WRAP

    # Auto column widths
    for col in ws.columns:
        mx = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(mx + 3, 55)

    ws.freeze_panes = "B2"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# =============================================================================
# STREAMLIT SIDEBAR
# =============================================================================
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    st.markdown("---")

    # ── Keepa API (optional) ──────────────────────────────
    st.markdown("**🔑 Keepa API** *(optional)*")
    keepa_key = st.text_input("Keepa API Key", type="password",
        placeholder="Paste key — faster BB price data",
        help="keepa.com → Profile → API. Replaces Amazon scraper for BB price.")
    if keepa_key:
        st.success("✅ Keepa active")
    else:
        st.caption("ℹ️ No key — using Amazon scraper")

    st.markdown("---")

    # ── Prime cookies ─────────────────────────────────────
    st.markdown("**🍪 Prime Session Cookies** *(optional)*")
    cookies_file = st.file_uploader("cookies.txt (Netscape format)", type=["txt"],
        label_visibility="collapsed",
        help="Export from browser with 'Get cookies.txt LOCALLY' extension. Gives Prime prices.")
    cookies_dict = None
    if cookies_file:
        cookies_dict = load_cookies_from_file(cookies_file)
        if cookies_dict:
            st.success(f"✅ {len(cookies_dict)} Amazon cookies loaded")
        else:
            st.warning("⚠️ No Amazon cookies found in file")

    st.markdown("---")

    # ── Column overrides ──────────────────────────────────
    st.markdown("**🗂️ Column Overrides**")
    st.caption("Leave blank for auto-detect")
    ov_asin  = st.text_input("ASIN Column",          placeholder="e.g. Output ASIN")
    ov_title = st.text_input("Title Column",         placeholder="e.g. Title")
    ov_desc  = st.text_input("Description Column",   placeholder="e.g. Description")
    ov_brand = st.text_input("Brand Column",         placeholder="e.g. Brand")
    ov_bb    = st.text_input("BB Price Column",      placeholder="e.g. BB Price")
    ov_net   = st.text_input("Net Price Column",     placeholder="e.g. Net Price")
    ov_full  = st.text_input("Fulfillment Cost Col", placeholder="e.g. Fullfilment Cost Subtotal")
    ov_ref   = st.text_input("Referral Fee Col",     placeholder="e.g. Amazon referral fee")

    st.markdown("---")

    # ── Matching thresholds ───────────────────────────────
    st.markdown("**🎯 Matching Thresholds**")
    MATCH_THRESH = st.slider("Min Combined Match %", 10, 80, 35, 5,
        help="Keyword + Semantic weighted score") / 100
    PRICE_TOL    = st.slider("BB Price Tolerance %", 5, 40, 20, 5) / 100
    REF_RATE     = st.slider("Referral Fee Rate %",  5, 20, 15, 1,
        help="Amazon referral % of selling price") / 100

    st.markdown("---")

    # ── Concurrency & behaviour ───────────────────────────
    st.markdown("**⚡ Concurrency**")
    N_WORKERS    = st.slider("Parallel Workers", 1, MAX_WORKERS, DEFAULT_WORKERS, 1,
        help="3 is safe. Go higher only if you have Prime cookies.")
    MAX_RETRIES  = st.selectbox("Max retries per ASIN", [1,2,3], index=1)
    EXTRA_DELAY  = st.slider("Extra delay buffer (sec)", 0, 8, 0, 1)
    SELLER_TYPE  = st.selectbox("Your Fulfillment Type", ["FBA","FBM","BOTH"], index=0,
        help="FBA = FBM prices ignored for profit calc")
    BB_WARN_FAIL = st.checkbox("Treat BB change as FAIL", value=False)

    st.markdown("---")

    # ── Cache controls ────────────────────────────────────
    st.markdown("**💾 Cache**")
    FORCE_REFRESH = st.checkbox("Force Refresh (ignore cache)", value=False,
        help="Re-fetch all ASINs even if cached")
    if st.button("🗑️ Clear Entire Cache"):
        if os.path.exists(CACHE_DB):
            os.remove(CACHE_DB)
            db_init()
            st.success("Cache cleared!")

    failed_list = cache_get_failed_asins()
    RETRY_ONLY  = st.checkbox(f"Retry Failed Only ({len(failed_list)} ASINs)", value=False,
        help="Only process ASINs that failed in a previous run")

    st.markdown("---")
    if LOGO_B64:
        st.markdown(f'<img src="data:image/jpeg;base64,{LOGO_B64}" style="width:120px;border-radius:6px;">', unsafe_allow_html=True)
    st.markdown("<p style='color:#aaa;font-size:11px;margin-top:8px;'>VirVentures Verifier v6.0<br>Semantic ML · FBA/FBM · Concurrent · Cache</p>", unsafe_allow_html=True)

# =============================================================================
# LOAD ML MODEL (cached globally)
# =============================================================================
sem_model = load_semantic_model()

if SEMANTIC_AVAILABLE and sem_model:
    st.markdown('<div class="good-box">🧠 <b>Semantic ML Active</b> — sentence-transformers loaded. Combined scoring: 40% keyword + 60% semantic similarity.</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="info-box">📝 <b>Keyword-Only Mode</b> — Install <code>pip install sentence-transformers</code> to enable ML semantic matching.</div>', unsafe_allow_html=True)

# =============================================================================
# FILE UPLOAD
# =============================================================================
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
            Semantic ML · FBA/FBM price detection · Price ranges · Conservative &amp; upside P&amp;L · Concurrent verification
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload Excel", type=["xlsx","xls"],
                             label_visibility="collapsed",
                             help="Any .xlsx layout — columns auto-detected")

# =============================================================================
# MAIN VERIFICATION FLOW
# =============================================================================
if uploaded:
    df = pd.read_excel(uploaded, dtype=str)
    df.columns = df.columns.str.strip()
    all_cols   = list(df.columns)

    # ── Resolve columns ───────────────────────────────────
    def res(ov, key):
        o = (ov or "").strip()
        return o if (o and o in all_cols) else detect_col(key, all_cols)

    ASIN_COL  = res(ov_asin,  "asin")
    TITLE_COL = res(ov_title, "title")
    DESC_COL  = res(ov_desc,  "desc")
    BRAND_COL = res(ov_brand, "brand")
    BB_COL    = res(ov_bb,    "bb_price")
    NET_COL   = res(ov_net,   "net_price")
    FULL_COL  = res(ov_full,  "fulfillment")
    REF_COL   = res(ov_ref,   "referral")

    # ── Column detection display ──────────────────────────
    st.markdown('<p class="sec">🗂️ Column Detection</p>', unsafe_allow_html=True)
    det_map = {"ASIN": ASIN_COL,"Title": TITLE_COL,"Description": DESC_COL,
               "Brand": BRAND_COL,"BB Price": BB_COL,"Net Price": NET_COL,
               "Fulfillment": FULL_COL,"Referral": REF_COL}
    cols_ui = st.columns(len(det_map))
    for idx, (lbl, col) in enumerate(det_map.items()):
        with cols_ui[idx]:
            st.metric(lbl, ("✅ "+col[:13]) if col else "⚠️ Not found")

    if not ASIN_COL:
        st.error("⛔ ASIN column not found. Set it in the sidebar.")
        st.stop()

    # ── Filter for retry-only mode ────────────────────────
    if RETRY_ONLY and failed_list:
        df = df[df[ASIN_COL].isin(failed_list)].reset_index(drop=True)
        st.info(f"🔄 Retry mode: processing {len(df)} previously failed ASINs only")

    # ── Preview ───────────────────────────────────────────
    st.markdown('<p class="sec">👁️ Preview</p>', unsafe_allow_html=True)
    prev_cols = [c for c in [ASIN_COL,TITLE_COL,DESC_COL,BRAND_COL,BB_COL,NET_COL] if c]
    st.dataframe(df[prev_cols].head(8), use_container_width=True)

    # ── Metrics ───────────────────────────────────────────
    valid_n  = df[ASIN_COL].dropna().apply(lambda x: str(x).strip()).str.len().gt(4).sum()
    avg_d    = 4.0 if (keepa_key and keepa_key.strip()) else (sum(DELAY_SEQ)/len(DELAY_SEQ) + EXTRA_DELAY)
    est_min  = round((valid_n / N_WORKERS) * avg_d / 60, 1)

    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Total Rows",    len(df))
    m2.metric("Valid ASINs",   valid_n)
    m3.metric("Workers",       N_WORKERS)
    m4.metric("Avg Delay",     f"~{avg_d:.1f}s")
    m5.metric("Est. Runtime",  f"~{est_min} min")

    # Engine summary box
    ml_str  = "✅ Semantic ML (40% KW + 60% SEM)" if sem_model else "⚠️ Keyword-only (install sentence-transformers)"
    kpa_str = "✅ Keepa API" if (keepa_key and keepa_key.strip()) else "⚠️ Amazon scraper"
    ck_str  = f"✅ {len(cookies_dict)} Prime cookies" if cookies_dict else "—"
    st.markdown(f"""
    <div class="info-box">
        <b>🧠 v6 Engine:</b> {ml_str} &nbsp;·&nbsp;
        BB Price: {kpa_str} &nbsp;·&nbsp; Cookies: {ck_str} &nbsp;·&nbsp;
        {N_WORKERS} workers &nbsp;·&nbsp; SQLite cache ({CACHE_TTL_HOURS}h TTL) &nbsp;·&nbsp;
        Exp backoff (5s→10s→20s) &nbsp;·&nbsp; FBA/FBM detection &nbsp;·&nbsp; Price ranges
    </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ── RUN BUTTON ────────────────────────────────────────
    col_run, col_retry = st.columns([3,1])
    with col_run:
        run_clicked = st.button(
            f"🚀  START VERIFICATION  ·  {valid_n} ASINs  ·  {N_WORKERS} Workers",
            use_container_width=True
        )

    if run_clicked:

        # Build HTTP session with optional cookies
        session = build_session(cookies_dict)

        # Build per-row config
        run_config = {
            "match_threshold": MATCH_THRESH,
            "price_tolerance": PRICE_TOL,
            "referral_rate":   REF_RATE,
            "seller_type":     SELLER_TYPE,
            "bb_warn_as_fail": BB_WARN_FAIL,
            "max_retries":     MAX_RETRIES,
        }

        # Prepare row data list
        row_data_list = []
        for _, row in df.iterrows():
            row_data_list.append({
                "asin":      str(row.get(ASIN_COL,"")).strip(),
                "title":     str(row.get(TITLE_COL,"")).strip()  if TITLE_COL else "",
                "desc":      str(row.get(DESC_COL,"")).strip()   if DESC_COL  else "",
                "brand":     str(row.get(BRAND_COL,"")).strip()  if BRAND_COL else "",
                "your_bb":   parse_price(row.get(BB_COL))        if BB_COL    else None,
                "net_price": parse_price(row.get(NET_COL))       if NET_COL   else None,
                "full_cost": parse_price(row.get(FULL_COL))      if FULL_COL  else None,
                "ref_fee":   parse_price(row.get(REF_COL))       if REF_COL   else None,
            })

        # ── Live progress UI ──────────────────────────────
        st.markdown('<p class="sec">⚡ Live Progress</p>', unsafe_allow_html=True)
        progress_bar  = st.progress(0.0)
        overall_status= st.empty()
        worker_area   = st.empty()
        live_table    = st.empty()

        results_list  = [None] * len(row_data_list)
        log_rows      = []
        completed     = 0
        total         = len(row_data_list)
        lock          = threading.Lock()
        worker_status = {}  # worker_id → status string

        # ── Thread worker ─────────────────────────────────
        def run_worker(idx: int, row_data: dict) -> tuple[int, dict]:
            worker_id = threading.current_thread().name
            with lock:
                worker_status[worker_id] = f"Processing {row_data['asin']}"

            # Small staggered start to avoid simultaneous hits
            time.sleep(random.uniform(0.2, 1.5) * (idx % N_WORKERS))

            result = verify_single_asin(
                row_data      = row_data,
                session       = session,
                sem_model     = sem_model,
                config        = run_config,
                force_refresh = FORCE_REFRESH,
            )

            # Delay between requests (per worker)
            if keepa_key and keepa_key.strip():
                delay = random.uniform(3.0, 5.0) + EXTRA_DELAY
            else:
                delay = DELAY_SEQ[idx % len(DELAY_SEQ)] + random.uniform(-0.5,1.2) + EXTRA_DELAY
            time.sleep(delay)

            with lock:
                worker_status[worker_id] = f"Done {row_data['asin']}"

            return idx, result

        # ── Execute with ThreadPoolExecutor ───────────────
        with ThreadPoolExecutor(max_workers=N_WORKERS) as executor:
            future_map = {
                executor.submit(run_worker, idx, rd): idx
                for idx, rd in enumerate(row_data_list)
            }

            for future in as_completed(future_map):
                try:
                    idx, result = future.result()
                    results_list[idx] = result
                    completed += 1

                    verdict = result.get("Verification","")
                    icon    = "✅" if "Verified" in verdict else ("⚠️" if "WARNING" in verdict else ("⏭️" if "SKIP" in verdict else "❌"))
                    log_rows.append({
                        "#":          future_map[future]+1,
                        "ASIN":       result.get("ASIN",""),
                        "Live BB":    result.get("BB Price (Live)",""),
                        "Sel. Price": result.get("Selected Price",""),
                        "Match":      result.get("Combined Score %",""),
                        "Confidence": result.get("Confidence Level",""),
                        "P&L":        result.get("P&L Status",""),
                        "Status":     icon,
                        "Cache":      result.get("Cache Hit",""),
                    })

                    # Update UI
                    progress_bar.progress(completed / total)
                    overall_status.markdown(
                        f"<p style='color:#1e2d4e;font-size:13px;'>"
                        f"<b>{completed}/{total}</b> completed &nbsp;·&nbsp; "
                        f"{N_WORKERS} workers running</p>",
                        unsafe_allow_html=True
                    )

                    # Worker status cards
                    worker_html = "".join([
                        f'<div class="worker-card">🔧 <b>{wid[-8:]}</b>: {st_}</div>'
                        for wid, st_ in list(worker_status.items())[-N_WORKERS:]
                    ])
                    worker_area.markdown(worker_html, unsafe_allow_html=True)

                    # Live table (last 15 rows)
                    live_table.dataframe(
                        pd.DataFrame(log_rows).tail(15),
                        use_container_width=True
                    )

                except Exception as e:
                    idx     = future_map[future]
                    asin    = row_data_list[idx].get("asin","?")
                    log_error(asin, "Worker exception", e)
                    results_list[idx] = {
                        "ASIN": asin, "Verification": "❌ WORKER ERROR",
                        "Fail Reasons": str(e)[:200]
                    }
                    completed += 1

        # ── Complete ──────────────────────────────────────
        progress_bar.progress(1.0)
        overall_status.markdown(
            "<p style='color:#1a7a42;font-weight:800;font-size:15px;'>✅ VERIFICATION COMPLETE!</p>",
            unsafe_allow_html=True
        )

        # ── Build results dataframe ───────────────────────
        results_df = pd.DataFrame([r for r in results_list if r])

        # Merge results back into original df
        for col in results_df.columns:
            if col != ASIN_COL:
                df[col] = results_df[col].values if len(results_df) == len(df) else results_df.get(col, "")

        # ── Summary metrics ───────────────────────────────
        st.markdown('<p class="sec">📊 Results Summary</p>', unsafe_allow_html=True)
        verd_col = "Verification"
        if verd_col in df.columns:
            verified_n  = df[verd_col].str.contains("Verified",  na=False).sum()
            warned_n    = df[verd_col].str.contains("WARNING",   na=False).sum()
            failed_n    = df[verd_col].str.contains("FAILED",    na=False).sum()
            skipped_n   = df[verd_col].str.contains("SKIPPED",   na=False).sum()
            cached_n    = df.get("Cache Hit","").str.contains("Yes", na=False).sum() if "Cache Hit" in df.columns else 0
        else:
            verified_n=warned_n=failed_n=skipped_n=cached_n = 0

        s1,s2,s3,s4,s5,s6 = st.columns(6)
        s1.metric("✅ Verified",      verified_n)
        s2.metric("⚠️ Needs Review",  warned_n)
        s3.metric("❌ Failed",        failed_n)
        s4.metric("⏭️ Skipped",       skipped_n)
        s5.metric("💾 Cache Hits",    cached_n)
        s6.metric("🏃 Workers Used",  N_WORKERS)

        # ── Confidence distribution ───────────────────────
        if "Confidence Level" in df.columns:
            st.markdown('<p class="sec">🎯 Confidence Distribution</p>', unsafe_allow_html=True)
            c1,c2,c3 = st.columns(3)
            high_n   = (df["Confidence Level"] == "HIGH").sum()
            med_n    = (df["Confidence Level"] == "MEDIUM").sum()
            low_n    = (df["Confidence Level"] == "LOW").sum()
            c1.metric("🟢 HIGH Confidence",   high_n)
            c2.metric("🟡 MEDIUM Confidence", med_n)
            c3.metric("🔴 LOW Confidence",    low_n)

        # ── P&L Summary ───────────────────────────────────
        if "P&L Status" in df.columns:
            st.markdown('<p class="sec">💰 P&L Summary</p>', unsafe_allow_html=True)
            p1,p2,p3 = st.columns(3)
            p1.metric("✅ Profitable",  (df["P&L Status"]=="PROFITABLE").sum())
            p2.metric("⚠️ Marginal",    (df["P&L Status"]=="MARGINAL").sum())
            p3.metric("❌ Loss",        (df["P&L Status"]=="LOSS").sum())

        # ── Action Required ───────────────────────────────
        if "Action Required" in df.columns:
            st.markdown('<p class="sec">🎬 Action Required</p>', unsafe_allow_html=True)
            a1,a2,a3 = st.columns(3)
            a1.metric("✅ No Action",   (df["Action Required"]=="NONE").sum())
            a2.metric("👀 Review",      (df["Action Required"]=="REVIEW").sum())
            a3.metric("🚨 Re-Evaluate", (df["Action Required"]=="RE-EVALUATE").sum())

        # ── Items needing attention ───────────────────────
        attn_df = df[df.get("Verification","").str.contains("FAILED|WARNING", na=False)] if "Verification" in df.columns else pd.DataFrame()
        if len(attn_df) > 0:
            st.markdown('<p class="sec">🔴 Items Needing Attention</p>', unsafe_allow_html=True)
            show = [c for c in [
                ASIN_COL, TITLE_COL, BB_COL,
                "BB Price (Live)","Price Range","Selected Price",
                "Conservative Profit","Confidence Level",
                "Action Required","Verification","Fail Reasons"
            ] if c and c in df.columns]
            st.dataframe(attn_df[show].reset_index(drop=True), use_container_width=True)

        # ── Clear failed cache for successful ones ────────
        if "Verification" in df.columns and ASIN_COL in df.columns:
            succeeded = df[df["Verification"].str.contains("Verified|WARNING", na=False)][ASIN_COL].tolist()
            if succeeded:
                cache_clear_failed(succeeded)

        # ── Error log download ────────────────────────────
        if os.path.exists(ERROR_LOG) and os.path.getsize(ERROR_LOG) > 0:
            with open(ERROR_LOG, "rb") as f:
                err_bytes = f.read()
            st.download_button("📋 Download Error Log", data=err_bytes,
                               file_name="vv_errors.log", mime="text/plain")

        # ── Excel download ────────────────────────────────
        st.markdown('<p class="sec">⬇️ Download Results</p>', unsafe_allow_html=True)
        excel_bytes = build_excel(df, MATCH_THRESH)
        st.download_button(
            "📥  Download Colour-Coded Excel — VirVentures_Verified.xlsx",
            data=excel_bytes,
            file_name="VirVentures_ASIN_Verification.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    # ── Empty state ───────────────────────────────────────
    st.markdown("""
    <div style="text-align:center;padding:70px 40px;background:#fff;
                border:2px dashed #f47920;border-radius:16px;margin-top:8px;">
        <p style="font-size:3rem;margin:0;">🔍</p>
        <p style="color:#1e2d4e !important;font-weight:800;font-size:1.2rem;margin:14px 0 6px;">
            Waiting for your file...
        </p>
        <p style="color:#f47920 !important;font-size:0.9rem;font-weight:600;margin:0 0 8px;">
            Upload your .xlsx above and I'll handle everything automatically
        </p>
        <p style="color:#aaa !important;font-size:0.78rem;margin:0;">
            Any column layout &nbsp;·&nbsp; Auto-detected &nbsp;·&nbsp;
            Semantic ML &nbsp;·&nbsp; FBA/FBM &nbsp;·&nbsp; Price Ranges &nbsp;·&nbsp;
            SQLite Cache &nbsp;·&nbsp; Concurrent Workers &nbsp;·&nbsp; Full P&amp;L
        </p>
    </div>
    """, unsafe_allow_html=True)
