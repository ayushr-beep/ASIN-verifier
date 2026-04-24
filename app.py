import streamlit as st
import pandas as pd
import requests
import re
import random
import time
import io
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ── Page config ────────────────────────────────────────────
st.set_page_config(
    page_title="ASIN Verifier Pro",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS (Kept Original Styling) ──────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.main { background-color: #0d0d0d; }
h1, h2, h3 { font-family: 'Space Mono', monospace !important; }
.stApp { background: #0d0d0d; }
section[data-testid="stSidebar"] { background: #111 !important; border-right: 1px solid #222; }
div[data-testid="metric-container"] { background: #1a1a1a; border: 1px solid #2a2a2a; border-radius: 8px; padding: 16px; }
.stButton > button { background: #00ff87 !important; color: #000 !important; font-family: 'Space Mono', monospace !important; font-weight: 700 !important; border: none !important; border-radius: 4px !important; padding: 12px 32px !important; font-size: 14px !important; letter-spacing: 1px !important; transition: all 0.2s ease !important; }
.stButton > button:hover { background: #00cc6a !important; transform: translateY(-1px); box-shadow: 0 4px 20px rgba(0,255,135,0.3) !important; }
.stDownloadButton > button { background: #1a1a1a !important; color: #00ff87 !important; border: 1px solid #00ff87 !important; font-family: 'Space Mono', monospace !important; font-weight: 700 !important; border-radius: 4px !important; padding: 10px 24px !important; }
.stDownloadButton > button:hover { background: #00ff8722 !important; }
.stProgress > div > div { background: #00ff87 !important; }
.stDataFrame { border: 1px solid #222 !important; border-radius: 8px !important; }
.app-header { padding: 2rem 0 1rem 0; border-bottom: 1px solid #222; margin-bottom: 2rem; }
.stFileUploader { border: 2px dashed #333 !important; border-radius: 8px !important; background: #111 !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# TWEAKED: Advanced Anti-Detection Constants
# ══════════════════════════════════════════════════════════

HEADERS_POOL = [
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Referer": "https://www.google.com/",
    },
    {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Referer": "https://www.bing.com/",
    },
    {
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://duckduckgo.com/",
    }
]

STOPWORDS = {"a","an","the","and","or","for","of","in","to","with","by","is","it","its","–","-","&"}

# ══════════════════════════════════════════════════════════
# Scraper functions
# ══════════════════════════════════════════════════════════

def fetch_page(asin):
    try:
        # TWEAK: Using a session to handle cookies if needed
        session = requests.Session()
        r = session.get(
            f"https://www.amazon.com/dp/{asin}",
            headers=random.choice(HEADERS_POOL),
            timeout=15
        )
        
        # TWEAK: Detect CAPTCHA/Block
        if "captcha" in r.text.lower() or "robot check" in r.text.lower():
            return "BLOCKED"
            
        return BeautifulSoup(r.text, "lxml") if r.status_code == 200 else None
    except Exception:
        return None

def get_live_bb_price(soup):
    if soup == "BLOCKED": return None, "BLOCKED"
    
    selectors = [
        "#corePriceDisplay_desktop_feature_div .a-price-whole",
        "#corePrice_feature_div .a-price-whole",
        ".a-price.apexPriceToPay .a-offscreen",
        "#price_inside_buybox",
        "#priceblock_ourprice",
        ".a-price .a-offscreen",
    ]
    for sel in selectors:
        tag = soup.select_one(sel)
        if tag:
            raw = re.sub(r"[^\d.]", "", tag.get_text(strip=True).replace(",", ""))
            if raw:
                try:
                    return float(raw), f"${float(raw):.2f}"
                except ValueError:
                    continue
    return None, "N/A"

def get_amz_title(soup):
    if soup == "BLOCKED": return "BOT DETECTED"
    tag = soup.select_one("#productTitle")
    return tag.get_text(strip=True) if tag else ""

def get_amz_full_text(soup):
    if soup == "BLOCKED": return ""
    parts = []
    t = soup.select_one("#productTitle")
    if t: parts.append(t.get_text(strip=True))
    for b in soup.select("#feature-bullets li span.a-list-item"):
        parts.append(b.get_text(strip=True))
    d = soup.select_one("#productDescription")
    if d: parts.append(d.get_text(strip=True))
    return " ".join(parts).lower()

def keyword_match(our_text, amz_text):
    if not our_text or not amz_text:
        return 0.0, "0/0"
    words = [w.lower() for w in re.findall(r"[a-zA-Z0-9]+", str(our_text))
             if w.lower() not in STOPWORDS and len(w) > 1]
    if not words:
        return 0.0, "0/0"
    hit = sum(1 for w in words if w in amz_text)
    return hit / len(words), f"{hit}/{len(words)}"

def parse_price(val):
    if not val or str(val).strip().lower() in ("nan", "", "none"):
        return None
    try:
        return float(re.sub(r"[^\d.]", "", str(val)))
    except ValueError:
        return None

def compare_bb(your_bb, live_bb, threshold):
    if live_bb == "BLOCKED": return "🚨 Amazon Blocked Request"
    if your_bb is None and live_bb is None: return "⚠️ No price data"
    if your_bb is None: return f"⚠️ Your BB blank | Live: {live_bb}"
    if live_bb is None or live_bb == "N/A": return "⚠️ No live BB on Amazon"
    
    # Extract float if live_bb is string
    if isinstance(live_bb, str):
        try: live_val = float(re.sub(r"[^\d.]", "", live_bb))
        except: return "⚠️ Price Error"
    else: live_val = live_bb

    diff = (live_val - your_bb) / your_bb
    if diff > threshold: return f"📈 Live ${live_val:.2f} (+{diff*100:.1f}%)"
    elif diff < -threshold: return f"📉 Live ${live_val:.2f} (-{abs(diff)*100:.1f}%)"
    return f"✅ BB Match ({diff*100:+.1f}%)"

# ══════════════════════════════════════════════════════════
# UI & Logic (Keeping your structure)
# ══════════════════════════════════════════════════════════

def build_output_excel(df):
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active
    # (Styles removed for brevity, kept same in final logic)
    final = io.BytesIO()
    wb.save(final)
    return final.getvalue()

st.markdown('<div class="app-header"><h1 style="color:#00ff87; font-size:2rem; margin:0;">🔍 ASIN VERIFIER PRO</h1></div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    ASIN_COL = st.text_input("ASIN Column", value="Output ASIN")
    BB_PRICE_COL = st.text_input("BB Price Column", value="BB Price")
    MATCH_THRESHOLD = st.slider("Min Match %", 10, 90, 40) / 100
    PRICE_DIFF = st.slider("Price Tolerance %", 1, 30, 10) / 100
    DELAY = st.slider("Base Delay (sec)", 2, 15, 8)

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, dtype=str)
    if st.button("🚀 RUN VERIFICATION"):
        results = {"Live BB Price": [], "BB Comparison": [], "Match %": [], "Verification": [], "Fail Reasons": []}
        
        progress_bar = st.progress(0)
        live_table = st.empty()
        log_rows = []

        for i, row in df_raw.iterrows():
            asin = str(row.get(ASIN_COL, "")).strip()
            your_bb = parse_price(row.get(BB_PRICE_COL))
            
            if not asin or len(asin) < 5:
                for k in results: results[k].append("SKIPPED")
                continue

            soup = fetch_page(asin)
            
            if soup == "BLOCKED":
                live_bb_str, match_pct, verdict, reasons = "BLOCKED", "0%", "❌ FAILED", "BOT DETECTED"
            elif soup is None:
                live_bb_str, match_pct, verdict, reasons = "ERROR", "0%", "❌ FAILED", "Fetch Failed"
            else:
                live_val, live_bb_str = get_live_bb_price(soup)
                score, _ = keyword_match(str(row.get("Title", "")) + str(row.get("Description", "")), get_amz_full_text(soup))
                match_pct = f"{score*100:.1f}%"
                bb_comp = compare_bb(your_bb, live_val, PRICE_DIFF)
                
                reasons_list = []
                if score < MATCH_THRESHOLD: reasons_list.append("Low Match")
                if "Live" in bb_comp: reasons_list.append("Price Flag")
                
                reasons = " | ".join(reasons_list)
                verdict = "❌ FAILED" if reasons_list else "✅ Verified"

            results["Live BB Price"].append(live_bb_str)
            results["Match %"].append(match_pct)
            results["Verification"].append(verdict)
            results["Fail Reasons"].append(reasons)
            
            log_rows.append({"ASIN": asin, "Live BB": live_bb_str, "Match": match_pct, "Status": verdict})
            live_table.dataframe(pd.DataFrame(log_rows).tail(5))
            
            # TWEAK: Advanced Jitter delay
            time.sleep(random.uniform(DELAY * 0.7, DELAY * 1.3))

        st.success("Verification Finished!")
