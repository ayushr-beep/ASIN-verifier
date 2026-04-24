import streamlit as st
import pandas as pd
import requests
import re
import random
import time
import io
import base64
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ── Page config ────────────────────────────────────────────
st.set_page_config(
    page_title="VirVentures ASIN Verifier",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS (Orange & White Theme) ──────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background-color: #ffffff; color: #333; }
section[data-testid="stSidebar"] { background: #f8f9fa !important; border-right: 1px solid #e0e0e0; }
h1, h2, h3 { color: #f37021 !important; }

/* Buttons */
.stButton > button { 
    background: #f37021 !important; 
    color: #ffffff !important; 
    border-radius: 6px !important; 
    border: none !important;
    font-weight: 700 !important;
}
.stDownloadButton > button { 
    background: #ffffff !important; 
    color: #f37021 !important; 
    border: 2px solid #f37021 !important; 
}

/* Dataframe styling */
.stDataFrame { border: 1px solid #eee !important; border-radius: 8px; }

/* Header and Logo Container */
.header-container {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 1rem 0;
    border-bottom: 3px solid #f37021;
    margin-bottom: 2rem;
}
</style>
""", unsafe_allow_html=True)

# ── Helper for Logo ────────────────────────────────────────
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# ══════════════════════════════════════════════════════════
# Constants & Scraper Logic
# ══════════════════════════════════════════════════════════

HEADERS_POOL = [
    {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0", "Referer": "https://www.google.com/"},
    {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) Chrome/123.0.0.0", "Referer": "https://www.bing.com/"}
]

STOPWORDS = {"a","an","the","and","or","for","of","in","to","with","by","is","it","its","–","-","&"}

def fetch_page(asin):
    try:
        r = requests.get(f"https://www.amazon.com/dp/{asin}", headers=random.choice(HEADERS_POOL), timeout=15)
        if "captcha" in r.text.lower() or "robot check" in r.text.lower(): return "BLOCKED"
        return BeautifulSoup(r.text, "lxml") if r.status_code == 200 else None
    except: return None

def get_live_bb_price(soup):
    if soup == "BLOCKED": return None, "BLOCKED"
    selectors = ["#corePriceDisplay_desktop_feature_div .a-price-whole", ".a-price .a-offscreen", "#price_inside_buybox"]
    for sel in selectors:
        tag = soup.select_one(sel)
        if tag:
            raw = re.sub(r"[^\d.]", "", tag.get_text(strip=True))
            if raw: return float(raw), f"${float(raw):.2f}"
    return None, "N/A"

def get_amz_full_text(soup):
    if soup == "BLOCKED": return ""
    parts = [t.get_text() for t in soup.select("#productTitle, #feature-bullets li span, #productDescription")]
    return " ".join(parts).lower()

def keyword_match(our_text, amz_text):
    words = [w.lower() for w in re.findall(r"[a-zA-Z0-9]+", str(our_text)) if w.lower() not in STOPWORDS and len(w) > 1]
    if not words or not amz_text: return 0.0, "0/0"
    hit = sum(1 for w in words if w in amz_text)
    return hit / len(words), f"{hit}/{len(words)}"

# ══════════════════════════════════════════════════════════
# UI Header with Logo
# ══════════════════════════════════════════════════════════

try:
    logo_base64 = get_base64_image("virventures_com_logo.jpg")
    st.markdown(f"""
    <div class="header-container">
        <div>
            <h1 style="margin:0;">🔍 ASIN VERIFIER PRO</h1>
            <p style="margin:0; color:#666;">VirVentures Inventory Authentication Tool</p>
        </div>
        <img src="data:image/jpeg;base64,{logo_base64}" width="180">
    </div>
    """, unsafe_allow_html=True)
except:
    st.title("🔍 VIRVENTURES ASIN VERIFIER")

# ══════════════════════════════════════════════════════════
# Main Logic
# ══════════════════════════════════════════════════════════

with st.sidebar:
    st.header("⚙️ Verifier Settings")
    ASIN_COL = st.text_input("ASIN Column", "Output ASIN")
    BB_COL = st.text_input("Price Column", "BB Price")
    MATCH_MIN = st.slider("Min Match %", 10, 90, 40) / 100
    DELAY = st.slider("Request Delay (sec)", 2, 15, 8)

uploaded_file = st.file_uploader("Upload Product File (Excel)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype=str)
    
    if st.button("🚀 START VERIFICATION"):
        res_price, res_match, res_verdict, res_reason = [], [], [], []
        progress = st.progress(0)
        
        for i, row in df.iterrows():
            asin = str(row.get(ASIN_COL, "")).strip()
            # Combine all previous box details for matching
            our_combined = f"{row.get('Title','')} {row.get('Description','')} {row.get('Brand','')}"
            
            if len(asin) < 5:
                res_price.append("N/A"); res_match.append("0%"); res_verdict.append("SKIPPED"); res_reason.append("No ASIN")
            else:
                soup = fetch_page(asin)
                l_price, l_price_str = get_live_bb_price(soup)
                score, _ = keyword_match(our_combined, get_amz_full_text(soup))
                
                match_pct = f"{score*100:.1f}%"
                issues = []
                if soup == "BLOCKED": issues.append("Amazon Blocked (Bot Detected)")
                elif score < MATCH_MIN: issues.append(f"Low match ({match_pct})")
                if l_price_str == "N/A": issues.append("No live price found")
                
                res_price.append(l_price_str)
                res_match.append(match_pct)
                res_verdict.append("❌ FAILED" if issues else "✅ Verified")
                res_reason.append(" | ".join(issues))
            
            progress.progress((i + 1) / len(df))
            time.sleep(random.uniform(DELAY * 0.8, DELAY * 1.2))

        df["Live Price"] = res_price
        df["Match %"] = res_match
        df["Verification"] = res_verdict
        df["Fail Reasons"] = res_reason
        
        st.success("Verification Complete!")
        
        # Display failed items with ALL original details (UPC, Description, etc.)
        st.subheader("🔴 Flagged Items (Requires Review)")
        failed_df = df[df["Verification"] == "❌ FAILED"]
        st.dataframe(failed_df, use_container_width=True)
        
        # Download
        output = io.BytesIO()
        df.to_excel(output, index=False)
        st.download_button("📥 Download Full Verified Report", output.getvalue(), "VirVentures_Report.xlsx")
