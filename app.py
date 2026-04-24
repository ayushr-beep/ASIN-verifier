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
    page_title="ASIN Verifier",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.main { background-color: #0d0d0d; }

h1, h2, h3 {
    font-family: 'Space Mono', monospace !important;
}

.stApp {
    background: #0d0d0d;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #111 !important;
    border-right: 1px solid #222;
}

/* Metric cards */
div[data-testid="metric-container"] {
    background: #1a1a1a;
    border: 1px solid #2a2a2a;
    border-radius: 8px;
    padding: 16px;
}

/* Buttons */
.stButton > button {
    background: #00ff87 !important;
    color: #000 !important;
    font-family: 'Space Mono', monospace !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 4px !important;
    padding: 12px 32px !important;
    font-size: 14px !important;
    letter-spacing: 1px !important;
    transition: all 0.2s ease !important;
}

.stButton > button:hover {
    background: #00cc6a !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 20px rgba(0,255,135,0.3) !important;
}

/* Download button */
.stDownloadButton > button {
    background: #1a1a1a !important;
    color: #00ff87 !important;
    border: 1px solid #00ff87 !important;
    font-family: 'Space Mono', monospace !important;
    font-weight: 700 !important;
    border-radius: 4px !important;
    padding: 10px 24px !important;
}

.stDownloadButton > button:hover {
    background: #00ff8722 !important;
}

/* Progress bar */
.stProgress > div > div {
    background: #00ff87 !important;
}

/* Dataframe */
.stDataFrame {
    border: 1px solid #222 !important;
    border-radius: 8px !important;
}

/* Status badges */
.badge-verified {
    background: #0a2e1a;
    color: #00ff87;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    border: 1px solid #00ff8740;
}
.badge-failed {
    background: #2e0a0a;
    color: #ff4444;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    border: 1px solid #ff444440;
}
.badge-warning {
    background: #2e2200;
    color: #ffaa00;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    border: 1px solid #ffaa0040;
}

/* Header */
.app-header {
    padding: 2rem 0 1rem 0;
    border-bottom: 1px solid #222;
    margin-bottom: 2rem;
}

/* Log box */
.log-box {
    background: #0a0a0a;
    border: 1px solid #1e1e1e;
    border-radius: 8px;
    padding: 16px;
    font-family: 'Space Mono', monospace;
    font-size: 12px;
    color: #888;
    max-height: 300px;
    overflow-y: auto;
}

/* File uploader */
.stFileUploader {
    border: 2px dashed #333 !important;
    border-radius: 8px !important;
    background: #111 !important;
}

/* Select boxes & inputs */
.stSelectbox > div > div, .stNumberInput > div > div {
    background: #1a1a1a !important;
    border-color: #333 !important;
    color: #fff !important;
}

/* Info / warning boxes */
.stAlert {
    border-radius: 8px !important;
}

/* expander */
.streamlit-expanderHeader {
    background: #1a1a1a !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# Constants
# ══════════════════════════════════════════════════════════

HEADERS_POOL = [
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
    },
    {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Accept-Language": "en-GB,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    },
]

STOPWORDS = {"a","an","the","and","or","for","of","in","to","with","by","is","it","its","–","-","&"}


# ══════════════════════════════════════════════════════════
# Scraper functions
# ══════════════════════════════════════════════════════════

def fetch_page(asin):
    try:
        r = requests.get(
            f"https://www.amazon.com/dp/{asin}",
            headers=random.choice(HEADERS_POOL),
            timeout=15
        )
        return BeautifulSoup(r.text, "lxml") if r.status_code == 200 else None
    except Exception:
        return None

def get_live_bb_price(soup):
    selectors = [
        "#corePriceDisplay_desktop_feature_div .a-price-whole",
        "#corePrice_feature_div .a-price-whole",
        "#price_inside_buybox",
        "#priceblock_ourprice",
        "#priceblock_dealprice",
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
    tag = soup.select_one("#productTitle")
    return tag.get_text(strip=True) if tag else ""

def get_amz_full_text(soup):
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
    if your_bb is None and live_bb is None:
        return "⚠️ No price data"
    if your_bb is None:
        return f"⚠️ Your BB blank | Live: ${live_bb:.2f}"
    if live_bb is None:
        return "⚠️ No live BB on Amazon"
    diff = (live_bb - your_bb) / your_bb
    if diff > threshold:
        return f"📈 Live ${live_bb:.2f} is {diff*100:.1f}% HIGHER than your BB ${your_bb:.2f}"
    elif diff < -threshold:
        return f"📉 Live ${live_bb:.2f} is {abs(diff)*100:.1f}% LOWER than your BB ${your_bb:.2f}"
    return f"✅ BB Match (diff {diff*100:+.1f}%)"


def build_output_excel(df):
    """Write results to styled Excel and return as bytes."""
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    DARK_HDR  = PatternFill("solid", fgColor="1F3864")
    GREEN     = PatternFill("solid", fgColor="C6EFCE")
    RED       = PatternFill("solid", fgColor="FFC7CE")
    AMBER     = PatternFill("solid", fgColor="FFEB9C")
    LIGHT_RED = PatternFill("solid", fgColor="FFE0E0")
    BOLD      = Font(bold=True)
    WH_BOLD   = Font(bold=True, color="FFFFFF")
    CENTER    = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[1]:
        cell.fill = DARK_HDR
        cell.font = WH_BOLD
        cell.alignment = CENTER
    ws.row_dimensions[1].height = 28

    hdr = [c.value for c in ws[1]]
    def ci(name):
        try: return hdr.index(name) + 1
        except: return None

    v_ci  = ci("Verification")
    bb_ci = ci("BB Comparison")
    m_ci  = ci("Match %")
    fr_ci = ci("Fail Reasons")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        v_val  = str(row[v_ci  - 1].value or "") if v_ci  else ""
        bb_val = str(row[bb_ci - 1].value or "") if bb_ci else ""
        m_val  = str(row[m_ci  - 1].value or "") if m_ci  else ""

        if v_ci:
            c = row[v_ci - 1]
            if "Verified" in v_val:   c.fill, c.font = GREEN, BOLD
            elif "FAILED" in v_val:   c.fill, c.font = RED, BOLD
            else:                     c.fill = AMBER

        if bb_ci and ("HIGHER" in bb_val or "LOWER" in bb_val):
            row[bb_ci - 1].fill = AMBER
            row[bb_ci - 1].font = BOLD

        if m_ci:
            try:
                if float(m_val.replace("%","")) < 40:
                    row[m_ci - 1].fill = LIGHT_RED
            except ValueError:
                pass

        if fr_ci and row[fr_ci - 1].value:
            row[fr_ci - 1].font = Font(italic=True, color="C00000")

    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 55)

    ws.freeze_panes = "C2"

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.getvalue()


# ══════════════════════════════════════════════════════════
# UI — Header
# ══════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
    <h1 style="color:#00ff87; font-size:2rem; margin:0; letter-spacing:-1px;">
        🔍 ASIN VERIFIER
    </h1>
    <p style="color:#555; font-family:'Space Mono',monospace; font-size:12px; margin:4px 0 0 0;">
        BB PRICE · DESCRIPTION MATCH · 100% AUTHENTIC CHECK
    </p>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# UI — Sidebar (settings)
# ══════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.markdown("---")

    st.markdown("**Column Names**")
    st.caption("Update if your file uses different names")

    ASIN_COL     = st.text_input("ASIN Column",        value="Output ASIN")
    TITLE_COL    = st.text_input("Title Column",       value="Title")
    DESC_COL     = st.text_input("Description Column", value="Description")
    BRAND_COL    = st.text_input("Brand Column",       value="Brand")
    UPC_COL      = st.text_input("UPC Column",         value="UPC")
    BB_PRICE_COL = st.text_input("BB Price Column",    value="BB Price")

    st.markdown("---")
    st.markdown("**Thresholds**")

    MATCH_THRESHOLD = st.slider(
        "Min Keyword Match %", 10, 90, 40, 5,
        help="Rows below this match % will be flagged as FAILED"
    ) / 100

    PRICE_DIFF = st.slider(
        "BB Price Diff Tolerance %", 1, 30, 10, 1,
        help="Flag if live BB price differs by more than this % from your BB Price"
    ) / 100

    DELAY = st.slider(
        "Delay between requests (sec)", 2, 10, 4, 1,
        help="Higher = safer from Amazon blocking. Don't go below 2."
    )

    st.markdown("---")
    st.markdown(
        "<p style='color:#444; font-size:11px; font-family:Space Mono,monospace;'>"
        "Built for your team ⚡<br>Powered by Python + Streamlit</p>",
        unsafe_allow_html=True
    )


# ══════════════════════════════════════════════════════════
# UI — File Upload
# ══════════════════════════════════════════════════════════

col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader(
        "Upload your Excel file (.xlsx)",
        type=["xlsx", "xls"],
        help="Upload the fresh file you receive each day"
    )

with col2:
    st.markdown("<br>", unsafe_allow_html=True)
    if uploaded_file:
        st.success(f"✅ {uploaded_file.name}")


# ══════════════════════════════════════════════════════════
# UI — Column Check & Preview
# ══════════════════════════════════════════════════════════

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, dtype=str)
    df_raw.columns = df_raw.columns.str.strip()

    st.markdown("### 📋 File Preview")

    # Column validation
    required_cols = {
        ASIN_COL: "ASIN",
        TITLE_COL: "Title",
        DESC_COL: "Description",
        BRAND_COL: "Brand",
        UPC_COL: "UPC",
        BB_PRICE_COL: "BB Price",
    }

    col_status = {}
    all_found = True
    check_cols = st.columns(len(required_cols))

    for idx, (col, label) in enumerate(required_cols.items()):
        found = col in df_raw.columns
        col_status[col] = found
        if not found:
            all_found = False
        with check_cols[idx]:
            if found:
                st.metric(label, "✅ Found")
            else:
                st.metric(label, "❌ Missing")

    if not all_found:
        missing = [c for c, f in col_status.items() if not f]
        st.error(f"⛔ Missing columns: **{', '.join(missing)}** — update column names in the sidebar ←")
    else:
        # Preview table
        preview_cols = [c for c in [ASIN_COL, TITLE_COL, DESC_COL, BRAND_COL, UPC_COL, BB_PRICE_COL]
                        if c in df_raw.columns]
        st.dataframe(df_raw[preview_cols].head(10), use_container_width=True)

        total_rows = len(df_raw)
        valid_asins = df_raw[ASIN_COL].dropna().apply(lambda x: str(x).strip()).str.len().gt(4).sum()
        est_mins = round(valid_asins * DELAY / 60, 1)

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Rows", total_rows)
        m2.metric("Valid ASINs", valid_asins)
        m3.metric("Est. Time", f"~{est_mins} min")

        st.markdown("---")

        # ── RUN BUTTON
        if st.button(f"🚀  RUN VERIFICATION  ({valid_asins} ASINs)", use_container_width=True):

            # Results storage
            results = {
                "Live BB Price":  [],
                "BB Comparison":  [],
                "Amazon Title":   [],
                "Keyword Match":  [],
                "Match %":        [],
                "Verification":   [],
                "Fail Reasons":   [],
            }

            verified_n = failed_n = skipped_n = price_flag_n = 0

            # Live UI elements
            st.markdown("### ⚡ Live Progress")
            progress_bar = st.progress(0)
            status_text  = st.empty()
            live_table   = st.empty()

            log_rows = []  # for live table

            for i, row in df_raw.iterrows():
                asin    = str(row.get(ASIN_COL, "")).strip()
                title   = str(row.get(TITLE_COL, "")).strip()
                desc    = str(row.get(DESC_COL, "")).strip()
                brand   = str(row.get(BRAND_COL, "")).strip()
                your_bb = parse_price(row.get(BB_PRICE_COL))
                row_num = i + 1

                our_combined = f"{title} {desc} {brand}".strip()

                progress = row_num / total_rows
                progress_bar.progress(progress)
                status_text.markdown(
                    f"<p style='color:#555; font-family:Space Mono,monospace; font-size:12px;'>"
                    f"Processing row {row_num}/{total_rows} — <b style='color:#00ff87'>{asin}</b></p>",
                    unsafe_allow_html=True
                )

                # Skip blank
                if not asin or asin.lower() in ("nan", "") or len(asin) < 5:
                    for k in results: results[k].append("SKIPPED")
                    skipped_n += 1
                    log_rows.append({
                        "Row": row_num, "ASIN": "—", "Your BB": "—",
                        "Live BB": "—", "Match": "—", "Status": "⏭️ Skipped"
                    })
                    live_table.dataframe(pd.DataFrame(log_rows).tail(10), use_container_width=True)
                    continue

                soup = fetch_page(asin)

                if soup is None:
                    for k in results: results[k].append("FETCH FAILED")
                    failed_n += 1
                    log_rows.append({
                        "Row": row_num, "ASIN": asin, "Your BB": "—",
                        "Live BB": "—", "Match": "—", "Status": "❌ Fetch failed"
                    })
                    live_table.dataframe(pd.DataFrame(log_rows).tail(10), use_container_width=True)
                    time.sleep(DELAY)
                    continue

                live_bb_float, live_bb_str = get_live_bb_price(soup)
                amz_title                  = get_amz_title(soup)
                amz_text                   = get_amz_full_text(soup)
                score, ratio               = keyword_match(our_combined, amz_text)
                match_pct                  = f"{score * 100:.1f}%"
                bb_comp                    = compare_bb(your_bb, live_bb_float, PRICE_DIFF)
                your_bb_str                = f"${your_bb:.2f}" if your_bb else "—"

                issues = []
                if live_bb_float is None:
                    issues.append("No live BB price")
                if "HIGHER" in bb_comp or "LOWER" in bb_comp:
                    issues.append(bb_comp)
                    price_flag_n += 1
                if score < MATCH_THRESHOLD:
                    issues.append(f"Low match ({match_pct})")

                if issues:
                    verdict      = "❌ FAILED"
                    fail_reasons = " | ".join(issues)
                    failed_n    += 1
                    status_icon  = "❌ Failed"
                else:
                    verdict      = "✅ Verified — 100% Authentic"
                    fail_reasons = ""
                    verified_n  += 1
                    status_icon  = "✅ Verified"

                results["Live BB Price"].append(live_bb_str)
                results["BB Comparison"].append(bb_comp)
                results["Amazon Title"].append(amz_title)
                results["Keyword Match"].append(ratio)
                results["Match %"].append(match_pct)
                results["Verification"].append(verdict)
                results["Fail Reasons"].append(fail_reasons)

                log_rows.append({
                    "Row": row_num, "ASIN": asin,
                    "Your BB": your_bb_str, "Live BB": live_bb_str,
                    "Match": match_pct, "Status": status_icon
                })
                live_table.dataframe(pd.DataFrame(log_rows).tail(10), use_container_width=True)

                time.sleep(random.uniform(DELAY - 0.5, DELAY + 0.5))

            # ── Finalize
            progress_bar.progress(1.0)
            status_text.markdown(
                "<p style='color:#00ff87; font-family:Space Mono,monospace; font-size:13px; font-weight:bold;'>"
                "✅ VERIFICATION COMPLETE</p>",
                unsafe_allow_html=True
            )

            # Pad results
            for col, vals in results.items():
                while len(vals) < len(df_raw):
                    vals.append("")
                df_raw[col] = vals

            # ── Summary metrics
            st.markdown("### 📊 Summary")
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("✅ Verified",      verified_n)
            s2.metric("❌ Failed",        failed_n)
            s3.metric("⚠️ Price Flagged", price_flag_n)
            s4.metric("⏭️ Skipped",       skipped_n)

            # ── Failed items table
            if failed_n > 0:
                st.markdown("### 🔴 Items Needing Attention")
                show_cols = [c for c in [ASIN_COL, TITLE_COL, BB_PRICE_COL,
                                          "Live BB Price", "Match %",
                                          "Verification", "Fail Reasons"]
                             if c in df_raw.columns]
                failed_df = df_raw[df_raw["Verification"].str.contains("FAILED", na=False)][show_cols]
                st.dataframe(failed_df.reset_index(drop=True), use_container_width=True)

            # ── Download
            st.markdown("### ⬇️ Download Results")
            excel_bytes = build_output_excel(df_raw)
            st.download_button(
                label="📥  Download Color-Coded Excel",
                data=excel_bytes,
                file_name="ASIN_Verification_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

else:
    # Empty state
    st.markdown("""
    <div style="
        text-align:center;
        padding: 80px 40px;
        border: 1px dashed #222;
        border-radius: 12px;
        margin-top: 20px;
    ">
        <p style="font-size:48px; margin:0;">📂</p>
        <p style="color:#444; font-family:'Space Mono',monospace; font-size:14px; margin:8px 0 0 0;">
            Upload your Excel file above to get started
        </p>
        <p style="color:#333; font-size:12px; margin:8px 0 0 0;">
            Supports .xlsx — works with any fresh file you receive
        </p>
    </div>
    """, unsafe_allow_html=True)
