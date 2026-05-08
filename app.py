"""
app.py  —  Energybae Solar Load Calculator Web App
────────────────────────────────────────────────────
Run with:  streamlit run app.py
"""
import os
import sys
import json
import tempfile
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent / "src"))
from bill_extractor import extract_bill_data, fill_excel_template
from create_template import create_template

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Energybae — Solar Load Calculator",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Space Grotesk', sans-serif; }

.main { background: #F1F8E9; }

.hero {
    background: linear-gradient(135deg, #1B5E20 0%, #2E7D32 60%, #4CAF50 100%);
    border-radius: 16px;
    padding: 2.2rem 2rem 1.6rem;
    margin-bottom: 1.5rem;
    color: white;
    box-shadow: 0 4px 24px rgba(27,94,32,0.25);
}
.hero h1 { font-size: 2rem; margin: 0; font-weight: 700; }
.hero p  { font-size: 1rem; margin: 0.5rem 0 0; opacity: 0.88; }

.step-card {
    background: white;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1rem;
    border-left: 4px solid #4CAF50;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.step-card h4 { margin: 0 0 0.4rem; color: #1B5E20; font-size: 1rem; }
.step-card p  { margin: 0; color: #555; font-size: 0.9rem; }

.metric-card {
    background: white;
    border-radius: 10px;
    padding: 1rem;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
}
.metric-card .val  { font-size: 1.6rem; font-weight: 700; color: #1B5E20; }
.metric-card .lbl  { font-size: 0.8rem; color: #777; margin-top: 2px; }

.success-banner {
    background: #E8F5E9;
    border: 1px solid #A5D6A7;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    color: #1B5E20;
    font-weight: 600;
    margin: 1rem 0;
}
.warning-banner {
    background: #FFF8E1;
    border: 1px solid #FFE082;
    border-radius: 10px;
    padding: 0.8rem 1.2rem;
    color: #F57F17;
    font-size: 0.88rem;
    margin: 0.8rem 0;
}
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = Path("templates/solar_load_calculator_template.xlsx")

# ── Ensure template exists ────────────────────────────────────────────────────
if not TEMPLATE_PATH.exists():
    TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    create_template(str(TEMPLATE_PATH))

# ── Hero banner ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <h1>⚡ Energybae Solar Load Calculator</h1>
  <p>Upload your MSEDCL electricity bill → AI extracts your data → Get a filled solar sizing Excel report instantly</p>
</div>
""", unsafe_allow_html=True)

# ── How it works ──────────────────────────────────────────────────────────────
with st.expander("ℹ️ How it works", expanded=False):
    cols = st.columns(4)
    steps = [
        ("1️⃣ Upload", "Upload your MSEDCL electricity bill (PDF or image)"),
        ("2️⃣ AI Reads", "Gemini AI extracts units, tariff, load & billing data"),
        ("3️⃣ Calculate", "Solar system size, costs & ROI are calculated automatically"),
        ("4️⃣ Download", "Get a ready-to-use Excel report with 3 sheets"),
    ]
    for col, (title, desc) in zip(cols, steps):
        col.markdown(f"""<div class="step-card"><h4>{title}</h4><p>{desc}</p></div>""",
                     unsafe_allow_html=True)

st.divider()

# ── API Key input ─────────────────────────────────────────────────────────────
api_key = os.environ.get("GOOGLE_API_KEY", "")
if not api_key:
    api_key = st.text_input(
        "🔑 Google AI API Key",
        type="password",
        placeholder="AIza...",
        help="Get your free key at aistudio.google.com. It stays in your session only."
    )

# ── File uploader ─────────────────────────────────────────────────────────────
st.markdown("### 📄 Upload Electricity Bill")
uploaded_file = st.file_uploader(
    "Drag & drop or browse",
    type=["pdf", "png", "jpg", "jpeg", "webp"],
    help="Supports MSEDCL Maharashtra bills. PDF or photo of the bill."
)

if uploaded_file:
    file_info = f"**{uploaded_file.name}** ({uploaded_file.size / 1024:.1f} KB)"
    st.success(f"✅ File ready: {file_info}")

    # Preview image
    if uploaded_file.type.startswith("image/"):
        with st.expander("Preview uploaded image"):
            st.image(uploaded_file, use_column_width=True)

# ── Process button ────────────────────────────────────────────────────────────
st.markdown("")
process_btn = st.button("⚡ Extract & Generate Solar Report", type="primary",
                        disabled=(not uploaded_file or not api_key),
                        use_container_width=True)

if process_btn:
    with tempfile.TemporaryDirectory() as tmpdir:
        # Save uploaded file
        suffix = Path(uploaded_file.name).suffix
        bill_path = os.path.join(tmpdir, f"bill{suffix}")
        with open(bill_path, "wb") as f:
            f.write(uploaded_file.getvalue())

        output_path = os.path.join(tmpdir, "solar_report.xlsx")

        try:
            # Step 1: Extract
            with st.spinner("🤖 AI is reading your electricity bill..."):
                data = extract_bill_data(bill_path, api_key)

            # Step 2: Fill Excel
            with st.spinner("📊 Building your Solar Load Calculator Excel..."):
                fill_excel_template(data, output_path, str(TEMPLATE_PATH))

            st.markdown('<div class="success-banner">✅ Report generated successfully!</div>',
                        unsafe_allow_html=True)

            # ── Show extracted data summary ──────────────────────────────────
            st.markdown("### 📋 Extracted Bill Data")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Customer Info**")
                st.write({
                    "Name":       data.get("consumer_name",    "—"),
                    "Cons. No.":  data.get("consumer_number",  "—"),
                    "Tariff":     data.get("tariff_category",  "—"),
                    "Bill Month": data.get("bill_month",       "—"),
                })
            with col2:
                st.markdown("**Billing Details**")
                units = data.get("units_consumed")
                net   = data.get("net_payable")
                rate  = data.get("rate_per_unit")
                load  = data.get("connected_load_kw")
                st.write({
                    "Units (kWh)":   f"{units:,.0f}" if units else "—",
                    "Net Payable":   f"₹{net:,.2f}" if net else "—",
                    "Rate/Unit":     f"₹{rate:.2f}/kWh" if rate else "—",
                    "Connected Load":f"{load:.1f} kW" if load else "—",
                })

            # ── Quick metrics ────────────────────────────────────────────────
            if units and net:
                avg_monthly = units
                daily_kwh   = avg_monthly / 30
                # Simple sizing estimate
                psh, eff = 4.5, 0.80
                system_kw  = round(daily_kwh / (psh * eff), 1)
                system_cost = system_kw * 45000 * 0.70  # after 30% subsidy
                savings_yr  = avg_monthly * 12 * (rate or 9)
                payback     = system_cost / savings_yr if savings_yr else 0

                st.markdown("### ⚡ Quick Estimates")
                m1, m2, m3, m4 = st.columns(4)
                m1.markdown(f'<div class="metric-card"><div class="val">{system_kw} kWp</div><div class="lbl">System Size</div></div>', unsafe_allow_html=True)
                m2.markdown(f'<div class="metric-card"><div class="val">₹{system_cost/100000:.1f}L</div><div class="lbl">Net Cost</div></div>', unsafe_allow_html=True)
                m3.markdown(f'<div class="metric-card"><div class="val">₹{savings_yr/1000:.0f}K</div><div class="lbl">Savings/Year</div></div>', unsafe_allow_html=True)
                m4.markdown(f'<div class="metric-card"><div class="val">{payback:.1f} yrs</div><div class="lbl">Payback</div></div>', unsafe_allow_html=True)

            # ── Download button ──────────────────────────────────────────────
            st.markdown("### 📥 Download Excel Report")
            with open(output_path, "rb") as f:
                excel_bytes = f.read()

            fname = Path(uploaded_file.name).stem
            dl_name = f"Energybae_Solar_Report_{fname}.xlsx"
            st.download_button(
                label="⬇️ Download Solar Load Calculator (Excel)",
                data=excel_bytes,
                file_name=dl_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

            st.markdown('<div class="warning-banner">⚠️ This is an AI-generated estimate. '
                        'Final system sizing is subject to on-site survey by an Energybae engineer.</div>',
                        unsafe_allow_html=True)

            # Show raw JSON
            with st.expander("🔍 View raw extracted JSON data"):
                st.json(data)

        except json.JSONDecodeError as e:
            st.error(f"❌ AI returned invalid JSON: {e}\nPlease try again or check bill quality.")
        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.info("Tip: Ensure your API key is valid and the bill file is readable.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.markdown("""
<div style='text-align:center;color:#888;font-size:0.82rem;padding:0.5rem 0'>
  <b>Energybae</b> · Pimpri, Pune · www.energybae.in · +91 9112233120
</div>
""", unsafe_allow_html=True)
