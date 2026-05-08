# ⚡ Energybae Solar Load Calculator
### Electricity Bill → AI Extraction → Excel Solar Report — Automated

> **AI Intern Practical Task — Energybae, Pimpri, Pune**  
> Automates the 15–30 minute manual process of reading an MSEDCL electricity bill
> and computing the correct solar system size, savings, and ROI.

---

## 🎯 What This Does

| Step | What Happens |
|------|-------------|
| **1. Upload** | User uploads an electricity bill (PDF or image) |
| **2. AI Reads** | Claude AI (claude-sonnet-4-20250514) extracts 20+ structured fields from the bill |
| **3. Calculate** | Solar system size, panel count, costs, payback, ROI auto-calculate via Excel formulas |
| **4. Download** | User receives a 3-sheet Excel report, ready to share with the customer |

**Before automation:** 15–30 min/customer, manual, error-prone  
**After automation:** ~15 seconds, zero manual entry, consistent

---

## 🗂️ Project Structure

```
energybae_solar/
│
├── app.py                          # Streamlit web app (bonus UI)
│
├── src/
│   ├── bill_extractor.py           # Core pipeline: extract + fill Excel
│   ├── create_template.py          # Builds the 3-sheet Excel template
│   └── generate_sample_bill.py     # Creates a realistic MSEDCL test bill
│
├── templates/
│   └── solar_load_calculator_template.xlsx   # Master Excel template
│
├── sample_data/
│   ├── sample_msedcl_bill.pdf      # Realistic test bill (generated)
│   └── extracted_sample.json       # Example AI extraction output
│
├── output/                         # Generated reports go here
│
├── requirements.txt
├── README.md                       # This file
└── ARCHITECTURE.md                 # System architecture + Mermaid diagram
```

---

## 🚀 Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Set your Anthropic API key
```bash
# Option A: Environment variable (recommended)
export ANTHROPIC_API_KEY="sk-ant-..."

# Option B: .env file
echo "ANTHROPIC_API_KEY=sk-ant-..." > .env
```

### 3a. Run via Command Line
```bash
python src/bill_extractor.py \
    --bill path/to/your_bill.pdf \
    --output output/solar_report.xlsx \
    --save-json
```

### 3b. Run via Web App (Streamlit)
```bash
streamlit run app.py
# Opens at http://localhost:8501
```

---

## 📊 Excel Template — 3 Sheets

### Sheet 1: `Bill Input`
All fields extracted from the electricity bill:
- **Customer Info** — Name, consumer number, address, tariff category, meter number, sanctioned load
- **Monthly Consumption** — 12-month kWh history + auto-calculated average & peak
- **Bill Details** — Units consumed, bill amount, fixed charges, electricity duty, FAC, meter rent, net payable, tariff slab, rate/unit, connected load, power factor

> 🟡 Yellow cells = input (filled by AI)  
> 🟠 Orange cells = auto-calculated  
> 🔵 Blue cells = key output results

### Sheet 2: `Solar Sizing`
Fully formula-driven. Calculates:
- Recommended solar system size (kWp)
- Panel count (assuming 400W panels)
- Rooftop area required
- Annual solar generation (kWh/year)
- System cost before and after government subsidy (PM Surya Ghar scheme)
- Annual savings (Year 1)
- Simple payback period
- 25-year cumulative savings (NPV with escalation)
- ROI percentage
- CO₂ avoided (lifetime and annual)

**Key assumptions** (editable blue cells):
- Peak Sun Hours: 4.5 hrs/day (Maharashtra)
- System Efficiency: 80%
- Performance Ratio: 75%
- Cost per kWp: ₹45,000 (on-grid)
- Government Subsidy: 30% (PM Surya Ghar)
- Annual tariff escalation: 5%
- Grid emission factor: 0.82 kg CO₂/kWh (CEA 2023)

### Sheet 3: `Customer Report`
A clean, printable summary for the customer — pulls all key values from Sheets 1 & 2.

---

## 🤖 AI Extraction — How It Works

The `extract_bill_data()` function in `bill_extractor.py`:

1. Encodes the bill file (PDF or image) as base64
2. Sends it to `claude-sonnet-4-20250514` via the Anthropic Messages API
   - PDFs use the `document` content block
   - Images use the `image` content block
3. Claude is prompted with a strict JSON schema and domain-specific instructions
4. The response is parsed and validated
5. Extracted fields are mapped to specific Excel cells

**Supported bill types:**
- ✅ MSEDCL (Maharashtra) — primary target
- ✅ Any readable PDF or image bill
- ✅ Bills with 12-month consumption history tables

**Extracted fields (20+):**
```
consumer_name, consumer_number, address, tariff_category, division,
bill_month, bill_date, meter_number, sanctioned_load_kw,
units_consumed, bill_amount, fixed_charges, electricity_duty,
fuel_adjustment_charge, meter_rent, subsidies_rebate, net_payable,
tariff_slab, rate_per_unit, connected_load_kw, power_factor,
monthly_consumption (12-month dict)
```

---

## 🖥️ Web App Features

The Streamlit app (`app.py`) provides:
- Drag-and-drop file upload
- Image preview for uploaded bills
- Live extraction progress indicators
- Quick metrics display (system size, cost, savings, payback)
- One-click Excel download
- Raw JSON viewer (for debugging/QA)

---

## ⚠️ Important Notes

1. **Do not overwrite Excel formulas** — the extractor only fills input cells (yellow). All formula cells are untouched.
2. **API Key** — Keep your `ANTHROPIC_API_KEY` secret. Never commit it to git (use `.env` + `.gitignore`).
3. **Accuracy** — AI extraction accuracy is high for clear, standard MSEDCL bills. Blurry or non-standard bills may need manual review of extracted JSON (`--save-json` flag).
4. **Assumptions** — Solar sizing assumptions in Sheet 2 are editable. The sales team should review these for each state/region.

---

## 🔧 CLI Reference

```
usage: bill_extractor.py [-h] --bill BILL [--output OUTPUT] 
                         [--template TEMPLATE] [--api-key API_KEY] [--save-json]

options:
  --bill      Path to electricity bill (PDF or image) [required]
  --output    Output Excel file path [default: output/solar_report.xlsx]
  --template  Custom Excel template path (optional)
  --api-key   Anthropic API key (or use ANTHROPIC_API_KEY env var)
  --save-json Also save extracted data as JSON (useful for debugging)
```

---

## 📈 What I'd Improve Next

1. **Multi-state support** — Add tariff logic for BESCOM (Karnataka), TPDDL (Delhi), etc.
2. **Confidence scores** — Flag low-confidence extractions for human review
3. **Database storage** — Save customer records to PostgreSQL for CRM integration
4. **Email automation** — Auto-send the filled Excel to the sales team via Zapier/Make
5. **Google Sheets output** — For teams who prefer cloud spreadsheets
6. **Bill validation** — Cross-check extracted values (e.g., verify net = sum of components)
7. **Multi-bill averaging** — Accept 3–6 months of bills for more accurate sizing

---

## 🏢 About Energybae

**Energybae** · Pimpri, Pune · Maharashtra  
Empowering People with Renewable Energy Solutions  
🌐 www.energybae.in | 📧 energybae.co@gmail.com | 📞 +91 9112233120

---

*This is a computer-generated document. Solar sizing estimates are indicative and subject to on-site survey.*
