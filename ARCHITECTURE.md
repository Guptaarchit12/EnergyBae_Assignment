```mermaid
flowchart TD
    A([👤 Sales Team / Customer]) -->|Uploads Bill PDF / Image| B

    subgraph INPUT["📥 INPUT LAYER"]
        B[fa:fa-file-upload Bill Upload\nPDF · PNG · JPG · WEBP]
    end

    subgraph EXTRACTION["🤖 AI EXTRACTION LAYER — Claude claude-sonnet-4-20250514"]
        C[Base64 Encode\nFile → API Payload]
        D[Claude Vision / Document\nOCR + Structured Extraction]
        E{JSON Schema\nValidation}
        F[Extracted Fields:\n• Consumer Name & Number\n• Units Consumed kWh\n• Tariff Slab & Rate\n• Bill Amount & Charges\n• 12-Month History\n• Connected Load\n• Net Payable]
    end

    subgraph EXCEL["📊 EXCEL AUTOMATION LAYER — openpyxl"]
        G[Load Excel Template\nsolar_load_calculator_template.xlsx]
        H[Fill Sheet 1: Bill Input\nCustomer + Billing Fields]
        I[Fill Sheet 2: Solar Sizing\nFormulas Auto-Calculate:\n• System Size kWp\n• Panel Count\n• Generation kWh/yr\n• CO₂ Avoided]
        J[Fill Sheet 3: Customer Report\nFinancial Summary:\n• System Cost ₹\n• Subsidy Amount\n• Payback Period\n• 25-yr Savings\n• ROI %]
    end

    subgraph OUTPUT["📤 OUTPUT LAYER"]
        K[fa:fa-file-excel Filled Excel Report\nEnergybae_Solar_Report.xlsx]
        L[Extracted JSON\nAudit Trail]
    end

    subgraph UI["🌐 WEB INTERFACE — Streamlit"]
        M[Upload Widget]
        N[Live Metrics Dashboard\nSystem Size · Savings · Payback]
        O[Download Button]
    end

    B --> C
    C --> D
    D --> E
    E -->|Valid| F
    E -->|Invalid JSON| D
    F --> H

    G --> H
    H --> I
    I --> J

    J --> K
    F --> L

    A --> M
    M --> B
    K --> O
    K --> N
    O --> A

    style INPUT    fill:#E8F5E9,stroke:#4CAF50,stroke-width:2px
    style EXTRACTION fill:#E3F2FD,stroke:#1976D2,stroke-width:2px
    style EXCEL    fill:#FFF9C4,stroke:#F9A825,stroke-width:2px
    style OUTPUT   fill:#F3E5F5,stroke:#7B1FA2,stroke-width:2px
    style UI       fill:#FBE9E7,stroke:#E64A19,stroke-width:2px

    style A fill:#1B5E20,color:#fff,stroke:none
    style K fill:#1B5E20,color:#fff,stroke:none
```

---

## System Architecture Description

### Overview
The Energybae Solar Load Calculator automates the manual process of reading an electricity bill and computing the recommended solar system size. The pipeline has 3 distinct layers:

### Layer 1 — AI Extraction (Claude claude-sonnet-4-20250514)
- The bill file (PDF or image) is base64-encoded and sent to Claude via the Anthropic Messages API
- Claude's vision/document capabilities parse the bill content and extract ~20 structured fields
- Output is validated JSON matching a predefined schema

### Layer 2 — Excel Automation (openpyxl)
- The Solar Load Calculator Excel template has 3 sheets:
  - **Bill Input** — all extracted data is written here (input cells only, formulas untouched)
  - **Solar Sizing** — formulas auto-calculate system size, generation, costs, ROI
  - **Customer Report** — a clean summary sheet that pulls values from the other two sheets
- Only input cells are populated; all formulas are preserved intact

### Layer 3 — Interface
- **CLI**: `python src/bill_extractor.py --bill bill.pdf --output report.xlsx`
- **Web App**: `streamlit run app.py` — drag-and-drop UI with live metrics and download button

### Data Flow (simplified)
```
Bill (PDF/Image)
    → Claude AI → JSON (20 fields)
        → openpyxl → Excel Template (3 sheets)
            → Download → Sales Team
```

### Key Design Decisions
| Decision | Rationale |
|---|---|
| Claude vision instead of regex OCR | Handles varied bill formats without custom parsers |
| Formulas preserved in Excel | Template stays dynamic; values update when inputs change |
| JSON audit trail | Every extraction is logged for QA and reprocessing |
| Streamlit UI | Zero-install web app; no front-end expertise needed by team |
| IFERROR wrappers on all formulas | Prevents #DIV/0! and #REF! errors on empty cells |
