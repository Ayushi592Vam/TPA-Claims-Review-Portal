# 🛡️ TPA Claims Review Portal

A professional, Streamlit web application for Third-Party Administrators (TPAs) and insurance teams to **upload, review, edit, and export** loss run data from Excel or CSV files — with full Excel cell traceability and structured JSON output.

---

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Screenshots](#screenshots)
- [Tech Stack](#tech-stack)
- [Installation](#installation)
- [Running the App](#running-the-app)
- [Usage Guide](#usage-guide)
- [JSON Export Format](#json-export-format)
- [Project Structure](#project-structure)
- [Configuration](#configuration)
- [Known Limitations](#known-limitations)

---

## Overview

The TPA Claims Review Portal streamlines the loss run review workflow for insurance professionals. It reads multi-sheet Excel workbooks (or CSVs), automatically classifies sheet types (Loss Run, Summary, Commercial), extracts claim records with their exact Excel cell coordinates, and lets reviewers audit, edit, and export clean structured data — all without touching the original file.

---

## ✨ Features

### 📂 File Ingestion
- Upload `.xlsx` or `.csv` loss run files via drag-and-drop
- Multi-sheet support with a sheet selector dropdown
- **Summary sheet always sorted to the top** and auto-selected on load
- Sheet type auto-detection: `LOSS_RUN`, `SUMMARY`, `COMMERCIAL_LOSS_RUN`, `UNKNOWN`

### 🧠 Intelligent Extraction
- Automatically locates the header row (no manual column mapping needed)
- Extracts every data row with its **exact Excel row and column coordinates**
- Detects and parses **merged cell regions** — Title rows, header bands, and data spans
- Extracts **totals/grand total rows** separately and aggregates financial columns
- Honours the cell's **actual Excel number format** for dates (e.g. `MM-DD-YYYY`, `DD/MM/YYYY`) — displays exactly as it appears in Excel
- Normalises Unicode punctuation (smart quotes, en/em dashes) to plain ASCII

### 🖥️ Review Interface
- **Three-panel layout**: TPA Records list | Field review | Export panel
- Scrollable left panel showing all claim records with status indicators
- Per-claim field table showing:
  - Original extracted value (read-only)
  - Editable modified value (press **Enter** or click away to save)
  - Orange dot indicator on fields that have been changed
- **👁 Eye icon** — click to open a Cell View popup showing:
  - The extracted value and its exact cell address (e.g. `B4`)
  - A rendered image of the actual Excel sheet with the target cell **highlighted in gold**
- **✔ All / ✘ None** — bulk-toggle field inclusion in export
- Per-field checkbox to include/exclude individual fields from export
- Sheet title banner displaying the report title from merged header rows
- Totals summary cards at the bottom of the review panel

### ✏️ Editing
- Click the **✏ pencil icon** to enter edit mode for any field
- Press **Enter** to save — works even if the value hasn't changed
- All edits tracked with `edited: true/false` flag in the export
- Original value always preserved alongside the modified value

### 📤 Export
- **Standard JSON** export with a sequential structure mirroring the Excel layout:
  1. `titleRows` — merged title/header regions (in top-to-bottom order)
  2. `records` — claim records keyed by Claim ID
  3. `totals` — aggregated financial totals from the sheet's totals row
- Exports only checked fields
- All exported strings sanitised (Unicode normalised)
- Auto-saved to local `feature_store/claims_json/` directory with timestamp
- One-click download button

---

## 🖼️ Screenshots
<img width="1745" height="760" alt="image" src="https://github.com/user-attachments/assets/65f595be-22e0-40b0-bdfc-7fcfaf12f68c" />
<img width="1692" height="548" alt="image" src="https://github.com/user-attachments/assets/d0bf02aa-ceb0-4eaa-b63d-297c9808534f" />


> Upload your loss run Excel file → select a sheet → review each claim → export as JSON.

| Panel | Description |
|---|---|
| Left | Scrollable TPA Records list with claim ID, name, and status |
| Centre | Field-by-field review with original vs. modified values and cell tracing |
| Right | Export format selector and merged region metadata |

---

## 🛠️ Tech Stack

| Library | Purpose |
|---|---|
| `streamlit` | Web UI framework |
| `openpyxl` | Excel file reading, cell formatting, merged cell detection |
| `Pillow` (PIL) | Excel sheet rendering and cell highlight image generation |
| `json` | Structured export serialisation |
| `re` | Excel number format parsing |
| `csv` | CSV file ingestion |
| `datetime` | Date formatting and export timestamps |

---

## ⚙️ Installation

### Prerequisites

- Python **3.9+**
- pip

### Steps

```bash
# 1. Clone the repository
git clone https://github.com/your-org/tpa-claims-portal.git
cd tpa-claims-portal

# 2. Create and activate a virtual environment (recommended)
python -m venv venv
source venv/bin/activate        # macOS/Linux
venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install streamlit openpyxl pillow
```

Or install from a requirements file:

```bash
pip install -r requirements.txt
```

**`requirements.txt`**
```
streamlit>=1.32.0
openpyxl>=3.1.0
Pillow>=10.0.0
```

---

## ▶️ Running the App

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`.

---

## 📖 Usage Guide

### 1. Upload a File
Click **"Upload Loss Run Excel/CSV"** and select your file. Supported formats: `.xlsx`, `.csv`.

### 2. Select a Sheet
Use the sheet dropdown (top right). The **Summary** sheet is always listed first and auto-selected if present.

### 3. Review Records
- The **left panel** lists all extracted claim records. Click **Select** to open a record.
- The **centre panel** shows every field for the selected record.
- The **extracted value** (left column) is the raw value from Excel — read-only.
- The **modified value** (right column) is editable. Click ✏ to edit, then press Enter to save.

### 4. Trace a Value to its Excel Cell
Click the **👁** eye icon on any field to open the **Cell View** popup. It shows:
- The field name and extracted value
- The exact cell address (e.g. `B4`, Row 4 · Col 2)
- A rendered snapshot of the Excel sheet with the cell **highlighted in gold**

### 5. Select Fields for Export
- Use **✔ All** to select all fields, **✘ None** to deselect all
- Use the checkbox on each row to toggle individual fields

### 6. Export
Click **⬇ Export as Standard JSON** in the right panel. A download button will appear. The JSON is also auto-saved to `feature_store/claims_json/`.

---

## 📄 JSON Export Format

The exported JSON mirrors the physical top-to-bottom layout of the Excel sheet:

```json
{
  "exportDate": "2026-03-05T11:23:30.616129",
  "sheetMeta": {
    "sheet_name": "CGL Loss Run",
    "record_count": 10
  },
  "titleRows": [
    {
      "type": "TITLE",
      "value": "COMMERCIAL GENERAL LIABILITY - LOSS RUN REPORT",
      "excel_row": 1,
      "excel_col": 1,
      "span_cols": 11,
      "span_rows": 1
    },
    {
      "type": "TITLE",
      "value": "Insured: Apex Construction LLC | Policy #: CGL-2021-00847",
      "excel_row": 2,
      "excel_col": 1,
      "span_cols": 11,
      "span_rows": 1
    }
  ],
  "records": {
    "CGL-001": {
      "Date of Loss": {
        "value": "03-15-2021",
        "original": "03-15-2021",
        "edited": false,
        "excel_row": 4,
        "excel_col": 2,
        "record_index": 0
      }
    }
  },
  "totals": {
    "excel_row": 14,
    "aggregated": {
      "Total Paid": 186100.0,
      "Reserve": 209500.0,
      "Total Incurred": 395600.0
    }
  },
  "recordCount": 10
}
```

### Key Fields

| Field | Description |
|---|---|
| `titleRows` | Merged header regions in Excel row order (report title, insured info, etc.) |
| `records` | Claim records keyed by Claim ID, each field with its Excel coordinates |
| `value` | The (possibly edited) display value |
| `original` | The original extracted value before any edits |
| `edited` | `true` if the field was modified by the reviewer |
| `excel_row` / `excel_col` | 1-based row and column position in the source Excel file |
| `totals.aggregated` | Sum of all numeric financial columns from the totals row |

---

## 📁 Project Structure

```
tpa-claims-portal/
│
├── app.py                          # Main Streamlit application
├── requirements.txt                # Python dependencies
├── README.md                       # This file
│
└── feature_store/
    └── claims_json/                # Auto-saved JSON exports (timestamped)
        ├── CGL Loss Run_20260305_112330.json
        └── ...
```

---

## 🔧 Configuration

No configuration file is required. The following constants can be changed directly in `app.py`:

| Constant | Location | Default | Description |
|---|---|---|---|
| `FEATURE_STORE_PATH` | Top of file | `feature_store/claims_json` | Directory for auto-saved exports |
| `scale` | `render_excel_sheet()` | `1.0` | Zoom scale for the cell view renderer |
| `pad_x`, `pad_y` | `crop_context()` | `300, 200` | Padding around highlighted cell in the popup |
| `container height` | Left panel | `700` | Height of the scrollable TPA records panel |

---


## 📜 License

This project is proprietary. All rights reserved.

---

*Built for insurance operations teams. Designed to make loss run review faster, traceable, and audit-ready.*
