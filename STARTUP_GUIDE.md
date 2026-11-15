# 🚀 Linwood AI Underwriting Engine — Quickstart & Operations Guide

Welcome to the **Linwood AI Underwriting Engine**. This institutional software autonomously underwrites commercial real estate properties by ingesting raw due diligence files (PDF Offering Memorandums, Excel Rent Rolls, and Trailing-12 Operating Statements) and generating a fully calculated, formula-intact **Waterman Acquisition Pro Forma Model in Excel** in under 3 seconds.

---

## 1. System Requirements

* **Operating System:** Windows 10 / Windows 11 / Windows Server
* **Python Version:** Python 3.10, 3.11, or 3.12 (64-bit recommended)
* **Spreadsheet Software:** Microsoft Excel (desktop installation for native COM formula recalculation)

---

## 2. Installation (1-Minute Setup)

1. Open **Command Prompt** or **PowerShell** in the project folder (`d:\Desktop\Linwood`).
2. Install all required dependencies by running:
   ```bash
   pip install -r requirements.txt
   ```

---

## 3. How to Launch the Software

### Method A: One-Click Launcher (Recommended)
Simply double-click the included batch file:
* **`start_ui.bat`**

This will automatically launch the FastAPI server and open the web dashboard in your default browser.

### Method B: Command Line (FastAPI Web Dashboard)
In your terminal, run:
```bash
python app.py
```
Then open your web browser to:
👉 **`http://localhost:8000`**

### Method C: Streamlit Alternative Portal
If you prefer a Streamlit analytics dashboard, run:
```bash
streamlit run streamlit_app.py
```

---

## 4. How to Use the Web Dashboard

### A. Underwriting a Deal via File Upload
1. Scroll down to the **"Upload Custom Data Room"** section.
2. Enter the **Property Name** (e.g., *Oakwood Ridge Apartments*) and target **Purchase Price** (e.g., *$25,000,000*).
3. Drag and drop your deal files:
   * **Rent Roll:** `.xlsx` or `.xls` spreadsheet listing units and tenant leases.
   * **T-12 Statement:** `.xlsx` spreadsheet showing historical operating expenses.
   * **Offering Memorandum (OM):** `.pdf` broker marketing package.
4. Click **"Run Automated Underwriting"**.
5. Within 2 to 3 seconds, the page renders:
   * Financial KPI Summary Cards (Purchase Price, Unit Count, Total RSF, Operating Expenses)
   * Unit Mix breakdown table (1BR, 2BR, 3BR counts, average SF, in-place rent, market rent)
   * T-12 Operating Expenses breakdown table
   * Live Formula Health Badge (confirming all 100,000+ formulas remain live)

### B. Downloading Your Deliverables
Once underwriting completes, download your ready files from the **Download Hub**:
* **`Download Pro Forma (.xlsx)`**: The finalized Waterman financial model ready for the Investment Committee.
* **`Download Audit Trail (.md)`**: The mathematical provenance report linking every number to its source document.
* **`Inspect Lineage`**: Click to open an interactive modal displaying the full data lineage without leaving the page.

### C. Deleting Deals & Resetting the Dashboard
* Click **`🗑️ Delete Deal & Reset`** at the bottom of the results section to purge temporary session files and reset the dashboard for the next deal.
* To clear selected upload files, click the **`(✕)`** button beside any file or click **`Clear Files`**.

---

## 5. Command-Line Interface (CLI Mode)

To run the pipeline headless via command line without the web browser:
```bash
python run_pipeline.py --dataroom "path/to/dataroom_folder"
```

---

## 6. Project Directory Structure

```
Linwood/
├── app.py                                   # FastAPI backend server
├── start_ui.bat                             # One-click Windows launcher
├── requirements.txt                         # Python dependencies
├── STARTUP_GUIDE.md                         # This user & startup guide
├── COMPLETE_SOFTWARE_MANUAL_AND_ARCHITECTURE.md # Full technical documentation
├── templates/
│   └── Waterman - ACQ Proforma - TEMPLATE.xlsx # Master 13-sheet Waterman template
├── static/
│   └── index.html                           # Modern Tailwind CSS web dashboard
├── underwriting/
│   ├── models.py                            # Canonical Intermediate Representation (IR)
│   ├── parsers.py                           # Multimodal parsers (PDF, XLSX, XLS)
│   ├── injector.py                          # Semantic Excel injector & COM recalculation
│   └── diff.py                              # Mathematical diff & verification engine
├── output/                                  # Completed Excel pro formas & audit reports
└── uploads/                                 # Temporary upload storage
```

---

## 7. Support & Maintenance

* All generated Excel files are saved into `output/`.
* The template file in `templates/Waterman - ACQ Proforma - TEMPLATE.xlsx` serves as the pristine base layout. It contains all 13 worksheets, 104,720 live formulas, and pre-configured forward SOFR interest rate curves.
