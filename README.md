# AI-Driven Design Validator (KADS)

## Design Intelligence KADS Features

The platform includes the KADS (Knowledge-Aided Design Suite) features via the **"Design Intelligence (KADS)" tab** in the PyQt6 app — each capability is a **separate module tab with its own button**, plus a **Run All Features** tab. The **Execution Console is a separate window**; every run is framed by delimiter lines so executions read as distinct blocks. All features run **fully offline** (no SQL server, no internet, no API keys) — knowledge is stored in `lessons_db.json`.

| Module Tab | What it does | Engine | Output |
|---|---|---|---|
| Document Intelligent Converter | Processes project-wise inspection reports, extracts & classifies observations (Material / Design-Drawing / Workmanship / Installation / Documentation / Testing & Commissioning), builds the searchable Lessons Learned repository (prompt-driven) | `doc_intelligence.py` | `KADS_Outputs/1_Lessons_Learned_Register.xlsx` |
| Designer Assistant | Proactive lessons learned, recurring issues and design considerations while querying a system | `doc_intelligence.py` | console + table |
| Binding Data Gap Analysis | Vendor submission vs POTS — missing drawings/calcs/certificates/manuals/test procedures, red flags, points to seek before approval | `binding_review.py` | `KADS_Outputs/3_Binding_Data_Gap_Analysis*.xlsx` |
| PreBid Queries | RFP analysis — ambiguities, contradictions, missing info, impractical requirements, execution risks + lessons-driven queries | `prebid_queries.py` | `KADS_Outputs/4_PreBid_Queries*.xlsx` |
| Design Review Checklists | System-wise checklists, risk register from historical data, recurring deficiencies, preventive measures | `design_review.py` | `KADS_Outputs/5_Design_Review_Checklist_*.xlsx` |
| Technical Compliance Matrix | Vendor offer vs SOTR/TTS — compliance matrix, deviations/exclusions/assumptions/ambiguities, auto technical queries | `offer_scrutiny.py` | `KADS_Outputs/2_Technical_Compliance_Matrix*.xlsx` |
| Run All Features | Executes all modules end-to-end in one click | `kads_cli.py` | all of the above |

### Quick start (KADS)

```bash
# GUI (recommended): login, then open the "Design Intelligence (KADS)" tab
python gui.py

# CLI - run everything end-to-end
python kads_cli.py --all

# CLI - prompt-driven conversion (user requirement based on prompt)
python kads_cli.py --prompt "convert inspection reports of DSV project only"
python kads_cli.py --prompt "show lessons learned for fire main system"
```

Inspection reports are read **project-wise** from `sample_data/inspection_reports/<Project>/`. Drop real HSL reports (.txt/.pdf/.docx) into project folders and rebuild — the repository, classifier and Excel registers update automatically.

### Build the Windows EXE

Run `build.bat` on a Windows machine — it produces `dist\DesignValidator.exe` (GUI with all KADS modules and sample data bundled) and `dist\DesignValidatorServer.exe` (optional API backend).

### KADS REST endpoints (server.py)

`POST /api/kads/convert-reports` · `GET /api/kads/lessons/search` · `POST /api/kads/lessons/proactive` · `POST /api/kads/compliance-matrix` · `POST /api/kads/binding-gap` · `POST /api/kads/prebid` · `POST /api/kads/design-review`

---

# Design Validation Platform (Proof of Concept)

An offline-capable, engineering-grade platform designed to automate the ingestion of equipment specs (PDF, Excel, JSON, Text) and validate physical, electrical, and thermal constraints against deployment site cabinet configurations.

The system is split into a **FastAPI backend server** (hosting a semantic rules knowledge base) and a **PyQt6 desktop client dashboard**.

---

## Key Features

1. **Multi-Format Ingestion**: Upload `.json` files, `.xlsx` Excel spreadsheets, `.pdf` manuals, or `.txt` datasheets. The layout extraction engine parses structured and unstructured data to retrieve metrics.
2. **Semantic Rules Retrieval**: Performs semantic retrieval using synonym-expansion Jaccard matching to query applicable standards inside a local standards database (`rules_db.json`).
3. **Standards-Compliant Valider**: Evaluates measurements against industry standards:
   - **IEC 60297-3-100**: Cabinet depth & fit checks.
   - **OSHA 29 CFR 1910.303**: Maintenance working clearance checks.
   - **NFPA 70 National Electrical Code (NEC)**: Electrical power loads.
   - **ASHRAE TC 9.9**: Thermal dissipation / HVAC heat loads.
4. **Role-Based Authentication**:
   - **Admin** (Username: `admin`, Password: `admin123`): Full edit privileges for cabinet setup metrics, with capability to update/save the baseline specs directly to the backend.
   - **User** (Username: `user`, Password: `user123`): Read-only baseline specifications (fetched from the server), but can validate equipment.
5. **High-Fidelity Reports Export**: Generates styled **PDF Reports** (via `reportlab`) and tabular **Excel Worksheets** (via `openpyxl`) detailing individual check results with exact regulatory citations as evidence for compliance packages.
6. **Offline Fallback Resilience**: Bypasses the server login dynamically if the backend is offline, running validation checks locally.

---

## Directory Structure

- `server.py`: FastAPI server hosting authentication, file uploads, validation runs, and report file delivery.
- `gui.py`: PyQt6 client app dashboard with a login view, specs panel, console log, and reports download triggers.
- `nlp_engine.py`: Regex metrics extractor, Excel parser, and synonym-expansion semantic rules retriever.
- `reports.py`: PDF (`reportlab`) and Excel (`openpyxl`) compiler module.
- `rules_db.json`: Curated standards database with citations.
- `db.json`: User registry and baseline cabinet setup data store.
- `requirements.txt`: Python package dependencies.
- `run_server.bat`: Windows batch file to launch the backend server.
- `build.bat`: PyInstaller compilation script.

---

## Installation & Setup

1. **Install Dependencies**:
   Ensure Python 3.10+ is installed. Run the command:
   ```bash
   pip install -r requirements.txt
   ```

2. **Launch the FastAPI Server**:
   Double-click `run_server.bat` or run:
   ```bash
   python server.py
   ```
   The backend API will run on `http://127.0.0.1:8000`.

3. **Launch the Client App**:
   Run the GUI client:
   ```bash
   python gui.py
   ```
   Or execute the standalone compiled application located at:
   ```text
   dist/DesignValidator.exe
   ```

---

## Testing the Platform

1. **Authentication Logins**:
   - Log in as **User** (`user` / `user123`). Notice that "Cabinet Limits Setup" is read-only (greyed out).
   - Log in as **Admin** (`admin` / `admin123`). Change a value (e.g. UPS capacity to `3.0kW`) and click **"Save Baseline to Server"**. Close and log back in to verify the changes persist.
2. **Uploading Documents**:
   - Click **"Upload Document Spec"**. Select a supported `.xlsx` Excel list, `.json` file, `.pdf` manual, or `.txt` spec sheet to extract parameters.
3. **Running Checks & Exporting Reports**:
   - Click **"Run Validation Rule Checks"**. Check the console for citations.
   - Click **"Download PDF Report"** or **"Download Excel Report"** to export validation documents.
