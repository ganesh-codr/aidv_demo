# AI-Driven Design Validator

## Features

The platform includes features **"Docu Intelligence" tab** and **"Docu Assistant" tab** in the PyQt6 app — each capability is a **separate module tab with its own button**. All features run **fully offline** (no SQL server, no internet, no API keys).

| Module Tab | What it does | Engine | Output |
|---|---|---|---|
| Document Intelligent Converter | Processes project-wise inspection reports, extracts & classifies observations (Material / Design-Drawing / Workmanship / Installation / Documentation / Testing & Commissioning), builds the searchable Lessons Learned repository (prompt-driven) | `doc_intelligence.py` | `KADS_Outputs/1_Lessons_Learned_Register.xlsx` |
| Document Assistant | Proactive lessons learned, recurring issues and design considerations while querying a system | `doc_intelligence.py` | console + table |

### Quick start (KADS)

```bash
# GUI (recommended): login, 
python main.py


### Build the Windows EXE

Run `build.bat` on a Windows machine — it produces `dist\DocuChatAI.exe` (GUI with all modules and sample data bundled).

### KADS REST endpoints (server.py)

---

# Design Validation Platform (Proof of Concept)

An offline-capable, engineering-grade platform designed to automate the ingestion of equipment specs (PDF, Excel, JSON, Text) and validate physical, electrical, and thermal constraints against deployment site cabinet configurations.

The system is split into a **FastAPI backend server** (hosting a semantic rules knowledge base) and a **PyQt6 desktop client dashboard**.

---


## Installation & Setup

1. **Install Dependencies**:
   Ensure Python 3.10+ is installed. Run the command:
   ```bash
   pip install -r requirements.txt
   ```

3. **Launch the Client App**:
   Run the GUI client:
   ```bash
   python main.py
   ```
   Or execute the standalone compiled application located at:
   ```text
   dist/DocuChatAI.exe
   ```

---

## Authentication

1. **Authentication Logins**:
   - Log in as **User** (`admin` / `admin`)
   - Log in as **Admin** (`engineer` / `eng`)
