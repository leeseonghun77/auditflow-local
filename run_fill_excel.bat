@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  python -m venv .venv
)
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python fill_ledger_from_pdfs.py
.venv\Scripts\python enhance_for_audit.py
pause