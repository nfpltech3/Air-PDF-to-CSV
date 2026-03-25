# Project Name

Air PDF to CSV Converter (Nagarkot GUI)

## Tech Stack
- Python 3.11+
- Tkinter, pandas, pdfplumber, openpyxl, Pillow

---

## Installation

### Clone
git clone <repository_url>
cd "Air PDF to CSV"

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
python -m venv venv

2. Activate (REQUIRED)

Windows:
venv\Scripts\activate

Mac/Linux:
source venv/bin/activate

3. Install dependencies
pip install -r requirements.txt

4. Run application
python AIR_PDF_to_CSV_7-Jan-2026.py

---

### Build Executable (For Desktop Apps)

1. Install PyInstaller (Inside venv):
pip install pyinstaller

2. Build using the included Spec file (Ensure you do not run main.py directly):
pyinstaller AIR_PDF_to_CSV_7-Jan-2026.spec

3. Locate Executable:
The application will be generated in the `dist/` folder.

---

## Environment Variables

Copy:
cp .env.example .env

Add required values.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit venv or node_modules.
- Run and test before pushing.
