# VAG Group (VW/Audi/Skoda) – Invoice Extractor

A unified extraction tool for parsing invoices and tabular data from VAG Group's PDF invoices. This tool supports extraction from two primary formats into configurable CSV datasets:

1. **Skoda Auto AS** (Mladá Boleslav / ČSN text-based format)
2. **VW / Audi Group** (VW AG, Audi AG, Audi Hungaria table-based format)

---

## Installation & Setup

⚠️ **IMPORTANT:** You must use a virtual environment.

### 1. Create Virtual Environment

```cmd
python -m venv venv
```

### 2. Activate Virtual Environment (REQUIRED)

**Windows:**
```cmd
venv\Scripts\activate
```

**Mac/Linux:**
```bash
source venv/bin/activate
```

### 3. Install Dependencies

```cmd
pip install -r requirements.txt
```

### 4. Run Application

```cmd
python VAG_Extractor_Combined_App.py
```

---

## Build Executable (For Desktop Application)

1. **Install PyInstaller** (Inside venv):
   ```cmd
   pip install pyinstaller
   ```

2. **Build** using the included Spec file (which includes the Nagarkot Logo):
   ```cmd
   pyinstaller VAG_Extractor_Combined.spec
   ```

3. **Locate Executable**:
   The standalone application will be generated in `dist/VAG_Extractor_Combined.exe`.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- `venv` and `__pycache__` are purposefully ignored in Git.
- Ensure the `Nagarkot Logo.png` is in the directory when running from source or building the executable.
