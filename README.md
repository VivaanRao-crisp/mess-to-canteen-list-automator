# Monthly Hostel Excel-to-Canteen-Register Generator

A deterministic, lightweight tool that converts a hostel's monthly student Excel list into a print-ready canteen register.

---

## 📥 Download for Windows (No Python Needed)

Hostel staff can download the standalone Windows application directly from the **[GitHub Releases Page](https://github.com/VivaanRao-crisp/mess-to-canteen-list-automator/releases)**:

1. Go to **[Releases](https://github.com/VivaanRao-crisp/mess-to-canteen-list-automator/releases)**.
2. Under **Assets**, click to download **`HostelRegisterGenerator.exe`**.
3. Double-click the downloaded `.exe` to run the application immediately.

> **Note on Windows Security**: Since the `.exe` is built freshly from source on GitHub Actions without an expensive commercial code-signing certificate, Windows SmartScreen may show an unrecognized app prompt on first launch (*"Windows protected your PC"*). Click **"More info"** and then **"Run anyway"**.

---

## 📋 What the Program Does

When the hostel sends a new monthly student list:
1. You select the original Excel file (`.xlsx`).
2. The program automatically identifies the **Roll Number**, **Student Name**, and **Room Number** columns, regardless of column order, case, or common header synonyms (e.g. *Student ID*, *Full Name*, *Room/Block*).
3. It creates a brand-new register Excel workbook where **every student occupies exactly TWO rows**:
   - **Roll Number**: Merged vertically across the 2 rows.
   - **Name**: Merged vertically across the 2 rows.
   - **Room Number**: Merged vertically across the 2 rows.
   - **Money / Record**: TWO separate, unmerged cells (one on each row) for handwriting entries and amounts.
4. It sets up print formatting (fit-to-page width, repeated header rows on every printed page, clear borders, comfortable handwriting row height).
5. The original monthly file is **never modified**.

---

## 🚀 Running from Source (macOS / Linux / Python)

### 1. Requirements

Ensure you have Python 3.8+ installed. Install the dependency (`openpyxl`):

```bash
pip install -r requirements.txt
```

### 2. Run the Application

Launch the desktop interface:

```bash
python3 main.py
```

### 3. Step-by-Step Workflow

1. Click **"📂 Choose Excel File (.xlsx)..."** and select your hostel file (e.g. `MESS LIST 2026.xlsx`).
2. The app detects the columns and shows a preview of the first few students.
3. Click **"✨ Generate Canteen Register"**.
4. Choose where to save the new file.
5. The register is created and validated instantly. You will be prompted to open it immediately!

---

## 🧪 Running Automated Tests

Run the full acceptance test suite:

```bash
# Option 1: Standalone acceptance script
python3 run_acceptance_test.py

# Option 2: Pytest
pytest -v
```

### What the acceptance tests verify:
- ✅ Student records and original order are preserved 1:1.
- ✅ Each student occupies exactly 2 rows.
- ✅ Roll, Name, and Room cells are merged vertically across pairs of rows.
- ✅ Money/Record cells (Column D) are strictly unmerged.
- ✅ Leading zeros in roll numbers (e.g. `002401`) and alphanumeric room numbers (e.g. `Block-A/102`) are preserved.
- ✅ The original hostel workbook remains completely unchanged.

---

## 🤖 Automated CI/CD (GitHub Actions)

A GitHub Actions workflow is configured in `.github/workflows/release.yml`. It runs on a `windows-latest` virtual machine to:
1. Set up Python 3.11.
2. Run automated tests and validation.
3. Compile `HostelRegisterGenerator.exe` using PyInstaller.
4. Automatically attach `HostelRegisterGenerator.exe` to the GitHub Release.

### How to Trigger a New Release:

**Via Git Tag**:
```bash
git tag v0.1.0-beta.1
git push origin v0.1.0-beta.1
```

**Via GitHub Actions Web Interface**:
1. Open the repository on GitHub.
2. Click the **Actions** tab.
3. Select **"Build & Release Windows Executable"**.
4. Click **Run workflow**, enter the version (e.g. `v0.1.0-beta.1`), and click **Run**.
