# Monthly Hostel Excel-to-Canteen-Register Generator

A deterministic, lightweight tool that converts a hostel's monthly student Excel list into a print-ready canteen register.

---

## 📋 What It Does

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

## 🚀 Quick Start

### 1. Requirements

Ensure you have Python 3.8+ installed. Install the single dependency (`openpyxl`):

```bash
pip install -r requirements.txt
```

### 2. Run the Application

Launch the simple desktop interface:

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

## 📅 How to Use It Next Month

Next month when you receive a new hostel Excel file:
1. Open the app: `python3 main.py`
2. Select next month's file.
3. Click **"Generate Canteen Register"**.
4. Save it as `Canteen_Register_<Month>.xlsx`.
5. Open the generated file and print directly (all page sizes, margins, repeated headers, and fit-to-width are already configured).

---

## 📦 Optional: Packaging as a Standalone Executable

If you want a standalone app that can be launched with a double-click without opening a terminal:

```bash
pip install pyinstaller

# On macOS:
pyinstaller --noconsole --onefile --name "CanteenRegisterGenerator" main.py

# On Windows:
pyinstaller --noconsole --onefile --name "CanteenRegisterGenerator" main.py
```

The compiled standalone executable will be located inside the `dist/` folder.
