======================================================================
               HOSTEL REGISTER GENERATOR - USER GUIDE
======================================================================

Welcome to the Hostel-to-Canteen Register Generator!
This application converts your monthly hostel student list Excel file
into a print-ready canteen register formatted with 2 rows per student.


1. HOW TO LAUNCH THE APPLICATION
----------------------------------------------------------------------
On Windows:
  - Simply double-click "HostelRegisterGenerator.exe" (inside the "dist" folder).
  - No installation, Python setup, or internet connection is required.


2. HOW TO GENERATE A CANTEEN REGISTER
----------------------------------------------------------------------
Step 1:
  Click the button:
    [ Choose Excel File (.xlsx)... ]
  and select this month's hostel Excel workbook (e.g., "MESS LIST 2026.xlsx").

Step 2:
  The program automatically detects:
    - Roll Number column
    - Student Name column
    - Room Number column
  You will see a live preview table showing the students found.

Step 3:
  Click the green button:
    [ Generate Canteen Register ]
  Choose where you want to save the new Excel file (for example,
  "Canteen_Register_October_2026.xlsx") and click "Save".

Step 4:
  A success message will confirm that the file was generated.
  Click "Yes" when prompted to open the newly generated Excel file!


3. WHAT THE GENERATED FILE LOOKS LIKE
----------------------------------------------------------------------
Each student occupies exactly TWO physical rows:
  • Roll Number  : Merged vertically across the student's 2 rows.
  • Student Name : Merged vertically across the student's 2 rows.
  • Room Number  : Merged vertically across the student's 2 rows.
  • Money / Record: TWO separate, unmerged cells (one on each row)
                   for handwriting amounts, dates, or signatures.

The file is print-ready for A4 paper:
  • Automatically fits to page width.
  • Repeats the header row on every printed page.
  • Features clean cell borders and comfortable handwriting row height.


4. SAFETY & INTEGRITY GUARANTEE
----------------------------------------------------------------------
  • Your original Excel file is NEVER modified or overwritten.
  • The program reads the data directly and saves a brand-new file.
  • Student order is preserved exactly as listed in the original file.


5. HOW TO REBUILD THE EXECUTABLE (FOR DEVELOPERS)
----------------------------------------------------------------------
If you ever update the Python source code in the future:
  1. On any Windows PC with Python 3.8+ installed, open this folder.
  2. Double-click "build_exe.bat".
  3. The script will automatically create a virtual environment, install
     all dependencies, and build a new "HostelRegisterGenerator.exe" in
     the "dist" folder.
======================================================================
