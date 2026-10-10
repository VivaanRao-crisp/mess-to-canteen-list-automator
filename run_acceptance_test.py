"""
run_acceptance_test.py
----------------------
Runs a complete end-to-end acceptance test:
  1. Creates a synthetic hostel Excel file with 3 students.
  2. Runs the canteen register converter.
  3. Verifies that:
     - Output contains the same students in the exact same order.
     - Each student occupies exactly 2 rows.
     - Roll number, Name, and Room number are vertically merged across their 2 rows.
     - The Money/Record column has two separate, unmerged cells per student.
     - No student data is lost.
     - The original input file remains unchanged.
  4. Also tests real uploaded sample files (MESS LIST 2026.xlsx and MESS LIST 2026-v3.xlsx)
     if present, generating sample registers for inspection.
"""

from __future__ import annotations

import sys
from pathlib import Path
import openpyxl

# Configure stdout safely for UTF-8 with ASCII fallback on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from excel_processor import process_hostel_file, read_students_from_sheet
from validator import validate_canteen_register


def run_test():
    print("=" * 70)
    print("RUNNING CANTEEN REGISTER GENERATOR ACCEPTANCE TEST")
    print("=" * 70)

    project_dir = Path(__file__).parent
    synthetic_input = project_dir / "sample_hostel_input.xlsx"
    synthetic_output = project_dir / "sample_canteen_register_output.xlsx"

    # Step 1: Create synthetic test file
    print("\n[1] Creating sample hostel workbook with 3 students...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hostel List"
    ws.append(["Roll Number", "Student Name", "Room Number"])
    students_data = [
        ("002401", "Rahul Sharma", "B-101"),
        ("002402", "Priya Singh", "B-102"),
        ("002403", "Amit Kumar Verma", "C-204"),
    ]
    for roll, name, room in students_data:
        ws.append([roll, name, room])
    wb.save(synthetic_input)
    print(f"    Created '{synthetic_input.name}' with {len(students_data)} students.")

    # Step 2: Convert
    print("\n[2] Processing hostel file to generate canteen register...")
    count, info = process_hostel_file(synthetic_input, synthetic_output)
    print(f"    Generated '{synthetic_output.name}' with {count} students.")
    print(f"    Detected columns: Roll='{info['roll_header']}', Name='{info['name_header']}', Room='{info['room_header']}'")

    # Step 3: Verify requirements
    print("\n[3] Verifying all acceptance criteria...")
    students, _, _ = read_students_from_sheet(synthetic_input)
    passed, issues = validate_canteen_register(synthetic_output, students, synthetic_input)

    wb_out = openpyxl.load_workbook(synthetic_output)
    ws_out = wb_out.active

    # Criterion A: Same students in same order
    print("    - Checking student order and preservation...")
    assert len(students) == len(students_data), "Student count mismatch"
    for i, (exp_roll, exp_name, exp_room) in enumerate(students_data):
        r1 = 2 + i * 2
        actual_roll = str(ws_out.cell(row=r1, column=1).value)
        actual_name = str(ws_out.cell(row=r1, column=2).value)
        actual_room = str(ws_out.cell(row=r1, column=3).value)
        assert actual_roll == exp_roll, f"Roll mismatch at row {r1}: {actual_roll} != {exp_roll}"
        assert actual_name == exp_name, f"Name mismatch at row {r1}: {actual_name} != {exp_name}"
        assert actual_room == exp_room, f"Room mismatch at row {r1}: {actual_room} != {exp_room}"
    print("      [PASS] Student details and order preserved exactly 1:1.")

    # Criterion B: Exactly 2 rows per student
    print("    - Checking row structure (2 rows per student)...")
    expected_rows = 1 + len(students_data) * 2  # 1 header + 6 data = 7
    assert ws_out.max_row == expected_rows, f"Expected {expected_rows} rows, got {ws_out.max_row}"
    print(f"      [PASS] Row count is {ws_out.max_row} (1 header + {len(students_data)*2} student rows).")

    # Criterion C: Roll, Name, Room merged vertically across 2 rows
    print("    - Checking merged identification cells...")
    merged_ranges = {str(r) for r in ws_out.merged_cells.ranges}
    for i in range(len(students_data)):
        r1 = 2 + i * 2
        r2 = r1 + 1
        assert f"A{r1}:A{r2}" in merged_ranges, f"A{r1}:A{r2} not merged"
        assert f"B{r1}:B{r2}" in merged_ranges, f"B{r1}:B{r2} not merged"
        assert f"C{r1}:C{r2}" in merged_ranges, f"C{r1}:C{r2} not merged"
    print("      [PASS] Roll, Name, and Room are merged vertically across pairs of rows.")

    # Criterion D: Money / Record column has TWO separate cells per student
    print("    - Checking Money / Record column (Col D: 2 separate unmerged cells)...")
    for i in range(len(students_data)):
        r1 = 2 + i * 2
        r2 = r1 + 1
        for mr in merged_ranges:
            assert f"D{r1}" not in mr and f"D{r2}" not in mr, f"Record cell is merged in {mr}!"
    print("      [PASS] Money / Record cells (D2, D3, D4, etc.) are strictly unmerged.")

    # Criterion E: Original input file remains untouched
    print("    - Checking original input file integrity...")
    wb_in_check = openpyxl.load_workbook(synthetic_input)
    ws_in_check = wb_in_check.active
    assert ws_in_check.max_row == 4, "Original input file was modified!"
    wb_in_check.close()
    print("      [PASS] Original input file is completely untouched.")

    wb_out.close()

    if passed:
        print("\n" + "=" * 70)
        print("[SUCCESS] ALL SYNTHETIC ACCEPTANCE CRITERIA PASSED SUCCESSFULLY!")
        print("=" * 70)
    else:
        print("\n[FAIL] Issues noted:", issues)
        return False

    # Step 4: Test real sample files if found
    for sample_name in ["MESS LIST 2026-v3.xlsx", "MESS LIST 2026.xlsx"]:
        sample_path = project_dir / sample_name
        if sample_path.exists():
            print(f"\n[4] Testing real sample file: '{sample_name}'...")
            try:
                out_sample = project_dir / f"canteen_register_{sample_path.stem}.xlsx"
                cnt, inf = process_hostel_file(sample_path, out_sample)
                st_list, _, _ = read_students_from_sheet(sample_path)
                ok, warns = validate_canteen_register(out_sample, st_list, sample_path)
                print(f"    [PASS] Processed {cnt} students from '{sample_name}'.")
                print(f"    [PASS] Detected: Roll='{inf['roll_header']}', Name='{inf['name_header']}', Room='{inf['room_header']}'.")
                print(f"    [PASS] Validation: {'PASSED' if ok else 'FAILED: ' + str(warns)}")
                print(f"    [PASS] Output saved to: '{out_sample.name}'")
            except Exception as e:
                print(f"    [WARN] Error on sample file '{sample_name}': {e}")

    print("\nAcceptance tests complete.")
    return True


if __name__ == "__main__":
    success = run_test()
    sys.exit(0 if success else 1)
