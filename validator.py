"""
validator.py
────────────
Deterministic validation check for generated canteen register files.
Verifies structure, row counts, cell merges, and data fidelity against the source.
"""

from __future__ import annotations

from pathlib import Path
import openpyxl
from excel_processor import Student


def validate_canteen_register(
    output_path: str | Path,
    expected_students: list[Student],
    input_path: str | Path,
) -> tuple[bool, list[str]]:
    """
    Validates that the output file meets all specification requirements:
      1. Output path != input path.
      2. File opens correctly.
      3. Exactly 4 columns with expected headers.
      4. Total data rows = len(expected_students) * 2.
      5. Each student's Roll, Name, and Room are vertically merged across rows (2*i, 2*i+1).
      6. Money / Record cells (column 4) are strictly unmerged.
      7. Student order and data values match expected records exactly.

    Returns (passed, list_of_issues).
    """
    issues: list[str] = []

    # Check 1: Output is distinct from input
    if Path(output_path).resolve() == Path(input_path).resolve():
        return False, ["Output file cannot be the same as the input file."]

    try:
        wb = openpyxl.load_workbook(str(output_path))
        ws = wb.active
    except Exception as e:
        return False, [f"Failed to open generated file: {e}"]

    # Check 2: Header row
    expected_headers = ["Roll Number", "Name", "Room Number", "Money / Record"]
    actual_headers = [ws.cell(row=1, column=c).value for c in range(1, 5)]
    if actual_headers != expected_headers:
        issues.append(f"Header mismatch. Expected {expected_headers}, got {actual_headers}")

    # Check 3: Total rows
    expected_total_rows = 1 + len(expected_students) * 2
    if ws.max_row != expected_total_rows:
        issues.append(
            f"Row count mismatch: expected {expected_total_rows} rows (1 header + "
            f"{len(expected_students)}*2 data rows), got {ws.max_row}."
        )

    # Gather merged ranges
    merged_ranges = {str(r) for r in ws.merged_cells.ranges}

    # Check 4: Inspect each student block
    for i, student in enumerate(expected_students):
        r1 = 2 + i * 2
        r2 = r1 + 1

        # Check Roll merge (Col A)
        roll_range = f"A{r1}:A{r2}"
        if roll_range not in merged_ranges:
            issues.append(f"Student #{i+1} ({student.name}): Roll cell {roll_range} is not merged.")
        val_roll = str(ws.cell(row=r1, column=1).value or "").strip()
        if val_roll != student.roll.strip():
            issues.append(f"Student #{i+1}: Roll mismatch. Expected '{student.roll}', got '{val_roll}'.")

        # Check Name merge (Col B)
        name_range = f"B{r1}:B{r2}"
        if name_range not in merged_ranges:
            issues.append(f"Student #{i+1} ({student.name}): Name cell {name_range} is not merged.")
        val_name = str(ws.cell(row=r1, column=2).value or "").strip()
        if val_name != student.name.strip():
            issues.append(f"Student #{i+1}: Name mismatch. Expected '{student.name}', got '{val_name}'.")

        # Check Room merge (Col C)
        room_range = f"C{r1}:C{r2}"
        if room_range not in merged_ranges:
            issues.append(f"Student #{i+1} ({student.name}): Room cell {room_range} is not merged.")
        val_room = str(ws.cell(row=r1, column=3).value or "").strip()
        if val_room != student.room.strip():
            issues.append(f"Student #{i+1}: Room mismatch. Expected '{student.room}', got '{val_room}'.")

        # Check Money/Record cells are NOT merged (Col D)
        cell_d1 = f"D{r1}"
        cell_d2 = f"D{r2}"
        for mr in merged_ranges:
            if cell_d1 in mr or cell_d2 in mr:
                issues.append(f"Student #{i+1}: Money/Record cell is merged in {mr}. Must remain unmerged!")

    wb.close()
    return (len(issues) == 0), issues
