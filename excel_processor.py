"""
excel_processor.py
──────────────────
Reads a hostel student list Excel file and creates a print-ready canteen register.

Workflow:
  1. Reads student rows in their exact original order.
  2. Creates a new workbook.
  3. Formats each student across exactly TWO rows:
     - Roll Number: merged vertically across the 2 rows
     - Name:        merged vertically across the 2 rows
     - Room Number: merged vertically across the 2 rows
     - Money / Record: TWO separate, unmerged cells (one on each row)
  4. Applies clean borders, clear headers, and print settings.
  5. Saves the new file without touching the original.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────────
# Recognized Header Synonyms
# ─────────────────────────────────────────────────────────────

ROLL_SYNONYMS = {
    "roll no", "roll number", "roll", "rollno", "rollnumber",
    "student id", "studentid", "enrollment number", "enrollment no",
    "enrolment number", "enrolment no", "reg no", "registration no",
    "registration number", "regno", "id no", "idno",
}

NAME_SYNONYMS = {
    "name", "student name", "studentname", "full name", "fullname",
    "student's name", "students name",
}

ROOM_SYNONYMS = {
    "room", "room no", "room number", "roomno", "roomnumber",
    "hostel room", "hostelroom", "room/block", "block/room",
    "room block", "block", "room no.", "hostel room no",
}

# Words that indicate a column is NOT a roll or room number
EXCLUDE_KEYWORDS = re.compile(
    r"(phone|mobile|fee|amount|date|email|address|branch|course|department|"
    r"semester|section|year|batch|gender|category|father|mother|contact|"
    r"aadhar|serial|sno|sr\.?\s*no|sl\.?\s*no)",
    re.IGNORECASE,
)


@dataclass
class Student:
    roll: str
    name: str
    room: str


def _clean_str(value: Any) -> str:
    """Return cell value as a clean string without trailing .0 on integers."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _normalize_header(value: Any) -> str:
    """Normalize text for synonym matching (lowercase, stripped, no punctuation)."""
    if value is None:
        return ""
    s = unicodedata.normalize("NFKC", str(value)).lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s.strip(".:-#/")


def detect_columns(header_row: list[Any]) -> tuple[Optional[int], Optional[int], Optional[int]]:
    """
    Given a list of header cell values, returns (roll_col, name_col, room_col)
    as 0-based column indices.
    """
    roll_col: Optional[int] = None
    name_col: Optional[int] = None
    room_col: Optional[int] = None

    for col_idx, raw_val in enumerate(header_row):
        norm = _normalize_header(raw_val)
        if not norm or EXCLUDE_KEYWORDS.search(norm):
            continue

        if roll_col is None and (norm in ROLL_SYNONYMS or "roll" in norm):
            roll_col = col_idx
        elif name_col is None and (norm in NAME_SYNONYMS or "name" in norm):
            name_col = col_idx
        elif room_col is None and (norm in ROOM_SYNONYMS or "room" in norm or "block" in norm):
            room_col = col_idx

    return roll_col, name_col, room_col


def read_students_from_sheet(
    filepath: str | Path,
    sheet_name: Optional[str] = None,
) -> tuple[list[Student], str, dict[str, Any]]:
    """
    Reads student records from the specified Excel sheet preserving order.
    Returns:
      (students, active_sheet_name, column_info)
    """
    wb = openpyxl.load_workbook(str(filepath), data_only=True, read_only=True)
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        selected_sheet = sheet_name
    else:
        ws = wb.active
        selected_sheet = ws.title

    rows: list[list[Any]] = []
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row))
    wb.close()

    if not rows:
        raise ValueError(f"Sheet '{selected_sheet}' is empty.")

    # Find the header row in the first 15 rows
    header_idx = None
    roll_col = name_col = room_col = None

    for idx, row in enumerate(rows[:15]):
        r_col, n_col, rm_col = detect_columns(row)
        if r_col is not None and n_col is not None:
            header_idx = idx
            roll_col, name_col, room_col = r_col, n_col, rm_col
            break

    if header_idx is None or roll_col is None or name_col is None:
        available_headers = [str(c) for c in (rows[0] if rows else []) if c is not None]
        raise ValueError(
            f"Could not automatically detect required columns ('Roll Number' and 'Name').\n"
            f"First row contains: {available_headers}"
        )

    header_names = rows[header_idx]
    col_info = {
        "header_row": header_idx + 1,
        "roll_header": str(header_names[roll_col]) if roll_col < len(header_names) else "",
        "name_header": str(header_names[name_col]) if name_col < len(header_names) else "",
        "room_header": str(header_names[room_col]) if (room_col is not None and room_col < len(header_names)) else "Not found",
        "has_room": room_col is not None,
    }

    # Extract students preserving order
    students: list[Student] = []
    for row in rows[header_idx + 1:]:
        # Skip completely empty rows
        if not any(cell is not None and str(cell).strip() != "" for cell in row):
            continue

        roll_val = _clean_str(row[roll_col]) if roll_col < len(row) else ""
        name_val = _clean_str(row[name_col]) if name_col < len(row) else ""
        room_val = _clean_str(row[room_col]) if (room_col is not None and room_col < len(row)) else ""

        # Skip repeated header row
        if roll_val.lower() in ROLL_SYNONYMS or name_val.lower() in NAME_SYNONYMS:
            continue

        # Skip rows where both roll and name are missing
        if not roll_val and not name_val:
            continue

        students.append(Student(roll=roll_val, name=name_val, room=room_val))

    return students, selected_sheet, col_info


def generate_canteen_register(
    students: list[Student],
    output_path: str | Path,
) -> None:
    """
    Generates a print-ready canteen register from the list of students.

    Columns:
      1. Roll Number     (A) - merged across 2 rows per student
      2. Name            (B) - merged across 2 rows per student
      3. Room Number     (C) - merged across 2 rows per student
      4. Money / Record  (D) - TWO separate unmerged cells per student
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Canteen Register"

    # ── Styling Definitions ───────────────────────────────────
    FONT_FAMILY = "Calibri"
    header_font = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    data_font = Font(name=FONT_FAMILY, size=10, bold=False, color="000000")

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    alt_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="808080")
    thin_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # ── Header Row (Row 1) ────────────────────────────────────
    headers = [
        ("Roll Number", 16),
        ("Name", 32),
        ("Room Number", 15),
        ("Money / Record", 30),
    ]

    ws.row_dimensions[1].height = 28
    for col_idx, (title, width) in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Student Rows (2 rows per student) ─────────────────────
    ROW_HEIGHT = 23

    for idx, student in enumerate(students):
        r1 = 2 + idx * 2
        r2 = r1 + 1

        ws.row_dimensions[r1].height = ROW_HEIGHT
        ws.row_dimensions[r2].height = ROW_HEIGHT

        # Subtle row alternation per student block
        block_fill = alt_fill if (idx % 2 == 1) else white_fill

        # 1. Roll Number (Merged A{r1}:A{r2})
        cell_roll_1 = ws.cell(row=r1, column=1, value=student.roll)
        cell_roll_1.font = data_font
        cell_roll_1.fill = block_fill
        cell_roll_1.alignment = align_center
        cell_roll_1.border = thin_border

        cell_roll_2 = ws.cell(row=r2, column=1)
        cell_roll_2.fill = block_fill
        cell_roll_2.border = thin_border

        ws.merge_cells(start_row=r1, start_column=1, end_row=r2, end_column=1)

        # 2. Name (Merged B{r1}:B{r2})
        cell_name_1 = ws.cell(row=r1, column=2, value=student.name)
        cell_name_1.font = data_font
        cell_name_1.fill = block_fill
        cell_name_1.alignment = align_left
        cell_name_1.border = thin_border

        cell_name_2 = ws.cell(row=r2, column=2)
        cell_name_2.fill = block_fill
        cell_name_2.border = thin_border

        ws.merge_cells(start_row=r1, start_column=2, end_row=r2, end_column=2)

        # 3. Room Number (Merged C{r1}:C{r2})
        cell_room_1 = ws.cell(row=r1, column=3, value=student.room)
        cell_room_1.font = data_font
        cell_room_1.fill = block_fill
        cell_room_1.alignment = align_center
        cell_room_1.border = thin_border

        cell_room_2 = ws.cell(row=r2, column=3)
        cell_room_2.fill = block_fill
        cell_room_2.border = thin_border

        ws.merge_cells(start_row=r1, start_column=3, end_row=r2, end_column=3)

        # 4. Money / Record (Col D: TWO UNMERGED CELLS)
        cell_rec_1 = ws.cell(row=r1, column=4, value="")
        cell_rec_1.fill = white_fill
        cell_rec_1.border = thin_border
        cell_rec_1.alignment = align_center

        cell_rec_2 = ws.cell(row=r2, column=4, value="")
        cell_rec_2.fill = white_fill
        cell_rec_2.border = thin_border
        cell_rec_2.alignment = align_center

    # ── Page & Print Setup ────────────────────────────────────
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = "A2"  # Freeze header row on scroll

    # Print settings: Repeat header row on every page
    ws.print_title_rows = "1:1"

    # Fit all columns to 1 page wide
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.fitToPage = True

    # Margins (inches)
    from openpyxl.worksheet.page import PageMargins
    ws.page_margins = PageMargins(
        left=0.4, right=0.4, top=0.6, bottom=0.6,
        header=0.3, footer=0.3,
    )
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.center.size = 9

    # Save to output path without modifying source
    wb.save(str(output_path))


def process_hostel_file(
    input_path: str | Path,
    output_path: str | Path,
    sheet_name: Optional[str] = None,
) -> tuple[int, dict[str, Any]]:
    """
    Main processing entry point:
      Reads input -> generates output register -> returns (student_count, info).
    """
    if Path(input_path).resolve() == Path(output_path).resolve():
        raise ValueError("Output file path must be different from input file path.")

    students, active_sheet, col_info = read_students_from_sheet(input_path, sheet_name)
    if not students:
        raise ValueError("No valid student records found in the input file.")

    generate_canteen_register(students, output_path)
    return len(students), col_info
