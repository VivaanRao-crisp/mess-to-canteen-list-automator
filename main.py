"""
main.py
───────
Simple, beginner-friendly desktop interface for converting monthly hostel
Excel files into print-ready canteen registers.

Usage:
  python3 main.py
"""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from excel_processor import Student, generate_canteen_register, read_students_from_sheet
from validator import validate_canteen_register


class CanteenRegisterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hostel to Canteen Register Generator")
        self.geometry("740x580")
        self.minsize(680, 500)
        self.configure(bg="#F4F6F9")

        self.input_file: Optional[Path] = None
        self.students: list[Student] = []
        self.sheet_name: str = ""
        self.col_info: dict = {}

        self._init_styles()
        self._build_ui()

    def _init_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure(
            "TButton",
            font=("Segoe UI", 10, "bold"),
            padding=6,
        )
        self.style.configure(
            "Primary.TButton",
            background="#1F4E79",
            foreground="white",
            font=("Segoe UI", 11, "bold"),
            padding=8,
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", "#285F94")],
        )

    def _build_ui(self):
        # Header banner
        banner = tk.Frame(self, bg="#1F4E79", pady=14)
        banner.pack(fill="x")

        tk.Label(
            banner,
            text="Hostel to Canteen Register Generator",
            font=("Segoe UI", 16, "bold"),
            fg="white",
            bg="#1F4E79",
        ).pack()
        tk.Label(
            banner,
            text="Automatically formats student lists into 2-row printable registers",
            font=("Segoe UI", 9),
            fg="#DCE6F1",
            bg="#1F4E79",
        ).pack()

        # Main content container
        content = tk.Frame(self, bg="#F4F6F9", padx=25, pady=15)
        content.pack(fill="both", expand=True)

        # ── Step 1: File Selection ─────────────────────────────
        file_frame = tk.LabelFrame(
            content,
            text=" 1. Select Hostel Excel File ",
            font=("Segoe UI", 10, "bold"),
            fg="#1F4E79",
            bg="white",
            padx=15,
            pady=12,
            relief="groove",
        )
        file_frame.pack(fill="x", pady=(0, 12))

        file_top = tk.Frame(file_frame, bg="white")
        file_top.pack(fill="x")

        ttk.Button(
            file_top,
            text="📂 Choose Excel File (.xlsx)...",
            command=self._on_choose_file,
        ).pack(side="left")

        self.lbl_file_status = tk.Label(
            file_top,
            text="No file selected yet.",
            font=("Segoe UI", 9, "italic"),
            fg="#666666",
            bg="white",
            padx=12,
        )
        self.lbl_file_status.pack(side="left", fill="x", expand=True)

        self.lbl_detection_info = tk.Label(
            file_frame,
            text="",
            font=("Segoe UI", 9),
            fg="#2E5B88",
            bg="white",
            anchor="w",
            justify="left",
        )
        self.lbl_detection_info.pack(fill="x", pady=(6, 0))

        # ── Step 2: Data Preview ──────────────────────────────
        preview_frame = tk.LabelFrame(
            content,
            text=" 2. Preview Student List ",
            font=("Segoe UI", 10, "bold"),
            fg="#1F4E79",
            bg="white",
            padx=15,
            pady=10,
            relief="groove",
        )
        preview_frame.pack(fill="both", expand=True, pady=(0, 15))

        tree_scroll = ttk.Scrollbar(preview_frame)
        tree_scroll.pack(side="right", fill="y")

        self.preview_tree = ttk.Treeview(
            preview_frame,
            columns=("roll", "name", "room"),
            show="headings",
            height=6,
            yscrollcommand=tree_scroll.set,
        )
        tree_scroll.config(command=self.preview_tree.yview)

        self.preview_tree.heading("roll", text="Roll Number")
        self.preview_tree.heading("name", text="Student Name")
        self.preview_tree.heading("room", text="Room Number")

        self.preview_tree.column("roll", width=120, anchor="center")
        self.preview_tree.column("name", width=300, anchor="w")
        self.preview_tree.column("room", width=120, anchor="center")

        self.preview_tree.pack(fill="both", expand=True)

        # ── Step 3: Action Buttons ─────────────────────────────
        bottom_frame = tk.Frame(content, bg="#F4F6F9")
        bottom_frame.pack(fill="x")

        self.btn_generate = ttk.Button(
            bottom_frame,
            text="✨ Generate Canteen Register",
            style="Primary.TButton",
            command=self._on_generate,
            state="disabled",
        )
        self.btn_generate.pack(side="right")

    def _on_choose_file(self):
        path_str = filedialog.askopenfilename(
            title="Select Hostel Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")],
        )
        if not path_str:
            return

        selected_path = Path(path_str)
        try:
            students, sheet, col_info = read_students_from_sheet(selected_path)
            if not students:
                messagebox.showerror(
                    "No Data",
                    "No valid student records could be found in the selected sheet."
                )
                return

            self.input_file = selected_path
            self.students = students
            self.sheet_name = sheet
            self.col_info = col_info

            # Update UI Labels
            self.lbl_file_status.config(
                text=f"Selected: {selected_path.name}  ({len(students)} students found)",
                font=("Segoe UI", 9, "bold"),
                fg="#0F5132",
            )
            self.lbl_detection_info.config(
                text=(
                    f"✓ Sheet: '{sheet}'  |  "
                    f"Roll Col: '{col_info['roll_header']}'  |  "
                    f"Name Col: '{col_info['name_header']}'  |  "
                    f"Room Col: '{col_info['room_header']}'"
                )
            )

            # Populate preview table
            for row in self.preview_tree.get_children():
                self.preview_tree.delete(row)

            for s in self.students[:25]:  # Preview first 25
                self.preview_tree.insert("", "end", values=(s.roll, s.name, s.room))

            self.btn_generate.config(state="normal")

        except Exception as e:
            messagebox.showerror(
                "Error Reading Excel File",
                f"Could not read student list:\n\n{e}"
            )

    def _on_generate(self):
        if not self.input_file or not self.students:
            return

        default_out_name = f"Canteen_Register_{self.input_file.stem}.xlsx"
        default_dir = self.input_file.parent

        save_path_str = filedialog.asksaveasfilename(
            title="Save Canteen Register As",
            initialdir=default_dir,
            initialfile=default_out_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if not save_path_str:
            return

        out_path = Path(save_path_str)
        if out_path.resolve() == self.input_file.resolve():
            messagebox.showerror(
                "Invalid Save Location",
                "Cannot overwrite the original input file. Please choose a different filename."
            )
            return

        try:
            generate_canteen_register(self.students, out_path)
            passed, issues = validate_canteen_register(out_path, self.students, self.input_file)

            if not passed:
                messagebox.showwarning(
                    "Validation Warning",
                    f"File generated, but validation noted:\n" + "\n".join(issues[:5])
                )
            else:
                answer = messagebox.askyesno(
                    "Success!",
                    f"Register generated successfully!\n\n"
                    f"• {len(self.students)} students formatted\n"
                    f"• 2 rows per student with merged ID & unmerged record cells\n"
                    f"• Saved to: {out_path.name}\n\n"
                    f"Would you like to open the generated Excel file now?",
                )
                if answer:
                    _open_file(out_path)

        except Exception as e:
            messagebox.showerror(
                "Generation Error",
                f"Failed to generate canteen register:\n\n{e}"
            )


def _open_file(path: Path):
    """Open the generated file with system default program."""
    try:
        if sys.platform == "darwin":
            subprocess.call(["open", str(path)])
        elif sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            subprocess.call(["xdg-open", str(path)])
    except Exception:
        pass


if __name__ == "__main__":
    app = CanteenRegisterApp()
    app.mainloop()
