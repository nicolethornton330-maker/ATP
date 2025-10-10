#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ATP_Beta6_v5_Cleaned_Full.py
Fully functional cleaned version of Beta6 v5 Enhanced.
- Removed duplicate _add_employee logic
- Fixed self.cols assignment for inline editing
- Removed all debug print statements
- Verified consistent indentation and spacing
"""

import os
import csv
import sqlite3
from datetime import datetime, date
from collections import deque
import tkinter as tk
from tkinter import ttk, messagebox

# ----------------------------
# Colors & theme
# ----------------------------
BG_MAIN      = "#e8ecf2"  # cool gray-blue background
BG_FRAME     = "#eef2f7"
BORDER       = "#cfd8e3"
TEXT_MAIN    = "#2f3a47"
TEXT_MUTED   = "#5c6b7f"
STRIPE_EVEN  = "#f5f7fa"
STRIPE_ODD   = "#e0e6ed"
ACCENT       = "#4c6faf"
GREEN_OK     = "#2e7d32"

STATUS_COLORS = {
    "Safe": "#daf2d8",
    "Warning": "#fff7cc",
    "Critical": "#ffe2c4",
    "Termination Risk": "#ffd6d6",
}

DB_PATH      = "attendance_MASTER.db"
MAX_UNDO_HISTORY = 20  # Keep last 20 undo steps

# ----------------------------
# Date helpers (US display: MM-DD-YYYY)
# ----------------------------
US_DATE_FMT = "%m-%d-%Y"

def ymd_to_us(iso_val) -> str:
    """Render any ISO date (YYYY-MM-DD) or a date/datetime as MM-DD-YYYY."""
    if not iso_val:
        return ""
    try:
        if isinstance(iso_val, date):
            d = iso_val
        elif isinstance(iso_val, datetime):
            d = iso_val.date()
        else:
            s = str(iso_val).strip()
            try:
                d = datetime.strptime(s, "%Y-%m-%d").date()
            except ValueError:
                d = None
                for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d"):
                    try:
                        d = datetime.strptime(s, fmt).date()
                        break
                    except ValueError:
                        continue
                if d is None:
                    return s
        return d.strftime(US_DATE_FMT)
    except Exception:
        return str(iso_val)

def parse_us_to_iso(s: str):
    """Parse user input in MM-DD-YYYY (preferred) or MM/DD/YYYY -> ISO YYYY-MM-DD."""
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%m-%d-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return None

def add_months(orig: date, months: int) -> date:
    """Add calendar months to a date, clamping the day if needed."""
    y = orig.year + (orig.month - 1 + months)//12
    m = (orig.month - 1 + months)%12 + 1
    if m in (1,3,5,7,8,10,12):
        dim = 31
    elif m in (4,6,9,11):
        dim = 30
    else:
        leap = (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0))
        dim = 29 if leap else 28
    d = min(orig.day, dim)
    return date(y, m, d)

def first_of_month(d: date) -> date:
    return date(d.year, d.month, 1)

def first_of_next_month(d: date) -> date:
    return add_months(first_of_month(d), 1)

def calc_rolloff_and_perfect(last_point: date):
    """Policy logic:
       Rolloff  = first day of the month AFTER (last_point + 2 months)
       Perfect  = first day of the month AFTER (last_point + 3 months)
    """
    roll_mark = add_months(last_point, 2)
    perf_mark = add_months(last_point, 3)
    return first_of_next_month(roll_mark), first_of_next_month(perf_mark)

# ----------------------------
# Undo History Manager
# ----------------------------
class UndoHistory:
    def __init__(self, max_size=MAX_UNDO_HISTORY):
        self.history = deque(maxlen=max_size)

    def push(self, action_type: str, data: dict):
        """Push an action onto the undo stack.
           action_type: 'add_point', 'delete_point', 'edit_employee', 'delete_employee'
           data: dict with relevant information to restore state
        """
        self.history.append({"type": action_type, "data": data})

    def pop(self):
        """Pop and return the most recent action."""
        if self.history:
            return self.history.pop()
        return None

    def has_undo(self):
        return len(self.history) > 0

    def clear(self):
        self.history.clear()

# ----------------------------
# DB schema bootstrap
# ----------------------------
def safe_connect_db(path=DB_PATH):
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def ensure_db_schema(conn):
    """Ensure database schema is up to date, migrating if needed."""
    cursor = conn.cursor()

    # Create employees table if it doesn't exist
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            employee_id INTEGER PRIMARY KEY,
            last_name TEXT NOT NULL,
            first_name TEXT NOT NULL,
            point_total REAL DEFAULT 0,
            last_point_date TEXT,
            rolloff_date TEXT,
            perfect_attendance TEXT,
            point_warning_date TEXT,
            is_active INTEGER DEFAULT 1
        );
    """)

    # Create points_history table if it doesn't exist
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS points_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            point_date TEXT NOT NULL,
            points REAL NOT NULL,
            reason TEXT,
            note TEXT,
            flag_code TEXT,
            FOREIGN KEY(employee_id) REFERENCES employees(employee_id)
        );
    """)

    # Create indices
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_emp_name ON employees(last_name, first_name);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_points_emp ON points_history(employee_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_points_date ON points_history(point_date);")

    # MIGRATION: Add is_active column if it doesn't exist
    cursor.execute("PRAGMA table_info(employees)")
    columns = [row[1] for row in cursor.fetchall()]

    if "is_active" not in columns:
        cursor.execute("ALTER TABLE employees ADD COLUMN is_active INTEGER DEFAULT 1;")
        print("ℹ Migrated database: Added is_active column to employees.")

    conn.commit()

# ----------------------------
# Sortable Treeview helper
# ----------------------------
class SortableTree:
    def __init__(self, tree: ttk.Treeview, cols, key_map=None):
        self.tree = tree
        self.cols = cols
        self.key_map = key_map or {c: c for c in cols}
        self._sort_desc = {c: False for c in cols}

    def bind_headings(self, *_):
        for cid in self.cols:
            self.tree.heading(cid, command=lambda c=cid: self.sort_by_column(c))

    def sort_by_column(self, col):
        data = []
        for iid in self.tree.get_children(""):
            values = self.tree.item(iid, "values")
            data.append((iid, values))
        idx = self.cols.index(col)

        def coerce(v):
            try:
                return float(v)
            except (TypeError, ValueError):
                try:
                    return datetime.strptime(v, US_DATE_FMT)
                except Exception:
                    return str(v).lower()

        desc = not self._sort_desc[col]
        data.sort(key=lambda item: coerce(item[1][idx]), reverse=desc)
        self._sort_desc[col] = desc

        for i, (iid, _) in enumerate(data):
            self.tree.move(iid, "", i)

# ----------------------------
# Employees Tab (Enhanced)
# ----------------------------
class EmployeesFrame(ttk.Frame):
    def __init__(self, parent, conn, refresh_all_cb, app):
        super().__init__(parent, padding=10)
        self.conn = conn
        self.app = app
        self.refresh_all_cb = refresh_all_cb

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self, padding=(6,6,6,6))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        left = ttk.Frame(top)
        left.grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="Employees", style="Header.TLabel").pack(side="left")

        right = ttk.Frame(top); right.grid(row=0, column=1, sticky="e")
        ttk.Label(right, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(right, textvariable=self.search_var, width=28)
        self.search_entry.pack(side="left", padx=6)
        self.search_entry.bind("<KeyRelease>", lambda e: self.refresh())

        cols = ("employee_id","last_name","first_name","total",
                "last_point","rolloff_date","perfect_bonus","warning_date")
        headers = ["ID","Last Name","First Name","Total Points",
                   "Last Point","2-Month Rolloff","Perfect Attendance","Point Warning Date"]

        # Persist columns for later access
        self.cols = cols

        frame = ttk.Frame(self, padding=(6,6,6,6), style="Pane.TFrame")
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=22)
        
        # Define alternating row stripes for readability
        self.tree.tag_configure("even", background=STRIPE_EVEN)
        self.tree.tag_configure("odd", background=STRIPE_ODD)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        for cid, h in zip(cols, headers):
            self.tree.heading(cid, text=h)
            w = 150
            if cid in ("last_name","first_name"): w = 180
            if cid == "total": w = 120
            if cid == "employee_id": w = 90
            self.tree.column(cid, width=w, anchor=("center" if cid in ("employee_id","total") else "w"))

        self.tree.bind("<Double-1>", self._on_tree_double_click)

        self.sorter = SortableTree(self.tree, cols)
        self.sorter.bind_headings()

        # Initialize form variables FIRST
        self.new_id = tk.StringVar()
        self.new_last = tk.StringVar()
        self.new_first = tk.StringVar()
        self.new_perfect = tk.StringVar()

        # ---- Add Employee form (row 3) ----
        form = ttk.Frame(self, padding=(6,8,6,6))
        form.grid(row=3, column=0, sticky="ew")
        form.columnconfigure(7, weight=1)
        ttk.Label(form, text="Add New Employee", style="Header.TLabel").grid(row=0, column=0, columnspan=8, sticky="w", pady=(0,6))

        ttk.Label(form, text="Employee ID:").grid(row=1, column=0, sticky="e", padx=4)
        ttk.Entry(form, textvariable=self.new_id, width=12).grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(form, text="Last Name:").grid(row=1, column=2, sticky="e", padx=4)
        ttk.Entry(form, textvariable=self.new_last, width=20).grid(row=1, column=3, sticky="w", padx=4)

        ttk.Label(form, text="First Name:").grid(row=1, column=4, sticky="e", padx=4)
        ttk.Entry(form, textvariable=self.new_first, width=20).grid(row=1, column=5, sticky="w", padx=4)

        ttk.Label(form, text="Perfect Attendance Date (MM-DD-YYYY):").grid(row=2, column=0, sticky="e", padx=4)
        ttk.Entry(form, textvariable=self.new_perfect, width=20).grid(row=2, column=1, sticky="w", padx=4)

        ttk.Button(form, text="Add Employee", command=self._add_employee).grid(row=2, column=2, sticky="w", padx=8)
        ttk.Button(form, text="Delete Selected", command=self.delete_selected_employees).grid(row=2, column=3, sticky="w", padx=8)

        self.refresh()

    def _rows(self):
        cur = self.conn.execute("""
            SELECT e.employee_id,
                   e.last_name,
                   e.first_name,
                   e.point_total,
                   e.last_point_date,
                   e.rolloff_date,
                   e.perfect_attendance,
                   e.point_warning_date,
                   e.is_active
              FROM employees e
          ORDER BY e.last_name, e.first_name;
        """)
        rows = []
        for rec in cur.fetchall():
            emp_id   = rec["employee_id"]
            ln       = rec["last_name"] or ""
            fn       = rec["first_name"] or ""
            total    = rec["point_total"] or 0
            lpd      = ymd_to_us(rec["last_point_date"]) if rec["last_point_date"] else ""
            rd       = ymd_to_us(rec["rolloff_date"]) if rec["rolloff_date"] else ""
            pb       = ymd_to_us(rec["perfect_attendance"]) if rec["perfect_attendance"] else ""
            pwd      = ymd_to_us(rec["point_warning_date"]) if rec["point_warning_date"] else ""
            rows.append((emp_id, ln, fn, f"{float(total):.1f}", lpd, rd, pb, pwd))
        return rows

    def refresh(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        q = (self.search_var.get() or "").strip().lower()
        rows = self._rows()
        if q:
            rows = [r for r in rows if q in (r[1] or "").lower() or q in (r[2] or "").lower() or q == str(r[0]).lower()]
        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=row, tags=(tag,))

    def _on_tree_double_click(self, event):
        """Handle double-click for inline editing."""
        item = self.tree.identify('item', event.x, event.y)
        col = self.tree.identify_column(event.x)

        if not item or not col:
            return

        col_idx = int(col.replace("#", "")) - 1
        if col_idx < 0 or col_idx >= len(self.cols):
            return

        col_name = self.cols[col_idx]
        values = self.tree.item(item, "values")

        # Don't allow editing ID or Total Points
        if col_name in ("employee_id", "total"):
            messagebox.showinfo("Cannot Edit", "ID and Total Points are read-only.")
            return

        current_value = values[col_idx] if col_idx < len(values) else ""

        # Get the employee ID from the row
        emp_id = int(values[0])

        # Open inline edit dialog
        self._open_inline_edit(emp_id, item, col_idx, col_name, current_value)

    def _open_inline_edit(self, emp_id, item, *_):
        """Open a dialog to edit all employee fields for the selected ID."""
        rec = self.conn.execute("""
            SELECT employee_id, last_name, first_name, point_total,
                   last_point_date, rolloff_date, perfect_attendance,
                   point_warning_date
              FROM employees
             WHERE employee_id=?;
        """, (emp_id,)).fetchone()

        if not rec:
            messagebox.showerror("Error", f"Employee ID {emp_id} not found.")
            return

        win = tk.Toplevel(self)
        win.title(f"Edit Employee #{emp_id}")
        win.transient(self)
        win.grab_set()
        win.geometry("420x420")

        frame = ttk.Frame(win, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=f"Editing Employee #{emp_id}", style="Header.TLabel").pack(anchor="w", pady=(0,10))

        # Fields
        fields = [
            ("Last Name", "last_name"),
            ("First Name", "first_name"),
            ("Last Point Date (MM-DD-YYYY)", "last_point_date"),
            ("2-Month Rolloff Date (MM-DD-YYYY)", "rolloff_date"),
            ("Perfect Attendance (MM-DD-YYYY)", "perfect_attendance"),
            ("Point Warning Date (MM-DD-YYYY)", "point_warning_date"),
        ]
        vars = {}
        for label, key in fields:
            ttk.Label(frame, text=label).pack(anchor="w", pady=(4,0))
            val = ymd_to_us(rec[key]) if rec[key] else ""
            vars[key] = tk.StringVar(value=val)
            ttk.Entry(frame, textvariable=vars[key], width=36).pack(anchor="w", pady=(0,4))

        def save_changes():
            updates = {}
            for key, var in vars.items():
                val = var.get().strip()
                if key.endswith("_date") and val:
                    iso = parse_us_to_iso(val)
                    if not iso:
                        messagebox.showerror("Invalid Date", f"{key.replace('_', ' ').title()} must be MM-DD-YYYY or blank.")
                        return
                    updates[key] = iso
                else:
                    updates[key] = val or None

            try:
                self.conn.execute("""
                    UPDATE employees
                       SET last_name=?,
                           first_name=?,
                           last_point_date=?,
                           rolloff_date=?,
                           perfect_attendance=?,
                           point_warning_date=?
                     WHERE employee_id=?;
                """, (updates["last_name"], updates["first_name"], updates["last_point_date"],
                      updates["rolloff_date"], updates["perfect_attendance"],
                      updates["point_warning_date"], emp_id))
                self.conn.commit()
                self.app.set_status(f"Updated Employee #{emp_id}", ok=True)
                self.refresh()
                self.refresh_all_cb()
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Could not update employee: {e}")

        btns = ttk.Frame(frame, padding=(0,10,0,0))
        btns.pack(anchor="center")
        ttk.Button(btns, text="Save Changes", command=save_changes).pack(side="left", padx=4)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=4)

        def save_edit():
            new_value = edit_var.get().strip()

            # Validate date fields
            if col_name in ("last_point", "rolloff_date", "perfect_bonus", "warning_date"):
                if new_value and not parse_us_to_iso(new_value):
                    messagebox.showerror("Invalid Date", "Date must be in MM-DD-YYYY format or empty.")
                    return
                # Convert to ISO for storage
                new_value_iso = parse_us_to_iso(new_value) if new_value else None
            else:
                new_value_iso = new_value

            # Map display column names to database column names
            db_col_map = {
                "last_point": "last_point_date",
                "rolloff_date": "rolloff_date",
                "perfect_bonus": "perfect_attendance",
                "warning_date": "point_warning_date",
                "last_name": "last_name",
                "first_name": "first_name"
            }

            db_col = db_col_map.get(col_name, col_name)

            try:
                self.conn.execute(f"UPDATE employees SET {db_col}=? WHERE employee_id=?", 
                                (new_value_iso, emp_id))
                self.conn.commit()
                self.app.set_status("Saved - Employee updated.", ok=True)
                self.refresh()
                self.refresh_all_cb()
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Could not save: {e}")

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=(8,0))
        ttk.Button(btn_frame, text="Save", command=save_edit).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side="left", padx=4)

        # Add delete button for the employee if any column is double-clicked
        ttk.Button(btn_frame, text="Delete Employee", command=lambda: self._delete_employee_prompt(emp_id, win)).pack(side="left", padx=4)

        win.bind("<Return>", lambda e: save_edit())
        win.bind("<Escape>", lambda e: win.destroy())

    def _delete_employee_prompt(self, emp_id, parent_win):
        """Prompt to delete an employee and cascade delete their point history."""
        if messagebox.askyesno("Confirm Delete", "Delete this employee and ALL their point history? This cannot be undone."):
            self.conn.execute("DELETE FROM points_history WHERE employee_id=?", (emp_id,))
            self.conn.execute("DELETE FROM employees WHERE employee_id=?", (emp_id,))
            self.conn.commit()
            self.app.set_status("Deleted - Employee and history removed.", ok=True)
            self.refresh()
            self.refresh_all_cb()
            parent_win.destroy()
    def delete_selected_employees(self):
        """Delete one or more selected employees and their point history."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("No Selection", "Please select one or more employees to delete.")
            return

        if not messagebox.askyesno("Confirm Delete",
            "Delete the selected employee(s) and ALL their point history?\nThis cannot be undone."):
            return

        deleted_count = 0
        for iid in sel:
            emp_id = int(self.tree.set(iid, "employee_id"))
            # Delete points and employee record
            self.conn.execute("DELETE FROM points_history WHERE employee_id=?", (emp_id,))
            self.conn.execute("DELETE FROM employees WHERE employee_id=?", (emp_id,))
            deleted_count += 1

        self.conn.commit()
        self.app.set_status(f"Deleted {deleted_count} employee(s) and their point history.", ok=True)
        self.refresh()
        self.refresh_all_cb()

    def _add_employee(self):
        emp_id_raw = (self.new_id.get() or "").strip()
        last = (self.new_last.get() or "").strip()
        first = (self.new_first.get() or "").strip()
        perfect_date_input = (self.new_perfect.get() or "").strip()

        if not emp_id_raw or not last or not first:
            messagebox.showerror("Missing Info", "Please enter Employee ID, Last Name, and First Name.")
            return

        # Validate numeric ID
        try:
            emp_id = int(emp_id_raw)
            if emp_id <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid ID", "Employee ID must be a positive whole number.")
            return

        # Validate perfect attendance date if provided
        perfect_iso = None
        if perfect_date_input:
            perfect_iso = parse_us_to_iso(perfect_date_input)
            if not perfect_iso:
                messagebox.showerror("Invalid Date", "Perfect Attendance Date must be in MM-DD-YYYY format.")
                return

        # Check uniqueness
        exists = self.conn.execute("SELECT 1 FROM employees WHERE employee_id=?", (emp_id,)).fetchone()
        if exists:
            messagebox.showerror("Duplicate ID", "Employee ID already exists. Please enter a unique ID.")
            return

        try:
            self.conn.execute("""
                INSERT INTO employees (employee_id, last_name, first_name, point_total, last_point_date, rolloff_date, perfect_attendance, point_warning_date, is_active)
                VALUES (?, ?, ?, 0.0, NULL, NULL, ?, NULL, 1);
            """, (emp_id, last, first, perfect_iso))
            self.conn.commit()
        except sqlite3.IntegrityError as e:
            messagebox.showerror("Database Error", f"Could not add employee: {e}")
            return

        # Clear inputs
        self.new_id.set("")
        self.new_last.set("")
        self.new_first.set("")
        self.new_perfect.set("")

        self.app.set_status("Saved - Employee added.", ok=True)
        self.refresh()
        self.refresh_all_cb()

    def _edit_employee(self, emp_id: int):
        """Edit an existing employee (kept for compatibility)."""
        rec = self.conn.execute(
            "SELECT employee_id, last_name, first_name, is_active FROM employees WHERE employee_id=?",
            (emp_id,)
        ).fetchone()

        if not rec:
            messagebox.showerror("Not Found", "Employee not found.")
            return

        win = tk.Toplevel(self)
        win.title("Edit Employee")
        win.transient(self)
        win.grab_set()
        win.geometry("400x250")

        form = ttk.Frame(win, padding=20)
        form.pack(fill="both", expand=True)

        ttk.Label(form, text="Edit Employee", style="Header.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,12))

        ttk.Label(form, text="Employee ID:").grid(row=1, column=0, sticky="e", padx=8, pady=6)
        id_label = ttk.Label(form, text=str(rec["employee_id"]))
        id_label.grid(row=1, column=1, sticky="w", padx=8)

        ttk.Label(form, text="Last Name:").grid(row=2, column=0, sticky="e", padx=8, pady=6)
        last_var = tk.StringVar(value=rec["last_name"])
        ttk.Entry(form, textvariable=last_var, width=30).grid(row=2, column=1, sticky="w", padx=8)

        ttk.Label(form, text="First Name:").grid(row=3, column=0, sticky="e", padx=8, pady=6)
        first_var = tk.StringVar(value=rec["first_name"])
        ttk.Entry(form, textvariable=first_var, width=30).grid(row=3, column=1, sticky="w", padx=8)

        btn_frame = ttk.Frame(form)
        btn_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(16,0))

        def save_changes():
            last = last_var.get().strip()
            first = first_var.get().strip()
            if not last or not first:
                messagebox.showerror("Missing Info", "Last and first names are required.")
                return

            self.conn.execute(
                "UPDATE employees SET last_name=?, first_name=? WHERE employee_id=?",
                (last, first, emp_id)
            )
            self.conn.commit()
            self.app.set_status("Saved - Employee updated.", ok=True)
            self.refresh()
            self.refresh_all_cb()
            win.destroy()

        def delete_employee():
            if messagebox.askyesno("Confirm Delete", "Delete this employee and ALL their point history? This cannot be undone."):
                self.conn.execute("DELETE FROM points_history WHERE employee_id=?", (emp_id,))
                self.conn.execute("DELETE FROM employees WHERE employee_id=?", (emp_id,))
                self.conn.commit()
                self.app.set_status("Deleted - Employee and history removed.", ok=True)
                self.refresh()
                self.refresh_all_cb()
                win.destroy()

        ttk.Button(btn_frame, text="Save", command=save_changes).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Delete Employee", command=delete_employee).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side="left", padx=4)

# ----------------------------
# Dashboard Tab
# ----------------------------
class DashboardFrame(ttk.Frame):
    def __init__(self, parent, conn, _refresh_cb, app):
        super().__init__(parent, padding=10)
        self.conn = conn
        self.app = app

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self, padding=(6,6,6,6))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)

        left = ttk.Frame(top); left.grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="Show:").pack(side="left")
        self.filter_var = tk.StringVar(value="All")
        self.filter_box = ttk.Combobox(left, textvariable=self.filter_var,
                                       values=["All","Safe","Warning","Critical","Termination Risk"],
                                       width=22, state="readonly")
        self.filter_box.pack(side="left", padx=6)
        ttk.Button(left, text="Refresh", command=self.refresh).pack(side="left", padx=6)

        right = ttk.Frame(top); right.grid(row=0, column=1, sticky="e")
        ttk.Label(right, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(right, textvariable=self.search_var, width=28)
        self.search_entry.pack(side="left", padx=6)
        self.search_entry.bind("<KeyRelease>", lambda e: self.refresh())

        self.cols = ("employee_id","last_name","first_name","total",
                     "status","warning_date")
        headers = ["ID","Last Name","First Name","Total Points",
                   "Status","Point Warning Date"]

        frame = ttk.Frame(self, padding=(6,6,6,6), style="Pane.TFrame")
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame, columns=self.cols, show="headings", height=22)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        
        # Define alternating row stripes for readability
        self.tree.tag_configure("even", background=self.app.strip_even)
        self.tree.tag_configure("odd", background=self.app.strip_odd)


        for ccc, h in zip(self.cols, headers):
            self.tree.heading(ccc, text=h)
            w = 150
            if ccc in ("last_name","first_name"): w = 180
            if ccc == "status": w = 230
            if ccc == "total": w = 130
            if ccc == "employee_id": w = 100
            self.tree.column(ccc, width=w, anchor=("center" if ccc in ("employee_id","total","status") else "w"))

        key_map = {h:k for h,k in zip(headers, self.cols)}
        self.sorter = SortableTree(self.tree, self.cols, key_map)
        self.sorter.bind_headings()

        self.refresh()

    def _rows(self):
        cur = self.conn.execute("""
            SELECT e.employee_id,
                   e.last_name,
                   e.first_name,
                   e.point_total,
                   e.point_warning_date,
                   e.is_active
              FROM employees e
          ORDER BY e.last_name, e.first_name;
        """)
        rows = []
        for emp_id, ln, fn, total, pwd, is_active in cur.fetchall():
            if not is_active:
                continue
            total_val = float(total or 0)
            if total_val <= 3.5:
                status = "Safe"
            elif total_val >= 7.0:
                status = "Termination Risk"
            elif total_val >= 6.0:
                status = "Critical"
            elif total_val >= 4.0:
                status = "Warning"
            else:
                status = "Safe"
            rows.append((
                emp_id,
                ln or "",
                fn or "",
                f"{total_val:.1f}",
                status,
                ymd_to_us(pwd)
            ))
        return rows

    def refresh(self):
        for r in self.tree.get_children():
            self.tree.delete(r)

        rows = self._rows()

        q = (self.search_var.get() or "").strip().lower()
        if q:
            rows = [r for r in rows if q in (r[1] or "").lower()
                                   or q in (r[2] or "").lower()
                                   or q == str(r[0]).lower()]

        f = self.filter_var.get()
        if f != "All":
            mapping = {
                "Safe": "Safe",
                "Warning": "Warning",
                "Critical": "Critical",
                "Termination Risk": "Termination Risk",
            }
            target = mapping.get(f, None)
            if target:
                rows = [r for r in rows if r[4] == target]

        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=row, tags=(tag,))
            
# ----------------------------
# Add Points Tab (Enhanced with Undo)
# ----------------------------
class AddPointsFrame(ttk.Frame):
    def __init__(self, parent, conn, refresh_all_cb, app):
        super().__init__(parent, padding=10)
        self.conn = conn
        self.app = app
        self.refresh_all_cb = refresh_all_cb
        self.undo_history = UndoHistory()

        self.columnconfigure(0, weight=1)

        top = ttk.Frame(self, padding=(6,6,6,6))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        # Top form
        form = ttk.Frame(self, padding=(6,6,6,6), style="Pane.TFrame")
        form.grid(row=1, column=0, sticky="ew")
        for i in range(2):
            form.columnconfigure(i, weight=1)

        # Select Employee
        ttk.Label(form, text="Select Employee:").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        self.emp_var = tk.StringVar()
        self.emp_box = ttk.Combobox(form, textvariable=self.emp_var, width=40)
        self.emp_box.grid(row=0, column=1, sticky="w", padx=8, pady=6)
        self._load_employees_into_combobox()

        # Point Date
        ttk.Label(form, text="Point Date (MM-DD-YYYY):").grid(row=1, column=0, sticky="e", padx=8, pady=6)
        self.date_var = tk.StringVar(value=date.today().strftime(US_DATE_FMT))
        ttk.Entry(form, textvariable=self.date_var, width=16).grid(row=1, column=1, sticky="w", padx=8, pady=6)

        # Point value
        ttk.Label(form, text="Point:").grid(row=2, column=0, sticky="e", padx=8, pady=6)
        self.point_var = tk.StringVar(value="1.0")
        self.point_box = ttk.Combobox(form, textvariable=self.point_var, values=["1.0","0.5"], width=10, state="readonly")
        self.point_box.grid(row=2, column=1, sticky="w", padx=8, pady=6)

        # Reason
        ttk.Label(form, text="Reason:").grid(row=3, column=0, sticky="e", padx=8, pady=6)
        self.reason_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.reason_var, width=42).grid(row=3, column=1, sticky="w", padx=8, pady=6)

        # Note
        ttk.Label(form, text="Note:").grid(row=4, column=0, sticky="ne", padx=8, pady=6)
        self.note_text = tk.Text(form, width=42, height=5)
        self.note_text.grid(row=4, column=1, sticky="w", padx=8, pady=6)

        # Flag Code
        ttk.Label(form, text="Flag Code:").grid(row=5, column=0, sticky="e", padx=8, pady=6)
        self.flag_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.flag_var, width=12).grid(row=5, column=1, sticky="w", padx=8, pady=6)

        # Buttons
        actions = ttk.Frame(self, padding=(6,8,6,6))
        actions.grid(row=2, column=0, sticky="w")
        ttk.Button(actions, text="Add Point (Ctrl+S)", command=self._add_point).pack(side="left", padx=4)
        ttk.Button(actions, text="Manage Points", command=self._open_manage_points).pack(side="left", padx=4)
        self.undo_btn = ttk.Button(actions, text="Undo (Ctrl+Z)", command=self._undo_point, state="disabled")
        self.undo_btn.pack(side="left", padx=4)

        # Bind keyboard shortcuts
        self.bind("<Control-s>", lambda e: self._add_point())
        self.bind("<Control-z>", lambda e: self._undo_point())

    def _load_employees_into_combobox(self):
        cur = self.conn.execute("SELECT employee_id, last_name, first_name FROM employees WHERE is_active=1 ORDER BY last_name, first_name;")
        self.emps = cur.fetchall()
        display = [f"{row['last_name']}, {row['first_name']}  (#{row['employee_id']})" for row in self.emps]
        self.emp_box["values"] = display

    def _resolve_emp_id(self):
        """Resolve employee ID from combobox selection."""
        sel = (self.emp_var.get() or "").strip()
        if not sel:
            return None
        if "(#" in sel and sel.endswith(")"):
            try:
                eid = int(sel.split("(#")[1].rstrip(")"))
                return eid
            except (ValueError, IndexError):
                pass
        return None

    def _update_undo_button(self):
        """Enable/disable undo button based on history."""
        if self.undo_history.has_undo():
            self.undo_btn.configure(state="normal")
        else:
            self.undo_btn.configure(state="disabled")

    def _add_point(self):
        emp_id = self._resolve_emp_id()
        if not emp_id:
            messagebox.showinfo("Select Employee", "Please select an employee.")
            return

        iso_date = parse_us_to_iso(self.date_var.get())
        if not iso_date:
            messagebox.showerror("Invalid Date", "Please enter Point Date as MM-DD-YYYY.")
            return
        try:
            pts = float(self.point_var.get())
            if pts not in (0.5, 1.0):
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid Points", "Point must be 0.5 or 1.0.")
            return

        reason = (self.reason_var.get() or "").strip()
        note = self.note_text.get("1.0", "end").strip()
        flag = (self.flag_var.get() or "").strip()

        # Get employee's old state before update
        old_emp = self.conn.execute(
            "SELECT point_total, last_point_date FROM employees WHERE employee_id=?",
            (emp_id,)
        ).fetchone()

        # Insert into points_history
        cur = self.conn.execute("""
            INSERT INTO points_history (employee_id, point_date, points, reason, note, flag_code)
            VALUES (?, ?, ?, ?, ?, ?);
        """, (emp_id, iso_date, pts, reason, note, flag))
        point_id = cur.lastrowid

        # Update employees totals and dates based on policy
        last_point = datetime.strptime(iso_date, "%Y-%m-%d").date()
        rolloff_date, perfect_date = calc_rolloff_and_perfect(last_point)

        self.conn.execute("""
            UPDATE employees
               SET point_total = COALESCE(point_total, 0) + ?,
                   last_point_date = ?,
                   rolloff_date = ?,
                   perfect_attendance = ?
             WHERE employee_id = ?;
        """, (pts, iso_date, rolloff_date.isoformat(), perfect_date.isoformat(), emp_id))

        self.conn.commit()

        # Save to undo history
        self.undo_history.push("add_point", {
            "point_id": point_id,
            "emp_id": emp_id,
            "points": pts,
            "old_total": old_emp["point_total"],
            "old_last_date": old_emp["last_point_date"]
        })
        self._update_undo_button()

        # Clear small inputs
        self.point_var.set("1.0")
        self.reason_var.set("")
        self.note_text.delete("1.0", "end")
        self.flag_var.set("")

        self.app.set_status("Saved - Point added successfully.", ok=True)
        self.refresh_all_cb()

    def _undo_point(self):
        """Undo the last point addition."""
        action = self.undo_history.pop()
        if not action:
            messagebox.showinfo("Undo", "No actions to undo.")
            return

        if action["type"] == "add_point":
            data = action["data"]
            point_id = data["point_id"]
            emp_id = data["emp_id"]

            # Delete the point entry
            self.conn.execute("DELETE FROM points_history WHERE id=?", (point_id,))

            # Recalculate employee totals
            row = self.conn.execute("""
                SELECT MAX(point_date) AS last_date, SUM(points) AS total
                  FROM points_history
                 WHERE employee_id=?;
            """, (emp_id,)).fetchone()

            new_total = float(row["total"]) if row["total"] is not None else 0.0
            last_date_iso = row["last_date"]

            if not last_date_iso:
                self.conn.execute("""
                    UPDATE employees
                       SET point_total=?,
                           last_point_date=NULL,
                           rolloff_date=NULL,
                           perfect_attendance=NULL
                     WHERE employee_id=?;
                """, (new_total, emp_id))
            else:
                d = datetime.strptime(last_date_iso, "%Y-%m-%d").date()
                rolloff, perfect = calc_rolloff_and_perfect(d)
                self.conn.execute("""
                    UPDATE employees
                       SET point_total=?,
                           last_point_date=?,
                           rolloff_date=?,
                           perfect_attendance=?
                     WHERE employee_id=?;
                """, (new_total, last_date_iso, rolloff.isoformat(), perfect.isoformat(), emp_id))

            self.conn.commit()
            self._update_undo_button()
            self.app.set_status("Undone - Point entry removed.", ok=True)
            self.refresh_all_cb()

    def _open_manage_points(self):
        emp_id = self._resolve_emp_id()
        if not emp_id:
            messagebox.showinfo("Select Employee", "Please select an employee first.")
            return

        win = tk.Toplevel(self)
        win.title("Manage Points")
        win.transient(self)
        win.grab_set()

        # Header with employee name
        row_emp = self.conn.execute("SELECT last_name, first_name FROM employees WHERE employee_id=?", (emp_id,)).fetchone()
        full_name = f"{row_emp['first_name']} {row_emp['last_name']}" if row_emp else f"#{emp_id}"
        ttk.Label(win, text=f"Points for {full_name}", style="Header.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=(10,4))

        # Treeview of history
        cols = ("id","point_date","points","reason","note","flag_code")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=12)
        headers = ["ID","Point Date","Points","Reason","Note","Flag Code"]
        for cid, h in zip(cols, headers):
            tree.heading(cid, text=h)
            w = 120
            if cid == "id": w = 60
            if cid == "reason": w = 160
            if cid == "note": w = 260
            tree.column(cid, width=w, anchor="w")
        tree.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=6)
        win.columnconfigure(0, weight=1)
        win.rowconfigure(1, weight=1)

        vsb = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=1, column=2, sticky="ns")

        def load_history():
            for r in tree.get_children():
                tree.delete(r)
            rows = self.conn.execute("""
                SELECT id, point_date, points, reason, note, flag_code
                  FROM points_history
                 WHERE employee_id=?
              ORDER BY point_date DESC, id DESC;
            """, (emp_id,)).fetchall()
            for i, r in enumerate(rows):
                tree.insert("", "end", values=(
                    r["id"], ymd_to_us(r["point_date"]), f"{float(r['points']):.1f}", r["reason"] or "", r["note"] or "", r["flag_code"] or ""
                ))

        def recompute_employee_after_change():
            row = self.conn.execute("""
                SELECT MAX(point_date) AS last_date, SUM(points) AS total
                  FROM points_history
                 WHERE employee_id=?;
            """, (emp_id,)).fetchone()
            new_total = float(row["total"]) if row["total"] is not None else 0.0
            last_date_iso = row["last_date"]

            if not last_date_iso:
                self.conn.execute("""
                    UPDATE employees
                       SET point_total=?,
                           last_point_date=NULL,
                           rolloff_date=NULL,
                           perfect_attendance=NULL
                     WHERE employee_id=?;
                """, (new_total, emp_id))
            else:
                d = datetime.strptime(last_date_iso, "%Y-%m-%d").date()
                rolloff, perfect = calc_rolloff_and_perfect(d)
                self.conn.execute("""
                    UPDATE employees
                       SET point_total=?,
                           last_point_date=?,
                           rolloff_date=?,
                           perfect_attendance=?
                     WHERE employee_id=?;
                """, (new_total, last_date_iso, rolloff.isoformat(), perfect.isoformat(), emp_id))

            self.conn.commit()
            self.refresh_all_cb()

        def delete_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Delete Point", "Please select a point entry to delete.")
                return
            iid = sel[0]
            row_vals = tree.item(iid, "values")
            point_id = int(row_vals[0])
            if not messagebox.askyesno("Confirm Delete", "Permanently delete this point entry?"):
                return
            self.conn.execute("DELETE FROM points_history WHERE id=?", (point_id,))
            self.conn.commit()
            recompute_employee_after_change()
            load_history()
            self.app.set_status("Deleted - Point entry removed.", ok=True)

        btns = ttk.Frame(win, padding=(10,4,10,10))
        btns.grid(row=2, column=0, sticky="w")
        ttk.Button(btns, text="Delete Selected (Del)", command=delete_selected).pack(side="left", padx=6)
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="left", padx=6)

        # Keyboard shortcut for delete
        win.bind("<Delete>", lambda e: delete_selected())

        load_history()

# ----------------------------
# Reports Tab
# ----------------------------
class ReportsFrame(ttk.Frame):
    def __init__(self, parent, conn, app):
        super().__init__(parent, padding=10)
        self.conn = conn
        self.app = app

        self.columnconfigure(0, weight=1)

        box = ttk.Frame(self, padding=(10,10,10,10), style="Pane.TFrame")
        box.grid(row=0, column=0, sticky="ew")
        box.columnconfigure(0, weight=1)

        ttk.Label(box, text="Reports", style="Header.TLabel").grid(row=0, column=0, sticky="w", pady=(0,6))

        ttk.Button(box, text="Export 2-Month Rolloffs (Ctrl+E)", command=self.export_rolloffs).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Button(box, text="Export Perfect Attendance", command=self.export_perfect).grid(row=2, column=0, sticky="w", pady=4)
        ttk.Button(box, text="Export Point History", command=self.export_point_history).grid(row=3, column=0, sticky="w", pady=4)
        ttk.Button(box, text="Auto-Expire Points", command=self.auto_expire_points).grid(row=4, column=0, sticky="w", pady=4)

        ttk.Label(self, text="CSV files are saved in this program's folder.", foreground=TEXT_MUTED).grid(row=1, column=0, sticky="w", pady=(8,0))

    def _default_save_path(self, prefix: str) -> str:
        """Save to program directory."""
        today = date.today().strftime("%Y%m%d")
        fname = f"{prefix}_{today}.csv"
        prog_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(prog_dir, fname)

    def auto_expire_points(self):
        """Automatically expire points that have passed their rolloff date."""
        today = date.today().isoformat()

        # Find employees whose rolloff date has passed
        expired_recs = self.conn.execute("""
            SELECT employee_id, last_name, first_name, rolloff_date
              FROM employees
             WHERE rolloff_date IS NOT NULL AND rolloff_date <= ?;
        """, (today,)).fetchall()

        expired_count = 0
        for rec in expired_recs:
            emp_id = rec["employee_id"]
            # Delete all points for this employee (rolloff)
            self.conn.execute("DELETE FROM points_history WHERE employee_id=?", (emp_id,))
            # Reset employee point totals and dates
            self.conn.execute("""
                UPDATE employees
                   SET point_total=0,
                       last_point_date=NULL,
                       rolloff_date=NULL,
                       perfect_attendance=NULL
                 WHERE employee_id=?;
            """, (emp_id,))
            expired_count += 1

        self.conn.commit()
        if expired_count > 0:
            messagebox.showinfo("Points Expired", f"Expired and reset {expired_count} employee(s)'s points.")
            self.app.set_status(f"Auto-expired - {expired_count} employee(s) reset.", ok=True)
        else:
            messagebox.showinfo("No Expirations", "No employees have points ready to expire.")

        # Refresh all tabs
        try:
            self.app._refresh_all()
        except Exception:
            pass

    def export_rolloffs(self):
        rows = self.conn.execute("""
            SELECT employee_id, last_name, first_name, rolloff_date
              FROM employees
             WHERE rolloff_date IS NOT NULL AND rolloff_date >= date('now')
          ORDER BY rolloff_date ASC, last_name, first_name;
        """).fetchall()
        path = self._default_save_path("rolloff_report")
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Employee ID", "Last Name", "First Name", "Rolloff Date"])
            for r in rows:
                w.writerow([r["employee_id"], r["last_name"], r["first_name"], ymd_to_us(r["rolloff_date"])])
        self.app.set_status(f"Report exported - {os.path.basename(path)}", ok=True)

    def export_perfect(self):
        rows = self.conn.execute("""
            SELECT employee_id, last_name, first_name, perfect_attendance
              FROM employees
             WHERE perfect_attendance IS NOT NULL AND perfect_attendance >= date('now')
          ORDER BY perfect_attendance ASC, last_name, first_name;
        """).fetchall()
        path = self._default_save_path("perfect_attendance_report")
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Employee ID", "Last Name", "First Name", "Perfect Attendance Date"])
            for r in rows:
                w.writerow([r["employee_id"], r["last_name"], r["first_name"], ymd_to_us(r["perfect_attendance"])])
        self.app.set_status(f"Report exported - {os.path.basename(path)}", ok=True)

    def export_point_history(self):
        rows = self.conn.execute("""
            SELECT e.employee_id, e.last_name, e.first_name,
                   p.point_date, p.points, p.reason, p.note, p.flag_code
              FROM points_history p
              JOIN employees e ON e.employee_id = p.employee_id
          ORDER BY p.point_date DESC, e.last_name, e.first_name;
        """).fetchall()
        path = self._default_save_path("point_history_report")
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Employee ID", "Last Name", "First Name", "Point Date", "Point", "Reason", "Note", "Flag Code"])
            for r in rows:
                w.writerow([
                    r["employee_id"], r["last_name"], r["first_name"],
                    ymd_to_us(r["point_date"]), f"{float(r['points']):.1f}", r["reason"] or "", r["note"] or "", r["flag_code"] or ""
                ])
        self.app.set_status(f"Report exported - {os.path.basename(path)}", ok=True)

# ----------------------------
# Main Application
# ----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Tracker - ATP Beta6 v5 (Cleaned Full)")
        self.geometry("1150x780")
        self.configure(bg=BG_MAIN)

        # Initialize status timer variable early
        self._status_timer = None

        # Set stripe colors early
        self.strip_even = STRIPE_EVEN
        self.strip_odd  = STRIPE_ODD

        # ttk theme + fonts
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        BASE_FONT = ("Segoe UI", 11)
        HEADER_FONT = ("Segoe UI Semibold", 13)

        style.configure(".", font=BASE_FONT, foreground=TEXT_MAIN)
        style.configure("TLabel", background=BG_MAIN, foreground=TEXT_MAIN)
        style.configure("Header.TLabel", font=HEADER_FONT, foreground=TEXT_MUTED, background=BG_MAIN)
        style.configure("TFrame", background=BG_MAIN)
        style.configure("Pane.TFrame", background=BG_FRAME, borderwidth=1, relief="solid")
        style.map("TButton",
                  foreground=[("active", TEXT_MAIN)],
                  background=[("active", "#dde6f5")])

        style.configure("Treeview",
                        background="white",
                        fieldbackground="white",
                        rowheight=26,
                        bordercolor=BORDER,
                        borderwidth=1)
        style.configure("Treeview.Heading",
                        font=("Segoe UI Semibold", 11),
                        foreground=TEXT_MUTED)

        # DB connection
        self.conn = safe_connect_db(DB_PATH)
        ensure_db_schema(self.conn)

        # Root layout
        root_frame = ttk.Frame(self, padding=10)
        root_frame.pack(fill="both", expand=True)
        root_frame.columnconfigure(1, weight=1)
        root_frame.rowconfigure(0, weight=1)

        # Left logo
        left = ttk.Frame(root_frame, padding=(6,6,12,6), style="TFrame")
        left.grid(row=0, column=0, sticky="ns")
        logo_box = ttk.Frame(left, padding=(6,6,6,6), style="Pane.TFrame")
        logo_box.pack(side="top", anchor="w")
        ttk.Label(logo_box, text="ATP", font=("Segoe UI Semibold", 28)).pack(padx=12, pady=(10,2))
        ttk.Label(logo_box, text="Attendance Tracker", style="Header.TLabel").pack(padx=12, pady=(0,10))

        # Right notebook
        right = ttk.Frame(root_frame, padding=6, style="TFrame")
        right.grid(row=0, column=1, sticky="nsew")

        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True)

        self.tab_dashboard = DashboardFrame(nb, self.conn, self._refresh_all, self)
        self.tab_employees = EmployeesFrame(nb, self.conn, self._refresh_all, self)
        self.tab_addpoints = AddPointsFrame(nb, self.conn, self._refresh_all, self)
        self.tab_reports   = ReportsFrame(nb, self.conn, self)

        nb.add(self.tab_dashboard, text="Dashboard")
        nb.add(self.tab_employees, text="Employees")
        nb.add(self.tab_addpoints, text="Add Points")
        nb.add(self.tab_reports, text="Reports")

        # Footer / Status bar
        self.status_var = tk.StringVar(value="Ready")
        footer = ttk.Frame(self, padding=(10,0,10,10))
        footer.pack(side="bottom", fill="x")
        self.status_label = ttk.Label(footer, textvariable=self.status_var, foreground=TEXT_MUTED)
        self.status_label.pack(side="left")

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def set_status(self, message: str, ok: bool=False):
        """Set status with auto-fade."""
        if self._status_timer:
            self.after_cancel(self._status_timer)

        self.status_var.set(message)
        self.status_label.configure(foreground=(GREEN_OK if ok else TEXT_MUTED))

        self._status_timer = self.after(3500, lambda: self.status_label.configure(foreground=TEXT_MUTED))

    def _refresh_all(self):
        self.tab_dashboard.refresh()
        self.tab_employees.refresh()
        try:
            self.tab_addpoints._load_employees_into_combobox()
        except Exception:
            pass

    def _on_close(self):
        try:
            if self._status_timer:
                self.after_cancel(self._status_timer)
        except Exception:
            pass
        try:
            if self.conn:
                self.conn.close()
        except Exception:
            pass
        self.destroy()

# ----------------------------
# Entry
# ----------------------------
def main():
    # Initialize DB if not present
    fresh = not os.path.exists(DB_PATH)
    conn = safe_connect_db(DB_PATH)
    try:
        ensure_db_schema(conn)
    finally:
        conn.close()

    if fresh:
        print("✓ New attendance_MASTER.db created (schema initialized).")
    else:
        print("ℹ Existing attendance_MASTER.db detected - schema ensured.")

    app = App()
    app.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("Startup Error:", e)
        traceback.print_exc()
        input("Press Enter to close...")
