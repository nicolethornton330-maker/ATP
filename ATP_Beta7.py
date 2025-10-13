#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ATP_Beta7
-----------
Stable Release — October 2025  
Author: Nicole Thornton  

Version Summary:
• Added “Entry ID” column to Point History export for full audit traceability.  
• Added running “Point Total” column at the end of each exported row for HRIS compatibility.  
• Verified legacy report exports (Rolloffs, Perfect Attendance) and Auto-Expire logic remain stable.  
• Preserved full backward compatibility with existing Beta 6 databases.  
• UI, schema, and logic synchronized for internal HR deployment.

Status:
✅ Official Locked Build — ATP_Beta7 (Baseline for all future enhancements)
"""

import os
import csv
import sqlite3
import calendar
from datetime import datetime, date
from collections import deque
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from PIL import Image, ImageTk
from tkinter import ttk, messagebox, filedialog
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
from datetime import date, datetime

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

def two_months_then_first(d: date) -> date:
    return first_of_next_month(add_months(d, 2))

def three_months_then_first(d: date) -> date:
    return first_of_next_month(add_months(d, 3))

def step_next_due(current_due: date, perfect_date: date) -> date:
    """
    Advance the next-due date one step:
      - If we haven't passed the perfect-attendance milestone yet,
        jump to 2 months after the perfect month (then first-of-next).
      - Otherwise, move forward by 2 months (then first-of-next).
    """
    if current_due < perfect_date:
        return two_months_then_first(perfect_date)
    return two_months_then_first(current_due)

def calc_rolloff_and_perfect(last_point: date):
    """Policy logic:
       Rolloff  = first day of the month AFTER (last_point + 2 months)
       Perfect  = first day of the month AFTER (last_point + 3 months)
    """
    roll_mark = add_months(last_point, 2)
    perf_mark = add_months(last_point, 3)
    return first_of_next_month(roll_mark), first_of_next_month(perf_mark)

# --- PATCH: reason choices helper -------------------------------------------
def get_reason_options(conn):
    """Return a de-duplicated, nicely ordered list of reasons.
    Starts with sensible defaults, then adds distinct reasons found in DB."""
    defaults = [
        "Tardy/Early Leave",
        "Absence", 
        "No Call/No Show",
    ]
    try:
        rows = conn.execute(
            "SELECT DISTINCT reason FROM points_history "
            "WHERE reason IS NOT NULL AND TRIM(reason) <> '' "
            "ORDER BY reason COLLATE NOCASE"
        ).fetchall()
        seen = set()
        ordered = []
        # defaults first
        for r in defaults:
            k = r.strip()
            if k and k.lower() not in seen:
                ordered.append(k); seen.add(k.lower())
        # then DB values
        for (r,) in rows:
            k = (r or "").strip()
            if k and k.lower() not in seen:
                ordered.append(k); seen.add(k.lower())
        return ordered
    except Exception:
        # if anything goes sideways, fall back to defaults
        return defaults[:]
# ---------------------------------------------------------------------------

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
                   "Last Point","2-Month Rolloff","Perfect Attendance","Warning Issued"]

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
            self.tree.column(cid, anchor=("center" if cid in ("employee_id","total") else "w"), stretch=True)

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
        ttk.Button(form, text="Import from CSV", command=self._import_employees).grid(row=2, column=4, sticky="w", padx=8)
        form.bind("<Return>", lambda e: self._add_employee())
        self.tree.unbind("<Double-1>")
        self.refresh()
        self.tree.bind("<Double-1>", self._on_tree_double_click)


        self.refresh()

    def _import_employees(self):
        """Import employees from a CSV file with validation and overwrite prompts."""
        path = filedialog.askopenfilename(
            title="Select Employee CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if not path:
            return

        added = skipped = overwritten = 0

        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    emp_id_raw = (row.get("employee_id") or "").strip()
                    last = (row.get("last_name") or "").strip()
                    first = (row.get("first_name") or "").strip()
                    total_raw = (row.get("point_total") or "").strip()
                    lp_input = (row.get("last_point_date") or "").strip()
                    roll_input = (row.get("rolloff_date") or "").strip()
                    perfect_input = (row.get("perfect_attendance_date") or "").strip()

                    # --- basic validation ---
                    if not emp_id_raw or not last or not first:
                        skipped += 1
                        continue
                    try:
                        emp_id = int(emp_id_raw)
                        if emp_id <= 0:
                            raise ValueError
                    except Exception:
                        skipped += 1
                        continue

                    # --- validate point total ---
                    try:
                        total = float(total_raw) if total_raw else 0.0
                    except Exception:
                        total = 0.0

                    # --- date conversion ---
                    lp_iso = parse_us_to_iso(lp_input) if lp_input else None
                    roll_iso = parse_us_to_iso(roll_input) if roll_input else None
                    perfect_iso = parse_us_to_iso(perfect_input) if perfect_input else None
                    if any(
                        d for d, raw in [
                            (lp_iso, lp_input),
                            (roll_iso, roll_input),
                            (perfect_iso, perfect_input)
                        ] if raw and not d
                    ):
                        skipped += 1
                        continue

                    # --- check for duplicates ---
                    cur = self.conn.execute(
                        "SELECT last_name, first_name FROM employees WHERE employee_id=?",
                        (emp_id,)
                    )
                    existing = cur.fetchone()
                    if existing:
                        msg = (f"Employee ID {emp_id} already exists as "
                               f"{existing[1]} {existing[0]}.\n"
                               f"Replace with {first} {last}?")
                        ans = messagebox.askyesnocancel("Duplicate Employee ID", msg)
                        if ans is None:
                            messagebox.showinfo("Import Cancelled", "Import process stopped by user.")
                            return
                        elif ans:
                            self.conn.execute("""
                                UPDATE employees
                                   SET last_name=?, first_name=?, point_total=?,
                                       last_point_date=?, rolloff_date=?, 
                                       perfect_attendance=?
                                 WHERE employee_id=?;
                            """, (last, first, total, lp_iso, roll_iso, perfect_iso, emp_id))
                            overwritten += 1
                        else:
                            skipped += 1
                            continue
                    else:
                        self.conn.execute("""
                            INSERT INTO employees 
                                (employee_id, last_name, first_name, point_total,
                                 last_point_date, rolloff_date, perfect_attendance)
                            VALUES (?, ?, ?, ?, ?, ?, ?);
                        """, (emp_id, last, first, total, lp_iso, roll_iso, perfect_iso))
                        added += 1

            self.conn.commit()
            self.refresh()
            messagebox.showinfo(
                "Import Complete",
                f"Added: {added}\nOverwritten: {overwritten}\nSkipped: {skipped}"
            )

        except Exception as e:
            messagebox.showerror("Import Failed", str(e))

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
        self.autosize_columns()
    def _on_tree_double_click(self, event):
        """Handle double-click on a row to open the inline edit dialog."""
        item = self.tree.identify('item', event.x, event.y)
        if not item:
            return

        values = self.tree.item(item, "values")
        if not values or len(values) == 0:
            return

        emp_id = int(values[0])
        self._open_inline_edit(emp_id, item, None, None, None)

    def autosize_columns(self):
        """Resize each column to best fit its longest visible text or header."""
        self.update_idletasks()

        # Try to match the Treeview font
        try:
            style = ttk.Style()
            font_name = style.lookup("Treeview", "font")
            if not font_name:
                raise tk.TclError
            tv_font = tkfont.nametofont(font_name)
        except Exception:
            tv_font = tkfont.Font(family="Segoe UI", size=11)

        try:
            header_font_name = ttk.Style().lookup("Treeview.Heading", "font")
            header_font = tkfont.nametofont(header_font_name) if header_font_name else tv_font
        except Exception:
            header_font = tv_font

        headers = {col: self.tree.heading(col)["text"] for col in self.cols}

        for col in self.cols:
            max_width = header_font.measure(headers[col])
            for iid in self.tree.get_children():
                text = str(self.tree.set(iid, col))
                width = tv_font.measure(text)
                if width > max_width:
                    max_width = width
            self.tree.column(col, width=max_width + 10)

        # ---- Stretch proportionally if window wider than content ----
        total_width = sum(self.tree.column(col, "width") for col in self.cols)
        tree_width = self.tree.winfo_width()
        if tree_width > 0 and total_width < tree_width:
            extra = tree_width - total_width
            for col in self.cols:
                current = self.tree.column(col, "width")
                share = (current / total_width) * extra
                self.tree.column(col, width=int(current + share))

    def _on_tree_double_click(self, event):
        """Handle double-click for inline editing."""
        item = self.tree.identify('item', event.x, event.y)
        if not item:
            return

        values = self.tree.item(item, "values")
        if not values or len(values) == 0:
            return

        emp_id = int(values[0])
        # Call the dedicated editor window
        self._open_inline_edit(emp_id, item, None, None, None)

    def _open_inline_edit(self, emp_id, *_):
        """Open a dialog to edit all employee fields for the selected ID (except ID)."""
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

        # --- Window setup ---
        win = tk.Toplevel(self)
        win.title(f"Edit Employee #{emp_id}")
        win.transient(self)
        win.grab_set()
        win.configure(bg="#f6f8fb")
        win.geometry("420x600")

        container = ttk.Frame(win, padding=20)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text=f"Editing Employee #{emp_id}",
                  style="Header.TLabel").pack(anchor="w", pady=(0,10))

        # --- Editable fields (excluding ID) ---
        fields = [
            ("Last Name", "last_name"),
            ("First Name", "first_name"),
            ("Point Total", "point_total"),
            ("Last Point Date (MM-DD-YYYY)", "last_point_date"),
            ("2-Month Rolloff Date (MM-DD-YYYY)", "rolloff_date"),
            ("Perfect Attendance (MM-DD-YYYY)", "perfect_attendance"),
            ("Warning Issued (MM-DD-YYYY)", "point_warning_date"),
        ]
        vars = {}

        for label, key in fields:
            ttk.Label(container, text=label).pack(anchor="w", pady=(4,0))
            if key.endswith("_date"):
                val = ymd_to_us(rec[key]) if rec[key] else ""
            elif key == "point_total":
                val = f"{float(rec[key]):.1f}" if rec[key] is not None else "0.0"
            else:
                val = rec[key] or ""
            vars[key] = tk.StringVar(value=val)
            ttk.Entry(container, textvariable=vars[key], width=36).pack(anchor="w", pady=(0,6))

        # --- Save logic ---
        def save_changes():
            updates = {}
            for key, var in vars.items():
                val = var.get().strip()
                if key == "point_total":
                    try:
                        updates[key] = float(val) if val else 0.0
                    except ValueError:
                        messagebox.showerror("Invalid Value", "Point Total must be a number.")
                        return
                elif key.endswith("_date") and val:
                    iso = parse_us_to_iso(val)
                    if not iso:
                        messagebox.showerror(
                            "Invalid Date",
                            f"{key.replace('_', ' ').title()} must be MM-DD-YYYY or blank."
                        )
                        return
                    updates[key] = iso
                else:
                    updates[key] = val or None

            try:
                self.conn.execute("""
                    UPDATE employees
                       SET last_name=?,
                           first_name=?,
                           point_total=?,
                           last_point_date=?,
                           rolloff_date=?,
                           perfect_attendance=?,
                           point_warning_date=?
                     WHERE employee_id=?;
                """, (updates["last_name"], updates["first_name"], updates["point_total"],
                      updates["last_point_date"], updates["rolloff_date"],
                      updates["perfect_attendance"], updates["point_warning_date"], emp_id))
                self.conn.commit()
                self.app.set_status(f"Updated Employee #{emp_id}", ok=True)
                self.refresh()
                self.refresh_all_cb()
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Could not update employee: {e}")

        # --- Buttons and shortcuts (note indentation) ---
        ttk.Button(container, text="💾 Save Changes", command=save_changes).pack(pady=(12,6))
        ttk.Button(container, text="Cancel", command=win.destroy).pack(pady=(0,8))

        win.bind("<Return>", lambda e: save_changes())
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
        
        messagebox.showinfo(
            "Employee Records Removed",
            f"🗑 {deleted_count} employee record(s) and all associated point history were deleted."
        )
        self.app.set_status(f"Deleted {deleted_count} employee(s).", ok=True)
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

        # Clear inputs and confirm visually
        self.new_id.set("")
        self.new_last.set("")
        self.new_first.set("")
        self.new_perfect.set("")

        self.app.set_status("Saved - Employee added.", ok=True)

        # Friendly confirmation popup
        messagebox.showinfo(
            "Employee Added",
            f"✅ {first} {last} (ID #{emp_id}) was successfully added to the system."
        )

        self.refresh()
        self.refresh_all_cb()

        def save_changes():
            last = last_var.get().strip()
            first = first_var.get().strip()
            if not last or not first:
                messagebox.showerror("Missing Info", "Last and first names are required.")
                return

            try:
                self.conn.execute(
                    "UPDATE employees SET last_name=?, first_name=? WHERE employee_id=?",
                    (last, first, emp_id)
                )
                self.conn.commit()
                self.app.set_status("Saved - Employee updated.", ok=True)
                self.refresh()
                self.refresh_all_cb()
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Could not update employee: {e}")

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

        self.cols = ("employee_id","last_name","first_name","total","last_point_date",
                     "status","warning_date")
        headers = ["ID","Last Name","First Name","Total","Last Point",
                   "Status","Warning Issued"]

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
            self.tree.column(ccc, anchor=("center" if ccc in ("employee_id","total","status") else "w"), stretch=True)

        key_map = {h:k for h,k in zip(headers, self.cols)}
        self.sorter = SortableTree(self.tree, self.cols, key_map)
        self.sorter.bind_headings()

        self.refresh()

    def _rows(self):
        """
        Row shape for the Dashboard:
          (id, last, first, total_points, last_point_date, status, warning_issued_date)
        """
        import sqlite3

        # Detect and cache the warning-issued column name once
        if not hasattr(self, "_warn_col_name"):
            try:
                cols = [r["name"] for r in self.conn.execute("PRAGMA table_info(employees)")]
            except Exception:
                cols = []
            candidates = [
                "warning_issued", "warning_issued_date", "warning_date",
                "warn_issued", "warn_date", "pwd", "point_warning_date"
            ]
            self._warn_col_name = next((c for c in candidates if c in cols), None)
            # Optional: debug print to console
            print("Dashboard: using warning column:", self._warn_col_name or "<none found>")

        # Build SELECT that aliases the chosen column to warn_col
        if self._warn_col_name:
            warn_select = f"{self._warn_col_name} AS warn_col"
        else:
            warn_select = "NULL AS warn_col"  # graceful fallback

        cur = self.conn.execute(f"""
            SELECT employee_id, last_name, first_name,
                   point_total, last_point_date,
                   {warn_select}
              FROM employees
             WHERE is_active = 1
          ORDER BY last_name COLLATE NOCASE, first_name COLLATE NOCASE;
        """)
        rs = cur.fetchall()

        def _fmt_date(val):
            if not val:
                return ""
            s = str(val)
            # Prefer ISO→US conversion if possible; otherwise return as-is
            try:
                return ymd_to_us(s)
            except Exception:
                return s

        out = []
        for r in rs:
            emp_id = r["employee_id"]
            last   = r["last_name"] or ""
            first  = r["first_name"] or ""

            total_f = float(r["point_total"] or 0.0)
            total   = f"{total_f:.1f}"

            last_pt = _fmt_date(r["last_point_date"])

            # Status label (use helper if present; else thresholds inline)
            try:
                status = self._status_for_total(total_f)
            except AttributeError:
                status = ("Termination Risk" if total_f >= 7.0 else
                          "Critical"         if total_f >=  5.0 else
                          "Warning"          if total_f >=  4.0 else
                          "Safe")

            warn_dt = _fmt_date(r["warn_col"])

            out.append((emp_id, last, first, total, last_pt, status, warn_dt))
        return out


    def _status_for_total(self, total: float) -> str:
        """Map total points to a dashboard status label."""
        total = float(total or 0.0)
        if total >= 7.0:
            return "Termination Risk"
        if total >= 5.0:
            return "Critical"
        if total >= 4.0:
            return "Warning"
        return "Safe"

    def refresh(self):
        # Clear existing rows
        for r in self.tree.get_children():
            self.tree.delete(r)

        rows = self._rows()  # (id, last, first, total, last_pt, status, warning_issued)

        # Search by ID / last / first
        q = (self.search_var.get() or "").strip().lower()
        if q:
            rows = [
                r for r in rows
                if q in (r[1] or "").lower()
                or q in (r[2] or "").lower()
                or q == str(r[0]).lower()
            ]

        # Filter by Status (index 5)
        f = self.filter_var.get()
        if f != "All":
            mapping = {
                "Safe": "Safe",
                "Warning": "Warning",
                "Critical": "Critical",
                "Termination Risk": "Termination Risk",
            }
            target = mapping.get(f)
            if target is not None:
                rows = [r for r in rows if r[5] == target]

        # Insert rows
        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=row, tags=(tag,))

        # Auto-resize columns after populating
        self.autosize_columns()

    def autosize_columns(self):
        """Resize each column to best fit its longest visible text or header."""
        self.update_idletasks()

        # Safely get the font actually used by Treeview rows
        try:
            # Try the ttk style lookup first
            style = ttk.Style()
            font_name = style.lookup("Treeview", "font")
            if not font_name:
                raise tk.TclError
            tv_font = tkfont.nametofont(font_name)
        except Exception:
            # Fallback to the main app font if theme doesn’t define one
            tv_font = tkfont.Font(family="Segoe UI", size=11)

        # Also get the header font (headings can differ)
        try:
            header_font_name = ttk.Style().lookup("Treeview.Heading", "font")
            if header_font_name:
                header_font = tkfont.nametofont(header_font_name)
            else:
                header_font = tv_font
        except Exception:
            header_font = tv_font

        # Map column identifiers to their header text
        headers = {col: self.tree.heading(col)["text"] for col in self.cols}

        for col in self.cols:
            # Start width from header text
            max_width = header_font.measure(headers[col])

            # Measure every visible cell in this column
            for iid in self.tree.get_children():
                text = str(self.tree.set(iid, col))
                width = tv_font.measure(text)
                if width > max_width:
                    max_width = width

            # Add a small buffer for padding
            self.tree.column(col, width=max_width + 10)
                    # Add a small buffer for readability
            self.tree.column(col, width=max_width + 10)

        # ---- Optional: stretch proportionally if there's extra space ----
        total_width = sum(self.tree.column(col, "width") for col in self.cols)
        tree_width = self.tree.winfo_width()
        if tree_width > 0 and total_width < tree_width:
            extra = tree_width - total_width
            # distribute extra space across all columns proportionally
            for col in self.cols:
                current = self.tree.column(col, "width")
                # scale extra space by column's relative width
                share = (current / total_width) * extra
                self.tree.column(col, width=int(current + share))

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

        # --- Select Employee (editable + live filter) ------------------------
        ttk.Label(form, text="Select Employee:").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        self.emp_var = tk.StringVar()
        self.emp_box = ttk.Combobox(form, textvariable=self.emp_var, width=40, state="normal")
        self.emp_box.grid(row=0, column=1, sticky="w", padx=8, pady=6)

        self._load_employees_into_combobox()   # populate values + maps

        # --- Live type-to-filter for Employee combobox (trace-based) ----------------
        def _emp_filter(*_):
            typed = (self.emp_var.get() or "").strip().lower()
            base = getattr(self, "_emp_all_values", [])
            if not typed:
                matches = base
            else:
                # contains match so you can type 3–4+ letters anywhere in the name
                matches = [v for v in base if typed in v.lower()]

            # update list without touching the entry text
            self.emp_box.configure(values=matches)

        # Use a variable trace instead of key bindings (works for all edits, incl. Backspace)
        try:
            self.emp_var.trace_remove("write", self._emp_trace_id)     # in case we re-open dialog
        except Exception:
            pass
        self._emp_trace_id = self.emp_var.trace_add("write", _emp_filter)
        
        def _emp_open_on_enter(e):
            self.emp_box.event_generate("<Down>")
            return "break"  # don't let Enter trigger the default button

        self.emp_box.bind("<Return>", _emp_open_on_enter)
        self.emp_box.bind("<KP_Enter>", _emp_open_on_enter)  # numpad Enter


        # Nuke any old KeyPress handlers that might swallow Backspace
        self.emp_box.unbind("<KeyPress>")
        # ---------------------------------------------------------------------------

        # Point Date
        ttk.Label(form, text="Point Date (MM-DD-YYYY):").grid(row=1, column=0, sticky="e", padx=8, pady=6)
        self.date_var = tk.StringVar(value="")
        ttk.Entry(form, textvariable=self.date_var, width=16).grid(row=1, column=1, sticky="w", padx=8, pady=6)

        # Point value
        ttk.Label(form, text="Point:").grid(row=2, column=0, sticky="e", padx=8, pady=6)
        self.point_var = tk.StringVar(value="1.0")
        self.point_box = ttk.Combobox(
            form, textvariable=self.point_var,
            values=["0.5","1.0","1.5"], width=10, state="readonly"
        )
        self.point_box.grid(row=2, column=1, sticky="w", padx=8, pady=6)

        # Reason (dropdown)
        ttk.Label(form, text="Reason:").grid(row=3, column=0, sticky="e", padx=8, pady=6)
        self.reason_var = tk.StringVar()
        self.reason_box = ttk.Combobox(
            form, textvariable=self.reason_var,
            values=get_reason_options(self.conn), width=36, state="normal"
        )
        self.reason_box.grid(row=3, column=1, sticky="w", padx=8, pady=6)
        self.reason_box.configure(postcommand=lambda:
            self.reason_box.configure(values=get_reason_options(self.conn))
        )
        try:
            if getattr(self, "_last_reason", ""):
                self.reason_box.set(self._last_reason)
        except Exception:
            pass

                # Note
        ttk.Label(form, text="Note:").grid(row=4, column=0, sticky="ne", padx=8, pady=6)
        self.note_text = tk.Text(form, width=42, height=5)
        self.note_text.grid(row=4, column=1, sticky="w", padx=8, pady=6)

        # Let Tab/Shift-Tab move focus instead of inserting a tab into the Text widget
        def _text_tab_next(e):
            nxt = e.widget.tk_focusNext()
            if nxt:
                nxt.focus_set()
            return "break"   # prevent a literal \t from being inserted

        def _text_tab_prev(e):
            prv = e.widget.tk_focusPrev()
            if prv:
                prv.focus_set()
            return "break"

        self.note_text.bind("<Tab>", _text_tab_next)
        self.note_text.bind("<Shift-Tab>", _text_tab_prev)
        self.note_text.bind("<ISO_Left_Tab>", _text_tab_prev)  # some systems send this for Shift-Tab

        # Flag Code
        ttk.Label(form, text="Flag Code:").grid(row=5, column=0, sticky="e", padx=8, pady=6)
        self.flag_var = tk.StringVar()
        self.flag_entry = ttk.Entry(form, textvariable=self.flag_var, width=12)
        self.flag_entry.grid(row=5, column=1, sticky="w", padx=8, pady=6)

        # Buttons
        actions = ttk.Frame(self, padding=(6,8,6,6))
        actions.grid(row=2, column=0, sticky="w")
        self.btn_addpoint = ttk.Button(actions, text="Add Point (Ctrl+S)", command=self._add_point)
        self.btn_addpoint.pack(side="left", padx=4)
        ttk.Button(actions, text="Manage Points", command=self._open_manage_points).pack(side="left", padx=4)
        self.undo_btn = ttk.Button(actions, text="Undo (Ctrl+Z)", command=self._undo_point, state="disabled")
        self.undo_btn.pack(side="left", padx=4)

        # ensure the button is tabbable (usually true, but explicit doesn't hurt)
        self.btn_addpoint['takefocus'] = True


    def _load_employees_into_combobox(self):
        """Fill employee combobox and keep full list + display→id map."""
        try:
            cur = self.conn.execute("""
                SELECT employee_id, last_name, first_name
                  FROM employees
                 WHERE is_active = 1
              ORDER BY last_name COLLATE NOCASE, first_name COLLATE NOCASE;
            """)
            rows = cur.fetchall()
        except Exception as e:
            rows = []
            print("WARN: failed to load employees:", e)

        values = [f"{row['last_name']}, {row['first_name']}  (#{row['employee_id']})" for row in rows]
        self._emp_all_values = values
        self._emp_display_to_id = {v: row['employee_id'] for v, row in zip(values, rows)}
        self.emp_box.configure(values=values)

    def _resolve_emp_id(self):
        """Resolve employee ID from combobox selection (exact, (#id), or prefix)."""
        sel = (self.emp_var.get() or "").strip()
        if not sel:
            return None
        mid = getattr(self, "_emp_display_to_id", {}).get(sel)
        if mid:
            return mid
        if "(#" in sel and sel.endswith(")"):
            try:
                return int(sel.split("(#")[1].rstrip(")"))
            except (ValueError, IndexError):
                pass
        base = getattr(self, "_emp_all_values", [])
        t = sel.lower()
        for v in base:
            if v.lower().startswith(t):
                return getattr(self, "_emp_display_to_id", {}).get(v)
        return None

    def _update_undo_button(self):
        self.undo_btn.configure(state="normal" if self.undo_history.has_undo() else "disabled")

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
            if pts not in (0.5, 1.0, 1.5):
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid Points", "Point must be 0.5, 1.0, or 1.5.")
            return

        reason = (self.reason_var.get() or "").strip()
        if not reason:
            messagebox.showinfo("Reason Missing", "Please enter or choose a reason.")
            self.reason_box.focus_set()
            return
        self._last_reason = reason

        note = self.note_text.get("1.0", "end").strip()
        flag = (self.flag_var.get() or "").strip()

        # old state
        old_emp = self.conn.execute(
            "SELECT point_total, last_point_date FROM employees WHERE employee_id=?",
            (emp_id,)
        ).fetchone()

        # insert history
        cur = self.conn.execute("""
            INSERT INTO points_history (employee_id, point_date, points, reason, note, flag_code)
            VALUES (?, ?, ?, ?, ?, ?);
        """, (emp_id, iso_date, pts, reason, note, flag))
        point_id = cur.lastrowid

        # update employee
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

        # undo info
        self.undo_history.push("add_point", {
            "point_id": point_id,
            "emp_id": emp_id,
            "points": pts,
            "old_total": old_emp["point_total"] if old_emp else 0.0,
            "old_last_date": old_emp["last_point_date"] if old_emp else None,
        })
        self._update_undo_button()

        # clear inputs
        self.point_var.set("1.0")
        self.reason_var.set("")
        self.note_text.delete("1.0", "end")
        self.flag_var.set("")

        self.app.set_status("Saved - Point added successfully.", ok=True)
        self.refresh_all_cb()

    def _undo_point(self):
        action = self.undo_history.pop()
        if not action:
            messagebox.showinfo("Undo", "No actions to undo.")
            return
        if action["type"] == "add_point":
            data = action["data"]
            point_id = data["point_id"]
            emp_id = data["emp_id"]
            self.conn.execute("DELETE FROM points_history WHERE id=?", (point_id,))
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
            for r in rows:
                tree.insert("", "end", values=(
                    r["id"], ymd_to_us(r["point_date"]), f"{float(r['points']):.1f}",
                    r["reason"] or "", r["note"] or "", r["flag_code"] or ""
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
            point_id = int(tree.item(iid, "values")[0])
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

        # Uniform button style (tweak padding if you like)
        style = ttk.Style(self)
        style.configure("Accent.TButton", padding=(10, 8))

        box = ttk.Frame(self, padding=(10,10,10,10), style="Pane.TFrame")
        box.grid(row=0, column=0, sticky="ew")
        box.columnconfigure(0, weight=1)

        ttk.Label(box, text="Accent", style="Header.TLabel").grid(row=0, column=0, sticky="w", pady=(0,6))

        ttk.Button(box, text="Export 2-Month Rolloffs (Ctrl+E)",
                   style="Accent.TButton",
                   command=self.export_rolloffs).grid(row=1, column=0, sticky="ew", pady=4)

        ttk.Button(box, text="Export Perfect Attendance",
                   style="Accent.TButton",
                   command=self.export_perfect).grid(row=2, column=0, sticky="ew", pady=4)

        ttk.Button(box, text="Export Point History",
                   style="Accent.TButton",
                   command=self.export_point_history).grid(row=3, column=0, sticky="ew", pady=4)

        ttk.Button(box, text="2 Month Rolloff",
                   style="Accent.TButton",
                   command=self.auto_expire_points).grid(row=4, column=0, sticky="ew", pady=4)

        # NEW: keep this in the same container, same style, same grid
        ttk.Button(box, text="Perfect Attendance Report",
                   style="Accent.TButton",
                   command=self.perfect_attendance_report).grid(row=5, column=0, sticky="ew", pady=4)

        ttk.Label(self, text="CSV files are saved in this program's folder.",
                  foreground=TEXT_MUTED).grid(row=1, column=0, sticky="w", pady=(8,0))

    def _default_save_path(self, prefix: str) -> str:
        """Save to program directory."""
        today = date.today().strftime("%Y%m%d")
        fname = f"{prefix}_{today}.csv"
        prog_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(prog_dir, fname)

    def auto_expire_points(self):
        """Automatically apply rolloff deductions using 2-month cadence with a one-time
        'perfect-month skip' after the 3-month milestone from the last point date.
        Generates an HRIS-ready CSV audit (one aggregated row per affected employee)."""

        today_date = date.today()
        today_iso = today_date.isoformat()

        # Find employees whose rolloff date has passed (i.e., due now or earlier)
        expired_recs = self.conn.execute("""
            SELECT employee_id, last_name, first_name,
                   rolloff_date, point_total,
                   NULLIF(last_point_date, '') AS last_point_iso
              FROM employees
             WHERE rolloff_date IS NOT NULL
               AND date(rolloff_date) <= date('now');
        """).fetchall()

        affected = 0
        total_points_removed = 0.0
        log_rows = []

        for rec in expired_recs:
            emp_id        = rec["employee_id"]
            last_name     = rec["last_name"]
            first_name    = rec["first_name"]
            current_total = float(rec["point_total"] or 0.0)
            next_roll_iso = rec["rolloff_date"]
            last_point_iso = rec["last_point_iso"]

            # Parse dates
            next_roll = datetime.strptime(next_roll_iso, "%Y-%m-%d").date()
            if last_point_iso:
                anchor = datetime.strptime(last_point_iso, "%Y-%m-%d").date()
                perfect_date = three_months_then_first(anchor)  # 1st of month after 3 months from last point
            else:
                # No anchor → no perfect-month to skip; just use straight 2-month cadence
                perfect_date = date.min  # ensures no skip condition triggers

            removed = 0

            if current_total > 0:
                # Apply all elapsed roll-offs (aggregate per employee per run)
                while next_roll <= today_date and current_total > 0:
                    current_total = max(0.0, round(current_total - 1.0, 2))
                    removed += 1
                    next_roll = step_next_due(next_roll, perfect_date)
            else:
                # No points to remove, but keep pushing the due date forward so it isn't perpetually overdue
                while next_roll <= today_date:
                    next_roll = step_next_due(next_roll, perfect_date)

            if removed > 0:
                affected += 1
                total_points_removed += removed

                # Update employee totals and the *next* due date
                self.conn.execute("""
                    UPDATE employees
                       SET point_total = ?,
                           rolloff_date = ?
                     WHERE employee_id = ?;
                """, (current_total, next_roll.isoformat(), emp_id))

                # Audit trail: one aggregated negative entry on today's date
                self.conn.execute("""
                    INSERT INTO points_history (employee_id, point_date, points, reason, note, flag_code)
                    VALUES (?, ?, ?, ?, ?, ?);
                """, (
                    emp_id,
                    today_iso,
                    -1.0 * removed,
                    "Auto Rolloff Adjustment",
                    "Automatic 2-Month Point Expiration (perfect-month skip)",
                    "AUTO"
                ))

                # CSV log row (aggregate)
                log_rows.append([
                    emp_id,
                    last_name,
                    first_name,
                    ymd_to_us(today_iso),     # Rolloff Date in MM-DD-YYYY
                    f"-{removed:.1f}",        # Total deduction applied this run
                    "2 Month Roll Off",
                    "",                       # Note
                    f"{current_total:.1f}"    # Updated total
                ])
            else:
                # No removal; still persist the bumped due date if it moved
                if next_roll_iso != next_roll.isoformat():
                    self.conn.execute("""
                        UPDATE employees
                           SET rolloff_date = ?
                         WHERE employee_id = ?;
                    """, (next_roll.isoformat(), emp_id))

        self.conn.commit()

        # --- Generate CSV audit if any rolloffs were applied ---
        if affected > 0:
            path = self._default_save_path("auto_rolloff_report")
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([
                    "Employee ID",
                    "Last Name",
                    "First Name",
                    "Rolloff Date",
                    "Point",
                    "Reason",
                    "Note",
                    "Point Total"
                ])
                w.writerows(log_rows)

            messagebox.showinfo(
                "Rolloff Applied",
                f"Applied rolloffs to {affected} employee(s) "
                f"(removed {int(total_points_removed)} point(s) total).\n"
                f"Audit saved as {os.path.basename(path)}"
            )
            try:
                self.app.set_status(
                    f"Auto rolloff complete — {affected} employee(s), "
                    f"-{int(total_points_removed)} point(s).",
                    ok=True
                )
            except Exception:
                pass
        else:
            messagebox.showinfo("No Expirations", "No employees have points ready to roll off.")

        # Refresh all tabs
        try:
            self.app._refresh_all()
        except Exception:
            pass

    def export_rolloffs(self):
        """Export employees with upcoming 2-month rolloff dates, formatted for HRIS import."""
        rows = self.conn.execute("""
            SELECT employee_id, last_name, first_name, rolloff_date, point_total
              FROM employees
             WHERE rolloff_date IS NOT NULL AND rolloff_date >= date('now')
          ORDER BY rolloff_date ASC, last_name, first_name;
        """).fetchall()

        path = self._default_save_path("rolloff_report")
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            # HRIS-aligned header order
            w.writerow([
                "Employee ID",
                "Last Name",
                "First Name",
                "Rolloff Date",
                "Note",
                "Reason",
                "Point Total"
            ])

            for r in rows:
                w.writerow([
                    r["employee_id"],
                    r["last_name"],
                    r["first_name"],
                    ymd_to_us(r["rolloff_date"]),
                    "",  # Empty Note column
                    "",  # Empty Reason column
                    f"{float(r['point_total'] or 0):.1f}"
                ])

        self.app.set_status(f"Report exported - {os.path.basename(path)}", ok=True)

    def export_perfect(self):
        """Export employees with future perfect-attendance dates, formatted for HRIS import."""
        rows = self.conn.execute("""
            SELECT employee_id, last_name, first_name, perfect_attendance, point_total
              FROM employees
             WHERE perfect_attendance IS NOT NULL AND perfect_attendance >= date('now')
          ORDER BY perfect_attendance ASC, last_name, first_name;
        """).fetchall()

        path = self._default_save_path("perfect_attendance_report")
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            # HRIS-specific header order
            w.writerow([
                "Employee ID",
                "Last Name",
                "First Name",
                "Perfect Attendance Date",
                "Note",
                "Reason",
                "Point Total"
            ])

            for r in rows:
                w.writerow([
                    r["employee_id"],
                    r["last_name"],
                    r["first_name"],
                    ymd_to_us(r["perfect_attendance"]),
                    "",  # Empty Note column
                    "",  # Empty Reason column
                    f"{float(r['point_total'] or 0):.1f}"
                ])

        self.app.set_status(f"Report exported - {os.path.basename(path)}", ok=True)

    def export_point_history(self):
        rows = self.conn.execute("""
            SELECT p.id, e.employee_id, e.last_name, e.first_name,
                   p.point_date, p.points, p.reason, p.note, p.flag_code
              FROM points_history p
              JOIN employees e ON e.employee_id = p.employee_id
          ORDER BY e.employee_id ASC, p.point_date ASC, p.id ASC;
        """).fetchall()

        path = self._default_save_path("point_history_report")
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Entry ID", "Employee ID", "Last Name", "First Name", "Point Date", "Point", "Reason", "Note", "Flag Code", "Point Total"])

            cumulative = {}  # running total per employee

            for r in rows:
                emp_id = r["employee_id"]
                pts = float(r["points"] or 0)

                # initialize if first time seeing this employee
                if emp_id not in cumulative:
                    # start from zero before first entry
                    cumulative[emp_id] = 0.0

                # increment running total
                cumulative[emp_id] += pts
                point_total = cumulative[emp_id]

                # write CSV row
                w.writerow([
                    r["id"],
                    emp_id,
                    r["last_name"],
                    r["first_name"],
                    ymd_to_us(r["point_date"]),
                    f"{pts:.1f}",
                    r["reason"] or "",
                    r["note"] or "",
                    r["flag_code"] or "",
                    f"{point_total:.1f}"
                ])

        self.app.set_status(f"Report exported - {os.path.basename(path)}", ok=True)
    def perfect_attendance_report(self, as_of=None, dry_run=False):
        """
        Export a CSV of all employees whose perfect_attendance date is due on or before `as_of`
        (default: today). After export, advance each due date to the first day of the month
        after 3 months. When dry_run=True, no DB changes are written.
        """
        as_of_date = as_of or date.today()
        as_of_iso = as_of_date.isoformat()

        rows = self.conn.execute("""
            SELECT employee_id,
                   last_name,
                   first_name,
                   COALESCE(point_total, 0.0) AS point_total,
                   NULLIF(perfect_attendance, '') AS perfect_iso
              FROM employees
             WHERE perfect_attendance IS NOT NULL
               AND date(perfect_attendance) <= date(?)
             ORDER BY last_name, first_name;
        """, (as_of_iso,)).fetchall()

        if not rows:
            msg = f"No perfect-attendance dates due as of {ymd_to_us(as_of_iso)}."
            if dry_run:
                messagebox.showinfo("Simulation Only", msg)
            else:
                messagebox.showinfo("No Perfect Attendance Due", msg)
            return

        log_rows = []
        updated = 0

        for rec in rows:
            emp_id       = rec["employee_id"]
            last_name    = rec["last_name"]
            first_name   = rec["first_name"]
            point_total  = float(rec["point_total"] or 0.0)
            perfect_iso  = rec["perfect_iso"]

            # Parse current due date and compute next due = +3 months, first of next
            due_d   = datetime.strptime(perfect_iso, "%Y-%m-%d").date()
            next_d  = three_months_then_first(due_d)

            # CSV row (use US format for dates)
            log_rows.append([
                emp_id,
                last_name,
                first_name,
                ymd_to_us(perfect_iso),          # Current due (being reported now)
                ymd_to_us(next_d.isoformat()),   # Next scheduled perfect-attendance date
                f"{point_total:.1f}",
            ])

            if not dry_run:
                self.conn.execute("""
                    UPDATE employees
                       SET perfect_attendance = ?
                     WHERE employee_id = ?;
                """, (next_d.isoformat(), emp_id))
                updated += 1

        if not dry_run:
            self.conn.commit()

        # Write/export report (always produce the file on real run)
        if dry_run:
            messagebox.showinfo(
                "Simulation Only",
                f"As of {ymd_to_us(as_of_iso)}, {len(log_rows)} employee(s) would be on the Perfect Attendance report.\n"
                "No changes were written."
            )
        else:
            path = self._default_save_path("perfect_attendance_report")
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([
                    "Employee ID",
                    "Last Name",
                    "First Name",
                    "Perfect Attendance Date",
                    "Next Perfect Attendance Date",
                    "Point Total",
                ])
                w.writerows(log_rows)

            messagebox.showinfo(
                "Perfect Attendance Report",
                f"Exported {len(log_rows)} employee(s). "
                f"Advanced {updated} perfect-attendance date(s).\n"
                f"Saved as {os.path.basename(path)}"
            )

        # Optional status bar & refresh
        try:
            if dry_run:
                self.app.set_status(
                    f"Perfect Attendance SIM: {len(log_rows)} due as of {ymd_to_us(as_of_iso)}.",
                    ok=True
                )
            else:
                self.app.set_status(
                    f"Perfect Attendance report saved — {len(log_rows)} employees; "
                    f"{updated} date(s) advanced.",
                    ok=True
                )
            self.app._refresh_all()
        except Exception:
            pass

# ----------------------------
# Main Application
# ----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Tracker - ATP Beta7")
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
            print("Active ttk theme:", style.theme_use())
            print("Available themes:", style.theme_names())

        except Exception:
            pass

        # --- Modern Cool Theme ---
        BASE_FONT = ("Segoe UI Variable", 10)
        HEADER_FONT = ("Segoe UI Semibold", 13)

        # Base UI
        style.configure(".", font=BASE_FONT, foreground="#2e3640")
        style.configure("TFrame", background="#f6f8fb")
        style.configure("Pane.TFrame", background="#ffffff", borderwidth=1, relief="solid")
        style.configure("Header.TLabel", font=HEADER_FONT, foreground="#3c4a60", background="#f6f8fb")
        style.configure("TLabel", background="#f6f8fb", foreground="#2e3640")

        # Entry Fields
        style.configure("TEntry",
                        fieldbackground="#ffffff",
                        bordercolor="#d0d7e2",
                        lightcolor="#ffffff",
                        darkcolor="#c7cfdb",
                        padding=6)
        style.map("TEntry",
                  fieldbackground=[("focus", "#fdfefe")],
                  bordercolor=[("focus", "#7ea6f7")])

        # Treeview (data tables)
        style.configure("Treeview",
                        background="#ffffff",
                        fieldbackground="#ffffff",
                        bordercolor="#d3dae6",
                        rowheight=32)
        style.configure("Treeview.Heading",
                        font=("Segoe UI Semibold", 10),
                        foreground="#1e2a36",
                        background="#eef2f7")
        style.map("Treeview",
                  background=[("selected", "#d9e5ff")],
                  foreground=[("selected", "#1e2a36")])

        # Modern Buttons
        style.configure("TButton",
                        font=("Segoe UI Semibold", 10),
                        background="#eaf0fb",
                        foreground="#1e2a36",
                        padding=(10, 6),
                        borderwidth=0,
                        relief="flat")
        style.map("TButton",
                  background=[("active", "#d6e4ff"), ("pressed", "#bcd2ff")],
                  relief=[("pressed", "flat")])

        # Accent & Danger Buttons
        style.configure("Accent.TButton",
                        background="#4c6faf", foreground="white",
                        font=("Segoe UI Semibold", 10))
        style.map("Accent.TButton",
                  background=[("active", "#3f5d94")],
                  relief=[("pressed", "flat")])

        style.configure("Danger.TButton",
                        background="#c94b4b", foreground="white",
                        font=("Segoe UI Semibold", 10))
        style.map("Danger.TButton",
                  background=[("active", "#a53e3e")])

        # Tree stripes
        self.strip_even = "#f8f9fc"
        self.strip_odd  = "#eef3f9"

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
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            img = Image.open(logo_path)
            img = img.resize((200, 200), Image.LANCZOS)  # adjust as needed
            self.logo_image = ImageTk.PhotoImage(img)
            ttk.Label(logo_box, image=self.logo_image).pack(padx=12, pady=(10,2))
        except Exception as e:
            # Fallback text if logo not found
            ttk.Label(logo_box, text="ATP", font=("Segoe UI Semibold", 28)).pack(padx=12, pady=(10,2))
            print(f"⚠ Could not load logo: {e}")

        # Subtitle beneath logo
        ttk.Label(logo_box, text="Point System", style="Header.TLabel").pack(padx=12, pady=(0,10))
        # Right notebook
        right = ttk.Frame(root_frame, padding=6, style="TFrame")
        right.grid(row=0, column=1, sticky="nsew")

        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True)

        # Adjust window size to fit contents
        self.update_idletasks()
        self.geometry(f"{self.winfo_reqwidth()}x{self.winfo_reqheight()}")


        self.tab_dashboard = DashboardFrame(nb, self.conn, self._refresh_all, self)
        self.tab_employees = EmployeesFrame(nb, self.conn, self._refresh_all, self)
        self.tab_addpoints = AddPointsFrame(nb, self.conn, self._refresh_all, self)
        self.tab_reports   = ReportsFrame(nb, self.conn, self)

        nb.add(self.tab_dashboard, text="Dashboard")
        nb.add(self.tab_employees, text="Employees")
        nb.add(self.tab_addpoints, text="Add Points")
        nb.add(self.tab_reports, text="Reports")

        self._refresh_all()
        self.update_idletasks()
        self.geometry(f"{self.winfo_reqwidth()}x{self.winfo_reqheight()}")
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