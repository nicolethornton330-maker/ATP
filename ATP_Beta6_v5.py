#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ATP_Beta6_v5.py
- Builds on v4 and PRESERVES: Add Points, Manage Points, Reports, MM-DD-YYYY UI, gray/blue theme
- NEW: Employees tab now REQUIRES manual Employee ID (numeric, unique) when adding a new employee
"""

import os
import csv
import sqlite3
from datetime import datetime, date
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

DB_PATH      = "attendance_MASTER.db"

# ----------------------------
# Date helpers (US display: MM-DD-YYYY)
# ----------------------------
US_DATE_FMT = "%m-%d-%Y"

def ymd_to_us(iso_val) -> str:
    """Render any ISO date (YYYY-MM-DD) — or a date/datetime — as MM-DD-YYYY.
       Tolerates a few legacy formats so the UI never shows raw strings."""
    if not iso_val:
        return ""
    try:
        if isinstance(iso_val, date):
            d = iso_val
        elif isinstance(iso_val, datetime):
            d = iso_val.date()
        else:
            s = str(iso_val).strip()
            # Strict ISO first
            try:
                d = datetime.strptime(s, "%Y-%m-%d").date()
            except ValueError:
                # Be forgiving for any legacy values that may exist
                d = None
                for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d"):
                    try:
                        d = datetime.strptime(s, fmt).date()
                        break
                    except ValueError:
                        continue
                if d is None:
                    return s  # give up gracefully
        return d.strftime(US_DATE_FMT)
    except Exception:
        return str(iso_val)

def parse_us_to_iso(s: str):
    """Parse user input in MM-DD-YYYY (preferred) or MM/DD/YYYY → ISO YYYY-MM-DD."""
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
    # Days in month
    if m in (1,3,5,7,8,10,12):
        dim = 31
    elif m in (4,6,9,11):
        dim = 30
    else:
        # February, check leap year
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
# DB schema bootstrap
# ----------------------------
SCHEMA_BASE = """
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS employees (
    employee_id INTEGER PRIMARY KEY,   -- now explicitly set by user (manual, unique, numeric)
    last_name TEXT NOT NULL,
    first_name TEXT NOT NULL,
    point_total REAL DEFAULT 0,
    last_point_date TEXT,          -- ISO YYYY-MM-DD
    rolloff_date TEXT,             -- ISO YYYY-MM-DD
    perfect_attendance TEXT,       -- ISO YYYY-MM-DD (next perfect attendance date)
    point_warning_date TEXT        -- ISO YYYY-MM-DD (optional policy use)
);

CREATE TABLE IF NOT EXISTS points_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id INTEGER NOT NULL,
    point_date TEXT NOT NULL,      -- ISO YYYY-MM-DD
    points REAL NOT NULL,          -- 0.5 or 1.0
    reason TEXT,
    note TEXT,
    flag_code TEXT,
    FOREIGN KEY(employee_id) REFERENCES employees(employee_id)
);

CREATE INDEX IF NOT EXISTS idx_emp_name ON employees(last_name, first_name);
CREATE INDEX IF NOT EXISTS idx_points_emp ON points_history(employee_id);
CREATE INDEX IF NOT EXISTS idx_points_date ON points_history(point_date);
"""

def safe_connect_db(path=DB_PATH):
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn

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
            # Try number, then date, then string
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
# Employees Tab
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

        frame = ttk.Frame(self, padding=(6,6,6,6), style="Pane.TFrame")
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=22)
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

        self.tree.tag_configure("even", background=self.app.strip_even)
        self.tree.tag_configure("odd", background=self.app.strip_odd)

        self.sorter = SortableTree(self.tree, cols)
        self.sorter.bind_headings()

        # ---- Add Employee form (row 3) ----
        form = ttk.Frame(self, padding=(6,8,6,6))
        form.grid(row=3, column=0, sticky="ew")
        ttk.Label(form, text="Add New Employee", style="Header.TLabel").grid(row=0, column=0, columnspan=6, sticky="w", pady=(0,6))

        ttk.Label(form, text="Employee ID:").grid(row=1, column=0, sticky="e", padx=4)
        self.new_id = tk.StringVar()
        ttk.Entry(form, textvariable=self.new_id, width=12).grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(form, text="Last Name:").grid(row=1, column=2, sticky="e", padx=4)
        self.new_last = tk.StringVar()
        ttk.Entry(form, textvariable=self.new_last, width=22).grid(row=1, column=3, sticky="w", padx=4)

        ttk.Label(form, text="First Name:").grid(row=1, column=4, sticky="e", padx=4)
        self.new_first = tk.StringVar()
        ttk.Entry(form, textvariable=self.new_first, width=22).grid(row=1, column=5, sticky="w", padx=4)

        ttk.Button(form, text="➕ Add Employee", command=self._add_employee).grid(row=1, column=6, sticky="w", padx=8)

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
                   e.point_warning_date
              FROM employees e
          ORDER BY e.last_name, e.first_name;
        """)
        rows = []
        for rec in cur.fetchall():
            emp_id   = rec["employee_id"]
            ln       = rec["last_name"] or ""
            fn       = rec["first_name"] or ""
            total    = rec["point_total"] or 0
            lpd      = ymd_to_us(rec["last_point_date"])
            rd       = ymd_to_us(rec["rolloff_date"])
            pb       = ymd_to_us(rec["perfect_attendance"])
            pwd      = ymd_to_us(rec["point_warning_date"])
            rows.append((emp_id, ln, fn, f"{float(total):.1f}", lpd, rd, pb, pwd))
        return rows

    def refresh(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        q = (self.search_var.get() or "").strip().lower()
        rows = self._rows()
        if q:
            rows = [r for r in rows if q in (r[1] or "").lower() or q in (r[2] or "").lower() or q == str(r[0])]
        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=row, tags=(tag,))

    def _add_employee(self):
        emp_id_raw = (self.new_id.get() or "").strip()
        last = (self.new_last.get() or "").strip()
        first = (self.new_first.get() or "").strip()

        # Validate presence
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

        # Check uniqueness
        exists = self.conn.execute("SELECT 1 FROM employees WHERE employee_id=?", (emp_id,)).fetchone()
        if exists:
            messagebox.showerror("Duplicate ID", "Employee ID already exists. Please enter a unique ID.")
            return

        try:
            self.conn.execute("""
                INSERT INTO employees (employee_id, last_name, first_name, point_total, last_point_date, rolloff_date, perfect_attendance, point_warning_date)
                VALUES (?, ?, ?, 0.0, NULL, NULL, NULL, NULL);
            """, (emp_id, last, first))
            self.conn.commit()
        except sqlite3.IntegrityError as e:
            messagebox.showerror("Database Error", f"Could not add employee: {e}")
            return

        # Clear inputs
        self.new_id.set("")
        self.new_last.set("")
        self.new_first.set("")

        self.app.set_status("Saved ✓ Employee added.", ok=True)
        self.refresh()
        self.refresh_all_cb()

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

        # Controls
        top = ttk.Frame(self, padding=(6,6,6,6))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)

        left = ttk.Frame(top); left.grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="Show:").pack(side="left")
        self.filter_var = tk.StringVar(value="All")
        self.filter_box = ttk.Combobox(left, textvariable=self.filter_var,
                                       values=["All","Safe","Warning","Critical","Termination"],
                                       width=22, state="readonly")
        self.filter_box.pack(side="left", padx=6)
        ttk.Button(left, text="🔄 Refresh", command=self.refresh).pack(side="left", padx=6)

        right = ttk.Frame(top); right.grid(row=0, column=1, sticky="e")
        ttk.Label(right, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(right, textvariable=self.search_var, width=28)
        self.search_entry.pack(side="left", padx=6)
        self.search_entry.bind("<KeyRelease>", lambda e: self.refresh())

        self.cols = ("employee_id","last_name","first_name","total",
                     "last_point","rolloff_date","perfect_bonus","status","warning_date")
        headers = ["ID","Last Name","First Name","Total Points","Last Point",
                   "2-Month Rolloff","Perfect Attendance","Status","Point Warning Date"]

        frame = ttk.Frame(self, padding=(6,6,6,6), style="Pane.TFrame")
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame, columns=self.cols, show="headings", height=22)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        for ccc, h in zip(self.cols, headers):
            self.tree.heading(ccc, text=h)
            w = 150
            if ccc in ("last_name","first_name"): w = 180
            if ccc == "status": w = 230
            if ccc == "total": w = 130
            if ccc == "employee_id": w = 100
            self.tree.column(ccc, width=w, anchor=("center" if ccc in ("employee_id","total","status") else "w"))

        self.tree.tag_configure("even", background=self.app.strip_even)
        self.tree.tag_configure("odd", background=self.app.strip_odd)

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
                   e.last_point_date,
                   e.rolloff_date,
                   e.perfect_attendance,
                   e.point_warning_date
              FROM employees e
          ORDER BY e.last_name, e.first_name;
        """)
        rows = []
        for emp_id, ln, fn, total, lpd, rd, pb, pwd in cur.fetchall():
            if not total or float(total) == 0:
                status = "✅ Safe"
            elif 5 <= float(total) <= 6:
                status = "⚠️ Warning"
            elif float(total) >= 8.0:
                status = "🚫 TERMINATION LEVEL"
            elif float(total) > 6:
                status = "🚫 Critical"
            else:
                status = ""
            rows.append((
                emp_id,
                ln or "",
                fn or "",
                f"{float(total or 0):.1f}",
                ymd_to_us(lpd),
                ymd_to_us(rd),
                ymd_to_us(pb),
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
                                   or q == str(r[0])]

        f = self.filter_var.get()
        if f != "All":
            mapping = {
                "Safe": "✅ Safe",
                "Warning": "⚠️ Warning",
                "Critical": "🚫 Critical",
                "Termination": "🚫 TERMINATION LEVEL",
            }
            target = mapping.get(f, None)
            if target:
                rows = [r for r in rows if r[7] == target]

        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=row, tags=(tag,))

# ----------------------------
# Add Points Tab
# ----------------------------
class AddPointsFrame(ttk.Frame):
    def __init__(self, parent, conn, refresh_all_cb, app):
        super().__init__(parent, padding=10)
        self.conn = conn
        self.app = app
        self.refresh_all_cb = refresh_all_cb

        self.columnconfigure(0, weight=1)

        # Top form
        form = ttk.Frame(self, padding=(6,6,6,6), style="Pane.TFrame")
        form.grid(row=0, column=0, sticky="ew")
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
        actions.grid(row=1, column=0, sticky="w")
        ttk.Button(actions, text="Add Point", command=self._add_point).pack(side="left", padx=4)
        ttk.Button(actions, text="Manage Points", command=self._open_manage_points).pack(side="left", padx=4)

        # Spacer
        ttk.Frame(self, height=8).grid(row=2, column=0)

    def _load_employees_into_combobox(self):
        cur = self.conn.execute("SELECT employee_id, last_name, first_name FROM employees ORDER BY last_name, first_name;")
        self.emps = cur.fetchall()
        display = [f"{row['last_name']}, {row['first_name']}  (#{row['employee_id']})" for row in self.emps]
        self.emp_box["values"] = display

    def _resolve_emp_id(self):
        sel = (self.emp_var.get() or "").strip()
        if not sel:
            return None
        # Match by id inside parentheses, else by name
        if sel.endswith(")") and "(#" in sel:
            try:
                eid = int(sel.split("(#")[-1].strip(") "))
                return eid
            except Exception:
                pass
        # Fallback: find by name
        for row in self.emps:
            label = f"{row['last_name']}, {row['first_name']}  (#{row['employee_id']})"
            if label == sel:
                return int(row["employee_id"])
        return None

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

        # Insert into points_history
        self.conn.execute("""
            INSERT INTO points_history (employee_id, point_date, points, reason, note, flag_code)
            VALUES (?, ?, ?, ?, ?, ?);
        """, (emp_id, iso_date, pts, reason, note, flag))

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

        # Clear small inputs (keep employee selected for fast entry)
        self.point_var.set("1.0")
        self.reason_var.set("")
        self.note_text.delete("1.0", "end")
        self.flag_var.set("")

        self.app.set_status("Saved ✓ Point added successfully.", ok=True)
        self.refresh_all_cb()

    # ---------- Manage Points (delete) ----------
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
                # No remaining points: clear date fields
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
            # Confirm delete
            if not messagebox.askyesno("Confirm Delete", "Permanently delete this point entry?"):
                return
            self.conn.execute("DELETE FROM points_history WHERE id=?", (point_id,))
            self.conn.commit()
            recompute_employee_after_change()
            load_history()
            self.app.set_status("Deleted ✓ Point entry removed.", ok=True)

        btns = ttk.Frame(win, padding=(10,4,10,10))
        btns.grid(row=2, column=0, sticky="w")
        ttk.Button(btns, text="🗑️ Delete Selected", command=delete_selected).pack(side="left", padx=6)
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="left", padx=6)

        load_history()

# ----------------------------
# Reports Tab (preserved from v4)
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

        # Vertically stacked buttons
        ttk.Button(box, text="📄 Export 2-Month Rolloffs", command=self.export_rolloffs).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Button(box, text="🏆 Export Perfect Attendance", command=self.export_perfect).grid(row=2, column=0, sticky="w", pady=4)
        ttk.Button(box, text="🗂️ Export Point History", command=self.export_point_history).grid(row=3, column=0, sticky="w", pady=4)

        ttk.Label(self, text="CSV files are saved in this program’s folder.", foreground=TEXT_MUTED).grid(row=1, column=0, sticky="w", pady=(8,0))

    def _default_save_path(self, prefix: str) -> str:
        today = date.today().strftime("%Y%m%d")
        fname = f"{prefix}_{today}.csv"
        return os.path.join(os.getcwd(), fname)

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
        self.app.set_status(f"Report exported ✓  {os.path.basename(path)}", ok=True)

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
        self.app.set_status(f"Report exported ✓  {os.path.basename(path)}", ok=True)

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
        self.app.set_status(f"Report exported ✓  {os.path.basename(path)}", ok=True)

# ----------------------------
# Main Application
# ----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Tracker - ATP Beta6 v5")
        self.geometry("1150x780")
        self.configure(bg=BG_MAIN)

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

        self.strip_even = STRIPE_EVEN
        self.strip_odd  = STRIPE_ODD

        # DB connection
        self.conn = safe_connect_db(DB_PATH)

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
        ttk.Label(logo_box, text="🗂️", font=("Segoe UI Emoji", 28)).pack(padx=12, pady=(10,2))
        ttk.Label(logo_box, text="ATP", style="Header.TLabel").pack(padx=12, pady=(0,10))

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
        self.status_var.set(message)
        self.status_label.configure(foreground=(GREEN_OK if ok else TEXT_MUTED))
        # Auto-fade back after a delay
        self.after(3500, lambda: self.status_label.configure(foreground=TEXT_MUTED))

    def _refresh_all(self):
        self.tab_dashboard.refresh()
        self.tab_employees.refresh()
        # reload combos in Add Points in case of new employee added
        try:
            self.tab_addpoints._load_employees_into_combobox()
        except Exception:
            pass

    def _on_close(self):
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
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.executescript(SCHEMA_BASE)
        conn.commit()
    finally:
        conn.close()
    if fresh:
        print("✅ New attendance_MASTER.db created (schema initialized).")
    else:
        print("ℹ️  Existing attendance_MASTER.db detected — schema ensured.")

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
