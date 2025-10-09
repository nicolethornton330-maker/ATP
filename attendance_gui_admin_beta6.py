#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Attendance Points Tracker — Admin HR Beta 6 (Patch 2 Final)
- Sidebar back to light background with dark readable text (hover tint #dfe2e8).
- Two-zone sidebar layout (top nav + bottom logo/caption) so it never collapses.
- Auto-scaling logo at bottom-left with "ATP" caption, ~20px above bottom edge.
- Silent autosave UX: no pop-up confirmations for routine edits/adds; instead a small green
  ✓ Saved label appears bottom-right for ~2 seconds after successful commits.
- All Beta 6 features preserved: inline editing, manual override, search/filters, column DnD,
  import, rolloffs (2M + YTD), dual Excel/CSV export, delete with backup, zoom, etc.
"""
import os, platform, shutil, sqlite3
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont
from ATP_beta6_v2_patch import apply_light_theme

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception:
    PANDAS_AVAILABLE = False

DB_FILE = "attendance.db"

# ----------------------------
# Date helpers
# ----------------------------
def ymd_to_us(iso: str) -> str:
    if not iso: return ""
    try:
        d = datetime.strptime(iso, "%Y-%m-%d").date()
        return d.strftime("%m/%d/%Y")
    except Exception:
        return str(iso)

def parse_us_to_iso(s: str):
    s = (s or "").strip()
    if not s: return None
    try:
        return datetime.strptime(s, "%m/%d/%Y").date().isoformat()
    except Exception:
        return None

def add_months(d: date, months: int) -> date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    if m in (1,3,5,7,8,10,12):
        dim = 31
    elif m in (4,6,9,11):
        dim = 30
    else:
        leap = (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0))
        dim = 29 if leap else 28
    day = min(d.day, dim)
    return date(y, m, day)

def first_of_month(d: date) -> date:
    return date(d.year, d.month, 1)

def first_of_next_month(d: date) -> date:
    return add_months(first_of_month(d), 1)

def today() -> date:
    return date.today()

def open_file(path: str):
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass

# ----------------------------
# DB schema
# ----------------------------
SCHEMA_BASE = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS employees (
    employee_id              INTEGER PRIMARY KEY,
    first_name               TEXT NOT NULL,
    last_name                TEXT NOT NULL,
    last_point_date          DATE,
    perfect_bonus            DATE,
    rolloff_date             DATE,
    last_rolloff2m_applied   DATE,
    point_warning_date       DATE
);
CREATE TABLE IF NOT EXISTS points (
    point_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id  INTEGER NOT NULL,
    date         DATE NOT NULL,
    value        REAL NOT NULL,  -- positive=infraction; negative=rolloff/adjustment
    reason       TEXT,
    note         TEXT,
    flag         TEXT,           -- '2M', 'YTD', 'ADJ'
    FOREIGN KEY(employee_id) REFERENCES employees(employee_id) ON DELETE CASCADE
);
"""

def conn():
    return sqlite3.connect(DB_FILE)

def init_db():
    with conn() as c:
        c.executescript(SCHEMA_BASE)

# ----------------------------
# Core helpers
# ----------------------------
def last_positive_point_date(c, emp_id):
    row = c.execute("SELECT MAX(date) FROM points WHERE employee_id=? AND value>0;", (emp_id,)).fetchone()
    return row[0] if row and row[0] else None

def current_total(c, emp_id) -> float:
    row = c.execute("SELECT COALESCE(ROUND(SUM(value),2),0.0) FROM points WHERE employee_id=?;", (emp_id,)).fetchone()
    return float(row[0] or 0.0)

def employee_exists(c, emp_id) -> bool:
    return c.execute("SELECT 1 FROM employees WHERE employee_id=?;", (emp_id,)).fetchone() is not None

def recalc_emp_dates(c, emp_id):
    lpp = last_positive_point_date(c, emp_id)
    if not lpp:
        return
    d = datetime.strptime(lpp, "%Y-%m-%d").date()
    pb = first_of_next_month(add_months(d, 3))
    rd = first_of_next_month(add_months(d, 2))
    t = today()
    while pb <= t:
        pb = add_months(pb, 3)
    while rd <= t:
        rd = add_months(rd, 2)
    c.execute("UPDATE employees SET last_point_date=?, perfect_bonus=?, rolloff_date=? WHERE employee_id=?;",
              (d.isoformat(), pb.isoformat(), rd.isoformat(), emp_id))

def recalc_all(c):
    for (emp_id,) in c.execute("SELECT employee_id FROM employees;").fetchall():
        recalc_emp_dates(c, emp_id)

# ----------------------------
# Import from Excel
# ----------------------------
def import_from_excel(c, emp_path, pts_path):
    if not PANDAS_AVAILABLE:
        raise RuntimeError("Install pandas/openpyxl: python -m pip install pandas openpyxl")

    emp_df = pd.read_excel(emp_path, dtype={"Employee #": "Int64"})
    for col in ["Employee #", "Last Name", "First Name"]:
        if col not in emp_df.columns:
            raise ValueError(f"Employees file missing required column: {col}")

    emp_added = 0
    for _, r in emp_df.iterrows():
        emp_id = r.get("Employee #")
        if pd.isna(emp_id):
            continue
        emp_id = int(emp_id)
        first = str(r.get("First Name", "") or "").strip()
        last  = str(r.get("Last Name", "") or "").strip()
        lpd_iso = None
        try:
            lpd = r.get("Last Point Date", None)
            if pd.notna(lpd):
                lpd_iso = pd.to_datetime(lpd, errors="coerce").date().isoformat()
        except Exception:
            lpd_iso = None

        if not employee_exists(c, emp_id):
            c.execute("""INSERT INTO employees (employee_id, first_name, last_name, last_point_date)
                         VALUES (?,?,?,?);""", (emp_id, first, last, lpd_iso))
            t = today()
            pb = first_of_next_month(add_months(t, 3)).isoformat()
            rd = first_of_next_month(add_months(t, 2)).isoformat()
            c.execute("UPDATE employees SET perfect_bonus=?, rolloff_date=? WHERE employee_id=?;", (pb, rd, emp_id))
            emp_added += 1
        else:
            c.execute("UPDATE employees SET first_name=?, last_name=? WHERE employee_id=?;", (first, last, emp_id))

    pts_df = pd.read_excel(pts_path, dtype={"EmployeeNumber": "Int64"})
    for col in ["EmployeeNumber", "PointedDate", "PointedAmount"]:
        if col not in pts_df.columns:
            raise ValueError(f"Points file missing required column: {col}")

    pts_added, pts_skipped = 0, 0
    for _, r in pts_df.iterrows():
        emp_id = r.get("EmployeeNumber")
        if pd.isna(emp_id):
            pts_skipped += 1
            continue
        emp_id = int(emp_id)

        if not employee_exists(c, emp_id):
            c.execute("INSERT INTO employees (employee_id, first_name, last_name) VALUES (?,?,?);",
                      (emp_id, "Employee", str(emp_id)))
            t = today()
            pb = first_of_next_month(add_months(t, 3)).isoformat()
            rd = first_of_next_month(add_months(t, 2)).isoformat()
            c.execute("UPDATE employees SET perfect_bonus=?, rolloff_date=? WHERE employee_id=?;", (pb, rd, emp_id))

        try:
            ds = r.get("PointedDate", None)
            ds_iso = pd.to_datetime(ds, errors="coerce").date().isoformat() if pd.notna(ds) else None
        except Exception:
            ds_iso = None
        if not ds_iso:
            pts_skipped += 1
            continue

        try:
            val = float(r.get("PointedAmount", None))
            if val not in (0.5, 1.0):
                pts_skipped += 1
                continue
        except Exception:
            pts_skipped += 1
            continue

        reason = r.get("Reason", None)
        note   = r.get("Note", None)
        reason = None if (reason is None or (isinstance(reason,float) and str(reason)=="nan")) else str(reason)
        note   = None if (note   is None or (isinstance(note,  float) and str(note)  =="nan")) else str(note)

        c.execute("""INSERT INTO points (employee_id, date, value, reason, note)
                     VALUES (?,?,?,?,?);""", (emp_id, ds_iso, val, reason, note))
        c.execute("UPDATE employees SET last_point_date=? WHERE employee_id=?;", (ds_iso, emp_id))
        recalc_emp_dates(c, emp_id)
        pts_added += 1

    recalc_all(c)
    return emp_added, pts_added, pts_skipped

# ----------------------------
# Rolloffs
# ----------------------------
def months_with_positive_points(c, emp_id):
    rows = c.execute("""
        SELECT DISTINCT CAST(strftime('%Y', date) AS INT), CAST(strftime('%m', date) AS INT)
        FROM points WHERE employee_id=? AND value>0;
    """, (emp_id,)).fetchall()
    return {(y,m) for (y,m) in rows}

def ensure_no_duplicate_rolloff_entry(c, emp_id, iso_date, flag):
    return c.execute("SELECT 1 FROM points WHERE employee_id=? AND date=? AND flag=?;",
                     (emp_id, iso_date, flag)).fetchone() is not None

def last_rolloff2m_applied(c, emp_id):
    row = c.execute("SELECT last_rolloff2m_applied FROM employees WHERE employee_id=?;", (emp_id,)).fetchone()
    return row[0] if row and row[0] else None

def set_last_rolloff2m_applied(c, emp_id, iso_date):
    c.execute("UPDATE employees SET last_rolloff2m_applied=? WHERE employee_id=?;", (iso_date, emp_id))

def apply_2m_rolloff(c):
    t = today()
    results = []
    emps = c.execute("SELECT employee_id FROM employees;").fetchall()
    for (emp_id,) in emps:
        pos = months_with_positive_points(c, emp_id)
        lpp_iso = last_positive_point_date(c, emp_id)
        if not lpp_iso:
            continue
        lpp = datetime.strptime(lpp_iso, "%Y-%m-%d").date()
        baseline_iso = last_rolloff2m_applied(c, emp_id)
        baseline = datetime.strptime(baseline_iso, "%Y-%m-%d").date() if baseline_iso else lpp
        current_m = first_of_next_month(baseline)
        consecutive = 0
        credit_dates = []
        limit = first_of_month(t)
        while current_m < limit:
            has_pts = (current_m.year, current_m.month) in pos
            if has_pts:
                consecutive = 0
            else:
                consecutive += 1
                if consecutive == 2:
                    ev = first_of_next_month(current_m)
                    if ev <= t:
                        credit_dates.append(ev)
                    consecutive = 0
            current_m = first_of_next_month(current_m)
        total_rolled = 0.0
        for ev in credit_dates:
            ev_iso = ev.isoformat()
            if ensure_no_duplicate_rolloff_entry(c, emp_id, ev_iso, "2M"):
                continue
            total_now = current_total(c, emp_id)
            if total_now <= 0:
                break
            amount = 1.0 if total_now >= 1.0 else total_now
            c.execute("""
                INSERT INTO points (employee_id, date, value, reason, note, flag)
                VALUES (?, ?, ?, '2-Month Rolloff', NULL, '2M');
            """, (emp_id, ev_iso, -amount))
            total_rolled += amount
            set_last_rolloff2m_applied(c, emp_id, ev_iso)
        if total_rolled > 0:
            results.append((emp_id, round(total_rolled,2)))
    recalc_all(c)
    return results

def apply_ytd_rolloff(c):
    t = today()
    this_first = first_of_month(t)
    prev_year = t.year - 1
    month = t.month
    results = []
    emps = c.execute("SELECT employee_id FROM employees;").fetchall()
    for (emp_id,) in emps:
        row = c.execute("""
            SELECT COALESCE(SUM(value),0.0) FROM points
             WHERE employee_id=? AND value>0
               AND CAST(strftime('%Y', date) AS INT)=?
               AND CAST(strftime('%m', date) AS INT)=?;
        """, (emp_id, prev_year, month)).fetchone()
        amt = float(row[0] or 0.0)
        if amt <= 0:
            continue
        ev_iso = this_first.isoformat()
        if ensure_no_duplicate_rolloff_entry(c, emp_id, ev_iso, "YTD"):
            continue
        total_now = current_total(c, emp_id)
        drop = min(amt, total_now)
        if drop <= 0:
            continue
        c.execute("""
            INSERT INTO points (employee_id, date, value, reason, note, flag)
            VALUES (?, ?, ?, 'YTD Rolloff', ?, 'YTD');
        """, (emp_id, ev_iso, -drop, f"Rolled month {month}/{prev_year}"))
        results.append((emp_id, round(drop,2)))
    recalc_all(c)
    return results

# ----------------------------
# Backup
# ----------------------------
def backup_db():
    if not os.path.exists(DB_FILE):
        return None
    os.makedirs("backups", exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    dest = os.path.join("backups", f"attendance_backup_{ts}.db")
    shutil.copy2(DB_FILE, dest)
    return dest

# ----------------------------
# Sorting / Dragging utilities
# ----------------------------
def parse_for_sort(col_key, value_str):
    if col_key in ("total",):
        try: return float(value_str)
        except: return -1e18
    if col_key in ("last_point","rolloff_date","perfect_bonus","warning_date"):
        if not value_str: return date.min
        try:
            return datetime.strptime(value_str, "%m/%d/%Y").date()
        except:
            return date.min
    return (value_str or "").lower()

class SortableTree:
    def __init__(self, tree, columns, key_map):
        self.tree = tree
        self.columns = tuple(columns)   # tuple of keys
        self.key_map = dict(key_map)    # heading text -> key
        self.sort_state = {}            # key -> bool (True asc, False desc)

    def bind_headings(self, on_drag_start=None, on_drag_motion=None, on_drag_end=None):
        for idx, key in enumerate(self.columns):
            text = self.tree.heading(f"#{idx+1}")["text"]
            self.tree.heading(f"#{idx+1}", command=lambda t=text: self.sort_by_heading(t))
        if on_drag_start:
            self.tree.bind("<ButtonPress-1>", on_drag_start, add="+")
        if on_drag_motion:
            self.tree.bind("<B1-Motion>", on_drag_motion, add="+")
        if on_drag_end:
            self.tree.bind("<ButtonRelease-1>", on_drag_end, add="+")

    def sort_by_heading(self, heading_text):
        key = self.key_map.get(heading_text)
        if not key: return
        items = [(iid, self.tree.item(iid, "values")) for iid in self.tree.get_children("")]
        col_index = self.columns.index(key)
        decorated = []
        for iid, vals in items:
            val_str = vals[col_index]
            decorated.append((parse_for_sort(key, val_str), iid, vals))
        asc = self.sort_state.get(key, True)
        decorated.sort(key=lambda x: x[0], reverse=not asc)
        for i, (_k, iid, _vals) in enumerate(decorated):
            self.tree.move(iid, "", i)
        self.sort_state[key] = not asc

# ----------------------------
# GUI Root
# ----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Points Tracker — HR Beta 6 (Patch 2 Final)")
        self.geometry("1360x900")
        try: self.state('zoomed')
        except Exception: pass
        try: self.attributes('-zoomed', True)
        except Exception: pass

        init_db()
        self.c = conn()

        # ---- Theme & Fonts ----
        self.sidebar_bg = "#f3f4f8"
        self.sidebar_fg = "#222222"   # dark text for readability
        self.sidebar_hover = "#dfe2e8"
        self.base_font_size = 15  # larger default
        self.base_font = tkfont.nametofont("TkDefaultFont")
        self.base_font.configure(size=self.base_font_size, family="Segoe UI")

        style = ttk.Style(self)
        apply_light_theme(style)  # 🎨 activate modern light mode

        style.configure("Treeview", rowheight=32, font=self.base_font, padding=4)
        style.configure("Treeview.Heading", font=(self.base_font.actual("family"), self.base_font_size, "bold"), padding=(6,4))
        style.map('Treeview', background=[('selected', '#cde1ff')])

        # Row tags for striping
        self.strip_even = "#f5f7fb"
        self.strip_odd  = "#ffffff"

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # ---- Sidebar (custom tk.Frame for color control) ----
        self.nav = tk.Frame(self, bg=self.sidebar_bg, padx=10, pady=10, width=240)
        self.nav.grid(row=0, column=0, sticky="nsw")
        self.nav.grid_propagate(False)  # don't shrink

        # Top and footer zones
        self.nav_top = tk.Frame(self.nav, bg=self.sidebar_bg)
        self.nav_top.pack(side="top", fill="x")
        self.nav_footer = tk.Frame(self.nav, bg=self.sidebar_bg)
        self.nav_footer.pack(side="bottom", fill="x", pady=(0,20))  # ~20px gap from bottom

        # Sidebar title
        title = tk.Label(self.nav_top, text="Menu", fg=self.sidebar_fg, bg=self.sidebar_bg, font=("Segoe UI", 14, "bold"))
        title.pack(pady=(0,8), anchor="w")

        # Button factory with hover
        def nav_btn(text, cmd):
            b = tk.Label(self.nav_top, text=text, bg=self.sidebar_bg, fg=self.sidebar_fg,
                         padx=12, pady=8, anchor="w", font=("Segoe UI", 12))
            b.pack(fill="x", pady=2)
            def on_enter(e): b.config(bg=self.sidebar_hover)
            def on_leave(e): b.config(bg=self.sidebar_bg)
            b.bind("<Enter>", on_enter)
            b.bind("<Leave>", on_leave)
            b.bind("<Button-1>", lambda e: cmd())
            return b

        nav_btn("Dashboard",   self.show_dashboard)
        nav_btn("Employees",   self.show_employees)
        nav_btn("Points",      self.show_points)
        nav_btn("Reports",     self.show_reports)
        nav_btn("Maintenance", self.show_maintenance)

        # Separator look using a thin line
        tk.Frame(self.nav_top, height=1, bg="#d6d9e4").pack(fill="x", pady=8)

        # Zoom controls
        zoom_row = tk.Frame(self.nav_top, bg=self.sidebar_bg)
        zoom_row.pack(fill="x", pady=(2,6))
        tk.Label(zoom_row, text="Zoom", fg=self.sidebar_fg, bg=self.sidebar_bg, font=("Segoe UI", 11)).pack(side="left")
        def mk_zoom(btn_text, delta):
            btn = tk.Label(zoom_row, text=btn_text, bg=self.sidebar_bg, fg=self.sidebar_fg, bd=1, relief="ridge", padx=8, pady=2)
            btn.pack(side="left", padx=4)
            btn.bind("<Button-1>", lambda e: self.adjust_zoom(delta))
            btn.bind("<Enter>", lambda e: btn.config(bg=self.sidebar_hover))
            btn.bind("<Leave>", lambda e: btn.config(bg=self.sidebar_bg))
        mk_zoom("A-", -1)
        mk_zoom("A+", +1)

        # Exit button
        exit_btn = tk.Label(self.nav_top, text="Exit", bg=self.sidebar_bg, fg=self.sidebar_fg,
                            bd=2, relief="raised", padx=12, pady=6, font=("Segoe UI", 12))
        exit_btn.pack(fill="x", pady=(8,2))
        exit_btn.bind("<Button-1>", lambda e: self.destroy())
        exit_btn.bind("<Enter>", lambda e: exit_btn.config(bg=self.sidebar_hover))
        exit_btn.bind("<Leave>", lambda e: exit_btn.config(bg=self.sidebar_bg))

        # Footer: Logo + APT caption
        self.logo_img = None
        try:
            if os.path.exists("logo.png"):
                self.logo_img = tk.PhotoImage(file="logo.png")
                # Scale down if too wide
                w = self.logo_img.width()
                if w > 200:
                    factor = max(1, int(w/200))
                    self.logo_img = self.logo_img.subsample(factor, factor)
                tk.Label(self.nav_footer, image=self.logo_img, bg=self.sidebar_bg).pack(side="left", anchor="s", padx=(0,6))
        except Exception:
            self.logo_img = None
        tk.Label(self.nav_footer, text="ATP", fg=self.sidebar_fg, bg=self.sidebar_bg, font=("Segoe UI", 12, "bold")).pack(side="left", anchor="s", pady=(8,0))

        # ---- Main content (ttk for native look) ----
        self.content = ttk.Frame(self, padding=8)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.rowconfigure(0, weight=1)
        self.content.columnconfigure(0, weight=1)

        # ---- Global status "✓ Saved" label (hidden by default) ----
        self.status_label = tk.Label(self, text="   ✓ Saved   ", bg="#2e7d32", fg="white", font=("Segoe UI", 10, "bold"))
        self.status_label_visible = False

        self.current = None
        self.show_dashboard()

    def flash_saved(self, duration_ms: int = 2000):
        try:
            # bottom-right, inside root window
            self.status_label.place(relx=1.0, rely=1.0, x=-16, y=-16, anchor="se")
            if hasattr(self, "_status_after_id"):
                self.after_cancel(self._status_after_id)
            self._status_after_id = self.after(duration_ms, lambda: self.status_label.place_forget())
        except Exception:
            pass

    def adjust_zoom(self, delta):
        self.base_font_size = max(10, min(20, self.base_font_size + delta))
        self.base_font.configure(size=self.base_font_size)
        ttk.Style(self).configure("Treeview", rowheight=int(24 + (self.base_font_size-10)*2))

    def swap(self, frame_cls):
        if self.current is not None:
            self.current.destroy()
        self.current = frame_cls(self.content, self.c, self.refresh_dashboard, self)
        self.current.grid(row=0, column=0, sticky="nsew")

    def refresh_dashboard(self):
        if isinstance(self.current, DashboardFrame):
            self.current.refresh()

    def show_dashboard(self):   self.swap(DashboardFrame)
    def show_employees(self):   self.swap(EmployeesFrame)
    def show_points(self):      self.swap(PointsFrame)
    def show_reports(self):     self.swap(ReportsFrame)
    def show_maintenance(self): self.swap(MaintenanceFrame)

# ----------------------------
# Dashboard
# ----------------------------
class DashboardFrame(ttk.Frame):
    def __init__(self, parent, c, _refresh_cb, app: App):
        super().__init__(parent)
        self.c = c
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self); top.grid(row=0, column=0, sticky="ew", pady=(0,8))
        top.columnconfigure(0, weight=1)
        left = ttk.Frame(top); left.grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="Show:").pack(side="left")
        self.filter_var = tk.StringVar(value="All")
        self.filter_box = ttk.Combobox(left, textvariable=self.filter_var, values=["All","Safe","Warning","Critical","Termination"],
                                       width=22, state="readonly")
        self.filter_box.pack(side="left", padx=6)
        ttk.Button(left, text="🔄 Refresh", command=self.refresh).pack(side="left", padx=6)

        right = ttk.Frame(top); right.grid(row=0, column=1, sticky="e")
        ttk.Label(right, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(right, textvariable=self.search_var, width=28)
        self.search_entry.pack(side="left", padx=6)
        self.search_entry.bind("<KeyRelease>", lambda e: self.refresh())

        self.cols = ("employee_id","last_name","first_name","total","last_point","rolloff_date","perfect_bonus","status","warning_date")
        headers = ["ID","Last Name","First Name","Total Points","Last Point","2-Month Rolloff","Perfect Attendance","Status","Point Warning Date"]
        self.tree = ttk.Treeview(self, columns=self.cols, show="headings", height=22)
        for ccc, h in zip(self.cols, headers):
            self.tree.heading(ccc, text=h)
            w = 150
            if ccc in ("last_name","first_name"): w = 180
            if ccc == "status": w = 230
            if ccc == "total": w = 130
            if ccc == "employee_id": w = 100
            self.tree.column(ccc, width=w, anchor="center")
        self.tree.grid(row=2, column=0, sticky="nsew")

        # striping
        self.tree.tag_configure("even", background=self.app.strip_even)
        self.tree.tag_configure("odd", background=self.app.strip_odd)

        key_map = {h:k for h,k in zip(headers, self.cols)}
        self.sorter = SortableTree(self.tree, self.cols, key_map)
        self._drag_from_col = None
        self.sorter.bind_headings(self._on_drag_start, None, self._on_drag_end)

        self.refresh()

    def _rows(self):
        cur = self.c.execute("""
            SELECT e.employee_id, e.last_name, e.first_name,
                   COALESCE(ROUND(SUM(p.value),2),0.0) AS total_points,
                   e.last_point_date, e.rolloff_date, e.perfect_bonus, e.point_warning_date
              FROM employees e
              LEFT JOIN points p ON p.employee_id = e.employee_id
          GROUP BY e.employee_id, e.last_name, e.first_name, e.last_point_date, e.perfect_bonus, e.rolloff_date, e.point_warning_date
          ORDER BY e.last_name, e.first_name;
        """)
        rows = []
        for emp_id, ln, fn, total, lpd, rd, pb, pwd in cur.fetchall():
            if total == 0:
                status = "✅ Safe"
            elif 5 <= total <= 6:
                status = "⚠️ Warning"
            elif total >= 8.0:
                status = "🚫 TERMINATION LEVEL"
            elif total > 6:
                status = "🚫 Critical"
            else:
                status = ""
            rows.append((emp_id, ln, fn, f"{total:.1f}", ymd_to_us(lpd), ymd_to_us(rd), ymd_to_us(pb), status, ymd_to_us(pwd)))
        return rows

    def refresh(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        rows = self._rows()
        q = (self.search_var.get() or "").strip().lower()
        if q:
            rows = [r for r in rows if q in (r[1] or "").lower() or q in (r[2] or "").lower() or q == str(r[0])]
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

    def _heading_at_x(self, x):
        for i in range(len(self.cols)):
            bbox = self.tree.bbox("#"+str(i+1))
            if not bbox:
                continue
            x0,y0,w,h = bbox
            if x0 <= x <= x0+w:
                return i
        return None

    def _on_drag_start(self, event):
        if self.tree.identify_region(event.x, event.y) != "heading":
            return
        idx = self._heading_at_x(event.x)
        if idx is None: return
        self._drag_from_col = idx

    def _on_drag_end(self, event):
        if self._drag_from_col is None:
            return
        if self.tree.identify_region(event.x, event.y) != "heading":
            self._drag_from_col = None; return
        to_idx = self._heading_at_x(event.x)
        if to_idx is None or to_idx == self._drag_from_col:
            self._drag_from_col = None; return
        current = list(self.tree["displaycolumns"])
        def dc_to_keys(dc):
            keys = []
            for item in dc:
                if isinstance(item, str) and item.startswith("#"):
                    keys.append(self.cols[int(item[1:])-1])
                else:
                    keys.append(item)
            return keys
        keys = dc_to_keys(current)
        key = keys.pop(self._drag_from_col)
        keys.insert(to_idx, key)
        self.tree["displaycolumns"] = keys
        self._drag_from_col = None

# ----------------------------
# Employees (inline edit + add + delete + drag headers + unified header)
# ----------------------------
class EmployeesFrame(ttk.Frame):
    def __init__(self, parent, c, dashboard_refresh_cb, app: App):
        super().__init__(parent)
        self.c = c
        self.dashboard_refresh_cb = dashboard_refresh_cb
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # Unified header
        top = ttk.Frame(self); top.grid(row=0, column=0, sticky="ew", pady=(0,8))
        top.columnconfigure(0, weight=1)
        left = ttk.Frame(top); left.grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(left, textvariable=self.search_var, width=28)
        self.search_entry.pack(side="left", padx=6)
        self.search_entry.bind("<KeyRelease>", lambda e: self.refresh())
        ttk.Button(left, text="Add Employee", command=self.add_employee_dialog).pack(side="left", padx=8)
        ttk.Button(left, text="Delete Selected", command=self.delete_employee).pack(side="left", padx=6)
        ttk.Button(left, text="Refresh", command=self.refresh).pack(side="left", padx=6)

        self.cols = ("employee_id","last_name","first_name","last_point_date","perfect_bonus","rolloff_date","manual_total_override","point_warning_date")
        headers = ["ID","Last Name","First Name","Last Point Date","Perfect Attendance Bonus","2 Month Rolloff Date","Total Points (Manual Override)","Point Warning Date"]
        self.tree = ttk.Treeview(self, columns=self.cols, show="headings", height=20)
        for ccc, h in zip(self.cols, headers):
            self.tree.heading(ccc, text=h)
            w = 210
            if ccc in ("employee_id",): w = 110
            if ccc in ("last_name","first_name"): w = 180
            if ccc == "manual_total_override": w = 210
            self.tree.column(ccc, width=w, anchor="center")
        self.tree.grid(row=1, column=0, sticky="nsew")

        self.tree.tag_configure("even", background=self.app.strip_even)
        self.tree.tag_configure("odd", background=self.app.strip_odd)

        # Inline edit
        self.tree.bind("<Double-1>", self.begin_edit)
        self.tree.bind("<Return>", self.begin_edit_from_key)

        # Sort + drag
        key_map = {h:k for h,k in zip(headers, self.cols)}
        self.sorter = SortableTree(self.tree, self.cols, key_map)
        self.sorter.bind_headings(self._on_drag_start, None, self._on_drag_end)
        self._drag_from_col = None

        self.refresh()
        self.editor = None
        self.edit_column = None
        self.edit_item = None

    def delete_employee(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select an employee to delete."); return
        vals = self.tree.item(sel[0], "values")
        emp_id = int(vals[0]); name = f"{vals[2]} {vals[1]}"
        if not messagebox.askyesno("Confirm Deletion",
                                   f"Are you sure you want to permanently delete {name} (ID {emp_id})?\n\n"
                                   "This will remove all related point records."):
            return
        backup = backup_db()
        if backup:
            messagebox.showinfo("Backup Created", f"Backup saved before deletion:\n{backup}")
        try:
            self.c.execute("DELETE FROM employees WHERE employee_id=?;", (emp_id,))
            self.c.commit()
            self.refresh()
            self.dashboard_refresh_cb()
            messagebox.showinfo("Deleted", f"Employee {name} has been deleted successfully.")
        except Exception as ex:
            messagebox.showerror("Deletion Error", f"Could not delete employee:\n{ex}")

    def add_employee_dialog(self):
        win = tk.Toplevel(self); win.title("Add Employee"); win.grab_set()
        ttk.Label(win, text="Employee ID (numeric)").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        id_var = tk.StringVar(); ttk.Entry(win, textvariable=id_var, width=14).grid(row=0, column=1, padx=6, pady=6)
        ttk.Label(win, text="First Name").grid(row=1, column=0, sticky="w", padx=6)
        fn_var = tk.StringVar(); ttk.Entry(win, textvariable=fn_var, width=24).grid(row=1, column=1, padx=6)
        ttk.Label(win, text="Last Name").grid(row=2, column=0, sticky="w", padx=6)
        ln_var = tk.StringVar(); ttk.Entry(win, textvariable=ln_var, width=24).grid(row=2, column=1, padx=6)

        def do_add():
            s = (id_var.get() or "").strip()
            if not s.isdigit():
                messagebox.showwarning("Validation","Employee ID must be numeric."); return
            emp_id = int(s)
            fn = (fn_var.get() or "").strip()
            ln = (ln_var.get() or "").strip()
            if not fn or not ln:
                messagebox.showwarning("Validation","Please enter both first and last name."); return
            if employee_exists(self.c, emp_id):
                messagebox.showwarning("Exists", f"Employee {emp_id} already exists."); return
            self.c.execute("INSERT INTO employees (employee_id, first_name, last_name) VALUES (?,?,?);", (emp_id, fn, ln))
            t = today()
            pb = first_of_next_month(add_months(t, 3)).isoformat()
            rd = first_of_next_month(add_months(t, 2)).isoformat()
            self.c.execute("UPDATE employees SET perfect_bonus=?, rolloff_date=? WHERE employee_id=?;", (pb, rd, emp_id))
            self.c.commit()
            self.refresh()
            self.dashboard_refresh_cb()
            self.app.flash_saved()
            win.destroy()

        ttk.Button(win, text="Save", command=do_add).grid(row=3, column=0, columnspan=2, pady=10)

    def _on_drag_start(self, event):
        if self.tree.identify_region(event.x, event.y) != "heading":
            return
        # find which column
        for i in range(len(self.cols)):
            bbox = self.tree.bbox("#"+str(i+1))
            if not bbox: continue
            x0,y0,w,h = bbox
            if x0 <= event.x <= x0+w:
                self._drag_from_col = i; break

    def _on_drag_end(self, event):
        if self._drag_from_col is None:
            return
        if self.tree.identify_region(event.x, event.y) != "heading":
            self._drag_from_col = None; return
        # determine target
        to_idx = None
        for i in range(len(self.cols)):
            bbox = self.tree.bbox("#"+str(i+1))
            if not bbox: continue
            x0,y0,w,h = bbox
            if x0 <= event.x <= x0+w:
                to_idx = i; break
        if to_idx is None or to_idx == self._drag_from_col:
            self._drag_from_col = None; return
        current = list(self.tree["displaycolumns"])
        def dc_to_keys(dc):
            keys = []
            for item in dc:
                if isinstance(item, str) and item.startswith("#"):
                    keys.append(self.cols[int(item[1:])-1])
                else:
                    keys.append(item)
            return keys
        keys = dc_to_keys(current)
        key = keys.pop(self._drag_from_col)
        keys.insert(to_idx, key)
        self.tree["displaycolumns"] = keys
        self._drag_from_col = None

    def refresh(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        cur = self.c.execute("""
            SELECT e.employee_id, e.last_name, e.first_name, e.last_point_date, e.perfect_bonus, e.rolloff_date,
                   COALESCE(ROUND(SUM(p.value),2),0.0) AS total_points, e.point_warning_date
              FROM employees e
              LEFT JOIN points p ON p.employee_id = e.employee_id
          GROUP BY e.employee_id, e.last_name, e.first_name, e.last_point_date, e.perfect_bonus, e.rolloff_date, e.point_warning_date
          ORDER BY e.last_name, e.first_name;
        """)
        rows = cur.fetchall()
        q = (self.search_var.get() or "").strip().lower()
        i_vis = 0
        for (emp_id, ln, fn, lpd, pb, rd, total, pwd) in rows:
            if q and (q not in (ln or "").lower() and q not in (fn or "").lower() and q != str(emp_id)):
                continue
            tag = "even" if i_vis % 2 == 0 else "odd"
            self.tree.insert("", "end", values=(emp_id, ln, fn, ymd_to_us(lpd), ymd_to_us(pb), ymd_to_us(rd), f"{total:.1f}", ymd_to_us(pwd)), tags=(tag,))
            i_vis += 1

    # -------- inline editing --------
    def _identify_cell(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return None, None, None
        col_index = int(col_id.strip("#")) - 1
        editable_cols = {"last_name","first_name","last_point_date","perfect_bonus","rolloff_date","manual_total_override","point_warning_date"}
        key = self.cols[col_index]
        if key not in editable_cols:
            return None, None, None
        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return None, None, None
        return row_id, key, bbox

    def begin_edit_from_key(self, event):
        focus = self.tree.focus()
        sel = self.tree.selection()
        if focus or sel:
            row_id = focus or sel[0]
            col_id = "#2"
            bbox = self.tree.bbox(row_id, col_id)
            if bbox:
                self._start_editor(row_id, "last_name", bbox)
        return "break"

    def begin_edit(self, event):
        row_id, key, bbox = self._identify_cell(event)
        if not row_id:
            return
        self._start_editor(row_id, key, bbox)

    def _start_editor(self, row_id, key, bbox):
        x, y, w, h = bbox
        value = self.tree.set(row_id, key)
        self.edit_item = row_id
        self.edit_column = key
        try:
            self.editor.destroy()
        except Exception:
            pass
        self.editor = tk.Entry(self.tree, font=self.app.base_font)
        self.editor.insert(0, value)
        self.editor.select_range(0, tk.END)
        self.editor.focus_set()
        self.editor.place(x=x, y=y, width=w, height=h)
        self.editor.bind("<Return>", self.finish_edit)
        self.editor.bind("<FocusOut>", self.finish_edit)
        self.editor.bind("<Escape>", self.cancel_edit)

    def finish_edit(self, event):
        if not self.editor:
            return
        new_val = self.editor.get().strip()
        row_vals = self.tree.item(self.edit_item, "values")
        emp_id = int(row_vals[0])
        col = self.edit_column
        try:
            if col == "last_name":
                self.c.execute("UPDATE employees SET last_name=? WHERE employee_id=?;", (new_val, emp_id))
            elif col == "first_name":
                self.c.execute("UPDATE employees SET first_name=? WHERE employee_id=?;", (new_val, emp_id))
            elif col in {"last_point_date","perfect_bonus","rolloff_date","point_warning_date"}:
                if new_val == "":
                    self.c.execute(f"UPDATE employees SET {col}=NULL WHERE employee_id=?;", (emp_id,))
                else:
                    iso = parse_us_to_iso(new_val)
                    if not iso:
                        messagebox.showerror("Invalid Date", "Dates must be MM/DD/YYYY.")
                        self.cancel_edit(None)
                        return
                    self.c.execute(f"UPDATE employees SET {col}=? WHERE employee_id=?;", (iso, emp_id))
            elif col == "manual_total_override":
                try:
                    target_total = float(new_val)
                except Exception:
                    messagebox.showerror("Invalid Number", "Enter a numeric total, e.g., 3.5")
                    self.cancel_edit(None)
                    return
                before = current_total(self.c, emp_id)
                diff = round(target_total - before, 2)
                if abs(diff) >= 1e-9:
                    today_iso = today().isoformat()
                    self.c.execute("""
                        INSERT INTO points (employee_id, date, value, reason, note, flag)
                        VALUES (?, ?, ?, 'Manual Override Adjustment', NULL, 'ADJ');
                    """, (emp_id, today_iso, diff))
                    if diff > 0:
                        self.c.execute("UPDATE employees SET last_point_date=? WHERE employee_id=?;", (today_iso, emp_id))
                    recalc_emp_dates(self.c, emp_id)
            self.c.commit()
        except Exception as ex:
            messagebox.showerror("Save Error", f"Could not save edit:\n{ex}")
            self.cancel_edit(None)
            return

        try:
            self.editor.destroy()
        except Exception:
            pass
        self.editor = None; self.edit_item=None; self.edit_column=None
        self.refresh(); self.dashboard_refresh_cb(); self.app.flash_saved()

    def cancel_edit(self, event):
        try:
            self.editor.destroy()
        except Exception:
            pass
        self.editor = None; self.edit_item=None; self.edit_column=None
        self.refresh()

# ----------------------------
# Points (Add infractions)
# ----------------------------
class PointsFrame(ttk.Frame):
    def __init__(self, parent, c, dashboard_refresh_cb, app: App):
        super().__init__(parent)
        self.c = c
        self.dashboard_refresh_cb = dashboard_refresh_cb
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        form = ttk.LabelFrame(self, text="Add Point (Infraction)", padding=8)
        form.grid(row=0, column=0, sticky="ew", pady=(0,8))

        ttk.Label(form, text="Employee").grid(row=0, column=0, sticky="w")
        self.emp_choices = self.load_employee_choices()
        self.emp_var = tk.StringVar()
        self.emp_combo = ttk.Combobox(form, textvariable=self.emp_var, values=[x[1] for x in self.emp_choices], width=32, state="readonly")
        self.emp_combo.grid(row=0, column=1, padx=6)
        self.emp_combo.bind("<<ComboboxSelected>>", self.on_emp_selected)

        ttk.Label(form, text="Employee ID").grid(row=0, column=2, sticky="w")
        self.empid_var = tk.StringVar(); ttk.Entry(form, textvariable=self.empid_var, width=12, state="readonly").grid(row=0, column=3, padx=6)

        ttk.Label(form, text="Date (MM/DD/YYYY)").grid(row=0, column=4, sticky="w")
        self.date_var = tk.StringVar(value=today().strftime("%m/%d/%Y")); ttk.Entry(form, textvariable=self.date_var, width=16).grid(row=0, column=5, padx=6)

        ttk.Label(form, text="Value (0.5 or 1.0)").grid(row=1, column=0, sticky="w")
        self.value_var = tk.StringVar(value="1.0")
        ttk.Combobox(form, textvariable=self.value_var, values=["0.5","1.0"], width=6, state="readonly").grid(row=1, column=1, padx=6)

        ttk.Label(form, text="Reason").grid(row=1, column=2, sticky="w")
        self.reason_var = tk.StringVar(); ttk.Entry(form, textvariable=self.reason_var, width=30).grid(row=1, column=3, padx=6, sticky="w")

        ttk.Label(form, text="Note").grid(row=1, column=4, sticky="w")
        self.note_var = tk.StringVar(); ttk.Entry(form, textvariable=self.note_var, width=30).grid(row=1, column=5, padx=6, sticky="w")

        ttk.Button(form, text="Add Point", command=self.add_point).grid(row=0, column=6, rowspan=2, padx=8)

        tbl = ttk.LabelFrame(self, text="Recent Points (All Employees)", padding=8); tbl.grid(row=1, column=0, sticky="nsew")
        cols = ("point_id","employee_id","employee_name","date","value","reason","note","flag")
        headers = ["Point ID","Employee ID","Employee Name","Date","Value","Reason","Note","Flag"]
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings", height=14)
        for ccc, h in zip(cols, headers):
            self.tree.heading(ccc, text=h)
            self.tree.column(ccc, width=120 if ccc not in ("reason","note","employee_name") else 240, anchor="center")
        self.tree.pack(fill="both", expand=True)

        self.tree.tag_configure("even", background=self.app.strip_even)
        self.tree.tag_configure("odd", background=self.app.strip_odd)

        bottom = ttk.Frame(self); bottom.grid(row=2, column=0, sticky="ew", pady=6)
        ttk.Button(bottom, text="Refresh", command=self.refresh).pack(side="left")
        ttk.Button(bottom, text="Reload Names", command=self.reload_names).pack(side="left", padx=6)
        self.refresh()

    def load_employee_choices(self):
        rows = self.c.execute("SELECT employee_id, last_name, first_name FROM employees ORDER BY last_name, first_name;").fetchall()
        return [(emp_id, f"{ln}, {fn}") for emp_id, ln, fn in rows]

    def reload_names(self):
        self.emp_choices = self.load_employee_choices()
        self.emp_combo["values"] = [x[1] for x in self.emp_choices]

    def on_emp_selected(self, _e):
        name = self.emp_var.get()
        for emp_id, disp in self.emp_choices:
            if disp == name:
                self.empid_var.set(str(emp_id)); return

    def refresh(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        cur = self.c.execute("""
            SELECT p.point_id, p.employee_id, e.last_name || ', ' || e.first_name as employee_name,
                   p.date, p.value, p.reason, p.note, p.flag
              FROM points p JOIN employees e ON e.employee_id = p.employee_id
          ORDER BY p.date DESC, p.point_id DESC LIMIT 200;
        """)
        rows = cur.fetchall()
        for i, (pid, eid, ename, ds, val, reason, note, flag) in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=(pid, eid, ename, ymd_to_us(ds), f"{val:.1f}", reason, note, flag or ""), tags=(tag,))

    def add_point(self):
        name = self.emp_var.get()
        emp_id = None
        for eid, disp in self.emp_choices:
            if disp == name: emp_id = eid; break
        if emp_id is None:
            messagebox.showwarning("Validation", "Please select an employee from the list."); return

        ds_iso = parse_us_to_iso(self.date_var.get())
        if not ds_iso:
            messagebox.showwarning("Validation", "Date must be MM/DD/YYYY."); return
        try:
            val = float(self.value_var.get())
            if val not in (0.5, 1.0):
                raise ValueError
        except Exception:
            messagebox.showwarning("Validation", "Value must be 0.5 or 1.0."); return

        if not employee_exists(self.c, emp_id):
            messagebox.showwarning("Validation", f"Employee {emp_id} does not exist."); return

        before = current_total(self.c, emp_id)
        if before >= 8.0:
            messagebox.showerror("Limit Reached", f"Employee {emp_id} already has {before:.1f} points (termination level).");
            return
        if before + val > 8.0:
            messagebox.showerror("Would Exceed 8.0", f"Employee {emp_id} currently has {before:.1f}.\nAdding {val:.1f} would exceed 8.0.");
            return

        reason = (self.reason_var.get() or "").strip() or None
        note   = (self.note_var.get()   or "").strip() or None
        self.c.execute("""
            INSERT INTO points (employee_id, date, value, reason, note)
            VALUES (?, ?, ?, ?, ?);
        """, (emp_id, ds_iso, val, reason, note))
        self.c.execute("UPDATE employees SET last_point_date=? WHERE employee_id=?;", (ds_iso, emp_id))
        recalc_emp_dates(self.c, emp_id)
        self.c.commit()

        after = current_total(self.c, emp_id)
        if before < 5.0 and 5.0 <= after <= 6.0:
            # Warning prompt retained
            row = self.c.execute("SELECT first_name, last_name FROM employees WHERE employee_id=?;", (emp_id,)).fetchone()
            if row:
                fn, ln = row
                if messagebox.askyesno("Warning Threshold", f"Employee {fn} {ln} (ID {emp_id}) has reached 5 points.\nRecord a Warning Date for today?"):
                    self.c.execute("UPDATE employees SET point_warning_date=? WHERE employee_id=?;", (today().isoformat(), emp_id))
                    self.c.commit()

        self.refresh(); self.dashboard_refresh_cb(); self.app.flash_saved()

# ----------------------------
# Reports (Dual-tab + CSV for detailed points)
# ----------------------------
class ReportsFrame(ttk.Frame):
    def __init__(self, parent, c, _refresh_cb, app: App):
        super().__init__(parent)
        self.c = c
        self.columnconfigure(0, weight=1)

        ttk.Label(self, text="Reports", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, sticky="w", pady=(0,6))
        ttk.Button(self, text="Export Dual-Tab Excel (Summary + Details)", command=self.export_dual).grid(row=1, column=0, sticky="w")
        ttk.Button(self, text="Export Detailed Points (CSV)", command=self.export_details_csv).grid(row=2, column=0, sticky="w", pady=6)

        if not PANDAS_AVAILABLE:
            ttk.Label(self, foreground="red", text="Note: pandas/openpyxl not found.\nInstall:\n  python -m pip install pandas openpyxl").grid(row=3, column=0, sticky="w")

    def fetch_summary(self):
        return self.c.execute("""
            SELECT e.employee_id, e.last_name, e.first_name,
                   COALESCE(ROUND(SUM(p.value),2),0.0) AS total_points,
                   e.last_point_date, e.perfect_bonus, e.rolloff_date
              FROM employees e LEFT JOIN points p ON p.employee_id = e.employee_id
          GROUP BY e.employee_id, e.last_name, e.first_name, e.last_point_date, e.perfect_bonus, e.rolloff_date
          ORDER BY e.last_name, e.first_name;
        """).fetchall()

    def fetch_details(self):
        return self.c.execute("""
            SELECT e.employee_id AS employee_no,
                   e.last_name,
                   e.first_name,
                   p.value AS point,
                   p.date   AS point_date,
                   p.reason,
                   p.note,
                   p.flag AS flag_code
              FROM points p
              JOIN employees e ON e.employee_id = p.employee_id
          ORDER BY p.date DESC, p.point_id DESC;
        """).fetchall()

    def export_dual(self):
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Missing Libraries", "Please install 'pandas' and 'openpyxl' to export Excel.\n\npython -m pip install pandas openpyxl")
            return
        fpath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Workbook","*.xlsx")], title="Save Dual-Tab Excel As")
        if not fpath:
            return
        try:
            summary = self.fetch_summary()
            details = self.fetch_details()

            def fmt_date(v):
                if isinstance(v, str) and len(v) == 10 and '-' in v:
                    return ymd_to_us(v)
                return v

            sum_rows = []
            for r in summary:
                r = list(r)
                r[4] = fmt_date(r[4])  # last_point_date
                r[5] = fmt_date(r[5])  # perfect_bonus
                r[6] = fmt_date(r[6])  # rolloff_date
                sum_rows.append(r)

            det_rows = []
            for r in details:
                r = list(r)
                r[4] = fmt_date(r[4])  # point_date
                det_rows.append(r)

            sum_cols = ["Employee #","Last Name","First Name","Total Points","Most Recent Point","Perfect Bonus","2-Month Rolloff"]
            det_cols = ["Employee #","Last Name","First Name","Point","Point Date","Reason","Note","Flag Code"]

            with pd.ExcelWriter(fpath, engine="openpyxl") as w:
                pd.DataFrame(sum_rows, columns=sum_cols).to_excel(w, sheet_name="Summary", index=False)
                pd.DataFrame(det_rows, columns=det_cols).to_excel(w, sheet_name="Details", index=False)

            open_file(fpath)
            messagebox.showinfo("Export Complete", f"Excel created:\n{fpath}")
        except Exception as ex:
            messagebox.showerror("Export Error", f"Could not export Excel:\n{ex}")

    def export_details_csv(self):
        base = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")], title="Save Detailed Points CSV As")
        if not base:
            return
        try:
            details = self.fetch_details()
            def fmt_date(v):
                if isinstance(v, str) and len(v) == 10 and '-' in v:
                    return ymd_to_us(v)
                return v
            det_rows = []
            for r in details:
                r = list(r); r[4] = fmt_date(r[4]); det_rows.append(r)
            det_cols = ["Employee #","Last Name","First Name","Point","Point Date","Reason","Note","Flag Code"]
            import csv
            with open(base, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(det_cols)
                writer.writerows(det_rows)
            open_file(base)
            messagebox.showinfo("Export Complete", f"CSV created:\n{base}")
        except Exception as ex:
            messagebox.showerror("Export Error", f"Could not export CSV:\n{ex}")

# ----------------------------
# Maintenance
# ----------------------------
class MaintenanceFrame(ttk.Frame):
    def __init__(self, parent, c, dashboard_refresh_cb, app: App):
        super().__init__(parent)
        self.c = c
        self.dashboard_refresh_cb = dashboard_refresh_cb
        self.columnconfigure(0, weight=1)

        ttk.Label(self, text="Maintenance", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, sticky="w", pady=(0,6))
        ttk.Button(self, text="Import Employees & Points from Excel", command=self.do_import).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Button(self, text="2 Month Rolloff", command=self.rolloff_2m).grid(row=2, column=0, sticky="w", pady=4)
        ttk.Button(self, text="YTD Rolloff", command=self.rolloff_ytd).grid(row=3, column=0, sticky="w", pady=4)
        ttk.Button(self, text="Recalculate All Dates", command=self.recalc_all).grid(row=4, column=0, sticky="w", pady=12)

        if not PANDAS_AVAILABLE:
            ttk.Label(self, foreground="red", text="Note: pandas/openpyxl not found.\nInstall:\n  python -m pip install pandas openpyxl").grid(row=5, column=0, sticky="w")

        ttk.Label(self, text=("2 Month: removes 1.0 per completed two clean months (dated 1st of following month).\n"
                              "YTD: removes points from same month last year (dated 1st of this month)."))\
            .grid(row=6, column=0, sticky="w", pady=6)

    def do_import(self):
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Missing Libraries", "Install pandas/openpyxl first:\n\npython -m pip install pandas openpyxl")
            return
        emp_path = filedialog.askopenfilename(title="Select Employees Excel", filetypes=[("Excel files","*.xlsx *.xls")])
        if not emp_path:
            return
        pts_path = filedialog.askopenfilename(title="Select Points Excel", filetypes=[("Excel files","*.xlsx *.xls")])
        if not pts_path:
            return
        try:
            emp_added, pts_added, pts_skipped = import_from_excel(self.c, emp_path, pts_path)
            self.c.commit()
            self.dashboard_refresh_cb()
            messagebox.showinfo("Import Complete", f"Imported {emp_added} employees and {pts_added} points.\nSkipped {pts_skipped} row(s).")
        except Exception as ex:
            messagebox.showerror("Import Error", f"Could not import data:\n{ex}")

    def rolloff_2m(self):
        backup = backup_db()
        if backup:
            messagebox.showinfo("Backup Created", f"Backup saved before rolloff:\n{backup}")
        rows = apply_2m_rolloff(self.c)
        self.c.commit()
        self.dashboard_refresh_cb()
        if not rows:
            messagebox.showinfo("2 Month Rolloff", "No eligible rolloff this run.")
            return
        if PANDAS_AVAILABLE:
            id2name = {eid: (ln + ", " + fn) for eid, ln, fn in self.c.execute("SELECT employee_id, last_name, first_name FROM employees;").fetchall()}
            data = [(eid, id2name.get(eid, ""), pts) for (eid, pts) in rows]
            df = pd.DataFrame(data, columns=["Employee ID","Name","Points Rolled Off"])
            out = f"Rolloff_Report_2M_{today().isoformat()}.xlsx"
            df.to_excel(out, index=False); open_file(out)
            messagebox.showinfo("2 Month Rolloff", f"Applied to {len(rows)} employee(s).\nReport: {out}")
        else:
            messagebox.showinfo("2 Month Rolloff", f"Applied to {len(rows)} employee(s).")

    def rolloff_ytd(self):
        backup = backup_db()
        if backup:
            messagebox.showinfo("Backup Created", f"Backup saved before rolloff:\n{backup}")
        rows = apply_ytd_rolloff(self.c)
        self.c.commit()
        self.dashboard_refresh_cb()
        if not rows:
            messagebox.showinfo("YTD Rolloff", "No eligible rolloff this run.")
            return
        if PANDAS_AVAILABLE:
            id2name = {eid: (ln + ", " + fn) for eid, ln, fn in self.c.execute("SELECT employee_id, last_name, first_name FROM employees;").fetchall()}
            data = [(eid, id2name.get(eid, ""), pts) for (eid, pts) in rows]
            df = pd.DataFrame(data, columns=["Employee ID","Name","Points Rolled Off"])
            out = f"Rolloff_Report_YTD_{today().isoformat()}.xlsx"
            df.to_excel(out, index=False); open_file(out)
            messagebox.showinfo("YTD Rolloff", f"Applied to {len(rows)} employee(s).\nReport: {out}")
        else:
            messagebox.showinfo("YTD Rolloff", f"Applied to {len(rows)} employee(s).")

    def recalc_all(self):
        recalc_all(self.c); self.c.commit(); self.dashboard_refresh_cb()
        messagebox.showinfo("Recalculation Complete", "Recalculated dates for all employees.")

# ----------------------------
# Entry
# ----------------------------
def main():
    init_db()
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