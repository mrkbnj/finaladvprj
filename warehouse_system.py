"""
Warehouse Management System
Copyright (c) 2026 Mark Benjamin H. Acob - All Rights Reserved

Proprietary Software - Internal Use Only
This software is proprietary and confidential.
Unauthorized copying, modification, or distribution is prohibited.

A comprehensive warehouse management system with QR code generation,
item staging, and shelf management capabilities.
Warehouse 1: General IT Equipment
Warehouse 2: Computer Peripherals (Monitor, Keyboard, Mouse, Headset)
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import uuid
import qrcode
import re
import hashlib
from datetime import datetime

import sys

if getattr(sys, 'frozen', False):
    # Running as a PyInstaller EXE
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Running as a normal Python script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(BASE_DIR, "warehouse.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "activity_log.xlsx")
USERS_FILE = os.path.join(BASE_DIR, "users.xlsx")
QR_FOLDER = os.path.join(BASE_DIR, "qr_codes")
QR_FOLDER_W1 = os.path.join(QR_FOLDER, "warehouse_1")
QR_FOLDER_W2 = os.path.join(QR_FOLDER, "warehouse_2")
QR_LABELS_FOLDER = os.path.join(BASE_DIR, "qr_labels")
QR_LABELS_FOLDER_W1 = os.path.join(QR_LABELS_FOLDER, "warehouse_1")
QR_LABELS_FOLDER_W2 = os.path.join(QR_LABELS_FOLDER, "warehouse_2")
EXCEL_FOLDER = os.path.join(BASE_DIR, "excel_exports")
EXCEL_FOLDER_W1 = os.path.join(EXCEL_FOLDER, "warehouse_1")
EXCEL_FOLDER_W2 = os.path.join(EXCEL_FOLDER, "warehouse_2")

SHELVES = [
    "Area A", "Area B", "Area C",
    "Rack 1 - Bay 1", "Rack 1 - Bay 2", "Rack 1 - Bay 3",
    "Rack 2 - Bay 1", "Rack 2 - Bay 2", "Rack 2 - Bay 3",
]
SHELVES_W1 = SHELVES_W2 = SHELVES

EQUIPMENT_TYPES = ["Monitor", "Keyboard", "Mouse", "Headset"]
STATUS_CHOICES  = ["No Issue", "Minimal", "Defective", "Missing"]

# ========== TOOLTIP ==========

class Tooltip:
    """Show a single tooltip at a time when the user hovers over a widget."""
    _active = None  # class-level: only one tooltip visible at a time

    def __init__(self, widget, text):
        self.widget   = widget
        self.text     = text
        self._tip_win = None
        widget.bind("<Enter>",   self._show)
        widget.bind("<Leave>",   self._hide)
        widget.bind("<Destroy>", self._hide)

    def _show(self, event=None):
        # Close any currently open tooltip first
        if Tooltip._active and Tooltip._active is not self:
            Tooltip._active._hide()
        if self._tip_win or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tip_win = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify="left",
                 background="#ffffe0", relief="solid", borderwidth=1,
                 font=("Helvetica", 8), wraplength=220,
                 padx=4, pady=3).pack()
        Tooltip._active = self

    def _hide(self, event=None):
        tw, self._tip_win = self._tip_win, None
        if tw:
            tw.destroy()
        if Tooltip._active is self:
            Tooltip._active = None

def tip(widget, text):
    """Attach a tooltip to widget. Returns the widget for inline use."""
    Tooltip(widget, text)
    return widget

staged_items = []
selected_staged_index = None
staged_sets = []
selected_set_index = None
current_user = ""
current_is_admin = False
session_start = ""

# Track last generated Excel paths per warehouse (for VIEW EXCEL button)
_last_excel_path = {1: None, 2: None}

# ── Checkbox selection state (warehouse table rows) ──────────
# Maps tree iid -> bool; populated/cleared on every table reload.
w1_row_checks: dict = {}   # W1 warehouse table checkboxes
w2_row_checks: dict = {}   # W2 warehouse table checkboxes
w1_pull_row_checks: dict = {}  # W1 pull history table checkboxes
w2_pull_row_checks: dict = {}  # W2 pull history table checkboxes

# ========== INITIALIZATION ==========

def initialize_file():
    # Always ensure all folders exist
    for folder in (QR_FOLDER_W1, QR_FOLDER_W2, QR_LABELS_FOLDER_W1, QR_LABELS_FOLDER_W2,
                   EXCEL_FOLDER_W1, EXCEL_FOLDER_W2):
        os.makedirs(folder, exist_ok=True)

    sheets_to_create = {}
    if not os.path.exists(FILE):
        sheets_to_create = {"items": None, "shelves": None, "pullouts": None,
                            "items_w2": None, "shelves_w2": None, "pullouts_w2": None}
        mode = 'w'
    else:
        try:
            with pd.ExcelFile(FILE) as xls:
                existing = xls.sheet_names
        except Exception as e:
            messagebox.showerror("File Error", f"Could not read '{FILE}':\n{e}")
            return
        needed = ["items", "shelves", "pullouts", "items_w2", "shelves_w2", "pullouts_w2"]
        sheets_to_create = {s: None for s in needed if s not in existing}
        mode = 'a'

    if not sheets_to_create:
        return

    default_dfs = {
        # W1 Sheets
        "items": pd.DataFrame(columns=["QR", "Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]),
        "shelves": pd.DataFrame({"Shelf": SHELVES_W1, "Status": ["AVAILABLE"] * len(SHELVES_W1), "Date_Full": [None] * len(SHELVES_W1)}),
        "pullouts": pd.DataFrame(columns=["Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]),

        # W2 Sheets
        "items_w2": pd.DataFrame(columns=["QR", "Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]),
        "shelves_w2": pd.DataFrame({"Shelf": SHELVES_W2, "Status": ["AVAILABLE"] * len(SHELVES_W2), "Date_Full": [None] * len(SHELVES_W2)}),
        "pullouts_w2": pd.DataFrame(columns=["Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]),
    }
    try:
        with pd.ExcelWriter(FILE, engine='openpyxl', mode=mode) as writer:
            for sheet in sheets_to_create:
                default_dfs[sheet].to_excel(writer, sheet_name=sheet, index=False)
            for ws in writer.book.worksheets:
                ws.protection.sheet = True
                ws.protection.enable()
    except Exception as e:
        messagebox.showerror("File Error", f"Could not create '{FILE}':\n{e}\n\nMake sure the file is not open in Excel.")

def initialize_log():
    if not os.path.exists(LOG_FILE):
        with pd.ExcelWriter(LOG_FILE, engine='openpyxl') as writer:
            pd.DataFrame(columns=["Timestamp", "User", "Action", "Details"]).to_excel(writer, sheet_name="logs", index=False)
            for ws in writer.book.worksheets:
                ws.protection.sheet = True
                ws.protection.enable()

# ========== USERS DATABASE ==========

def _hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()

def _is_admin_password(pw: str) -> bool:
    """Admin passwords must contain at least one of: ! @ #"""
    return bool(re.search(r'[!@#]', pw))

def initialize_users():
    """Create users.xlsx with a default admin account if it doesn't exist."""
    if not os.path.exists(USERS_FILE):
        default_pw = "Admin@123"          # contains @  → admin role
        df = pd.DataFrame([{
            "Username": "admin",
            "Password": _hash_password(default_pw),
            "Role":     "admin",
            "Created":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        with pd.ExcelWriter(USERS_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="users", index=False)
            for ws in writer.book.worksheets:
                ws.protection.sheet = True
                ws.protection.enable()

def load_users() -> pd.DataFrame:
    """Return the users DataFrame (always with expected columns)."""
    initialize_users()
    from openpyxl import load_workbook
    wb = load_workbook(USERS_FILE, data_only=True)
    ws = wb["users"]
    ws.protection.sheet = False
    import tempfile
    tmp = USERS_FILE + ".~tmp.xlsx"
    wb.save(tmp)
    df = pd.read_excel(tmp, sheet_name="users")
    try:
        os.remove(tmp)
    except Exception:
        pass
    for col in ["Username", "Password", "Role", "Created"]:
        if col not in df.columns:
            df[col] = ""
    return df

def _save_users(df: pd.DataFrame):
    with pd.ExcelWriter(USERS_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="users", index=False)
        for ws in writer.book.worksheets:
            ws.protection.sheet = True
            ws.protection.enable()

def authenticate_user(username: str, password: str):
    """Return (ok, role) or (False, None)."""
    df = load_users()
    row = df[df["Username"].str.lower() == username.strip().lower()]
    if row.empty:
        return False, None
    stored_hash = str(row.iloc[0]["Password"])
    if stored_hash != _hash_password(password):
        return False, None
    role = str(row.iloc[0].get("Role", "user")).lower()
    return True, role

def create_account(username: str, password: str) -> str:
    """Create a new account. Returns error string or '' on success."""
    username = username.strip()
    if not username:
        return "Username cannot be empty."
    if len(username) < 3:
        return "Username must be at least 3 characters."
    if not re.match(r'^[A-Za-z0-9_]+$', username):
        return "Username may only contain letters, digits, and underscores."
    if len(password) < 6:
        return "Password must be at least 6 characters."
    df = load_users()
    if username.lower() in df["Username"].str.lower().tolist():
        return f"Username '{username}' already exists."
    role = "admin" if _is_admin_password(password) else "user"
    new_row = pd.DataFrame([{
        "Username": username,
        "Password": _hash_password(password),
        "Role":     role,
        "Created":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }])
    df = pd.concat([df, new_row], ignore_index=True)
    _save_users(df)
    return ""

def delete_account(username: str) -> str:
    """Delete account by username. Returns error string or '' on success."""
    df = load_users()
    match = df[df["Username"].str.lower() == username.strip().lower()]
    if match.empty:
        return f"Account '{username}' not found."
    if str(match.iloc[0]["Role"]).lower() == "admin":
        # prevent deleting the last admin
        remaining_admins = df[df["Role"].str.lower() == "admin"]
        if len(remaining_admins) <= 1:
            return "Cannot delete the only admin account."
    df = df[df["Username"].str.lower() != username.strip().lower()].reset_index(drop=True)
    _save_users(df)
    return ""

def change_password(username: str, new_password: str) -> str:
    """Change password and update role. Returns error string or '' on success."""
    if len(new_password) < 6:
        return "Password must be at least 6 characters."
    df = load_users()
    idx = df[df["Username"].str.lower() == username.strip().lower()].index
    if idx.empty:
        return f"Account '{username}' not found."
    i = idx[0]
    df.at[i, "Password"] = _hash_password(new_password)
    df.at[i, "Role"]     = "admin" if _is_admin_password(new_password) else "user"
    _save_users(df)
    return ""



def _load_sheet(file, sheet, init_fn):
    def _read():
        # Read with openpyxl and strip sheet protection so data loads correctly
        from openpyxl import load_workbook
        wb = load_workbook(file, data_only=True)
        if sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet}' not found")
        ws = wb[sheet]
        ws.protection.sheet = False   # temporarily lift for pandas
        import tempfile, os
        tmp = file + ".~tmp.xlsx"
        wb.save(tmp)
        df = pd.read_excel(tmp, sheet_name=sheet)
        try:
            os.remove(tmp)
        except Exception:
            pass
        return df

    try:
        return _read()
    except Exception:
        try:
            init_fn()
            return _read()
        except Exception as e:
            messagebox.showerror("File Error",
                f"Could not load sheet '{sheet}' from '{file}':\n{e}\n\n"
                "Make sure the file is not open in Excel and the folder is writable.")
            return pd.DataFrame()

def load_items():
    df = _load_sheet(FILE, "items", initialize_file)
    expected = ["QR", "Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
    for col in expected:
        if col not in df.columns:
            df[col] = pd.Series(dtype=str)
    return df
def load_shelves():     return _load_sheet(FILE, "shelves", initialize_file) # W1 Only
def load_shelves_w2():  return _load_sheet(FILE, "shelves_w2", initialize_file) # W2 Only
def load_pullouts():    return _load_sheet(FILE, "pullouts", initialize_file)
def load_items_w2():
    df = _load_sheet(FILE, "items_w2", initialize_file)
    expected = ["QR", "Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
    for col in expected:
        if col not in df.columns:
            df[col] = pd.Series(dtype=str)
    return df
def load_pullouts_w2(): return _load_sheet(FILE, "pullouts_w2", initialize_file)
def load_logs():        return _load_sheet(LOG_FILE, "logs", initialize_log)

def _excel_locked_error():
    messagebox.showerror(
        "File Locked",
        f"Cannot save — '{FILE}' is open in Excel or another program.\n\n"
        "Please close the file and try again."
    )

def _write_all_sheets(df_items, df_shelves, df_pullouts, df_items_w2, df_shelves_w2, df_po2):
    """Single write point for the warehouse Excel file. Raises on failure."""
    try:
        with pd.ExcelWriter(FILE, engine='openpyxl') as writer:
            df_items.to_excel(writer,     sheet_name="items",       index=False)
            df_shelves.to_excel(writer,   sheet_name="shelves",     index=False)
            df_pullouts.to_excel(writer,  sheet_name="pullouts",    index=False)
            df_items_w2.to_excel(writer,  sheet_name="items_w2",    index=False)
            df_shelves_w2.to_excel(writer,sheet_name="shelves_w2",  index=False)
            df_po2.to_excel(writer,       sheet_name="pullouts_w2", index=False)
            for ws in writer.book.worksheets:
                ws.protection.sheet = True
                ws.protection.enable()
    except PermissionError:
        _excel_locked_error()
        raise

def save_warehouse_1(df_items, df_shelves, df_pullouts=None):
    """Saves Warehouse 1 sheets while preserving Warehouse 2."""
    if df_pullouts is None: df_pullouts = load_pullouts()
    _write_all_sheets(df_items, df_shelves, df_pullouts,
                      load_items_w2(), load_shelves_w2(), load_pullouts_w2())

def save_warehouse_2(df_items_w2, df_shelves_w2, df_pullouts_w2=None):
    """Saves Warehouse 2 sheets while preserving Warehouse 1."""
    if df_pullouts_w2 is None: df_pullouts_w2 = load_pullouts_w2()
    _write_all_sheets(load_items(), load_shelves(), load_pullouts(),
                      df_items_w2, df_shelves_w2, df_pullouts_w2)

def save_log(action, details=""):
    initialize_log()
    df_log = load_logs()
    new_row = {"Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "User": current_user, "Action": action, "Details": details}
    df_log = pd.concat([df_log, pd.DataFrame([new_row])], ignore_index=True)
    with pd.ExcelWriter(LOG_FILE, engine='openpyxl') as writer:
        df_log.to_excel(writer, sheet_name="logs", index=False)
        for ws in writer.book.worksheets:
            ws.protection.sheet = True
            ws.protection.enable()

# ========== QR HELPERS ==========

def qr_path_for(hostname, warehouse=1):
    folder = QR_FOLDER_W1 if warehouse == 1 else QR_FOLDER_W2
    return os.path.join(folder, f"{hostname.replace(' ', '_')}.png")

def generate_qr(hostname, data, warehouse=1):
    folder = QR_FOLDER_W1 if warehouse == 1 else QR_FOLDER_W2
    os.makedirs(folder, exist_ok=True)
    qr_img = qrcode.make(data)
    qr_img.save(qr_path_for(hostname, warehouse))

def delete_qr(hostname, warehouse=1):
    path = qr_path_for(hostname, warehouse)
    if os.path.exists(path):
        try:
            os.remove(path)
        except Exception as e:
            messagebox.showwarning("Warning", f"QR file not deleted: {e}")

# ========== MINI CALENDAR PICKER ==========

def pick_date(parent, target_var, title="Select Date"):
    """Pop up a simple month calendar; sets target_var to 'YYYY-MM-DD' on pick."""
    from calendar import monthcalendar
    cal_win = tk.Toplevel(parent)
    cal_win.title(title)
    cal_win.resizable(False, False)
    cal_win.transient(parent)
    cal_win.grab_set()

    today = datetime.today()
    # mutable state inside closure
    state = {"year": today.year, "month": today.month}

    header = tk.Frame(cal_win, bg="#2c3e50")
    header.pack(fill="x")
    prev_btn = tk.Button(header, text="◀", bg="#2c3e50", fg="white", bd=0, font=("Helvetica", 10),
                         command=lambda: _change_month(-1))
    prev_btn.pack(side="left", padx=8, pady=4)
    month_lbl = tk.Label(header, text="", bg="#2c3e50", fg="white", font=("Helvetica", 10, "bold"), width=16)
    month_lbl.pack(side="left", expand=True)
    next_btn = tk.Button(header, text="▶", bg="#2c3e50", fg="white", bd=0, font=("Helvetica", 10),
                         command=lambda: _change_month(1))
    next_btn.pack(side="right", padx=8, pady=4)

    day_names = tk.Frame(cal_win, bg="#dce3f0")
    day_names.pack(fill="x")
    for i, d in enumerate(["Su","Mo","Tu","We","Th","Fr","Sa"]):
        tk.Label(day_names, text=d, width=4, font=("Helvetica", 8, "bold"),
                 bg="#dce3f0", fg="#2c3e50").grid(row=0, column=i, padx=2, pady=3)

    grid_frame = tk.Frame(cal_win, bg="white", padx=4, pady=4)
    grid_frame.pack()

    def _change_month(delta):
        m = state["month"] + delta
        y = state["year"]
        if m > 12: m, y = 1, y + 1
        if m < 1:  m, y = 12, y - 1
        state["month"], state["year"] = m, y
        _draw()

    def _draw():
        for w in grid_frame.winfo_children():
            w.destroy()
        y, m = state["year"], state["month"]
        month_lbl.config(text=datetime(y, m, 1).strftime("%B %Y"))
        weeks = monthcalendar(y, m)
        for r, week in enumerate(weeks):
            for c, day in enumerate(week):
                if day == 0:
                    tk.Label(grid_frame, text="", width=4, bg="white").grid(row=r, column=c, padx=1, pady=1)
                else:
                    is_today = (day == today.day and m == today.month and y == today.year)
                    btn = tk.Button(
                        grid_frame, text=str(day), width=4,
                        font=("Helvetica", 9, "bold" if is_today else "normal"),
                        bg="#2980b9" if is_today else "white",
                        fg="white" if is_today else "#2c3e50",
                        relief="flat", cursor="hand2",
                        command=lambda d=day: _select(d)
                    )
                    btn.grid(row=r, column=c, padx=1, pady=1)

    def _select(day):
        chosen = f"{state['year']}-{state['month']:02d}-{day:02d}"
        target_var.set(chosen)
        cal_win.destroy()

    # "Clear" button
    clear_frame = tk.Frame(cal_win)
    clear_frame.pack(fill="x", pady=(0, 4))
    tk.Button(clear_frame, text="Clear date", fg="gray", bd=0, font=("Helvetica", 8),
              command=lambda: [target_var.set(""), cal_win.destroy()]).pack()

    _draw()
    cal_win.update_idletasks()
    # centre over parent
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    cw, ch = cal_win.winfo_reqwidth(), cal_win.winfo_reqheight()
    cal_win.geometry(f"+{px + (pw - cw)//2}+{py + (ph - ch)//2}")
    cal_win.focus_force()


def _date_picker_widget(parent, var, label_text):
    """Return a frame with a calendar-button + date label. var is a StringVar."""
    frame = tk.Frame(parent)
    tk.Label(frame, text=label_text).pack(side="left", padx=(0, 2))
    date_lbl = tk.Label(frame, textvariable=var, width=11, relief="sunken",
                        bg="white", font=("Helvetica", 9), anchor="w", padx=4)
    date_lbl.pack(side="left", padx=(0, 2))
    tk.Button(frame, text="📅", width=2,
              command=lambda: pick_date(parent.winfo_toplevel(), var)).pack(side="left")
    return frame

# ========== COLUMN SORT ==========

def attach_sort_headers(tree):
    """No-op: sorting removed, headings are display-only."""
    pass


def next_set_id():
    df = load_items_w2()
    # Also include staged sets
    existing_ids = set(df["Set ID"].dropna().tolist()) if "Set ID" in df.columns else set()
    for s in staged_sets:
        existing_ids.add(s["set_id"])
    n = 1
    while True:
        sid = f"SET-{n:03d}"
        if sid not in existing_ids:
            return sid
        n += 1

# ========== W1 STAGING ==========

def update_staged_display():
    staged_listbox.delete(0, tk.END)
    if not staged_items:
        staged_listbox.insert(tk.END, "No staged items")
        return
    for item in staged_items:
        staged_listbox.insert(tk.END, f"{item['Hostname']} → {item['Shelf']} → {item.get('Status', '')}")

def select_staged_item(event):
    global selected_staged_index
    selection = staged_listbox.curselection()
    if not selection:
        return
    index = selection[0]
    selected_staged_index = index
    item = staged_items[index]
    _fill_input_fields(item["Hostname"], item.get("Serial Number", ""), item.get("Checked By", ""), item["Shelf"], item.get("Status", ""), item.get("Remarks", ""))

# ========== W1 INPUT HELPERS ==========

def _fill_input_fields(hostname="", serial="", checked_by="", shelf="", status="", remarks=""):
    hostname_entry.delete(0, tk.END);   hostname_entry.insert(0, hostname)
    serial_entry.delete(0, tk.END);     serial_entry.insert(0, serial)
    checked_by_entry.delete(0, tk.END); checked_by_entry.insert(0, checked_by)
    shelf_var.set(shelf)
    remarks_var.set(status)
    remarks_text_var.set(remarks)

def _clear_input_fields():
    _fill_input_fields()

# ========== W1 CORE ==========

def remove_from_staging():
    global selected_staged_index
    if selected_staged_index is not None:
        index = selected_staged_index
        if index >= len(staged_items):
            messagebox.showerror("Error", "Invalid staged selection")
            selected_staged_index = None
            return
        removed = staged_items.pop(index)
        selected_staged_index = None
        _clear_input_fields()
        messagebox.showinfo("Removed", f"'{removed['Hostname']}' removed from staging")
    else:
        if not staged_items:
            messagebox.showinfo("Info", "No staged items to clear")
            return
        if not messagebox.askyesno("Confirm", f"Clear all {len(staged_items)} staged item(s)?"):
            return
        staged_items.clear()
        selected_staged_index = None
        _clear_input_fields()
        messagebox.showinfo("Cleared", "All staged items cleared")
    update_staged_display()

def put_item():
    hostname   = hostname_entry.get().strip()
    shelf      = shelf_var.get()
    serial     = serial_entry.get().strip()
    checked_by = checked_by_entry.get().strip()
    status     = remarks_var.get()
    remarks    = remarks_text_var.get().strip()

    if not hostname:
        messagebox.showerror("Error", "Please enter a Hostname"); return
    if not serial:
        messagebox.showerror("Error", "Please enter a Serial Number"); return
    if not checked_by:
        messagebox.showerror("Error", "Please enter Checked By"); return
    if not shelf:
        messagebox.showerror("Error", "Please select a Shelf"); return
    if not status:
        messagebox.showerror("Error", "Please select a Status"); return

    df_items = load_items()
    df_shelves = load_shelves()

    if hostname in df_items["Hostname"].values:
        messagebox.showerror("Error", "Hostname already exists in warehouse"); return
    if any(item['Hostname'] == hostname for item in staged_items):
        messagebox.showerror("Error", "Hostname already staged"); return

    if serial in df_items["Serial Number"].astype(str).values:
        match = df_items[df_items["Serial Number"].astype(str) == serial].iloc[0]
        messagebox.showerror("Error",
            f"Serial Number '{serial}' is already assigned to:\n"
            f"Hostname: {match['Hostname']} | Shelf: {match['Shelf']}"); return
    if any(item.get('Serial Number') == serial for item in staged_items):
        match = next(item for item in staged_items if item.get('Serial Number') == serial)
        messagebox.showerror("Error",
            f"Serial Number '{serial}' is already staged under:\n"
            f"Hostname: {match['Hostname']} | Shelf: {match['Shelf']}"); return

    shelf_status = df_shelves[df_shelves["Shelf"] == shelf]["Status"].values
    if len(shelf_status) > 0 and shelf_status[0] == "FULL":
        messagebox.showerror("Error", "Shelf is marked FULL"); return

    staged_items.append({
        "Hostname":      hostname,
        "Serial Number": serial,
        "Checked By":    checked_by,
        "Shelf":         shelf,
        "Status":        status,
        "Remarks":       remarks,
    })
    _clear_input_fields()
    messagebox.showinfo("Staged", f"'{hostname}' added to staging queue")
    update_staged_display()

def put_warehouse():
    if not staged_items:
        messagebox.showerror("Error", "No staged items to put"); return
    if not messagebox.askyesno("Confirm", f"Put {len(staged_items)} item(s) to warehouse?"):
        return

    try:
        df_items = load_items()
        df_shelves = load_shelves()
        for col in ["Serial Number", "Checked By"]:
            if col not in df_items.columns:
                df_items[col] = ""

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for item in staged_items:
            qr_code = str(uuid.uuid4())
            generate_qr(item['Hostname'], item['Hostname'], warehouse=1)
            df_items = pd.concat([df_items, pd.DataFrame([{
                "QR": qr_code,
                "Hostname": item['Hostname'],
                "Serial Number": item.get('Serial Number', ''),
                "Checked By": item.get('Checked By', ''),
                "Shelf": item['Shelf'],
                "Status": item.get('Status', ''),
                "Remarks": item.get('Remarks', ''),
                "Date": now_str
            }])], ignore_index=True)

        save_warehouse_1(df_items, df_shelves)

        count = len(staged_items)
        for item in staged_items:
            save_log("PUT WAREHOUSE", f"[W1] Hostname: {item['Hostname']} | Shelf: {item['Shelf']}")

        staged_items.clear()
        messagebox.showinfo("Success", f"{count} item(s) added to Warehouse 1.\nQR codes generated.\nUse 'GENERATE FILES' to create PDF labels and export to Excel.")
        update_staged_display()
        w1_refresh_all()

    except Exception as e:
        messagebox.showerror("Save Error",
            f"Failed to save to Excel:\n{str(e)}\n\n"
            "Common causes:\n• Excel file is open → close it\n• Wrong folder")
        


def update_item():
    global selected_staged_index
    new_hostname   = hostname_entry.get().strip()
    new_serial     = serial_entry.get().strip()
    new_checked_by = checked_by_entry.get().strip()
    new_shelf      = shelf_var.get()
    new_status     = remarks_var.get()
    new_remarks    = remarks_text_var.get().strip()

    if not new_hostname:
        messagebox.showerror("Error", "Hostname cannot be empty"); return
    if not new_serial:
        messagebox.showerror("Error", "Serial Number cannot be empty"); return
    if not new_checked_by:
        messagebox.showerror("Error", "Checked By cannot be empty"); return
    if not new_shelf:
        messagebox.showerror("Error", "Please select a Shelf"); return
    if not new_status:
        messagebox.showerror("Error", "Please select a Status"); return

    if selected_staged_index is not None:
        index = selected_staged_index
        if index >= len(staged_items):
            messagebox.showerror("Error", "Invalid staged selection")
            selected_staged_index = None
            return
        if any(i != index and item['Hostname'] == new_hostname for i, item in enumerate(staged_items)):
            messagebox.showerror("Error", "Hostname already exists in staging"); return
        if any(i != index and item.get('Serial Number') == new_serial for i, item in enumerate(staged_items)):
            match = next(item for i, item in enumerate(staged_items) if i != index and item.get('Serial Number') == new_serial)
            messagebox.showerror("Error",
                f"Serial Number '{new_serial}' is already staged under:\n"
                f"Hostname: {match['Hostname']}"); return
        df_items = load_items()
        if new_serial in df_items["Serial Number"].astype(str).values:
            existing_idx = df_items[df_items["Serial Number"].astype(str) == new_serial].index[0]
            if existing_idx != index:
                match = df_items.iloc[existing_idx]
                messagebox.showerror("Error",
                    f"Serial Number '{new_serial}' is already in warehouse:\n"
                    f"Hostname: {match['Hostname']} | Shelf: {match['Shelf']}"); return
        staged_items[index].update({
            "Hostname":      new_hostname,
            "Serial Number": new_serial,
            "Checked By":    new_checked_by,
            "Shelf":         new_shelf,
            "Status":        new_status,
            "Remarks":       new_remarks,
        })
        messagebox.showinfo("Updated", "Staged item updated")
        update_staged_display()
        selected_staged_index = None
        return

    messagebox.showerror(
        "Update Not Allowed",
        "Items cannot be updated directly from the warehouse.\n\n"
        "Double-click the row to move it back to staging,\n"
        "then select it in the staged list and click UPDATE.")

def delete_item():
    selected = tree_warehouse.selection()
    if not selected:
        messagebox.showerror("Error", "Select item"); return

    df_items = load_items()
    df_shelves = load_shelves()
    # values: ☐(0), QR(1), Hostname(2), Serial(3), Checked By(4), Shelf(5), Status(6), Remarks(7), Date(8)
    hostname = tree_warehouse.item(selected[0], "values")[2]

    if not messagebox.askyesno("Confirm Delete", f"Delete '{hostname}'?\nThis cannot be undone."):
        return

    delete_qr(hostname, warehouse=1)
    df_items = df_items[df_items["Hostname"] != hostname].reset_index(drop=True)
    save_warehouse_1(df_items, df_shelves)
    save_log("DELETE ITEM", f"[W1] Hostname: {hostname}")
    messagebox.showinfo("Deleted", "Record and QR code deleted")
    w1_refresh_all()

def pull_search_live(event=None):
    """Filter whichever W1 view is currently active based on the search box."""
    keyword = pull_item_entry.get().strip().lower()

    if tree_available.winfo_ismapped():
        # Shelf status view is active
        df = load_shelves().sort_values("Shelf")
        df_items_all = load_items()
        if keyword:
            mask = False
            for col in ["Shelf", "Status", "Date_Full"]:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        tree_available.delete(*tree_available.get_children())
        for _, row in df.iterrows():
            date_full  = row.get("Date_Full", "")
            shelf_name = row["Shelf"]
            item_count = int((df_items_all["Shelf"] == shelf_name).sum()) if "Shelf" in df_items_all.columns else 0
            tree_available.insert("", "end", values=(shelf_name, row["Status"], item_count, date_full if pd.notna(date_full) else ""))
        w1_search_label.config(text=f"{len(df)} match(es)" if keyword else "")

    elif tree_pullouts.winfo_ismapped():
        # Pull history view is active — respect all active filters
        df = load_pullouts()
        shelf_filter   = pull_shelf_var.get()
        remarks_filter = pull_remarks_var.get()
        date_from      = w1_date_from_var.get().strip()
        date_to        = w1_date_to_var.get().strip()
        if keyword:
            search_cols = ["Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
            mask = False
            for col in search_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        if shelf_filter:   df = df[df["Shelf"] == shelf_filter]
        if remarks_filter: df = df[df["Status"] == remarks_filter]
        df = _filter_by_date(df, date_from, date_to)
        tree_pullouts.delete(*tree_pullouts.get_children())
        w1_pull_row_checks.clear()
        for _, row in df.iterrows():
            iid = tree_pullouts.insert("", "end", values=(
                "☐",
                *tuple(row.get(c, "") for c in ["Hostname", "Serial Number", "Shelf", "Status", "Remarks", "Pull Reason", "Date"])
            ))
            w1_pull_row_checks[iid] = False
        active = bool(keyword or shelf_filter or remarks_filter or date_from or date_to)
        w1_search_label.config(text=f"{len(df)} match(es)" if active else "", fg="darkorange" if active else "blue")

    else:
        # Warehouse view (default)
        show_warehouse()
        if keyword:
            df = load_items()
            search_cols = ["QR", "Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
            mask = False
            for col in search_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
            _populate_warehouse_tree(df)
            w1_search_label.config(text=f"{len(df)} match(es)")
        else:
            w1_search_label.config(text="")

def pull_item():
    hostname_input = pull_item_entry.get().strip()
    reason = pull_reason_filter_var.get().strip()
    if not hostname_input:
        messagebox.showerror("Error", "No item selected for pull out"); return
    if not reason:
        messagebox.showerror("Error", "Please enter a pull reason"); return

    df_items = load_items()
    df_shelves = load_shelves()
    df_pullouts = load_pullouts()

    # Exact match first, then case-insensitive partial match
    match = df_items[df_items["Hostname"] == hostname_input]
    if match.empty:
        match = df_items[df_items["Hostname"].astype(str).str.lower().str.contains(hostname_input.lower(), na=False)]
    if match.empty:
        messagebox.showerror("Error", f"'{hostname_input}' not found in warehouse"); return
    if len(match) > 1:
        names = "\n".join(match["Hostname"].tolist())
        messagebox.showerror("Error", f"Multiple matches found. Be more specific:\n{names}"); return

    hostname = match.iloc[0]["Hostname"]
    if not messagebox.askyesno("Confirm Pull Out", f"Pull out '{hostname}' from warehouse?\nReason: {reason}"):
        return

    item_row = match.iloc[0]
    shelf = str(item_row.get("Shelf", ""))

    delete_qr(hostname, warehouse=1)
    df_items = df_items[df_items["Hostname"] != hostname].reset_index(drop=True)
    df_pullouts = pd.concat([df_pullouts, pd.DataFrame([{
        "Hostname": hostname,
        "Serial Number": str(item_row.get("Serial Number", "")),
        "Checked By": str(item_row.get("Checked By", "")),
        "Shelf": shelf,
        "Status": str(item_row.get("Status", "")),
        "Remarks": str(item_row.get("Remarks", "")),
        "Pull Reason": reason,
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }])], ignore_index=True)

    save_warehouse_1(df_items, df_shelves, df_pullouts)
    save_log("WAREHOUSE PULL", f"[W1] Hostname: {hostname} | Shelf: {shelf} | Reason: {reason}")
    messagebox.showinfo("Success", f"'{hostname}' pulled out successfully")
    pull_item_entry.delete(0, tk.END)
    pull_reason_filter_var.set("")
    w1_refresh_all()

def undo_pull(event=None):
    # Collect checked rows; fall back to treeview selection if nothing checked
    checked = [iid for iid, state in w1_pull_row_checks.items() if state]
    if not checked:
        sel = tree_pullouts.selection()
        if sel:
            checked = [sel[0]]
    if not checked:
        messagebox.showinfo("Back to Warehouse", "Check at least one row in the Pull History table.")
        return

    restored = 0
    for item_id in checked:
        values = tree_pullouts.item(item_id, "values")
        if not values:
            continue
        # values: ☐(0), Hostname(1), Serial(2), Shelf(3), Status(4), Remarks(5), PullReason(6), Date(7)
        hostname, shelf, status, remarks = values[1], values[3], values[4], values[5]
        if not messagebox.askyesno("Undo Pull", f"Restore '{hostname}' back to the warehouse?\n\nShelf: {shelf}\nStatus: {status}"):
            continue

        df_items = load_items()
        df_shelves = load_shelves()
        df_pullouts = load_pullouts()

        if hostname in df_items["Hostname"].values:
            messagebox.showerror("Error", f"'{hostname}' already exists in warehouse")
            continue

        match = df_pullouts[df_pullouts["Hostname"] == hostname]
        if match.empty:
            messagebox.showerror("Error", f"'{hostname}' not found in pull history")
            continue

        pull_row = match.iloc[0]
        qr_code = ""
        try:
            qr_code = str(uuid.uuid4())
            generate_qr(hostname, qr_code, warehouse=1)
        except Exception as e:
            messagebox.showwarning("Warning", f"QR code not regenerated: {e}")

        for col in ["Serial Number", "Checked By"]:
            if col not in df_items.columns:
                df_items[col] = ""

        df_items = pd.concat([df_items, pd.DataFrame([{
            "QR": qr_code,
            "Hostname": hostname,
            "Serial Number": str(pull_row.get("Serial Number", "")),
            "Checked By": str(pull_row.get("Checked By", "")),
            "Shelf": shelf,
            "Status": status,
            "Remarks": remarks,
            "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }])], ignore_index=True)

        df_pullouts = df_pullouts[df_pullouts["Hostname"] != hostname].reset_index(drop=True)
        save_warehouse_1(df_items, df_shelves, df_pullouts)
        save_log("UNDO PULL", f"[W1] Hostname: {hostname} | Shelf: {shelf}")
        restored += 1

    if restored:
        messagebox.showinfo("Restored", f"{restored} item(s) restored to Warehouse 1.")
    show_pullouts()

def unstage_from_warehouse(event=None):
    # Collect checked rows; fall back to treeview selection if nothing checked
    checked = [iid for iid, state in w1_row_checks.items() if state]
    if not checked:
        sel = tree_warehouse.selection()
        if sel:
            checked = [sel[0]]
    if not checked:
        messagebox.showinfo("Back to Stage", "Check at least one row in the Warehouse table.")
        return

    moved = 0
    for item_id in checked:
        values = tree_warehouse.item(item_id, "values")
        if not values:
            continue
        # values: ☐(0), QR(1), Hostname(2), Serial(3), Checked By(4), Shelf(5), Status(6), Remarks(7), Date(8)
        hostname, serial, checked_by, shelf, status, remarks = values[2], values[3], values[4], values[5], values[6], values[7]
        if not messagebox.askyesno("Move to Staging", f"Move '{hostname}' back to staging?\n\nShelf: {shelf}\nStatus: {status}"):
            continue
        if any(item['Hostname'] == hostname for item in staged_items):
            messagebox.showerror("Error", f"'{hostname}' is already in staging")
            continue

        df_items = load_items()
        df_shelves = load_shelves()
        delete_qr(hostname, warehouse=1)
        df_items = df_items[df_items["Hostname"] != hostname].reset_index(drop=True)
        save_warehouse_1(df_items, df_shelves)
        staged_items.append({"Hostname": hostname, "Serial Number": serial, "Checked By": checked_by, "Shelf": shelf, "Status": status, "Remarks": remarks})
        save_log("UNSTAGE", f"[W1] Hostname: {hostname} | Shelf: {shelf}")
        moved += 1

    if moved:
        messagebox.showinfo("Moved", f"{moved} item(s) moved back to staging.")
        update_staged_display()
        w1_refresh_all()

# ========== W1 SHELF MANAGEMENT ==========

def add_shelf():
    new_shelf = remove_shelf_var.get().strip()
    if not new_shelf:
        messagebox.showerror("Error", "Enter shelf name"); return
    df_shelves = load_shelves()
    if new_shelf in df_shelves["Shelf"].values:
        messagebox.showerror("Error", "Shelf already exists"); return
    df_shelves = pd.concat([df_shelves, pd.DataFrame([{"Shelf": new_shelf, "Status": "AVAILABLE"}])], ignore_index=True)
    df_shelves = df_shelves.sort_values("Shelf", ignore_index=True)
    save_warehouse_1(load_items(), df_shelves)
    messagebox.showinfo("Success", f"Shelf '{new_shelf}' added")
    remove_shelf_var.set("")
    update_all_shelf_dropdowns()

def remove_shelf():
    shelf_name = remove_shelf_var.get().strip()
    if not shelf_name:
        messagebox.showerror("Error", "Select a shelf to remove"); return
    df_items = load_items()
    df_shelves = load_shelves()
    items_in_shelf = df_items[df_items["Shelf"] == shelf_name]
    if not items_in_shelf.empty:
        messagebox.showerror("Error", f"Cannot remove shelf '{shelf_name}' - it has {len(items_in_shelf)} item(s)"); return
    if shelf_name not in df_shelves["Shelf"].values:
        messagebox.showerror("Error", f"Shelf '{shelf_name}' does not exist"); return
    df_shelves = df_shelves[df_shelves["Shelf"] != shelf_name].sort_values("Shelf", ignore_index=True)
    save_warehouse_1(df_items, df_shelves)
    messagebox.showinfo("Success", f"Shelf '{shelf_name}' removed")
    remove_shelf_var.set("")
    update_all_shelf_dropdowns()

def set_shelf_status(new_status):
    shelf = shelf_control_var.get()
    if not shelf:
        messagebox.showerror("Error", "Select a shelf from Shelf Control"); return
    df_items = load_items()
    df_shelves = load_shelves()
    idx = df_shelves[df_shelves["Shelf"] == shelf].index
    if len(idx) == 0:
        return
    df_shelves.at[idx[0], "Status"] = new_status
    df_shelves.at[idx[0], "Date_Full"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if new_status == "FULL" else None
    save_warehouse_1(df_items, df_shelves)
    save_log("SHELF STATUS", f"Shelf: {shelf} → {new_status}")
    w1_status_label.config(text=f"{shelf} → {new_status}")
    w1_refresh_all()

# ========== SHARED HELPERS ==========

def _filter_by_date(df, date_from, date_to):
    """Filter a DataFrame by date range. Both args are optional strings 'YYYY-MM-DD'."""
    if not date_from and not date_to:
        return df
    try:
        df = df.copy()
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        if date_from:
            df = df[df["Date"] >= pd.to_datetime(date_from)]
        if date_to:
            df = df[df["Date"] <= pd.to_datetime(date_to) + pd.Timedelta(days=1)]
    except Exception:
        pass
    return df

# ========== W1 DISPLAY ==========

def _show_tree(tree):
    for t in (tree_warehouse, tree_available, tree_pullouts, tree_qr):
        if t is not tree:
            t.pack_forget()
    tree.pack(fill="both", expand=True)

def _open_qr_gallery(warehouse, filter_keys=None):
    """Shared QR gallery window for both warehouses.
    filter_keys: if provided, only show items whose QR key is in this list.
    """
    from PIL import Image, ImageTk
    wh_label = f"Warehouse {warehouse}"
    bg_color = "#2c3e50" if warehouse == 1 else "#1a5276"
    btn_color = "#1a252f" if warehouse == 1 else "#154360"

    qr_win = tk.Toplevel(root)
    qr_win.title(f"Stored QR Codes — {wh_label}"
                 + (f"  [{len(filter_keys)} selected]" if filter_keys else ""))
    qr_win.geometry("860x560")

    toolbar = tk.Frame(qr_win, bg=bg_color)
    toolbar.pack(fill="x")
    tk.Label(toolbar, text=f"Stored QR Codes — {wh_label}",
             bg=bg_color, fg="white", font=("Helvetica", 10, "bold")).pack(side="left", padx=10, pady=6)
    search_var = tk.StringVar()
    tk.Label(toolbar, text="Search:", bg=bg_color, fg="white").pack(side="left", padx=(20, 2))
    tk.Entry(toolbar, textvariable=search_var, width=18).pack(side="left", pady=4)
    count_lbl = tk.Label(toolbar, text="", bg=bg_color, fg="#aed6f1")
    count_lbl.pack(side="left", padx=10)
    tk.Button(toolbar, text="↻", command=lambda: _load_gallery(search_var.get()),
              bg=btn_color, fg="white", relief="flat", padx=8).pack(side="right", padx=8, pady=4)

    container = tk.Frame(qr_win)
    container.pack(fill="both", expand=True)
    canvas_qr = tk.Canvas(container, bg="#f4f6f7", highlightthickness=0)
    scrollbar_qr = ttk.Scrollbar(container, orient="vertical", command=canvas_qr.yview)
    canvas_qr.configure(yscrollcommand=scrollbar_qr.set)
    scrollbar_qr.pack(side="right", fill="y")
    canvas_qr.pack(side="left", fill="both", expand=True)
    canvas_qr.bind("<MouseWheel>", lambda e: canvas_qr.yview_scroll(int(-1*(e.delta/120)), "units"))

    inner = tk.Frame(canvas_qr, bg="#f4f6f7")
    canvas_window_id = canvas_qr.create_window((0, 0), window=inner, anchor="nw")
    _img_refs = []

    def _load_gallery(keyword=""):
        for w in inner.winfo_children():
            w.destroy()
        _img_refs.clear()
        COLS, THUMB, PAD = 4, 120, 14
        row_f = col_f = shown = 0
        df = load_items() if warehouse == 1 else load_items_w2()

        for _, row in df.iterrows():
            if warehouse == 1:
                key   = str(row.get("Hostname", ""))
                sub   = str(row.get("Serial Number", ""))
                shelf = str(row.get("Shelf", ""))
                path  = qr_path_for(key, warehouse=1)
                kw_fields = [key, shelf]
                cell_labels = [(key, ("Helvetica", 8, "bold"), "#2c3e50", 0),
                               (f"S/N: {sub}", ("Helvetica", 7), "#555", 0),
                               (f"Shelf: {shelf}", ("Helvetica", 7), "gray", 0)]
            else:
                set_id  = str(row.get("Set ID", ""))
                eq_type = str(row.get("Equipment Type", ""))
                shelf   = str(row.get("Shelf", ""))
                key     = f"{set_id}-{eq_type}"
                sub     = str(row.get("Serial Number", ""))
                host    = str(row.get("Hostname", ""))
                path    = qr_path_for(key, warehouse=2)
                kw_fields = [set_id, eq_type, shelf]
                cell_labels = [(key, ("Helvetica", 8, "bold"), "#2c3e50", 0),
                               (host, ("Helvetica", 7, "italic"), "#2c3e50", 0),
                               (f"S/N: {sub}", ("Helvetica", 7), "#555", 0),
                               (f"Shelf: {shelf}", ("Helvetica", 7), "gray", 0)]

            # Filter by selection keys (from Stored QR button) first,
            # then by the gallery's own search box
            if filter_keys is not None and key not in filter_keys:
                continue
            if keyword and not any(keyword.lower() in f.lower() for f in kw_fields):
                continue

            cell = tk.Frame(inner, bg="white", bd=1, relief="solid", padx=PAD, pady=PAD)
            cell.grid(row=row_f, column=col_f, padx=8, pady=8, sticky="n")
            if os.path.exists(path):
                try:
                    img = Image.open(path).resize((THUMB, THUMB), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    _img_refs.append(photo)
                    tk.Label(cell, image=photo, bg="white").pack()
                except Exception:
                    tk.Label(cell, text="[Error]", bg="white", fg="red", width=14).pack()
            else:
                tk.Label(cell, text="[No QR File]", bg="#fdf2f2", fg="#c0392b",
                         width=14, height=6, font=("Helvetica", 8)).pack()
            for text, font, fg, _ in cell_labels:
                tk.Label(cell, text=text, bg="white", font=font, fg=fg, wraplength=130).pack(pady=(4,0) if _ == 0 else 0)

            col_f += 1; shown += 1
            if col_f >= COLS:
                col_f = 0; row_f += 1

        if shown == 0:
            tk.Label(inner, text="No QR codes found.", bg="#f4f6f7",
                     font=("Helvetica", 10), fg="gray").grid(row=0, column=0, padx=20, pady=40)
        count_lbl.config(text=f"{shown} QR code(s)")
        inner.update_idletasks()
        canvas_qr.configure(scrollregion=canvas_qr.bbox("all"))
        canvas_qr.itemconfigure(canvas_window_id, width=canvas_qr.winfo_width())

    canvas_qr.bind("<Configure>", lambda e: canvas_qr.itemconfigure(canvas_window_id, width=e.width))
    search_var.trace_add("write", lambda *_: _load_gallery(search_var.get()))
    _load_gallery()

def show_qr_codes():    _open_qr_gallery(warehouse=1)
def w2_show_qr_codes(): _open_qr_gallery(warehouse=2)

# ── Checkbox helpers ──────────────────────────────────────────

def _get_w1_selected_rows():
    """Return list of row-value tuples for checked (or all) W1 warehouse rows.
    w1_row_checks maps iid -> bool (plain bool, not BooleanVar)."""
    checked = [iid for iid, state in w1_row_checks.items() if state]
    if checked:
        return [tree_warehouse.item(iid, "values") for iid in checked]
    # Fall back to all visible rows when nothing is explicitly checked
    return [tree_warehouse.item(iid, "values") for iid in tree_warehouse.get_children()]

def _get_w2_selected_rows():
    """Return list of row-value tuples for checked (or all) W2 warehouse rows.
    w2_row_checks maps iid -> bool (plain bool, not BooleanVar)."""
    checked = [iid for iid, state in w2_row_checks.items() if state]
    if checked:
        return [tree_w2_warehouse.item(iid, "values") for iid in checked]
    return [tree_w2_warehouse.item(iid, "values") for iid in tree_w2_warehouse.get_children()]

def generate_stored_qr(warehouse=1):
    """Generate QR PNGs for selected/filtered items, open the QR gallery filtered
    to those items, and write a 'qr_selection_w1' / 'qr_selection_w2' sheet to
    the warehouse Excel with the details of those items.

    W1 column layout: ☐(0), QR(1), Hostname(2), Serial(3), Checked By(4), Shelf(5), Status(6), Remarks(7), Date(8)
    W2 column layout: ☐(0), QR(1), Set ID(2), Hostname(3), Equip Type(4), Serial(5), Checked By(6), Shelf(7), Status(8), Remarks(9), Date(10)
    """
    if warehouse == 1:
        rows = _get_w1_selected_rows()
    else:
        rows = _get_w2_selected_rows()

    if not rows:
        messagebox.showinfo("Stored QR", "No items to generate QR codes for.")
        return

    # ── 1. Generate QR PNGs ───────────────────────────────────
    count_ok = count_skip = 0
    qr_keys = []   # track keys generated so the gallery can filter to them

    for values in rows:
        try:
            if warehouse == 1:
                hostname = str(values[2])
                generate_qr(hostname, hostname, warehouse=1)
                qr_keys.append(hostname)
            else:
                set_id  = str(values[2])
                eq_type = str(values[4])
                qr_key  = f"{set_id}-{eq_type}"
                generate_qr(qr_key, qr_key, warehouse=2)
                qr_keys.append(qr_key)
            count_ok += 1
        except Exception:
            count_skip += 1

    if count_skip:
        messagebox.showwarning(
            "Generate Files",
            f"{count_ok} QR code(s) generated.\n{count_skip} item(s) skipped due to errors."
        )

    # ── 2. Write selection sheet to Excel ─────────────────────
    sheet_name = "qr_selection_w1" if warehouse == 1 else "qr_selection_w2"
    try:
        if warehouse == 1:
            cols = ["QR", "Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
            # values indices:  1       2          3              4           5       6        7          8
            records = [
                {c: v for c, v in zip(cols, [values[1], values[2], values[3],
                                              values[4], values[5], values[6],
                                              values[7], values[8]])}
                for values in rows
            ]
        else:
            cols = ["QR", "Set ID", "Hostname", "Equipment Type", "Serial Number",
                    "Checked By", "Shelf", "Status", "Remarks", "Date"]
            # values indices: 1     2        3          4                 5
            #                 6          7       8         9          10
            records = [
                {c: v for c, v in zip(cols, [values[1], values[2], values[3], values[4],
                                              values[5], values[6], values[7], values[8],
                                              values[9], values[10]])}
                for values in rows
            ]

        df_sel = pd.DataFrame(records, columns=cols)
        df_sel.insert(0, "Generated At",
                      datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        # Append / replace the selection sheet without touching other sheets
        try:
            from openpyxl import load_workbook
            wb = load_workbook(FILE)
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            ws = wb.create_sheet(sheet_name)
            # Write header
            ws.append(["Generated At"] + cols)
            for rec in df_sel.itertuples(index=False):
                ws.append(list(rec))
            ws.protection.sheet = True
            ws.protection.enable()
            wb.save(FILE)
        except PermissionError:
            _excel_locked_error()

    except Exception as e:
        messagebox.showwarning("Stored QR", f"QR codes generated but Excel sheet could not be saved:\n{e}")

    # ── 3. Prompt user for PDF + Excel export ────────────────
    # ── 3. Prompt user for PDF + Excel export ────────────────
    def _do_generate_files():
        # ── Scan existing PDFs for this warehouse ─────────────
        pdf_folder = QR_LABELS_FOLDER_W1 if warehouse == 1 else QR_LABELS_FOLDER_W2
        existing_pdfs = []
        if os.path.exists(pdf_folder):
            existing_pdfs = sorted(
                [os.path.splitext(f)[0] for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")],
                reverse=True
            )

        # ── Scan existing Excel files in warehouse-specific export folder ──
        excel_folder = EXCEL_FOLDER_W1 if warehouse == 1 else EXCEL_FOLDER_W2
        os.makedirs(excel_folder, exist_ok=True)
        existing_excels = sorted(
            [os.path.splitext(f)[0] for f in os.listdir(excel_folder)
             if f.lower().endswith(".xlsx")],
            reverse=True
        )

        # ── Build dialog ───────────────────────────────────────
        name_win = tk.Toplevel(root)
        name_win.title("Generate Files — Name Your Export")
        name_win.resizable(False, False)
        name_win.transient(root)
        name_win.grab_set()

        tk.Label(name_win, text="Generate PDF Labels & Excel Export",
                 font=("Helvetica", 10, "bold"), bg="#6c3483", fg="white",
                 padx=10, pady=6).pack(fill="x")

        form = tk.Frame(name_win, padx=16, pady=12)
        form.pack()

        # ── Section header helper ──────────────────────────────
        def _section(parent, text, row):
            tk.Label(parent, text=text, font=("Helvetica", 8, "bold"), fg="#6c3483",
                     anchor="w").grid(row=row, column=0, columnspan=3, sticky="w", pady=(10, 2))

        # ── PDF row ────────────────────────────────────────────
        _section(form, "▸ PDF Label File", 0)
        tk.Label(form, text="File Name:", anchor="w", width=14).grid(row=1, column=0, sticky="w", pady=3)
        pdf_name_var = tk.StringVar(value="")
        pdf_name_cb  = ttk.Combobox(form, textvariable=pdf_name_var, width=28,
                                     values=existing_pdfs)
        pdf_name_cb.grid(row=1, column=1, pady=3, padx=(4, 0))
        tk.Label(form, text=".pdf", fg="gray", font=("Helvetica", 8)).grid(row=1, column=2, sticky="w", padx=(3, 0))
        tk.Label(form, text="  ↳ Select existing PDF to append pages to, or type a new name",
                 fg="gray", font=("Helvetica", 7)).grid(row=2, column=1, columnspan=2, sticky="w")

        # ── Excel file row ─────────────────────────────────────
        _section(form, "▸ Excel File", 3)
        tk.Label(form, text="File Name:", anchor="w", width=14).grid(row=4, column=0, sticky="w", pady=3)
        file_name_var = tk.StringVar(value="")
        file_name_cb  = ttk.Combobox(form, textvariable=file_name_var, width=28,
                                      values=existing_excels)
        file_name_cb.grid(row=4, column=1, pady=3, padx=(4, 0))
        tk.Label(form, text=".xlsx", fg="gray", font=("Helvetica", 8)).grid(row=4, column=2, sticky="w", padx=(3, 0))
        tk.Label(form, text="  ↳ Select existing Excel to append a sheet to, or type a new name",
                 fg="gray", font=("Helvetica", 7)).grid(row=5, column=1, columnspan=2, sticky="w")

        # ── Sheet name row — dynamically lists sheets when an existing Excel is chosen ──
        _section(form, "▸ Excel Sheet", 6)
        tk.Label(form, text="Sheet Name:", anchor="w", width=14).grid(row=7, column=0, sticky="w", pady=3)
        sheet_name_var = tk.StringVar(value="")
        sheet_name_cb  = ttk.Combobox(form, textvariable=sheet_name_var, width=28)
        sheet_name_cb.grid(row=7, column=1, pady=3, padx=(4, 0))
        tk.Label(form, text="  ↳ New sheet name to add (existing sheet of same name will be replaced)",
                 fg="gray", font=("Helvetica", 7)).grid(row=8, column=1, columnspan=2, sticky="w")

        def _refresh_sheet_list(*_):
            """Populate sheet dropdown whenever the chosen Excel file changes."""
            chosen = file_name_var.get().strip()
            if not chosen:
                sheet_name_cb["values"] = []
                return
            safe = chosen if chosen.lower().endswith(".xlsx") else chosen + ".xlsx"
            xl_peek_folder = EXCEL_FOLDER_W1 if warehouse == 1 else EXCEL_FOLDER_W2
            path = os.path.join(xl_peek_folder, safe.replace("/", "-").replace("\\", "-"))
            if os.path.exists(path):
                try:
                    from openpyxl import load_workbook
                    wb_peek = load_workbook(path, read_only=True)
                    sheet_name_cb["values"] = wb_peek.sheetnames
                    wb_peek.close()
                    return
                except Exception:
                    pass
            sheet_name_cb["values"] = []

        file_name_cb.bind("<<ComboboxSelected>>", _refresh_sheet_list)
        file_name_var.trace_add("write", _refresh_sheet_list)

        _wh_label = "Warehouse 1" if warehouse == 1 else "Warehouse 2"
        tk.Label(form, text=f"(Excel files saved to excel_exports/{_wh_label.lower().replace(' ', '_')}/)",
                 fg="gray", font=("Helvetica", 8)).grid(row=9, column=0, columnspan=3, sticky="w", pady=(8, 0))

        error_lbl = tk.Label(form, text="", fg="red", font=("Helvetica", 8))
        error_lbl.grid(row=10, column=0, columnspan=3, sticky="w", pady=(4, 0))

        confirmed = [False]

        def on_confirm():
            pn = pdf_name_var.get().strip()
            fn = file_name_var.get().strip()
            sn = sheet_name_var.get().strip()
            if not pn:
                error_lbl.config(text="Please enter a PDF file name."); return
            if not fn:
                error_lbl.config(text="Please enter an Excel file name."); return
            if not sn:
                error_lbl.config(text="Please enter a sheet name."); return
            confirmed[0] = True
            name_win.destroy()

        btn_row = tk.Frame(name_win, pady=10)
        btn_row.pack()
        tk.Button(btn_row, text="GENERATE", command=on_confirm,
                  bg="#6c3483", fg="white", width=12).pack(side="left", padx=6)
        tk.Button(btn_row, text="Cancel", command=name_win.destroy,
                  width=10).pack(side="left", padx=6)

        name_win.update_idletasks()
        px, py = root.winfo_rootx(), root.winfo_rooty()
        pw, ph = root.winfo_width(), root.winfo_height()
        nw, nh = name_win.winfo_reqwidth(), name_win.winfo_reqheight()
        name_win.geometry(f"+{px+(pw-nw)//2}+{py+(ph-nh)//2}")
        name_win.focus_force()
        root.wait_window(name_win)

        if not confirmed[0]:
            return

        pdf_name_str   = pdf_name_var.get().strip()
        file_name_str  = file_name_var.get().strip()
        sheet_name_str = sheet_name_var.get().strip()

        # ── Check for already-generated items ─────────────────
        already_in_pdf   = []
        already_in_excel = []

        # Check PDF sidecar
        import json
        pdf_folder  = QR_LABELS_FOLDER_W1 if warehouse == 1 else QR_LABELS_FOLDER_W2
        safe_pdf    = pdf_name_str.replace(" ", "_").replace("/", "-").replace("\\", "-")
        if not safe_pdf.lower().endswith(".pdf"):
            safe_pdf += ".pdf"
        sidecar_path = os.path.join(pdf_folder, safe_pdf + ".keys.json")
        if os.path.exists(sidecar_path):
            try:
                with open(sidecar_path, "r") as kf:
                    existing_keys = set(json.load(kf).get("keys", []))
                for values in rows:
                    if warehouse == 1:
                        key = str(values[2])
                        label = f"  • {values[2]}  (Serial: {values[3]})"
                    else:
                        key   = f"{values[2]}-{values[4]}"
                        label = f"  • {values[2]} — {values[4]}  (Serial: {values[5]})"
                    if key in existing_keys:
                        already_in_pdf.append(label)
            except Exception:
                pass

        # Check Excel sheet
        safe_xl = file_name_str.replace("/", "-").replace("\\", "-")
        if not safe_xl.lower().endswith(".xlsx"):
            safe_xl += ".xlsx"
        xl_check_folder = EXCEL_FOLDER_W1 if warehouse == 1 else EXCEL_FOLDER_W2
        excel_check_path = os.path.join(xl_check_folder, safe_xl)
        if os.path.exists(excel_check_path):
            try:
                from openpyxl import load_workbook
                wb_chk = load_workbook(excel_check_path, data_only=True)
                for ws_p in wb_chk.worksheets:
                    ws_p.protection.sheet = False
                if sheet_name_str in wb_chk.sheetnames:
                    ws_chk = wb_chk[sheet_name_str]
                    existing_xl_keys = set()
                    for xl_row in ws_chk.iter_rows(min_row=2, values_only=True):
                        if xl_row and xl_row[1] is not None:
                            if warehouse == 1:
                                existing_xl_keys.add((str(xl_row[1]), str(xl_row[2])))
                            else:
                                existing_xl_keys.add((str(xl_row[1]), str(xl_row[3]), str(xl_row[4])))
                    for values in rows:
                        if warehouse == 1:
                            key   = (str(values[2]), str(values[3]))
                            label = f"  • {values[2]}  (Serial: {values[3]})"
                        else:
                            key   = (str(values[2]), str(values[4]), str(values[5]))
                            label = f"  • {values[2]} — {values[4]}  (Serial: {values[5]})"
                        if key in existing_xl_keys:
                            already_in_excel.append(label)
                wb_chk.close()
            except Exception:
                pass

        if already_in_pdf or already_in_excel:
            parts = []
            if already_in_pdf:
                parts.append(
                    f"Already in PDF '{pdf_name_str}'  ({len(already_in_pdf)} item(s)):\n"
                    + "\n".join(already_in_pdf)
                )
            if already_in_excel:
                parts.append(
                    f"Already in Excel '{file_name_str}' / sheet '{sheet_name_str}'  ({len(already_in_excel)} item(s)):\n"
                    + "\n".join(already_in_excel)
                )
            msg = (
                "Some selected items were already generated before:\n\n"
                + "\n\n".join(parts)
                + "\n\nDo you want to continue? (duplicates will be skipped)"
            )
            if not messagebox.askyesno("Already Generated", msg):
                return

        # --- Generate PDF ---
        pdf_msg = ""
        try:
            if warehouse == 1:
                pdf_items = [
                    {
                        "Hostname":      str(values[2]),
                        "Serial Number": str(values[3]),
                        "Checked By":    str(values[4]),
                        "Shelf":         str(values[5]),
                        "Status":        str(values[6]),
                        "Remarks":       str(values[7]),
                        "_warehouse":    1,
                    }
                    for values in rows
                ]
            else:
                pdf_items = [
                    {
                        "Set ID":         str(values[2]),
                        "Hostname":       str(values[3]),
                        "Equipment Type": str(values[4]),
                        "Serial Number":  str(values[5]),
                        "Checked By":     str(values[6]),
                        "Shelf":          str(values[7]),
                        "Status":         str(values[8]),
                        "Remarks":        str(values[9]),
                        "_warehouse":     2,
                    }
                    for values in rows
                ]
            pdf_path = generate_qr_pdf(pdf_items, custom_name=pdf_name_str)
            pdf_msg = f"PDF saved to:\n{pdf_path}"
        except Exception as pdf_err:
            pdf_msg = f"PDF generation failed: {pdf_err}"

        # --- Generate Excel ---
        excel_msg = ""
        try:
            import stat
            from openpyxl import load_workbook
            from openpyxl import Workbook as _Workbook

            safe_fname = file_name_str.replace("/", "-").replace("\\", "-")
            if not safe_fname.lower().endswith(".xlsx"):
                safe_fname += ".xlsx"
            xl_save_folder = EXCEL_FOLDER_W1 if warehouse == 1 else EXCEL_FOLDER_W2
            os.makedirs(xl_save_folder, exist_ok=True)
            excel_path = os.path.join(xl_save_folder, safe_fname)

            if warehouse == 1:
                cols_xl = ["Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
                records_xl = [
                    [values[2], values[3], values[4], values[5], values[6], values[7], values[8]]
                    for values in rows
                ]
            else:
                cols_xl = ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
                records_xl = [
                    [values[2], values[3], values[4], values[5], values[6], values[7], values[8], values[9], values[10]]
                    for values in rows
                ]

            gen_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            header = ["Generated At"] + cols_xl

            # ── Lift read-only if file exists ─────────────────────
            if os.path.exists(excel_path):
                try:
                    import ctypes
                    ctypes.windll.kernel32.SetFileAttributesW(excel_path, 0x80)
                except Exception:
                    pass
                os.chmod(excel_path, stat.S_IWRITE | stat.S_IREAD)

                # Read all existing sheets into plain Python lists (read_only=True is safe)
                wb_read = load_workbook(excel_path, read_only=True)
                existing_sheets = {}
                for sname in wb_read.sheetnames:
                    existing_sheets[sname] = [
                        list(row) for row in wb_read[sname].iter_rows(values_only=True)
                    ]
                wb_read.close()

                # Build a fresh workbook with all existing data preserved
                wb_xl = _Workbook()
                wb_xl.remove(wb_xl.active)  # remove blank default sheet
                for sname, srows in existing_sheets.items():
                    ws_p = wb_xl.create_sheet(sname)
                    for row in srows:
                        ws_p.append([v if v is not None else "" for v in row])

                # Append new rows into target sheet, skipping duplicates
                if sheet_name_str in existing_sheets:
                    ws_xl = wb_xl[sheet_name_str]
                    existing_xl_keys = set()
                    for row in existing_sheets[sheet_name_str][1:]:  # skip header row
                        if row and len(row) > 2 and row[1] is not None:
                            if warehouse == 1:
                                existing_xl_keys.add((str(row[1]), str(row[2])))
                            else:
                                existing_xl_keys.add((str(row[1]), str(row[3]), str(row[4])))
                    for rec in records_xl:
                        if warehouse == 1:
                            key = (str(rec[0]), str(rec[1]))
                        else:
                            key = (str(rec[0]), str(rec[2]), str(rec[3]))
                        if key not in existing_xl_keys:
                            ws_xl.append([gen_at] + [str(v) for v in rec])
                else:
                    ws_xl = wb_xl.create_sheet(sheet_name_str)
                    ws_xl.append(header)
                    for rec in records_xl:
                        ws_xl.append([gen_at] + [str(v) for v in rec])

                for ws_p in wb_xl.worksheets:
                    ws_p.protection.sheet = True
                    ws_p.protection.enable()
                wb_xl.save(excel_path)
                wb_xl.close()

            else:
                # Brand new file
                wb_xl = _Workbook()
                ws_xl = wb_xl.active
                ws_xl.title = sheet_name_str
                ws_xl.append(header)
                for rec in records_xl:
                    ws_xl.append([gen_at] + [str(v) for v in rec])
                for ws_p in wb_xl.worksheets:
                    ws_p.protection.sheet = True
                    ws_p.protection.enable()
                wb_xl.save(excel_path)
                wb_xl.close()

            # ── Lock read-only via Windows API + chmod ────────────
            try:
                import ctypes
                ctypes.windll.kernel32.SetFileAttributesW(excel_path, 0x01)
            except Exception:
                pass
            os.chmod(excel_path, stat.S_IREAD | stat.S_IRGRP | stat.S_IROTH)

            excel_msg = f"Excel saved to:\n{excel_path}"
            _last_excel_path[warehouse] = excel_path
        except Exception as xl_err:
            excel_msg = f"Excel export failed: {xl_err}"

        messagebox.showinfo("Generate Files",
            f"{count_ok} QR code(s) processed.\n\n{pdf_msg}\n\n{excel_msg}")

    _do_generate_files()

    # ── 4. Open gallery filtered to the generated items ───────
    _open_qr_gallery(warehouse=warehouse, filter_keys=qr_keys)

def w1_generate_stored_qr(): generate_stored_qr(warehouse=1)
def w2_generate_stored_qr(): generate_stored_qr(warehouse=2)

def view_excel(warehouse=None):
    """Open an Excel file manager dialog similar to the QR Label Manager."""
    manager = tk.Toplevel(root)
    manager.title(f"Excel File Manager — Warehouse {warehouse}" if warehouse else "Excel File Manager")
    manager.geometry("700x500")
    manager.resizable(False, False)

    # ── Header ──────────────────────────────────────────────
    hdr = tk.Frame(manager, bg="#1e8449")
    hdr.pack(fill="x")
    tk.Label(hdr, text="Generated Excel Files", font=("Helvetica", 10, "bold"),
             bg="#1e8449", fg="white").pack(side="left", padx=10, pady=6)
    sel_count_lbl = tk.Label(hdr, text="", font=("Helvetica", 9),
                              bg="#1e8449", fg="#f0b429")
    sel_count_lbl.pack(side="right", padx=10)

    # ── Checklist canvas area ────────────────────────────────
    list_frame = tk.Frame(manager, bd=1, relief="sunken")
    list_frame.pack(fill="both", expand=True, padx=10, pady=(8, 4))

    canvas_cl = tk.Canvas(list_frame, bg="white", highlightthickness=0)
    sb_cl = ttk.Scrollbar(list_frame, orient="vertical", command=canvas_cl.yview)
    canvas_cl.configure(yscrollcommand=sb_cl.set)
    sb_cl.pack(side="right", fill="y")
    canvas_cl.pack(side="left", fill="both", expand=True)
    canvas_cl.bind("<MouseWheel>", lambda e: canvas_cl.yview_scroll(int(-1*(e.delta/120)), "units"))

    inner_cl = tk.Frame(canvas_cl, bg="white")
    cw_id = canvas_cl.create_window((0, 0), window=inner_cl, anchor="nw")
    canvas_cl.bind("<Configure>", lambda e: canvas_cl.itemconfigure(cw_id, width=e.width))

    row_data   = []   # (iid, full_path, wh_label, filename, date_str, size_str, var)
    check_vars = {}   # iid -> BooleanVar

    def _refresh_sel_count():
        n = sum(1 for v in check_vars.values() if v.get())
        sel_count_lbl.config(text=f"{n} selected" if n else "")
        clear_btn.config(state="normal" if n else "disabled")

    def _toggle_all():
        want = not all(v.get() for v in check_vars.values())
        for v in check_vars.values():
            v.set(want)
        _refresh_sel_count()
        _repaint_rows()

    def _repaint_rows():
        for iid, _, _, _, _, _, var in row_data:
            try:
                fr = inner_cl.nametowidget(f"row_{iid}")
                fr.config(bg="#e8f5e9" if var.get() else ("white" if row_data.index(
                    next(r for r in row_data if r[0] == iid)) % 2 == 0 else "#f7f9fc"))
            except Exception:
                pass

    EVEN_BG, ODD_BG, SEL_BG = "white", "#f7f9fc", "#e8f5e9"

    def load_excel_files():
        for w in inner_cl.winfo_children():
            w.destroy()
        row_data.clear()
        check_vars.clear()

        now = datetime.now()
        all_warehouses = [("Warehouse 1", 1), ("Warehouse 2", 2)]
        filtered = [(wl, wn) for wl, wn in all_warehouses if warehouse is None or wn == warehouse]

        hdr_row = tk.Frame(inner_cl, bg="#dce3f0")
        hdr_row.pack(fill="x")
        tk.Label(hdr_row, text="✔", width=3, bg="#dce3f0", font=("Helvetica", 9, "bold")).pack(side="left", padx=(6,0))
        for txt, w in [("Warehouse", 110), ("Filename", 240), ("Created", 155), ("Size", 60)]:
            tk.Label(hdr_row, text=txt, width=w//7, bg="#dce3f0",
                     font=("Helvetica", 9, "bold"), anchor="w").pack(side="left", padx=4, pady=5)

        idx = 0
        for warehouse_label, wh_num in filtered:
            # Scan warehouse-specific Excel export folder only
            xl_folder = EXCEL_FOLDER_W1 if wh_num == 1 else EXCEL_FOLDER_W2
            if not os.path.exists(xl_folder):
                continue
            files = sorted(
                [f for f in os.listdir(xl_folder) if f.lower().endswith(".xlsx")],
                reverse=True
            )
            for f in files:
                full_path = os.path.join(xl_folder, f)
                size_kb = round(os.path.getsize(full_path) / 1024, 1)
                try:
                    mtime = os.path.getmtime(full_path)
                    file_dt = datetime.fromtimestamp(mtime)
                    delta = now - file_dt
                    age = "Today" if delta.days == 0 else ("1 day ago" if delta.days == 1 else f"{delta.days} days ago")
                    date_str = f"{file_dt.strftime('%Y-%m-%d')}  ({age})"
                except Exception:
                    date_str = "Unknown"

                iid = f"row{idx}"
                var = tk.BooleanVar(value=False)
                check_vars[iid] = var
                bg = EVEN_BG if idx % 2 == 0 else ODD_BG

                row_fr = tk.Frame(inner_cl, bg=bg, name=f"row_{iid}", cursor="hand2")
                row_fr.pack(fill="x")

                def _make_toggle(v=var):
                    def _toggle(e=None):
                        v.set(not v.get())
                        _refresh_sel_count()
                        _repaint_rows()
                    return _toggle

                cb = tk.Checkbutton(row_fr, variable=var, bg=bg,
                                    command=lambda: [_refresh_sel_count(), _repaint_rows()])
                cb.pack(side="left", padx=(6, 0), pady=4)
                for txt, w in [(warehouse_label, 16), (f, 34), (date_str, 22), (f"{size_kb} kb", 8)]:
                    lbl = tk.Label(row_fr, text=txt, bg=bg, anchor="w", width=w,
                                   font=("Helvetica", 9))
                    lbl.pack(side="left", padx=4, pady=4)
                    lbl.bind("<Button-1>", _make_toggle(var))
                row_fr.bind("<Button-1>", _make_toggle(var))

                row_data.append((iid, full_path, warehouse_label, f, date_str, f"{size_kb} kb", var))
                idx += 1
            # Only show files once (not duplicated per warehouse label)
            break

        if idx == 0:
            tk.Label(inner_cl, text="No generated Excel files found.", fg="gray",
                     font=("Helvetica", 10), bg="white").pack(pady=30)

        inner_cl.update_idletasks()
        canvas_cl.configure(scrollregion=canvas_cl.bbox("all"))
        _refresh_sel_count()

    def open_selected():
        chosen = [rd for rd in row_data if rd[6].get()]
        if not chosen:
            messagebox.showwarning("Warning", "Check a file to open.", parent=manager); return
        if len(chosen) > 1:
            messagebox.showerror("Error", "Only 1 file can be opened at a time.\nPlease check only one file.", parent=manager); return
        full_path = chosen[0][1]
        if os.path.exists(full_path):
            os.startfile(full_path)
        else:
            messagebox.showerror("Error", "File not found.", parent=manager)

    def clear_selected():
        chosen = [rd for rd in row_data if rd[6].get()]
        if not chosen:
            messagebox.showwarning("Warning", "Check at least one file to delete.", parent=manager); return
        count = len(chosen)
        prompt = (f"Delete '{chosen[0][3]}'?" if count == 1
                  else f"Delete {count} checked file(s)?\nThis cannot be undone.")
        if not messagebox.askyesno("Confirm Delete", prompt, parent=manager):
            return
        failed = []
        import stat, ctypes
        for _, full_path, _, fname, _, _, _ in chosen:
            try:
                if os.path.exists(full_path):
                    try:
                        ctypes.windll.kernel32.SetFileAttributesW(full_path, 0x80)
                    except Exception:
                        pass
                    os.chmod(full_path, stat.S_IWRITE | stat.S_IREAD)
                    os.remove(full_path)
            except Exception as e:
                failed.append(f"{fname}: {e}")
        load_excel_files()
        if failed:
            messagebox.showerror("Error", "Some files could not be deleted:\n" + "\n".join(failed), parent=manager)
        else:
            messagebox.showinfo("Deleted", f"{count} file(s) deleted.", parent=manager)

    # ── Bottom toolbar ───────────────────────────────────────
    btn_frame_m = tk.Frame(manager)
    btn_frame_m.pack(pady=8)
    tk.Button(btn_frame_m, text="☑", command=_toggle_all, width=4).pack(side="left", padx=4)
    tk.Button(btn_frame_m, text="OPEN", command=open_selected, width=10).pack(side="left", padx=4)
    clear_btn = tk.Button(btn_frame_m, text="✕", command=clear_selected,
                           width=10, bg="#922b21", fg="white", state="disabled")
    clear_btn.pack(side="left", padx=4)

    load_excel_files()

def w1_view_excel(): view_excel(warehouse=1)
def w2_view_excel(): view_excel(warehouse=2)

def view_stored_qr(warehouse=1):
    """Open the QR gallery filtered to items currently visible in the warehouse table.
    If no filter is active, shows all existing QR PNG files."""
    folder = QR_FOLDER_W1 if warehouse == 1 else QR_FOLDER_W2
    if not os.path.exists(folder):
        messagebox.showinfo("View QR", "No QR codes folder found.")
        return
    files = [f for f in os.listdir(folder) if f.lower().endswith(".png")]
    if not files:
        messagebox.showinfo("View QR", "No QR codes have been generated yet.")
        return

    # Build filter_keys from what is currently visible in the warehouse tree
    if warehouse == 1:
        visible_iids = list(tree_warehouse.get_children())
        if visible_iids:
            filter_keys = []
            for iid in visible_iids:
                values = tree_warehouse.item(iid, "values")
                # values: ☐(0), QR(1), Hostname(2), ...
                filter_keys.append(str(values[2]))
        else:
            filter_keys = [os.path.splitext(f)[0].replace("_", " ") for f in files]
    else:
        visible_iids = list(tree_w2_warehouse.get_children())
        if visible_iids:
            filter_keys = []
            for iid in visible_iids:
                values = tree_w2_warehouse.item(iid, "values")
                # values: ☐(0), QR(1), Set ID(2), Hostname(3), Equip Type(4), ...
                filter_keys.append(f"{values[2]}-{values[4]}")
        else:
            filter_keys = [os.path.splitext(f)[0].replace("_", " ") for f in files]

    _open_qr_gallery(warehouse=warehouse, filter_keys=filter_keys)

def w1_view_stored_qr(): view_stored_qr(warehouse=1)
def w2_view_stored_qr(): view_stored_qr(warehouse=2)

def _w1_refresh_select_all_label():
    """Update the Select All / Deselect All button label for W1.
    Safe to call before UI exists — errors are silently swallowed."""
    try:
        all_iids = list(tree_warehouse.get_children())
        checked  = [iid for iid in all_iids if w1_row_checks.get(iid)]
        w1_select_all_btn.config(
            text="DESELECT ALL" if all_iids and len(checked) == len(all_iids) else "SELECT ALL")
    except (NameError, tk.TclError, AttributeError):
        pass

def _w2_refresh_select_all_label():
    """Update the Select All / Deselect All button label for W2.
    Safe to call before UI exists — errors are silently swallowed."""
    try:
        all_iids = list(tree_w2_warehouse.get_children())
        checked  = [iid for iid in all_iids if w2_row_checks.get(iid)]
        w2_select_all_btn.config(
            text="DESELECT ALL" if all_iids and len(checked) == len(all_iids) else "SELECT ALL")
    except (NameError, tk.TclError, AttributeError):
        pass


def show_warehouse():
    try: w1_back_to_wh_btn.pack_forget()
    except Exception: pass
    try: w1_back_to_stage_btn.pack(side="left", padx=(0, 6))
    except Exception: pass
    w1_update_full_shelves_display()
    try: w1_back_to_wh_btn.pack_forget()
    except Exception: pass
    df_items = load_items()
    if "Date" not in df_items.columns:
        df_items["Date"] = ""
    # Apply active date filters if set
    try:
        date_from = w1_date_from_var.get().strip()
        date_to   = w1_date_to_var.get().strip()
        df_items = _filter_by_date(df_items, date_from, date_to)
    except (NameError, Exception):
        pass  # vars not yet created during startup
    _populate_warehouse_tree(df_items)

def show_available():
    try: w1_back_to_wh_btn.pack_forget()
    except Exception: pass
    try: w1_back_to_stage_btn.pack_forget()
    except Exception: pass
    _show_tree(tree_available)
    try: w1_back_to_wh_btn.pack_forget()
    except Exception: pass
    tree_available.delete(*tree_available.get_children())
    try:
        keyword = search_entry.get().strip().lower()
    except Exception:
        keyword = ""
    df = load_shelves().sort_values("Shelf")
    df_items = load_items()
    if keyword:
        mask = False
        for col in ["Shelf", "Status", "Date_Full"]:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]
    for _, row in df.iterrows():
        date_full = row.get("Date_Full", "")
        shelf_name = row["Shelf"]
        item_count = int((df_items["Shelf"] == shelf_name).sum()) if "Shelf" in df_items.columns else 0
        tree_available.insert("", "end", values=(shelf_name, row["Status"], item_count, date_full if pd.notna(date_full) else ""))

def show_pullouts():
    try: w1_back_to_stage_btn.pack_forget()
    except Exception: pass
    try: w1_back_to_wh_btn.pack(side="left", padx=(0, 6))
    except Exception: pass
    _show_tree(tree_pullouts)
    w1_back_to_wh_btn.pack(side="left", padx=(0, 6))
    tree_pullouts.delete(*tree_pullouts.get_children())
    df_po = load_pullouts()
    # Re-apply active search/filter state so it survives view switches
    try:
        keyword        = search_entry.get().strip().lower()
        shelf_filter   = pull_shelf_var.get()
        remarks_filter = pull_remarks_var.get()
        date_from      = w1_date_from_var.get().strip()
        date_to        = w1_date_to_var.get().strip()
    except (NameError, Exception):
        keyword = shelf_filter = remarks_filter = date_from = date_to = ""
    if keyword:
        search_cols = ["Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
        mask = False
        for col in search_cols:
            if col in df_po.columns:
                mask = mask | df_po[col].astype(str).str.lower().str.contains(keyword, na=False)
        df_po = df_po[mask]
    if shelf_filter:   df_po = df_po[df_po["Shelf"] == shelf_filter]
    if remarks_filter: df_po = df_po[df_po["Status"] == remarks_filter]
    df_po = _filter_by_date(df_po, date_from, date_to)
    w1_pull_row_checks.clear()
    for _, row in df_po.iterrows():
        iid = tree_pullouts.insert("", "end", values=(
            "☐",
            *tuple(row.get(c, "") for c in ["Hostname", "Serial Number", "Shelf", "Status", "Remarks", "Pull Reason", "Date"])
        ))
        w1_pull_row_checks[iid] = False
    try:
        all_reasons = sorted(load_pullouts()["Pull Reason"].dropna().unique().tolist())
        pull_reason_filter_entry["values"] = [""] + all_reasons
    except Exception:
        pass
    active = bool(keyword or shelf_filter or remarks_filter or date_from or date_to)
    if active:
        parts = []
        if keyword:        parts.append(f"Search: \"{keyword}\"")
        if shelf_filter:   parts.append(f"Shelf: {shelf_filter}")
        if remarks_filter: parts.append(f"Status: {remarks_filter}")
        if date_from:      parts.append(f"From: {date_from}")
        if date_to:        parts.append(f"To: {date_to}")
        try:
            w1_search_label.config(text=f"{len(df_po)} result(s) — " + " | ".join(parts), fg="darkorange")
        except (NameError, Exception):
            pass

def _populate_warehouse_tree(df):
    _show_tree(tree_warehouse)
    tree_warehouse.delete(*tree_warehouse.get_children())
    w1_row_checks.clear()
    for _, row in df.iterrows():
        iid = tree_warehouse.insert("", "end", values=(
            "☐",
            *tuple(row.get(c, "") for c in ["QR", "Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"])
        ))
        w1_row_checks[iid] = False
    _w1_refresh_select_all_label()

def search_item():
    keyword        = search_entry.get().strip().lower()
    shelf_filter   = pull_shelf_var.get()
    remarks_filter = pull_remarks_var.get()
    date_from      = w1_date_from_var.get().strip()
    date_to        = w1_date_to_var.get().strip()

    # ── Pull history view is active: filter it instead of warehouse ──
    if tree_pullouts.winfo_ismapped():
        df = load_pullouts()
        if keyword:
            search_cols = ["Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
            mask = False
            for col in search_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        if shelf_filter:   df = df[df["Shelf"] == shelf_filter]
        if remarks_filter: df = df[df["Status"] == remarks_filter]
        df = _filter_by_date(df, date_from, date_to)
        tree_pullouts.delete(*tree_pullouts.get_children())
        w1_pull_row_checks.clear()
        for _, row in df.iterrows():
            iid = tree_pullouts.insert("", "end", values=(
                "☐",
                *tuple(row.get(c, "") for c in ["Hostname", "Serial Number", "Shelf", "Status", "Remarks", "Pull Reason", "Date"])
            ))
            w1_pull_row_checks[iid] = False
        parts = []
        if keyword:        parts.append(f"Search: \"{keyword}\"")
        if shelf_filter:   parts.append(f"Shelf: {shelf_filter}")
        if remarks_filter: parts.append(f"Status: {remarks_filter}")
        if date_from:      parts.append(f"From: {date_from}")
        if date_to:        parts.append(f"To: {date_to}")
        label = f"{len(df)} result(s)" + (" — " + " | ".join(parts) if parts else "")
        w1_search_label.config(text=label if parts else "", fg="darkorange" if parts else "blue")
        return

    # ── Default: filter warehouse view ───────────────────────────────
    df = load_items()
    if keyword:
        search_cols = ["QR", "Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]
    if shelf_filter:   df = df[df["Shelf"] == shelf_filter]
    if remarks_filter: df = df[df["Status"] == remarks_filter]
    df = _filter_by_date(df, date_from, date_to)
    _populate_warehouse_tree(df)
    parts = []
    if keyword:        parts.append(f"Search: \"{keyword}\"")
    if shelf_filter:   parts.append(f"Shelf: {shelf_filter}")
    if remarks_filter: parts.append(f"Status: {remarks_filter}")
    if date_from:      parts.append(f"From: {date_from}")
    if date_to:        parts.append(f"To: {date_to}")
    label = f"{len(df)} result(s)" + (" — " + " | ".join(parts) if parts else "")
    w1_search_label.config(text=label if parts else "")

def filter_pull_history():
    """Filter and display the W1 pull history table only. Completely separate from warehouse search."""
    reason = pull_reason_filter_var.get().strip().lower()
    date_from = w1_pull_date_from_var.get().strip()
    date_to = w1_pull_date_to_var.get().strip()
    show_pullouts()  # always switch to pull history view first
    if not reason and not date_from and not date_to:
        return
    tree_pullouts.delete(*tree_pullouts.get_children())
    df = load_pullouts()
    if reason:
        search_cols = ["Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(reason, na=False)
        df = df[mask]
    df = _filter_by_date(df, date_from, date_to)
    # Refresh pull reason dropdown
    all_reasons = sorted(load_pullouts()["Pull Reason"].dropna().unique().tolist())
    pull_reason_filter_entry["values"] = [""] + all_reasons
    w1_pull_row_checks.clear()
    for _, row in df.iterrows():
        iid = tree_pullouts.insert("", "end", values=(
            "☐",
            *tuple(row.get(c, "") for c in ["Hostname", "Serial Number", "Shelf", "Status", "Remarks", "Pull Reason", "Date"])
        ))
        w1_pull_row_checks[iid] = False
    w1_search_label.config(text=f"Pull History Filtered: {len(df)} result(s)")

def w1_update_full_shelves_display():
    df_shelves = load_shelves()
    full_shelves = df_shelves[df_shelves["Status"] == "FULL"]["Shelf"].tolist()
    w1_full_label.config(text="FULL Shelves:\n" + "\n".join(full_shelves) if full_shelves else "FULL Shelves: None")

def w1_refresh_all():
    show_warehouse()
    update_all_shelf_dropdowns()

def update_all_shelf_dropdowns():
    # Load separate lists
    w1_shelf_list = sorted(load_shelves()["Shelf"].tolist())
    w2_shelf_list = sorted(load_shelves_w2()["Shelf"].tolist())

    # Update W1 specific dropdowns
    for dropdown in (shelf_dropdown, shelf_control_dropdown, remove_shelf_dropdown, pull_shelf_dropdown):
        try:
            dropdown["values"] = w1_shelf_list
        except NameError:
            pass # Ignore if UI isn't initialized yet

    # Update W2 specific dropdowns
    for dropdown in (w2_shelf_control_dropdown, w2_remove_shelf_dropdown, w2_pull_shelf_dropdown):
        try:
            dropdown["values"] = w2_shelf_list
        except NameError:
            pass

def select_pull_item(event):
    """Toggle checkbox on click in W1 pull history table."""
    selected = tree_pullouts.selection()
    if selected:
        iid = selected[0]
        if iid in w1_pull_row_checks:
            w1_pull_row_checks[iid] = not w1_pull_row_checks[iid]
            tree_pullouts.set(iid, "CP0", "☑" if w1_pull_row_checks[iid] else "☐")

def w2_select_pull_item(event):
    """Toggle checkbox on click in W2 pull history table."""
    selected = tree_w2_pullouts.selection()
    if selected:
        iid = selected[0]
        if iid in w2_pull_row_checks:
            w2_pull_row_checks[iid] = not w2_pull_row_checks[iid]
            tree_w2_pullouts.set(iid, "CP0", "☑" if w2_pull_row_checks[iid] else "☐")

def select_item(event):
    selected = tree_warehouse.selection()
    if selected:
        iid = selected[0]
        values = tree_warehouse.item(iid, "values")
        # values: ☐(0), QR(1), Hostname(2), Serial Number(3), Checked By(4), Shelf(5), Status(6), Remarks(7), Date(8)

        # Toggle checkbox on single-click
        if iid in w1_row_checks:
            w1_row_checks[iid] = not w1_row_checks[iid]
            tree_warehouse.set(iid, "C0", "☑" if w1_row_checks[iid] else "☐")
            _w1_refresh_select_all_label()

        pull_item_entry.delete(0, tk.END)
        pull_item_entry.insert(0, values[2])
        w1_status_label.config(
            text=f"Selected → Hostname: {values[2]}  |  Shelf: {values[5]}  |  Serial: {values[3]}",
            fg="#1a5276")

# ========== W1 RESET ==========

def reset_ui():
    _clear_input_fields()
    for s in tree_warehouse.selection(): tree_warehouse.selection_remove(s)
    w1_status_label.config(text="")
    w1_search_label.config(text="")
    show_warehouse()

def reset_shelf_control():
    shelf_control_var.set("")
    w1_status_label.config(text="")

def reset_shelf_addition():
    remove_shelf_var.set("")
    w1_status_label.config(text="")

def reset_pull_out():
    pull_item_entry.delete(0, tk.END)
    pull_reason_filter_var.set("")
    w1_pull_date_from_var.set("")
    w1_pull_date_to_var.set("")
    for s in tree_warehouse.selection(): tree_warehouse.selection_remove(s)
    w1_status_label.config(text="")
    w1_search_label.config(text="")
    show_warehouse()

def clear_pull_filters():
    pull_shelf_var.set("")
    pull_remarks_var.set("")
    search_entry.delete(0, tk.END)
    pull_reason_filter_var.set("")
    w1_pull_date_from_var.set("")
    w1_pull_date_to_var.set("")
    w1_date_from_var.set("")
    w1_date_to_var.set("")
    w1_search_label.config(text="")
    show_warehouse()

# ========== W2 STAGING ==========

def update_w2_staged_display():
    w2_staged_listbox.delete(0, tk.END)
    if not staged_sets:
        w2_staged_listbox.insert(tk.END, "No staged sets")
        return
    for s in staged_sets:
        items = s["items"]
        # Collect distinct shelves
        shelves = list(dict.fromkeys(i.get("Shelf", "") for i in items if i.get("Shelf")))
        shelf_str = shelves[0] if len(shelves) == 1 else (f"{len(shelves)} shelves" if shelves else "—")
        # Check completeness (all 4 types present)
        missing = [t for t in EQUIPMENT_TYPES if t not in [i["Equipment Type"] for i in items]]
        flag = " ⚠ missing: " + ", ".join(missing) if missing else ""
        w2_staged_listbox.insert(tk.END, f"{s['set_id']} | {len(items)} items | {shelf_str}{flag}")

def w2_build_set():
    """Open a dialog to build a new equipment set."""
    selected_types = []
    for eq_type, var in w2_equip_vars.items():
        if var.get():
            selected_types.append(eq_type)

    if not selected_types:
        messagebox.showerror("Error", "Please select at least one equipment type"); return

    set_id = next_set_id()

    build_win = tk.Toplevel(root)
    build_win.title(f"Build {set_id}")
    build_win.resizable(False, False)
    build_win.transient(root)

    tk.Label(build_win, text=f"Fill in details for {set_id}", font=("Helvetica", 10, "bold")).pack(pady=(10, 5))

    shelf_list = sorted(load_shelves_w2()["Shelf"].tolist())

    COL_WIDTHS = [12, 18, 18, 16, 16, 13, 20]
    HEADERS    = ["Equipment", "Hostname", "Serial Number", "Checked By", "Shelf", "Status", "Remarks"]

    outer = tk.Frame(build_win, padx=10, pady=5)
    outer.pack(fill="both", expand=True)

    # Fixed header
    hdr_frame = tk.Frame(outer, bg="#dce3f0")
    hdr_frame.pack(fill="x")
    for col, (h, cw) in enumerate(zip(HEADERS, COL_WIDTHS)):
        tk.Label(hdr_frame, text=h, font=("Helvetica", 9, "bold"), width=cw, anchor="w",
                 bg="#dce3f0", padx=5).grid(row=0, column=col, padx=5, pady=4, sticky="w")

    ROW_COLORS = ("#ffffff", "#f0f4ff")

    rows = {}
    for r, eq_type in enumerate(selected_types):
        bg = ROW_COLORS[r % 2]
        row_bg = tk.Frame(outer, bg=bg, bd=1, relief="flat")
        row_bg.pack(fill="x", pady=1)
        tk.Label(row_bg, text=eq_type, width=COL_WIDTHS[0], anchor="w", bg=bg,
                 font=("Helvetica", 9, "bold")).grid(row=0, column=0, padx=5, pady=8, sticky="w")
        hostname_e = tk.Entry(row_bg, width=COL_WIDTHS[1], font=("Helvetica", 9)); hostname_e.grid(row=0, column=1, padx=5, pady=8)
        serial_e   = tk.Entry(row_bg, width=COL_WIDTHS[2], font=("Helvetica", 9)); serial_e.grid(row=0, column=2, padx=5, pady=8)
        checked_e  = tk.Entry(row_bg, width=COL_WIDTHS[3], font=("Helvetica", 9)); checked_e.grid(row=0, column=3, padx=5, pady=8)
        shelf_v = tk.StringVar()
        ttk.Combobox(row_bg, textvariable=shelf_v, values=shelf_list, width=COL_WIDTHS[4], state="readonly",
                     font=("Helvetica", 9)).grid(row=0, column=4, padx=5, pady=8)
        status_v = tk.StringVar()
        ttk.Combobox(row_bg, textvariable=status_v, values=STATUS_CHOICES,
                     width=COL_WIDTHS[5], state="readonly", font=("Helvetica", 9)).grid(row=0, column=5, padx=5, pady=8)
        remarks_e = tk.Entry(row_bg, width=COL_WIDTHS[6], font=("Helvetica", 9)); remarks_e.grid(row=0, column=6, padx=5, pady=8)
        rows[eq_type] = (hostname_e, serial_e, checked_e, shelf_v, status_v, remarks_e)

    error_lbl = tk.Label(outer, text="", fg="red", font=("Helvetica", 8))
    error_lbl.pack(pady=(6, 0))

    def confirm_set():
        df_items_w2 = load_items_w2()
        existing_serials_w2 = df_items_w2["Serial Number"].astype(str).tolist() if "Serial Number" in df_items_w2.columns else []
        staged_serials = []
        for ss in staged_sets:
            for it in ss["items"]:
                if it.get("Serial Number"):
                    staged_serials.append(it["Serial Number"])

        items = []
        for eq_type, (hostname_e, serial_e, checked_e, shelf_v, status_v, remarks_e) in rows.items():
            hostname   = hostname_e.get().strip()
            serial     = serial_e.get().strip()
            checked_by = checked_e.get().strip()
            shelf      = shelf_v.get().strip()
            status     = status_v.get().strip()
            remarks    = remarks_e.get().strip()

            if not hostname:
                error_lbl.config(text=f"Please enter a Hostname for {eq_type}"); return
            if not serial:
                error_lbl.config(text=f"Please enter a Serial Number for {eq_type}"); return
            if not checked_by:
                error_lbl.config(text=f"Please enter Checked By for {eq_type}"); return
            if not shelf:
                error_lbl.config(text=f"Please select a Shelf for {eq_type}"); return
            if not status:
                error_lbl.config(text=f"Please select a Status for {eq_type}"); return

            if serial in existing_serials_w2:
                error_lbl.config(text=f"Serial '{serial}' already exists in Warehouse 2"); return
            if serial in staged_serials:
                error_lbl.config(text=f"Serial '{serial}' already staged in another set"); return
            if serial in [i.get("Serial Number") for i in items]:
                error_lbl.config(text=f"Duplicate serial '{serial}' within this set"); return

            items.append({
                "Equipment Type": eq_type,
                "Hostname":       hostname,
                "Serial Number":  serial,
                "Checked By":     checked_by,
                "Shelf":          shelf,
                "Status":         status,
                "Remarks":        remarks,
            })

        staged_sets.append({"set_id": set_id, "items": items})
        for var in w2_equip_vars.values():
            var.set(False)
        build_win.destroy()
        update_w2_staged_display()
        messagebox.showinfo("Staged", f"{set_id} added to staging with {len(items)} item(s)")

    btn_f = tk.Frame(outer)
    btn_f.pack(pady=8)
    tk.Button(btn_f, text="STAGE", command=confirm_set, width=16).pack(side="left", padx=5)
    tk.Button(btn_f, text="CANCEL", command=build_win.destroy, width=10).pack(side="left", padx=5)

    # Auto-size AFTER all widgets (including buttons) are packed
    build_win.update_idletasks()
    build_win.geometry(f"{build_win.winfo_reqwidth()}x{build_win.winfo_reqheight()}")
    build_win.grab_set()
    build_win.focus_force()

def w2_remove_staged_set():
    global selected_set_index
    sel = w2_staged_listbox.curselection()
    if not sel:
        if not staged_sets:
            messagebox.showinfo("Info", "No staged sets to clear"); return
        if not messagebox.askyesno("Confirm", f"Clear all {len(staged_sets)} staged set(s)?"):
            return
        staged_sets.clear()
        selected_set_index = None
        update_w2_staged_display()
        messagebox.showinfo("Cleared", "All staged sets cleared")
        return
    index = sel[0]
    if index >= len(staged_sets):
        return
    removed = staged_sets.pop(index)
    selected_set_index = None
    update_w2_staged_display()
    messagebox.showinfo("Removed", f"{removed['set_id']} removed from staging")

def w2_put_warehouse():
    if not staged_sets:
        messagebox.showerror("Error", "No staged sets to put"); return

    total_items = sum(len(s["items"]) for s in staged_sets)
    if not messagebox.askyesno("Confirm",
        f"Put {len(staged_sets)} set(s) ({total_items} item(s)) to Warehouse 2?"):
        return

    try:
        df_w2 = load_items_w2()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for s in staged_sets:
            set_id = s["set_id"]
            for item in s["items"]:
                eq_type = item["Equipment Type"]
                qr_label = f"{set_id}-{eq_type}"
                qr_code = str(uuid.uuid4())
                generate_qr(qr_label, qr_label, warehouse=2)
                df_w2 = pd.concat([df_w2, pd.DataFrame([{
                    "QR":             qr_code,
                    "Set ID":         set_id,
                    "Hostname":       item.get("Hostname", ""),
                    "Equipment Type": eq_type,
                    "Serial Number":  item.get("Serial Number", ""),
                    "Checked By":     item.get("Checked By", ""),
                    "Shelf":          item["Shelf"],
                    "Status":         item.get("Status", ""),
                    "Remarks":        item.get("Remarks", ""),
                    "Date":           now_str
                }])], ignore_index=True)
            save_log("PUT WAREHOUSE", f"[W2] Set: {set_id} | Items: {len(s['items'])}")

        save_warehouse_2(df_w2, load_shelves_w2())

        count = len(staged_sets)
        staged_sets.clear()
        update_w2_staged_display()
        w2_pull_item_entry.delete(0, tk.END)
        w2_search_label.config(text="", fg="blue")
        w2_status_label.config(text="")
        w2_refresh_all()
        messagebox.showinfo("Success", f"{count} set(s) added to Warehouse 2.\nQR codes generated.\nUse 'GENERATE FILES' to create PDF labels and export to Excel.")

    except Exception as e:
        messagebox.showerror("Save Error", f"Failed to save:\n{str(e)}")

# ========== W2 DISPLAY ==========

def _show_w2_tree(tree):
    for t in (tree_w2_warehouse, tree_w2_available, tree_w2_pullouts, tree_w2_qr):
        if t is not tree:
            t.pack_forget()
    tree.pack(fill="both", expand=True)

def w2_show_warehouse():
    w2_update_full_shelves_display()
    try: w2_back_to_wh_btn.pack_forget()
    except Exception: pass
    try: w2_back_to_stage_btn.pack(side="left", padx=(0, 6))
    except Exception: pass
    df = load_items_w2()
    if "Date" not in df.columns:
        df["Date"] = ""
    # Re-apply any active search/filter so state survives view switches
    try:
        keyword = w2_search_entry.get().strip().lower()
        shelf_f = w2_pull_shelf_var.get()
        type_f  = w2_type_filter_var.get()
        date_from = w2_date_from_var.get().strip()
        date_to   = w2_date_to_var.get().strip()
    except (NameError, Exception):
        keyword = shelf_f = type_f = date_from = date_to = ""
    active = bool(keyword or shelf_f or type_f or date_from or date_to)
    if keyword:
        search_cols = ["QR", "Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]
    if shelf_f: df = df[df["Shelf"] == shelf_f]
    if type_f:  df = df[df["Equipment Type"] == type_f]
    df = _filter_by_date(df, date_from, date_to)
    _populate_w2_warehouse_tree(df)
    if active:
        w2_search_label.config(text=f"🔍 Active filter — {len(df)} result(s) shown", fg="darkorange")

def w2_show_available():
    _show_w2_tree(tree_w2_available)
    try: w2_back_to_wh_btn.pack_forget()
    except Exception: pass
    try: w2_back_to_stage_btn.pack_forget()
    except Exception: pass
    tree_w2_available.delete(*tree_w2_available.get_children())
    try:
        keyword = w2_search_entry.get().strip().lower()
    except Exception:
        keyword = ""
    df = load_shelves_w2().sort_values("Shelf")
    df_items_w2 = load_items_w2()
    if keyword:
        mask = False
        for col in ["Shelf", "Status", "Date_Full"]:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]
    for _, row in df.iterrows():
        date_full  = row.get("Date_Full", "")
        shelf_name = row["Shelf"]
        item_count = int((df_items_w2["Shelf"] == shelf_name).sum()) if "Shelf" in df_items_w2.columns else 0
        tree_w2_available.insert("", "end", values=(shelf_name, row["Status"], item_count, date_full if pd.notna(date_full) else ""))

def w2_show_pullouts():
    _show_w2_tree(tree_w2_pullouts)
    try: w2_back_to_stage_btn.pack_forget()
    except Exception: pass
    w2_back_to_wh_btn.pack(side="left", padx=(0, 6))
    tree_w2_pullouts.delete(*tree_w2_pullouts.get_children())
    df_po2 = load_pullouts_w2()
    # Re-apply active search/filter state so it survives view switches
    try:
        keyword   = w2_search_entry.get().strip().lower()
        shelf_f   = w2_pull_shelf_var.get()
        type_f    = w2_type_filter_var.get()
        date_from = w2_date_from_var.get().strip()
        date_to   = w2_date_to_var.get().strip()
    except (NameError, Exception):
        keyword = shelf_f = type_f = date_from = date_to = ""
    if keyword:
        search_cols = ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
        mask = False
        for col in search_cols:
            if col in df_po2.columns:
                mask = mask | df_po2[col].astype(str).str.lower().str.contains(keyword, na=False)
        df_po2 = df_po2[mask]
    if shelf_f: df_po2 = df_po2[df_po2["Shelf"] == shelf_f]
    if type_f:  df_po2 = df_po2[df_po2["Equipment Type"] == type_f]
    df_po2 = _filter_by_date(df_po2, date_from, date_to)
    w2_pull_row_checks.clear()
    for _, row in df_po2.iterrows():
        iid = tree_w2_pullouts.insert("", "end", values=(
            "☐",
            *tuple(row.get(c, "") for c in ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Shelf", "Status", "Remarks", "Pull Reason", "Date"])
        ))
        w2_pull_row_checks[iid] = False
    try:
        all_reasons = sorted(load_pullouts_w2()["Pull Reason"].dropna().unique().tolist())
        w2_pull_reason_filter_entry["values"] = [""] + all_reasons
    except Exception:
        pass
    active = bool(keyword or shelf_f or type_f or date_from or date_to)
    if active:
        parts = []
        if keyword:   parts.append(f"Search: \"{keyword}\"")
        if shelf_f:   parts.append(f"Shelf: {shelf_f}")
        if type_f:    parts.append(f"Type: {type_f}")
        if date_from: parts.append(f"From: {date_from}")
        if date_to:   parts.append(f"To: {date_to}")
        try:
            w2_search_label.config(text=f"{len(df_po2)} result(s) — " + " | ".join(parts), fg="darkorange")
        except (NameError, Exception):
            pass

def _populate_w2_warehouse_tree(df):
    _show_w2_tree(tree_w2_warehouse)
    tree_w2_warehouse.delete(*tree_w2_warehouse.get_children())
    w2_row_checks.clear()
    for _, row in df.iterrows():
        iid = tree_w2_warehouse.insert("", "end", values=(
            "☐",
            *tuple(row.get(c, "") for c in ["QR", "Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"])
        ))
        w2_row_checks[iid] = False
    _w2_refresh_select_all_label()

def w2_search_item():
    keyword   = w2_search_entry.get().strip().lower()
    shelf_f   = w2_pull_shelf_var.get()
    type_f    = w2_type_filter_var.get()
    date_from = w2_date_from_var.get().strip()
    date_to   = w2_date_to_var.get().strip()

    # ── Pull history view is active: filter it instead of warehouse ──
    if tree_w2_pullouts.winfo_ismapped():
        df = load_pullouts_w2()
        if keyword:
            search_cols = ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
            mask = False
            for col in search_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        if shelf_f: df = df[df["Shelf"] == shelf_f]
        if type_f:  df = df[df["Equipment Type"] == type_f]
        df = _filter_by_date(df, date_from, date_to)
        tree_w2_pullouts.delete(*tree_w2_pullouts.get_children())
        w2_pull_row_checks.clear()
        for _, row in df.iterrows():
            iid = tree_w2_pullouts.insert("", "end", values=(
                "☐",
                *tuple(row.get(c, "") for c in ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Shelf", "Status", "Remarks", "Pull Reason", "Date"])
            ))
            w2_pull_row_checks[iid] = False
        parts = []
        if keyword:   parts.append(f"Search: \"{keyword}\"")
        if shelf_f:   parts.append(f"Shelf: {shelf_f}")
        if type_f:    parts.append(f"Type: {type_f}")
        if date_from: parts.append(f"From: {date_from}")
        if date_to:   parts.append(f"To: {date_to}")
        label = (f"{len(df)} result(s)" + (" — " + " | ".join(parts) if parts else "")) if parts else ""
        w2_search_label.config(text=label, fg="darkorange" if parts else "blue")
        return

    # ── Default: filter warehouse view ───────────────────────────────
    df = load_items_w2()
    if keyword:
        search_cols = ["QR", "Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]
    if shelf_f: df = df[df["Shelf"] == shelf_f]
    if type_f:  df = df[df["Equipment Type"] == type_f]
    df = _filter_by_date(df, date_from, date_to)
    _populate_w2_warehouse_tree(df)
    parts = []
    if keyword:   parts.append(f"Search: \"{keyword}\"")
    if shelf_f:   parts.append(f"Shelf: {shelf_f}")
    if type_f:    parts.append(f"Type: {type_f}")
    if date_from: parts.append(f"From: {date_from}")
    if date_to:   parts.append(f"To: {date_to}")
    label = (f"{len(df)} result(s)" + (" — " + " | ".join(parts) if parts else "")) if parts else ""
    w2_search_label.config(text=label, fg="darkorange" if parts else "blue")

def w2_clear_filters():
    w2_pull_shelf_var.set("")
    w2_type_filter_var.set("")
    w2_search_entry.delete(0, tk.END)
    w2_date_from_var.set("")
    w2_date_to_var.set("")
    w2_pull_reason_filter_var.set("")
    w2_pull_date_from_var.set("")
    w2_pull_date_to_var.set("")
    w2_search_label.config(text="", fg="blue")
    w2_show_warehouse()

def w2_update_full_shelves_display():
    df_shelves = load_shelves_w2()
    full_shelves = df_shelves[df_shelves["Status"] == "FULL"]["Shelf"].tolist()
    w2_full_label.config(text="FULL Shelves:\n" + "\n".join(full_shelves) if full_shelves else "FULL Shelves: None")

def w2_refresh_all():
    w2_show_warehouse()
    update_all_shelf_dropdowns()

# ========== W2 PULL OUT ==========

def w2_pull_search_live(event=None):
    """Filter whichever W2 view is currently active based on the search box."""
    keyword = w2_pull_item_entry.get().strip().lower()

    if tree_w2_available.winfo_ismapped():
        # Shelf status view is active
        df = load_shelves_w2().sort_values("Shelf")
        df_items_all = load_items_w2()
        if keyword:
            mask = False
            for col in ["Shelf", "Status", "Date_Full"]:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        tree_w2_available.delete(*tree_w2_available.get_children())
        for _, row in df.iterrows():
            date_full  = row.get("Date_Full", "")
            shelf_name = row["Shelf"]
            item_count = int((df_items_all["Shelf"] == shelf_name).sum()) if "Shelf" in df_items_all.columns else 0
            tree_w2_available.insert("", "end", values=(shelf_name, row["Status"], item_count, date_full if pd.notna(date_full) else ""))
        w2_search_label.config(text=f"{len(df)} match(es)" if keyword else "", fg="blue")

    elif tree_w2_pullouts.winfo_ismapped():
        # Pull history view is active — respect all active filters
        df = load_pullouts_w2()
        shelf_f   = w2_pull_shelf_var.get()
        type_f    = w2_type_filter_var.get()
        date_from = w2_date_from_var.get().strip()
        date_to   = w2_date_to_var.get().strip()
        if keyword:
            search_cols = ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
            mask = False
            for col in search_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        if shelf_f: df = df[df["Shelf"] == shelf_f]
        if type_f:  df = df[df["Equipment Type"] == type_f]
        df = _filter_by_date(df, date_from, date_to)
        tree_w2_pullouts.delete(*tree_w2_pullouts.get_children())
        w2_pull_row_checks.clear()
        for _, row in df.iterrows():
            iid = tree_w2_pullouts.insert("", "end", values=(
                "☐",
                *tuple(row.get(c, "") for c in ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Shelf", "Status", "Remarks", "Pull Reason", "Date"])
            ))
            w2_pull_row_checks[iid] = False
        active = bool(keyword or shelf_f or type_f or date_from or date_to)
        w2_search_label.config(text=f"{len(df)} match(es)" if active else "", fg="darkorange" if active else "blue")

    else:
        # Warehouse view (default)
        w2_show_warehouse()
        if keyword:
            df = load_items_w2()
            search_cols = ["QR", "Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
            mask = False
            for col in search_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
            _populate_w2_warehouse_tree(df)
            w2_search_label.config(text=f"{len(df)} match(es)", fg="blue")
        else:
            w2_search_label.config(text="", fg="blue")

def w2_select_item(event):
    selected = tree_w2_warehouse.selection()
    if selected:
        iid = selected[0]
        values = tree_w2_warehouse.item(iid, "values")
        # values: ☐(0), QR(1), Set ID(2), Hostname(3), Equipment Type(4), Serial(5), Checked By(6), Shelf(7), Status(8), Remarks(9), Date(10)

        # Toggle checkbox on single-click
        if iid in w2_row_checks:
            w2_row_checks[iid] = not w2_row_checks[iid]
            tree_w2_warehouse.set(iid, "C0", "☑" if w2_row_checks[iid] else "☐")
            _w2_refresh_select_all_label()

        w2_pull_item_entry.delete(0, tk.END)
        w2_pull_item_entry.insert(0, f"{values[2]} - {values[4]}")  # SET-001 - Monitor
        w2_status_label.config(
            text=f"Selected → {values[2]} ({values[4]})  |  Hostname: {values[3]}  |  Shelf: {values[7]}  |  Serial: {values[5]}",
            fg="#1a5276")

def w2_unstage_from_warehouse(event=None):
    # Collect checked rows; fall back to treeview selection if nothing checked
    checked = [iid for iid, state in w2_row_checks.items() if state]
    if not checked:
        sel = tree_w2_warehouse.selection()
        if sel:
            checked = [sel[0]]
    if not checked:
        messagebox.showinfo("Back to Stage", "Check at least one row in the Warehouse table.")
        return

    moved = 0
    for item_id in checked:
        values = tree_w2_warehouse.item(item_id, "values")
        if not values:
            continue
        # values: ☐(0), QR(1), Set ID(2), Hostname(3), Equipment Type(4), Serial(5), Checked By(6), Shelf(7), Status(8), Remarks(9), Date(10)
        set_id, hostname, eq_type, serial, checked_by, shelf, status, remarks = values[2], values[3], values[4], values[5], values[6], values[7], values[8], values[9]
        if not messagebox.askyesno("Move to Staging",
            f"Move {eq_type} ({set_id}) back to staging?\n\nShelf: {shelf}"):
            continue

        df_w2 = load_items_w2()
        match = df_w2[(df_w2["Set ID"] == set_id) & (df_w2["Equipment Type"] == eq_type)]
        if match.empty:
            messagebox.showerror("Error", "Item not found in warehouse")
            continue

        qr_label = f"{set_id}-{eq_type}"
        delete_qr(qr_label, warehouse=2)
        df_w2 = df_w2.drop(match.index).reset_index(drop=True)
        save_warehouse_2(df_w2, load_shelves_w2())

        staged_sets.append({"set_id": set_id, "items": [{
            "Equipment Type": eq_type,
            "Hostname":       hostname,
            "Serial Number":  serial,
            "Checked By":     checked_by,
            "Shelf":          shelf,
            "Status":         status,
            "Remarks":        remarks,
        }]})
        save_log("UNSTAGE", f"[W2] Set: {set_id} | Item: {eq_type} | Shelf: {shelf}")
        moved += 1

    if moved:
        messagebox.showinfo("Moved", f"{moved} item(s) moved back to staging.")
        update_w2_staged_display()
        w2_pull_item_entry.delete(0, tk.END)
        w2_search_label.config(text="", fg="blue")
        w2_status_label.config(text="")
        w2_show_warehouse()
        update_all_shelf_dropdowns()

def w2_pull_item():
    selection_text = w2_pull_item_entry.get().strip()
    reason = w2_pull_reason_filter_var.get().strip()
    if not selection_text:
        messagebox.showerror("Error", "No item selected for pull out"); return
    if not reason:
        messagebox.showerror("Error", "Please enter a pull reason"); return

    # Parse "SET-001 - Monitor"
    try:
        set_id, eq_type = [x.strip() for x in selection_text.split(" - ", 1)]
    except ValueError:
        messagebox.showerror("Error", "Invalid selection format"); return

    df_w2 = load_items_w2()
    df_po2 = load_pullouts_w2()

    match = df_w2[(df_w2["Set ID"] == set_id) & (df_w2["Equipment Type"] == eq_type)]
    if match.empty:
        messagebox.showerror("Error", f"'{selection_text}' not found in Warehouse 2"); return

    if not messagebox.askyesno("Confirm Pull Out",
        f"Pull out {eq_type} from {set_id}?\nReason: {reason}"):
        return

    item_row = match.iloc[0]
    qr_label = f"{set_id}-{eq_type}"
    delete_qr(qr_label, warehouse=2)

    df_w2 = df_w2.drop(match.index).reset_index(drop=True)
    df_po2 = pd.concat([df_po2, pd.DataFrame([{
        "Set ID": set_id,
        "Hostname": str(item_row.get("Hostname", "")),
        "Equipment Type": eq_type,
        "Serial Number": str(item_row.get("Serial Number", "")),
        "Checked By": str(item_row.get("Checked By", "")),
        "Shelf": str(item_row.get("Shelf", "")),
        "Status": str(item_row.get("Status", "")),
        "Remarks": str(item_row.get("Remarks", "")),
        "Pull Reason": reason,
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }])], ignore_index=True)

    save_warehouse_2(df_w2, load_shelves_w2(), df_po2)
    save_log("WAREHOUSE PULL", f"[W2] Set: {set_id} | Item: {eq_type} | Reason: {reason}")
    messagebox.showinfo("Success", f"{eq_type} from {set_id} pulled out successfully")
    w2_pull_item_entry.delete(0, tk.END)
    w2_pull_reason_filter_var.set("")
    w2_refresh_all()

def w2_undo_pull(event=None):
    # Collect checked rows; fall back to treeview selection if nothing checked
    checked = [iid for iid, state in w2_pull_row_checks.items() if state]
    if not checked:
        sel = tree_w2_pullouts.selection()
        if sel:
            checked = [sel[0]]
    if not checked:
        messagebox.showinfo("Back to Warehouse", "Check at least one row in the Pull History table.")
        return

    restored = 0
    for item_id in checked:
        values = tree_w2_pullouts.item(item_id, "values")
        if not values:
            continue
        # values: ☐(0), Set ID(1), Hostname(2), Equip Type(3), Serial(4), Shelf(5), Status(6), Remarks(7), PullReason(8), Date(9)
        set_id, hostname, eq_type, shelf, status, remarks = values[1], values[2], values[3], values[5], values[6], values[7]
        if not messagebox.askyesno("Undo Pull",
            f"Restore {eq_type} ({set_id}) back to Warehouse 2?\nShelf: {shelf}"):
            continue

        df_w2 = load_items_w2()
        df_po2 = load_pullouts_w2()

        match = df_po2[(df_po2["Set ID"] == set_id) & (df_po2["Equipment Type"] == eq_type)]
        if match.empty:
            messagebox.showerror("Error", "Record not found in pull history")
            continue

        pull_row = match.iloc[0]
        qr_label = f"{set_id}-{eq_type}"
        qr_code = str(uuid.uuid4())
        try:
            generate_qr(qr_label, qr_code, warehouse=2)
        except Exception as e:
            messagebox.showwarning("Warning", f"QR not regenerated: {e}")

        df_w2 = pd.concat([df_w2, pd.DataFrame([{
            "QR":             qr_code,
            "Set ID":         set_id,
            "Hostname":       str(pull_row.get("Hostname", "")),
            "Equipment Type": eq_type,
            "Serial Number":  str(pull_row.get("Serial Number", "")),
            "Checked By":     str(pull_row.get("Checked By", "")),
            "Shelf":          shelf,
            "Status":         status,
            "Remarks":        remarks,
            "Date":           datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }])], ignore_index=True)

        df_po2 = df_po2.drop(match.index).reset_index(drop=True)
        save_warehouse_2(df_w2, load_shelves_w2(), df_po2)
        save_log("UNDO PULL", f"[W2] Set: {set_id} | Item: {eq_type} | Shelf: {shelf}")
        restored += 1

    if restored:
        messagebox.showinfo("Restored", f"{restored} item(s) restored to Warehouse 2.")
    w2_show_pullouts()

# ========== W2 ITEM MANAGEMENT ==========

def w2_update_item():
    """Update a staged set item OR a warehouse item, depending on what is selected."""

    # ── Branch A: edit a STAGED set ──────────────────────────────────────────
    staged_sel = w2_staged_listbox.curselection()
    if staged_sel:
        s_idx = staged_sel[0]
        if s_idx >= len(staged_sets):
            return
        staged_set = staged_sets[s_idx]
        set_id     = staged_set["set_id"]
        items      = staged_set["items"]
        shelf_list = sorted(load_shelves_w2()["Shelf"].tolist())

        edit_win = tk.Toplevel(root)
        edit_win.title(f"Edit Staged Set — {set_id}")
        edit_win.resizable(False, False)
        edit_win.transient(root)

        tk.Label(edit_win, text=f"Edit Staged Set: {set_id}",
                 font=("Helvetica", 10, "bold"), bg="#27ae60", fg="white",
                 padx=10, pady=6).grid(row=0, column=0, columnspan=2, sticky="we")

        COL_WIDTHS = [12, 18, 18, 16, 16, 13, 20]
        HEADERS    = ["Equipment", "Hostname", "Serial Number",
                      "Checked By", "Shelf", "Status", "Remarks"]

        outer = tk.Frame(edit_win, padx=10, pady=5)
        outer.grid(row=1, column=0, columnspan=2)

        hdr_frame = tk.Frame(outer, bg="#dce3f0")
        hdr_frame.pack(fill="x")
        for col, (h, cw) in enumerate(zip(HEADERS, COL_WIDTHS)):
            tk.Label(hdr_frame, text=h, font=("Helvetica", 9, "bold"), width=cw, anchor="w",
                     bg="#dce3f0", padx=5).grid(row=0, column=col, padx=5, pady=4, sticky="w")

        ROW_COLORS = ("#ffffff", "#f0f4ff")
        row_widgets = {}
        for r, item in enumerate(items):
            bg = ROW_COLORS[r % 2]
            eq_type = item["Equipment Type"]
            row_bg  = tk.Frame(outer, bg=bg, bd=1, relief="flat")
            row_bg.pack(fill="x", pady=1)
            tk.Label(row_bg, text=eq_type, width=COL_WIDTHS[0], anchor="w",
                     bg=bg, font=("Helvetica", 9, "bold")).grid(row=0, column=0, padx=5, pady=8, sticky="w")
            hostname_e = tk.Entry(row_bg, width=COL_WIDTHS[1], font=("Helvetica", 9))
            hostname_e.insert(0, item.get("Hostname", ""))
            hostname_e.grid(row=0, column=1, padx=5, pady=8)
            serial_e = tk.Entry(row_bg, width=COL_WIDTHS[2], font=("Helvetica", 9))
            serial_e.insert(0, item.get("Serial Number", ""))
            serial_e.grid(row=0, column=2, padx=5, pady=8)
            checked_e = tk.Entry(row_bg, width=COL_WIDTHS[3], font=("Helvetica", 9))
            checked_e.insert(0, item.get("Checked By", ""))
            checked_e.grid(row=0, column=3, padx=5, pady=8)
            shelf_v = tk.StringVar(value=item.get("Shelf", ""))
            ttk.Combobox(row_bg, textvariable=shelf_v, values=shelf_list,
                         width=COL_WIDTHS[4], state="readonly",
                         font=("Helvetica", 9)).grid(row=0, column=4, padx=5, pady=8)
            status_v = tk.StringVar(value=item.get("Status", ""))
            ttk.Combobox(row_bg, textvariable=status_v,
                         values=STATUS_CHOICES,
                         width=COL_WIDTHS[5], state="readonly",
                         font=("Helvetica", 9)).grid(row=0, column=5, padx=5, pady=8)
            remarks_e = tk.Entry(row_bg, width=COL_WIDTHS[6], font=("Helvetica", 9))
            remarks_e.insert(0, item.get("Remarks", ""))
            remarks_e.grid(row=0, column=6, padx=5, pady=8)
            row_widgets[eq_type] = (hostname_e, serial_e, checked_e,
                                    shelf_v, status_v, remarks_e)

        error_lbl = tk.Label(edit_win, text="", fg="red", font=("Helvetica", 8))
        error_lbl.grid(row=2, column=0, columnspan=2, pady=(4, 0))

        def save_staged_update():
            df_items_w2      = load_items_w2()
            existing_serials = df_items_w2["Serial Number"].astype(str).tolist() \
                               if "Serial Number" in df_items_w2.columns else []
            # serials from OTHER staged sets
            other_staged_serials = [
                it["Serial Number"]
                for si, ss in enumerate(staged_sets) if si != s_idx
                for it in ss["items"] if it.get("Serial Number")
            ]
            seen_in_dialog = []
            updated_items  = []
            for eq_type, (hostname_e, serial_e, checked_e,
                          shelf_v, status_v, remarks_e) in row_widgets.items():
                hn  = hostname_e.get().strip()
                sn  = serial_e.get().strip()
                cb  = checked_e.get().strip()
                sh  = shelf_v.get().strip()
                st  = status_v.get().strip()
                rm  = remarks_e.get().strip()
                if not hn:  error_lbl.config(text=f"Hostname required for {eq_type}");       return
                if not sn:  error_lbl.config(text=f"Serial Number required for {eq_type}");  return
                if not cb:  error_lbl.config(text=f"Checked By required for {eq_type}");     return
                if not sh:  error_lbl.config(text=f"Please select a Shelf for {eq_type}");   return
                if not st:  error_lbl.config(text=f"Please select a Status for {eq_type}");  return
                # Find original serial for this item to allow keeping same value
                orig_sn = next((it["Serial Number"] for it in items
                                if it["Equipment Type"] == eq_type), None)
                if sn != orig_sn:
                    if sn in existing_serials:
                        error_lbl.config(text=f"Serial '{sn}' already in Warehouse 2"); return
                    if sn in other_staged_serials:
                        error_lbl.config(text=f"Serial '{sn}' in another staged set"); return
                if sn in seen_in_dialog:
                    error_lbl.config(text=f"Duplicate serial '{sn}' within this set"); return
                seen_in_dialog.append(sn)
                updated_items.append({
                    "Equipment Type": eq_type, "Hostname": hn,
                    "Serial Number": sn, "Checked By": cb,
                    "Shelf": sh, "Status": st, "Remarks": rm,
                })
            staged_sets[s_idx]["items"] = updated_items
            edit_win.destroy()
            update_w2_staged_display()
            messagebox.showinfo("Updated", f"{set_id} staging updated successfully")

        btn_f = tk.Frame(edit_win)
        btn_f.grid(row=3, column=0, columnspan=2, pady=10)
        tk.Button(btn_f, text="SAVE UPDATE", command=save_staged_update, width=14,
                  bg="#27ae60", fg="white").pack(side="left", padx=6)
        tk.Button(btn_f, text="Cancel", command=edit_win.destroy, width=10).pack(side="left", padx=6)

        edit_win.update_idletasks()
        edit_win.geometry(f"{edit_win.winfo_reqwidth()}x{edit_win.winfo_reqheight()}")
        edit_win.grab_set()
        edit_win.focus_force()
        return

    # ── Branch B: warehouse-table update is intentionally disabled ────────────
    messagebox.showerror(
        "Update Not Allowed",
        "Items cannot be updated directly from the warehouse.\n\n"
        "Double-click the row to move it back to staging,\n"
        "then select it in the staged sets list and click UPDATE ITEM.")

# ========== W2 SHELF MANAGEMENT ==========

def w2_set_shelf_status(new_status):
    shelf = w2_shelf_control_var.get()
    if not shelf:
        messagebox.showerror("Error", "Select a shelf"); return
    df_items_w2 = load_items_w2()
    df_shelves_w2 = load_shelves_w2()
    idx = df_shelves_w2[df_shelves_w2["Shelf"] == shelf].index
    if len(idx) == 0:
        return
    df_shelves_w2.at[idx[0], "Status"] = new_status
    df_shelves_w2.at[idx[0], "Date_Full"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if new_status == "FULL" else None
    save_warehouse_2(df_items_w2, df_shelves_w2)
    save_log("SHELF STATUS", f"[W2] Shelf: {shelf} → {new_status}")
    w2_status_label.config(text=f"{shelf} → {new_status}")
    w2_refresh_all()

def w2_add_shelf():
    new_shelf = w2_remove_shelf_var.get().strip()
    if not new_shelf:
        messagebox.showerror("Error", "Enter shelf name"); return
    df_shelves_w2 = load_shelves_w2()
    if new_shelf in df_shelves_w2["Shelf"].values:
        messagebox.showerror("Error", "Shelf already exists in W2"); return
    df_shelves_w2 = pd.concat([df_shelves_w2, pd.DataFrame([{"Shelf": new_shelf, "Status": "AVAILABLE"}])], ignore_index=True)
    df_shelves_w2 = df_shelves_w2.sort_values("Shelf", ignore_index=True)
    save_warehouse_2(load_items_w2(), df_shelves_w2)
    messagebox.showinfo("Success", f"Shelf '{new_shelf}' added to Warehouse 2")
    w2_remove_shelf_var.set("")
    update_all_shelf_dropdowns()

def w2_remove_shelf():
    shelf_name = w2_remove_shelf_var.get().strip()
    if not shelf_name:
        messagebox.showerror("Error", "Select a shelf to remove"); return
    df_items_w2 = load_items_w2()
    df_shelves_w2 = load_shelves_w2()
    if not df_items_w2[df_items_w2["Shelf"] == shelf_name].empty:
        messagebox.showerror("Error", f"Cannot remove shelf '{shelf_name}' — it still has items"); return
    if shelf_name not in df_shelves_w2["Shelf"].values:
        messagebox.showerror("Error", f"Shelf '{shelf_name}' does not exist in W2"); return
    df_shelves_w2 = df_shelves_w2[df_shelves_w2["Shelf"] != shelf_name].sort_values("Shelf", ignore_index=True)
    save_warehouse_2(df_items_w2, df_shelves_w2)
    messagebox.showinfo("Success", f"Shelf '{shelf_name}' removed from Warehouse 2")
    w2_remove_shelf_var.set("")
    update_all_shelf_dropdowns()

def w2_reset_shelf_control():
    w2_shelf_control_var.set("")
    w2_status_label.config(text="")

def w2_reset_shelf_addition():
    w2_remove_shelf_var.set("")
    w2_status_label.config(text="")

def w2_reset_pull_out():
    w2_pull_item_entry.delete(0, tk.END)
    w2_pull_reason_filter_var.set("")
    w2_pull_date_from_var.set("")
    w2_pull_date_to_var.set("")
    w2_status_label.config(text="")
    w2_search_label.config(text="")
    w2_show_warehouse()

def w2_filter_pull_history():
    """Filter and display the W2 pull history table only. Completely separate from warehouse search."""
    reason = w2_pull_reason_filter_var.get().strip().lower()
    date_from = w2_pull_date_from_var.get().strip()
    date_to = w2_pull_date_to_var.get().strip()
    w2_show_pullouts()  # always switch to pull history view first
    if not reason and not date_from and not date_to:
        return
    tree_w2_pullouts.delete(*tree_w2_pullouts.get_children())
    df = load_pullouts_w2()
    if reason:
        search_cols = ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(reason, na=False)
        df = df[mask]
    df = _filter_by_date(df, date_from, date_to)
    # Populate the pull reason dropdown with known values from data
    all_reasons = sorted(load_pullouts_w2()["Pull Reason"].dropna().unique().tolist())
    w2_pull_reason_filter_entry["values"] = [""] + all_reasons
    w2_pull_row_checks.clear()
    for _, row in df.iterrows():
        iid = tree_w2_pullouts.insert("", "end", values=(
            "☐",
            *tuple(row.get(c, "") for c in ["Set ID", "Hostname", "Equipment Type", "Serial Number", "Shelf", "Status", "Remarks", "Pull Reason", "Date"])
        ))
        w2_pull_row_checks[iid] = False
    w2_search_label.config(text=f"Pull History Filtered: {len(df)} result(s)", fg="darkorange")

# ========== QR LABEL PDF ==========

def _lock_pdf(pdf_path):
    """Restrict PDF to view-only: tries pikepdf encryption first, always falls back to OS read-only."""
    tmp_lock = pdf_path + ".~lock.pdf"
    try:
        import pikepdf
        permissions = pikepdf.Permissions(
            accessibility=True,
            extract=True,
            modify_annotation=False,
            modify_assembly=False,
            modify_form=False,
            modify_other=False,
            print_highres=True,
            print_lowres=True,
        )
        with pikepdf.open(pdf_path) as pdf_obj:
            pdf_obj.save(
                tmp_lock,
                encryption=pikepdf.Encryption(
                    owner="WMS_READONLY_2026",
                    user="",
                    R=4,
                    allow=permissions,
                )
            )
        os.replace(tmp_lock, pdf_path)
    except ImportError:
        pass
    except Exception:
        try:
            os.remove(tmp_lock)
        except Exception:
            pass
    # Always apply OS-level read-only regardless of pikepdf outcome
    try:
        import stat
        os.chmod(pdf_path, stat.S_IREAD | stat.S_IRGRP | stat.S_IROTH)
    except Exception:
        pass


def generate_qr_pdf(items_batch, custom_name=None):
    import json
    from fpdf import FPDF
    PAGE_W, LABEL_W, COLS = 210, 54, 3
    MARGIN_X, MARGIN_Y, ROW_GAP = 12, 10, 4
    GAP_X = (PAGE_W - (COLS * LABEL_W) - (2 * MARGIN_X)) / (COLS - 1)
    QR_SIZE, QR_PAD_TOP, QR_PAD_BTM = 20, 3, 3
    LINE_H, FIELD_PAD_BTM = 4.8, 2

    if not items_batch:
        return None

    # ── Resolve output path first (needed for sidecar lookup) ──
    is_w2_batch = items_batch[0].get('_warehouse', 1) == 2
    if is_w2_batch:
        output_folder = QR_LABELS_FOLDER_W2
    else:
        output_folder = QR_LABELS_FOLDER_W1

    os.makedirs(output_folder, exist_ok=True)

    if custom_name:
        safe_name = custom_name.replace(" ", "_").replace("/", "-").replace("\\", "-")
        if safe_name.lower().endswith(".pdf"):
            safe_name = safe_name[:-4]
    else:
        if is_w2_batch:
            set_ids = list(dict.fromkeys(item.get("Set ID", "") for item in items_batch))
            label_name = "_".join(set_ids[:3])
            if len(set_ids) > 3:
                label_name += f"_and_{len(set_ids)-3}_more"
        else:
            existing = [f for f in os.listdir(output_folder) if f.startswith("BATCH_") and f.endswith(".pdf")]
            batch_nums = []
            for f in existing:
                try:
                    batch_nums.append(int(f.split("_")[1]))
                except (IndexError, ValueError):
                    pass
            label_name = f"BATCH_{max(batch_nums, default=0) + 1}"
        safe_name = label_name.replace(" ", "_").replace("/", "-")

    pdf_path   = os.path.join(output_folder, f"{safe_name}.pdf")
    index_path = pdf_path + ".keys.json"

    # ── Load existing sidecar (item data + keys) ──
    existing_items = []
    existing_keys  = set()
    if os.path.exists(index_path) and os.path.exists(pdf_path):
        try:
            with open(index_path, "r") as kf:
                sidecar = json.load(kf)
            existing_items = sidecar.get("items", [])
            existing_keys  = set(sidecar.get("keys",  []))
        except Exception:
            existing_items = []
            existing_keys  = set()

    def _item_key(item):
        if item.get('_warehouse', 1) == 2:
            return f"{item.get('Set ID', '')}-{item.get('Equipment Type', '')}"
        return str(item.get('Hostname', ''))

    # Only keep truly new items
    new_items = [it for it in items_batch if _item_key(it) not in existing_keys]
    if not new_items:
        return pdf_path  # nothing to add

    # ── Combine: existing items first, then new ones ──
    all_items = existing_items + new_items

    # ── Lift OS read-only before writing (re-applied after via _lock_pdf) ──
    if os.path.exists(pdf_path):
        try:
            import stat
            os.chmod(pdf_path, stat.S_IWRITE | stat.S_IREAD)
        except Exception:
            pass

    # ── Render all items into a fresh PDF so grid is always packed ──
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    col = row = 0

    for item in all_items:
        is_w2 = item.get('_warehouse', 1) == 2

        if is_w2:
            fields = [
                ("Set ID:",    str(item.get("Set ID", ""))),
                ("Type:",      str(item.get("Equipment Type", ""))),
                ("Serial No:", str(item.get("Serial Number", ""))),
                ("Shelf:",     str(item.get("Shelf", ""))),
                ("Status:",    str(item.get("Status", ""))),
                ("Remarks:",   str(item.get("Remarks", ""))),
                ("Date:",      str(item.get("_date", datetime.now().strftime("%Y-%m-%d")))),
            ]
            qr_key = f"{item.get('Set ID', '')}-{item.get('Equipment Type', '')}"
            path   = qr_path_for(qr_key, warehouse=2)
        else:
            fields = [
                ("Hostname:",   str(item.get("Hostname", ""))),
                ("Serial No:",  str(item.get("Serial Number", ""))),
                ("Checked By:", str(item.get("Checked By", ""))),
                ("Shelf:",      str(item.get("Shelf", ""))),
                ("Status:",     str(item.get("Status", ""))),
                ("Remarks:",    str(item.get("Remarks", ""))),
                ("Date:",       str(item.get("_date", datetime.now().strftime("%Y-%m-%d")))),
            ]
            path = qr_path_for(item.get('Hostname', ''), warehouse=1)

        LABEL_H = QR_PAD_TOP + QR_SIZE + QR_PAD_BTM + len(fields) * LINE_H + FIELD_PAD_BTM

        x = MARGIN_X + col * (LABEL_W + GAP_X)
        y = MARGIN_Y + row * (LABEL_H + ROW_GAP)
        if y + LABEL_H > 297 - MARGIN_Y:
            pdf.add_page(); col = row = 0
            x = MARGIN_X
            y = MARGIN_Y

        pdf.set_draw_color(150, 150, 150)
        pdf.rect(x, y, LABEL_W, LABEL_H)

        if os.path.exists(path):
            pdf.image(path, x=x + (LABEL_W - QR_SIZE) / 2, y=y + QR_PAD_TOP, w=QR_SIZE, h=QR_SIZE)

        label_x = x + 2
        value_x = x + 19
        text_y  = y + QR_PAD_TOP + QR_SIZE + QR_PAD_BTM
        for lbl, val in fields:
            pdf.set_font("Helvetica", style="B", size=5.5)
            pdf.set_xy(label_x, text_y); pdf.cell(17, LINE_H, lbl, ln=0)
            pdf.set_font("Helvetica", size=5.5)
            pdf.set_xy(value_x, text_y); pdf.cell(LABEL_W - 21, LINE_H, val[:22], ln=0)
            text_y += LINE_H

        col += 1
        if col >= COLS:
            col = 0; row += 1

    pdf.output(pdf_path)

    # ── Stamp date on new items before saving to sidecar ──
    today = datetime.now().strftime("%Y-%m-%d")
    for it in new_items:
        it.setdefault("_date", today)

    # ── Update sidecar with all item data + keys ──
    all_keys = list(existing_keys | {_item_key(it) for it in new_items})
    try:
        with open(index_path, "w") as kf:
            json.dump({"keys": all_keys, "items": all_items}, kf)
    except Exception:
        pass

    _lock_pdf(pdf_path)
    try:
        import stat
        os.chmod(pdf_path, stat.S_IREAD | stat.S_IRGRP | stat.S_IROTH)
    except Exception:
        pass
    return pdf_path

# ========== DIALOGS ==========

def open_label_manager(warehouse=None):
    manager = tk.Toplevel(root)
    manager.title(f"QR Label Manager — Warehouse {warehouse}" if warehouse else "QR Label Manager")
    manager.geometry("700x500")
    manager.resizable(False, False)

    # ── Header ──────────────────────────────────────────────
    hdr = tk.Frame(manager, bg="#2c3e50")
    hdr.pack(fill="x")
    tk.Label(hdr, text="QR Label Files", font=("Helvetica", 10, "bold"),
             bg="#2c3e50", fg="white").pack(side="left", padx=10, pady=6)
    sel_count_lbl = tk.Label(hdr, text="", font=("Helvetica", 9),
                              bg="#2c3e50", fg="#f0b429")
    sel_count_lbl.pack(side="right", padx=10)

    # ── Checklist canvas area ────────────────────────────────
    list_frame = tk.Frame(manager, bd=1, relief="sunken")
    list_frame.pack(fill="both", expand=True, padx=10, pady=(8, 4))

    canvas_cl = tk.Canvas(list_frame, bg="white", highlightthickness=0)
    sb_cl = ttk.Scrollbar(list_frame, orient="vertical", command=canvas_cl.yview)
    canvas_cl.configure(yscrollcommand=sb_cl.set)
    sb_cl.pack(side="right", fill="y")
    canvas_cl.pack(side="left", fill="both", expand=True)
    canvas_cl.bind("<MouseWheel>", lambda e: canvas_cl.yview_scroll(int(-1*(e.delta/120)), "units"))

    inner_cl = tk.Frame(canvas_cl, bg="white")
    cw_id = canvas_cl.create_window((0, 0), window=inner_cl, anchor="nw")
    canvas_cl.bind("<Configure>", lambda e: canvas_cl.itemconfigure(cw_id, width=e.width))

    # ── State ────────────────────────────────────────────────
    row_data    = []   # list of (iid, full_path, wh_label, filename, date_str, size_str, var)
    check_vars  = {}   # iid -> BooleanVar

    def _refresh_sel_count():
        n = sum(1 for v in check_vars.values() if v.get())
        sel_count_lbl.config(text=f"{n} selected" if n else "")
        clear_btn.config(state="normal" if n else "disabled")

    def _toggle_all():
        want = not all(v.get() for v in check_vars.values())
        for v in check_vars.values():
            v.set(want)
        _refresh_sel_count()
        _repaint_rows()

    def _repaint_rows():
        for iid, _, _, _, _, _, var in row_data:
            try:
                fr = inner_cl.nametowidget(f"row_{iid}")
                fr.config(bg="#e8f4fd" if var.get() else ("white" if row_data.index(
                    next(r for r in row_data if r[0] == iid)) % 2 == 0 else "#f7f9fc"))
            except Exception:
                pass

    EVEN_BG, ODD_BG, SEL_BG = "white", "#f7f9fc", "#e8f4fd"

    def load_label_files():
        # clear
        for w in inner_cl.winfo_children():
            w.destroy()
        row_data.clear()
        check_vars.clear()

        now = datetime.now()
        all_warehouses = [("Warehouse 1", QR_LABELS_FOLDER_W1), ("Warehouse 2", QR_LABELS_FOLDER_W2)]
        filtered = [(wl, f) for wl, f in all_warehouses if warehouse is None or wl == f"Warehouse {warehouse}"]

        # Column header row
        hdr_row = tk.Frame(inner_cl, bg="#dce3f0")
        hdr_row.pack(fill="x")
        tk.Label(hdr_row, text="✔", width=3, bg="#dce3f0", font=("Helvetica", 9, "bold")).pack(side="left", padx=(6,0))
        for txt, w in [("Warehouse", 110), ("Filename", 240), ("Created", 155), ("Size", 60)]:
            tk.Label(hdr_row, text=txt, width=w//7, bg="#dce3f0",
                     font=("Helvetica", 9, "bold"), anchor="w").pack(side="left", padx=4, pady=5)

        idx = 0
        for warehouse_label, folder in filtered:
            if not os.path.exists(folder):
                continue
            for f in sorted([fn for fn in os.listdir(folder) if fn.endswith(".pdf")], reverse=True):
                full_path = os.path.join(folder, f)
                size_kb = round(os.path.getsize(full_path) / 1024, 1)
                try:
                    mtime = os.path.getmtime(full_path)
                    file_dt = datetime.fromtimestamp(mtime)
                    delta = now - file_dt
                    age = "Today" if delta.days == 0 else ("1 day ago" if delta.days == 1 else f"{delta.days} days ago")
                    date_str = f"{file_dt.strftime('%Y-%m-%d')}  ({age})"
                except Exception:
                    date_str = "Unknown"

                iid = f"row{idx}"
                var = tk.BooleanVar(value=False)
                check_vars[iid] = var
                bg = EVEN_BG if idx % 2 == 0 else ODD_BG

                row_fr = tk.Frame(inner_cl, bg=bg, name=f"row_{iid}", cursor="hand2")
                row_fr.pack(fill="x")

                def _make_toggle(v=var):
                    def _toggle(e=None):
                        v.set(not v.get())
                        _refresh_sel_count()
                        _repaint_rows()
                    return _toggle

                cb = tk.Checkbutton(row_fr, variable=var, bg=bg,
                                    command=lambda: [_refresh_sel_count(), _repaint_rows()])
                cb.pack(side="left", padx=(6, 0), pady=4)
                for txt, w in [(warehouse_label, 16), (f, 34), (date_str, 22), (f"{size_kb} kb", 8)]:
                    lbl = tk.Label(row_fr, text=txt, bg=bg, anchor="w", width=w,
                                   font=("Helvetica", 9))
                    lbl.pack(side="left", padx=4, pady=4)
                    lbl.bind("<Button-1>", _make_toggle(var))
                row_fr.bind("<Button-1>", _make_toggle(var))

                row_data.append((iid, full_path, warehouse_label, f, date_str, f"{size_kb} kb", var))
                idx += 1

        if idx == 0:
            tk.Label(inner_cl, text="No QR label PDF files found.", fg="gray",
                     font=("Helvetica", 10), bg="white").pack(pady=30)

        inner_cl.update_idletasks()
        canvas_cl.configure(scrollregion=canvas_cl.bbox("all"))
        _refresh_sel_count()

    def open_selected():
        chosen = [rd for rd in row_data if rd[6].get()]
        if not chosen:
            messagebox.showwarning("Warning", "Check a file to open.", parent=manager); return
        if len(chosen) > 1:
            messagebox.showerror("Error", "Only 1 file can be opened at a time.\nPlease check only one file.", parent=manager); return
        full_path = chosen[0][1]
        if os.path.exists(full_path):
            os.startfile(full_path)
        else:
            messagebox.showerror("Error", "File not found.", parent=manager)

    def clear_selected():
        chosen = [rd for rd in row_data if rd[6].get()]
        if not chosen:
            messagebox.showwarning("Warning", "Check at least one file to clear.", parent=manager); return
        count = len(chosen)
        prompt = (f"Delete '{chosen[0][3]}'?" if count == 1
                  else f"Delete {count} checked file(s)?\nThis cannot be undone.")
        if not messagebox.askyesno("Confirm Clear", prompt, parent=manager):
            return
        failed = []
        import stat
        for _, full_path, _, fname, _, _, _ in chosen:
            try:
                if os.path.exists(full_path):
                    os.chmod(full_path, stat.S_IWRITE | stat.S_IREAD)
                    os.remove(full_path)
                    sidecar = full_path + ".keys.json"
                    if os.path.exists(sidecar):
                        os.remove(sidecar)
            except Exception as e:
                failed.append(f"{fname}: {e}")
        load_label_files()
        if failed:
            messagebox.showerror("Error", "Some files could not be deleted:\n" + "\n".join(failed), parent=manager)
        else:
            messagebox.showinfo("Cleared", f"{count} file(s) deleted.", parent=manager)

    # ── Bottom toolbar ───────────────────────────────────────
    btn_frame_m = tk.Frame(manager)
    btn_frame_m.pack(pady=8)
    tk.Button(btn_frame_m, text="☑", command=_toggle_all,   width=4).pack(side="left", padx=4)
    tk.Button(btn_frame_m, text="OPEN", command=open_selected, width=10).pack(side="left", padx=4)
    clear_btn = tk.Button(btn_frame_m, text="✕", command=clear_selected,
                           width=4, bg="#922b21", fg="white", state="disabled")
    clear_btn.pack(side="left", padx=4)

    load_label_files()

def open_activity_log():
    log_win = tk.Toplevel(root)
    log_win.title("Activity Log")
    log_win.geometry("820x500")

    filter_frame = tk.Frame(log_win)
    filter_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(filter_frame, text="Filter Action:").pack(side="left", padx=(0, 2))
    filter_action_var = tk.StringVar()
    filter_action_cb = ttk.Combobox(filter_frame, textvariable=filter_action_var, state="readonly", width=15,
        values=["", "LOGIN", "LOGOUT", "PUT WAREHOUSE", "WAREHOUSE PULL", "UPDATE ITEM",
                "DELETE ITEM", "UNDO PULL", "UNSTAGE", "SHELF STATUS"]
    )
    filter_action_cb.pack(side="left", padx=(0, 10))

    count_label = tk.Label(filter_frame, text="", fg="blue")
    count_label.pack(side="right", padx=10)

    content_frame = tk.Frame(log_win)
    content_frame.pack(fill="both", expand=True, padx=10, pady=(0, 5))

    user_panel = tk.LabelFrame(content_frame, text="Users", padx=5, pady=5)
    user_panel.pack(side="left", fill="y", padx=(0, 8))

    user_scrollbar = ttk.Scrollbar(user_panel, orient="vertical")
    user_scrollbar.pack(side="right", fill="y")
    user_listbox = tk.Listbox(user_panel, width=18, yscrollcommand=user_scrollbar.set,
                               selectmode="single", exportselection=False, font=("Helvetica", 9))
    user_listbox.pack(side="left", fill="y")
    user_scrollbar.config(command=user_listbox.yview)

    btn_frame = tk.Frame(user_panel)
    btn_frame.pack(pady=(5, 0))
    tk.Button(btn_frame, text="↻", command=lambda: reset_filters(), width=3).pack(padx=2)

    table_frame_l = tk.Frame(content_frame)
    table_frame_l.pack(side="left", fill="both", expand=True)
    scrollbar_y = ttk.Scrollbar(table_frame_l, orient="vertical")
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x = ttk.Scrollbar(table_frame_l, orient="horizontal")
    scrollbar_x.pack(side="bottom", fill="x")
    tree_log = ttk.Treeview(table_frame_l,
                             columns=("Timestamp", "User", "Action", "Details"),
                             show="headings",
                             yscrollcommand=scrollbar_y.set,
                             xscrollcommand=scrollbar_x.set)
    for col, width in [("Timestamp", 140), ("User", 110), ("Action", 130), ("Details", 350)]:
        tree_log.heading(col, text=col); tree_log.column(col, width=width)
    tree_log.pack(fill="both", expand=True)
    scrollbar_y.config(command=tree_log.yview)
    scrollbar_x.config(command=tree_log.xview)

    def populate_user_listbox():
        try:
            df_log = load_logs()
            users = sorted(df_log["User"].dropna().unique().tolist())
        except Exception:
            users = []
        user_listbox.delete(0, tk.END)
        user_listbox.insert(tk.END, "(All Users)")
        for u in users:
            user_listbox.insert(tk.END, u)
        user_listbox.selection_set(0)

    def get_selected_user():
        sel = user_listbox.curselection()
        if not sel: return None
        val = user_listbox.get(sel[0])
        return None if val == "(All Users)" else val

    def load_log_data():
        tree_log.delete(*tree_log.get_children())
        df_log = load_logs()
        action_f = filter_action_var.get().strip()
        user_f = get_selected_user()
        if user_f:   df_log = df_log[df_log["User"] == user_f]
        if action_f: df_log = df_log[df_log["Action"] == action_f]
        df_log = df_log.iloc[::-1].reset_index(drop=True)
        for _, row in df_log.iterrows():
            tree_log.insert("", "end", values=tuple(
                row.get(c, "") for c in ["Timestamp", "User", "Action", "Details"]))
        count_label.config(text=f"{len(df_log)} record(s)")

    def reset_filters():
        filter_action_var.set("")
        user_listbox.selection_clear(0, tk.END)
        user_listbox.selection_set(0)
        load_log_data()

    filter_action_cb.bind("<<ComboboxSelected>>", lambda e: load_log_data())
    user_listbox.bind("<<ListboxSelect>>", lambda e: load_log_data())
    populate_user_listbox()
    load_log_data()



# ========== SWITCH USER ==========

def switch_user():
    global current_user, current_is_admin, session_start

    sw_win = tk.Toplevel(root)
    sw_win.title("Switch User — Login")
    sw_win.geometry("320x290")
    sw_win.resizable(False, False)
    sw_win.transient(root)
    sw_win.grab_set()

    hdr = tk.Frame(sw_win, bg="#2c3e50")
    hdr.pack(fill="x")
    tk.Label(hdr, text="Switch User", font=("Helvetica", 10, "bold"),
             bg="#2c3e50", fg="white", pady=8).pack()
    tk.Label(hdr, text="Enter credentials of the account to switch to",
             font=("Helvetica", 8), bg="#2c3e50", fg="#95a5a6", pady=0).pack(pady=(0, 8))

    form = tk.Frame(sw_win, padx=26, pady=14)
    form.pack(fill="x")

    tk.Label(form, text="Username", anchor="w", font=("Helvetica", 9)).grid(row=0, column=0, sticky="w")
    uv = tk.StringVar()
    u_entry = tk.Entry(form, textvariable=uv, width=26, font=("Helvetica", 10))
    u_entry.grid(row=1, column=0, pady=(2, 10), sticky="we")

    tk.Label(form, text="Password", anchor="w", font=("Helvetica", 9)).grid(row=2, column=0, sticky="w")
    pv = tk.StringVar()
    pw_e = tk.Entry(form, textvariable=pv, width=26, show="●", font=("Helvetica", 10))
    pw_e.grid(row=3, column=0, pady=(2, 4), sticky="we")

    show_var = tk.BooleanVar(value=False)
    tk.Checkbutton(form, text="Show password", variable=show_var,
                   command=lambda: pw_e.config(show="" if show_var.get() else "●"),
                   font=("Helvetica", 8)).grid(row=4, column=0, sticky="w")

    err = tk.Label(form, text="", fg="#e74c3c", font=("Helvetica", 8), wraplength=260)
    err.grid(row=5, column=0, sticky="w", pady=(6, 0))

    def do_switch(event=None):
        global current_user, current_is_admin, session_start
        uname = uv.get().strip()
        pw    = pv.get()
        if not uname or not pw:
            err.config(text="Please enter both username and password."); return
        if uname.lower() == current_user.lower():
            err.config(text="You are already logged in as this user."); return
        ok, role = authenticate_user(uname, pw)
        if not ok:
            err.config(text="Invalid username or password."); return
        save_log("LOGOUT", f"Session ended for '{current_user}'")
        current_user     = uname
        current_is_admin = (role == "admin")
        session_start    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        _refresh_user_bar()
        save_log("LOGIN", f"Session started by '{current_user}' (role: {role})")
        sw_win.destroy()
        messagebox.showinfo("User Switched", f"Switched to: {current_user}\nRole: {role.capitalize()}")

    pw_e.bind("<Return>", do_switch)
    u_entry.bind("<Return>", lambda e: pw_e.focus_set())
    tk.Button(sw_win, text="SWITCH USER", command=do_switch,
              bg="#2c3e50", fg="white", font=("Helvetica", 9, "bold"),
              pady=4, padx=20).pack(pady=(0, 10))
    u_entry.focus_set()
    sw_win.wait_window()


# ========== LOGIN ==========

def show_login():
    global current_user, current_is_admin, session_start
    initialize_users()

    login_win = tk.Tk()
    login_win.title("Warehouse System — Login")
    login_win.geometry("360x310")
    login_win.resizable(False, False)
    login_win.eval('tk::PlaceWindow . center')

    # ── Header ────────────────────────────────────────────────
    hdr = tk.Frame(login_win, bg="#2c3e50")
    hdr.pack(fill="x")
    tk.Label(hdr, text="Warehouse Management System",
             font=("Helvetica", 12, "bold"), bg="#2c3e50", fg="white",
             pady=12).pack()
    tk.Label(hdr, text="Please log in to continue",
             font=("Helvetica", 9), bg="#2c3e50", fg="#95a5a6",
             pady=0).pack(pady=(0, 10))

    form = tk.Frame(login_win, padx=30, pady=18)
    form.pack(fill="x")

    tk.Label(form, text="Username", anchor="w", font=("Helvetica", 9)).grid(row=0, column=0, sticky="w", pady=(0, 2))
    user_var = tk.StringVar()
    user_entry = tk.Entry(form, textvariable=user_var, width=26, font=("Helvetica", 10))
    user_entry.grid(row=1, column=0, pady=(0, 10), sticky="we")

    tk.Label(form, text="Password", anchor="w", font=("Helvetica", 9)).grid(row=2, column=0, sticky="w", pady=(0, 2))
    pw_var = tk.StringVar()
    pw_entry = tk.Entry(form, textvariable=pw_var, width=26, show="●", font=("Helvetica", 10))
    pw_entry.grid(row=3, column=0, pady=(0, 6), sticky="we")

    show_pw_var = tk.BooleanVar(value=False)
    def _toggle_pw():
        pw_entry.config(show="" if show_pw_var.get() else "●")
    tk.Checkbutton(form, text="Show password", variable=show_pw_var,
                   command=_toggle_pw, font=("Helvetica", 8)).grid(row=4, column=0, sticky="w")

    error_label = tk.Label(form, text="", fg="#e74c3c", font=("Helvetica", 8), wraplength=260)
    error_label.grid(row=5, column=0, pady=(6, 0), sticky="w")

    def attempt_login(event=None):
        global current_user, current_is_admin, session_start
        uname = user_var.get().strip()
        pw    = pw_var.get()
        if not uname or not pw:
            error_label.config(text="Please enter both username and password."); return
        ok, role = authenticate_user(uname, pw)
        if not ok:
            error_label.config(text="Invalid username or password."); return
        current_user     = uname
        current_is_admin = (role == "admin")
        session_start    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        login_win.quit()

    btn_row = tk.Frame(login_win, padx=30)
    btn_row.pack(fill="x")
    tk.Button(btn_row, text="LOGIN", command=attempt_login,
              bg="#2c3e50", fg="white", font=("Helvetica", 10, "bold"),
              width=14, pady=4).pack(side="left", padx=(0, 8))

    def open_register():
        _show_register_window(login_win)

    tk.Button(btn_row, text="CREATE ACCOUNT", command=open_register,
              font=("Helvetica", 10), width=17, pady=4).pack(side="left")

    pw_entry.bind("<Return>", attempt_login)
    user_entry.bind("<Return>", lambda e: pw_entry.focus_set())

    def on_close():
        if messagebox.askyesno("Exit", "Exit the Warehouse System?", parent=login_win):
            login_win.destroy()
            import sys; sys.exit(0)

    login_win.protocol("WM_DELETE_WINDOW", on_close)
    user_entry.focus_set()
    login_win.mainloop()
    login_win.destroy()
    initialize_log()
    save_log("LOGIN", f"Session started by '{current_user}' (role: {('admin' if current_is_admin else 'user')})")


def _show_register_window(parent):
    """Registration dialog callable from login screen or admin panel."""
    reg_win = tk.Toplevel(parent)
    reg_win.title("Create Account")
    reg_win.geometry("320x300")
    reg_win.resizable(False, False)
    reg_win.transient(parent)
    reg_win.grab_set()

    hdr = tk.Frame(reg_win, bg="#1a5276")
    hdr.pack(fill="x")
    tk.Label(hdr, text="Create New Account", font=("Helvetica", 10, "bold"),
             bg="#1a5276", fg="white", pady=8).pack()

    form = tk.Frame(reg_win, padx=24, pady=14)
    form.pack(fill="x")

    tk.Label(form, text="Username", anchor="w", font=("Helvetica", 9)).grid(row=0, column=0, sticky="w")
    uv = tk.StringVar()
    tk.Entry(form, textvariable=uv, width=26, font=("Helvetica", 10)).grid(row=1, column=0, pady=(2, 10), sticky="we")

    tk.Label(form, text="Password", anchor="w", font=("Helvetica", 9)).grid(row=2, column=0, sticky="w")
    pv = tk.StringVar()
    pw_e = tk.Entry(form, textvariable=pv, width=26, show="●", font=("Helvetica", 10))
    pw_e.grid(row=3, column=0, pady=(2, 4), sticky="we")

    tk.Label(form, text="Confirm Password", anchor="w", font=("Helvetica", 9)).grid(row=4, column=0, sticky="w")
    cpv = tk.StringVar()
    tk.Entry(form, textvariable=cpv, width=26, show="●", font=("Helvetica", 10)).grid(row=5, column=0, pady=(2, 4), sticky="we")

    err = tk.Label(form, text="", fg="#e74c3c", font=("Helvetica", 8), wraplength=260)
    err.grid(row=6, column=0, sticky="w", pady=(6, 0))

    def do_create():
        pw   = pv.get()
        cpw  = cpv.get()
        if pw != cpw:
            err.config(text="Passwords do not match."); return
        msg = create_account(uv.get(), pw)
        if msg:
            err.config(text=msg); return
        role = "Admin" if _is_admin_password(pw) else "User"
        messagebox.showinfo("Account Created",
            f"Account '{uv.get()}' created successfully!\nRole: {role}", parent=reg_win)
        reg_win.destroy()

    btn_row = tk.Frame(reg_win, pady=10)
    btn_row.pack()
    tk.Button(btn_row, text="CREATE ACCOUNT", command=do_create,
              font=("Helvetica", 10), bg="#1a5276", fg="white", padx=16, pady=4).pack(side="left")
    reg_win.focus_force()
    reg_win.wait_window()



# ========== ADMIN PANEL ==========

def open_admin_panel():
    """Account management window — admin only."""
    if not current_is_admin:
        messagebox.showerror("Access Denied", "Only admin accounts can access the Account Manager.")
        return

    panel = tk.Toplevel(root)
    panel.title("Account Manager")
    panel.geometry("700x480")
    panel.resizable(False, False)

    hdr = tk.Frame(panel, bg="#1a5276")
    hdr.pack(fill="x")
    tk.Label(hdr, text="Account Manager", font=("Helvetica", 11, "bold"),
             bg="#1a5276", fg="white", pady=8).pack(side="left", padx=12)
    tk.Label(hdr, text="(Admin only)", font=("Helvetica", 8, "italic"),
             bg="#1a5276", fg="#aed6f1").pack(side="left")

    # ── Table ─────────────────────────────────────────────────
    tbl_frame = tk.Frame(panel, bd=1, relief="sunken")
    tbl_frame.pack(fill="both", expand=True, padx=10, pady=(8, 4))

    sb_y = ttk.Scrollbar(tbl_frame, orient="vertical")
    sb_y.pack(side="right", fill="y")
    tree_acc = ttk.Treeview(tbl_frame,
        columns=("Username", "Role", "Created"),
        show="headings",
        yscrollcommand=sb_y.set)
    for col, w in [("Username", 200), ("Role", 120), ("Created", 180)]:
        tree_acc.heading(col, text=col)
        tree_acc.column(col, width=w, anchor="w")
    tree_acc.pack(fill="both", expand=True)
    sb_y.config(command=tree_acc.yview)

    def _load_table():
        tree_acc.delete(*tree_acc.get_children())
        df = load_users()
        for _, row in df.iterrows():
            role = str(row.get("Role", "user")).capitalize()
            tree_acc.insert("", "end", values=(
                row.get("Username", ""),
                role,
                row.get("Created", ""),
            ))

    # ── Bottom action area ────────────────────────────────────
    act_frame = tk.Frame(panel, padx=10, pady=8)
    act_frame.pack(fill="x")

    # Create account sub-frame
    create_lf = tk.LabelFrame(act_frame, text="Create Account", padx=8, pady=6)
    create_lf.pack(side="left", fill="y", padx=(0, 10))

    tk.Label(create_lf, text="Username", font=("Helvetica", 8)).grid(row=0, column=0, sticky="w")
    new_user_var = tk.StringVar()
    tk.Entry(create_lf, textvariable=new_user_var, width=16, font=("Helvetica", 9)).grid(row=1, column=0, pady=(2, 6))

    tk.Label(create_lf, text="Password", font=("Helvetica", 8)).grid(row=2, column=0, sticky="w")
    new_pw_var = tk.StringVar()
    tk.Entry(create_lf, textvariable=new_pw_var, width=16, show="●", font=("Helvetica", 9)).grid(row=3, column=0, pady=(2, 4))

    tk.Label(create_lf,
        text="Include ! @ # → Admin role",
        font=("Helvetica", 7, "italic"), fg="#7f8c8d").grid(row=4, column=0, sticky="w")

    create_err = tk.Label(create_lf, text="", fg="#e74c3c", font=("Helvetica", 7), wraplength=140)
    create_err.grid(row=5, column=0, sticky="w")

    def do_create():
        err = create_account(new_user_var.get(), new_pw_var.get())
        if err:
            create_err.config(text=err); return
        role = "Admin" if _is_admin_password(new_pw_var.get()) else "User"
        create_err.config(text="")
        new_user_var.set("")
        new_pw_var.set("")
        _load_table()
        messagebox.showinfo("Created",
            f"Account '{new_user_var.get() or '(created)'}' added as {role}.", parent=panel)

    tk.Button(create_lf, text="CREATE", command=do_create,
              bg="#1a5276", fg="white", font=("Helvetica", 8, "bold"), pady=3).grid(row=6, column=0, pady=(6, 0))

    # Delete / change-password sub-frame
    del_lf = tk.LabelFrame(act_frame, text="Manage Selected Account", padx=8, pady=6)
    del_lf.pack(side="left", fill="y", padx=(0, 10))

    del_err = tk.Label(del_lf, text="", fg="#e74c3c", font=("Helvetica", 7), wraplength=160)
    del_err.grid(row=0, column=0, columnspan=2, sticky="w")

    def do_delete():
        sel = tree_acc.selection()
        if not sel:
            del_err.config(text="Select an account first."); return
        uname = tree_acc.item(sel[0], "values")[0]
        if uname.lower() == current_user.lower():
            del_err.config(text="Cannot delete your own account."); return
        if not messagebox.askyesno("Confirm Delete",
                f"Delete account '{uname}'?\nThis cannot be undone.", parent=panel):
            return
        err = delete_account(uname)
        if err:
            del_err.config(text=err); return
        del_err.config(text="")
        _load_table()
        messagebox.showinfo("Deleted", f"Account '{uname}' deleted.", parent=panel)

    tk.Button(del_lf, text="DELETE ACCOUNT", command=do_delete,
              bg="#922b21", fg="white", font=("Helvetica", 8, "bold"), pady=3,
              width=16).grid(row=1, column=0, pady=(6, 10), sticky="w")

    tk.Label(del_lf, text="New Password", font=("Helvetica", 8)).grid(row=2, column=0, sticky="w")
    chpw_var = tk.StringVar()
    tk.Entry(del_lf, textvariable=chpw_var, width=18, show="●",
             font=("Helvetica", 9)).grid(row=3, column=0, pady=(2, 4))

    def do_change_pw():
        sel = tree_acc.selection()
        if not sel:
            del_err.config(text="Select an account first."); return
        uname = tree_acc.item(sel[0], "values")[0]
        err = change_password(uname, chpw_var.get())
        if err:
            del_err.config(text=err); return
        del_err.config(text="")
        chpw_var.set("")
        _load_table()
        messagebox.showinfo("Updated", f"Password for '{uname}' updated.", parent=panel)

    tk.Button(del_lf, text="CHANGE PASSWORD", command=do_change_pw,
              bg="#117a65", fg="white", font=("Helvetica", 8, "bold"), pady=3,
              width=16).grid(row=4, column=0, pady=(4, 0), sticky="w")

    _load_table()


def _refresh_user_bar():
    """Update the user bar label and show/hide admin-only buttons after a user switch."""
    role_badge = " 🔑" if current_is_admin else ""
    user_label.config(text=f"👤  {current_user}{role_badge}")
    session_label.config(text=f"Session started: {session_start}")
    # Show/hide admin buttons
    if current_is_admin:
        activity_log_btn.pack(side="right", padx=4, pady=2)
        admin_panel_btn.pack(side="right", padx=4, pady=2)
    else:
        activity_log_btn.pack_forget()
        admin_panel_btn.pack_forget()


def _guarded_activity_log():
    if not current_is_admin:
        messagebox.showerror("Access Denied", "Only admin accounts can view the Activity Log.")
        return
    open_activity_log()


# ========== UI SETUP ==========

show_login()

root = tk.Tk()
root.title("Warehouse System — Developed by Mark Benjamin (IT Intern 2026)")
root.geometry("1280x780")
root.eval('tk::PlaceWindow . center')

# ── User bar ──────────────────────────────────────────────
user_bar = tk.Frame(root, bg="#2c3e50", height=28)
user_bar.pack(fill="x")

clock_label = tk.Label(user_bar, text="", bg="#2c3e50", fg="#95a5a6", font=("Helvetica", 8))
clock_label.pack(side="right", padx=10, pady=4)

tip(tk.Button(user_bar, text="Change User", command=switch_user,
          bg="#34495e", fg="white", bd=0, padx=10),
    "Switch to a different user session. Credentials required.").pack(side="right", padx=10, pady=2)

admin_panel_btn = tip(tk.Button(user_bar, text="👥 Accounts", command=open_admin_panel,
          bg="#34495e", fg="white", bd=0, padx=10),
    "Manage user accounts (Admin only).")
activity_log_btn = tip(tk.Button(user_bar, text="📋 Activity Log", command=_guarded_activity_log,
          bg="#34495e", fg="white", bd=0, padx=10),
    "View the full history of all actions performed in this system (Admin only).")

# Pack admin buttons only if current user is admin
if current_is_admin:
    activity_log_btn.pack(side="right", padx=4, pady=2)
    admin_panel_btn.pack(side="right", padx=4, pady=2)

role_badge = " 🔑" if current_is_admin else ""
user_label = tk.Label(user_bar, text=f"👤  {current_user}{role_badge}", bg="#2c3e50", fg="white", font=("Helvetica", 9, "bold"))
user_label.pack(side="left", padx=10, pady=4)

session_label = tk.Label(user_bar, text=f"Session started: {session_start}", bg="#2c3e50", fg="#95a5a6", font=("Helvetica", 8))
session_label.pack(side="left", padx=5, pady=4)

def update_clock():
    clock_label.config(text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    root.after(1000, update_clock)
update_clock()

# ── Notebook (tabs) ───────────────────────────────────────
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=6, pady=6)

tab1 = tk.Frame(notebook)
tab2 = tk.Frame(notebook)
notebook.add(tab1, text="  Warehouse 1 — Laptops")
notebook.add(tab2, text="  Warehouse 2 — Computer Peripherals / Equipment")  
# ══════════════════════════════════════════════════════════
#  WAREHOUSE 1 TAB
# ══════════════════════════════════════════════════════════

w1_main = tk.Frame(tab1)
w1_main.pack(fill="both", expand=True, padx=8, pady=8)

w1_row1 = tk.Frame(w1_main)
w1_row1.pack(fill="both", expand=True)

# Item Management
input_frame = tk.LabelFrame(w1_row1, text="Item Management", padx=10, pady=5)
input_frame.pack(side="left", fill="both", padx=5)
tip(input_frame, "Add new items to the staging queue before committing to the warehouse.")

tk.Label(input_frame, text="Hostname").grid(row=0, column=0, sticky="w")
hostname_entry = tk.Entry(input_frame, width=22); hostname_entry.grid(row=0, column=1, pady=3)
tip(hostname_entry, "Enter the device hostname (e.g. PC-001). Must be unique.")

tk.Label(input_frame, text="Serial Number").grid(row=1, column=0, sticky="w")
serial_entry = tk.Entry(input_frame, width=22); serial_entry.grid(row=1, column=1, pady=3)
tip(serial_entry, "Enter the manufacturer serial number. Must be unique.")

tk.Label(input_frame, text="Checked By").grid(row=2, column=0, sticky="w")
checked_by_entry = tk.Entry(input_frame, width=22); checked_by_entry.grid(row=2, column=1, pady=3)
tip(checked_by_entry, "Name of the person who physically checked this item.")

tk.Label(input_frame, text="Shelf").grid(row=3, column=0, sticky="w")
shelf_var = tk.StringVar()
shelf_dropdown = ttk.Combobox(input_frame, textvariable=shelf_var, width=19)
shelf_dropdown.grid(row=3, column=1, pady=3)
tip(shelf_dropdown, "Select the shelf/area where this item will be stored.")

tk.Label(input_frame, text="Status").grid(row=4, column=0, sticky="w")
remarks_var = tk.StringVar()
ttk.Combobox(input_frame, textvariable=remarks_var, values=STATUS_CHOICES, width=19, state="readonly").grid(row=4, column=1, pady=3)

tk.Label(input_frame, text="Remarks").grid(row=5, column=0, sticky="w")
remarks_text_var = tk.StringVar()
remarks_text_entry = tk.Entry(input_frame, textvariable=remarks_text_var, width=22)
remarks_text_entry.grid(row=5, column=1, pady=3)
tip(remarks_text_entry, "Optional notes about the item's condition.")

crud_frame = tk.Frame(input_frame)
crud_frame.grid(row=6, column=0, columnspan=2, pady=5)
tip(tk.Button(crud_frame, text="PUT",    command=put_item,    width=8), "Add this item to the staging queue.").grid(row=0, column=0, padx=3)
tip(tk.Button(crud_frame, text="UPDATE", command=update_item, width=8), "Update the selected staged item with new values.").grid(row=0, column=1, padx=3)
tip(tk.Button(crud_frame, text="↻",      command=reset_ui,    width=3), "Clear all input fields and reset the view.").grid(row=0, column=2, padx=3)

tk.Label(input_frame, text="Staged Items (Click to Edit)", fg="green", font=("Helvetica", 9, "bold")).grid(row=7, column=0, columnspan=2, sticky="w")
staged_listbox = tk.Listbox(input_frame, width=32)
staged_listbox.grid(row=8, column=0, columnspan=2, sticky="nswe", pady=3)
staged_listbox.bind("<<ListboxSelect>>", select_staged_item)
tip(staged_listbox, "Items waiting to be committed. Click a row to load it back into the fields for editing.")
input_frame.rowconfigure(8, weight=1)
input_frame.columnconfigure(0, weight=1)
input_frame.columnconfigure(1, weight=1)

staging_btn_frame = tk.Frame(input_frame)
staging_btn_frame.grid(row=9, column=0, columnspan=2, pady=3)
tip(tk.Button(staging_btn_frame, text="CLEAR ITEMS",   command=remove_from_staging, width=13),
    "Remove the selected staged item (or all if none selected).").pack(side="left", padx=2)
tip(tk.Button(staging_btn_frame, text="PUT WAREHOUSE", command=put_warehouse,       width=13),
    "Commit all staged items to Warehouse 1 and generate QR codes. Use GENERATE FILES to create PDF labels.").pack(side="left", padx=2)

# Shelf Controls W1
shelf_mid_frame = tk.Frame(w1_row1)
shelf_mid_frame.pack(side="left", fill="both", expand=True, padx=5)

# Shelf + View sub-row (side by side inside shelf_mid_frame)
w1_shelf_view_row = tk.Frame(shelf_mid_frame)
w1_shelf_view_row.pack(fill="x")

shelf_control_frame = tk.LabelFrame(w1_shelf_view_row, text="Shelf Control & Management", padx=10, pady=5)
shelf_control_frame.pack(side="left", fill="both", expand=True)
tip(shelf_control_frame, "Manage shelf availability and add/remove shelves for Warehouse 1.")

status_control_frame = tk.LabelFrame(shelf_control_frame, text="Status Control", padx=8, pady=5)
status_control_frame.pack(fill="x", pady=(0, 5))
tip(status_control_frame, "Mark a shelf as FULL (no more items) or AVAILABLE.")
shelf_control_var = tk.StringVar()
shelf_control_dropdown = ttk.Combobox(status_control_frame, textvariable=shelf_control_var, width=22, state="readonly")
shelf_control_dropdown.pack(side="left", padx=5)
tip(shelf_control_dropdown, "Select the shelf whose status you want to change.")
tip(tk.Button(status_control_frame, text="SET FULL",      command=lambda: set_shelf_status("FULL"),      width=10),
    "Mark selected shelf as FULL — prevents new items from being placed there.").pack(side="left", padx=3)
tip(tk.Button(status_control_frame, text="SET AVAILABLE", command=lambda: set_shelf_status("AVAILABLE"), width=12),
    "Mark selected shelf as AVAILABLE — allows new items to be placed.").pack(side="left", padx=3)
tip(tk.Button(status_control_frame, text="↻",             command=reset_shelf_control,                   width=3),
    "Clear the shelf selection.").pack(side="left", padx=3)

add_remove_frame = tk.LabelFrame(shelf_control_frame, text="Add / Remove", padx=8, pady=5)
add_remove_frame.pack(fill="x")
tip(add_remove_frame, "Add a new shelf by typing its name, or remove an existing empty shelf.")
remove_shelf_var = tk.StringVar()
remove_shelf_dropdown = ttk.Combobox(add_remove_frame, textvariable=remove_shelf_var, width=22)
remove_shelf_dropdown.pack(side="left", padx=5)
tip(remove_shelf_dropdown, "Type a new shelf name to add, or select an existing one to remove.")
tip(tk.Button(add_remove_frame, text="ADD",    command=add_shelf   ), "Add the typed shelf name as a new shelf.").pack(side="left", padx=3)
tip(tk.Button(add_remove_frame, text="REMOVE", command=remove_shelf ), "Remove the selected shelf (only if it has no items).").pack(side="left", padx=3)
tip(tk.Button(add_remove_frame, text="↻",      command=reset_shelf_addition, width=3), "Clear the shelf name field.").pack(side="left", padx=3)

# View W1 — sits to the right of Shelf Control inside shelf_mid_frame
view_frame = tk.LabelFrame(w1_shelf_view_row, text="View", padx=10, pady=5)
view_frame.pack(side="left", fill="y", padx=(5, 0))
tip(view_frame, "Switch between different table views for Warehouse 1.")
for text, cmd, tooltip in [
    ("SHOW WAREHOUSE", show_warehouse,  "View all items currently stored in Warehouse 1."),
    ("SHELF STATUS",   show_available,  "View each shelf's status and how many items it holds."),
    ("PULL HISTORY",   show_pullouts,   "View items that have been pulled out of Warehouse 1."),
    ("QR LABELS",      lambda: open_label_manager(warehouse=1), "Open and manage printable QR label PDF files."),
    ("VIEW EXCEL",     w1_view_excel,   "Open the last Excel file generated by GENERATE FILES for Warehouse 1."),
]:
    tip(tk.Button(view_frame, text=text, command=cmd, width=15), tooltip).pack(anchor="w", pady=3)

# Search & Filter W1
w1_pullout_frame = tk.LabelFrame(shelf_mid_frame, text="Warehouse 1", padx=10, pady=8)
w1_pullout_frame.pack(fill="x", pady=5)

w1_search_filter = tk.LabelFrame(w1_pullout_frame, text="Search & Filter", padx=8, pady=5)
w1_search_filter.pack(fill="x", pady=(0, 5))

# Row 1: search + shelf + remarks + single 🔍 button + clear
w1_sf_row1 = tk.Frame(w1_search_filter)
w1_sf_row1.pack(fill="x", pady=(0, 3))
tk.Label(w1_sf_row1, text="Search:").pack(side="left", padx=(5, 2))
search_entry = tk.Entry(w1_sf_row1, width=20); search_entry.pack(side="left", padx=(0, 2))
search_entry.bind("<Return>", lambda e: search_item())
search_entry.bind("<KeyRelease>", pull_search_live)

tk.Label(w1_sf_row1, text="Shelf:").pack(side="left", padx=(5, 2))
pull_shelf_var = tk.StringVar()
pull_shelf_dropdown = ttk.Combobox(w1_sf_row1, textvariable=pull_shelf_var, width=16, state="readonly")
pull_shelf_dropdown.pack(side="left", padx=(0, 8))

tk.Label(w1_sf_row1, text="Remarks:").pack(side="left", padx=(5, 2))
pull_remarks_var = tk.StringVar()
ttk.Combobox(w1_sf_row1, textvariable=pull_remarks_var, values=[""] + STATUS_CHOICES, width=14, state="readonly").pack(side="left", padx=(0, 8))

tip(tk.Button(w1_sf_row1, text="🔍", command=search_item,       width=3), "Search and filter the warehouse table.").pack(side="left", padx=3)
tip(tk.Button(w1_sf_row1, text="↻",  command=clear_pull_filters, width=3), "Clear all search and filter fields.").pack(side="left", padx=3)

# Row 2: date range (calendar pickers)
w1_sf_row2 = tk.Frame(w1_search_filter)
w1_sf_row2.pack(fill="x", pady=(0, 2))
w1_date_from_var = tk.StringVar()
w1_date_to_var   = tk.StringVar()
_date_picker_widget(w1_sf_row2, w1_date_from_var, "From:").pack(side="left", padx=(5, 8))
_date_picker_widget(w1_sf_row2, w1_date_to_var,   "To:").pack(side="left", padx=(0, 5))

pull_item_entry = search_entry

# Pull Out W1
w1_pull_action = tk.LabelFrame(w1_pullout_frame, text="Pull Out", padx=8, pady=5)
w1_pull_action.pack(fill="x", pady=(5, 0))

# Pull Out Row 1: pull reason dropdown + redo + warehouse pull buttons only
w1_po_row1 = tk.Frame(w1_pull_action)
w1_po_row1.pack(fill="x", pady=(0, 3))
tk.Label(w1_po_row1, text="Pull Reason:").pack(side="left", padx=(5, 2))
pull_reason_filter_var = tk.StringVar()
pull_reason_filter_entry = ttk.Combobox(w1_po_row1, textvariable=pull_reason_filter_var, width=20)
pull_reason_filter_entry.pack(side="left", padx=(0, 8))
w1_pull_date_from_var = tk.StringVar()
w1_pull_date_to_var   = tk.StringVar()
tip(tk.Button(w1_po_row1, text="↻",  command=reset_pull_out,      width=3),  "Clear pull-out fields and reset the view.").pack(side="left", padx=2)
tip(tk.Button(w1_po_row1, text="WAREHOUSE PULL", command=pull_item, width=16), "Pull the selected item out of Warehouse 1. A pull reason is required.").pack(side="left", padx=(10, 3))

# Status bar W1
w1_status_bar = tk.Frame(shelf_mid_frame)
w1_status_bar.pack(fill="x")
w1_full_label   = tk.Label(w1_status_bar, text="FULL Shelves: None", fg="red");  w1_full_label.pack(side="left", padx=10)
w1_search_label = tk.Label(w1_status_bar, text="", fg="blue");                   w1_search_label.pack(side="left", padx=10)
w1_status_label = tk.Label(w1_status_bar, text="", fg="green");                  w1_status_label.pack(side="left", padx=10)

# ── W1 Table toolbar (Select All + Stored QR) ──────────────
w1_table_toolbar = tk.Frame(shelf_mid_frame)
w1_table_toolbar.pack(fill="x", padx=5, pady=(2, 0))

def w1_toggle_select_all():
    all_iids = list(tree_warehouse.get_children())
    checked  = [iid for iid in all_iids if w1_row_checks.get(iid)]
    # If pull history is visible, toggle that instead
    if tree_pullouts.winfo_ismapped():
        all_iids = list(tree_pullouts.get_children())
        checked  = [iid for iid in all_iids if w1_pull_row_checks.get(iid)]
        new_state = len(checked) < len(all_iids)
        for iid in all_iids:
            w1_pull_row_checks[iid] = new_state
            tree_pullouts.set(iid, "CP0", "☑" if new_state else "☐")
        return
    new_state = len(checked) < len(all_iids)
    for iid in all_iids:
        w1_row_checks[iid] = new_state
        tree_warehouse.set(iid, "C0", "☑" if new_state else "☐")
    _w1_refresh_select_all_label()

w1_select_all_btn = tk.Button(w1_table_toolbar, text="SELECT ALL",
    command=w1_toggle_select_all, width=12)
w1_select_all_btn.pack(side="left", padx=(0, 6))
tip(w1_select_all_btn, "Select or deselect all visible rows in the active table.")
tip(tk.Button(w1_table_toolbar, text="GENERATE FILES", command=w1_generate_stored_qr,
          bg="#6c3483", fg="white", width=15),
    "Generate QR PNGs, PDF labels and Excel export for selected items.").pack(side="left", padx=(0, 6))
tip(tk.Button(w1_table_toolbar, text="VIEW QR", command=w1_view_stored_qr,
          bg="#1a5276", fg="white", width=10),
    "View all existing QR codes without regenerating.").pack(side="left", padx=(0, 6))
w1_back_to_stage_btn = tip(tk.Button(w1_table_toolbar, text="BACK TO STAGE", command=unstage_from_warehouse,
          bg="#117a65", fg="white", width=15),
    "Move checked warehouse items back to the Item Management staging list.")
w1_back_to_stage_btn.pack(side="left", padx=(0, 6))
w1_back_to_wh_btn = tip(tk.Button(w1_table_toolbar, text="BACK TO WAREHOUSE", command=undo_pull,
          bg="#922b21", fg="white", width=18),
    "Restore checked pull history items back to Warehouse 1.")
# Hidden by default; shown only when Pull History view is active

# Tables W1
w1_table_frame = tk.Frame(shelf_mid_frame)
w1_table_frame.pack(fill="both", expand=True, pady=5)

tree_warehouse = ttk.Treeview(w1_table_frame, columns=("C0","C1","C2","C3","C4","C5","C6","C7","C8"), show='headings')
for col, text, width in zip(("C0","C1","C2","C3","C4","C5","C6","C7","C8"),
    ("✔","QR","Hostname","Serial Number","Checked By","Shelf","Status","Remarks","Date"),
    (30,180,150,120,115,130,95,145,145)):
    tree_warehouse.heading(col, text=text); tree_warehouse.column(col, width=width)
tree_warehouse.column("C0", anchor="center", stretch=False)
tree_warehouse.bind("<<TreeviewSelect>>", select_item)

tree_available = ttk.Treeview(w1_table_frame, columns=("C1","C2","C3","C4"), show='headings')
for col, text, width in zip(("C1","C2","C3","C4"), ("Shelf","Status","Items Stored","Date Set Full"), (250,140,110,200)):
    tree_available.heading(col, text=text); tree_available.column(col, width=width)

tree_pullouts = ttk.Treeview(w1_table_frame, columns=("CP0","C1","C2","C3","C4","C5","C6","C7"), show='headings')
for col, text, width in zip(("CP0","C1","C2","C3","C4","C5","C6","C7"),
    ("✔","Hostname","Serial Number","Shelf","Status","Remarks","Pull Reason","Date"), (30,145,125,125,90,155,205,150)):
    tree_pullouts.heading(col, text=text); tree_pullouts.column(col, width=width)
tree_pullouts.column("CP0", anchor="center", stretch=False)
tree_pullouts.bind("<<TreeviewSelect>>", select_pull_item)

tree_qr = ttk.Treeview(w1_table_frame, columns=("C1","C2","C3"), show='headings')
for col, text, width in zip(("C1","C2","C3"),
    ("Hostname","QR UUID String","File Status (PNG)"), (200,400,150)):
    tree_qr.heading(col, text=text); tree_qr.column(col, width=width)

# ══════════════════════════════════════════════════════════
#  WAREHOUSE 2 TAB
# ══════════════════════════════════════════════════════════

w2_main = tk.Frame(tab2)
w2_main.pack(fill="both", expand=True, padx=8, pady=8)

w2_row1 = tk.Frame(w2_main)
w2_row1.pack(fill="both", expand=True)

# Equipment selection + staging panel
w2_input_frame = tk.LabelFrame(w2_row1, text="Set Staging", padx=10, pady=5)
w2_input_frame.pack(side="left", fill="both", padx=5)
tip(w2_input_frame, "Build equipment sets (Monitor, Keyboard, Mouse, Headset) and stage them before committing to Warehouse 2.")

tip(tk.Label(w2_input_frame, text="Select Equipment:", font=("Helvetica", 9, "bold")),
    "Check the equipment types to include in this set.").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

w2_equip_vars = {}
for i, eq in enumerate(EQUIPMENT_TYPES):
    var = tk.BooleanVar()
    w2_equip_vars[eq] = var
    tk.Checkbutton(w2_input_frame, text=eq, variable=var, width=10, anchor="w").grid(
        row=1 + i // 2, column=i % 2, sticky="w", padx=4)

tip(tk.Button(w2_input_frame, text="BUILD SET", command=w2_build_set,
          bg="#2980b9", fg="white", width=28),
    "Open a form to enter details for each selected equipment type, then stage the set.").grid(row=3, column=0, columnspan=2, pady=(8, 4))

tip(tk.Label(w2_input_frame, text="Staged Sets", fg="green", font=("Helvetica", 9, "bold")),
    "Sets waiting to be committed to Warehouse 2.").grid(row=4, column=0, columnspan=2, sticky="w")
w2_staged_listbox = tk.Listbox(w2_input_frame, width=34)
w2_staged_listbox.grid(row=5, column=0, columnspan=2, sticky="nswe", pady=3)
tip(w2_staged_listbox, "Click a set to select it for editing. ⚠ means a required equipment type is missing.")
w2_input_frame.rowconfigure(5, weight=1)
w2_input_frame.columnconfigure(0, weight=1)
w2_input_frame.columnconfigure(1, weight=1)

w2_stage_btns = tk.Frame(w2_input_frame)
w2_stage_btns.grid(row=6, column=0, columnspan=2, pady=3)
tip(tk.Button(w2_stage_btns, text="CLEAR SETS",    command=w2_remove_staged_set, width=13),
    "Remove the selected staged set (or all sets if none selected).").pack(side="left", padx=2)
tip(tk.Button(w2_stage_btns, text="PUT WAREHOUSE", command=w2_put_warehouse,     width=13),
    "Commit all staged sets to Warehouse 2 and generate QR codes. Use GENERATE FILES to create PDF labels.").pack(side="left", padx=2)

# Item Management (UPDATE / DELETE) for W2
w2_item_mgmt_frame = tk.LabelFrame(w2_input_frame, text="Item Management", padx=6, pady=4)
w2_item_mgmt_frame.grid(row=7, column=0, columnspan=2, pady=(6, 0), sticky="we")
tip(w2_item_mgmt_frame, "Edit a staged set's details before committing it to the warehouse.")
tk.Label(w2_item_mgmt_frame,
         text="Select a row in the staged sets table,\nthen use buttons below:",
         font=("Helvetica", 8), fg="gray", justify="left").pack(anchor="w", pady=(0, 4))
w2_item_action_btns = tk.Frame(w2_item_mgmt_frame)
w2_item_action_btns.pack()
tip(tk.Button(w2_item_action_btns, text="UPDATE ITEM", command=w2_update_item,
          bg="#1a5276", fg="white", width=13),
    "Edit the items in the selected staged set.").pack(side="left", padx=2)

# Shelf Controls W2
w2_shelf_mid = tk.Frame(w2_row1)
w2_shelf_mid.pack(side="left", fill="both", expand=True, padx=5)

# Shelf + View sub-row (side by side inside w2_shelf_mid)
w2_shelf_view_row = tk.Frame(w2_shelf_mid)
w2_shelf_view_row.pack(fill="x")

w2_shelf_ctrl_frame = tk.LabelFrame(w2_shelf_view_row, text="Shelf Control & Management", padx=10, pady=5)
w2_shelf_ctrl_frame.pack(side="left", fill="both", expand=True)
tip(w2_shelf_ctrl_frame, "Manage shelf availability and add/remove shelves for Warehouse 2.")

w2_status_ctrl = tk.LabelFrame(w2_shelf_ctrl_frame, text="Status Control", padx=8, pady=5)
w2_status_ctrl.pack(fill="x", pady=(0, 5))
tip(w2_status_ctrl, "Mark a shelf as FULL (no more items) or AVAILABLE.")
w2_shelf_control_var = tk.StringVar()
w2_shelf_control_dropdown = ttk.Combobox(w2_status_ctrl, textvariable=w2_shelf_control_var, width=22, state="readonly")
w2_shelf_control_dropdown.pack(side="left", padx=5)
tip(w2_shelf_control_dropdown, "Select the shelf whose status you want to change.")
tip(tk.Button(w2_status_ctrl, text="SET FULL",      command=lambda: w2_set_shelf_status("FULL"),      width=10),
    "Mark selected shelf as FULL — prevents new items from being placed there.").pack(side="left", padx=3)
tip(tk.Button(w2_status_ctrl, text="SET AVAILABLE", command=lambda: w2_set_shelf_status("AVAILABLE"), width=12),
    "Mark selected shelf as AVAILABLE — allows new items to be placed.").pack(side="left", padx=3)
tip(tk.Button(w2_status_ctrl, text="↻",             command=w2_reset_shelf_control,                   width=3),
    "Clear the shelf selection.").pack(side="left", padx=3)

w2_add_remove = tk.LabelFrame(w2_shelf_ctrl_frame, text="Add / Remove", padx=8, pady=5)
w2_add_remove.pack(fill="x")
tip(w2_add_remove, "Add a new shelf by typing its name, or remove an existing empty shelf.")
w2_remove_shelf_var = tk.StringVar()
w2_remove_shelf_dropdown = ttk.Combobox(w2_add_remove, textvariable=w2_remove_shelf_var, width=22)
w2_remove_shelf_dropdown.pack(side="left", padx=5)
tip(w2_remove_shelf_dropdown, "Type a new shelf name to add, or select an existing one to remove.")
tip(tk.Button(w2_add_remove, text="ADD",    command=w2_add_shelf   ), "Add the typed shelf name as a new shelf.").pack(side="left", padx=3)
tip(tk.Button(w2_add_remove, text="REMOVE", command=w2_remove_shelf ), "Remove the selected shelf (only if it has no items).").pack(side="left", padx=3)
tip(tk.Button(w2_add_remove, text="↻",      command=w2_reset_shelf_addition, width=3), "Clear the shelf name field.").pack(side="left", padx=3)

# View W2 — sits to the right of Shelf Control inside w2_shelf_mid
w2_view_frame = tk.LabelFrame(w2_shelf_view_row, text="View", padx=10, pady=5)
w2_view_frame.pack(side="left", fill="y", padx=(5, 0))
tip(w2_view_frame, "Switch between different table views for Warehouse 2.")
for text, cmd, tooltip in [
    ("SHOW WAREHOUSE", w2_show_warehouse, "View all items currently stored in Warehouse 2."),
    ("SHELF STATUS",   w2_show_available, "View each shelf's status and how many items it holds."),
    ("PULL HISTORY",   w2_show_pullouts,  "View items that have been pulled out of Warehouse 2."),
    ("QR LABELS",      lambda: open_label_manager(warehouse=2), "Open and manage printable QR label PDF files."),
    ("VIEW EXCEL",     w2_view_excel,     "Open the last Excel file generated by GENERATE FILES for Warehouse 2."),
]:
    tip(tk.Button(w2_view_frame, text=text, command=cmd, width=15), tooltip).pack(anchor="w", pady=3)

# Search & Filter W2
w2_pullout_frame = tk.LabelFrame(w2_shelf_mid, text="Warehouse 2", padx=10, pady=8)
w2_pullout_frame.pack(fill="x", pady=5)

w2_search_filter = tk.LabelFrame(w2_pullout_frame, text="Search & Filter", padx=8, pady=5)
w2_search_filter.pack(fill="x", pady=(0, 5))

# Row 1: search + shelf + type + single 🔍 button + clear
w2_sf_row1 = tk.Frame(w2_search_filter)
w2_sf_row1.pack(fill="x", pady=(0, 3))
tk.Label(w2_sf_row1, text="Search:").pack(side="left", padx=(5, 2))
w2_search_entry = tk.Entry(w2_sf_row1, width=18); w2_search_entry.pack(side="left", padx=(0, 2))
w2_search_entry.bind("<Return>", lambda e: w2_search_item())
w2_search_entry.bind("<KeyRelease>", w2_pull_search_live)

tk.Label(w2_sf_row1, text="Shelf:").pack(side="left", padx=(5, 2))
w2_pull_shelf_var = tk.StringVar()
w2_pull_shelf_dropdown = ttk.Combobox(w2_sf_row1, textvariable=w2_pull_shelf_var, width=16, state="readonly")
w2_pull_shelf_dropdown.pack(side="left", padx=(0, 10))

tk.Label(w2_sf_row1, text="Type:").pack(side="left", padx=(5, 2))
w2_type_filter_var = tk.StringVar()
ttk.Combobox(w2_sf_row1, textvariable=w2_type_filter_var,
             values=[""] + EQUIPMENT_TYPES, width=12, state="readonly").pack(side="left", padx=(0, 8))

tip(tk.Button(w2_sf_row1, text="🔍", command=w2_search_item,  width=3), "Search and filter the Warehouse 2 table.").pack(side="left", padx=3)
tip(tk.Button(w2_sf_row1, text="↻",  command=w2_clear_filters, width=3), "Clear all search and filter fields.").pack(side="left", padx=3)

# Row 2: date range (calendar pickers)
w2_sf_row2 = tk.Frame(w2_search_filter)
w2_sf_row2.pack(fill="x", pady=(0, 2))
w2_date_from_var = tk.StringVar()
w2_date_to_var   = tk.StringVar()
_date_picker_widget(w2_sf_row2, w2_date_from_var, "From:").pack(side="left", padx=(5, 8))
_date_picker_widget(w2_sf_row2, w2_date_to_var,   "To:").pack(side="left", padx=(0, 5))

w2_pull_item_entry = w2_search_entry

# Pull Out W2
w2_pull_action = tk.LabelFrame(w2_pullout_frame, text="Pull Out", padx=8, pady=5)
w2_pull_action.pack(fill="x", pady=(5, 0))

# Pull Out Row 1: pull reason dropdown + redo + warehouse pull buttons only
w2_po_row1 = tk.Frame(w2_pull_action)
w2_po_row1.pack(fill="x", pady=(0, 3))
tk.Label(w2_po_row1, text="Pull Reason:").pack(side="left", padx=(5, 2))
w2_pull_reason_filter_var = tk.StringVar()
w2_pull_reason_filter_entry = ttk.Combobox(w2_po_row1, textvariable=w2_pull_reason_filter_var, width=20)
w2_pull_reason_filter_entry.pack(side="left", padx=(0, 8))
w2_pull_date_from_var = tk.StringVar()
w2_pull_date_to_var   = tk.StringVar()
tip(tk.Button(w2_po_row1, text="↻",  command=w2_reset_pull_out,      width=3),  "Clear pull-out fields and reset the view.").pack(side="left", padx=2)
tip(tk.Button(w2_po_row1, text="WAREHOUSE PULL", command=w2_pull_item, width=16), "Pull the selected item out of Warehouse 2. A pull reason is required.").pack(side="left", padx=(10, 3))

# Status bar W2
w2_status_bar = tk.Frame(w2_shelf_mid)
w2_status_bar.pack(fill="x")
w2_full_label   = tk.Label(w2_status_bar, text="FULL Shelves: None", fg="red");  w2_full_label.pack(side="left", padx=10)
w2_search_label = tk.Label(w2_status_bar, text="", fg="blue");                   w2_search_label.pack(side="left", padx=10)
w2_status_label = tk.Label(w2_status_bar, text="", fg="green");                  w2_status_label.pack(side="left", padx=10)

# ── W2 Table toolbar (Select All + Stored QR) ──────────────
w2_table_toolbar = tk.Frame(w2_shelf_mid)
w2_table_toolbar.pack(fill="x", padx=5, pady=(2, 0))

def w2_toggle_select_all():
    # If pull history is visible, toggle that instead
    if tree_w2_pullouts.winfo_ismapped():
        all_iids = list(tree_w2_pullouts.get_children())
        checked  = [iid for iid in all_iids if w2_pull_row_checks.get(iid)]
        new_state = len(checked) < len(all_iids)
        for iid in all_iids:
            w2_pull_row_checks[iid] = new_state
            tree_w2_pullouts.set(iid, "CP0", "☑" if new_state else "☐")
        return
    all_iids = list(tree_w2_warehouse.get_children())
    checked  = [iid for iid in all_iids if w2_row_checks.get(iid)]
    new_state = len(checked) < len(all_iids)
    for iid in all_iids:
        w2_row_checks[iid] = new_state
        tree_w2_warehouse.set(iid, "C0", "☑" if new_state else "☐")
    _w2_refresh_select_all_label()

w2_select_all_btn = tk.Button(w2_table_toolbar, text="SELECT ALL",
    command=w2_toggle_select_all, width=12)
w2_select_all_btn.pack(side="left", padx=(0, 6))
tip(w2_select_all_btn, "Select or deselect all visible rows in the active table.")

tip(tk.Button(w2_table_toolbar, text="GENERATE FILES", command=w2_generate_stored_qr,
          bg="#6c3483", fg="white", width=15),
    "Generate QR PNGs, PDF labels and Excel export for selected items.").pack(side="left", padx=(0, 6))
tip(tk.Button(w2_table_toolbar, text="VIEW QR", command=w2_view_stored_qr,
          bg="#1a5276", fg="white", width=10),
    "View all existing QR codes without regenerating.").pack(side="left", padx=(0, 6))
w2_back_to_stage_btn = tip(tk.Button(w2_table_toolbar, text="BACK TO STAGE", command=w2_unstage_from_warehouse,
          bg="#117a65", fg="white", width=15),
    "Move checked warehouse items back to the Set Staging list.")
w2_back_to_stage_btn.pack(side="left", padx=(0, 6))
w2_back_to_wh_btn = tip(tk.Button(w2_table_toolbar, text="BACK TO WAREHOUSE", command=w2_undo_pull,
          bg="#922b21", fg="white", width=18),
    "Restore checked pull history items back to Warehouse 2.")
# Hidden by default; shown only when Pull History view is active

# Tables W2
w2_table_frame = tk.Frame(w2_shelf_mid)
w2_table_frame.pack(fill="both", expand=True, pady=5)

tree_w2_warehouse = ttk.Treeview(w2_table_frame,
    columns=("C0","C1","C2","C3","C4","C5","C6","C7","C8","C9","C10"), show='headings')
for col, text, width in zip(
    ("C0","C1","C2","C3","C4","C5","C6","C7","C8","C9","C10"),
    ("✔","QR","Set ID","Hostname","Equipment Type","Serial Number","Checked By","Shelf","Status","Remarks","Date"),
    (30,170,90,120,115,115,115,125,90,150,135)):
    tree_w2_warehouse.heading(col, text=text); tree_w2_warehouse.column(col, width=width)
tree_w2_warehouse.column("C0", anchor="center", stretch=False)
tree_w2_warehouse.bind("<<TreeviewSelect>>", w2_select_item)

tree_w2_available = ttk.Treeview(w2_table_frame, columns=("C1","C2","C3","C4"), show='headings')
for col, text, width in zip(("C1","C2","C3","C4"), ("Shelf","Status","Items Stored","Date Set Full"), (250,140,110,200)):
    tree_w2_available.heading(col, text=text); tree_w2_available.column(col, width=width)

tree_w2_pullouts = ttk.Treeview(w2_table_frame,
    columns=("CP0","C1","C2","C3","C4","C5","C6","C7","C8","C9"), show='headings')
for col, text, width in zip(
    ("CP0","C1","C2","C3","C4","C5","C6","C7","C8","C9"),
    ("✔","Set ID","Hostname","Equipment Type","Serial Number","Shelf","Status","Remarks","Pull Reason","Date"),
    (30,85,115,115,115,110,80,150,180,135)):
    tree_w2_pullouts.heading(col, text=text); tree_w2_pullouts.column(col, width=width)
tree_w2_pullouts.column("CP0", anchor="center", stretch=False)
tree_w2_pullouts.bind("<<TreeviewSelect>>", w2_select_pull_item)

tree_w2_qr = ttk.Treeview(w2_table_frame, columns=("C1","C2","C3","C4"), show='headings')
for col, text, width in zip(("C1","C2","C3","C4"),
    ("Set ID","Equipment Type","QR UUID String","File Status (PNG)"), (100,120,380,130)):
    tree_w2_qr.heading(col, text=text); tree_w2_qr.column(col, width=width)
# ── Init ──────────────────────────────────────────────────
initialize_file()
update_all_shelf_dropdowns()
update_staged_display()
update_w2_staged_display()
show_warehouse()
w2_show_warehouse()

# Attach click-to-sort on all main treeviews
for _t in (tree_warehouse, tree_available, tree_pullouts, tree_qr,
           tree_w2_warehouse, tree_w2_available, tree_w2_pullouts, tree_w2_qr):
    attach_sort_headers(_t)

def on_main_close():
    if messagebox.askyesno("Exit", f"Log out '{current_user}' and exit the system?"):
        save_log("LOGOUT", f"Session ended for '{current_user}'")
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_main_close)
root.mainloop()