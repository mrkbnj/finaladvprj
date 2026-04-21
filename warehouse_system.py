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
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(BASE_DIR, "warehouse.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "activity_log.xlsx")
QR_FOLDER = os.path.join(BASE_DIR, "qr_codes")
QR_FOLDER_W1 = os.path.join(QR_FOLDER, "warehouse_1")
QR_FOLDER_W2 = os.path.join(QR_FOLDER, "warehouse_2")
QR_LABELS_FOLDER = os.path.join(BASE_DIR, "qr_labels")
QR_LABELS_FOLDER_W1 = os.path.join(QR_LABELS_FOLDER, "warehouse_1")
QR_LABELS_FOLDER_W2 = os.path.join(QR_LABELS_FOLDER, "warehouse_2")

SHELVES = [
    "Area A", "Area B", "Area C",
    "Rack 1 - Bay 1", "Rack 1 - Bay 2", "Rack 1 - Bay 3",
    "Rack 2 - Bay 1", "Rack 2 - Bay 2", "Rack 2 - Bay 3",
]
SHELVES_W1 = SHELVES_W2 = SHELVES

EQUIPMENT_TYPES = ["Monitor", "Keyboard", "Mouse", "Headset"]

staged_items = []
selected_staged_index = None
staged_sets = []
selected_set_index = None
current_user = ""
session_start = ""

# ========== INITIALIZATION ==========

def initialize_file():
    # Always ensure all folders exist
    for folder in (QR_FOLDER_W1, QR_FOLDER_W2, QR_LABELS_FOLDER_W1, QR_LABELS_FOLDER_W2):
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
        "items": pd.DataFrame(columns=["QR", "Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]),
        "shelves": pd.DataFrame({"Shelf": SHELVES_W1, "Status": ["AVAILABLE"] * len(SHELVES_W1), "Date_Full": [None] * len(SHELVES_W1)}),
        "pullouts": pd.DataFrame(columns=["Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]),

        # W2 Sheets
        "items_w2": pd.DataFrame(columns=["QR", "Set ID", "Hostname", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]),
        "shelves_w2": pd.DataFrame({"Shelf": SHELVES_W2, "Status": ["AVAILABLE"] * len(SHELVES_W2), "Date_Full": [None] * len(SHELVES_W2)}),
        "pullouts_w2": pd.DataFrame(columns=["Set ID", "Hostname", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]),
    }
    try:
        with pd.ExcelWriter(FILE, engine='openpyxl', mode=mode) as writer:
            for sheet in sheets_to_create:
                default_dfs[sheet].to_excel(writer, sheet_name=sheet, index=False)
    except Exception as e:
        messagebox.showerror("File Error", f"Could not create '{FILE}':\n{e}\n\nMake sure the file is not open in Excel.")

def initialize_log():
    if not os.path.exists(LOG_FILE):
        with pd.ExcelWriter(LOG_FILE, engine='openpyxl') as writer:
            pd.DataFrame(columns=["Timestamp", "User", "Action", "Details"]).to_excel(writer, sheet_name="logs", index=False)

# ========== LOAD / SAVE ==========

def _load_sheet(file, sheet, init_fn):
    try:
        return pd.read_excel(file, sheet_name=sheet)
    except Exception:
        try:
            init_fn()
            return pd.read_excel(file, sheet_name=sheet)
        except Exception as e:
            messagebox.showerror("File Error",
                f"Could not load sheet '{sheet}' from '{file}':\n{e}\n\n"
                "Make sure the file is not open in Excel and the folder is writable.")
            return pd.DataFrame()

def load_items():
    df = _load_sheet(FILE, "items", initialize_file)
    expected = ["QR", "Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
    for col in expected:
        if col not in df.columns:
            df[col] = pd.Series(dtype=str)
    return df
def load_shelves():     return _load_sheet(FILE, "shelves", initialize_file) # W1 Only
def load_shelves_w2():  return _load_sheet(FILE, "shelves_w2", initialize_file) # W2 Only
def load_pullouts():    return _load_sheet(FILE, "pullouts", initialize_file)
def load_items_w2():    return _load_sheet(FILE, "items_w2", initialize_file)
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
    prev_btn = tk.Button(header, text="◀", bg="#2c3e50", fg="white", bd=0, font=("Arial", 10),
                         command=lambda: _change_month(-1))
    prev_btn.pack(side="left", padx=8, pady=4)
    month_lbl = tk.Label(header, text="", bg="#2c3e50", fg="white", font=("Arial", 10, "bold"), width=16)
    month_lbl.pack(side="left", expand=True)
    next_btn = tk.Button(header, text="▶", bg="#2c3e50", fg="white", bd=0, font=("Arial", 10),
                         command=lambda: _change_month(1))
    next_btn.pack(side="right", padx=8, pady=4)

    day_names = tk.Frame(cal_win, bg="#dce3f0")
    day_names.pack(fill="x")
    for i, d in enumerate(["Su","Mo","Tu","We","Th","Fr","Sa"]):
        tk.Label(day_names, text=d, width=4, font=("Arial", 8, "bold"),
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
                        font=("Arial", 9, "bold" if is_today else "normal"),
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
    tk.Button(clear_frame, text="Clear date", fg="gray", bd=0, font=("Arial", 8),
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
                        bg="white", font=("Arial", 9), anchor="w", padx=4)
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
    _fill_input_fields(item["Hostname"], item.get("Brand/Model", ""), item.get("Serial Number", ""), item.get("Checked By", ""), item["Shelf"], item.get("Status", ""), item.get("Remarks", ""))

# ========== W1 INPUT HELPERS ==========

def _fill_input_fields(hostname="", brand="", serial="", checked_by="", shelf="", status="", remarks=""):
    hostname_entry.delete(0, tk.END);   hostname_entry.insert(0, hostname)
    brand_entry.delete(0, tk.END);      brand_entry.insert(0, brand)
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
    brand      = brand_entry.get().strip()
    shelf      = shelf_var.get()
    serial     = serial_entry.get().strip()
    checked_by = checked_by_entry.get().strip()
    status     = remarks_var.get()
    remarks    = remarks_text_var.get().strip()

    if not hostname:
        messagebox.showerror("Error", "Please enter a Hostname"); return
    if not brand:
        messagebox.showerror("Error", "Please enter a Brand/Model"); return
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
        "Brand/Model":   brand,
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
        for col in ["Brand/Model", "Serial Number", "Checked By"]:
            if col not in df_items.columns:
                df_items[col] = ""

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for item in staged_items:
            qr_code = str(uuid.uuid4())
            generate_qr(item['Hostname'], (
                f"Hostname: {item['Hostname']}\n"
                f"Brand/Model: {item.get('Brand/Model', '')}\n"
                f"Serial Number: {item.get('Serial Number', '')}\n"
                f"Checked By: {item.get('Checked By', '')}\n"
                f"Shelf: {item['Shelf']}\n"
                f"Status: {item.get('Status', '')}\n"
                f"Remarks: {item.get('Remarks', '')}\n"
                f"Date: {now_str}"
            ), warehouse=1)
            df_items = pd.concat([df_items, pd.DataFrame([{
                "QR": qr_code,
                "Hostname": item['Hostname'],
                "Brand/Model": item.get('Brand/Model', ''),
                "Serial Number": item.get('Serial Number', ''),
                "Checked By": item.get('Checked By', ''),
                "Shelf": item['Shelf'],
                "Status": item.get('Status', ''),
                "Remarks": item.get('Remarks', ''),
                "Date": now_str
            }])], ignore_index=True)

        save_warehouse_1(df_items, df_shelves)

        try:
            pdf_path = generate_qr_pdf([{**item, '_warehouse': 1} for item in staged_items])
            pdf_msg = f"\nQR labels saved to:\n{pdf_path}"
        except Exception as pdf_err:
            pdf_msg = f"\nPDF generation failed: {pdf_err}"

        count = len(staged_items)
        for item in staged_items:
            save_log("PUT WAREHOUSE", f"[W1] Hostname: {item['Hostname']} | Shelf: {item['Shelf']}")

        staged_items.clear()
        messagebox.showinfo("Success", f"{count} item(s) added to Warehouse 1{pdf_msg}")
        update_staged_display()
        w1_refresh_all()

    except Exception as e:
        messagebox.showerror("Save Error",
            f"Failed to save to Excel:\n{str(e)}\n\n"
            "Common causes:\n• Excel file is open → close it\n• Wrong folder")
        


def update_item():
    global selected_staged_index
    new_hostname   = hostname_entry.get().strip()
    new_brand      = brand_entry.get().strip()
    new_serial     = serial_entry.get().strip()
    new_checked_by = checked_by_entry.get().strip()
    new_shelf      = shelf_var.get()
    new_status     = remarks_var.get()
    new_remarks    = remarks_text_var.get().strip()

    if not new_hostname:
        messagebox.showerror("Error", "Hostname cannot be empty"); return
    if not new_brand:
        messagebox.showerror("Error", "Brand/Model cannot be empty"); return
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
            "Brand/Model":   new_brand,
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
    index = tree_warehouse.index(selected[0])
    hostname = df_items.at[index, "Hostname"]

    if not messagebox.askyesno("Confirm Delete", f"Delete '{hostname}'?\nThis cannot be undone."):
        return

    delete_qr(hostname, warehouse=1)
    df_items = df_items.drop(index).reset_index(drop=True)
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
        if keyword:
            mask = False
            for col in ["Shelf", "Status", "Date_Full"]:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        tree_available.delete(*tree_available.get_children())
        for _, row in df.iterrows():
            date_full = row.get("Date_Full", "")
            tree_available.insert("", "end", values=(row["Shelf"], row["Status"], date_full if pd.notna(date_full) else ""))
        w1_search_label.config(text=f"{len(df)} match(es)" if keyword else "")

    elif tree_pullouts.winfo_ismapped():
        # Pull history view is active
        df = load_pullouts()
        if keyword:
            search_cols = ["Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
            mask = False
            for col in search_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        tree_pullouts.delete(*tree_pullouts.get_children())
        for _, row in df.iterrows():
            tree_pullouts.insert("", "end", values=tuple(row.get(c, "") for c in ["Hostname", "Brand/Model", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]))
        w1_search_label.config(text=f"{len(df)} match(es)" if keyword else "")

    else:
        # Warehouse view (default)
        show_warehouse()
        if keyword:
            df = load_items()
            search_cols = ["QR", "Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
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
        "Brand/Model": str(item_row.get("Brand/Model", "")),
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

def undo_pull(event):
    item_id = tree_pullouts.identify_row(event.y)
    if not item_id:
        return
    values = tree_pullouts.item(item_id, "values")
    if not values:
        return

    hostname, brand, shelf, status, remarks = values[0], values[1], values[2], values[3], values[4]
    if not messagebox.askyesno("Undo Pull", f"Restore '{hostname}' back to the warehouse?\n\nShelf: {shelf}\nStatus: {status}"):
        return

    df_items = load_items()
    df_shelves = load_shelves()
    df_pullouts = load_pullouts()

    if hostname in df_items["Hostname"].values:
        messagebox.showerror("Error", f"'{hostname}' already exists in warehouse"); return

    match = df_pullouts[df_pullouts["Hostname"] == hostname]
    if match.empty:
        messagebox.showerror("Error", f"'{hostname}' not found in pull history"); return

    pull_row = match.iloc[0]
    qr_code = ""
    try:
        qr_code = str(uuid.uuid4())
        generate_qr(hostname, qr_code, warehouse=1)
    except Exception as e:
        messagebox.showwarning("Warning", f"QR code not regenerated: {e}")

    for col in ["Brand/Model", "Serial Number", "Checked By"]:
        if col not in df_items.columns:
            df_items[col] = ""

    df_items = pd.concat([df_items, pd.DataFrame([{
        "QR": qr_code,
        "Hostname": hostname,
        "Brand/Model": str(pull_row.get("Brand/Model", "")),
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
    messagebox.showinfo("Restored", f"'{hostname}' has been restored to the warehouse")
    show_pullouts()

def unstage_from_warehouse(event):
    item_id = tree_warehouse.identify_row(event.y)
    if not item_id:
        return
    values = tree_warehouse.item(item_id, "values")
    if not values:
        return

    hostname, brand, serial, checked_by, shelf, status, remarks = values[1], values[2], values[3], values[4], values[5], values[6], values[7]
    if not messagebox.askyesno("Move to Staging", f"Move '{hostname}' back to staging?\n\nShelf: {shelf}\nStatus: {status}"):
        return
    if any(item['Hostname'] == hostname for item in staged_items):
        messagebox.showerror("Error", f"'{hostname}' is already in staging"); return

    df_items = load_items()
    df_shelves = load_shelves()
    delete_qr(hostname, warehouse=1)
    df_items = df_items[df_items["Hostname"] != hostname].reset_index(drop=True)
    save_warehouse_1(df_items, df_shelves)
    staged_items.append({"Hostname": hostname, "Brand/Model": brand, "Serial Number": serial, "Checked By": checked_by, "Shelf": shelf, "Status": status, "Remarks": remarks})
    save_log("UNSTAGE", f"[W1] Hostname: {hostname} | Shelf: {shelf}")
    messagebox.showinfo("Moved", f"'{hostname}' moved back to staging")
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

def _open_qr_gallery(warehouse):
    """Shared QR gallery window for both warehouses."""
    from PIL import Image, ImageTk
    wh_label = f"Warehouse {warehouse}"
    bg_color = "#2c3e50" if warehouse == 1 else "#1a5276"
    btn_color = "#1a252f" if warehouse == 1 else "#154360"

    qr_win = tk.Toplevel(root)
    qr_win.title(f"Stored QR Codes — {wh_label}")
    qr_win.geometry("860x560")

    toolbar = tk.Frame(qr_win, bg=bg_color)
    toolbar.pack(fill="x")
    tk.Label(toolbar, text=f"Stored QR Codes — {wh_label}",
             bg=bg_color, fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=10, pady=6)
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
                cell_labels = [(key, ("Arial", 8, "bold"), "#2c3e50", 0),
                               (f"S/N: {sub}", ("Arial", 7), "#555", 0),
                               (f"Shelf: {shelf}", ("Arial", 7), "gray", 0)]
            else:
                set_id  = str(row.get("Set ID", ""))
                eq_type = str(row.get("Equipment Type", ""))
                shelf   = str(row.get("Shelf", ""))
                key     = f"{set_id}-{eq_type}"
                sub     = str(row.get("Serial Number", ""))
                host    = str(row.get("Hostname", ""))
                path    = qr_path_for(key, warehouse=2)
                kw_fields = [set_id, eq_type, shelf]
                cell_labels = [(key, ("Arial", 8, "bold"), "#2c3e50", 0),
                               (host, ("Arial", 7, "italic"), "#2c3e50", 0),
                               (f"S/N: {sub}", ("Arial", 7), "#555", 0),
                               (f"Shelf: {shelf}", ("Arial", 7), "gray", 0)]

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
                         width=14, height=6, font=("Arial", 8)).pack()
            for text, font, fg, _ in cell_labels:
                tk.Label(cell, text=text, bg="white", font=font, fg=fg, wraplength=130).pack(pady=(4,0) if _ == 0 else 0)

            col_f += 1; shown += 1
            if col_f >= COLS:
                col_f = 0; row_f += 1

        if shown == 0:
            tk.Label(inner, text="No QR codes found.", bg="#f4f6f7",
                     font=("Arial", 10), fg="gray").grid(row=0, column=0, padx=20, pady=40)
        count_lbl.config(text=f"{shown} QR code(s)")
        inner.update_idletasks()
        canvas_qr.configure(scrollregion=canvas_qr.bbox("all"))
        canvas_qr.itemconfigure(canvas_window_id, width=canvas_qr.winfo_width())

    canvas_qr.bind("<Configure>", lambda e: canvas_qr.itemconfigure(canvas_window_id, width=e.width))
    search_var.trace_add("write", lambda *_: _load_gallery(search_var.get()))
    _load_gallery()

def show_qr_codes():    _open_qr_gallery(warehouse=1)
def w2_show_qr_codes(): _open_qr_gallery(warehouse=2)


def show_warehouse():
    w1_update_full_shelves_display()
    _show_tree(tree_warehouse)
    tree_warehouse.delete(*tree_warehouse.get_children())
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
    for _, row in df_items.iterrows():
        tree_warehouse.insert("", "end", values=tuple(row.get(c, "") for c in ["QR", "Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]))

def show_available():
    _show_tree(tree_available)
    tree_available.delete(*tree_available.get_children())
    try:
        keyword = search_entry.get().strip().lower()
    except Exception:
        keyword = ""
    df = load_shelves().sort_values("Shelf")
    if keyword:
        mask = False
        for col in ["Shelf", "Status", "Date_Full"]:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]
    for _, row in df.iterrows():
        date_full = row.get("Date_Full", "")
        tree_available.insert("", "end", values=(row["Shelf"], row["Status"], date_full if pd.notna(date_full) else ""))

def show_pullouts():
    _show_tree(tree_pullouts)
    tree_pullouts.delete(*tree_pullouts.get_children())
    df_po = load_pullouts()
    for _, row in df_po.iterrows():
        tree_pullouts.insert("", "end", values=tuple(row.get(c, "") for c in ["Hostname", "Brand/Model", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]))
    try:
        all_reasons = sorted(df_po["Pull Reason"].dropna().unique().tolist())
        pull_reason_filter_entry["values"] = [""] + all_reasons
    except Exception:
        pass

def _populate_warehouse_tree(df):
    _show_tree(tree_warehouse)
    tree_warehouse.delete(*tree_warehouse.get_children())
    for _, row in df.iterrows():
        tree_warehouse.insert("", "end", values=tuple(row.get(c, "") for c in ["QR", "Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]))

def search_item():
    keyword      = search_entry.get().strip().lower()
    shelf_filter = pull_shelf_var.get()
    remarks_filter = pull_remarks_var.get()
    date_from    = w1_date_from_var.get().strip()
    date_to      = w1_date_to_var.get().strip()

    df = load_items()

    if keyword:
        search_cols = ["QR", "Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]

    if shelf_filter:   df = df[df["Shelf"] == shelf_filter]
    if remarks_filter: df = df[df["Status"] == remarks_filter]

    df = _filter_by_date(df, date_from, date_to)
    _populate_warehouse_tree(df)
    if shelf_filter:   parts.append(f"Shelf: {shelf_filter}")
    if remarks_filter: parts.append(f"Status: {remarks_filter}")
    if date_from:      parts.append(f"From: {date_from}")
    if date_to:        parts.append(f"To: {date_to}")
    label = (f"{len(df)} result(s)" + (" — " + " | ".join(parts) if parts else "")) if parts else ""
    w1_search_label.config(text=label)

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
        search_cols = ["Set ID", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(reason, na=False)
        df = df[mask]
    df = _filter_by_date(df, date_from, date_to)
    # Refresh pull reason dropdown
    all_reasons = sorted(load_pullouts()["Pull Reason"].dropna().unique().tolist())
    pull_reason_filter_entry["values"] = [""] + all_reasons
    for _, row in df.iterrows():
        tree_pullouts.insert("", "end", values=tuple(row.get(c, "") for c in ["Hostname", "Brand/Model", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]))
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

def select_item(event):
    selected = tree_warehouse.selection()
    if selected:
        values = tree_warehouse.item(selected[0], "values")
        # values: QR, Hostname, Brand/Model, Serial Number, Checked By, Shelf, Status, Remarks, Date
        pull_item_entry.delete(0, tk.END)
        pull_item_entry.insert(0, values[1])
        w1_status_label.config(
            text=f"Selected → Hostname: {values[1]}  |  Shelf: {values[5]}  |  Serial: {values[3]}",
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

    tk.Label(build_win, text=f"Fill in details for {set_id}", font=("Arial", 10, "bold")).pack(pady=(10, 5))

    shelf_list = sorted(load_shelves_w2()["Shelf"].tolist())

    COL_WIDTHS = [12, 18, 20, 18, 16, 16, 13, 20]
    HEADERS    = ["Equipment", "Hostname", "Brand / Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks"]

    outer = tk.Frame(build_win, padx=10, pady=5)
    outer.pack(fill="both", expand=True)

    # Fixed header
    hdr_frame = tk.Frame(outer, bg="#dce3f0")
    hdr_frame.pack(fill="x")
    for col, (h, cw) in enumerate(zip(HEADERS, COL_WIDTHS)):
        tk.Label(hdr_frame, text=h, font=("Arial", 9, "bold"), width=cw, anchor="w",
                 bg="#dce3f0", padx=5).grid(row=0, column=col, padx=5, pady=4, sticky="w")

    ROW_COLORS = ("#ffffff", "#f0f4ff")

    rows = {}
    for r, eq_type in enumerate(selected_types):
        bg = ROW_COLORS[r % 2]
        row_bg = tk.Frame(outer, bg=bg, bd=1, relief="flat")
        row_bg.pack(fill="x", pady=1)
        tk.Label(row_bg, text=eq_type, width=COL_WIDTHS[0], anchor="w", bg=bg,
                 font=("Arial", 9, "bold")).grid(row=0, column=0, padx=5, pady=8, sticky="w")
        hostname_e = tk.Entry(row_bg, width=COL_WIDTHS[1], font=("Arial", 9)); hostname_e.grid(row=0, column=1, padx=5, pady=8)
        brand_e    = tk.Entry(row_bg, width=COL_WIDTHS[2], font=("Arial", 9)); brand_e.grid(row=0, column=2, padx=5, pady=8)
        serial_e   = tk.Entry(row_bg, width=COL_WIDTHS[3], font=("Arial", 9)); serial_e.grid(row=0, column=3, padx=5, pady=8)
        checked_e  = tk.Entry(row_bg, width=COL_WIDTHS[4], font=("Arial", 9)); checked_e.grid(row=0, column=4, padx=5, pady=8)
        shelf_v = tk.StringVar()
        ttk.Combobox(row_bg, textvariable=shelf_v, values=shelf_list, width=COL_WIDTHS[5], state="readonly",
                     font=("Arial", 9)).grid(row=0, column=5, padx=5, pady=8)
        status_v = tk.StringVar()
        ttk.Combobox(row_bg, textvariable=status_v, values=["No Issue", "Minimal", "Defective"],
                     width=COL_WIDTHS[6], state="readonly", font=("Arial", 9)).grid(row=0, column=6, padx=5, pady=8)
        remarks_e = tk.Entry(row_bg, width=COL_WIDTHS[7], font=("Arial", 9)); remarks_e.grid(row=0, column=7, padx=5, pady=8)
        rows[eq_type] = (hostname_e, brand_e, serial_e, checked_e, shelf_v, status_v, remarks_e)

    error_lbl = tk.Label(outer, text="", fg="red", font=("Arial", 8))
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
        for eq_type, (hostname_e, brand_e, serial_e, checked_e, shelf_v, status_v, remarks_e) in rows.items():
            hostname   = hostname_e.get().strip()
            brand      = brand_e.get().strip()
            serial     = serial_e.get().strip()
            checked_by = checked_e.get().strip()
            shelf      = shelf_v.get().strip()
            status     = status_v.get().strip()
            remarks    = remarks_e.get().strip()

            if not hostname:
                error_lbl.config(text=f"Please enter a Hostname for {eq_type}"); return
            if not brand:
                error_lbl.config(text=f"Please enter a Brand/Model for {eq_type}"); return
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
                "Brand/Model":    brand,
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
        pdf_items = []

        for s in staged_sets:
            set_id = s["set_id"]
            for item in s["items"]:
                eq_type = item["Equipment Type"]
                qr_label = f"{set_id}-{eq_type}"
                qr_code = str(uuid.uuid4())
                generate_qr(qr_label, (
                    f"Set ID: {set_id}\n"
                    f"Hostname: {item.get('Hostname', '')}\n"
                    f"Equipment: {eq_type}\n"
                    f"Brand/Model: {item.get('Brand/Model', '')}\n"
                    f"Serial Number: {item.get('Serial Number', '')}\n"
                    f"Checked By: {item.get('Checked By', '')}\n"
                    f"Shelf: {item['Shelf']}\n"
                    f"Remarks: {item['Remarks']}\n"
                    f"Date: {now_str}"
                ), warehouse=2)
                df_w2 = pd.concat([df_w2, pd.DataFrame([{
                    "QR":             qr_code,
                    "Set ID":         set_id,
                    "Hostname":       item.get("Hostname", ""),
                    "Equipment Type": eq_type,
                    "Brand/Model":    item.get("Brand/Model", ""),
                    "Serial Number":  item.get("Serial Number", ""),
                    "Checked By":     item.get("Checked By", ""),
                    "Shelf":          item["Shelf"],
                    "Status":         item.get("Status", ""),
                    "Remarks":        item.get("Remarks", ""),
                    "Date":           now_str
                }])], ignore_index=True)
                pdf_items.append({
                    "Hostname":       item.get("Hostname", qr_label),
                    "Set ID":         set_id,
                    "Equipment Type": eq_type,
                    "Brand/Model":    item.get("Brand/Model", ""),
                    "Serial Number":  item.get("Serial Number", ""),
                    "Checked By":     item.get("Checked By", ""),
                    "Shelf":          item["Shelf"],
                    "Status":         item.get("Status", ""),
                    "Remarks":        item.get("Remarks", ""),
                    "_warehouse":     2,
                })
            save_log("PUT WAREHOUSE", f"[W2] Set: {set_id} | Items: {len(s['items'])}")

        save_warehouse_2(df_w2, load_shelves_w2())

        try:
            pdf_path = generate_qr_pdf(pdf_items)
            pdf_msg = f"\nQR labels saved to:\n{pdf_path}"
        except Exception as pdf_err:
            pdf_msg = f"\nPDF generation failed: {pdf_err}"

        count = len(staged_sets)
        staged_sets.clear()
        update_w2_staged_display()
        w2_pull_item_entry.delete(0, tk.END)
        w2_search_label.config(text="", fg="blue")
        w2_status_label.config(text="")
        w2_refresh_all()
        messagebox.showinfo("Success", f"{count} set(s) added to Warehouse 2{pdf_msg}")

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
    _show_w2_tree(tree_w2_warehouse)
    tree_w2_warehouse.delete(*tree_w2_warehouse.get_children())
    df = load_items_w2()
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
        search_cols = ["QR", "Set ID", "Hostname", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]
    if shelf_f: df = df[df["Shelf"] == shelf_f]
    if type_f:  df = df[df["Equipment Type"] == type_f]
    df = _filter_by_date(df, date_from, date_to)
    for _, row in df.iterrows():
        tree_w2_warehouse.insert("", "end", values=tuple(
            row.get(c, "") for c in ["QR", "Set ID", "Hostname", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]))
    if active:
        w2_search_label.config(text=f"🔍 Active filter — {len(df)} result(s) shown", fg="darkorange")

def w2_show_available():
    _show_w2_tree(tree_w2_available)
    tree_w2_available.delete(*tree_w2_available.get_children())
    try:
        keyword = w2_search_entry.get().strip().lower()
    except Exception:
        keyword = ""
    df = load_shelves_w2().sort_values("Shelf")
    if keyword:
        mask = False
        for col in ["Shelf", "Status", "Date_Full"]:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
        df = df[mask]
    for _, row in df.iterrows():
        date_full = row.get("Date_Full", "")
        tree_w2_available.insert("", "end", values=(row["Shelf"], row["Status"], date_full if pd.notna(date_full) else ""))

def w2_show_pullouts():
    _show_w2_tree(tree_w2_pullouts)
    tree_w2_pullouts.delete(*tree_w2_pullouts.get_children())
    df_po2 = load_pullouts_w2()
    for _, row in df_po2.iterrows():
        tree_w2_pullouts.insert("", "end", values=tuple(
            row.get(c, "") for c in ["Set ID", "Hostname", "Equipment Type", "Brand/Model", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]))
    try:
        all_reasons = sorted(df_po2["Pull Reason"].dropna().unique().tolist())
        w2_pull_reason_filter_entry["values"] = [""] + all_reasons
    except Exception:
        pass

def _populate_w2_warehouse_tree(df):
    _show_w2_tree(tree_w2_warehouse)
    tree_w2_warehouse.delete(*tree_w2_warehouse.get_children())
    for _, row in df.iterrows():
        tree_w2_warehouse.insert("", "end", values=tuple(
            row.get(c, "") for c in ["QR", "Set ID", "Hostname", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]))

def w2_search_item():
    keyword   = w2_search_entry.get().strip().lower()
    shelf_f   = w2_pull_shelf_var.get()
    type_f    = w2_type_filter_var.get()
    date_from = w2_date_from_var.get().strip()
    date_to   = w2_date_to_var.get().strip()

    df = load_items_w2()

    if keyword:
        search_cols = ["QR", "Set ID", "Hostname", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
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
        if keyword:
            mask = False
            for col in ["Shelf", "Status", "Date_Full"]:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        tree_w2_available.delete(*tree_w2_available.get_children())
        for _, row in df.iterrows():
            date_full = row.get("Date_Full", "")
            tree_w2_available.insert("", "end", values=(row["Shelf"], row["Status"], date_full if pd.notna(date_full) else ""))
        w2_search_label.config(text=f"{len(df)} match(es)" if keyword else "", fg="blue")

    elif tree_w2_pullouts.winfo_ismapped():
        # Pull history view is active
        df = load_pullouts_w2()
        if keyword:
            search_cols = ["Set ID", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
            mask = False
            for col in search_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            df = df[mask]
        tree_w2_pullouts.delete(*tree_w2_pullouts.get_children())
        for _, row in df.iterrows():
            tree_w2_pullouts.insert("", "end", values=tuple(row.get(c, "") for c in ["Set ID", "Hostname", "Equipment Type", "Brand/Model", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]))
        w2_search_label.config(text=f"{len(df)} match(es)" if keyword else "", fg="blue")

    else:
        # Warehouse view (default)
        w2_show_warehouse()
        if keyword:
            df = load_items_w2()
            search_cols = ["QR", "Set ID", "Hostname", "Equipment Type", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Date"]
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
        values = tree_w2_warehouse.item(selected[0], "values")
        # values: QR, Set ID, Hostname, Equipment Type, Brand/Model, Serial, Checked By, Shelf, Remarks, Date
        w2_pull_item_entry.delete(0, tk.END)
        w2_pull_item_entry.insert(0, f"{values[1]} - {values[3]}")  # SET-001 - Monitor
        w2_status_label.config(
            text=f"Selected → {values[1]} ({values[3]})  |  Hostname: {values[2]}  |  Shelf: {values[7]}  |  Serial: {values[5]}",
            fg="#1a5276")

def w2_unstage_from_warehouse(event):
    item_id = tree_w2_warehouse.identify_row(event.y)
    if not item_id:
        return
    values = tree_w2_warehouse.item(item_id, "values")
    if not values:
        return
    # values: QR, Set ID, Hostname, Equipment Type, Brand/Model, Serial, Checked By, Shelf, Status, Remarks, Date
    set_id, hostname, eq_type, brand, serial, checked_by, shelf, status, remarks = values[1], values[2], values[3], values[4], values[5], values[6], values[7], values[8], values[9]
    if not messagebox.askyesno("Move to Staging",
        f"Move {eq_type} ({set_id}) back to staging?\n\nShelf: {shelf}"):
        return

    df_w2 = load_items_w2()
    match = df_w2[(df_w2["Set ID"] == set_id) & (df_w2["Equipment Type"] == eq_type)]
    if match.empty:
        messagebox.showerror("Error", "Item not found in warehouse"); return

    qr_label = f"{set_id}-{eq_type}"
    delete_qr(qr_label, warehouse=2)
    df_w2 = df_w2.drop(match.index).reset_index(drop=True)
    save_warehouse_2(df_w2, load_shelves_w2())

    # Add back as a staged set (single-item set)
    staged_sets.append({"set_id": set_id, "items": [{
        "Equipment Type": eq_type,
        "Hostname":       hostname,
        "Brand/Model":    brand,
        "Serial Number":  serial,
        "Checked By":     checked_by,
        "Shelf":          shelf,
        "Status":         status,
        "Remarks":        remarks,
    }]})
    save_log("UNSTAGE", f"[W2] Set: {set_id} | Item: {eq_type} | Shelf: {shelf}")
    messagebox.showinfo("Moved", f"{eq_type} ({set_id}) moved back to staging")
    update_w2_staged_display()
    # Clear the pull search entry and any status text so w2_show_warehouse
    # doesn't re-apply a stale filter that makes the table appear blank
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
        "Brand/Model": str(item_row.get("Brand/Model", "")),
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

def w2_undo_pull(event):
    item_id = tree_w2_pullouts.identify_row(event.y)
    if not item_id:
        return
    values = tree_w2_pullouts.item(item_id, "values")
    if not values:
        return
    set_id, hostname, eq_type, brand, shelf, status, remarks = values[0], values[1], values[2], values[3], values[4], values[5], values[6]
    if not messagebox.askyesno("Undo Pull",
        f"Restore {eq_type} ({set_id}) back to Warehouse 2?\nShelf: {shelf}"):
        return

    df_w2 = load_items_w2()
    df_po2 = load_pullouts_w2()

    match = df_po2[(df_po2["Set ID"] == set_id) & (df_po2["Equipment Type"] == eq_type)]
    if match.empty:
        messagebox.showerror("Error", "Record not found in pull history"); return

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
        "Brand/Model":    str(pull_row.get("Brand/Model", "")),
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
    messagebox.showinfo("Restored", f"{eq_type} from {set_id} restored to Warehouse 2")
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
                 font=("Arial", 10, "bold"), bg="#27ae60", fg="white",
                 padx=10, pady=6).grid(row=0, column=0, columnspan=2, sticky="we")

        COL_WIDTHS = [12, 18, 20, 18, 16, 16, 13, 20]
        HEADERS    = ["Equipment", "Hostname", "Brand / Model", "Serial Number",
                      "Checked By", "Shelf", "Status", "Remarks"]

        outer = tk.Frame(edit_win, padx=10, pady=5)
        outer.grid(row=1, column=0, columnspan=2)

        hdr_frame = tk.Frame(outer, bg="#dce3f0")
        hdr_frame.pack(fill="x")
        for col, (h, cw) in enumerate(zip(HEADERS, COL_WIDTHS)):
            tk.Label(hdr_frame, text=h, font=("Arial", 9, "bold"), width=cw, anchor="w",
                     bg="#dce3f0", padx=5).grid(row=0, column=col, padx=5, pady=4, sticky="w")

        ROW_COLORS = ("#ffffff", "#f0f4ff")
        row_widgets = {}
        for r, item in enumerate(items):
            bg = ROW_COLORS[r % 2]
            eq_type = item["Equipment Type"]
            row_bg  = tk.Frame(outer, bg=bg, bd=1, relief="flat")
            row_bg.pack(fill="x", pady=1)
            tk.Label(row_bg, text=eq_type, width=COL_WIDTHS[0], anchor="w",
                     bg=bg, font=("Arial", 9, "bold")).grid(row=0, column=0, padx=5, pady=8, sticky="w")
            hostname_e = tk.Entry(row_bg, width=COL_WIDTHS[1], font=("Arial", 9))
            hostname_e.insert(0, item.get("Hostname", ""))
            hostname_e.grid(row=0, column=1, padx=5, pady=8)
            brand_e = tk.Entry(row_bg, width=COL_WIDTHS[2], font=("Arial", 9))
            brand_e.insert(0, item.get("Brand/Model", ""))
            brand_e.grid(row=0, column=2, padx=5, pady=8)
            serial_e = tk.Entry(row_bg, width=COL_WIDTHS[3], font=("Arial", 9))
            serial_e.insert(0, item.get("Serial Number", ""))
            serial_e.grid(row=0, column=3, padx=5, pady=8)
            checked_e = tk.Entry(row_bg, width=COL_WIDTHS[4], font=("Arial", 9))
            checked_e.insert(0, item.get("Checked By", ""))
            checked_e.grid(row=0, column=4, padx=5, pady=8)
            shelf_v = tk.StringVar(value=item.get("Shelf", ""))
            ttk.Combobox(row_bg, textvariable=shelf_v, values=shelf_list,
                         width=COL_WIDTHS[5], state="readonly",
                         font=("Arial", 9)).grid(row=0, column=5, padx=5, pady=8)
            status_v = tk.StringVar(value=item.get("Status", ""))
            ttk.Combobox(row_bg, textvariable=status_v,
                         values=["No Issue", "Minimal", "Defective"],
                         width=COL_WIDTHS[6], state="readonly",
                         font=("Arial", 9)).grid(row=0, column=6, padx=5, pady=8)
            remarks_e = tk.Entry(row_bg, width=COL_WIDTHS[7], font=("Arial", 9))
            remarks_e.insert(0, item.get("Remarks", ""))
            remarks_e.grid(row=0, column=7, padx=5, pady=8)
            row_widgets[eq_type] = (hostname_e, brand_e, serial_e, checked_e,
                                    shelf_v, status_v, remarks_e)

        error_lbl = tk.Label(edit_win, text="", fg="red", font=("Arial", 8))
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
            for eq_type, (hostname_e, brand_e, serial_e, checked_e,
                          shelf_v, status_v, remarks_e) in row_widgets.items():
                hn  = hostname_e.get().strip()
                br  = brand_e.get().strip()
                sn  = serial_e.get().strip()
                cb  = checked_e.get().strip()
                sh  = shelf_v.get().strip()
                st  = status_v.get().strip()
                rm  = remarks_e.get().strip()
                if not hn:  error_lbl.config(text=f"Hostname required for {eq_type}");       return
                if not br:  error_lbl.config(text=f"Brand/Model required for {eq_type}");    return
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
                    "Brand/Model": br, "Serial Number": sn,
                    "Checked By": cb, "Shelf": sh,
                    "Status": st,     "Remarks": rm,
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
        search_cols = ["Hostname", "Brand/Model", "Serial Number", "Checked By", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]
        mask = False
        for col in search_cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(reason, na=False)
        df = df[mask]
    df = _filter_by_date(df, date_from, date_to)
    # Populate the pull reason dropdown with known values from data
    all_reasons = sorted(load_pullouts_w2()["Pull Reason"].dropna().unique().tolist())
    w2_pull_reason_filter_entry["values"] = [""] + all_reasons
    for _, row in df.iterrows():
        tree_w2_pullouts.insert("", "end", values=tuple(
            row.get(c, "") for c in ["Set ID", "Hostname", "Equipment Type", "Brand/Model", "Shelf", "Status", "Remarks", "Pull Reason", "Date"]))
    w2_search_label.config(text=f"Pull History Filtered: {len(df)} result(s)", fg="darkorange")

# ========== QR LABEL PDF ==========

def generate_qr_pdf(items_batch):
    from fpdf import FPDF
    PAGE_W, LABEL_W, COLS = 210, 54, 3
    MARGIN_X, MARGIN_Y, ROW_GAP = 12, 10, 4
    GAP_X = (PAGE_W - (COLS * LABEL_W) - (2 * MARGIN_X)) / (COLS - 1)
    QR_SIZE, QR_PAD_TOP, QR_PAD_BTM = 20, 3, 3   # mm above/below QR image
    LINE_H, FIELD_PAD_BTM = 4.8, 2                # line height + bottom padding

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    col = row = 0

    for item in items_batch:
        is_w2 = item.get('_warehouse', 1) == 2

        # Build field list
        if is_w2:
            fields = [
                ("Set ID:",      str(item.get("Set ID", ""))),
                ("Type:",        str(item.get("Equipment Type", ""))),
                ("Brand/Model:", str(item.get("Brand/Model", ""))),
                ("Serial No:",   str(item.get("Serial Number", ""))),
                ("Shelf:",       str(item.get("Shelf", ""))),
                ("Status:",      str(item.get("Status", ""))),
                ("Remarks:",     str(item.get("Remarks", ""))),
                ("Date:",        datetime.now().strftime("%Y-%m-%d")),
            ]
            qr_key = f"{item.get('Set ID', '')}-{item.get('Equipment Type', '')}"
            path = qr_path_for(qr_key, warehouse=2)
        else:
            fields = [
                ("Hostname:",    str(item.get("Hostname", ""))),
                ("Brand/Model:", str(item.get("Brand/Model", ""))),
                ("Serial No:",   str(item.get("Serial Number", ""))),
                ("Checked By:",  str(item.get("Checked By", ""))),
                ("Shelf:",       str(item.get("Shelf", ""))),
                ("Status:",      str(item.get("Status", ""))),
                ("Remarks:",     str(item.get("Remarks", ""))),
                ("Date:",        datetime.now().strftime("%Y-%m-%d")),
            ]
            path = qr_path_for(item['Hostname'], warehouse=1)

        # Compute exact label height: QR block + text block + bottom padding
        LABEL_H = QR_PAD_TOP + QR_SIZE + QR_PAD_BTM + len(fields) * LINE_H + FIELD_PAD_BTM

        x = MARGIN_X + col * (LABEL_W + GAP_X)
        y = MARGIN_Y + row * (LABEL_H + ROW_GAP)
        if y + LABEL_H > 297 - MARGIN_Y:
            pdf.add_page(); col = row = 0
            x = MARGIN_X
            y = MARGIN_Y

        # Draw border tightly around content
        pdf.set_draw_color(150, 150, 150)
        pdf.rect(x, y, LABEL_W, LABEL_H)

        # QR image centred at top of label
        if os.path.exists(path):
            pdf.image(path, x=x + (LABEL_W - QR_SIZE) / 2, y=y + QR_PAD_TOP, w=QR_SIZE, h=QR_SIZE)

        # Text fields
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

    date_str = datetime.now().strftime("%Y-%m-%d")
    if items_batch and items_batch[0].get('_warehouse', 1) == 2:
        output_folder = QR_LABELS_FOLDER_W2
        set_ids = list(dict.fromkeys(item.get("Set ID", "") for item in items_batch))
        label_name = "_".join(set_ids[:3])
        if len(set_ids) > 3:
            label_name += f"_and_{len(set_ids)-3}_more"
    else:
        output_folder = QR_LABELS_FOLDER_W1
        os.makedirs(output_folder, exist_ok=True)
        existing = [f for f in os.listdir(output_folder) if f.startswith("BATCH_") and f.endswith(".pdf")]
        batch_nums = []
        for f in existing:
            try:
                num = int(f.split("_")[1])
                batch_nums.append(num)
            except (IndexError, ValueError):
                pass
        next_batch = max(batch_nums, default=0) + 1
        label_name = f"BATCH_{next_batch}"

    os.makedirs(output_folder, exist_ok=True)
    safe_name = label_name.replace(" ", "_").replace("/", "-")
    pdf_path = os.path.join(output_folder, f"{safe_name}_{date_str}.pdf")
    pdf.output(pdf_path)
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
    tk.Label(hdr, text="QR Label Files", font=("Arial", 10, "bold"),
             bg="#2c3e50", fg="white").pack(side="left", padx=10, pady=6)
    sel_count_lbl = tk.Label(hdr, text="", font=("Arial", 9),
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
        tk.Label(hdr_row, text="✔", width=3, bg="#dce3f0", font=("Arial", 9, "bold")).pack(side="left", padx=(6,0))
        for txt, w in [("Warehouse", 110), ("Filename", 240), ("Created", 155), ("Size", 60)]:
            tk.Label(hdr_row, text=txt, width=w//7, bg="#dce3f0",
                     font=("Arial", 9, "bold"), anchor="w").pack(side="left", padx=4, pady=5)

        idx = 0
        for warehouse_label, folder in filtered:
            if not os.path.exists(folder):
                continue
            for f in sorted([fn for fn in os.listdir(folder) if fn.endswith(".pdf")], reverse=True):
                full_path = os.path.join(folder, f)
                size_kb = round(os.path.getsize(full_path) / 1024, 1)
                try:
                    date_part = f.replace(".pdf", "")[-10:]
                    file_dt = datetime.strptime(date_part, "%Y-%m-%d")
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
                                   font=("Arial", 9))
                    lbl.pack(side="left", padx=4, pady=4)
                    lbl.bind("<Button-1>", _make_toggle(var))
                row_fr.bind("<Button-1>", _make_toggle(var))

                row_data.append((iid, full_path, warehouse_label, f, date_str, f"{size_kb} kb", var))
                idx += 1

        if idx == 0:
            tk.Label(inner_cl, text="No QR label PDF files found.", fg="gray",
                     font=("Arial", 10), bg="white").pack(pady=30)

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
        for _, full_path, _, fname, _, _, _ in chosen:
            try:
                if os.path.exists(full_path):
                    os.remove(full_path)
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
                               selectmode="single", exportselection=False, font=("Arial", 9))
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
    global current_user, session_start

    switch_win = tk.Toplevel(root)
    switch_win.title("Change User")
    switch_win.geometry("300x200")
    switch_win.resizable(False, False)
    switch_win.transient(root)

    tk.Label(switch_win, text="Enter new user name:", font=("Arial", 10)).pack(pady=(20, 5))
    name_var = tk.StringVar()
    name_entry = tk.Entry(switch_win, textvariable=name_var, width=25, font=("Arial", 10))
    name_entry.pack(pady=5)
    error_label = tk.Label(switch_win, text="", fg="red", font=("Arial", 8))
    error_label.pack()

    def validate_name(name):
        if not name:
            return "Please enter a name to continue"
        if not re.match(r'^[A-Za-z][A-Za-z ]*$', name):
            return "Name must contain letters and spaces only"
        if '  ' in name:
            return "Name cannot contain consecutive spaces"
        return None

    def on_key_release(event):
        val = name_var.get()
        cleaned = re.sub(r'[^A-Za-z ]', '', val)
        if cleaned != val:
            name_var.set(cleaned)
            name_entry.icursor(len(cleaned))
            error_label.config(text="Numbers and special characters are not allowed")
        else:
            error_label.config(text="")

    def apply_switch():
        global current_user, session_start
        new_name = name_var.get().strip()
        error = validate_name(new_name)
        if error:
            error_label.config(text=error); return
        save_log("LOGOUT", f"Session ended for '{current_user}'")
        current_user = new_name
        session_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        user_label.config(text=f"👤  {current_user}")
        session_label.config(text=f"Session started: {session_start}")
        save_log("LOGIN", f"Session started by '{current_user}'")
        switch_win.destroy()
        messagebox.showinfo("User Changed", f"Successfully switched to: {current_user}")

    name_entry.bind("<KeyRelease>", on_key_release)
    tk.Button(switch_win, text="Switch", command=apply_switch, width=15).pack(pady=10)
    switch_win.bind("<Return>", lambda e: apply_switch())
    switch_win.update_idletasks()
    switch_win.grab_set()
    switch_win.focus_force()
    name_entry.focus_set()
    switch_win.wait_window()

# ========== LOGIN ==========

def show_login():
    global current_user, session_start
    login_win = tk.Tk()
    login_win.title("Warehouse System — Login")
    login_win.geometry("320x180")
    login_win.resizable(False, False)
    login_win.eval('tk::PlaceWindow . center')

    tk.Label(login_win, text="Warehouse System", font=("Arial", 13, "bold")).pack(pady=(20, 5))
    tk.Label(login_win, text="Who is using the system?", font=("Arial", 9)).pack()

    name_var = tk.StringVar()
    name_entry = tk.Entry(login_win, textvariable=name_var, width=25, font=("Arial", 10))
    name_entry.pack(pady=10)
    name_entry.focus()

    error_label = tk.Label(login_win, text="", fg="red", font=("Arial", 8))
    error_label.pack()

    def attempt_login():
        global current_user, session_start
        name = name_var.get().strip()
        if not name:
            error_label.config(text="Please enter your name to continue"); return
        if not re.match(r'^[A-Za-z][A-Za-z ]*$', name):
            error_label.config(text="Name must contain letters and spaces only"); return
        if '  ' in name:
            error_label.config(text="Name cannot contain consecutive spaces"); return
        current_user = name
        session_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        login_win.quit()

    tk.Button(login_win, text="LOGIN", command=attempt_login, width=15).pack(pady=5)
    name_entry.bind("<Return>", lambda e: attempt_login())

    def on_close():
        if messagebox.askyesno("Exit", "Exit the Warehouse System?", parent=login_win):
            login_win.destroy()
            import sys; sys.exit(0)

    login_win.protocol("WM_DELETE_WINDOW", on_close)
    login_win.mainloop()
    login_win.destroy()
    initialize_log()
    save_log("LOGIN", f"Session started by '{current_user}'")

# ========== UI SETUP ==========

show_login()

root = tk.Tk()
root.title("Warehouse Management System")
root.geometry("1280x780")
root.eval('tk::PlaceWindow . center')

# ── User bar ──────────────────────────────────────────────
user_bar = tk.Frame(root, bg="#2c3e50", height=28)
user_bar.pack(fill="x")

clock_label = tk.Label(user_bar, text="", bg="#2c3e50", fg="#95a5a6", font=("Arial", 8))
clock_label.pack(side="right", padx=10, pady=4)

tk.Button(user_bar, text="Change User", command=switch_user,
          bg="#34495e", fg="white", bd=0, padx=10).pack(side="right", padx=10, pady=2)
tk.Button(user_bar, text="📋 Activity Log", command=open_activity_log,
          bg="#34495e", fg="white", bd=0, padx=10).pack(side="right", padx=4, pady=2)

user_label = tk.Label(user_bar, text=f"👤  {current_user}", bg="#2c3e50", fg="white", font=("Arial", 9, "bold"))
user_label.pack(side="left", padx=10, pady=4)

session_label = tk.Label(user_bar, text=f"Session started: {session_start}", bg="#2c3e50", fg="#95a5a6", font=("Arial", 8))
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
w1_row1.pack(fill="x")

# Item Management
input_frame = tk.LabelFrame(w1_row1, text="Item Management", padx=10, pady=5)
input_frame.pack(side="left", fill="both", padx=5)

tk.Label(input_frame, text="Hostname").grid(row=0, column=0, sticky="w")
hostname_entry = tk.Entry(input_frame, width=22); hostname_entry.grid(row=0, column=1, pady=3)
tk.Label(input_frame, text="Brand / Model").grid(row=1, column=0, sticky="w")
brand_entry = tk.Entry(input_frame, width=22); brand_entry.grid(row=1, column=1, pady=3)
tk.Label(input_frame, text="Serial Number").grid(row=2, column=0, sticky="w")
serial_entry = tk.Entry(input_frame, width=22); serial_entry.grid(row=2, column=1, pady=3)
tk.Label(input_frame, text="Checked By").grid(row=3, column=0, sticky="w")
checked_by_entry = tk.Entry(input_frame, width=22); checked_by_entry.grid(row=3, column=1, pady=3)

tk.Label(input_frame, text="Shelf").grid(row=4, column=0, sticky="w")
shelf_var = tk.StringVar()
shelf_dropdown = ttk.Combobox(input_frame, textvariable=shelf_var, width=19)
shelf_dropdown.grid(row=4, column=1, pady=3)

tk.Label(input_frame, text="Status").grid(row=5, column=0, sticky="w")
remarks_var = tk.StringVar()
ttk.Combobox(input_frame, textvariable=remarks_var, values=["No Issue", "Minimal", "Defective"], width=19, state="readonly").grid(row=5, column=1, pady=3)

tk.Label(input_frame, text="Remarks").grid(row=6, column=0, sticky="w")
remarks_text_var = tk.StringVar()
tk.Entry(input_frame, textvariable=remarks_text_var, width=22).grid(row=6, column=1, pady=3)

crud_frame = tk.Frame(input_frame)
crud_frame.grid(row=7, column=0, columnspan=2, pady=5)
tk.Button(crud_frame, text="PUT",    command=put_item,    width=8).grid(row=0, column=0, padx=3)
tk.Button(crud_frame, text="UPDATE", command=update_item, width=8).grid(row=0, column=1, padx=3)
tk.Button(crud_frame, text="↻",      command=reset_ui,    width=3).grid(row=0, column=2, padx=3)

tk.Label(input_frame, text="Staged Items (Click to Edit)", fg="green", font=("Arial", 9, "bold")).grid(row=8, column=0, columnspan=2, sticky="w")
staged_listbox = tk.Listbox(input_frame, height=4, width=32)
staged_listbox.grid(row=9, column=0, columnspan=2, sticky="we", pady=3)
staged_listbox.bind("<<ListboxSelect>>", select_staged_item)

staging_btn_frame = tk.Frame(input_frame)
staging_btn_frame.grid(row=10, column=0, columnspan=2, pady=3)
tk.Button(staging_btn_frame, text="CLEAR ITEMS",   command=remove_from_staging, width=13).pack(side="left", padx=2)
tk.Button(staging_btn_frame, text="PUT WAREHOUSE", command=put_warehouse,       width=13).pack(side="left", padx=2)

# Shelf Controls W1
shelf_mid_frame = tk.Frame(w1_row1)
shelf_mid_frame.pack(side="left", fill="both", expand=True, padx=5)
shelf_control_frame = tk.LabelFrame(shelf_mid_frame, text="Shelf Control & Management", padx=10, pady=5)
shelf_control_frame.pack(fill="x")

status_control_frame = tk.LabelFrame(shelf_control_frame, text="Status Control", padx=8, pady=5)
status_control_frame.pack(fill="x", pady=(0, 5))
shelf_control_var = tk.StringVar()
shelf_control_dropdown = ttk.Combobox(status_control_frame, textvariable=shelf_control_var, width=22, state="readonly")
shelf_control_dropdown.pack(side="left", padx=5)
tk.Button(status_control_frame, text="SET FULL",      command=lambda: set_shelf_status("FULL"),      width=10).pack(side="left", padx=3)
tk.Button(status_control_frame, text="SET AVAILABLE", command=lambda: set_shelf_status("AVAILABLE"), width=12).pack(side="left", padx=3)
tk.Button(status_control_frame, text="↻",             command=reset_shelf_control,                   width=3).pack(side="left", padx=3)

add_remove_frame = tk.LabelFrame(shelf_control_frame, text="Add / Remove", padx=8, pady=5)
add_remove_frame.pack(fill="x")
remove_shelf_var = tk.StringVar()
remove_shelf_dropdown = ttk.Combobox(add_remove_frame, textvariable=remove_shelf_var, width=22)
remove_shelf_dropdown.pack(side="left", padx=5)
tk.Button(add_remove_frame, text="ADD",    command=add_shelf,           ).pack(side="left", padx=3)
tk.Button(add_remove_frame, text="REMOVE", command=remove_shelf,         ).pack(side="left", padx=3)
tk.Button(add_remove_frame, text="↻",      command=reset_shelf_addition, width=3).pack(side="left", padx=3)

# View W1
view_frame = tk.LabelFrame(w1_row1, text="View", padx=10, pady=5)
view_frame.pack(side="right", fill="both", padx=5)
for text, cmd in [
    ("SHOW WAREHOUSE", show_warehouse),
    ("SHELF STATUS",   show_available),
    ("PULL HISTORY",   show_pullouts),
    ("STORED QR",      show_qr_codes),
    ("QR LABELS",      lambda: open_label_manager(warehouse=1)),
]:
    tk.Button(view_frame, text=text, command=cmd, width=15).pack(anchor="w", pady=3)

# Search & Filter W1
w1_pullout_frame = tk.LabelFrame(w1_main, text="Warehouse 1", padx=10, pady=8)
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
ttk.Combobox(w1_sf_row1, textvariable=pull_remarks_var, values=["No Issue", "Minimal", "Defective"], width=14, state="readonly").pack(side="left", padx=(0, 8))

tk.Button(w1_sf_row1, text="🔍", command=search_item,       width=3).pack(side="left", padx=3)
tk.Button(w1_sf_row1, text="↻",  command=clear_pull_filters, width=3).pack(side="left", padx=3)

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
tk.Button(w1_po_row1, text="↻",  command=reset_pull_out,      width=3).pack(side="left", padx=2)
tk.Button(w1_po_row1, text="WAREHOUSE PULL", command=pull_item, width=16).pack(side="left", padx=(10, 3))

# Status bar W1
w1_status_bar = tk.Frame(w1_main)
w1_status_bar.pack(fill="x")
w1_full_label   = tk.Label(w1_status_bar, text="FULL Shelves: None", fg="red");  w1_full_label.pack(side="left", padx=10)
w1_search_label = tk.Label(w1_status_bar, text="", fg="blue");                   w1_search_label.pack(side="left", padx=10)
w1_status_label = tk.Label(w1_status_bar, text="", fg="green");                  w1_status_label.pack(side="left", padx=10)

# Tables W1
w1_table_frame = tk.Frame(w1_main)
w1_table_frame.pack(fill="both", expand=True, pady=5)

tree_warehouse = ttk.Treeview(w1_table_frame, columns=("C1","C2","C3","C4","C5","C6","C7","C8","C9"), show='headings')
for col, text, width in zip(("C1","C2","C3","C4","C5","C6","C7","C8","C9"),
    ("QR","Hostname","Brand/Model","Serial Number","Checked By","Shelf","Status","Remarks","Date"),
    (180,130,120,110,105,120,90,130,140)):
    tree_warehouse.heading(col, text=text); tree_warehouse.column(col, width=width)
tree_warehouse.bind("<<TreeviewSelect>>", select_item)
tree_warehouse.bind("<Double-1>", unstage_from_warehouse)

tree_available = ttk.Treeview(w1_table_frame, columns=("C1","C2","C3"), show='headings')
for col, text, width in zip(("C1","C2","C3"), ("Shelf","Status","Date_Full"), (250,150,200)):
    tree_available.heading(col, text=text); tree_available.column(col, width=width)

tree_pullouts = ttk.Treeview(w1_table_frame, columns=("C1","C2","C3","C4","C5","C6","C7"), show='headings')
for col, text, width in zip(("C1","C2","C3","C4","C5","C6","C7"),
    ("Hostname","Brand/Model","Shelf","Status","Remarks","Pull Reason","Date"), (140,120,120,90,150,200,150)):
    tree_pullouts.heading(col, text=text); tree_pullouts.column(col, width=width)
tree_pullouts.bind("<Double-1>", undo_pull)

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
w2_row1.pack(fill="x")

# Equipment selection + staging panel
w2_input_frame = tk.LabelFrame(w2_row1, text="Set Staging", padx=10, pady=5)
w2_input_frame.pack(side="left", fill="both", padx=5)

tk.Label(w2_input_frame, text="Select Equipment:", font=("Arial", 9, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

w2_equip_vars = {}
for i, eq in enumerate(EQUIPMENT_TYPES):
    var = tk.BooleanVar()
    w2_equip_vars[eq] = var
    tk.Checkbutton(w2_input_frame, text=eq, variable=var, width=10, anchor="w").grid(
        row=1 + i // 2, column=i % 2, sticky="w", padx=4)

tk.Button(w2_input_frame, text="BUILD SET", command=w2_build_set,
          bg="#2980b9", fg="white", width=28).grid(row=3, column=0, columnspan=2, pady=(8, 4))

tk.Label(w2_input_frame, text="Staged Sets", fg="green", font=("Arial", 9, "bold")).grid(row=4, column=0, columnspan=2, sticky="w")
w2_staged_listbox = tk.Listbox(w2_input_frame, height=5, width=34)
w2_staged_listbox.grid(row=5, column=0, columnspan=2, sticky="we", pady=3)

w2_stage_btns = tk.Frame(w2_input_frame)
w2_stage_btns.grid(row=6, column=0, columnspan=2, pady=3)
tk.Button(w2_stage_btns, text="CLEAR SETS",    command=w2_remove_staged_set, width=13).pack(side="left", padx=2)
tk.Button(w2_stage_btns, text="PUT WAREHOUSE", command=w2_put_warehouse,     width=13).pack(side="left", padx=2)

# Item Management (UPDATE / DELETE) for W2
w2_item_mgmt_frame = tk.LabelFrame(w2_input_frame, text="Item Management", padx=6, pady=4)
w2_item_mgmt_frame.grid(row=7, column=0, columnspan=2, pady=(6, 0), sticky="we")
tk.Label(w2_item_mgmt_frame,
         text="Select a row in the staged sets table,\nthen use buttons below:",
         font=("Arial", 8), fg="gray", justify="left").pack(anchor="w", pady=(0, 4))
w2_item_action_btns = tk.Frame(w2_item_mgmt_frame)
w2_item_action_btns.pack()
tk.Button(w2_item_action_btns, text="UPDATE ITEM", command=w2_update_item,
          bg="#1a5276", fg="white", width=13).pack(side="left", padx=2)

# Shelf Controls W2
w2_shelf_mid = tk.Frame(w2_row1)
w2_shelf_mid.pack(side="left", fill="both", expand=True, padx=5)
w2_shelf_ctrl_frame = tk.LabelFrame(w2_shelf_mid, text="Shelf Control & Management", padx=10, pady=5)
w2_shelf_ctrl_frame.pack(fill="x")

w2_status_ctrl = tk.LabelFrame(w2_shelf_ctrl_frame, text="Status Control", padx=8, pady=5)
w2_status_ctrl.pack(fill="x", pady=(0, 5))
w2_shelf_control_var = tk.StringVar()
w2_shelf_control_dropdown = ttk.Combobox(w2_status_ctrl, textvariable=w2_shelf_control_var, width=22, state="readonly")
w2_shelf_control_dropdown.pack(side="left", padx=5)
tk.Button(w2_status_ctrl, text="SET FULL",      command=lambda: w2_set_shelf_status("FULL"),      width=10).pack(side="left", padx=3)
tk.Button(w2_status_ctrl, text="SET AVAILABLE", command=lambda: w2_set_shelf_status("AVAILABLE"), width=12).pack(side="left", padx=3)
tk.Button(w2_status_ctrl, text="↻",             command=w2_reset_shelf_control,                   width=3).pack(side="left", padx=3)

w2_add_remove = tk.LabelFrame(w2_shelf_ctrl_frame, text="Add / Remove", padx=8, pady=5)
w2_add_remove.pack(fill="x")
w2_remove_shelf_var = tk.StringVar()
w2_remove_shelf_dropdown = ttk.Combobox(w2_add_remove, textvariable=w2_remove_shelf_var, width=22)
w2_remove_shelf_dropdown.pack(side="left", padx=5)
tk.Button(w2_add_remove, text="ADD",    command=w2_add_shelf,           ).pack(side="left", padx=3)
tk.Button(w2_add_remove, text="REMOVE", command=w2_remove_shelf,         ).pack(side="left", padx=3)
tk.Button(w2_add_remove, text="↻",      command=w2_reset_shelf_addition, width=3).pack(side="left", padx=3)

# View W2
w2_view_frame = tk.LabelFrame(w2_row1, text="View", padx=10, pady=5)
w2_view_frame.pack(side="right", fill="both", padx=5)
for text, cmd in [
    ("SHOW WAREHOUSE", w2_show_warehouse),
    ("SHELF STATUS",   w2_show_available),
    ("PULL HISTORY",   w2_show_pullouts),
    ("STORED QR",      w2_show_qr_codes),
    ("QR LABELS",      lambda: open_label_manager(warehouse=2)),
]:
    tk.Button(w2_view_frame, text=text, command=cmd, width=15).pack(anchor="w", pady=3)

# Search & Filter W2
w2_pullout_frame = tk.LabelFrame(w2_main, text="Warehouse 2", padx=10, pady=8)
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

tk.Button(w2_sf_row1, text="🔍", command=w2_search_item,  width=3).pack(side="left", padx=3)
tk.Button(w2_sf_row1, text="↻",  command=w2_clear_filters, width=3).pack(side="left", padx=3)

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
tk.Button(w2_po_row1, text="↻",  command=w2_reset_pull_out,      width=3).pack(side="left", padx=2)
tk.Button(w2_po_row1, text="WAREHOUSE PULL", command=w2_pull_item, width=16).pack(side="left", padx=(10, 3))

# Status bar W2
w2_status_bar = tk.Frame(w2_main)
w2_status_bar.pack(fill="x")
w2_full_label   = tk.Label(w2_status_bar, text="FULL Shelves: None", fg="red");  w2_full_label.pack(side="left", padx=10)
w2_search_label = tk.Label(w2_status_bar, text="", fg="blue");                   w2_search_label.pack(side="left", padx=10)
w2_status_label = tk.Label(w2_status_bar, text="", fg="green");                  w2_status_label.pack(side="left", padx=10)

# Tables W2
w2_table_frame = tk.Frame(w2_main)
w2_table_frame.pack(fill="both", expand=True, pady=5)

tree_w2_warehouse = ttk.Treeview(w2_table_frame,
    columns=("C1","C2","C3","C4","C5","C6","C7","C8","C9","C10","C11"), show='headings')
for col, text, width in zip(
    ("C1","C2","C3","C4","C5","C6","C7","C8","C9","C10","C11"),
    ("QR","Set ID","Hostname","Equipment Type","Brand/Model","Serial Number","Checked By","Shelf","Status","Remarks","Date"),
    (170,85,110,105,125,105,105,115,85,140,130)):
    tree_w2_warehouse.heading(col, text=text); tree_w2_warehouse.column(col, width=width)
tree_w2_warehouse.bind("<<TreeviewSelect>>", w2_select_item)
tree_w2_warehouse.bind("<Double-1>", w2_unstage_from_warehouse)

tree_w2_available = ttk.Treeview(w2_table_frame, columns=("C1","C2","C3"), show='headings')
for col, text, width in zip(("C1","C2","C3"), ("Shelf","Status","Date_Full"), (250,150,200)):
    tree_w2_available.heading(col, text=text); tree_w2_available.column(col, width=width)

tree_w2_pullouts = ttk.Treeview(w2_table_frame,
    columns=("C1","C2","C3","C4","C5","C6","C7","C8","C9"), show='headings')
for col, text, width in zip(
    ("C1","C2","C3","C4","C5","C6","C7","C8","C9"),
    ("Set ID","Hostname","Equipment Type","Brand/Model","Shelf","Status","Remarks","Pull Reason","Date"),
    (85,110,115,130,110,80,160,180,130)):
    tree_w2_pullouts.heading(col, text=text); tree_w2_pullouts.column(col, width=width)
tree_w2_pullouts.bind("<Double-1>", w2_undo_pull)

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