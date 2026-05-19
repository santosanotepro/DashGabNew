import json
import os
import datetime
import tkinter as tk
from tkinter import ttk, messagebox

def load_config(config_path: str) -> dict:
    if not os.path.exists(config_path):
        return {}
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        return {}

def save_config(config: dict, config_path: str):
    config["version"] = "2.0"
    config["created"] = datetime.datetime.now().isoformat()
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

def auto_detect_column_types(df) -> dict:
    types = {}
    for col in df.columns:
        col_name = str(col).strip()
        ctype = "text"
        role = ""

        try:
            col_data = df[col_name].dropna()
            import pandas as pd
            pd.to_numeric(col_data)
            is_numeric = True
        except:
            is_numeric = False

        unique_vals = df[col_name].nunique()
        lower_col = col_name.lower()

        if is_numeric:
            ctype = "number"
            if "kwl" in lower_col or "kanwil" in lower_col:
                role = "kanwil_code"
            elif "kpp" in lower_col:
                role = "kpp_code"
        elif lower_col in ["lat", "latitude"]:
            ctype = "koordinat"
            role = "lat"
        elif lower_col in ["long", "longitude", "lng"]:
            ctype = "koordinat"
            role = "long"
        elif unique_vals < 20:
            ctype = "kategori"

        types[col_name] = {
            "type": ctype,
            "include": True,
            "role": role if role else None
        }
    return types

class SetupDialog(tk.Toplevel):
    def __init__(self, parent, sheet_name, df, current_config=None):
        super().__init__(parent)
        self.title(f"Setup Kolom: {sheet_name}")
        self.geometry("800x600")
        self.sheet_name = sheet_name
        self.df = df
        self.result = None
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#f0f2f5")

        hdr = tk.Frame(self, bg="#0a2346", height=40)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"Konfigurasi Sheet: {sheet_name}", bg="#0a2346", fg="#e8c96a", font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=10, pady=10)

        body = tk.Frame(self, bg="#f0f2f5", padx=10, pady=10)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(body, text="Preview Data (5 baris pertama):", bg="#f0f2f5", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        preview_frame = tk.Frame(body)
        preview_frame.pack(fill=tk.X, pady=(0, 15))

        cols = tuple(str(c) for c in df.columns)
        self.tree = ttk.Treeview(preview_frame, columns=cols, show='headings', height=5)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in df.head(5).iterrows():
            self.tree.insert("", tk.END, values=tuple(str(v) for v in row))

        tree_scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=tree_scroll_x.set)

        self.tree.pack(side=tk.TOP, fill=tk.X)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        tk.Label(body, text="Konfigurasi Kolom:", bg="#f0f2f5", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        config_frame = tk.Frame(body, bg="white", bd=1, relief=tk.SOLID)
        config_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        canvas = tk.Canvas(config_frame, bg="white")
        scrollbar = ttk.Scrollbar(config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tk.Label(scrollable_frame, text="✓", width=3, bg="white", font=("Arial", 9, "bold")).grid(row=0, column=0, padx=2, pady=2)
        tk.Label(scrollable_frame, text="Kolom", width=20, anchor=tk.W, bg="white", font=("Arial", 9, "bold")).grid(row=0, column=1, padx=2, pady=2)
        tk.Label(scrollable_frame, text="Tipe", width=15, anchor=tk.W, bg="white", font=("Arial", 9, "bold")).grid(row=0, column=2, padx=2, pady=2)
        tk.Label(scrollable_frame, text="Role", width=15, anchor=tk.W, bg="white", font=("Arial", 9, "bold")).grid(row=0, column=3, padx=2, pady=2)

        self.col_vars = {}
        default_types = auto_detect_column_types(df)
        if current_config and "columns" in current_config:
            for c, props in current_config["columns"].items():
                if c in default_types: default_types[c] = props

        self.type_options = ["text", "number", "kategori", "koordinat"]
        self.role_options = ["", "kanwil_code", "kpp_code", "lat", "long"]

        row_idx = 1
        for col in cols:
            info = default_types.get(col, {"type": "text", "include": True, "role": ""})
            inc_var = tk.BooleanVar(value=info.get("include", True))
            tk.Checkbutton(scrollable_frame, variable=inc_var, bg="white").grid(row=row_idx, column=0, padx=2, pady=2)
            tk.Label(scrollable_frame, text=col, width=20, anchor=tk.W, bg="white").grid(row=row_idx, column=1, padx=2, pady=2)

            type_var = tk.StringVar(value=info.get("type", "text"))
            type_cb = ttk.Combobox(scrollable_frame, textvariable=type_var, values=self.type_options, width=12, state="readonly")
            type_cb.grid(row=row_idx, column=2, padx=2, pady=2)

            role_var = tk.StringVar(value=info.get("role", "") or "")
            role_cb = ttk.Combobox(scrollable_frame, textvariable=role_var, values=self.role_options, width=12, state="readonly")
            role_cb.grid(row=row_idx, column=3, padx=2, pady=2)

            def update_role_state(*args, cb=role_cb, tvar=type_var, rvar=role_var):
                t = tvar.get()
                if t == "number":
                    cb.config(values=["", "kanwil_code", "kpp_code"], state="readonly")
                    if rvar.get() not in ["", "kanwil_code", "kpp_code"]: rvar.set("")
                elif t == "koordinat":
                    cb.config(values=["", "lat", "long"], state="readonly")
                    if rvar.get() not in ["", "lat", "long"]: rvar.set("")
                else:
                    rvar.set("")
                    cb.config(state="disabled")

            type_var.trace("w", update_role_state)
            update_role_state()
            self.col_vars[col] = {"include": inc_var, "type": type_var, "role": role_var}
            row_idx += 1

        lbl_frame = tk.Frame(body, bg="#f0f2f5")
        lbl_frame.pack(fill=tk.X, pady=10)
        tk.Label(lbl_frame, text="Label Dashboard:", bg="#f0f2f5", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.label_var = tk.StringVar(value=current_config.get("label", f"Dashboard {sheet_name}") if current_config else f"Dashboard {sheet_name}")
        tk.Entry(lbl_frame, textvariable=self.label_var, width=40).pack(side=tk.LEFT, padx=10)

        btn_frame = tk.Frame(body, bg="#f0f2f5")
        btn_frame.pack(fill=tk.X, pady=10)
        tk.Button(btn_frame, text="Simpan & Lanjutkan", bg="#1d4ed8", fg="white", font=("Arial", 10, "bold"), command=self.on_save).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="Lewati Sheet Ini", command=self.on_skip).pack(side=tk.RIGHT, padx=5)

    def on_save(self):
        config_cols = {}
        has_kanwil = False
        has_kpp = False
        has_lat = False
        has_long = False
        filters = []

        for col, vars_dict in self.col_vars.items():
            ctype = vars_dict["type"].get()
            role = vars_dict["role"].get()
            inc = vars_dict["include"].get()

            config_cols[col] = {"type": ctype, "include": inc, "role": role if role else None}
            if inc:
                if role == "kanwil_code": has_kanwil = True
                if role == "kpp_code": has_kpp = True
                if role == "lat": has_lat = True
                if role == "long": has_long = True
                if ctype == "kategori": filters.append(col)

        self.result = {
            "label": self.label_var.get(),
            "columns": config_cols,
            "filters": filters,
            "has_kanwil_kpp": (has_kanwil and has_kpp),
            "has_map": (has_lat and has_long)
        }
        self.destroy()

    def on_skip(self):
        self.result = None
        self.destroy()
