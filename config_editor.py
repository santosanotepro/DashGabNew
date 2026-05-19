import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os

from config_manager import load_config, save_config

class ConfigEditorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Config Editor Dashboard DJP")
        self.geometry("600x500")
        self.config_path = tk.StringVar(value="dashboard_config.json")
        self.config_data = {}
        self._build_ui()
        self._load_file()

    def _build_ui(self):
        hdr = tk.Frame(self, bg="#0a2346", height=40)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="Config Editor Dashboard DJP", bg="#0a2346", fg="#e8c96a", font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=10, pady=10)

        body = tk.Frame(self, padx=15, pady=15)
        body.pack(fill=tk.BOTH, expand=True)

        f1 = tk.Frame(body)
        f1.pack(fill=tk.X, pady=5)
        tk.Label(f1, text="File:", width=5, anchor=tk.W).pack(side=tk.LEFT)
        tk.Entry(f1, textvariable=self.config_path, width=40).pack(side=tk.LEFT, padx=5)
        tk.Button(f1, text="Buka", command=self._browse).pack(side=tk.LEFT)
        tk.Button(f1, text="Muat Ulang", command=self._load_file).pack(side=tk.LEFT, padx=5)

        tk.Label(body, text="Sheet yang dikonfigurasi:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(15, 5))
        self.sheet_listbox = tk.Listbox(body, height=6)
        self.sheet_listbox.pack(fill=tk.X, pady=5)
        self.sheet_listbox.bind('<<ListboxSelect>>', self._on_select_sheet)

        self.detail_frame = tk.LabelFrame(body, text="Detail Sheet", padx=10, pady=10)
        self.detail_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.lbl_detail = tk.Label(self.detail_frame, text="Pilih sheet di atas.")
        self.lbl_detail.pack(anchor=tk.W)

        btn_frame = tk.Frame(body)
        btn_frame.pack(fill=tk.X, pady=5)
        tk.Button(btn_frame, text="Hapus Sheet", fg="red", command=self._delete_sheet).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Simpan Perubahan", bg="#1d4ed8", fg="white", command=self._save_changes).pack(side=tk.RIGHT)

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if p:
            self.config_path.set(p)
            self._load_file()

    def _load_file(self):
        p = self.config_path.get()
        self.config_data = load_config(p)
        self.sheet_listbox.delete(0, tk.END)
        if "sheets" in self.config_data:
            for s, cfg in self.config_data["sheets"].items():
                cols_count = len(cfg.get("columns", {}))
                hm = cfg.get("has_map", False)
                self.sheet_listbox.insert(tk.END, f"{s} ({cols_count} kolom, has_map={str(hm).lower()})")
        else:
            self.config_data["sheets"] = {}
        self.lbl_detail.config(text="Pilih sheet di atas.")

    def _on_select_sheet(self, evt):
        sel = self.sheet_listbox.curselection()
        if not sel: return
        s = self.sheet_listbox.get(sel[0]).split(" (")[0]
        if s in self.config_data["sheets"]:
            cfg = self.config_data["sheets"][s]
            detail_text = f"Label: {cfg.get('label', '')}\nMap: {cfg.get('has_map')}, Kanwil/KPP: {cfg.get('has_kanwil_kpp')}\n\nKolom:\n"
            for c, ccfg in cfg.get("columns", {}).items():
                if ccfg.get("include"):
                    r = f" (Role: {ccfg['role']})" if ccfg.get("role") else ""
                    detail_text += f" - {c}: {ccfg['type']}{r}\n"
            self.lbl_detail.config(text=detail_text)

    def _delete_sheet(self):
        sel = self.sheet_listbox.curselection()
        if not sel: return
        s = self.sheet_listbox.get(sel[0]).split(" (")[0]
        if messagebox.askyesno("Konfirmasi", f"Hapus konfigurasi untuk sheet '{s}'?\n\nSheet ini akan meminta setup ulang di generator."):
            del self.config_data["sheets"][s]
            self._save_changes()
            self._load_file()

    def _save_changes(self):
        p = self.config_path.get()
        save_config(self.config_data, p)
        messagebox.showinfo("Sukses", "Konfigurasi disimpan.")

if __name__ == "__main__":
    app = ConfigEditorApp()
    app.mainloop()
