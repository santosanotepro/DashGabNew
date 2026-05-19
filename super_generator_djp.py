import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import datetime
import traceback
import os

from engine_lha import build_lha_module, dash_process_data, peta_process_data, DASH_REQUIRED_COLS, KPP_MAPPING, build_agg
from engine_grup import read_excel as grup_read_excel, build_data_model as grup_build_data_model, build_grup_module
from engine_adaptive import process_adaptive_sheet
from html_shell import build_super_html
from config_manager import load_config, save_config, SetupDialog

class SuperGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LHA Super Generator Dashboard DJP v2")
        self.geometry("720x640")
        self.configure(bg="#f0f2f5")

        self.config_path = "dashboard_config.json"
        self.config_data = load_config(self.config_path)

        self.excel_path = tk.StringVar()
        self.title_var = tk.StringVar(value="Super Dashboard DJP")
        self.output_path = tk.StringVar()
        self.encode_b64_var = tk.BooleanVar(value=True)

        self.detected_sheets = []
        self.has_lha = False
        self.has_grup = False
        self.adaptive_sheets = {}
        self.adaptive_needs_setup = []

        self._build_ui()

    def _build_ui(self):
        hdr = tk.Frame(self, bg="#0a2346", height=60)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="██ LHA ██  Super Generator Dashboard DJP v2", bg="#0a2346", fg="#e8c96a", font=("Arial", 14, "bold")).pack(anchor=tk.W, padx=15, pady=(10, 0))
        tk.Label(hdr, text="Direktorat Jenderal Pajak | Multi-Dashboard", bg="#0a2346", fg="white", font=("Arial", 10)).pack(anchor=tk.W, padx=15, pady=(0, 10))

        body = tk.Frame(self, bg="#f0f2f5", padx=20, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        f1 = tk.Frame(body, bg="#f0f2f5")
        f1.pack(fill=tk.X, pady=5)
        tk.Label(f1, text="File Excel:", width=15, anchor=tk.W, bg="#f0f2f5", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        tk.Entry(f1, textvariable=self.excel_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(f1, text="Pilih", command=self._browse_excel).pack(side=tk.LEFT)

        tk.Label(body, text="Sheet yang terdeteksi:", bg="#f0f2f5", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(15, 5))
        self.sheet_listbox = tk.Listbox(body, height=6, bg="white", font=("Courier", 10))
        self.sheet_listbox.pack(fill=tk.X, pady=5)

        f2 = tk.Frame(body, bg="#f0f2f5")
        f2.pack(fill=tk.X, pady=10)
        tk.Label(f2, text="Judul Dashboard:", width=15, anchor=tk.W, bg="#f0f2f5", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        tk.Entry(f2, textvariable=self.title_var, width=50).pack(side=tk.LEFT, padx=5)

        f3 = tk.Frame(body, bg="#f0f2f5")
        f3.pack(fill=tk.X, pady=5)
        tk.Label(f3, text="Simpan ke:", width=15, anchor=tk.W, bg="#f0f2f5", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        tk.Entry(f3, textvariable=self.output_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(f3, text="Pilih", command=self._browse_output).pack(side=tk.LEFT)

        tk.Checkbutton(body, text="Encode Base64 (proteksi file)", variable=self.encode_b64_var, bg="#f0f2f5", font=("Arial", 10)).pack(anchor=tk.W, pady=5)

        tk.Label(body, text="Status:", bg="#f0f2f5", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(10, 5))
        self.log_text = tk.Text(body, height=6, bg="#1c2b3a", fg="white", font=("Courier", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(body, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)

        btn_frame = tk.Frame(body, bg="#f0f2f5")
        btn_frame.pack(fill=tk.X, pady=5)

        self.btn_setup = tk.Button(btn_frame, text="⚙️ Setup Sheet Baru", state=tk.DISABLED, command=self._setup_new_sheets)
        self.btn_setup.pack(side=tk.LEFT, padx=5)

        self.btn_generate = tk.Button(btn_frame, text="▶️ Generate Dashboard", bg="#1d4ed8", fg="white", font=("Arial", 10, "bold"), state=tk.DISABLED, command=self._generate)
        self.btn_generate.pack(side=tk.RIGHT, padx=5)

    def _log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def _update_progress(self, pct, msg=None):
        self.progress_var.set(pct)
        if msg: self._log(msg)

    def _browse_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if p:
            self.excel_path.set(p)
            out_p = p.rsplit(".", 1)[0] + "_dashboard.html"
            self.output_path.set(out_p)
            self._detect_sheets(p)

    def _browse_output(self):
        p = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML Files", "*.html")])
        if p: self.output_path.set(p)

    def _detect_sheets(self, file_path):
        self.sheet_listbox.delete(0, tk.END)
        self.detected_sheets = []
        self.has_lha = False
        self.has_grup = False
        self.adaptive_sheets = {}
        self.adaptive_needs_setup = []

        try:
            xl = pd.ExcelFile(file_path)
            sheet_names = xl.sheet_names

            if "Gabungan" in sheet_names:
                self.has_lha = True
                self.sheet_listbox.insert(tk.END, "✅ Gabungan         → Engine LHA")

            if "Perusahaan" in sheet_names and "LHA" in sheet_names:
                self.has_grup = True
                self.sheet_listbox.insert(tk.END, "✅ Perusahaan + LHA → Engine Grup")

            reserved = ["Gabungan", "Perusahaan", "LHA", "Sheet1"]
            if "sheets" not in self.config_data: self.config_data["sheets"] = {}

            for s in sheet_names:
                if s in reserved: continue
                if s in self.config_data["sheets"]:
                    self.adaptive_sheets[s] = self.config_data["sheets"][s]
                    self.sheet_listbox.insert(tk.END, f"⚙️ {s.ljust(16)} → Adaptif (config ada)")
                else:
                    self.adaptive_needs_setup.append(s)
                    self.sheet_listbox.insert(tk.END, f"🆕 {s.ljust(16)} → Adaptif (perlu setup)")

            if self.adaptive_needs_setup:
                self.btn_setup.config(state=tk.NORMAL)
                self.btn_generate.config(state=tk.DISABLED)
            else:
                self.btn_setup.config(state=tk.DISABLED)
                if self.has_lha or self.has_grup or self.adaptive_sheets:
                    self.btn_generate.config(state=tk.NORMAL)

            self._log(f"Berhasil mendeteksi sheet dari {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca Excel: {str(e)}")

    def _setup_new_sheets(self):
        if not self.adaptive_needs_setup: return
        path = self.excel_path.get()
        s = self.adaptive_needs_setup[0]
        try:
            df = pd.read_excel(path, sheet_name=s, nrows=50)
            dlg = SetupDialog(self, s, df, self.config_data["sheets"].get(s, {}))
            self.wait_window(dlg)

            if dlg.result:
                self.config_data["sheets"][s] = dlg.result
                save_config(self.config_data, self.config_path)
                self._detect_sheets(path)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal setup sheet {s}: {str(e)}")

    def _generate(self):
        self.btn_generate.config(state=tk.DISABLED)
        self.btn_setup.config(state=tk.DISABLED)
        threading.Thread(target=self._run_generation, daemon=True).start()

    def _run_generation(self):
        try:
            path = self.excel_path.get()
            out_path = self.output_path.get()
            modules = []

            if self.has_lha:
                self._update_progress(10, "Membaca data LHA Gabungan...")
                df_lha = pd.read_excel(path, sheet_name="Gabungan", dtype=str)
                col_map = {k:k for k in DASH_REQUIRED_COLS}
                payload_dict, years_list = dash_process_data(df_lha, pd.DataFrame(KPP_MAPPING), col_map, lambda p, m: self._update_progress(10 + (p*0.15), m))
                peta_b64, pub_vals, jenis_vals, klu_vals = peta_process_data(df_lha, col_map, self._log)
                lha_mod = build_lha_module(payload_dict, years_list, peta_b64, pub_vals, jenis_vals, klu_vals, "Dashboard LHA", self._log)
                modules.append({
                    "id": "lha", "label": "Dashboard LHA", "icon": "📊",
                    "css": lha_mod["css"], "js": lha_mod["js"],
                    "sub_tabs": [
                        {"id": "dash", "label": "Dashboard Monitoring", "html": lha_mod["body_dash"]},
                        {"id": "peta", "label": "Peta Sebaran", "html": lha_mod["body_peta"]}
                    ]
                })
                self._update_progress(35, "Engine LHA selesai.")

            if self.has_grup:
                self._update_progress(40, "Membaca data Grup...")
                perusahaan_rows, lha_rows = grup_read_excel(path)
                perusahaan = grup_build_data_model(perusahaan_rows, lha_rows)
                grup_mod = build_grup_module(perusahaan, os.path.basename(path), self._log)
                modules.append({
                    "id": "grup", "label": "Dashboard Grup", "icon": "🏢",
                    "css": grup_mod["css"], "js": grup_mod["js"], "body_extra": grup_mod.get("body_extra", ""),
                    "sub_tabs": [
                        {"id": "dash", "label": "Dashboard & Tabel", "html": grup_mod["body_dash"]},
                        {"id": "net", "label": "Peta Jaringan", "html": grup_mod["body_net"]}
                    ]
                })
                self._update_progress(55, "Engine Grup selesai.")

            if self.adaptive_sheets:
                prog_start = 55
                prog_step = 30 / len(self.adaptive_sheets)
                for i, (s, cfg) in enumerate(self.adaptive_sheets.items()):
                    self._update_progress(prog_start + (i*prog_step), f"Memproses sheet Adaptif: {s}...")
                    df_adp = pd.read_excel(path, sheet_name=s)
                    adp_mod = process_adaptive_sheet(df_adp, s, cfg, KPP_MAPPING, self._log)
                    modules.append({
                        "id": f"adaptive_{s}", "label": cfg.get("label", s), "icon": "📋",
                        "css": adp_mod["css"], "js": adp_mod["js"], "sub_tabs": adp_mod["sub_tabs"]
                    })

            self._update_progress(90, "Menyusun Master HTML...")
            final_html = build_super_html(
                modules=modules, title=self.title_var.get(),
                generated_at=datetime.datetime.now().strftime("%d %B %Y %H:%M"),
                encode_b64=self.encode_b64_var.get()
            )

            with open(out_path, "w", encoding="utf-8") as f: f.write(final_html)
            self._update_progress(100, f"Selesai! Tersimpan di: {out_path}")
            messagebox.showinfo("Sukses", "Dashboard berhasil di-generate!")

        except Exception as e:
            self._log(f"ERROR: {str(e)}\n{traceback.format_exc()}")
            messagebox.showerror("Error", "Terjadi kesalahan. Lihat log untuk detail.")
        finally:
            self.btn_generate.config(state=tk.NORMAL)
            if self.adaptive_needs_setup: self.btn_setup.config(state=tk.NORMAL)

if __name__ == "__main__":
    app = SuperGeneratorApp()
    app.mainloop()
