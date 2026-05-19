import pandas as pd
import random
import os
from engine_lha import DASH_REQUIRED_COLS, PETA_REQUIRED_COLS, KPP_MAPPING

def generate():
    print("Generating Test Excel...")

    kpp_key = "KPP" if "KPP" in KPP_MAPPING[0] else "KODE_KPP" if "KODE_KPP" in KPP_MAPPING[0] else list(KPP_MAPPING[0].keys())[0]

    # 1. Gabungan (LHA)
    n_lha = 100
    gab_data = {}

    all_req_cols = list(DASH_REQUIRED_COLS.keys()) + PETA_REQUIRED_COLS

    for col in set(all_req_cols):
        if col in ["POTENSI", "REALISASI", "JML_TUNGGAKAN", "CAIR_TUNGGAKAN", "POTENSI_LHA", "POTENSI_AWAL", "POTENSI_AKHIR", "NILAI_PEMBAYARAN"]:
            gab_data[col] = [random.randint(1000000, 100000000) for _ in range(n_lha)]
        elif col in ["KODE_KPP", "KPP_ADM", "KPP_WP", "KPP"]:
            kpps = [k[kpp_key] for k in KPP_MAPPING]
            gab_data[col] = [random.choice(kpps) for _ in range(n_lha)]
        elif col in ["KODE_KANWIL", "KWL_ADM", "KWL_WP", "KWL"]:
            gab_data[col] = [random.choice(["010", "020", "030", "040", "050"]) for _ in range(n_lha)]
        elif col == "TAHUN_PAJAK":
            gab_data[col] = [random.choice([2020, 2021, 2022, 2023, 2024]) for _ in range(n_lha)]
        elif col in ["NAMA_WP", "NAMA_KPP", "KANWIL"]:
            gab_data[col] = [f"DUMMY DATA {i}" for i in range(n_lha)]
        elif col == "NPWP15":
            gab_data[col] = [f"012345678{i:06d}" for i in range(n_lha)]
        elif col in ["NAMA_KLU", "NM_KATEGORI", "KATEGORI"]:
            gab_data[col] = [random.choice(["Pertanian", "Pertambangan", "Industri Pengolahan"]) for _ in range(n_lha)]
        elif col in ["JENIS_WP", "NAMA_JENIS_WP"]:
            gab_data[col] = [random.choice(["Badan", "Orang Pribadi"]) for _ in range(n_lha)]
        elif col in ["PENERBIT", "Status SP2DK", "Status LHP2DK", "JENIS_LHA"]:
            gab_data[col] = [random.choice(["Kantor Pusat", "Kanwil", "KPP", "Terbit"]) for _ in range(n_lha)]
        elif col in ["LAT", "LONG"]:
            if col == "LAT": gab_data[col] = [random.uniform(-10.0, 5.0) for _ in range(n_lha)]
            else: gab_data[col] = [random.uniform(95.0, 140.0) for _ in range(n_lha)]
        else:
            gab_data[col] = ["-" for _ in range(n_lha)]

    df_gab = pd.DataFrame(gab_data)

    # 2. Perusahaan (Grup)
    n_perus = 20
    perus_data = {
        "ID_PERUSAHAAN": [f"P{i:03d}" for i in range(n_perus)],
        "NAMA_PERUSAHAAN": [f"PT GRUP MAKMUR {i}" for i in range(n_perus)],
        "NAMA_GRUP": [random.choice(["GRUP A", "GRUP B", "GRUP C"]) for _ in range(n_perus)],
        "LEVEL_PERUSAHAAN": [random.choice(["Holding", "Subsidiary 1", "Subsidiary 2"]) for _ in range(n_perus)],
        "NPWP": [f"012345678{i:06d}" for i in range(n_perus)],
        "STATUS_WP": [random.choice(["Aktif", "Non Efektif"]) for _ in range(n_perus)]
    }
    df_perus = pd.DataFrame(perus_data)

    # 3. LHA (Grup)
    n_lha_grup = 30
    lha_grup_data = {
        "ID_LHA": [f"L{i:03d}" for i in range(n_lha_grup)],
        "ID_PERUSAHAAN": [random.choice(perus_data["ID_PERUSAHAAN"]) for _ in range(n_lha_grup)],
        "POTENSI": [random.randint(5000000, 500000000) for _ in range(n_lha_grup)],
        "REALISASI": [random.randint(1000000, 400000000) for _ in range(n_lha_grup)],
        "STATUS_LHA": [random.choice(["Terbit LHA", "Belum LHA"]) for _ in range(n_lha_grup)],
        "TAHUN_LHA": [random.choice(["2022", "2023", "2024"]) for _ in range(n_lha_grup)]
    }
    df_lha_grup = pd.DataFrame(lha_grup_data)

    # 4. WP_PKH (Adaptive)
    n_pkh = 50
    kpps = [k[kpp_key] for k in KPP_MAPPING[:5]]
    pkh_data = {
        "NAMA_WP": [f"WP PKH {i}" for i in range(n_pkh)],
        "KWL": [random.choice(["010", "020", "030", "040", "050"]) for _ in range(n_pkh)],
        "KPP": [random.choice(kpps) for _ in range(n_pkh)],
        "SEKTOR": [random.choice(["Industri", "Perdagangan", "Jasa"]) for _ in range(n_pkh)],
        "POTENSI": [random.randint(100000, 10000000) for _ in range(n_pkh)],
        "LAT": [random.uniform(-10.0, 5.0) for _ in range(n_pkh)],
        "LONG": [random.uniform(95.0, 140.0) for _ in range(n_pkh)]
    }
    df_pkh = pd.DataFrame(pkh_data)

    out_path = "test_data.xlsx"
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df_gab.to_excel(writer, sheet_name="Gabungan", index=False)
        df_perus.to_excel(writer, sheet_name="Perusahaan", index=False)
        df_lha_grup.to_excel(writer, sheet_name="LHA", index=False)
        df_pkh.to_excel(writer, sheet_name="WP_PKH", index=False)

    print(f"Created {out_path}")

    import json
    cfg = {
        "sheets": {
            "WP_PKH": {
                "label": "Dashboard WP PKH",
                "columns": {
                    "NAMA_WP": {"type": "text", "include": True},
                    "KWL": {"type": "number", "include": True, "role": "kanwil_code"},
                    "KPP": {"type": "number", "include": True, "role": "kpp_code"},
                    "SEKTOR": {"type": "kategori", "include": True},
                    "POTENSI": {"type": "number", "include": True},
                    "LAT": {"type": "koordinat", "include": True, "role": "lat"},
                    "LONG": {"type": "koordinat", "include": True, "role": "long"}
                },
                "filters": ["SEKTOR"],
                "has_kanwil_kpp": True,
                "has_map": True
            }
        }
    }
    with open("dashboard_config.json", "w") as f:
        json.dump(cfg, f, indent=2)
    print("Configured WP_PKH in dashboard_config.json")

    print("Running headless test...")
    from engine_lha import build_lha_module, dash_process_data, peta_process_data
    from engine_grup import read_excel as grup_read_excel, build_data_model as grup_build_data_model, build_grup_module
    from engine_adaptive import process_adaptive_sheet
    from html_shell import build_super_html
    import datetime

    modules = []

    print("Testing Engine LHA")
    df_lha = pd.read_excel('test_data.xlsx', sheet_name="Gabungan", dtype=str)
    col_map = {k:k for k in DASH_REQUIRED_COLS}
    for c in PETA_REQUIRED_COLS:
        col_map[c] = c

    payload_dict, years_list = dash_process_data(df_lha, pd.DataFrame(KPP_MAPPING), col_map, lambda p, m: None)
    peta_b64, pub_vals, jenis_vals, klu_vals = peta_process_data(df_lha, col_map, print)
    lha_mod = build_lha_module(payload_dict, years_list, peta_b64, pub_vals, jenis_vals, klu_vals, "Dashboard LHA", print)
    modules.append({
        "id": "lha", "label": "Dashboard LHA", "icon": "📊",
        "css": lha_mod["css"], "js": lha_mod["js"],
        "sub_tabs": [
            {"id": "dash", "label": "Dashboard Monitoring", "html": lha_mod["body_dash"]},
            {"id": "peta", "label": "Peta Sebaran", "html": lha_mod["body_peta"]}
        ]
    })

    print("Testing Engine Grup")
    perusahaan_rows, lha_rows = grup_read_excel('test_data.xlsx')
    perusahaan = grup_build_data_model(perusahaan_rows, lha_rows)
    grup_mod = build_grup_module(perusahaan, 'test_data.xlsx', print)
    modules.append({
        "id": "grup", "label": "Dashboard Grup", "icon": "🏢",
        "css": grup_mod["css"], "js": grup_mod["js"], "body_extra": grup_mod.get("body_extra", ""),
        "sub_tabs": [
            {"id": "dash", "label": "Dashboard & Tabel", "html": grup_mod["body_dash"]},
            {"id": "net", "label": "Peta Jaringan", "html": grup_mod["body_net"]}
        ]
    })

    print("Testing Engine Adaptive")
    s = "WP_PKH"
    df_adp = pd.read_excel('test_data.xlsx', sheet_name=s)
    adp_mod = process_adaptive_sheet(df_adp, s, cfg["sheets"][s], KPP_MAPPING, print)
    modules.append({
        "id": f"adaptive_{s}", "label": cfg["sheets"][s].get("label", s), "icon": "📋",
        "css": adp_mod["css"], "js": adp_mod["js"], "sub_tabs": adp_mod["sub_tabs"]
    })

    print("Building Super HTML Shell")
    final_html = build_super_html(
        modules=modules, title="Test Dashboard",
        generated_at=datetime.datetime.now().strftime("%d %B %Y %H:%M"), encode_b64=True
    )

    with open('test_dashboard.html', "w", encoding="utf-8") as f:
        f.write(final_html)

    print(f"Generated test_dashboard.html with {len(final_html)} chars.")

if __name__ == "__main__":
    generate()
