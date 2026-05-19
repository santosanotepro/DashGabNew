import json
import base64
import pandas as pd

def process_adaptive_sheet(df, sheet_name, config, kpp_mapping, log=None, progress_cb=None):
    if log: log(f"Memproses Adaptive Sheet: {sheet_name}")

    cols_config = config.get("columns", {})
    filters = config.get("filters", [])
    has_kanwil_kpp = config.get("has_kanwil_kpp", False)
    has_map = config.get("has_map", False)

    included_cols = [c for c, cfg in cols_config.items() if cfg.get("include")]
    if not included_cols: return {"css": "", "sub_tabs": [], "js": ""}

    df = df[included_cols].copy()

    for c in included_cols:
        ctype = cols_config[c].get("type")
        if ctype == "number":
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        elif ctype == "koordinat":
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        else:
            df[c] = df[c].fillna("").astype(str)

    payload_json = df.to_json(orient='records')
    payload_b64 = base64.b64encode(payload_json.encode('utf-8')).decode('ascii')

    prefix = f"adp_{sheet_name}_"

    css = f"""
    #{prefix}wrapper {{ padding: 10px; }}
    #{prefix}kpi-strip {{ display: flex; gap: 15px; margin-bottom: 20px; overflow-x: auto; }}
    .{prefix}kpi-card {{ background: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); flex: 1; min-width: 200px; border-top: 4px solid var(--djp-gold); }}
    .{prefix}kpi-val {{ font-size: 20px; font-weight: bold; color: var(--djp-dark); margin-top: 5px; }}
    #{prefix}filter-bar {{ display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; background: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
    .{prefix}filter-item {{ display: flex; flex-direction: column; }}
    .{prefix}filter-item label {{ font-size: 12px; font-weight: bold; color: #555; margin-bottom: 4px; }}
    .{prefix}filter-item select {{ padding: 6px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; min-width: 150px; }}
    #{prefix}data-table {{ width: 100%; border-collapse: collapse; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1); font-size: 13px; }}
    #{prefix}data-table th, #{prefix}data-table td {{ padding: 10px; border: 1px solid #eee; text-align: left; }}
    #{prefix}data-table th {{ background: var(--djp-dark); color: var(--djp-gold); cursor: pointer; position: sticky; top: 0; }}
    #{prefix}data-table tr:nth-child(even) {{ background: #f9f9f9; }}
    #{prefix}data-table tr:hover {{ background: #f0f2f5; }}
    .{prefix}num-col {{ text-align: right !important; }}
    #{prefix}map-container {{ height: 500px; width: 100%; border-radius: 8px; border: 1px solid #ccc; z-index: 1; }}
    #{prefix}search-box {{ padding: 8px; border: 1px solid #ccc; border-radius: 4px; width: 300px; margin-bottom: 10px; }}
    """

    number_cols = [c for c in included_cols if cols_config[c].get("type") == "number"]
    kategori_cols = [c for c in included_cols if cols_config[c].get("type") == "kategori"]

    filter_html = ""
    if kategori_cols:
        filter_html += f"<div id='{prefix}filter-bar'>"
        for col in kategori_cols:
            filter_html += f"""
            <div class='{prefix}filter-item'>
                <label>{col}</label>
                <select id='{prefix}filter_{col}' onchange='{prefix}applyFilters()'>
                    <option value='ALL'>Semua</option>
                </select>
            </div>
            """
        filter_html += "</div>"

    kpi_html = f"<div id='{prefix}kpi-strip'>"
    if number_cols:
        for col in number_cols:
            kpi_html += f"""
            <div class='{prefix}kpi-card'>
                <div style='font-size: 12px; color: #666;'>Total {col}</div>
                <div class='{prefix}kpi-val' id='{prefix}kpi_{col}'>0</div>
            </div>
            """
    else:
        kpi_html += f"""
        <div class='{prefix}kpi-card'>
            <div style='font-size: 12px; color: #666;'>Total Data</div>
            <div class='{prefix}kpi-val' id='{prefix}kpi_count'>0</div>
        </div>
        """
    kpi_html += "</div>"

    th_html = "".join([f"<th onclick='{prefix}sortBy(\"{c}\")'>{c}</th>" for c in included_cols])
    explorer_html = f"""
    <div id='{prefix}wrapper'>
        {filter_html}
        {kpi_html}
        <input type='text' id='{prefix}search-box' placeholder='Cari data...' onkeyup='{prefix}applyFilters()'>
        <div style='overflow-x: auto; max-height: 600px; border: 1px solid #ddd; border-radius: 4px;'>
            <table id='{prefix}data-table'>
                <thead><tr>{th_html}</tr></thead>
                <tbody id='{prefix}tbody'></tbody>
            </table>
        </div>
    </div>
    """

    js_config = {
        "prefix": prefix,
        "cols": included_cols,
        "numCols": number_cols,
        "catCols": kategori_cols,
        "hasMap": has_map,
        "hasKwl": has_kanwil_kpp
    }

    if has_map:
        for c, cfg in cols_config.items():
            if cfg.get("role") == "lat": js_config["latCol"] = c
            if cfg.get("role") == "long": js_config["lngCol"] = c

    if has_kanwil_kpp:
        for c, cfg in cols_config.items():
            if cfg.get("role") == "kanwil_code": js_config["kwlCol"] = c
            if cfg.get("role") == "kpp_code": js_config["kppCol"] = c

    js = f"""
    const {prefix}CFG = {json.dumps(js_config)};
    const {prefix}DATA = (function(){{ return JSON.parse(decodeURIComponent(escape(atob("{payload_b64}")))); }})();
    let {prefix}filtered = {prefix}DATA;
    let {prefix}map = null;
    let {prefix}markers = [];

    function {prefix}fmtRp(val) {{
        return 'Rp ' + Number(val).toLocaleString('id-ID');
    }}

    function {prefix}init() {{
        {prefix}CFG.catCols.forEach(col => {{
            let uniqueVals = [...new Set({prefix}DATA.map(d => d[col]))].filter(v => v);
            uniqueVals.sort();
            let sel = document.getElementById({prefix}CFG.prefix + 'filter_' + col);
            if(sel) {{
                uniqueVals.forEach(v => {{
                    let opt = document.createElement('option');
                    opt.value = v; opt.textContent = v;
                    sel.appendChild(opt);
                }});
            }}
        }});

        {prefix}applyFilters();
    }}

    function {prefix}applyFilters() {{
        let term = document.getElementById({prefix}CFG.prefix + 'search-box');
        term = term ? term.value.toLowerCase() : "";

        {prefix}filtered = {prefix}DATA.filter(d => {{
            if(term) {{
                let match = false;
                for(let k in d) {{
                    if(String(d[k]).toLowerCase().includes(term)) {{ match = true; break; }}
                }}
                if(!match) return false;
            }}

            for(let col of {prefix}CFG.catCols) {{
                let sel = document.getElementById({prefix}CFG.prefix + 'filter_' + col);
                if(sel && sel.value !== 'ALL' && d[col] !== sel.value) return false;
            }}
            return true;
        }});

        {prefix}renderKPI();
        {prefix}renderTable();
        if({prefix}CFG.hasMap && {prefix}map) {prefix}renderMap();
        if({prefix}CFG.hasKwl) {prefix}renderKwlTable();
    }}

    function {prefix}renderKPI() {{
        if({prefix}CFG.numCols.length > 0) {{
            let sums = {{}};
            {prefix}CFG.numCols.forEach(c => sums[c] = 0);

            {prefix}filtered.forEach(d => {{
                {prefix}CFG.numCols.forEach(c => sums[c] += Number(d[c]) || 0);
            }});

            {prefix}CFG.numCols.forEach(c => {{
                let el = document.getElementById({prefix}CFG.prefix + 'kpi_' + c);
                if(el) el.textContent = {prefix}fmtRp(sums[c]);
            }});
        }} else {{
            let el = document.getElementById({prefix}CFG.prefix + 'kpi_count');
            if(el) el.textContent = {prefix}filtered.length.toLocaleString('id-ID');
        }}
    }}

    function {prefix}renderTable() {{
        let tbody = document.getElementById({prefix}CFG.prefix + 'tbody');
        if(!tbody) return;

        let html = '';
        let maxRows = Math.min({prefix}filtered.length, 100);
        for(let i=0; i<maxRows; i++) {{
            let d = {prefix}filtered[i];
            html += '<tr>';
            {prefix}CFG.cols.forEach(c => {{
                let val = d[c];
                if({prefix}CFG.numCols.includes(c)) {{
                    html += '<td class="' + {prefix}CFG.prefix + 'num-col">' + {prefix}fmtRp(val) + '</td>';
                }} else {{
                    html += '<td>' + (val || '-') + '</td>';
                }}
            }});
            html += '</tr>';
        }}
        if({prefix}filtered.length > 100) {{
            html += '<tr><td colspan="' + {prefix}CFG.cols.length + '" style="text-align:center; color:#888;">Menampilkan 100 baris pertama dari ' + {prefix}filtered.length + ' data. Gunakan filter/pencarian untuk spesifik.</td></tr>';
        }}
        tbody.innerHTML = html;
    }}

    document.addEventListener('DOMContentLoaded', {prefix}init);
    """

    sub_tabs = []

    if has_kanwil_kpp:
        kwl_html = f"""
        <div id='{prefix}wrapper'>
            <h3>Tabel Agregasi Kanwil/KPP</h3>
            <div id='{prefix}kwl-container'></div>
        </div>
        """
        kwl_js = f"""
        function {prefix}renderKwlTable() {{
            let container = document.getElementById('{prefix}kwl-container');
            if(!container) return;

            let kwlCol = {prefix}CFG.kwlCol;
            let kppCol = {prefix}CFG.kppCol;
            let numCols = {prefix}CFG.numCols;

            let grouped = {{}};
            {prefix}filtered.forEach(d => {{
                let kwl = d[kwlCol] || 'Unknown';
                if(!grouped[kwl]) grouped[kwl] = {{ count: 0 }};
                grouped[kwl].count++;
                numCols.forEach(nc => {{
                    grouped[kwl][nc] = (grouped[kwl][nc] || 0) + (Number(d[nc]) || 0);
                }});
            }});

            let html = '<table id="{prefix}data-table"><thead><tr><th>Kanwil</th><th>Jml Data</th>';
            numCols.forEach(nc => html += '<th>' + nc + '</th>');
            html += '</tr></thead><tbody>';

            for(let k in grouped) {{
                html += '<tr><td>' + k + '</td><td>' + grouped[k].count + '</td>';
                numCols.forEach(nc => html += '<td class="{prefix}num-col">' + {prefix}fmtRp(grouped[k][nc]) + '</td>');
                html += '</tr>';
            }}
            html += '</tbody></table>';
            container.innerHTML = html;
        }}
        """
        js += "\n" + kwl_js
        sub_tabs.append({"id": "tabel_kwl", "label": "Tabel Kanwil/KPP", "html": kwl_html, "js": ""})

    if has_map:
        map_html = f"""
        <div id='{prefix}wrapper'>
            <div id='{prefix}map-container'></div>
        </div>
        """
        map_js = f"""
        function {prefix}initMap() {{
            if({prefix}map) return;
            let container = document.getElementById('{prefix}map-container');
            if(!container) return;

            {prefix}map = L.map('{prefix}map-container').setView([-2.5489, 118.0149], 5);
            L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
                attribution: '&copy; OpenStreetMap contributors'
            }}).addTo({prefix}map);

            {prefix}renderMap();
        }}

        function {prefix}renderMap() {{
            if(!{prefix}map) return;

            {prefix}markers.forEach(m => {prefix}map.removeLayer(m));
            {prefix}markers = [];

            let latCol = {prefix}CFG.latCol;
            let lngCol = {prefix}CFG.lngCol;

            {prefix}filtered.forEach(d => {{
                let lat = parseFloat(d[latCol]);
                let lng = parseFloat(d[lngCol]);
                if(!isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0) {{
                    let popupHtml = '<b>Detail Data</b><br>';
                    {prefix}CFG.cols.forEach(c => {{
                        if({prefix}CFG.numCols.includes(c)) popupHtml += c + ': ' + {prefix}fmtRp(d[c]) + '<br>';
                        else popupHtml += c + ': ' + (d[c]||'-') + '<br>';
                    }});

                    let m = L.circleMarker([lat, lng], {{
                        radius: 6,
                        fillColor: '#1d4ed8',
                        color: '#fff',
                        weight: 1,
                        opacity: 1,
                        fillOpacity: 0.8
                    }}).bindPopup(popupHtml);

                    m.addTo({prefix}map);
                    {prefix}markers.push(m);
                }}
            }});
        }}

        document.addEventListener('tabswitched', function(e) {{
            if(e.detail.tabId === 'peta' && e.detail.moduleId.startsWith('adaptive_{sheet_name}')) {{
                setTimeout({prefix}initMap, 200);
            }}
        }});
        """
        js += "\n" + map_js
        sub_tabs.append({"id": "peta", "label": "Peta", "html": map_html, "js": ""})

    sub_tabs.append({"id": "explorer", "label": "Data Explorer", "html": explorer_html, "js": ""})

    return {
        "css": css,
        "sub_tabs": sub_tabs,
        "js": js
    }
