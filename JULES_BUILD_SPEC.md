# JULES_BUILD_SPEC.md — Super Generator Dashboard DJP v2

## OVERVIEW

Build a Python Tkinter application (`super_generator_djp.py`) that reads **one Excel file with multiple sheets** and produces **one self-contained HTML file** with a sidebar-navigated multi-dashboard layout.

The project reuses two existing generators as importable modules and adds a new Adaptive Engine for arbitrary sheets.

### Key Constraints
- **Python 3.11+**, only `pandas`, `openpyxl`, `tkinter` (built-in)
- **Offline-first**: no pip install during runtime, no internet required for output HTML
- Output HTML must work in Chrome/Edge opened from `file:///`
- All data embedded as Base64 JSON inside the HTML
- **CRITICAL**: Final HTML output must be **Base64-encoded** before writing to disk. The saved `.html` file contains a minimal bootstrap page that decodes and renders the real dashboard. This prevents casual users from opening the raw source.

---

## FILE STRUCTURE

```
project/
├── super_generator_djp.py          # NEW — Main orchestrator + Tkinter UI
├── engine_lha.py                   # NEW — Extracted from Generator_LHA_DJP_Gabungan_Final.py
├── engine_grup.py                  # NEW — Extracted from generator_grup_djp.py
├── engine_adaptive.py              # NEW — Adaptive engine for unknown sheets
├── config_manager.py               # NEW — JSON config read/write + Setup Dialog
├── html_shell.py                   # NEW — Master HTML template with sidebar
├── config_editor.py                # NEW — Standalone config editor (separate app)
├── _peta_assets.py                 # EXISTING — GeoJSON + KPP coords (DO NOT MODIFY)
├── Generator_LHA_DJP_Gabungan_Final.py  # EXISTING — Reference only
├── generator_grup_djp.py           # EXISTING — Reference only
└── dashboard_config.json           # AUTO-GENERATED at runtime
```

---

## FILE 1: `engine_lha.py`

### Purpose
Extract all data-processing and HTML-generation logic from `Generator_LHA_DJP_Gabungan_Final.py` into a clean importable module. **No Tkinter code in this file.**

### What to extract from `Generator_LHA_DJP_Gabungan_Final.py`:

**Constants** (copy verbatim):
- `DASH_REQUIRED_COLS` (dict, line 31-46)
- `KPP_MAPPING` (list of dicts, line 52-405) — all 295 KPP entries
- `PETA_REQUIRED_COLS` (list, line 412-416)
- `KPP_NAME_FIX` (dict, line 419-423)
- `KLU_ORDER` (list, line 426-452)
- `KLU_COLORS` (list, line 454-458)
- `JENIS_COLORS` (list, line 461-464)
- `PUB_COLORS` (dict, line 466)
- `PUB_LABELS` (dict, line 467)
- `_HTML_TEMPLATE` (raw string, line 472-801) — the full HTML template with `{{PLACEHOLDERS}}`

**Functions** (copy verbatim):
- `build_agg(data)` (line 805-844) — aggregation logic
- `dash_process_data(df, df_map, col_map, progress_cb=None)` (line 847-896) — returns `(payload_dict, years_list)`
- `build_html(payload, years)` (line 899-914) — legacy single-tab builder (keep for reference)
- `validate_columns(df, log)` (line 916-933) — peta column validation
- `peta_process_data(df, col_map, log)` (line 936-1015) — returns `(payload_b64, pub_vals, jenis_vals, klu_vals)`
- `_chip(item, color, chip_class, fn, label=None)` (line 1260-1267) — HTML chip helper
- `build_combined_html(dash_payload, dash_years, peta_b64, pub_vals, jenis_vals, klu_vals, title, log)` (line 1272-1317) — **THE MAIN BUILDER**, returns complete HTML string with both Dashboard + Peta tabs

**Also extract** (for the Super Generator to use for column mapping):
- `class MappingDialog(tk.Toplevel)` (line 1320-1358) — Tkinter dialog for column mapping. This one CAN have Tkinter since it's a dialog called by the main app.

### Public API

```python
# engine_lha.py public API:
from engine_lha import (
    DASH_REQUIRED_COLS, KPP_MAPPING, PETA_REQUIRED_COLS,
    build_agg, dash_process_data, validate_columns, peta_process_data,
    build_combined_html, MappingDialog,
    _ASSETS_LOADED  # bool: whether _peta_assets.py was found
)
```

### Import of `_peta_assets.py`

At the top of `engine_lha.py`, replicate the same try/except import pattern:
```python
_ASSETS_LOADED = False
try:
    _dir = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, _dir)
    from _peta_assets import GEO_B64, KREF_B64, KWP_B64, JS_B64
    _ASSETS_LOADED = True
except ImportError:
    GEO_B64 = KREF_B64 = KWP_B64 = JS_B64 = ""
```

---

## FILE 2: `engine_grup.py`

### Purpose
Extract all data-processing and HTML-generation logic from `generator_grup_djp.py` into a clean importable module. **No Tkinter code.**

### What to extract from `generator_grup_djp.py`:

**Functions** (copy verbatim):
- `read_excel(path)` (line 42-76) — reads sheets "Perusahaan" and "LHA", returns `(perusahaan_rows, lha_rows)`
- `build_data_model(perusahaan_rows, lha_rows)` (line 79-143) — returns dict of perusahaan with children/LHA attached
- `fmt_rp(val)` (line 150-158) — format Rupiah
- `generate_html(perusahaan, output_path, source_filename)` (line 161-1709) — **MODIFY THIS**: instead of writing to file, **return the HTML string**. Remove the `with open(output_path, ...) as f: f.write(html)` at the end and replace with `return html`.

### Public API

```python
# engine_grup.py public API:
from engine_grup import (
    read_excel as grup_read_excel,
    build_data_model as grup_build_data_model,
    generate_html as grup_generate_html,  # returns HTML string (modified)
)
```

### Critical Modification

In `generate_html()`, change the last 3 lines from:
```python
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
```
To:
```python
    return html
```

And change the function signature to NOT require `output_path`:
```python
def generate_html(perusahaan, source_filename):
```

---

## FILE 3: `engine_adaptive.py`

### Purpose
New engine that creates a dashboard from ANY Excel sheet based on user-configured column types.

### Input
```python
def process_adaptive_sheet(
    df: pd.DataFrame,
    sheet_name: str,
    config: dict,         # from dashboard_config.json for this sheet
    kpp_mapping: list,    # KPP_MAPPING from engine_lha
    log: callable,        # log function
    progress_cb: callable # progress callback(pct, msg)
) -> str:                 # returns HTML string for this module
```

### Config Structure (per sheet)
```json
{
  "label": "Dashboard WP PKH",
  "columns": {
    "NAMA_WP": {"type": "text", "include": true},
    "KWL": {"type": "number", "include": true, "role": "kanwil_code"},
    "KPP": {"type": "number", "include": true, "role": "kpp_code"},
    "SEKTOR": {"type": "kategori", "include": true},
    "POTENSI": {"type": "number", "include": true},
    "LAT": {"type": "koordinat", "include": true, "coord_role": "lat"},
    "LONG": {"type": "koordinat", "include": true, "coord_role": "long"}
  },
  "filters": ["SEKTOR"],
  "has_kanwil_kpp": true,
  "has_map": true
}
```

### Column Types
| Type | Behavior |
|------|----------|
| `text` | Shown in Data Explorer table, not aggregated |
| `number` | SUM in aggregations, shown in KPI cards and table |
| `kategori` | Becomes filter dropdown/chips, COUNT per unique value |
| `koordinat` | Pair of lat/long for map markers |

### HTML Output Structure

The function must return a complete `<div>` block (not a full HTML page) containing:

1. **Sub-tab bar** with up to 3 tabs:
   - "Tabel Kanwil/KPP" (only if `has_kanwil_kpp == true`)
   - "Peta" (only if `has_map == true`)
   - "Data Explorer" (always present)

2. **KPI Strip** — one card per `number` column showing SUM, formatted in Rupiah (Rp)

3. **Filter Bar** — one dropdown per `kategori` column

4. **Tabel Kanwil/KPP sub-tab** (if applicable):
   - Hierarchical table: Kanwil → KPP (expandable)
   - Use `KPP_MAPPING` for the Kanwil/KPP structure
   - Each `number` column becomes a data column in the table
   - SUM aggregation per Kanwil and per KPP
   - Matching the visual style of the LHA dashboard table (DJP dark theme header, gold accents)

5. **Peta sub-tab** (if applicable):
   - Use Leaflet.js (same CDN links as LHA engine)
   - Point mode: markers at coordinate columns
   - Bubble size proportional to the first `number` column
   - Popup with all included columns for that row
   - If `_peta_assets.py` is available, also support Area (choropleth) mode using province GeoJSON

6. **Data Explorer sub-tab** (always):
   - Full interactive table with all included columns
   - Search box (filters across all text columns)
   - Sort by clicking column headers
   - Filter dropdowns for `kategori` columns
   - Pagination: 100 rows per page with prev/next buttons
   - Number columns formatted as Rupiah with thousand separators

### Data Embedding
All data must be embedded as Base64 JSON inside `<script>` tags, same pattern as LHA engine:
```javascript
var _DATA = (function(){
    var s = "BASE64_CHUNKS_HERE";
    return JSON.parse(atob(s));
})();
```

### CSS
Use CSS variables matching the DJP theme:
```css
--djp-dark: #0a2346;
--djp-gold: #c8a84b;
--djp-gold2: #e8c96a;
--accent: #1d4ed8;
```

---

## FILE 4: `config_manager.py`

### Purpose
Handle JSON config file and provide a Tkinter Setup Dialog for new sheets.

### Functions

```python
def load_config(config_path: str) -> dict:
    """Load dashboard_config.json. Return empty structure if not found."""

def save_config(config: dict, config_path: str):
    """Write config to JSON file."""

def auto_detect_column_types(df: pd.DataFrame) -> dict:
    """
    Auto-suggest column types based on data analysis:
    - All numeric → 'number'
    - <20 unique values → 'kategori'
    - Column name matches LAT/LONG/LATITUDE/LONGITUDE → 'koordinat'
    - Otherwise → 'text'
    Returns: {col_name: {"type": str, "include": True, ...}}
    """
```

### Class: `SetupDialog(tk.Toplevel)`

A Tkinter dialog that appears when a new sheet is detected without config.

**UI Layout:**
```
┌─ Setup Kolom: [Sheet Name] ────────────────────┐
│                                                   │
│  Preview Data (5 baris pertama):                  │
│  ┌──────────────────────────────────────────┐    │
│  │ [scrollable table showing first 5 rows]   │    │
│  └──────────────────────────────────────────┘    │
│                                                   │
│  Konfigurasi Kolom:                               │
│  ┌──────┬────────┬──────────┬───────────┐        │
│  │ ✓    │ Kolom  │ Tipe     │ Role      │        │
│  ├──────┼────────┼──────────┼───────────┤        │
│  │ [x]  │ KWL    │ [number▾]│ [kwl_code]│        │
│  │ [x]  │ KPP    │ [number▾]│ [kpp_code]│        │
│  │ [x]  │ SEKTOR │ [kateg▾] │           │        │
│  │ [x]  │ POTENSI│ [number▾]│           │        │
│  │ [ ]  │ NOTES  │ [text▾]  │           │        │
│  └──────┴────────┴──────────┴───────────┘        │
│                                                   │
│  Label Dashboard: [________________]              │
│                                                   │
│  [Simpan & Lanjutkan]    [Lewati Sheet Ini]       │
└───────────────────────────────────────────────────┘
```

**Tipe dropdown options:** `text`, `number`, `kategori`, `koordinat`
**Role dropdown options (conditional):**
- If type == `number`: `(kosong)`, `kanwil_code`, `kpp_code`
- If type == `koordinat`: `lat`, `long`

**Output:** Updates config dict and saves to JSON.

### Config JSON Schema

```json
{
  "version": "2.0",
  "created": "ISO-8601 timestamp",
  "excel_filename": "original_filename.xlsx",
  "sheets": {
    "SheetName1": {
      "label": "Display Name",
      "columns": {
        "COL_NAME": {
          "type": "text|number|kategori|koordinat",
          "include": true,
          "role": "kanwil_code|kpp_code|lat|long|null"
        }
      },
      "filters": ["COL1", "COL2"],
      "has_kanwil_kpp": true,
      "has_map": false
    }
  }
}
```

The `has_kanwil_kpp` is automatically `true` when any column has `role: "kanwil_code"` AND another has `role: "kpp_code"`.

The `has_map` is automatically `true` when any column has `role: "lat"` AND another has `role: "long"`.

The `filters` list is automatically populated from columns with `type: "kategori"`.

---

## FILE 5: `html_shell.py`

### Purpose
Provide the master HTML template that wraps all dashboard modules into one page with sidebar navigation.

### Function

```python
def build_super_html(
    modules: list[dict],  # [{id, label, icon, html_content, sub_tabs}]
    title: str,
    generated_at: str,
    encode_b64: bool = True  # Base64 encode the entire output
) -> str:
```

### Module Dict Structure
```python
{
    "id": "lha",                    # unique identifier
    "label": "Dashboard LHA",       # sidebar label
    "icon": "📊",                   # emoji icon
    "sub_tabs": [                   # list of sub-tabs
        {"id": "monitoring", "label": "Dashboard Monitoring", "html": "..."},
        {"id": "peta", "label": "Peta Sebaran", "html": "..."},
    ]
}
```

### HTML Output Layout

```html
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{{TITLE}}</title>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>
    /* --- MASTER LAYOUT CSS --- */
    /* See CSS specification below */
  </style>
</head>
<body>
  <!-- MASTER HEADER -->
  <div id="master-header">...</div>

  <div id="app-body">
    <!-- SIDEBAR -->
    <nav id="sidebar">
      <!-- Menu items generated from modules list -->
      <!-- Each item shows icon + label -->
      <!-- Sub-items (sub-tabs) shown nested when parent active -->
    </nav>

    <!-- CONTENT AREA -->
    <main id="content">
      <!-- One panel per module, display:none except active -->
      <div class="module-panel" id="mod-lha">
        <!-- Sub-tab bar -->
        <div class="subtab-bar">
          <button class="subtab active">Dashboard Monitoring</button>
          <button class="subtab">Peta Sebaran</button>
        </div>
        <!-- Sub-tab content panels -->
        <div class="subtab-panel active">...LHA dashboard HTML...</div>
        <div class="subtab-panel">...LHA peta HTML...</div>
      </div>
      <div class="module-panel" id="mod-grup" style="display:none">
        ...
      </div>
    </main>
  </div>

  <script>
    // Sidebar navigation logic
    // Sub-tab switching logic
    // Lazy initialization for peta/heavy modules
  </script>
</body>
</html>
```

### CSS Specification

**Master Header:**
- Background: `#0a2346` (DJP dark)
- Bottom border: 3px solid `#c8a84b` (DJP gold)
- Height: ~56px
- Contains: Badge "DJP", title, "RAHASIA" badge, generated timestamp

**Sidebar:**
- Width: 240px, collapsible to 0px
- Background: `#0a2346`
- Position: fixed left, full height below header
- Menu items: icon + label, color `#c8a84b` (gold)
- Active item: left border 3px gold, lighter background `rgba(200,168,75,0.12)`
- Sub-items: indent 20px, font-size smaller, shown only when parent active
- Collapse button at bottom: toggle icon `◀`/`▶`
- Scrollable if items exceed viewport

**Content Area:**
- Takes remaining width (100% - sidebar)
- Each module panel: full height, hidden unless active
- Sub-tab bar: horizontal, inside content area, below header
  - Style: pill buttons, active = `#c8a84b` bg with dark text, inactive = transparent with gold text
  - Position: sticky top below master header

**Responsive:**
- Sidebar auto-collapses on width < 768px
- Sub-tab bar scrolls horizontally on overflow

### JavaScript Requirements

```javascript
const STATE = {
    activeModule: null,        // currently visible module id
    activeSubTabs: {},         // {moduleId: subTabId}
    initialized: {},           // {moduleId_subTabId: bool} for lazy init
    sidebarOpen: true
};

function switchModule(moduleId) {
    // Hide all module panels
    // Show selected module panel
    // Update sidebar active state
    // If module has peta sub-tab and it's active but not initialized, init it
}

function switchSubTab(moduleId, subTabId) {
    // Hide all sub-tab panels in this module
    // Show selected sub-tab panel
    // Update sub-tab bar active state
    // Lazy init if needed (especially for Leaflet maps)
}

function toggleSidebar() {
    // Toggle sidebar width 240px <-> 0px
    // Animate transition
    // Adjust content area width
}
```

### Base64 Encoding Wrapper

When `encode_b64=True`, the function wraps the entire HTML in a bootstrap page:

```python
def build_super_html(..., encode_b64=True):
    full_html = "..."  # the complete dashboard HTML
    
    if encode_b64:
        encoded = base64.b64encode(full_html.encode('utf-8')).decode('ascii')
        # Split into chunks for JS string concatenation
        chunks = [encoded[i:i+76] for i in range(0, len(encoded), 76)]
        data_js = '+\n'.join([f'"{c}"' for c in chunks])
        
        wrapper = f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>Dashboard DJP</title>
<style>
body{{margin:0;background:#0a2346;display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif}}
.loader{{color:#c8a84b;font-size:14px;text-align:center}}
.loader h2{{margin-bottom:8px}}
</style></head><body>
<div class="loader"><h2>🔓 Memuat Dashboard...</h2><p>Mohon tunggu sebentar</p></div>
<script>
(function(){{
var d={data_js};
var h=atob(d);
document.open();document.write(h);document.close();
}})();
</script></body></html>'''
        return wrapper
    
    return full_html
```

This means the saved `.html` file is a small bootstrap that decodes and renders the real dashboard. The raw HTML source is not directly readable.

---

## FILE 6: `super_generator_djp.py`

### Purpose
Main application. Tkinter UI that orchestrates everything.

### Tkinter UI Layout

```
┌──────────────────────────────────────────────────────┐
│  ██ LHA ██  Super Generator Dashboard DJP v2         │
│  Direktorat Jenderal Pajak | Multi-Dashboard         │
├══════════════════════════════════════════════════════╡
│                                                      │
│  File Excel:  [_________________________] [Pilih]    │
│                                                      │
│  Sheet yang terdeteksi:                              │
│  ┌──────────────────────────────────────────────┐   │
│  │ ✅ Gabungan         → Engine LHA             │   │
│  │ ✅ Perusahaan + LHA → Engine Grup            │   │
│  │ ⚙️ WP_PKH           → Adaptif (config ada)  │   │
│  │ 🆕 Data_Baru        → Adaptif (perlu setup)  │   │
│  └──────────────────────────────────────────────┘   │
│                                                      │
│  Judul Dashboard: [_________________________]        │
│  Simpan ke:       [_________________________] [Pilih]│
│                                                      │
│  ☐ Encode Base64 (proteksi file)                     │
│                                                      │
│  Status:                                             │
│  ┌──────────────────────────────────────────────┐   │
│  │ Membaca sheet Gabungan... 16,494 baris        │   │
│  │ Engine LHA: memproses...                      │   │
│  │ Engine LHA: dashboard selesai (8 tahun)       │   │
│  └──────────────────────────────────────────────┘   │
│                                                      │
│  [████████████████████░░░░░░░░] 65%                  │
│  Memproses Engine Grup...                            │
│                                                      │
│  [ ⚙️ Setup Sheet Baru ]  [ ▶️ Generate Dashboard ] │
│                                                      │
└──────────────────────────────────────────────────────┘
```

### Flow Logic

```python
class SuperGeneratorApp(tk.Tk):
    def __init__(self):
        # Window setup: 720x640, DJP dark theme header
        # UI elements as shown above
    
    def _browse_excel(self):
        # Open file dialog for .xlsx
        # Read all sheet names
        # Auto-detect each sheet type:
        #   1. Check if sheet name == "Gabungan" OR has columns matching DASH_REQUIRED_COLS → LHA
        #   2. Check if "Perusahaan" AND "LHA" sheets exist with ID_PERUSAHAAN/ID_LHA → Grup
        #   3. Everything else → check dashboard_config.json
        #      - If config exists for this sheet → "Adaptif (config ada)"
        #      - If no config → "Adaptif (perlu setup)"
        # Display detection results in the sheet list
    
    def _setup_new_sheets(self):
        # For each unconfig'd sheet, open SetupDialog
        # Save config after each dialog completes
    
    def _generate(self):
        # Thread-based generation:
        threading.Thread(target=self._run_generation, daemon=True).start()
    
    def _run_generation(self):
        modules = []
        
        # ── Step 1: LHA Engine (progress 0-35%) ──
        if self.has_lha:
            self._update_progress(0, "Memproses Engine LHA...")
            # Read the LHA sheet
            df_lha = pd.read_excel(path, sheet_name=lha_sheet, dtype=str)
            # Convert numeric columns
            # Run column mapping if needed
            # Call dash_process_data() → dashboard payload
            # Call peta_process_data() → peta payload
            # Call build_combined_html() → full LHA HTML
            # BUT: we need to extract the BODY content from this HTML,
            # not the full page. Split the HTML into:
            #   - CSS (extract from <style>...</style>)
            #   - Body content (extract from <body>...</body>)
            #   - JS (extract from <script>...</script>)
            # Package as module with sub_tabs:
            #   sub_tab 1: "Dashboard Monitoring" → dashboard part
            #   sub_tab 2: "Peta Sebaran" → peta part
            modules.append({
                "id": "lha",
                "label": "Dashboard LHA",
                "icon": "📊",
                "sub_tabs": [...]
            })
        
        # ── Step 2: Grup Engine (progress 35-55%) ──
        if self.has_grup:
            self._update_progress(35, "Memproses Engine Grup...")
            # Call grup_read_excel() → perusahaan_rows, lha_rows
            # Call grup_build_data_model() → perusahaan dict
            # Call grup_generate_html() → full Grup HTML string
            # Extract body content, CSS, JS
            # Package as module with sub_tabs:
            #   sub_tab 1: "Dashboard & Tabel"
            #   sub_tab 2: "Peta Jaringan"
            modules.append({
                "id": "grup",
                "label": "Dashboard Grup",
                "icon": "🏢",
                "sub_tabs": [...]
            })
        
        # ── Step 3: Adaptive Engines (progress 55-85%) ──
        for sheet_name, sheet_config in adaptive_sheets.items():
            self._update_progress(pct, f"Memproses {sheet_name}...")
            df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
            html_content = process_adaptive_sheet(df, sheet_name, sheet_config, ...)
            # Package as module
            sub_tabs = []
            if sheet_config.get("has_kanwil_kpp"):
                sub_tabs.append({"id": "tabel_kwl", "label": "Tabel Kanwil/KPP", ...})
            if sheet_config.get("has_map"):
                sub_tabs.append({"id": "peta", "label": "Peta", ...})
            sub_tabs.append({"id": "explorer", "label": "Data Explorer", ...})
            
            modules.append({
                "id": f"adaptive_{sheet_name}",
                "label": sheet_config["label"],
                "icon": "📋",
                "sub_tabs": sub_tabs
            })
        
        # ── Step 4: Assemble HTML (progress 85-100%) ──
        self._update_progress(85, "Menyusun HTML gabungan...")
        final_html = build_super_html(
            modules=modules,
            title=self.title_var.get(),
            generated_at=datetime.datetime.now().strftime("%d %B %Y %H:%M"),
            encode_b64=self.encode_b64_var.get()
        )
        
        # Write to file
        with open(self.output_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        self._update_progress(100, "Selesai!")
```

### CRITICAL: Extracting Sub-Modules from Existing HTML

Both existing engines produce **complete standalone HTML pages**. To embed them as modules inside the sidebar-navigated shell, we need to extract their inner content.

**Strategy: Use string splitting to extract CSS, body, and JS sections.**

```python
def extract_html_parts(full_html: str) -> dict:
    """
    Extract CSS, body content, and JS from a full HTML page.
    Returns: {"css": str, "body": str, "js": str}
    """
    import re
    
    # Extract all <style> blocks
    css_blocks = re.findall(r'<style[^>]*>(.*?)</style>', full_html, re.DOTALL)
    css = '\n'.join(css_blocks)
    
    # Extract body content (between <body> and </body>)
    body_match = re.search(r'<body[^>]*>(.*?)</body>', full_html, re.DOTALL)
    body = body_match.group(1) if body_match else ""
    
    # Extract all <script> blocks
    js_blocks = re.findall(r'<script[^>]*>(.*?)</script>', full_html, re.DOTALL)
    js = '\n'.join(js_blocks)
    
    # Remove script tags from body (we'll put JS separately)
    body = re.sub(r'<script[^>]*>.*?</script>', '', body, flags=re.DOTALL).strip()
    
    return {"css": css, "body": body, "js": js}
```

**For LHA engine**, the HTML has two panels:
- `<div class="tab-panel active" id="panel-dash">` → sub-tab "Dashboard Monitoring"
- `<div class="tab-panel" id="panel-peta">` → sub-tab "Peta Sebaran"

Extract each panel separately and wrap them as sub-tab contents.

**For Grup engine**, the HTML has two panels:
- `<div id="tab-dashboard" class="tab-pane active">` → sub-tab "Dashboard & Tabel"
- `<div id="tab-network" class="tab-pane">` → sub-tab "Peta Jaringan"

Plus the modal overlay (`<div class="modal-overlay">`), tooltip (`<div class="net-tooltip">`), which should be placed outside the sub-tab panels but inside the module panel.

### ID Namespacing

**CRITICAL**: Since multiple dashboards share one HTML page, all element IDs must be namespaced to avoid conflicts.

For LHA engine content, prefix all IDs with `lha_`:
- `#tbody` → `#lha_tbody`
- `#map` → `#lha_map`
- `#ysel` → `#lha_ysel`
- etc.

For Grup engine content, prefix with `grup_`:
- `#main-table` → `#grup_main-table`
- `#network-svg` → `#grup_network-svg`
- etc.

For Adaptive engine, prefix with `adp_{sheet_name}_`.

**This namespacing must also be applied to the JavaScript** — every `document.getElementById('...')` must reference the namespaced ID.

**Implementation approach**: After extracting HTML parts, do a find-and-replace on IDs:
```python
def namespace_html(html: str, js: str, prefix: str) -> tuple[str, str]:
    """Add prefix to all element IDs in HTML and corresponding JS getElementById calls."""
    # Find all id="..." in HTML
    ids = re.findall(r'id="([^"]+)"', html)
    for old_id in set(ids):
        new_id = f"{prefix}{old_id}"
        html = html.replace(f'id="{old_id}"', f'id="{new_id}"')
        html = html.replace(f'#{old_id}', f'#{new_id}')  # CSS selectors
        js = js.replace(f"'{old_id}'", f"'{new_id}'")
        js = js.replace(f'"{old_id}"', f'"{new_id}"')
        js = js.replace(f"getElementById('{old_id}')", f"getElementById('{new_id}')")
        js = js.replace(f'getElementById("{old_id}")', f'getElementById("{new_id}")')
    return html, js
```

---

## FILE 7: `config_editor.py`

### Purpose
Standalone Tkinter app to view/edit `dashboard_config.json`.

### Features
- Open JSON config file
- List all configured sheets
- Edit column types, include/exclude, roles
- Delete sheet config (forces re-setup on next generator run)
- Save changes
- Simple UI, no Excel file needed

### UI
```
┌── Config Editor Dashboard DJP ──────────────────────┐
│  File: [dashboard_config.json    ] [Buka]            │
│                                                      │
│  Sheet yang dikonfigurasi:                           │
│  ┌──────────────────────────────────────────────┐   │
│  │ > WP_PKH  (8 kolom, has_map=true)            │   │
│  │   Data_Baru (5 kolom, explorer only)          │   │
│  └──────────────────────────────────────────────┘   │
│                                                      │
│  Detail WP_PKH:                                      │
│  ┌──────┬────────┬──────────┬───────────┐           │
│  │ ✓    │ Kolom  │ Tipe     │ Role      │           │
│  │ ...  │ ...    │ ...      │ ...       │           │
│  └──────┴────────┴──────────┴───────────┘           │
│                                                      │
│  [Hapus Sheet]  [Simpan Perubahan]                   │
└──────────────────────────────────────────────────────┘
```

---

## TESTING REQUIREMENTS

### Test 1: LHA Only
- Input: Excel with sheet "Gabungan" containing LHA data
- Expected: HTML with sidebar showing "Dashboard LHA" with 2 sub-tabs (Monitoring + Peta)
- Verify: Dashboard table loads, peta renders (if _peta_assets.py present)

### Test 2: Grup Only
- Input: Excel with sheets "Perusahaan" and "LHA"
- Expected: HTML with sidebar showing "Dashboard Grup" with 2 sub-tabs (Dashboard + Network)
- Verify: KPI cards, table, network SVG

### Test 3: LHA + Grup
- Input: Excel with "Gabungan", "Perusahaan", "LHA" sheets
- Expected: HTML with 2 sidebar items, each with their sub-tabs
- Verify: Switching between LHA and Grup via sidebar works

### Test 4: Adaptive Sheet
- Input: Excel with an extra sheet (e.g., "WP_PKH") with numeric + kategori columns
- Expected: Setup dialog appears, after config → HTML includes 3rd sidebar item
- Verify: Data Explorer shows all rows, filters work

### Test 5: Base64 Encoding
- Input: Any valid Excel
- Check "Encode Base64" checkbox
- Expected: Output .html file contains bootstrap page with Base64 payload
- Verify: Opening in browser shows "Memuat Dashboard..." then renders correctly

### Test 6: Config Persistence
- Run generator once with adaptive sheet → config saved
- Run generator again with same Excel → no setup dialog, uses saved config
- Run config_editor, delete sheet config, run generator → setup dialog appears again

---

## IMPORTANT IMPLEMENTATION NOTES

1. **Do NOT modify `_peta_assets.py`** — it's a 750KB binary asset file. Import it as-is.

2. **Do NOT modify the existing generators** (`Generator_LHA_DJP_Gabungan_Final.py`, `generator_grup_djp.py`) — they are reference files. Extract their logic into `engine_lha.py` and `engine_grup.py`.

3. **All Tkinter UI should use the DJP theme:**
   - Header bg: `#0a2346`, text: `#e8c96a`
   - Body bg: `#f0f2f5`
   - Accent: `#1d4ed8`
   - Log area: dark bg `#1c2b3a`, light text

4. **The LHA engine's `build_combined_html()` returns a complete HTML page with Dashboard + Peta in one.** Instead of calling this function, you may need to call `dash_process_data()` and `peta_process_data()` separately, then build the sub-tab HTML yourself, embedding the CSS and JS appropriately.

5. **Similarly for Grup engine** — `generate_html()` returns a complete page. You need to extract its body and JS, namespace the IDs, and embed as a module.

6. **Error handling**: Every engine should be wrapped in try/except. If one engine fails, the others should still produce output. Show the error in the log but continue.

7. **Progress reporting**: Use a callback pattern `progress_cb(percentage: int, message: str)` for all engines.

8. **Character encoding**: Always use `utf-8` for reading Excel and writing HTML.

9. **Large data handling**: LHA data can be 16,000+ rows. Use pandas vectorized operations, not row-by-row loops. The existing engines already handle this efficiently — don't rewrite their aggregation logic.

10. **Leaflet.js**: Used by both LHA peta and adaptive peta. CDN link: `https://unpkg.com/leaflet@1.9.4/`. Only include once in the master HTML `<head>`, not per module.

---

## EXECUTION ORDER

```
1. Create engine_lha.py      (extract from Generator_LHA_DJP_Gabungan_Final.py)
2. Create engine_grup.py      (extract from generator_grup_djp.py)
3. Create config_manager.py   (config + setup dialog)
4. Create engine_adaptive.py  (new adaptive engine)
5. Create html_shell.py       (master HTML template)
6. Create super_generator_djp.py (main orchestrator)
7. Create config_editor.py    (standalone editor)
8. Test with sample data
```

---

## VALIDATION COMMANDS

After building, run these to verify:

```bash
# Check all files exist
ls -la engine_lha.py engine_grup.py engine_adaptive.py config_manager.py html_shell.py super_generator_djp.py config_editor.py

# Check imports work
python -c "from engine_lha import DASH_REQUIRED_COLS, KPP_MAPPING, build_agg, dash_process_data; print('LHA engine OK')"
python -c "from engine_grup import grup_read_excel, grup_build_data_model, grup_generate_html; print('Grup engine OK')" 
python -c "from engine_adaptive import process_adaptive_sheet; print('Adaptive engine OK')"
python -c "from config_manager import load_config, save_config, auto_detect_column_types; print('Config manager OK')"
python -c "from html_shell import build_super_html; print('HTML shell OK')"

# Check main app launches (will open Tkinter window)
python super_generator_djp.py

# Check config editor launches
python config_editor.py
```
