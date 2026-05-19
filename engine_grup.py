import json, base64, math, pandas as pd, datetime, openpyxl

DASH_HTML = r'''<div id="grup_tab-dashboard" class="tab-pane active">
<div class="content">

<!-- KPI STRIP -->
<div class="kpi-strip" id="grup_kpi-strip">
  <div class="kpi-card dark">
    <div class="kpi-icon">🏢</div>
    <div class="kpi-val" id="grup_kpi-grup">{total_grup}</div>
    <div class="kpi-label">Total Grup</div>
    <div class="kpi-sub" id="grup_kpi-grup-sub">&nbsp;</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-icon">🏗</div>
    <div class="kpi-val" id="grup_kpi-perus">{total_perusahaan}</div>
    <div class="kpi-label">Perusahaan</div>
    <div class="kpi-sub" id="grup_kpi-perus-sub">&nbsp;</div>
  </div>
  <div class="kpi-card green">
    <div class="kpi-icon">✅</div>
    <div class="kpi-val" id="grup_kpi-terbit">{total_lha_terbit}</div>
    <div class="kpi-label">LHA Diterbitkan</div>
    <div class="kpi-sub" id="grup_kpi-terbit-sub">&nbsp;</div>
  </div>
  <div class="kpi-card red">
    <div class="kpi-icon">⏳</div>
    <div class="kpi-val" id="grup_kpi-belum">{total_lha_belum}</div>
    <div class="kpi-label">Belum Terbit</div>
    <div class="kpi-sub" id="grup_kpi-belum-sub">&nbsp;</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-icon">💰</div>
    <div class="kpi-val" id="grup_kpi-potensi">--</div>
    <div class="kpi-label">Total Potensi</div>
    <div class="kpi-sub" id="grup_kpi-potensi-sub">(Rp000)</div>
  </div>
  <div class="kpi-card amber">
    <div class="kpi-icon">💵</div>
    <div class="kpi-val" id="grup_kpi-realisasi">--</div>
    <div class="kpi-label">Realisasi</div>
    <div class="kpi-sub" id="grup_kpi-eff-sub">Efektivitas: --</div>
  </div>
</div>

<!-- FILTER BAR (full-width, semua dalam satu baris) -->
<div class="filter-bar-row">

  <div class="filter-item">
    <div class="filter-label">🏢 Grup Perusahaan</div>
    <select id="grup_sel-grup" onchange="filterByGrup()">
      <option value="ALL">— Semua Grup —</option>
    </select>
  </div>

  <div class="filter-item">
    <div class="filter-label">📋 Status LHA</div>
    <select id="grup_sel-status" onchange="filterByGrup()">
      <option value="ALL">Semua Status</option>
      <option value="DITERBITKAN">✅ Diterbitkan</option>
      <option value="BELUM DITERBITKAN">⏳ Belum Diterbitkan</option>
    </select>
  </div>

  <div class="filter-item">
    <div class="filter-label">📅 Tahun Pajak LHA</div>
    <select id="grup_sel-tahun" onchange="filterByGrup()">
      <option value="ALL">Semua Tahun</option>
    </select>
  </div>

  <div class="filter-item" style="flex:2; min-width:200px;">
    <div class="filter-label">🔍 Cari Perusahaan</div>
    <input type="text" id="grup_search-box" placeholder="Nama perusahaan..." oninput="filterByGrup()">
  </div>

</div>

<!-- DISTRIBUSI LEVEL (full width, horizontal) -->
<div class="level-dist-full" id="grup_level-dist-block">
  <div class="level-dist-title">📊 Distribusi Level Perusahaan</div>
  <div id="grup_level-dist-body" style="display:flex; gap:10px; flex-wrap:wrap; align-items:center;"><!-- diisi JS --></div>
</div>

<!-- TABLE PERUSAHAAN -->
<div class="table-card">
  <div class="table-card-header">
    <h3>📋 Daftar Perusahaan &amp; Status LHA</h3>
    <span id="grup_tbl-count">-- perusahaan</span>
  </div>
  <div class="table-scroll">
    <table id="grup_main-table">
      <thead>
        <tr>
          <th style="width:34px">#</th>
          <th style="min-width:240px">Nama Perusahaan</th>
          <th>Grup</th>
          <th>NPWP</th>
          <th>Sektor</th>
          <th>KPP</th>
          <th style="text-align:center">Kepemilikan</th>
          <th style="text-align:center">Status</th>
          <th style="text-align:center">LHA Terbit</th>
          <th style="text-align:center">LHA Belum</th>
          <th style="text-align:right">Total Potensi</th>
          <th style="text-align:right">Realisasi</th>
        </tr>
      </thead>
      <tbody id="grup_main-tbody">
        <tr><td colspan="12" class="empty-msg">Memuat data...</td></tr>
      </tbody>
    </table>
  </div>
</div>

</div><!-- /content -->
</div>'''
NET_HTML = r'''<div id="grup_tab-network" class="tab-pane">
<div class="network-layout">

  <!-- Sidebar kiri -->
  <div class="net-sidebar" id="grup_net-sidebar">
    <div class="net-sidebar-header">📋 Ringkasan Grup</div>
    <div class="net-sidebar-body" id="grup_net-sidebar-body"><!-- diisi JS --></div>
  </div>

  <!-- Canvas -->
  <div class="net-canvas-wrap">
    <div class="net-top-bar">
      <button class="net-toggle-sidebar" onclick="toggleSidebar()">
        <span id="grup_sidebar-icon">◀</span> <span id="grup_sidebar-label">Sembunyikan</span>
      </button>
      <div class="net-ctrl-group">
        <button class="net-btn" onclick="zoomNetwork(1.25)">＋ Zoom In</button>
        <button class="net-btn" onclick="zoomNetwork(0.8)">－ Zoom Out</button>
        <button class="net-btn" onclick="resetNetwork()">↺ Reset</button>
      </div>
    </div>

    <div class="net-filter-bar">
      <select id="grup_net-sel-grup" onchange="renderNetwork()" style="min-width:200px;">
        <option value="ALL">— Semua Grup —</option>
      </select>
    </div>

    <div class="net-legend" id="grup_net-legend">
      <div class="net-legend-title">Level Hierarki</div>
    </div>

    <svg id="grup_network-svg" xmlns="http://www.w3.org/2000/svg"></svg>
  </div>
</div>
</div>'''
EXTRA_HTML = r'''

<!-- MODAL DETAIL -->
<div class="modal-overlay" id="grup_modal-overlay" onclick="closeModal(event)">
  <div class="modal" id="grup_modal-box">
    <div class="modal-head">
      <div>
        <h2 id="grup_modal-nama"></h2>
        <div class="npwp" id="grup_modal-npwp"></div>
      </div>
      <button class="modal-close" onclick="closeModalBtn()">✕</button>
    </div>
    <div class="modal-body">
      <div class="info-grid" id="grup_modal-info-grid"></div>
      <div class="section-hdr">📄 Daftar LHA</div>
      <div id="grup_modal-lha-wrap"></div>
    </div>
  </div>
</div>

<div class="net-tooltip" id="grup_net-tooltip"></div>

<!-- ════════════════════════════════════════════
     JAVASCRIPT
════════════════════════════════════════════ -->'''
CSS_CODE = r'''@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');

:root {{
  --djp-900: #001D3D;
  --djp-800: #003566;
  --djp-700: #0052A5;
  --djp-600: #0070D8;
  --djp-400: #38A3F5;
  --djp-100: #DBEEFF;
  --djp-50:  #F0F7FF;
  --green:   #0A7C4F;
  --green-bg:#DCFCE7;
  --red:     #B91C1C;
  --red-bg:  #FEE2E2;
  --amber:   #D97706;
  --amber-bg:#FEF3C7;
  --slate:   #475569;
  --slate-light: #94A3B8;
  --border:  #CBD5E1;
  --bg:      #F0F7FF;
  --white:   #FFFFFF;
  --shadow-sm: 0 1px 3px rgba(0,0,0,.07), 0 1px 2px rgba(0,0,0,.05);
  --shadow:    0 4px 12px rgba(0,30,80,.10);
  --shadow-lg: 0 10px 30px rgba(0,30,80,.14);
  --radius:  12px;
  --radius-sm: 8px;
}}

* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
  font-family: 'Plus Jakarta Sans', sans-serif;
  background: var(--bg);
  color: var(--djp-900);
  font-size: 13px;
  line-height: 1.5;
}}
code, .mono {{ font-family: 'DM Mono', monospace; }}

/* ── TOPBAR ── */
.topbar {{
  background: linear-gradient(135deg, var(--djp-900) 0%, var(--djp-800) 60%, var(--djp-700) 100%);
  color: white;
  padding: 0 28px;
  height: 62px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  position: sticky;
  top: 0;
  z-index: 100;
  box-shadow: 0 2px 12px rgba(0,0,0,.25);
}}
.topbar-brand {{ display: flex; align-items: center; gap: 14px; }}
.topbar-icon {{
  width: 40px; height: 40px;
  background: rgba(255,255,255,.15);
  border-radius: 10px;
  display: flex; align-items: center; justify-content: center;
  font-size: 20px;
  border: 1px solid rgba(255,255,255,.2);
}}
.topbar-title {{ font-size: 16px; font-weight: 800; letter-spacing: -.2px; }}
.topbar-sub {{ font-size: 10px; opacity: .65; margin-top: 1px; font-weight: 500; letter-spacing: .3px; text-transform: uppercase; }}
.topbar-meta {{ font-size: 10px; opacity: .55; text-align: right; line-height: 1.6; }}

/* ── TABS ── */
.tabs {{
  background: var(--djp-800);
  display: flex;
  padding: 0 28px;
  border-bottom: 2px solid var(--djp-900);
}}
.tab-btn {{
  padding: 11px 20px;
  color: rgba(255,255,255,.55);
  border: none;
  background: transparent;
  cursor: pointer;
  font-family: inherit;
  font-size: 12px;
  font-weight: 700;
  letter-spacing: .3px;
  text-transform: uppercase;
  border-bottom: 3px solid transparent;
  transition: all .2s;
  margin-bottom: -2px;
}}
.tab-btn:hover {{ color: rgba(255,255,255,.85); }}
.tab-btn.active {{ color: white; border-bottom-color: #F5B800; }}
.tab-pane {{ display: none; }}
.tab-pane.active {{ display: block; }}

/* ── CONTENT AREA ── */
.content {{ padding: 22px 28px; }}

/* ── KPI STRIP ── */
.kpi-strip {{
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
  gap: 14px;
  margin-bottom: 20px;
}}
.kpi-card {{
  background: var(--white);
  border-radius: var(--radius);
  padding: 16px 18px;
  box-shadow: var(--shadow-sm);
  border: 1px solid var(--border);
  position: relative;
  overflow: hidden;
}}
.kpi-card::before {{
  content: '';
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 3px;
  background: var(--djp-600);
}}
.kpi-card.green::before {{ background: var(--green); }}
.kpi-card.amber::before {{ background: var(--amber); }}
.kpi-card.red::before   {{ background: var(--red); }}
.kpi-card.dark::before  {{ background: var(--djp-900); }}
.kpi-icon {{ font-size: 24px; margin-bottom: 6px; }}
.kpi-val  {{ font-size: 24px; font-weight: 800; line-height: 1; color: var(--djp-900); }}
.kpi-label{{ font-size: 10px; color: var(--slate); font-weight: 700; text-transform: uppercase; letter-spacing: .5px; margin-top: 4px; }}
.kpi-sub  {{ font-size: 10px; color: var(--slate-light); margin-top: 2px; }}

/* ── FILTER BAR (full-width, satu baris) ── */
.filter-bar-row {{
  background: var(--white);
  border-radius: var(--radius);
  padding: 14px 18px;
  margin-bottom: 12px;
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
  align-items: flex-end;
  box-shadow: var(--shadow-sm);
  border: 1px solid var(--border);
}}
.filter-item {{
  display: flex;
  flex-direction: column;
  gap: 5px;
  flex: 1;
  min-width: 160px;
}}
.filter-label {{
  font-size: 10px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .5px; color: var(--slate);
}}
.filter-item select,
.filter-item input {{
  width: 100%;
  border: 1.5px solid var(--border);
  border-radius: var(--radius-sm);
  padding: 7px 10px;
  font-family: inherit;
  font-size: 12px;
  color: var(--djp-900);
  background: var(--bg);
  cursor: pointer;
  outline: none;
  transition: border-color .15s;
}}
.filter-item select:focus,
.filter-item input:focus {{
  border-color: var(--djp-600);
  background: white;
}}

/* ── DISTRIBUSI LEVEL (full-width, horizontal chips) ── */
.level-dist-full {{
  background: var(--white);
  border-radius: var(--radius);
  padding: 14px 18px;
  margin-bottom: 16px;
  box-shadow: var(--shadow-sm);
  border: 1px solid var(--border);
}}
.level-dist-title {{
  font-size: 11px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .5px; color: var(--djp-800); margin-bottom: 10px;
  display: flex; align-items: center; gap: 6px;
}}
.lvl-row {{
  display: flex;
  align-items: center;
  gap: 8px;
  flex: 0 0 auto;
  background: var(--bg);
  border-radius: 8px;
  padding: 7px 12px;
  border: 1px solid var(--border);
}}
.lvl-chip {{
  padding: 3px 10px;
  border-radius: 20px;
  font-size: 10px;
  font-weight: 700;
  text-align: center;
  white-space: nowrap;
  min-width: 52px;
}}
.lvl-track {{
  background: var(--djp-50);
  border-radius: 4px;
  height: 8px;
  width: 60px;
}}
.lvl-fill {{ height: 8px; border-radius: 4px; transition: width .4s ease; }}
.lvl-count {{
  font-size: 13px; font-weight: 800;
  color: var(--djp-800); min-width: 16px; text-align: right;
}}

/* ── TABLE CARD ── */
.table-card {{
  background: var(--white);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  border: 1px solid var(--border);
  overflow: hidden;
  margin-bottom: 20px;
}}
.table-card-header {{
  background: var(--djp-800);
  color: white;
  padding: 13px 18px;
  display: flex;
  justify-content: space-between;
  align-items: center;
}}
.table-card-header h3 {{ font-size: 13px; font-weight: 700; }}
.table-card-header span {{ font-size: 11px; opacity: .7; }}
.table-scroll {{ overflow-x: auto; }}
table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
thead th {{
  background: #1B3A6B;
  color: white;
  padding: 10px 10px;
  text-align: left;
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: .4px;
  white-space: nowrap;
  border-right: 1px solid rgba(255,255,255,.08);
  position: sticky;
  top: 0;
  z-index: 2;
}}
thead th:last-child {{ border-right: none; }}
tbody tr {{ cursor: pointer; transition: background .1s; }}
tbody tr:hover td {{ background: #EFF6FF !important; }}
tbody td {{
  padding: 9px 10px;
  border-bottom: 1px solid #F1F5F9;
  vertical-align: middle;
}}
tbody tr:last-child td {{ border-bottom: none; }}

/* ── BADGES ── */
.badge {{
  display: inline-flex; align-items: center; gap: 4px;
  padding: 3px 9px;
  border-radius: 20px;
  font-size: 10px;
  font-weight: 700;
  white-space: nowrap;
}}
.badge-green {{ background: var(--green-bg); color: var(--green); }}
.badge-red   {{ background: var(--red-bg);   color: var(--red);   }}
.badge-amber {{ background: var(--amber-bg); color: var(--amber); }}
.badge-blue  {{ background: var(--djp-100);  color: var(--djp-700); }}
.badge-gray  {{ background: #F1F5F9; color: var(--slate); }}

/* ── LEVEL CHIP IN TABLE ── */
.level-tag {{
  display: inline-block;
  padding: 2px 8px;
  border-radius: 10px;
  font-size: 10px;
  font-weight: 700;
}}

/* ── NAMA INDENTASI ── */
.nama-cell {{ display: flex; align-items: center; gap: 0; }}
.indent-line {{
  display: inline-block;
  width: 16px;
  flex-shrink: 0;
}}
.tree-icon {{ font-size: 12px; opacity: .5; margin-right: 3px; flex-shrink: 0; }}
.expand-btn {{
  background: rgba(0,82,165,.12);
  border: none;
  color: var(--djp-700);
  cursor: pointer;
  border-radius: 4px;
  width: 18px; height: 18px;
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 10px;
  font-weight: 700;
  margin-right: 4px;
  flex-shrink: 0;
  transition: all .15s;
}}
.expand-btn:hover {{ background: var(--djp-600); color: white; }}

/* ── PROGRESS BAR ── */
.pct-wrap {{ display: flex; align-items: center; gap: 5px; }}
.pct-bar  {{ flex: 1; background: #EFF6FF; border-radius: 3px; height: 6px; min-width: 50px; }}
.pct-fill {{ height: 6px; border-radius: 3px; background: var(--djp-600); }}
.pct-text {{ font-size: 10px; font-weight: 600; color: var(--slate); white-space: nowrap; }}

/* ── HIDDEN ROW ── */
.child-row {{ }}
.child-row.collapsed {{ display: none !important; }}

/* ── MODAL ── */
.modal-overlay {{
  display: none;
  position: fixed; inset: 0;
  background: rgba(0,10,30,.55);
  backdrop-filter: blur(3px);
  z-index: 1000;
  align-items: center;
  justify-content: center;
}}
.modal-overlay.open {{ display: flex; }}
.modal {{
  background: white;
  border-radius: 16px;
  width: 760px; max-width: 96vw; max-height: 88vh;
  overflow-y: auto;
  box-shadow: var(--shadow-lg);
  animation: modalIn .2s ease;
}}
@keyframes modalIn {{
  from {{ opacity: 0; transform: translateY(12px) scale(.98); }}
  to   {{ opacity: 1; transform: translateY(0)   scale(1);    }}
}}
.modal-head {{
  padding: 18px 22px;
  background: var(--djp-800);
  color: white;
  border-radius: 16px 16px 0 0;
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
}}
.modal-head h2 {{ font-size: 15px; font-weight: 800; line-height: 1.3; }}
.modal-head .npwp {{ font-size: 11px; opacity: .7; margin-top: 3px; font-family: 'DM Mono', monospace; }}
.modal-close {{
  background: rgba(255,255,255,.15);
  border: none; color: white; cursor: pointer;
  border-radius: 8px; width: 30px; height: 30px;
  font-size: 16px; display: flex; align-items: center; justify-content: center;
  transition: background .15s;
}}
.modal-close:hover {{ background: rgba(255,255,255,.28); }}
.modal-body {{ padding: 22px; }}
.info-grid {{
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
  margin-bottom: 18px;
}}
.info-item {{
  background: var(--bg);
  border-radius: var(--radius-sm);
  padding: 10px 12px;
  border: 1px solid var(--border);
}}
.info-item label {{ font-size: 9px; color: var(--slate-light); font-weight: 700; text-transform: uppercase; letter-spacing: .4px; display: block; margin-bottom: 3px; }}
.info-item span  {{ font-size: 13px; font-weight: 700; color: var(--djp-900); }}
.section-hdr {{
  font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .5px;
  color: var(--djp-800); margin: 16px 0 10px;
  display: flex; align-items: center; gap: 8px;
}}
.section-hdr::after {{
  content: '';
  flex: 1; height: 2px;
  background: linear-gradient(to right, var(--djp-100), transparent);
  border-radius: 2px;
}}
.modal-lha-table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
.modal-lha-table th {{ background: #1B3A6B; color: white; padding: 8px 10px; text-align: center; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .3px; }}
.modal-lha-table td {{ padding: 8px 10px; border-bottom: 1px solid #F1F5F9; }}
.modal-lha-table tr.terbit td {{ background: #F0FFF8; }}
.modal-lha-table tr.belum  td {{ background: #FFF8F8; }}
.modal-lha-table tr:last-child td {{ border-bottom: none; }}
.no-lha {{ text-align: center; padding: 22px; color: var(--slate-light); font-size: 12px; }}

/* ── NETWORK TAB ── */
#grup_tab-network {{ position: relative; }}
.network-layout {{
  display: flex;
  height: calc(100vh - 160px);
  min-height: 500px;
  gap: 0;
}}

/* Sidebar */
.net-sidebar {{
  width: 280px;
  background: var(--white);
  border-right: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
  transition: width .25s ease, opacity .25s ease;
  overflow: hidden;
}}
.net-sidebar.hidden {{
  width: 0;
  opacity: 0;
  pointer-events: none;
}}
.net-sidebar-header {{
  background: var(--djp-800);
  color: white;
  padding: 14px 16px;
  font-size: 11px;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: .5px;
  flex-shrink: 0;
}}
.net-sidebar-body {{
  padding: 14px;
  overflow-y: auto;
  flex: 1;
}}
.net-grup-card {{
  border: 1.5px solid var(--border);
  border-radius: var(--radius-sm);
  margin-bottom: 10px;
  overflow: hidden;
  transition: border-color .15s;
  cursor: pointer;
}}
.net-grup-card:hover {{ border-color: var(--djp-400); }}
.net-grup-card.selected {{ border-color: var(--djp-600); }}
.net-grup-card-head {{
  padding: 9px 12px;
  display: flex; align-items: center; gap: 8px;
  font-size: 11px; font-weight: 700;
}}
.net-grup-dot {{
  width: 10px; height: 10px;
  border-radius: 50%;
  flex-shrink: 0;
}}
.net-grup-stat {{
  padding: 8px 12px 10px;
  background: var(--bg);
  font-size: 10px;
  color: var(--slate);
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 5px;
}}
.net-stat-item span {{ display: block; font-weight: 700; font-size: 12px; color: var(--djp-900); }}

/* Network canvas area */
.net-canvas-wrap {{
  flex: 1;
  position: relative;
  background: linear-gradient(145deg, #EBF4FF 0%, #F5F9FF 50%, #EDF7F2 100%);
  overflow: hidden;
}}
#grup_network-svg {{ width: 100%; height: 100%; }}
.net-top-bar {{
  position: absolute; top: 12px; left: 12px; right: 12px;
  display: flex; align-items: center; gap: 8px;
  pointer-events: none;
  z-index: 10;
}}
.net-toggle-sidebar {{
  pointer-events: all;
  background: white;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 7px 12px;
  font-family: inherit;
  font-size: 11px; font-weight: 700;
  cursor: pointer;
  box-shadow: var(--shadow-sm);
  color: var(--djp-800);
  display: flex; align-items: center; gap: 5px;
  transition: all .15s;
}}
.net-toggle-sidebar:hover {{ background: var(--djp-600); color: white; border-color: var(--djp-600); }}
.net-ctrl-group {{
  pointer-events: all;
  display: flex; gap: 6px;
  margin-left: auto;
}}
.net-btn {{
  background: rgba(255,255,255,.92);
  border: 1px solid var(--border);
  border-radius: 7px;
  padding: 6px 12px;
  font-family: inherit;
  font-size: 11px; font-weight: 700;
  color: var(--djp-800);
  cursor: pointer;
  box-shadow: var(--shadow-sm);
  transition: all .15s;
}}
.net-btn:hover {{ background: var(--djp-700); color: white; border-color: var(--djp-700); }}
.net-filter-bar {{
  position: absolute; bottom: 12px; left: 12px;
  display: flex; align-items: center; gap: 8px;
  z-index: 10;
}}
.net-filter-bar select {{
  background: rgba(255,255,255,.92);
  border: 1px solid var(--border);
  border-radius: 7px;
  padding: 6px 10px;
  font-family: inherit;
  font-size: 11px;
  color: var(--djp-900);
  cursor: pointer;
  box-shadow: var(--shadow-sm);
  outline: none;
}}
.net-legend {{
  position: absolute; bottom: 12px; right: 12px;
  background: rgba(255,255,255,.92);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 10px 12px;
  font-size: 10px;
  z-index: 10;
  box-shadow: var(--shadow-sm);
}}
.net-legend-title {{ font-weight: 800; color: var(--djp-800); margin-bottom: 6px; text-transform: uppercase; letter-spacing: .4px; }}
.net-legend-item {{ display: flex; align-items: center; gap: 6px; margin-bottom: 3px; }}
.net-legend-dot {{ width: 10px; height: 10px; border-radius: 50%; }}

/* ── TOOLTIP ── */
.net-tooltip {{
  position: fixed;
  background: rgba(0,15,40,.88);
  color: white;
  padding: 10px 14px;
  border-radius: 10px;
  font-size: 11px;
  pointer-events: none;
  z-index: 999;
  max-width: 260px;
  line-height: 1.6;
  display: none;
  box-shadow: var(--shadow-lg);
}}
.net-tooltip strong {{ font-size: 12px; font-weight: 700; display: block; margin-bottom: 3px; }}

/* ── MISC ── */
.empty-msg {{ text-align: center; padding: 28px; color: var(--slate-light); font-size: 12px; }}
.hidden {{ display: none !important; }}'''
JS_TEMPLATE = r'''const _GRUP_DATA = (function(){ return JSON.parse(decodeURIComponent(escape(atob('{b64_data}')))); })();
// ── DATA ──────────────────────────────────────────────────────────────────
const PERUSAHAAN = {p_json};

// ── CONSTANTS ─────────────────────────────────────────────────────────────
const LEVEL_COLORS = [
  {{ bg:'#1B3A6B', text:'#FFFFFF', name:'Induk / Level 0' }},
  {{ bg:'#0070D8', text:'#FFFFFF', name:'Anak / Level 1'  }},
  {{ bg:'#0EA5E9', text:'#FFFFFF', name:'Cucu / Level 2'  }},
  {{ bg:'#38BDF8', text:'#0C4A6E', name:'Cicit / Level 3' }},
  {{ bg:'#7DD3FC', text:'#0C4A6E', name:'Level 4'         }},
  {{ bg:'#BAE6FD', text:'#0C4A6E', name:'Level 5'         }},
  {{ bg:'#E0F2FE', text:'#0369A1', name:'Level 6'         }},
  {{ bg:'#F0F9FF', text:'#0369A1', name:'Level 7'         }},
];
const GRUP_COLORS = [
  '#0052A5','#0A7C4F','#B91C1C','#D97706','#7C3AED',
  '#0891B2','#BE185D','#15803D','#C2410C','#1D4ED8',
];

// ── SETUP ─────────────────────────────────────────────────────────────────
const grupList = [...new Set(Object.values(PERUSAHAAN).map(p => p.nama_grup).filter(Boolean))].sort();
let grupColorMap = {{}};
grupList.forEach((g,i) => {{ grupColorMap[g] = GRUP_COLORS[i % GRUP_COLORS.length]; }});

// Isi dropdown grup
['grup_sel-grup','grup_net-sel-grup'].forEach(id => {{
  const sel = document.getElementById(id);
  grupList.forEach(g => {{
    const op = document.createElement('option');
    op.value = g; op.textContent = g;
    sel.appendChild(op);
  }});
}});

// Isi dropdown tahun
const allTahun = [...new Set(
  Object.values(PERUSAHAAN).flatMap(p => p.lha_list.map(l => l.tahun_pajak)).filter(Boolean)
)].sort((a,b) => a-b);
const selTahun = document.getElementById('grup_sel-tahun');
allTahun.forEach(t => {{
  const op = document.createElement('option');
  op.value = t; op.textContent = t;
  selTahun.appendChild(op);
}});

// ── TABS ──────────────────────────────────────────────────────────────────
function switchTab(id, btn) {{
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('tab-' + id).classList.add('active');
  if (id === 'network') renderNetwork();
}}

// ── FORMAT HELPERS ────────────────────────────────────────────────────────
function fmtRp(v) {{
  if (!v || v === 0) return '—';
  if (v >= 1000000) return 'Rp ' + (v/1000000).toFixed(1) + ' M';
  if (v >= 1000)    return 'Rp ' + (v/1000).toFixed(1) + ' rb';
  return 'Rp ' + v.toFixed(0);
}}
function fmtPct(v) {{
  if (!v || v === 0) return '—';
  return v.toFixed(0) + '%';
}}

// ── MAIN FILTER & RENDER TABLE ────────────────────────────────────────────

// Kembalikan semua ID perusahaan sesuai filter GRUP + SEARCH saja
// (status LHA dan tahun pajak TIDAK menyaring perusahaan)
function getGrupIds() {{
  const selGrup = document.getElementById('grup_sel-grup').value;
  const search  = document.getElementById('grup_search-box').value.toLowerCase().trim();

  return Object.values(PERUSAHAAN).filter(p => {{
    if (selGrup !== 'ALL' && p.nama_grup !== selGrup) return false;
    if (search && !p.nama.toLowerCase().includes(search) && !p.nama_grup.toLowerCase().includes(search)) return false;
    return true;
  }}).map(p => p.id);
}}

// Helper: hitung LHA per perusahaan sesuai filter status + tahun
function calcLha(p, selStatus, selTahun2) {{
  let terbit = 0, belum = 0, pot = 0, real = 0;
  p.lha_list.forEach(l => {{
    const okS = selStatus === 'ALL' || l.status === selStatus;
    const okT = selTahun2 === 'ALL' || l.tahun_pajak == selTahun2;
    if (okS && okT) {{
      if (l.status === 'DITERBITKAN') terbit++; else belum++;
      pot  += l.potensi_total;
      real += l.realisasi;
    }}
  }});
  return {{ terbit, belum, pot, real }};
}}

function filterByGrup() {{
  const ids = new Set(getGrupIds());
  const selStatus = document.getElementById('grup_sel-status').value;
  const selTahun2 = document.getElementById('grup_sel-tahun').value;
  updateKPI(ids, selStatus, selTahun2);
  updateLevelDist(ids, selStatus, selTahun2);
  renderTable(ids, selStatus, selTahun2);
}}

function updateKPI(ids, selStatus, selTahun2) {{
  const selGrup = document.getElementById('grup_sel-grup').value;
  let terbit = 0, belum = 0, pot = 0, real = 0;
  ids.forEach(id => {{
    const r = calcLha(PERUSAHAAN[id], selStatus, selTahun2);
    terbit += r.terbit; belum += r.belum; pot += r.pot; real += r.real;
  }});
  const eff = pot > 0 ? (real / pot * 100).toFixed(1) : '0';

  document.getElementById('grup_kpi-grup').textContent      = selGrup === 'ALL' ? grupList.length : 1;
  document.getElementById('grup_kpi-perus').textContent     = ids.size;
  document.getElementById('grup_kpi-terbit').textContent    = terbit;
  document.getElementById('grup_kpi-belum').textContent     = belum;
  document.getElementById('grup_kpi-potensi').textContent   = fmtRp(pot);
  document.getElementById('grup_kpi-realisasi').textContent = fmtRp(real);
  document.getElementById('grup_kpi-eff-sub').textContent   = 'Efektivitas: ' + eff + '%';
}}

function updateLevelDist(ids, selStatus, selTahun2) {{
  // total per level (semua perusahaan dalam grup)
  const totals = {{}};
  // match per level (perusahaan yang punya LHA sesuai filter)
  const matches = {{}};
  for (let i = 0; i <= 7; i++) {{ totals[i] = 0; matches[i] = 0; }}

  ids.forEach(id => {{
    const p  = PERUSAHAAN[id];
    const lv = p.level;
    totals[lv]++;
    // Perusahaan "match" jika filter ALL/ALL -> semua match
    // atau jika ada setidaknya 1 LHA yang cocok filter
    const isAll = selStatus === 'ALL' && selTahun2 === 'ALL';
    if (isAll) {{
      matches[lv]++;
    }} else {{
      const hasMatch = p.lha_list.some(l => {{
        const okS = selStatus === 'ALL' || l.status === selStatus;
        const okT = selTahun2 === 'ALL' || l.tahun_pajak == selTahun2;
        return okS && okT;
      }});
      if (hasMatch) matches[lv]++;
    }}
  }});

  const names = ['Induk','Anak','Cucu','Cicit','Lv.4','Lv.5','Lv.6','Lv.7'];
  const maxTotal = Math.max(...Object.values(totals), 1);
  let html = '';
  for (let i = 0; i <= 7; i++) {{
    const tot = totals[i] || 0;
    if (tot === 0) continue;
    const mat = matches[i] || 0;
    const lc  = LEVEL_COLORS[i];
    const pct = (tot / maxTotal * 100).toFixed(0);
    const isFiltered = !(selStatus === 'ALL' && selTahun2 === 'ALL');
    // Label: "Anak" atau "Anak (5/3)"
    const chipLabel = isFiltered
      ? names[i] + ' <span style="font-size:9px;opacity:.85">(' + mat + '/' + tot + ')</span>'
      : names[i] + ' <span style="font-size:9px;opacity:.75">(' + tot + ')</span>';

    html += '<div class="lvl-row">';
    html += '<div class="lvl-chip" style="background:' + lc.bg + ';color:' + lc.text + '">' + chipLabel + '</div>';
    // Bar berdasarkan total, overlay match
    html += '<div class="lvl-track" style="position:relative;">';
    html += '<div class="lvl-fill" style="width:' + pct + '%;background:' + lc.bg + ';opacity:.3;position:absolute;left:0;top:0;height:100%"></div>';
    if (isFiltered) {{
      const pctM = (mat / maxTotal * 100).toFixed(0);
      html += '<div class="lvl-fill" style="width:' + pctM + '%;background:' + lc.bg + ';position:absolute;left:0;top:0;height:100%"></div>';
    }} else {{
      html += '<div class="lvl-fill" style="width:' + pct + '%;background:' + lc.bg + ';position:absolute;left:0;top:0;height:100%"></div>';
    }}
    html += '</div>';
    html += '<div class="lvl-count">' + (isFiltered ? mat + '<span style="color:#94A3B8;font-weight:400">/' + tot + '</span>' : tot) + '</div>';
    html += '</div>';
  }}
  document.getElementById('grup_level-dist-body').innerHTML = html || '<div class="empty-msg">Tidak ada data</div>';
}}

// ── RENDER TABLE ──────────────────────────────────────────────────────────
let expandedNodes = new Set();

function renderTable(ids, selStatus, selTahun2) {{
  // ids = semua perusahaan dalam grup (tidak disaring LHA)
  // Urutkan: root dulu, lalu DFS
  const allIds = [...ids];
  const roots  = allIds.filter(id => !PERUSAHAAN[id].id_induk || !ids.has(PERUSAHAAN[id].id_induk));

  function dfsOrder(id, result) {{
    result.push(id);
    (PERUSAHAAN[id].children || []).forEach(cid => {{
      if (ids.has(cid)) dfsOrder(cid, result);
    }});
    return result;
  }}

  const ordered = [];
  roots.forEach(rid => dfsOrder(rid, ordered));
  allIds.forEach(id => {{ if (!ordered.includes(id)) ordered.push(id); }});

  document.getElementById('grup_tbl-count').textContent = ordered.length + ' perusahaan';

  let rows = '';
  let rowNum = 0;
  ordered.forEach(id => {{
    const p  = PERUSAHAAN[id];
    const lv = p.level;
    const lc = LEVEL_COLORS[Math.min(lv, 7)];
    const hasChildren = (p.children || []).some(cid => ids.has(cid));
    const isExpanded  = expandedNodes.has(id);
    const parentInIds = p.id_induk && ids.has(p.id_induk);

    // Collapse visibility
    let cls = 'perusahaan-row child-row';
    let ancestorId = p.id_induk;
    while (ancestorId && ids.has(ancestorId)) {{
      if (!expandedNodes.has(ancestorId)) {{ cls += ' collapsed'; break; }}
      ancestorId = PERUSAHAAN[ancestorId] ? PERUSAHAAN[ancestorId].id_induk : null;
    }}

    rowNum++;
    const bgRow = lv % 2 === 0 ? '' : 'style="background:#FAFCFF"';

    // Hitung LHA sesuai filter (hanya kolom LHA yang berubah, perusahaan tetap tampil)
    const lhaCalc = calcLha(p, selStatus, selTahun2);
    const {{ terbit, belum, pot, real }} = lhaCalc;

    // Indentasi & expand button
    let indent = '';
    for (let i = 0; i < lv; i++) indent += '<span class="indent-line">&nbsp;&nbsp;&nbsp;&nbsp;</span>';
    let treeIcon = lv > 0 ? '<span class="tree-icon">└</span>' : '';
    let expBtn   = hasChildren
      ? '<button class="expand-btn" data-expand="' + id + '">' + (isExpanded ? '−' : '+') + '</button>'
      : '<span style="display:inline-block;width:22px"></span>';

    const kepChip = p.kepemilikan > 0
      ? '<div class="pct-wrap"><div class="pct-bar"><div class="pct-fill" style="width:' + Math.min(p.kepemilikan,100) + '%"></div></div><span class="pct-text">' + p.kepemilikan.toFixed(0) + '%</span></div>'
      : '<span style="color:#CBD5E1;font-size:10px">—</span>';

    const statusBadge = p.status_aktif === 'AKTIF'
      ? '<span class="badge badge-green">✔ Aktif</span>'
      : '<span class="badge badge-red">✖ ' + (p.status_aktif||'?') + '</span>';

    const terbitBadge = terbit > 0 ? '<span class="badge badge-green">' + terbit + '</span>' : '<span style="color:#CBD5E1">—</span>';
    const belumBadge  = belum  > 0 ? '<span class="badge badge-red">'   + belum  + '</span>' : '<span style="color:#CBD5E1">—</span>';

    rows += '<tr class="' + cls + '" data-id="' + id + '" data-parent="' + (p.id_induk||'') + '" data-modal="' + id + '" ' + bgRow + '>';
    rows += '<td style="color:#94A3B8;font-size:10px">' + rowNum + '</td>';
    rows += '<td><div class="nama-cell">' + indent + treeIcon + expBtn;
    rows += '<div>';
    rows += '<div style="font-weight:700;color:#0F172A;font-size:12px">' + escHtml(p.nama) + '</div>';
    rows += '<div style="font-size:10px;color:#94A3B8;font-family:monospace">' + escHtml(p.id) + '</div>';
    rows += '</div></div></td>';
    rows += '<td><span class="badge" style="background:' + (grupColorMap[p.nama_grup]||'#666') + '22;color:' + (grupColorMap[p.nama_grup]||'#333') + '">' + escHtml(p.nama_grup||'—') + '</span></td>';
    rows += '<td><span class="mono" style="font-size:10px;color:#475569">' + escHtml(p.npwp||'—') + '</span></td>';
    rows += '<td style="color:#475569">' + escHtml(p.sektor||'—') + '</td>';
    rows += '<td style="font-size:11px;color:#475569">' + escHtml(p.kpp||'—') + '</td>';
    rows += '<td style="text-align:center">' + kepChip + '</td>';
    rows += '<td style="text-align:center">' + statusBadge + '</td>';
    rows += '<td style="text-align:center">' + terbitBadge + '</td>';
    rows += '<td style="text-align:center">' + belumBadge + '</td>';
    rows += '<td style="text-align:right;font-weight:600;color:#0052A5">' + (pot > 0 ? fmtRp(pot) : '—') + '</td>';
    rows += '<td style="text-align:right;font-weight:600;color:#0A7C4F">' + (real > 0 ? fmtRp(real) : '—') + '</td>';
    rows += '</tr>';
  }});

  document.getElementById('grup_main-tbody').innerHTML = rows || '<tr><td colspan="12" class="empty-msg">Tidak ada data sesuai filter</td></tr>';
}}

// toggleExpand handled via event delegation in init()

function escHtml(s) {{
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}}

// ── MODAL ─────────────────────────────────────────────────────────────────
function openModal(id) {{
  const p = PERUSAHAAN[id];
  if (!p) return;
  document.getElementById('grup_modal-nama').textContent = p.nama;
  document.getElementById('grup_modal-npwp').textContent = 'NPWP: ' + (p.npwp||'—') + '  |  ' + p.id;
  const lc = LEVEL_COLORS[Math.min(p.level,7)];

  const selStatus = document.getElementById('grup_sel-status').value;
  const selTahun2 = document.getElementById('grup_sel-tahun').value;

  const infos = [
    ['Nama Grup',   p.nama_grup||'—'],
    ['Level',       `<span class="level-tag" style="background:${{lc.bg}};color:${{lc.text}}">${{LEVEL_COLORS[p.level]?.name||'Lv.'+p.level}}</span>`],
    ['Sektor',      p.sektor||'—'],
    ['KPP',         p.kpp||'—'],
    ['Status',      p.status_aktif||'—'],
    ['Berdiri',     p.tahun_berdiri||'—'],
    ['Kepemilikan', p.kepemilikan ? p.kepemilikan + '%' : '—'],
    ['Keterangan',  p.keterangan||'—'],
    ['ID Induk',    p.id_induk||'— (Induk)'],
  ];
  let igHtml = '';
  infos.forEach(([label,val]) => {{
    igHtml += '<div class="info-item"><label>' + label + '</label><span>' + val + '</span></div>';
  }});
  document.getElementById('grup_modal-info-grid').innerHTML = igHtml;

  // LHA table
  const lhaFiltered = p.lha_list.filter(l => {{
    const okS = selStatus === 'ALL' || l.status === selStatus;
    const okT = selTahun2 === 'ALL' || l.tahun_pajak == selTahun2;
    return okS && okT;
  }});

  if (lhaFiltered.length === 0) {{
    document.getElementById('grup_modal-lha-wrap').innerHTML =
      '<div class="no-lha">📭 Tidak ada LHA' + (selStatus!=='ALL'||selTahun2!=='ALL' ? ' sesuai filter aktif' : '') + '</div>';
  }} else {{
    let tHtml = '<table class="modal-lha-table"><thead><tr>';
    ['ID LHA','Tahun','Jenis','Status','Tgl Terbit','No. Dokumen','Potensi PPh','Potensi PPN','Total Potensi','Realisasi','Keterangan'].forEach(h => {{
      tHtml += '<th>' + h + '</th>';
    }});
    tHtml += '</tr></thead><tbody>';
    lhaFiltered.forEach(l => {{
      const cls2 = l.status === 'DITERBITKAN' ? 'terbit' : 'belum';
      const stBadge = l.status === 'DITERBITKAN'
        ? '<span class="badge badge-green">✔ Diterbitkan</span>'
        : '<span class="badge badge-red">⏳ Belum</span>';
      tHtml += '<tr class="' + cls2 + '">';
      tHtml += '<td class="mono" style="font-size:10px">' + escHtml(l.id_lha) + '</td>';
      tHtml += '<td style="text-align:center;font-weight:700">' + l.tahun_pajak + '</td>';
      tHtml += '<td><span class="badge badge-blue">' + escHtml(l.jenis) + '</span></td>';
      tHtml += '<td>' + stBadge + '</td>';
      tHtml += '<td style="text-align:center;font-size:10px">' + escHtml(l.tgl_terbit||'—') + '</td>';
      tHtml += '<td class="mono" style="font-size:10px">' + escHtml(l.no_lha||'—') + '</td>';
      tHtml += '<td style="text-align:right">' + fmtRp(l.potensi_pph) + '</td>';
      tHtml += '<td style="text-align:right">' + fmtRp(l.potensi_ppn) + '</td>';
      tHtml += '<td style="text-align:right;font-weight:700;color:#0052A5">' + fmtRp(l.potensi_total) + '</td>';
      tHtml += '<td style="text-align:right;font-weight:700;color:#0A7C4F">' + fmtRp(l.realisasi) + '</td>';
      tHtml += '<td style="font-size:10px;color:#64748B">' + escHtml(l.keterangan||'—') + '</td>';
      tHtml += '</tr>';
    }});
    tHtml += '</tbody></table>';
    document.getElementById('grup_modal-lha-wrap').innerHTML = tHtml;
  }}

  document.getElementById('grup_modal-overlay').classList.add('open');
}}

function closeModal(e) {{
  if (e.target === document.getElementById('grup_modal-overlay'))
    document.getElementById('grup_modal-overlay').classList.remove('open');
}}
function closeModalBtn() {{
  document.getElementById('grup_modal-overlay').classList.remove('open');
}}

// ── NETWORK / PETA JARINGAN ───────────────────────────────────────────────
let netTransform = {{ x:0, y:0, scale:1 }};
let sidebarVisible = true;

function toggleSidebar() {{
  sidebarVisible = !sidebarVisible;
  const sb = document.getElementById('grup_net-sidebar');
  sb.classList.toggle('hidden', !sidebarVisible);
  document.getElementById('grup_sidebar-icon').textContent  = sidebarVisible ? '◀' : '▶';
  document.getElementById('grup_sidebar-label').textContent = sidebarVisible ? 'Sembunyikan' : 'Tampilkan';
  setTimeout(renderNetwork, 300);
}}

function buildSidebar() {{
  let html = '';
  grupList.forEach((g,gi) => {{
    const ps = Object.values(PERUSAHAAN).filter(p => p.nama_grup === g);
    const terbit2 = ps.reduce((s,p) => s + p.lha_terbit, 0);
    const belum2  = ps.reduce((s,p) => s + p.lha_belum,  0);
    const pot2    = ps.reduce((s,p) => s + p.total_potensi, 0);
    const real2   = ps.reduce((s,p) => s + p.total_realisasi, 0);
    const color   = grupColorMap[g] || '#0052A5';
    html += '<div class="net-grup-card" data-grup="' + escHtml(g) + '">';
    html += '<div class="net-grup-card-head" style="background:' + color + '18">';
    html += '<div class="net-grup-dot" style="background:' + color + '"></div>';
    html += '<span style="color:' + color + ';font-size:11px;font-weight:700">' + escHtml(g) + '</span>';
    html += '</div>';
    html += '<div class="net-grup-stat">';
    html += '<div class="net-stat-item"><span>' + ps.length + '</span>Perusahaan</div>';
    html += '<div class="net-stat-item"><span style="color:#0A7C4F">' + terbit2 + '</span>LHA Terbit</div>';
    html += '<div class="net-stat-item"><span style="color:#0052A5">' + fmtRp(pot2) + '</span>Potensi</div>';
    html += '<div class="net-stat-item"><span style="color:#0A7C4F">' + fmtRp(real2) + '</span>Realisasi</div>';
    html += '</div>';
    html += '</div>';
  }});
  document.getElementById('grup_net-sidebar-body').innerHTML = html;
}}

function selectGrupNet(g) {{
  document.getElementById('grup_net-sel-grup').value = g;
  // Highlight card
  document.querySelectorAll('.net-grup-card').forEach((c,i) => {{
    c.classList.toggle('selected', grupList[i] === g);
  }});
  renderNetwork();
}}

function buildLegend() {{
  let html = '<div class="net-legend-title">Level Hierarki</div>';
  LEVEL_COLORS.forEach((lc,i) => {{
    html += '<div class="net-legend-item"><div class="net-legend-dot" style="background:' + lc.bg + ';border:1px solid #ccc"></div><span>' + lc.name + '</span></div>';
  }});
  document.getElementById('grup_net-legend').innerHTML = html;
}}

// Tree layout — top-down
function computeTreeLayout(rootId, ids) {{
  const nodeMap = {{}};
  let maxDepth = 0;

  function dfs(id, depth) {{
    if (!ids.has(id)) return;
    if (maxDepth < depth) maxDepth = depth;
    nodeMap[id] = {{ id, depth, children: [], x: 0, y: 0, subtreeW: 0 }};
    const p = PERUSAHAAN[id];
    (p.children||[]).forEach(cid => {{
      if (ids.has(cid)) {{
        dfs(cid, depth+1);
        if (nodeMap[cid]) nodeMap[id].children.push(cid);
      }}
    }});
  }}
  dfs(rootId, 0);

  const NODE_W = 160, NODE_H = 56, GAP_X = 24, GAP_Y = 80;

  function measureSubtree(id) {{
    const n = nodeMap[id];
    if (!n) return 0;
    if (n.children.length === 0) {{ n.subtreeW = NODE_W; return NODE_W; }}
    let total = 0;
    n.children.forEach(cid => {{ total += measureSubtree(cid) + GAP_X; }});
    total -= GAP_X;
    n.subtreeW = Math.max(NODE_W, total);
    return n.subtreeW;
  }}
  measureSubtree(rootId);

  function assignPos(id, startX, depth) {{
    const n = nodeMap[id];
    if (!n) return;
    n.y = depth * (NODE_H + GAP_Y) + GAP_Y;
    if (n.children.length === 0) {{
      n.x = startX + n.subtreeW/2;
      return;
    }}
    let cx = startX;
    n.children.forEach(cid => {{
      const cn = nodeMap[cid];
      if (!cn) return;
      assignPos(cid, cx, depth+1);
      cx += cn.subtreeW + GAP_X;
    }});
    const first = nodeMap[n.children[0]];
    const last  = nodeMap[n.children[n.children.length-1]];
    n.x = (first.x + last.x) / 2;
  }}
  assignPos(rootId, 0, 0);

  return {{ nodeMap, maxDepth, NODE_W, NODE_H }};
}}

function renderNetwork() {{
  const selGrup = document.getElementById('grup_net-sel-grup').value;
  const svg = document.getElementById('grup_network-svg');
  svg.innerHTML = '';

  // Highlight sidebar card
  document.querySelectorAll('.net-grup-card').forEach((c,i) => {{
    c.classList.toggle('selected', grupList[i] === selGrup);
  }});

  const filteredPs = selGrup === 'ALL'
    ? Object.values(PERUSAHAAN)
    : Object.values(PERUSAHAAN).filter(p => p.nama_grup === selGrup);
  const ids = new Set(filteredPs.map(p => p.id));

  const roots = filteredPs.filter(p => !p.id_induk || !ids.has(p.id_induk));

  const TREE_GAP = 80;
  let allNodes = [], allEdges = [];
  let offsetX = 40;

  roots.forEach(r => {{
    const {{ nodeMap, NODE_W, NODE_H }} = computeTreeLayout(r.id, ids);
    Object.values(nodeMap).forEach(n => {{
      n.x += offsetX;
      allNodes.push({{ ...n, data: PERUSAHAAN[n.id] }});
      n.children.forEach(cid => {{
        if (nodeMap[cid]) allEdges.push({{ from: n, to: nodeMap[cid] }});
      }});
    }});
    const rootNode = nodeMap[r.id];
    if (rootNode) offsetX += rootNode.subtreeW + TREE_GAP;
  }});

  if (allNodes.length === 0) {{
    svg.innerHTML = '<text x="50%" y="50%" text-anchor="middle" font-family="Plus Jakarta Sans" font-size="14" fill="#94A3B8">Tidak ada data untuk ditampilkan</text>';
    return;
  }}

  const maxX = Math.max(...allNodes.map(n => n.x + 80)) + 40;
  const maxY = Math.max(...allNodes.map(n => n.y + 60)) + 40;

  const NODE_W = 156, NODE_H = 54;
  const NS = 'http://www.w3.org/2000/svg';

  const g = document.createElementNS(NS, 'g');
  g.setAttribute('id', 'net-g');
  svg.appendChild(g);

  // Definisi gradient & shadow
  const defs = document.createElementNS(NS, 'defs');
  defs.innerHTML = `
    <filter id="node-shadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="2" stdDeviation="3" flood-color="rgba(0,30,80,0.15)"/>
    </filter>
    ${{LEVEL_COLORS.map((lc,i) => `<linearGradient id="grad${{i}}" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0%" stop-color="${{lc.bg}}"/>
      <stop offset="100%" stop-color="${{lc.bg}}dd"/>
    </linearGradient>`).join('')}}
  `;
  svg.insertBefore(defs, g);

  // Edges
  allEdges.forEach(e => {{
    const x1 = e.from.x, y1 = e.from.y + NODE_H/2;
    const x2 = e.to.x,   y2 = e.to.y   - NODE_H/2;
    const my = (y1+y2)/2;
    const path = document.createElementNS(NS, 'path');
    path.setAttribute('d', `M${{x1}},${{y1}} C${{x1}},${{my}} ${{x2}},${{my}} ${{x2}},${{y2}}`);
    path.setAttribute('fill', 'none');
    path.setAttribute('stroke', '#94A3B8');
    path.setAttribute('stroke-width', '1.5');
    path.setAttribute('stroke-dasharray', '4,3');
    path.setAttribute('opacity', '0.6');
    g.appendChild(path);
  }});

  // Nodes
  allNodes.forEach(n => {{
    const p   = n.data;
    const lv  = Math.min(p.level, 7);
    const lc  = LEVEL_COLORS[lv];
    const x   = n.x - NODE_W/2, y = n.y - NODE_H/2;

    const ng = document.createElementNS(NS, 'g');
    ng.setAttribute('transform', `translate(${{x}},${{y}})`);
    ng.setAttribute('cursor', 'pointer');
    ng.style.filter = 'url(#node-shadow)';

    // Card background
    const rect = document.createElementNS(NS, 'rect');
    rect.setAttribute('width', NODE_W); rect.setAttribute('height', NODE_H);
    rect.setAttribute('rx', '8'); rect.setAttribute('ry', '8');
    rect.setAttribute('fill', 'white');
    rect.setAttribute('stroke', lc.bg);
    rect.setAttribute('stroke-width', '1.5');
    ng.appendChild(rect);

    // Top color bar
    const bar = document.createElementNS(NS, 'rect');
    bar.setAttribute('width', NODE_W); bar.setAttribute('height', '5');
    bar.setAttribute('rx', '8'); bar.setAttribute('ry', '8');
    bar.setAttribute('fill', `url(#grad${{lv}})`);
    ng.appendChild(bar);

    // Level badge
    const lvRect = document.createElementNS(NS, 'rect');
    lvRect.setAttribute('x', '8'); lvRect.setAttribute('y', '8');
    lvRect.setAttribute('width', '36'); lvRect.setAttribute('height', '14');
    lvRect.setAttribute('rx', '7'); lvRect.setAttribute('fill', lc.bg);
    ng.appendChild(lvRect);
    const lvTxt = document.createElementNS(NS, 'text');
    lvTxt.setAttribute('x', '26'); lvTxt.setAttribute('y', '19');
    lvTxt.setAttribute('text-anchor', 'middle');
    lvTxt.setAttribute('font-family', 'Plus Jakarta Sans, sans-serif');
    lvTxt.setAttribute('font-size', '8');
    lvTxt.setAttribute('font-weight', '700');
    lvTxt.setAttribute('fill', lc.text);
    lvTxt.textContent = 'Lv.' + lv;
    ng.appendChild(lvTxt);

    // LHA badge if any
    if (p.lha_terbit > 0 || p.lha_belum > 0) {{
      const lhaBg = document.createElementNS(NS, 'rect');
      lhaBg.setAttribute('x', String(NODE_W-44)); lhaBg.setAttribute('y', '8');
      lhaBg.setAttribute('width', '36'); lhaBg.setAttribute('height', '14');
      lhaBg.setAttribute('rx', '7');
      lhaBg.setAttribute('fill', p.lha_belum > 0 ? '#FEE2E2' : '#DCFCE7');
      ng.appendChild(lhaBg);
      const lhaTxt = document.createElementNS(NS, 'text');
      lhaTxt.setAttribute('x', String(NODE_W-26)); lhaTxt.setAttribute('y', '19');
      lhaTxt.setAttribute('text-anchor', 'middle');
      lhaTxt.setAttribute('font-family', 'Plus Jakarta Sans, sans-serif');
      lhaTxt.setAttribute('font-size', '8');
      lhaTxt.setAttribute('font-weight', '700');
      lhaTxt.setAttribute('fill', p.lha_belum > 0 ? '#B91C1C' : '#0A7C4F');
      lhaTxt.textContent = p.lha_terbit + '+' + p.lha_belum + ' LHA';
      ng.appendChild(lhaTxt);
    }}

    // Company name (truncate)
    const nama = p.nama.length > 22 ? p.nama.substring(0,20) + '…' : p.nama;
    const namaTxt = document.createElementNS(NS, 'text');
    namaTxt.setAttribute('x', String(NODE_W/2)); namaTxt.setAttribute('y', '35');
    namaTxt.setAttribute('text-anchor', 'middle');
    namaTxt.setAttribute('font-family', 'Plus Jakarta Sans, sans-serif');
    namaTxt.setAttribute('font-size', '10');
    namaTxt.setAttribute('font-weight', '700');
    namaTxt.setAttribute('fill', '#0F172A');
    namaTxt.textContent = nama;
    ng.appendChild(namaTxt);

    // Kepemilikan
    if (p.kepemilikan > 0) {{
      const kepTxt = document.createElementNS(NS, 'text');
      kepTxt.setAttribute('x', String(NODE_W/2)); kepTxt.setAttribute('y', '47');
      kepTxt.setAttribute('text-anchor', 'middle');
      kepTxt.setAttribute('font-family', 'DM Mono, monospace');
      kepTxt.setAttribute('font-size', '9');
      kepTxt.setAttribute('fill', '#475569');
      kepTxt.textContent = p.kepemilikan + '% kepemilikan';
      ng.appendChild(kepTxt);
    }}

    // Hover & click
    ng.addEventListener('mouseenter', function(e) {{
      rect.setAttribute('fill', '#F0F7FF');
      showTooltip(e, p);
    }});
    ng.addEventListener('mousemove', function(e) {{ moveTooltip(e); }});
    ng.addEventListener('mouseleave', function() {{
      rect.setAttribute('fill', 'white');
      hideTooltip();
    }});
    ng.addEventListener('click', function() {{ openModal(p.id); }});

    g.appendChild(ng);
  }});

  // Set viewBox with current transform
  const pad = 20;
  svg.setAttribute('viewBox', `${{-netTransform.x/netTransform.scale - pad}} ${{-netTransform.y/netTransform.scale - pad}} ${{(maxX+pad*2)/netTransform.scale}} ${{(maxY+pad*2)/netTransform.scale}}`);
  svg._baseViewBox = {{ x: -pad, y: -pad, w: maxX+pad*2, h: maxY+pad*2 }};

  // Drag to pan
  setupSvgDrag(svg);
}}

function zoomNetwork(factor) {{
  netTransform.scale *= factor;
  const svg = document.getElementById('grup_network-svg');
  const vb = svg._baseViewBox;
  if (!vb) return;
  const cw = vb.w / netTransform.scale;
  const ch = vb.h / netTransform.scale;
  svg.setAttribute('viewBox', `${{vb.x}} ${{vb.y}} ${{cw}} ${{ch}}`);
}}

function resetNetwork() {{
  netTransform = {{ x:0, y:0, scale:1 }};
  renderNetwork();
}}

function setupSvgDrag(svg) {{
  let dragging = false, startX, startY;
  svg.addEventListener('mousedown', e => {{
    dragging = true;
    startX = e.clientX; startY = e.clientY;
    svg.style.cursor = 'grabbing';
  }});
  window.addEventListener('mouseup', () => {{ dragging = false; svg.style.cursor = ''; }});
  svg.addEventListener('mousemove', e => {{
    if (!dragging || !svg._baseViewBox) return;
    const vb = svg._baseViewBox;
    const rect2 = svg.getBoundingClientRect();
    const scaleX = (vb.w/netTransform.scale) / rect2.width;
    const scaleY = (vb.h/netTransform.scale) / rect2.height;
    const dx = (e.clientX - startX) * scaleX;
    const dy = (e.clientY - startY) * scaleY;
    startX = e.clientX; startY = e.clientY;
    const [cx,cy,cw,ch] = svg.getAttribute('viewBox').split(' ').map(Number);
    svg.setAttribute('viewBox', `${{cx-dx}} ${{cy-dy}} ${{cw}} ${{ch}}`);
  }});
}}

// ── TOOLTIP ───────────────────────────────────────────────────────────────
function showTooltip(e, p) {{
  const tt = document.getElementById('grup_net-tooltip');
  const lha_info = p.lha_terbit + ' terbit, ' + p.lha_belum + ' belum';
  tt.innerHTML = '<strong>' + escHtml(p.nama) + '</strong>'
    + escHtml(p.sektor||'') + ' &nbsp;|&nbsp; ' + escHtml(p.kpp||'') + '<br>'
    + 'LHA: ' + lha_info + '<br>'
    + 'Potensi: ' + fmtRp(p.total_potensi) + '&nbsp; Realisasi: ' + fmtRp(p.total_realisasi);
  tt.style.display = 'block';
  moveTooltip(e);
}}
function moveTooltip(e) {{
  const tt = document.getElementById('grup_net-tooltip');
  tt.style.left = (e.clientX + 14) + 'px';
  tt.style.top  = (e.clientY + 14) + 'px';
}}
function hideTooltip() {{
  document.getElementById('grup_net-tooltip').style.display = 'none';
}}

// ── INIT ──────────────────────────────────────────────────────────────────
(function init() {{
  const ids       = new Set(getGrupIds());
  const selStatus = document.getElementById('grup_sel-status').value;
  const selTahun2 = document.getElementById('grup_sel-tahun').value;
  // Auto-expand level 0 nodes
  Object.values(PERUSAHAAN).forEach(p => {{ if (p.level === 0) expandedNodes.add(p.id); }});
  updateKPI(ids, selStatus, selTahun2);
  updateLevelDist(ids, selStatus, selTahun2);
  renderTable(ids, selStatus, selTahun2);
  buildSidebar();
  buildLegend();

  // ── Event delegation: tabel row click & expand button ──
  document.getElementById('grup_main-tbody').addEventListener('click', function(e) {{
    const expBtn = e.target.closest('[data-expand]');
    if (expBtn) {{
      e.stopPropagation();
      const eid = expBtn.getAttribute('data-expand');
      if (expandedNodes.has(eid)) expandedNodes.delete(eid);
      else expandedNodes.add(eid);
      const filteredIds = new Set(getGrupIds());
      const ss = document.getElementById('grup_sel-status').value;
      const st = document.getElementById('grup_sel-tahun').value;
      renderTable(filteredIds, ss, st);
      return;
    }}
    const row = e.target.closest('[data-modal]');
    if (row) openModal(row.getAttribute('data-modal'));
  }});

  // ── Event delegation: sidebar grup card click ──
  document.getElementById('grup_net-sidebar-body').addEventListener('click', function(e) {{
    const card = e.target.closest('[data-grup]');
    if (card) selectGrupNet(card.getAttribute('data-grup'));
  }});
}})();'''

def read_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    # Sheet Perusahaan
    ws_p = wb["Perusahaan"]
    perusahaan_rows = []
    header_found = False
    p_headers = []
    for row in ws_p.iter_rows(values_only=True):
        if not header_found:
            if row[0] == "ID_PERUSAHAAN":
                p_headers = [str(c).strip() if c else "" for c in row]
                header_found = True
            continue
        if row[0] is None:
            continue
        perusahaan_rows.append(dict(zip(p_headers, row)))

    # Sheet LHA
    ws_l = wb["LHA"]
    lha_rows = []
    header_found2 = False
    l_headers = []
    for row in ws_l.iter_rows(values_only=True):
        if not header_found2:
            if row[0] == "ID_LHA":
                l_headers = [str(c).strip() if c else "" for c in row]
                header_found2 = True
            continue
        if row[0] is None:
            continue
        lha_rows.append(dict(zip(l_headers, row)))

    wb.close()
    return perusahaan_rows, lha_rows


def build_data_model(perusahaan_rows, lha_rows):
    """Bangun dictionary PERUSAHAAN yang kaya dengan info LHA & children."""
    perusahaan = {}
    for r in perusahaan_rows:
        pid = str(r.get("ID_PERUSAHAAN", "")).strip()
        if not pid:
            continue
        perusahaan[pid] = {
            "id":           pid,
            "id_induk":     str(r.get("ID_INDUK", "") or "").strip(),
            "nama":         str(r.get("NAMA_PERUSAHAAN", "") or "").strip(),
            "npwp":         str(r.get("NPWP", "") or "").strip(),
            "level":        int(r.get("LEVEL", 0) or 0),
            "nama_grup":    str(r.get("NAMA_GRUP", "") or "").strip(),
            "sektor":       str(r.get("SEKTOR_USAHA", "") or "").strip(),
            "kpp":          str(r.get("KPP_TERDAFTAR", "") or "").strip(),
            "status_aktif": str(r.get("STATUS_AKTIF", "") or "").strip(),
            "kepemilikan":  float(r.get("KEPEMILIKAN_PCT", 0) or 0),
            "keterangan":   str(r.get("KETERANGAN", "") or "").strip(),
            "children":     [],
            "lha_list":     [],
            "lha_terbit":   0,
            "lha_belum":    0,
            "total_potensi":    0.0,
            "total_realisasi":  0.0,
        }

    # Pasang children
    for pid, p in perusahaan.items():
        if p["id_induk"] and p["id_induk"] in perusahaan:
            perusahaan[p["id_induk"]]["children"].append(pid)

    # Pasang LHA
    for r in lha_rows:
        pid = str(r.get("ID_PERUSAHAAN", "")).strip()
        if pid not in perusahaan:
            continue
        tgl = r.get("TGL_TERBIT", "")
        if tgl and not isinstance(tgl, str):
            try:
                tgl = tgl.strftime("%d/%m/%Y")
            except Exception:
                tgl = str(tgl)
        lha = {
            "id_lha":        str(r.get("ID_LHA", "") or "").strip(),
            "tahun_pajak":   int(r.get("TAHUN_PAJAK", 0) or 0),
            "jenis":         str(r.get("JENIS_LHA", "") or "").strip(),
            "status":        str(r.get("STATUS_LHA", "") or "").strip(),
            "tgl_terbit":    str(tgl or "").strip(),
            "no_lha":        str(r.get("NO_LHA", "") or "").strip(),
            "potensi_pph":   float(r.get("POTENSI_PPH", 0) or 0),
            "potensi_ppn":   float(r.get("POTENSI_PPN", 0) or 0),
            "potensi_total": float(r.get("POTENSI_TOTAL", 0) or 0),
            "realisasi":     float(r.get("REALISASI", 0) or 0),
            "keterangan":    str(r.get("KETERANGAN", "") or "").strip(),
        }
        perusahaan[pid]["lha_list"].append(lha)
        perusahaan[pid]["total_potensi"]   += lha["potensi_total"]
        perusahaan[pid]["total_realisasi"] += lha["realisasi"]
        if lha["status"] == "DITERBITKAN":
            perusahaan[pid]["lha_terbit"] += 1
        else:
            perusahaan[pid]["lha_belum"] += 1

    return perusahaan


# ─────────────────────────────────────────────────────────────────────────────
# GENERATOR HTML
# ─────────────────────────────────────────────────────────────────────────────

def fmt_rp(val):
    """Format angka ke Rp dalam miliar/juta/ribu (Rp000)."""
    if val >= 1_000_000:
        return f"Rp {val/1_000_000:.1f} M"
    elif val >= 1_000:
        return f"Rp {val/1_000:.1f} rb"
    elif val > 0:
        return f"Rp {val:.0f}"
    return "Rp 0"



def build_grup_module(perusahaan, source_filename, log=None):

        now = datetime.datetime.now().strftime("%d %B %Y %H:%M")
        p_json = json.dumps(perusahaan, ensure_ascii=False, default=str)

        # Hitung statistik global
        total_perusahaan = len(perusahaan)
        total_lha_terbit = sum(p["lha_terbit"] for p in perusahaan.values())
        total_lha_belum  = sum(p["lha_belum"]  for p in perusahaan.values())
        total_potensi    = sum(p["total_potensi"]   for p in perusahaan.values())
        total_realisasi  = sum(p["total_realisasi"] for p in perusahaan.values())
        efektivitas      = (total_realisasi / total_potensi * 100) if total_potensi > 0 else 0

        grup_list = sorted(set(p["nama_grup"] for p in perusahaan.values() if p["nama_grup"]))
        total_grup = len(grup_list)


        dash_html_content = DASH_HTML.replace('{summary_counts["grup"]}', str(total_grup))
        dash_html_content = dash_html_content.replace('{summary_counts["perusahaan"]}', str(total_perusahaan))
        dash_html_content = dash_html_content.replace('{summary_counts["lha_terbit"]}', str(total_lha_terbit))
        dash_html_content = dash_html_content.replace('{summary_counts["lha_belum"]}', str(total_lha_belum))
        dash_html_content = dash_html_content.replace('{fmt_rp(summary_counts["potensi"])}', fmt_rp(total_potensi))
        dash_html_content = dash_html_content.replace('{fmt_rp(summary_counts["realisasi"])}', fmt_rp(total_realisasi))
        dash_html_content = dash_html_content.replace('{summary_counts["eff"]}', f"{efektivitas:.1f}")

        opt_grup = ""
        for g in sorted(list(grup_list)):
            opt_grup += f'<option value="{g}">{g}</option>'

        dash_html_content = dash_html_content.replace('{opt_grup}', opt_grup)

        net_html_content = NET_HTML.replace('{opt_grup}', opt_grup)

        import json as _json
        import base64 as _base64
        p_json = _json.dumps(perusahaan, ensure_ascii=False, default=str)
        b64_data = _base64.b64encode(p_json.encode('utf-8')).decode('ascii')
        js_content = JS_TEMPLATE.replace('{b64_data}', b64_data)

        return {
            "css": CSS_CODE,
            "body_dash": dash_html_content,
            "body_net": net_html_content,
            "body_extra": EXTRA_HTML,
            "js": js_content,
            "js_dash": js_content,
            "js_peta": ""
        }

def generate_html(perusahaan, source_filename, log=None):
    return build_grup_module(perusahaan, source_filename, log)
