import base64

def build_super_html(modules: list[dict], title: str, generated_at: str, encode_b64: bool = True) -> str:
    master_css = """
    :root {
        --djp-dark: #0a2346;
        --djp-gold: #c8a84b;
        --djp-gold2: #e8c96a;
        --djp-bg: #f0f2f5;
        --accent: #1d4ed8;
        --sidebar-w: 240px;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Plus Jakarta Sans', sans-serif; background-color: var(--djp-bg); overflow: hidden; height: 100vh; display: flex; flex-direction: column; }
    #master-header { height: 56px; background-color: var(--djp-dark); border-bottom: 3px solid var(--djp-gold); display: flex; align-items: center; padding: 0 20px; color: white; z-index: 100; flex-shrink: 0; justify-content: space-between; }
    #master-header h1 { font-size: 16px; margin: 0; color: var(--djp-gold2); }
    .hdr-badge { background: var(--djp-gold); color: var(--djp-dark); padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 12px; margin-right: 10px; }
    .hdr-right { font-size: 11px; color: #aaa; text-align: right; }
    #app-body { display: flex; flex: 1; overflow: hidden; position: relative; }
    #sidebar { width: var(--sidebar-w); background-color: var(--djp-dark); transition: width 0.3s ease; display: flex; flex-direction: column; color: white; flex-shrink: 0; overflow-y: auto; overflow-x: hidden; position: relative; }
    #sidebar.collapsed { width: 0; }
    .sb-menu-item { display: flex; align-items: center; padding: 12px 15px; cursor: pointer; color: #ccc; font-size: 14px; border-left: 3px solid transparent; transition: all 0.2s; white-space: nowrap; }
    .sb-menu-item:hover { background: rgba(200,168,75,0.05); color: var(--djp-gold2); }
    .sb-menu-item.active { background: rgba(200,168,75,0.12); color: var(--djp-gold2); border-left-color: var(--djp-gold); font-weight: 600; }
    .sb-icon { margin-right: 10px; font-size: 16px; }
    .sb-subitems { background: rgba(0,0,0,0.15); display: none; }
    .sb-subitems.open { display: block; }
    .sb-sub-item { padding: 8px 15px 8px 40px; cursor: pointer; color: #aaa; font-size: 13px; transition: color 0.2s; white-space: nowrap; }
    .sb-sub-item:hover, .sb-sub-item.active { color: white; }
    .sb-sub-item.active { font-weight: 600; position: relative; }
    .sb-sub-item.active::before { content: '•'; position: absolute; left: 25px; color: var(--djp-gold); }
    #sb-toggle { position: absolute; bottom: 10px; right: 10px; background: rgba(255,255,255,0.1); border: none; color: white; width: 30px; height: 30px; border-radius: 4px; cursor: pointer; z-index: 10; display: flex; justify-content: center; align-items: center; }
    #sb-toggle:hover { background: rgba(255,255,255,0.2); }
    #content { flex: 1; display: flex; flex-direction: column; overflow: hidden; background: var(--djp-bg); }
    .module-panel { display: none; flex: 1; flex-direction: column; overflow: hidden; }
    .module-panel.active { display: flex; }
    .subtab-bar { display: flex; background: #fff; padding: 10px 20px; border-bottom: 1px solid #ddd; gap: 10px; overflow-x: auto; flex-shrink: 0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    .subtab { padding: 6px 16px; border-radius: 20px; border: 1px solid var(--djp-gold); background: transparent; color: var(--djp-dark); cursor: pointer; font-size: 13px; font-weight: 600; transition: all 0.2s; white-space: nowrap; }
    .subtab:hover { background: rgba(200,168,75,0.1); }
    .subtab.active { background: var(--djp-gold); color: var(--djp-dark); }
    .subtab-panel { display: none; flex: 1; overflow: auto; padding: 20px; }
    .subtab-panel.active { display: block; }
    @media (max-width: 768px) { #sidebar { width: 0; } #sidebar.open-mobile { width: var(--sidebar-w); position: absolute; height: 100%; z-index: 50; } }
    """

    all_css = master_css
    sidebar_html = ""
    modules_html = ""
    all_js = ""

    for i, mod in enumerate(modules):
        mod_id = mod['id']
        is_active = (i == 0)

        if 'css' in mod: all_css += f"\n/* CSS for {mod_id} */\n" + mod['css']
        if 'js' in mod: all_js += f"\n// JS for {mod_id}\n" + mod['js']

        active_cls = " active" if is_active else ""
        sidebar_html += f'<div class="sb-menu-item{active_cls}" onclick="switchModule(\'{mod_id}\')" id="sb-mod-{mod_id}"><span class="sb-icon">{mod["icon"]}</span><span class="sb-label">{mod["label"]}</span></div>'

        subitems_cls = " open" if is_active else ""
        sidebar_html += f'<div class="sb-subitems{subitems_cls}" id="sb-subs-{mod_id}">'
        for j, tab in enumerate(mod.get('sub_tabs', [])):
            tab_id = tab['id']
            tab_active = " active" if j == 0 else ""
            sidebar_html += f'<div class="sb-sub-item{tab_active}" onclick="switchSubTab(\'{mod_id}\', \'{tab_id}\')" id="sb-sub-{mod_id}-{tab_id}">{tab["label"]}</div>'
        sidebar_html += '</div>'

        panel_active = " active" if is_active else ""
        modules_html += f'<div class="module-panel{panel_active}" id="mod-{mod_id}">'

        modules_html += '<div class="subtab-bar">'
        for j, tab in enumerate(mod.get('sub_tabs', [])):
            tab_id = tab['id']
            tab_active = " active" if j == 0 else ""
            modules_html += f'<button class="subtab{tab_active}" id="tab-btn-{mod_id}-{tab_id}" onclick="switchSubTab(\'{mod_id}\', \'{tab_id}\')">{tab["label"]}</button>'
        modules_html += '</div>'

        for j, tab in enumerate(mod.get('sub_tabs', [])):
            tab_id = tab['id']
            tab_active = " active" if j == 0 else ""
            modules_html += f'<div class="subtab-panel{tab_active}" id="tab-panel-{mod_id}-{tab_id}">{tab["html"]}</div>'

        if 'body_extra' in mod: modules_html += mod['body_extra']
        modules_html += '</div>'

    master_js = f"""
    const STATE = {{
        activeModule: '{modules[0]['id'] if modules else ""}',
        activeSubTabs: {{}},
        sidebarOpen: window.innerWidth > 768
    }};

    {chr(10).join([f"STATE.activeSubTabs['{m['id']}'] = '{m['sub_tabs'][0]['id']}';" for m in modules if m.get('sub_tabs')])}

    function switchModule(moduleId) {{
        document.querySelectorAll('.module-panel').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.sb-menu-item').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.sb-subitems').forEach(el => el.classList.remove('open'));

        const modPanel = document.getElementById('mod-' + moduleId);
        if(modPanel) modPanel.classList.add('active');

        const sbMod = document.getElementById('sb-mod-' + moduleId);
        if(sbMod) sbMod.classList.add('active');

        const sbSubs = document.getElementById('sb-subs-' + moduleId);
        if(sbSubs) sbSubs.classList.add('open');

        STATE.activeModule = moduleId;
        window.dispatchEvent(new Event('resize'));
        if(STATE.activeSubTabs[moduleId]) switchSubTab(moduleId, STATE.activeSubTabs[moduleId]);
    }}

    function switchSubTab(moduleId, subTabId) {{
        document.querySelectorAll(`#mod-${{moduleId}} .subtab`).forEach(el => el.classList.remove('active'));
        const btn = document.getElementById(`tab-btn-${{moduleId}}-${{subTabId}}`);
        if(btn) btn.classList.add('active');

        document.querySelectorAll(`#mod-${{moduleId}} .subtab-panel`).forEach(el => el.classList.remove('active'));
        const panel = document.getElementById(`tab-panel-${{moduleId}}-${{subTabId}}`);
        if(panel) panel.classList.add('active');

        document.querySelectorAll(`#sb-subs-${{moduleId}} .sb-sub-item`).forEach(el => el.classList.remove('active'));
        const sbSub = document.getElementById(`sb-sub-${{moduleId}}-${{subTabId}}`);
        if(sbSub) sbSub.classList.add('active');

        STATE.activeSubTabs[moduleId] = subTabId;
        window.dispatchEvent(new Event('resize'));
        document.dispatchEvent(new CustomEvent('tabswitched', {{ detail: {{ moduleId, tabId: subTabId }} }}));
    }}

    function toggleSidebar() {{
        const sb = document.getElementById('sidebar');
        if(window.innerWidth <= 768) sb.classList.toggle('open-mobile');
        else sb.classList.toggle('collapsed');
    }}

    window.addEventListener('resize', () => {{
        const sb = document.getElementById('sidebar');
        if(window.innerWidth <= 768) sb.classList.remove('collapsed');
        else sb.classList.remove('open-mobile');
    }});
    """

    full_html = f"""<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{title}</title>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>{all_css}</style>
</head>
<body>
  <div id="master-header">
      <div style="display:flex;align-items:center;"><span class="hdr-badge">DJP</span><h1>{title}</h1></div>
      <div class="hdr-right"><div style="color:var(--djp-gold);font-weight:bold;font-size:12px;margin-bottom:2px;">&#x1F512; RAHASIA</div><div>Generated: {generated_at}</div></div>
  </div>
  <div id="app-body">
    <nav id="sidebar">{sidebar_html}<button id="sb-toggle" onclick="toggleSidebar()">☰</button></nav>
    <main id="content">{modules_html}</main>
  </div>
  <script>{master_js}\n{all_js}</script>
</body>
</html>
"""

    if encode_b64:
        encoded = base64.b64encode(full_html.encode('utf-8')).decode('ascii')
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
var h=decodeURIComponent(escape(atob(d)));
document.open();document.write(h);document.close();
}})();
</script></body></html>'''
        return wrapper

    return full_html
