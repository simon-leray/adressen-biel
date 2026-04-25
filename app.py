import streamlit as st
import pandas as pd
import os
import re
import urllib.parse
import json
import pydeck as pdk

# --- KONSTANTEN ---
EXCEL_FILE = 'Biel_Adressregister_Final.xlsx'
GEOJSON_FILE = 'Eigentum.md'
LOGO_FILE = 'logo_dark.png'

EIGENTUEMER_MAP = {
    "01": ("der Stadt Biel", "die Stadt Biel"),
    "02": ("einer öffentlichen Institution", "eine öffentliche Institution"),
    "03": ("einer Privatperson oder privaten Firma", "eine Privatperson oder private Firma"),
}

# --- 1. SEITENKONFIGURATION & STATE ---
st.set_page_config(page_title="Immobilienregister Biel", layout="wide")

if 'results_limit' not in st.session_state:
    st.session_state.results_limit = 20

# --- 2. CSS DESIGN (LIGHT MODE) ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
#MainMenu, footer, header {visibility: hidden;}
[data-testid="stAppViewContainer"] { font-family: 'Inter', sans-serif !important; background-color: #FAFAFA !important; }
.block-container { padding-top: 1rem; padding-bottom: 4rem; max-width: 900px; }
.stTextInput > div > div > input { border-radius: 12px; padding: 1.2rem 1.5rem; font-size: 1.2rem; background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; box-shadow: 0 8px 30px rgba(0,0,0,0.04); }
div[data-testid="stExpander"] { border-radius: 12px; margin-bottom: 1rem; background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; }
.main-title { text-align: center; font-weight: 700; font-size: 2.8rem; letter-spacing: -0.03em; margin-top: 1rem; color: #111111 !important; }
.title-subtext { text-align: center; color: #888888 !important; margin-bottom: 2rem; font-size: 1.05rem; }
.label-text { font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: #86868B !important; margin-bottom: 0.5rem; }
.methodology-box { margin-top: 4rem; padding: 2rem; border-radius: 12px; background-color: #F2F2F7; font-size: 0.9rem; color: #555555; line-height: 1.6; }
.stTabs [aria-selected="true"] { color: #111111 !important; border-bottom: 2px solid #111111 !important; }
.stRadio [data-testid="stWidgetLabel"] { display: none; }
div[role="radiogroup"] { gap: 8px !important; margin-top: 0.5rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
div[role="radiogroup"] > label { background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; padding: 8px 16px !important; border-radius: 30px !important; cursor: pointer; }
div[role="radiogroup"] > label:has(input:checked) { background-color: #111111 !important; border-color: #111111 !important; }
div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; }
.legend-box { padding: 15px; border-radius: 12px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 15px; display: flex; gap: 15px; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; font-size: 0.85rem; font-weight: 500; color: #111111; }
.legend-color { width: 14px; height: 14px; border-radius: 4px; margin-right: 8px; border: 1px solid rgba(0,0,0,0.1); }
.maps-link { font-size: 0.85rem; color: #0066CC; text-decoration: none; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

# --- 3. HELFER-FUNKTIONEN ---

def format_methodology(text):
    """Konvertiert Markdown-Bold sauber zu HTML (Fix von Claude)."""
    text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)
    return text.replace('\n', '<br>')

def get_owner_text(raw_text, case_index):
    """Nutzt Map für korrekte Bezeichnungen (Fix von Claude)."""
    code = next((k for k in EIGENTUEMER_MAP if k in str(raw_text)), None)
    return EIGENTUEMER_MAP[code][case_index] if code else ("einem unbekannten Eigentümer" if case_index == 0 else "ein unbekannter Eigentümer")

def lv95_to_wgs84(y, x):
    y_p, x_p = (y - 2600000) / 1000000, (x - 1200000) / 1000000
    lon = 2.6779094 + 4.728982 * y_p + 0.791484 * y_p * x_p + 0.1306 * y_p * x_p**2 - 0.0436 * y_p**3
    lat = 16.9023892 + 3.238272 * x_p - 0.270978 * y_p**2 - 0.002528 * x_p**2 - 0.0447 * y_p**2 * x_p - 0.0140 * x_p**3
    return [lon * 100 / 36, lat * 100 / 36]

def recursive_convert(coords):
    if isinstance(coords[0], (int, float)): return lv95_to_wgs84(coords[0], coords[1])
    return [recursive_convert(c) for c in coords]

@st.cache_data
def load_data():
    if not os.path.exists(EXCEL_FILE): return None
    df = pd.read_excel(EXCEL_FILE, sheet_name='Adress-Verzeichnis').fillna("")
    extracted = df['Fläche(n)'].str.extract(r'(\d+)')
    df['Fläche_Zahl'] = extracted[0].astype(float).fillna(0) if not extracted.empty else 0
    
    def bestimme_kategorie(row):
        besitz, nummern = str(row['Eigentumsverhältnis']).split(" / "), str(row['Grundstücksnummer(n)']).split(" / ")
        boden, bau = [], []
        for i in range(len(besitz)):
            b, n = besitz[i], nummern[i] if i < len(nummern) else ""
            if "Baurecht" in n: bau.append(b)
            elif "Quellenrecht" not in n: boden.append(b)
        s_boden, s_bau = any("01" in b for b in boden), any("01" in b for b in bau)
        if s_boden and not bau: return "Vollbesitz"
        if s_boden and bau and not s_bau: return "Bodenbesitz"
        if not s_boden and s_bau: return "Gebäudebesitz"
        if s_boden and s_bau: return "Bodenbesitz" if any("01" not in b for b in bau) else "Vollbesitz"
        return "Andere"
    df['Filter_Kategorie'] = df.apply(bestimme_kategorie, axis=1)
    return df

@st.cache_data
def load_geojson_data(df):
    if not os.path.exists(GEOJSON_FILE): return None
    try:
        with open(GEOJSON_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        features = []
        if data.get("type") == "FeatureCollection": features = data.get("features", [])
        else:
            for val in data.values():
                if isinstance(val, dict) and val.get("type") == "FeatureCollection": features.extend(val.get("features", []))
        mapping = {}
        for _, row in df.iterrows():
            clean_nums = re.findall(r'\d+', str(row['Grundstücksnummer(n)']))
            for n in clean_nums: mapping[str(n)] = row['Filter_Kategorie']
        for f in features:
            f['geometry']['coordinates'] = recursive_convert(f['geometry']['coordinates'])
            gn = str(f["properties"].get("grst_nummer", "")).strip()
            kat = mapping.get(gn, "Vollbesitz" if str(f["properties"].get("eigentuemer_gr", "")) == "01" else "Andere")
            if kat == "Vollbesitz": f["properties"].update({"lc": [0,122,255,255], "fc": [0,122,255,60], "kn": "Vollbesitz Stadt"})
            elif kat == "Bodenbesitz": f["properties"].update({"lc": [90,200,250,255], "fc": [90,200,250,60], "kn": "Bodenbesitz Stadt"})
            elif kat == "Gebäudebesitz": f["properties"].update({"lc": [255,179,64,255], "fc": [255,179,64,60], "kn": "Gebäudebesitz Stadt"})
            else: f["properties"].update({"lc": [255,149,0,255], "fc": [255,149,0,60], "kn": "Privat / Andere"})
        return features
    except: return None

def build_info_text(besitz_str, nummern_str):
    if not besitz_str: return "Keine Daten."
    b_list, n_list = str(besitz_str).split(" / "), str(nummern_str).split(" / ")
    boden, bau, quelle = [], [], []
    for i in range(len(b_list)):
        b, info = b_list[i], n_list[i] if i < len(n_list) else ""
        if "Quellenrecht" in info: quelle.append(b)
        elif "Baurecht" in info: bau.append(b)
        else: boden.append(b)
    if quelle:
        return f"<strong>QUELLENRECHT</strong><br><br>Der Boden gehört <strong>{get_owner_text(boden[0], 0)}</strong>. Jedoch besitzt <strong>{get_owner_text(quelle[0], 1)}</strong> hier ein Quellenrecht."
    if bau:
        txt_boden = " sowie ".join(list(dict.fromkeys([get_owner_text(b, 0) for b in boden])))
        txt_bau_nom = " sowie ".join(list(dict.fromkeys([get_owner_text(b, 1) for b in bau])))
        txt_bau_dat = " sowie ".join(list(dict.fromkeys([get_owner_text(b, 0) for b in bau])))
        if txt_boden == txt_bau_dat: return f"<strong>BAURECHT</strong><br><br>Sowohl Grund als auch Gebäude gehören <strong>{txt_boden}</strong>, sind jedoch rechtlich getrennt geführt."
        return f"<strong>BAURECHT</strong><br><br>Der Grund gehört <strong>{txt_boden}</strong>. Jedoch besitzt <strong>{txt_bau_nom}</strong> hier ein Baurecht. Das Gebäude gehört somit rechtlich <strong>{txt_bau_dat}</strong>."
    if len(boden) > 1: return f"<strong>GRENZFALL / MITBESITZ</strong><br><br>Dieses Objekt gehört <strong>{' sowie '.join(list(dict.fromkeys([get_owner_text(b, 0) for b in boden])))}</strong> gemeinsam."
    return f"<strong>VOLLEIGENTUM</strong><br><br>Grund und Gebäude gehören vollumfänglich <strong>{get_owner_text(boden[0], 0)}</strong>."

# --- 4. METHODIK TEXT ---
methodik_raw = """
**Methodik & Datenquellen:**

Die diesem Tool zugrundeliegenden Daten basieren auf den öffentlich zugänglichen Geodaten des WebGIS der Stadt Biel (Stand: 26.11.2025). Die Kategorisierung der Eigentumsverhältnisse (Aufschlüsselung nach Stadt, öffentliche Institutionen und Privaten) erfolgt anhand der städtischen Codierungs-Struktur.

Die physischen Adressen und Grundstücksnummern wurden ergänzend mit den offiziellen Datensätzen des Bundes via [map.geo.admin.ch](https://map.geo.admin.ch) abgeglichen, um eine möglichst hohe geografische Präzision zu gewährleisten.

Dieses Tool dient ausschliesslich der Orientierung. Es bietet keine verbindliche Auskunft. Bei komplexen Grenz- oder Stockwerkeigentums-Fällen können vereinfachte Darstellungen auftreten.
"""

# --- 5. APP RENDERING ---

# Popup-Logik (Optimiert für Speed)
if 'disclaimer_shown' not in st.session_state:
    @st.dialog("Wichtiger Hinweis")
    def show_disclaimer():
        st.markdown(methodik_raw)
        if st.button("Verstanden"):
            st.session_state.disclaimer_shown = True
            st.rerun()
    show_disclaimer()

df = load_data()
if df is not None:
    if os.path.exists(LOGO_FILE): st.columns([1, 1.5, 1])[1].image(LOGO_FILE, use_container_width=True)
    st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
    st.markdown("<div class='title-subtext'>Suchportal für den Immobilienbesitz der Stadt Biel</div>", unsafe_allow_html=True)
    
    t1, t2 = st.tabs(["🔍 Suche & Recherche", "🗺️ Interaktive Karte"])
    
    with t1:
        search = st.text_input("Suche", placeholder="Strasse und Hausnummer...", key="main_search", label_visibility="collapsed")
        f_mode = st.radio("Filter", ["Alle Adressen", "Vollbesitz der Stadt (Gebäude und Land)", "Bodenbesitz der Stadt (Land im Baurecht abgegeben)", "Gebäudebesitz der Stadt (Land im Baurecht erhalten)"], horizontal=True, label_visibility="collapsed")
        
        # Filter-Beschreibungen
        desc = ""
        if "Vollbesitz" in f_mode: desc = "💡 Gebäude und Boden gehören vollständig der Stadt Biel."
        elif "Bodenbesitz" in f_mode: desc = "💡 Die Stadt besitzt das Land, hat es aber an Dritte im Baurecht abgegeben."
        elif "Gebäudebesitz" in f_mode: desc = "💡 Der Boden gehört jemand anderem, aber die Stadt besitzt das Gebäude im Baurecht."
        if desc: st.markdown(f"<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>{desc}</p>", unsafe_allow_html=True)

        if not (f_mode == "Alle Adressen" and search == ""):
            # Schnelle Filterung
            f_df = df
            if "Vollbesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Vollbesitz"]
            elif "Bodenbesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Bodenbesitz"]
            elif "Gebäudebesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Gebäudebesitz"]
            if search: f_df = f_df[f_df['Adresse'].str.contains(search, case=False, na=False)]
            
            if not f_df.empty:
                st.markdown(f"<div style='margin-bottom:1rem; opacity:0.6; font-size:0.8rem;'>{len(f_df)} Treffer</div>", unsafe_allow_html=True)
                for _, r in f_df.iloc[:st.session_state.results_limit].iterrows():
                    with st.expander(f"{r['Adresse']}"):
                        st.markdown(f"<div class='info-text'>{build_info_text(r['Eigentumsverhältnis'], r['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
                        m_q = urllib.parse.quote(f"{r['Adresse']}, Biel")
                        st.markdown(f'<a href="https://www.google.com/maps/search/?api=1&query={m_q}" target="_blank" class="maps-link">📍 Auf Google Maps anzeigen</a>', unsafe_allow_html=True)
                        st.write("---")
                        c1, c2, c3 = st.columns(3)
                        c1.markdown(f"<div class='label-text'>Parzelle</div>{r['Grundstücksnummer(n)']}", unsafe_allow_html=True)
                        c2.markdown(f"<div class='label-text'>Eigentum</div>{re.sub(r'\d{2}:\s*', '', str(r['Eigentumsverhältnis']))}", unsafe_allow_html=True)
                        c3.markdown(f"<div class='label-text'>Fläche</div>{r['Fläche(n)']}", unsafe_allow_html=True)
                if len(f_df) > st.session_state.results_limit:
                    if st.button("Weitere laden"): 
                        st.session_state.results_limit += 30
                        st.rerun()
            else: st.info("Keine Treffer.")
        else: st.info("Bitte Adresse eingeben oder Filter wählen.")

    with t2:
        st.markdown("<div class='legend-box'><div class='legend-item'><div class='legend-color' style='background-color:rgba(0,122,255,0.3);'></div> Vollbesitz (Stadt)</div><div class='legend-item'><div class='legend-color' style='background-color:rgba(90,200,250,0.3);'></div> Bodenbesitz (Baurecht abg.)</div><div class='legend-item'><div class='legend-color' style='background-color:rgba(255,149,0,0.3);'></div> Privat / Andere</div><div class='legend-item'><div class='legend-color' style='background-color:rgba(255,179,64,0.3);'></div> Gebäudebesitz (Baurecht erh.)</div></div>", unsafe_allow_html=True)
        geo = load_geojson_data(df)
        if geo:
            st.pydeck_chart(pdk.Deck(map_provider="carto", map_style="light", initial_view_state=pdk.ViewState(latitude=47.1368, longitude=7.2468, zoom=14.0), layers=[pdk.Layer("GeoJsonLayer", data={"type": "FeatureCollection", "features": geo}, opacity=1.0, stroked=True, filled=True, get_fill_color="properties.fc", get_line_color="properties.lc", line_width_min_pixels=2, pickable=True)], tooltip={"html": "<b>Parzelle:</b> {grst_nummer}<br/><b>Kategorie:</b> {kn}", "style": {"backgroundColor": "steelblue", "color": "white"}}))

    st.markdown(f"<div class='methodology-box'>{format_methodology(methodik_raw)}</div>", unsafe_allow_html=True)
