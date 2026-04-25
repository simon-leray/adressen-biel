import streamlit as st
import pandas as pd
import os
import re
import urllib.parse
import json
import pydeck as pdk

# --- 1. SEITENKONFIGURATION ---
st.set_page_config(
    page_title="Immobilienregister Biel", 
    layout="wide"
)

# --- 2. CSS DESIGN (NUR LIGHT MODE) ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

#MainMenu, footer, header {visibility: hidden;}

[data-testid="stAppViewContainer"] {
    font-family: 'Inter', sans-serif !important;
    background-color: #FAFAFA !important;
}

.block-container {
    padding-top: 1rem;
    padding-bottom: 4rem;
    max-width: 900px;
}

/* Suchfeld Hero-Style */
.stTextInput > div > div > input {
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    font-size: 1.2rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    color: #111111 !important;
    box-shadow: 0 8px 30px rgba(0,0,0,0.04);
}

div[data-testid="stExpander"] {
    border-radius: 12px;
    margin-bottom: 1rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
}

.main-title {
    text-align: center; font-weight: 700; font-size: 2.8rem; letter-spacing: -0.03em; margin-top: 1rem; color: #111111 !important;
}

.title-subtext {
    text-align: center; color: #888888 !important; margin-bottom: 2rem; font-size: 1.05rem;
}

.label-text {
    font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: #86868B !important; margin-bottom: 0.5rem;
}

.methodology-box {
    margin-top: 4rem; padding: 2rem; border-radius: 12px; background-color: #F2F2F7; font-size: 0.9rem; color: #555555; line-height: 1.6;
}

.stTabs [aria-selected="true"] { color: #111111 !important; border-bottom: 2px solid #111111 !important; }

/* Filter Buttons Styling (Chips) */
.stRadio [data-testid="stWidgetLabel"] { display: none; }
div[role="radiogroup"] { 
    gap: 8px !important; 
    margin-top: 0.5rem; 
    margin-bottom: 1.5rem; 
    flex-wrap: wrap; 
}

div[role="radiogroup"] > label {
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    padding: 8px 16px !important;
    border-radius: 30px !important;
    cursor: pointer;
}

div[role="radiogroup"] > label:has(input:checked) {
    background-color: #111111 !important;
    border-color: #111111 !important;
}
div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; }

.legend-box { padding: 15px; border-radius: 12px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 15px; display: flex; gap: 15px; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; font-size: 0.85rem; font-weight: 500; color: #111111; }
.legend-color { width: 14px; height: 14px; border-radius: 4px; margin-right: 8px; border: 1px solid rgba(0,0,0,0.1); }
.maps-link { font-size: 0.85rem; color: #0066CC; text-decoration: none; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

# --- 3. METHODIK TEXT & POPUP LOGIK ---

methodik_text = """
**Methodik & Datenquellen:**

Die diesem Tool zugrundeliegenden Daten basieren auf den öffentlich zugänglichen Geodaten des WebGIS der Stadt Biel (Stand: 26.11.2025). Die Kategorisierung der Eigentumsverhältnisse (Aufschlüsselung nach Stadt, öffentliche Institutionen und Privaten) erfolgt anhand der städtischen Codierungs-Struktur.

Die physischen Adressen und Grundstücksnummern wurden ergänzend mit den offiziellen Datensätzen des Bundes via [map.geo.admin.ch](https://map.geo.admin.ch) abgeglichen, um eine möglichst hohe geografische Präzision zu gewährleisten.

Dieses Tool dient ausschliesslich der Orientierung. Es bietet keine verbindliche Auskunft. Bei komplexen Grenz- oder Stockwerkeigentums-Fällen können vereinfachte Darstellungen auftreten.
"""

# Popup beim ersten Laden anzeigen
if 'disclaimer_shown' not in st.session_state:
    @st.dialog("Wichtiger Hinweis")
    def show_disclaimer():
        st.markdown(methodik_text)
        if st.button("Verstanden"):
            st.session_state.disclaimer_shown = True
            st.rerun()
    show_disclaimer()

# --- 4. DATEN-LOGIK & UMRECHNUNG ---

def lv95_to_wgs84(y, x):
    y_prime, x_prime = (y - 2600000) / 1000000, (x - 1200000) / 1000000
    lon = 2.6779094 + 4.728982 * y_prime + 0.791484 * y_prime * x_prime + 0.1306 * y_prime * x_prime**2 - 0.0436 * y_prime**3
    lat = 16.9023892 + 3.238272 * x_prime - 0.270978 * y_prime**2 - 0.002528 * x_prime**2 - 0.0447 * y_prime**2 * x_prime - 0.0140 * x_prime**3
    return [lon * 100 / 36, lat * 100 / 36]

def recursive_convert(coords):
    if isinstance(coords[0], (int, float)): return lv95_to_wgs84(coords[0], coords[1])
    return [recursive_convert(c) for c in coords]

@st.cache_data
def load_data():
    if not os.path.exists('Biel_Adressregister_Final.xlsx'): return None
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis').fillna("")
    df['Fläche_Zahl'] = df['Fläche(n)'].str.extract(r'(\d+)').astype(float).fillna(0)
    
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
def load_geojson_map_data(df):
    if not os.path.exists('Eigentum.md'): return None
    with open('Eigentum.md', 'r', encoding='utf-8') as f:
        content = f.read().replace("```json", "").replace("```", "").strip()
    try:
        raw_data = json.loads(content)
        features = []
        if isinstance(raw_data, dict):
            if raw_data.get("type") == "FeatureCollection": features = raw_data.get("features", [])
            else:
                for val in raw_data.values():
                    if isinstance(val, dict) and val.get("type") == "FeatureCollection": features.extend(val.get("features", []))
        mapping = {}
        for _, row in df.iterrows():
            clean_nums = re.findall(r'\d+', str(row['Grundstücksnummer(n)']))
            for n in clean_nums: mapping[str(n)] = row['Filter_Kategorie']
        for f in features:
            f['geometry']['coordinates'] = recursive_convert(f['geometry']['coordinates'])
            gn = str(f["properties"].get("grst_nummer", "")).strip()
            kat = mapping.get(gn, "Vollbesitz" if str(f["properties"].get("eigentuemer_gr", "")) == "01" else "Andere")
            if kat == "Vollbesitz": f["properties"]["lc"], f["properties"]["fc"], f["properties"]["kn"] = [0, 122, 255, 255], [0, 122, 255, 60], "Vollbesitz Stadt"
            elif kat == "Bodenbesitz": f["properties"]["lc"], f["properties"]["fc"], f["properties"]["kn"] = [90, 200, 250, 255], [90, 200, 250, 60], "Bodenbesitz Stadt (Baurecht abgegeben)"
            elif kat == "Andere": f["properties"]["lc"], f["properties"]["fc"], f["properties"]["kn"] = [255, 149, 0, 255], [255, 149, 0, 60], "Privat / Andere"
            elif kat == "Gebäudebesitz": f["properties"]["lc"], f["properties"]["fc"], f["properties"]["kn"] = [255, 179, 64, 255], [255, 179, 64, 60], "Gebäudebesitz Stadt (Baurecht erhalten)"
        return features
    except: return None

def generiere_besitz_text(besitz_string, nummern_string):
    if not besitz_string: return "Keine Daten verfügbar."
    def d(t):
        t = str(t).lower()
        if "01" in t: return "der Stadt Biel"
        if "03" in t: return "einer Privatperson oder privaten Firma"
        if "02" in t: return "einer öffentlichen Institution (Bund, Kanton, SBB oder Ähnliche)"
        return "einem unbekannten Eigentümer"
    def n(t):
        t = str(t).lower()
        if "01" in t: return "die Stadt Biel"
        if "03" in t: return "eine Privatperson oder private Firma"
        if "02" in t: return "eine öffentliche Institution (Bund, Kanton, SBB oder Ähnliche)"
        return "ein unbekannter Eigentümer"
    b_list, n_list = str(besitz_string).split(" / "), str(nummern_string).split(" / ")
    boden, bau, quelle = [], [], []
    for i in range(len(b_list)):
        b, info = b_list[i], n_list[i] if i < len(n_list) else ""
        if "Quellenrecht" in info: quelle.append(b)
        elif "Baurecht" in info: bau.append(b)
        else: boden.append(b)
    if quelle:
        return f"<strong>QUELLENRECHT</strong><br><br>Der Grund und Boden dieser Parzelle gehört <strong>{d(boden[0])}</strong>. Jedoch besitzt <strong>{n(quelle[0])}</strong> hier ein Quellenrecht zur Wassernutzung."
    if bau:
        txt_boden = " sowie ".join(list(dict.fromkeys([d(b) for b in boden])))
        txt_bau_nom = " sowie ".join(list(dict.fromkeys([n(b) for b in bau])))
        txt_bau_dat = " sowie ".join(list(dict.fromkeys([d(b) for b in bau])))
        if txt_boden == txt_bau_dat: return f"<strong>BAURECHT</strong><br><br>Sowohl der Grund und Boden als auch das Gebäude gehören <strong>{txt_boden}</strong>. Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke, die im Register unabhängig voneinander behandelt werden."
        return f"<strong>BAURECHT</strong><br><br>Der Grund gehört <strong>{txt_boden}</strong>. Jedoch besitzt <strong>{txt_bau_nom}</strong> hier ein Baurecht. Das Gebäude gehört somit rechtlich <strong>{txt_bau_dat}</strong>, obwohl der Boden weiterhin <strong>{txt_boden}</strong> gehört."
    if len(boden) > 1: return f"<strong>GRENZFALL / MITBESITZ / STOCKWERKEIGENTUM</strong><br><br>Dieses Objekt steht auf mehreren Parzellen oder gehört <strong>{' sowie '.join(list(dict.fromkeys([d(b) for b in boden])))}</strong> gemeinsam. Dies ist z.B. bei Stockwerkeigentum der Fall."
    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl der Grund und Boden als auch das darauf stehende Gebäude gehören vollumfänglich <strong>{d(boden[0])}</strong>."

# --- 5. APP RENDERING ---
try:
    df = load_data()
    if df is not None:
        col_logo = st.columns([1, 1.5, 1])[1]
        if os.path.exists("logo_dark.png"): col_logo.image("logo_dark.png", use_container_width=True)
        
        st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
        st.markdown("<div class='title-subtext'>Recherche-Portal für das Immobilienregister Biel</div>", unsafe_allow_html=True)
        
        t1, t2 = st.tabs(["🔍 Suche & Recherche", "🗺️ Interaktive Karte"])
        with t1:
            search = st.text_input("Suche", placeholder="Strasse und Hausnummer (z.B. Ring 16)...", label_visibility="collapsed")
            f_opt = ["Alle Adressen", "Vollbesitz der Stadt (Gebäude und Land)", "Bodenbesitz der Stadt (Land im Baurecht abgegeben)", "Gebäudebesitz der Stadt (Land im Baurecht erhalten)"]
            f_mode = st.radio("Filter", f_opt, horizontal=True, label_visibility="collapsed")
            
            if f_mode == "Alle Adressen": st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Zeigt das gesamte Register. <strong>Bitte Suchbegriff eingeben.</strong></p>", unsafe_allow_html=True)
            elif "Vollbesitz" in f_mode: st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Adressen, bei denen Boden und Gebäude vollständig der Stadt Biel gehören.</p>", unsafe_allow_html=True)
            elif "Bodenbesitz" in f_mode: st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Die Stadt besitzt das Land, hat es aber an Dritte im Baurecht abgegeben.</p>", unsafe_allow_html=True)
            elif "Gebäudebesitz" in f_mode: st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Der Boden gehört jemand anderem, aber die Stadt besitzt darauf ein Gebäude im Baurecht.</p>", unsafe_allow_html=True)

            f_df = df.copy()
            if "Vollbesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Vollbesitz"]
            elif "Bodenbesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Bodenbesitz"]
            elif "Gebäudebesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Gebäudebesitz"]
            if search: f_df = f_df[f_df['Adresse'].str.contains(search, case=False, na=False)]
            
            if not (f_mode == "Alle Adressen" and search == ""):
                if not f_df.empty:
                    st.markdown(f"<div style='margin-bottom:1rem; opacity:0.6; font-size:0.8rem;'>{len(f_df)} Treffer</div>", unsafe_allow_html=True)
                    if "lc" not in st.session_state: st.session_state.lc = 20
                    for _, r in f_df.iloc[:st.session_state.lc].iterrows():
                        with st.expander(f"{r['Adresse']}"):
                            st.markdown(f"<div class='info-text'>{generiere_besitz_text(r['Eigentumsverhältnis'], r['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
                            st.markdown(f'<a href="https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(f"{r["Adresse"]}, Biel")}" target="_blank" class="maps-link">📍 Auf Google Maps anzeigen</a>', unsafe_allow_html=True)
                            st.write("---")
                            c1, c2, c3 = st.columns(3)
                            c1.markdown(f"<div class='label-text'>Parzelle</div>{r['Grundstücksnummer(n)']}", unsafe_allow_html=True)
                            c2.markdown(f"<div class='label-text'>Eigentum</div>{re.sub(r'\d{2}:\s*', '', str(r['Eigentumsverhältnis']))}", unsafe_allow_html=True)
                            c3.markdown(f"<div class='label-text'>Fläche</div>{r['Fläche(n)']}", unsafe_allow_html=True)
                    if len(f_df) > st.session_state.lc:
                        if st.button("Weitere laden"): st.session_state.lc += 30; st.rerun()
                else: st.info("Keine Treffer.")
            else: st.info("Bitte Adresse eingeben oder Filter wählen.")
        with t2:
            st.markdown("<div class='legend-box'><div class='legend-item'><div class='legend-color' style='background-color:rgba(0,122,255,0.3); border-color:#007AFF;'></div> Vollbesitz (Stadt)</div><div class='legend-item'><div class='legend-color' style='background-color:rgba(90,200,250,0.3); border-color:#5AC8FA;'></div> Bodenbesitz (Baurecht abg.)</div><div class='legend-item'><div class='legend-color' style='background-color:rgba(255,149,0,0.3); border-color:#FF9500;'></div> Privat / Andere</div><div class='legend-item'><div class='legend-color' style='background-color:rgba(255,179,64,0.3); border-color:#FFB340;'></div> Gebäudebesitz (Baurecht erh.)</div></div>", unsafe_allow_html=True)
            with st.spinner("Lade Karte..."):
                geo = load_geojson_map_data(df)
                if geo:
                    st.pydeck_chart(pdk.Deck(map_provider="carto", map_style="light", initial_view_state=pdk.ViewState(latitude=47.1368, longitude=7.2468, zoom=14.0), layers=[pdk.Layer("GeoJsonLayer", data={"type": "FeatureCollection", "features": geo}, opacity=1.0, stroked=True, filled=True, get_fill_color="properties.fc", get_line_color="properties.lc", line_width_min_pixels=2, pickable=True)], tooltip={"html": "<b>Parzelle:</b> {grst_nummer}<br/><b>Kategorie:</b> {kn}", "style": {"backgroundColor": "steelblue", "color": "white"}}))
        
        st.markdown(f"<div class='methodology-box'>{methodik_text.replace('**', '<strong>').replace('\n', '<br>')}</div>", unsafe_allow_html=True)
except Exception as e: st.error(f"Fehler: {e}")
