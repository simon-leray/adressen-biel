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

# --- 1. SEITENKONFIGURATION ---
st.set_page_config(page_title="Immobilienregister Biel", layout="wide")

# Session State Initialisierung
if 'results_limit' not in st.session_state:
    st.session_state.results_limit = 20
if 'disclaimer_shown' not in st.session_state:
    st.session_state.disclaimer_shown = False

# --- 2. CSS DESIGN ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
#MainMenu, footer, header {visibility: hidden;}
[data-testid="stAppViewContainer"] { font-family: 'Inter', sans-serif !important; background-color: #FAFAFA !important; }
.block-container { padding-top: 1rem; padding-bottom: 4rem; max-width: 900px; }
.stTextInput > div > div > input { border-radius: 12px; padding: 1.2rem 1.5rem; font-size: 1.2rem; background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; }
div[data-testid="stExpander"] { border-radius: 12px; margin-bottom: 0.5rem; background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; }
.main-title { text-align: center; font-weight: 700; font-size: 2.8rem; letter-spacing: -0.03em; margin-top: 1rem; color: #111111 !important; }
.title-subtext { text-align: center; color: #888888 !important; margin-bottom: 2rem; font-size: 1.05rem; }
.label-text { font-size: 0.75rem; font-weight: 600; text-transform: uppercase; color: #86868B !important; margin-bottom: 0.2rem; }
.methodology-box { margin-top: 4rem; padding: 2rem; border-radius: 12px; background-color: #F2F2F7; font-size: 0.9rem; color: #555555; }
.stTabs [aria-selected="true"] { color: #111111 !important; border-bottom: 2px solid #111111 !important; }
div[role="radiogroup"] { gap: 8px !important; margin-bottom: 1.5rem; }
div[role="radiogroup"] > label { background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; padding: 8px 16px !important; border-radius: 30px !important; }
div[role="radiogroup"] > label:has(input:checked) { background-color: #111111 !important; border-color: #111111 !important; }
div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; }
.legend-box { padding: 15px; border-radius: 12px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 15px; display: flex; gap: 15px; }
.legend-item { display: flex; align-items: center; font-size: 0.85rem; font-weight: 500; }
.legend-color { width: 14px; height: 14px; border-radius: 4px; margin-right: 8px; }
.maps-link { font-size: 0.85rem; color: #0066CC; text-decoration: none; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

# --- 3. HELFER & DATEN ---

def format_methodology(text):
    text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)
    return text.replace('\n', '<br>')

def get_owner_text(raw_text, case_index):
    code = next((k for k in EIGENTUEMER_MAP if k in str(raw_text)), None)
    return EIGENTUEMER_MAP[code][case_index] if code else ("einem unbekannten Eigentümer" if case_index == 0 else "ein unbekannter Eigentümer")

@st.cache_data(show_spinner=False)
def load_data():
    if not os.path.exists(EXCEL_FILE): return None
    df = pd.read_excel(EXCEL_FILE, sheet_name='Adress-Verzeichnis').fillna("")
    extracted = df['Fläche(n)'].str.extract(r'(\d+)')
    df['Fläche_Zahl'] = extracted[0].astype(float).fillna(0) if not extracted.empty else 0
    
    def bestimme_kategorie(row):
        besitz = str(row['Eigentumsverhältnis'])
        nummern = str(row['Grundstücksnummer(n)'])
        s_boden = "01" in besitz and "Baurecht" not in nummern and "Quellenrecht" not in nummern
        s_bau = "01" in besitz and "Baurecht" in nummern
        if s_boden and not s_bau: return "Vollbesitz"
        if s_boden and s_bau: return "Bodenbesitz"
        if not s_boden and s_bau: return "Gebäudebesitz"
        return "Andere"
    df['Filter_Kategorie'] = df.apply(bestimme_kategorie, axis=1)
    return df

@st.cache_data(show_spinner=False)
def load_geojson_data(df):
    if not os.path.exists(GEOJSON_FILE): return None
    try:
        with open(GEOJSON_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        features = data.get("features", []) if data.get("type") == "FeatureCollection" else []
        if not features:
            for val in data.values():
                if isinstance(val, dict) and val.get("type") == "FeatureCollection":
                    features.extend(val.get("features", []))
        
        mapping = {}
        for _, row in df.iterrows():
            clean_nums = re.findall(r'\d+', str(row['Grundstücksnummer(n)']))
            for n in clean_nums: mapping[str(n)] = row['Filter_Kategorie']

        # Koordinaten-Konvertierung (Simplified)
        def conv(c):
            if isinstance(c[0], (int, float)):
                y_p, x_p = (c[0] - 2600000) / 1000000, (c[1] - 1200000) / 1000000
                return [(2.6779094 + 4.728982 * y_p + 0.791484 * y_p * x_p) * 100 / 36, (16.9023892 + 3.238272 * x_p - 0.270978 * y_p**2) * 100 / 36]
            return [conv(i) for i in c]

        for f in features:
            f['geometry']['coordinates'] = conv(f['geometry']['coordinates'])
            gn = str(f["properties"].get("grst_nummer", "")).strip()
            kat = mapping.get(gn, "Vollbesitz" if str(f["properties"].get("eigentuemer_gr", "")) == "01" else "Andere")
            if kat == "Vollbesitz": f["properties"].update({"fc": [0,122,255,60], "lc": [0,122,255,255], "kn": "Vollbesitz Stadt"})
            elif kat == "Bodenbesitz": f["properties"].update({"fc": [90,200,250,60], "lc": [90,200,250,255], "kn": "Bodenbesitz Stadt"})
            elif kat == "Gebäudebesitz": f["properties"].update({"fc": [255,179,64,60], "lc": [255,179,64,255], "kn": "Gebäudebesitz Stadt"})
            else: f["properties"].update({"fc": [255,149,0,60], "lc": [255,149,0,255], "kn": "Privat / Andere"})
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
    if quelle: return f"<strong>QUELLENRECHT</strong><br><br>Boden: <strong>{get_owner_text(boden[0], 0)}</strong>. Quellenrecht: <strong>{get_owner_text(quelle[0], 1)}</strong>."
    if bau:
        txt_boden = " sowie ".join(list(dict.fromkeys([get_owner_text(b, 0) for b in boden])))
        txt_bau_nom = " sowie ".join(list(dict.fromkeys([get_owner_text(b, 1) for b in bau])))
        return f"<strong>BAURECHT</strong><br><br>Grund: <strong>{txt_boden}</strong>. Baurecht: <strong>{txt_bau_nom}</strong>."
    return f"<strong>EIGENTUM</strong><br><br>Besitz: <strong>{get_owner_text(boden[0], 0)}</strong>."

# --- 4. TEXTE ---
methodik_raw = "**Methodik & Datenquellen:**\nDaten basieren auf WebGIS Biel (26.11.2025). Abgleich mit map.geo.admin.ch. Tool dient der Orientierung, keine verbindliche Auskunft."

# --- 5. RENDER LOGIK ---

if not st.session_state.disclaimer_shown:
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
        search = st.text_input("Suche", placeholder="Strasse eingeben (min. 3 Zeichen)...", label_visibility="collapsed")
        f_mode = st.radio("Filter", ["Alle Adressen", "Vollbesitz der Stadt (Gebäude und Land)", "Bodenbesitz der Stadt (Land im Baurecht abgegeben)", "Gebäudebesitz der Stadt (Land im Baurecht erhalten)"], horizontal=True, label_visibility="collapsed")
        
        # Performance-Filterung
        if (len(search) >= 3) or (f_mode != "Alle Adressen"):
            f_df = df
            if "Vollbesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Vollbesitz"]
            elif "Bodenbesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Bodenbesitz"]
            elif "Gebäudebesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Gebäudebesitz"]
            
            if len(search) >= 3:
                f_df = f_df[f_df['Adresse'].str.contains(search, case=False, na=False)]
            
            if not f_df.empty:
                st.markdown(f"<div style='opacity:0.5; font-size:0.8rem; margin-bottom:10px;'>{len(f_df)} Treffer</div>", unsafe_allow_html=True)
                # Nur anzeigen, was nötig ist
                for _, r in f_df.iloc[:st.session_state.results_limit].iterrows():
                    with st.expander(f"{r['Adresse']}"):
                        st.markdown(f"<div>{build_info_text(r['Eigentumsverhältnis'], r['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
                        m_q = urllib.parse.quote(f"{r['Adresse']}, Biel")
                        st.markdown(f'<a href="https://www.google.com/maps/search/?api=1&query={m_q}" target="_blank" class="maps-link">📍 Google Maps</a>', unsafe_allow_html=True)
                        st.write("---")
                        c1, c2, c3 = st.columns(3)
                        c1.markdown(f"<div class='label-text'>Parzelle</div>{r['Grundstücksnummer(n)']}", unsafe_allow_html=True)
                        c2.markdown(f"<div class='label-text'>Eigentum</div>{re.sub(r'\d{2}:\s*', '', str(r['Eigentumsverhältnis']))}", unsafe_allow_html=True)
                        c3.markdown(f"<div class='label-text'>Fläche</div>{r['Fläche(n)']}", unsafe_allow_html=True)
                
                if len(f_df) > st.session_state.results_limit:
                    if st.button("Mehr laden"):
                        st.session_state.results_limit += 30
                        st.rerun()
            else: st.info("Keine Treffer.")
        else: st.info("Bitte Adresse eingeben (min. 3 Zeichen) oder Filter wählen.")

    with t2:
        st.markdown("<div class='legend-box'><div class='legend-item'><div class='legend-color' style='background-color:#007AFF;'></div> Vollbesitz</div><div class='legend-item'><div class='legend-color' style='background-color:#5AC8FA;'></div> Bodenbesitz</div><div class='legend-item'><div class='legend-color' style='background-color:#FF9500;'></div> Privat</div><div class='legend-item'><div class='legend-color' style='background-color:#FFB340;'></div> Gebäudebesitz</div></div>", unsafe_allow_html=True)
        geo = load_geojson_data(df)
        if geo:
            st.pydeck_chart(pdk.Deck(map_provider="carto", map_style="light", initial_view_state=pdk.ViewState(latitude=47.1368, longitude=7.2468, zoom=14), layers=[pdk.Layer("GeoJsonLayer", data={"type": "FeatureCollection", "features": geo}, stroked=True, filled=True, get_fill_color="properties.fc", get_line_color="properties.lc", line_width_min_pixels=1, pickable=True)], tooltip={"html": "<b>{kn}</b><br>Parzelle: {grst_nummer}"}))

    st.markdown(f"<div class='methodology-box'>{format_methodology(methodik_raw)}</div>", unsafe_allow_html=True)
