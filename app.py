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

# --- 2. THEME LOGIK ---
col_space, col_toggle = st.columns([5, 1.5])
with col_toggle:
    dark_mode = st.toggle("Dark Mode", value=False)

# --- 3. CSS DESIGN (DASHBOARD-STYLE) ---
base_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
#MainMenu, footer, header {visibility: hidden;}
[data-testid="stAppViewContainer"] { font-family: 'Inter', sans-serif !important; background-color: #FAFAFA !important; }
.block-container { padding-top: 1rem; padding-bottom: 4rem; max-width: 900px; }
.stTextInput > div > div > input { border-radius: 12px; padding: 1.2rem 1.5rem; font-size: 1.2rem; background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; color: #111111 !important; box-shadow: 0 8px 30px rgba(0,0,0,0.04); }
div[data-testid="stExpander"] { border-radius: 12px; margin-bottom: 1rem; background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; }
.main-title { text-align: center; font-weight: 700; font-size: 2.8rem; letter-spacing: -0.03em; margin-top: 1rem; color: #111111 !important; }
.title-subtext { text-align: center; color: #888888 !important; margin-bottom: 2rem; font-size: 1.05rem; }
.label-text { font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: #86868B !important; margin-bottom: 0.5rem; }
.fact-card { padding: 2rem; border-radius: 16px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 1rem; }
.methodology-box { margin-top: 4rem; padding: 2rem; border-radius: 12px; background-color: #F2F2F7; font-size: 0.9rem; color: #555555; }
.stTabs [aria-selected="true"] { color: #111111 !important; border-bottom: 2px solid #111111 !important; }
div[role="radiogroup"] { gap: 8px !important; margin-top: 0.5rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
div[role="radiogroup"] > label { background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; padding: 8px 16px !important; border-radius: 30px !important; cursor: pointer; }
div[role="radiogroup"] > label:has(input:checked) { background-color: #111111 !important; border-color: #111111 !important; }
div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; }
.legend-box { padding: 15px; border-radius: 12px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 15px; display: flex; gap: 15px; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; font-size: 0.85rem; font-weight: 500; color: #111111; }
.legend-color { width: 14px; height: 14px; border-radius: 4px; margin-right: 8px; }
.maps-link { font-size: 0.85rem; color: #0066CC; text-decoration: none; font-weight: 500; }
</style>
"""

dark_css = """
<style>
[data-testid="stAppViewContainer"], .stApp { background-color: #000000 !important; }
div[data-testid="stExpander"], div[data-testid="stExpander"] *, div[data-testid="stExpanderDetails"] { background-color: #1C1C1E !important; border-color: #333336 !important; }
div[data-testid="stExpander"] summary p, .main-title, .info-text, .value-text { color: #F5F5F7 !important; }
.stTextInput > div > div > input { background-color: #1C1C1E !important; border-color: #333336 !important; color: #F5F5F7 !important; }
.fact-card, .legend-box { background-color: #1C1C1E !important; border-color: #333336 !important; }
.legend-item { color: #F5F5F7 !important; }
div[role="radiogroup"] > label { background-color: #1C1C1E !important; border-color: #333336 !important; }
div[role="radiogroup"] > label:has(input:checked) { background-color: #FFFFFF !important; border-color: #FFFFFF !important; }
div[role="radiogroup"] > label:has(input:checked) p { color: #111111 !important; }
.methodology-box { background-color: #1C1C1E !important; color: #A1A1A6 !important; }
</style>
"""

st.markdown(base_css + (dark_css if dark_mode else ""), unsafe_allow_html=True)

# --- 4. DATEN-LOGIK & UMRECHNUNG ---

def lv95_to_wgs84(y, x):
    y_prime = (y - 2600000) / 1000000
    x_prime = (x - 1200000) / 1000000
    lon = 2.6779094 + 4.728982 * y_prime + 0.791484 * y_prime * x_prime + 0.1306 * y_prime * x_prime**2 - 0.0436 * y_prime**3
    lat = 16.9023892 + 3.238272 * x_prime - 0.270978 * y_prime**2 - 0.002528 * x_prime**2 - 0.0447 * y_prime**2 * x_prime - 0.0140 * x_prime**3
    return [lon * 100 / 36, lat * 100 / 36]

def recursive_convert(coords):
    if isinstance(coords[0], (int, float)): return lv95_to_wgs84(coords[0], coords[1])
    return [recursive_convert(c) for c in coords]

@st.cache_data
def load_data():
    if not os.path.exists('Biel_Adressregister_Final.xlsx'): return None
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    df = df.fillna("")
    df['Fläche_Zahl'] = df['Fläche(n)'].str.extract(r'(\d+)').astype(float).fillna(0)
    
    def bestimme_kategorie(row):
        besitz = str(row['Eigentumsverhältnis']).split(" / ")
        nummern = str(row['Grundstücksnummer(n)']).split(" / ")
        boden, bau = [], []
        for i in range(len(besitz)):
            b, n = besitz[i], nummern[i] if i < len(nummern) else ""
            if "Baurecht" in n: bau.append(b)
            elif "Quellenrecht" not in n: boden.append(b)
        stadt_boden = any("01" in b for b in boden)
        stadt_bau = any("01" in b for b in bau)
        if stadt_boden and not bau: return "Vollbesitz"
        elif stadt_boden and bau and not stadt_bau: return "Bodenbesitz"
        elif not stadt_boden and stadt_bau: return "Gebäudebesitz"
        elif stadt_boden and stadt_bau: return "Bodenbesitz" if any("01" not in b for b in bau) else "Vollbesitz"
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
                    if isinstance(val, dict) and val.get("type") == "FeatureCollection":
                        features.extend(val.get("features", []))

        mapping = {}
        for _, row in df.iterrows():
            clean_nums = re.findall(r'\d+', str(row['Grundstücksnummer(n)']))
            kat = row['Filter_Kategorie']
            for n in clean_nums:
                mapping[str(n)] = kat
        
        for f in features:
            f['geometry']['coordinates'] = recursive_convert(f['geometry']['coordinates'])
            gn = str(f["properties"].get("grst_nummer", "")).strip()
            
            kat = mapping.get(gn, None)

            if kat is None:
                egr = str(f["properties"].get("eigentuemer_gr", ""))
                if egr == "01": kat = "Vollbesitz"
                else: kat = "Andere"

            # FARBEN-LOGIK AKTUALISIERT NACH DEINEM WUNSCH (ohne Alpha-Kanal für reine Farben)
            if kat == "Vollbesitz":
                f["properties"]["fill_color"] = [0, 122, 255]    # Blau
                f["properties"]["kat_name"] = "Vollbesitz Stadt"
            elif kat == "Bodenbesitz":
                f["properties"]["fill_color"] = [90, 200, 250]   # Hellblau
                f["properties"]["kat_name"] = "Bodenbesitz Stadt (Baurecht abgegeben)"
            elif kat == "Andere":
                f["properties"]["fill_color"] = [255, 149, 0]    # Orange
                f["properties"]["kat_name"] = "Privat / Andere"
            elif kat == "Gebäudebesitz":
                f["properties"]["fill_color"] = [255, 179, 64]   # Hellorange
                f["properties"]["kat_name"] = "Gebäudebesitz Stadt (Baurecht erhalten)"
                
        return features
    except Exception as e: 
        print(f"Fehler: {e}")
        return None

def bereinige_eigentum_text(text): return re.sub(r'\d{2}:\s*', '', str(text))

def generiere_besitz_text(besitz, nummern):
    if not besitz: return "Keine Daten."
    b_list, n_list = str(besitz).split(" / "), str(nummern).split(" / ")
    boden, bau, quelle = [], [], []
    for i in range(len(b_list)):
        b, info = b_list[i], n_list[i] if i < len(n_list) else ""
        if "Quellenrecht" in info: quelle.append(b)
        elif "Baurecht" in info: bau.append(b)
        else: boden.append(b)
    def d(t):
        t = str(t).lower()
        if "01" in t: return "der Stadt Biel"
        if "03" in t: return "einer Privatperson/Firma"
        return "einer öff. Institution"
    if quelle: return f"<strong>QUELLENRECHT</strong><br>Boden gehört {d(boden[0])}, Quellenrecht besteht."
    if bau:
        s_boden, s_bau = any("01" in str(b) for b in boden), any("01" in str(b) for b in bau)
        if s_boden and not s_bau: return f"<strong>BAURECHT (ABGEGEBEN)</strong><br>Der Boden gehört {d(boden[0])}, ist aber im Baurecht abgegeben."
        if s_bau and not s_boden: return f"<strong>GEBÄUDEBESITZ (BAURECHT ERHALTEN)</strong><br>Der Boden gehört Dritten, die Stadt besitzt das Gebäude."
        return "<strong>BAURECHT</strong><br>Komplexes Verhältnis."
    return f"<strong>VOLLEIGENTUM</strong><br>Gehört vollumfänglich <strong>{d(boden[0])}</strong>."

# --- 5. APP RENDERING ---
try:
    df = load_data()
    if df is not None:
        col_logo = st.columns([1, 1.5, 1])[1]
        with col_logo:
            logo = "logo_light.png" if dark_mode else "logo_dark.png"
            if os.path.exists(logo): st.image(logo, use_container_width=True)
        st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
        st.markdown("<div class='title-subtext'>Recherche-Portal für das Immobilienregister Biel</div>", unsafe_allow_html=True)
        t1, t2 = st.tabs(["🔍 Suche & Recherche", "🗺️ Interaktive Areal-Karte"])
        with t1:
            st.write("")
            search = st.text_input("Suche", placeholder="Strasse und Hausnummer (z.B. Ring 16)...", label_visibility="collapsed")
            f_opt = ["Alle Adressen", "Vollbesitz der Stadt (Gebäude und Land)", "Bodenbesitz der Stadt (Land im Baurecht abgegeben)", "Gebäudebesitz der Stadt (Land im Baurecht erhalten)"]
            f_mode = st.radio("Eigentumstyp", f_opt, horizontal=True, label_visibility="collapsed")
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
                            m_q = urllib.parse.quote(f"{r['Adresse']}, Biel")
                            st.markdown(f'<a href="https://www.google.com/maps/search/?api=1&query={m_q}" target="_blank" class="maps-link">📍 Auf Google Maps anzeigen</a>', unsafe_allow_html=True)
                            st.write("---")
                            c1, c2, c3 = st.columns(3)
                            c1.markdown(f"<div class='label-text'>Parzelle</div>{r['Grundstücksnummer(n)']}", unsafe_allow_html=True)
                            c2.markdown(f"<div class='label-text'>Eigentum</div>{bereinige_eigentum_text(r['Eigentumsverhältnis'])}", unsafe_allow_html=True)
                            c3.markdown(f"<div class='label-text'>Fläche</div>{r['Fläche(n)']}", unsafe_allow_html=True)
                    if len(f_df) > st.session_state.lc:
                        if st.button("Weitere laden"): st.session_state.lc += 30; st.rerun()
                else: st.info("Keine Treffer.")
            else: st.info("Bitte Adresse eingeben oder Filter wählen.")
        
        with t2:
            st.write("")
            
            st.markdown("""
            <div class='legend-box'>
                <div class='legend-item'><div class='legend-color' style='background-color:#007AFF;'></div> Vollbesitz (Stadt)</div>
                <div class='legend-item'><div class='legend-color' style='background-color:#5AC8FA;'></div> Bodenbesitz (Baurecht abg.)</div>
                <div class='legend-item'><div class='legend-color' style='background-color:#FF9500;'></div> Privat / Andere</div>
                <div class='legend-item'><div class='legend-color' style='background-color:#FFB340;'></div> Gebäudebesitz (Baurecht erh.)</div>
            </div>
            """, unsafe_allow_html=True)
            
            with st.spinner("Lade Karte..."):
                geo = load_geojson_map_data(df)
                if geo:
                    layer = pdk.Layer(
                        "GeoJsonLayer", 
                        data={"type": "FeatureCollection", "features": geo}, 
                        opacity=0.75, # Opacity etwas runter, damit die Strassennamen von unten durchschimmern 
                        stroked=True, 
                        filled=True, 
                        get_fill_color="properties.fill_color", 
                        get_line_color=[255,255,255,100], 
                        line_width_min_pixels=1, 
                        pickable=True
                    )
                    
                    # Hier setzen wir "carto" ein, damit die Karte darunter s/w ist und die Labels sauber zu sehen sind
                    map_style = "dark" if dark_mode else "light"
                    
                    st.pydeck_chart(pdk.Deck(
                        map_provider="carto",
                        map_style=map_style,
                        layers=[layer], 
                        initial_view_state=pdk.ViewState(latitude=47.1368, longitude=7.2468, zoom=13.5), 
                        tooltip={
                            "html": "<b>Parzelle:</b> {grst_nummer}<br/><b>Kategorie:</b> {kat_name}",
                            "style": {"backgroundColor": "steelblue", "color": "white"}
                        }
                    ))
                else: st.warning("Kartendaten fehlen.")
        st.markdown("<div class='methodology-box'><strong>Methodik:</strong> WebGIS Biel (26.11.2025).</div>", unsafe_allow_html=True)
except Exception as e: st.error(f"Fehler: {e}")
