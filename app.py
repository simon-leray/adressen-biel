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

# --- 3. CSS DESIGN ---
base_css = """
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
    box-shadow: 0 2px 10px rgba(0,0,0,0.02);
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
    margin-top: 4rem; padding: 2rem; border-radius: 12px; background-color: #F2F2F7; font-size: 0.9rem; color: #555555;
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
    box-shadow: 0 2px 5px rgba(0,0,0,0.02);
    cursor: pointer;
    transition: all 0.2s ease;
}

div[role="radiogroup"] > label > div:first-child { display: none !important; }
div[role="radiogroup"] > label p { margin: 0 !important; font-size: 0.85rem !important; font-weight: 500 !important; }

div[role="radiogroup"] > label:has(input:checked) {
    background-color: #111111 !important;
    border-color: #111111 !important;
}
div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; }

.maps-link {
    font-size: 0.85rem; color: #0066CC; text-decoration: none; font-weight: 500;
}
.maps-link:hover { text-decoration: underline; }

/* Legenden Styling */
.legend-box {
    padding: 15px; border-radius: 12px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 15px; display: flex; gap: 15px; flex-wrap: wrap;
}
.legend-item { display: flex; align-items: center; font-size: 0.85rem; font-weight: 500; color: #111111; }
.legend-color { width: 14px; height: 14px; border-radius: 4px; margin-right: 8px; }
</style>
"""

dark_css = """
<style>
[data-testid="stAppViewContainer"], .stApp { background-color: #000000 !important; }
div[data-testid="stWidgetLabel"] p { color: #FFFFFF !important; }

div[data-testid="stExpander"], div[data-testid="stExpander"] *, div[data-testid="stExpanderDetails"] { 
    background-color: #1C1C1E !important; border-color: #333336 !important; 
}
div[data-testid="stExpander"] summary p, .main-title, .info-text, .value-text, div[data-testid="stExpander"] strong { color: #F5F5F7 !important; }

.stTextInput > div > div > input { background-color: #1C1C1E !important; border-color: #333336 !important; color: #F5F5F7 !important; }
.methodology-box { background-color: #1C1C1E !important; color: #A1A1A6 !important; }
hr { border-top-color: #333336 !important; }

div[role="radiogroup"] > label { background-color: #1C1C1E !important; border-color: #333336 !important; }
div[role="radiogroup"] > label p { color: #F5F5F7 !important; }
div[role="radiogroup"] > label:has(input:checked) { background-color: #FFFFFF !important; border-color: #FFFFFF !important; }
div[role="radiogroup"] > label:has(input:checked) p { color: #111111 !important; }

.maps-link { color: #58A6FF; }

.legend-box { background-color: #1C1C1E !important; border-color: #333336 !important; }
.legend-item { color: #F5F5F7 !important; }
</style>
"""

st.markdown(base_css + (dark_css if dark_mode else ""), unsafe_allow_html=True)

# --- 4. DATEN-LOGIK ---

# Mathe-Umrechner von Schweizer Koordinaten (LV95) auf GPS (WGS84)
def lv95_to_wgs84(y, x):
    y_prime = (y - 2600000) / 1000000
    x_prime = (x - 1200000) / 1000000
    lon = 2.6779094 + 4.728982 * y_prime + 0.791484 * y_prime * x_prime + 0.1306 * y_prime * x_prime**2 - 0.0436 * y_prime**3
    lat = 16.9023892 + 3.238272 * x_prime - 0.270978 * y_prime**2 - 0.002528 * x_prime**2 - 0.0447 * y_prime**2 * x_prime - 0.0140 * x_prime**3
    return [lon * 100 / 36, lat * 100 / 36]

def recursive_convert(coords):
    if isinstance(coords[0], (int, float)):
        return lv95_to_wgs84(coords[0], coords[1])
    return [recursive_convert(c) for c in coords]

@st.cache_data
def load_data():
    if not os.path.exists('Biel_Adressregister_Final.xlsx'):
        return None
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    df = df.fillna("")
    
    def bestimme_kategorie(row):
        besitz = str(row['Eigentumsverhältnis']).split(" / ")
        nummern = str(row['Grundstücksnummer(n)']).split(" / ")
        boden, bau = [], []
        
        for i in range(len(besitz)):
            b = besitz[i]
            n = nummern[i] if i < len(nummern) else ""
            if "Baurecht" in n:
                bau.append(b)
            elif "Quellenrecht" not in n:
                boden.append(b)
                
        stadt_boden = any("01" in b for b in boden)
        stadt_bau = any("01" in b for b in bau)
        
        if stadt_boden and not bau: return "Vollbesitz"
        elif stadt_boden and bau and not stadt_bau: return "Bodenbesitz"
        elif not stadt_boden and stadt_bau: return "Gebäudebesitz"
        elif stadt_boden and stadt_bau:
            if any("01" not in b for b in bau): return "Bodenbesitz"
            return "Vollbesitz"
        return "Andere"
        
    df['Filter_Kategorie'] = df.apply(bestimme_kategorie, axis=1)
    return df

@st.cache_data
def load_geojson_map_data(df):
    if not os.path.exists('Eigentum.md'):
        return None
    
    with open('Eigentum.md', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Bereinigung, falls das File Markdown Formatierung wie ```json hat
    content = content.replace("```json", "").replace("```", "").strip()
    
    try:
        data = json.loads(content)
        features = []
        
        # Durchsuche das File nach FeatureCollections
        if "type" in data and data["type"] == "FeatureCollection":
            features = data.get("features", [])
        else:
            for key, val in data.items():
                if isinstance(val, dict) and val.get("type") == "FeatureCollection":
                    features.extend(val.get("features", []))
                    
        # Mapping von Grundstücksnummern zu Kategorien erstellen
        mapping = {}
        for _, row in df.iterrows():
            nums = str(row['Grundstücksnummer(n)']).replace('Baurecht', '').replace('Quellenrecht', '').split('/')
            cat = row['Filter_Kategorie']
            for n in nums:
                n_clean = n.strip()
                if n_clean:
                    mapping[n_clean] = cat
        
        # Koordinaten umrechnen & Farben zuweisen
        for f in features:
            f['geometry']['coordinates'] = recursive_convert(f['geometry']['coordinates'])
            
            gn = str(f["properties"].get("grst_nummer", ""))
            kat = mapping.get(gn, "Andere")
            
            # Farben Zuweisung
            if kat == "Vollbesitz":
                f["properties"]["fill_color"] = [0, 122, 255, 200]  # Apple Blau
                f["properties"]["kat_name"] = "Vollbesitz Stadt"
            elif kat == "Bodenbesitz":
                f["properties"]["fill_color"] = [255, 149, 0, 200]  # Apple Orange
                f["properties"]["kat_name"] = "Bodenbesitz Stadt (Baurecht abgegeben)"
            elif kat == "Gebäudebesitz":
                f["properties"]["fill_color"] = [52, 199, 89, 200]  # Apple Grün
                f["properties"]["kat_name"] = "Gebäudebesitz Stadt (Baurecht erhalten)"
            else:
                f["properties"]["fill_color"] = [142, 142, 147, 80] # Grau transparent
                f["properties"]["kat_name"] = "Privat / Andere"
                
        return features
    except Exception as e:
        print(f"GeoJSON Parsing Fehler: {e}")
        return None

def bereinige_eigentum_text(text):
    return re.sub(r'\d{2}:\s*', '', str(text))

def generiere_besitz_text(besitz_string, nummern_string):
    if not besitz_string: return "Keine detaillierten Daten verfügbar."
    besitzer_liste = str(besitz_string).split(" / ")
    info_liste = str(nummern_string).split(" / ")
    boden, bau, quelle = [], [], []
    
    for i in range(len(besitzer_liste)):
        b, info = besitzer_liste[i], info_liste[i] if i < len(info_liste) else ""
        if "Quellenrecht" in info: quelle.append(b)
        elif "Baurecht" in info: bau.append(b)
        else: boden.append(b)

    def d(text):
        t = str(text).lower()
        if "01" in t: return "der Stadt Biel"
        if "03" in t: return "einer Privatperson oder privaten Firma"
        if "02" in t: return "einer öffentlichen Institution"
        return "einem unbekannten Eigentümer"

    if quelle:
        return f"<strong>QUELLENRECHT</strong><br><br>Der Boden gehört <strong>{d(boden[0])}</strong>. Jedoch besitzt eine andere Partei hier ein Quellenrecht zur Wassernutzung."
    
    if bau:
        s_boden = any("01" in str(b) for b in boden)
        s_bau = any("01" in str(b) for b in bau)
        if s_boden and not s_bau:
            return f"<strong>BAURECHT (ABGEGEBEN)</strong><br><br>Der Boden gehört <strong>{d(boden[0])}</strong>. Die Stadt hat das Land jedoch an Dritte im Baurecht abgegeben. Diese besitzen das Gebäude, während die Stadt die Kontrolle über den Boden behält."
        if s_bau and not s_boden:
            return f"<strong>GEBÄUDEBESITZ (BAURECHT ERHALTEN)</strong><br><br>Der Boden gehört einem Dritten. Die <strong>Stadt Biel</strong> besitzt hier jedoch das Gebäude im Baurecht."
        return f"<strong>BAURECHT</strong><br><br>Hier besteht ein komplexes Baurechtsverhältnis."

    if len(boden) > 1:
        txt = " sowie ".join(list(dict.fromkeys([d(b) for b in boden])))
        return f"<strong>GRENZFALL / MITBESITZ</strong><br><br>Dieses Objekt steht auf mehreren Parzellen oder gehört <strong>{txt}</strong> gemeinsam."

    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl Boden als auch Gebäude gehören vollumfänglich <strong>{d(boden[0])}</strong>."

# --- 5. APP RENDERING ---
try:
    df = load_data()
    if df is not None:
        col_l, col_logo, col_r = st.columns([1, 1.5, 1])
        with col_logo:
            logo = "logo_light.png" if dark_mode else "logo_dark.png"
            if os.path.exists(logo): st.image(logo, use_container_width=True)
                
        st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
        st.markdown("<div class='title-subtext'>Recherche-Portal für das Immobilienregister Biel</div>", unsafe_allow_html=True)

        t1, t2 = st.tabs(["🔍 Suche & Recherche", "🗺️ Interaktive Areal-Karte"])

        with t1:
            st.write("")
            search = st.text_input("Suche", placeholder="Strasse und Hausnummer (z.B. Ring 16)...", label_visibility="collapsed")
            
            filter_options = [
                "Alle Adressen", 
                "Vollbesitz der Stadt (Gebäude und Land)", 
                "Bodenbesitz der Stadt (Land im Baurecht abgegeben)", 
                "Gebäudebesitz der Stadt (Land im Baurecht erhalten)"
            ]
            f_mode = st.radio("Eigentumstyp", filter_options, horizontal=True, label_visibility="collapsed")
            
            if f_mode == "Alle Adressen":
                st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Zeigt das gesamte Immobilienregister. <strong>Bitte Suchbegriff eingeben.</strong></p>", unsafe_allow_html=True)
            elif f_mode == "Vollbesitz der Stadt (Gebäude und Land)":
                st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Adressen, bei denen Boden und Gebäude vollständig der Stadt Biel gehören.</p>", unsafe_allow_html=True)
            elif f_mode == "Bodenbesitz der Stadt (Land im Baurecht abgegeben)":
                st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Die Stadt besitzt das Land, hat es aber an Dritte im Baurecht abgegeben.</p>", unsafe_allow_html=True)
            elif f_mode == "Gebäudebesitz der Stadt (Land im Baurecht erhalten)":
                st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Der Boden gehört jemand anderem, aber die Stadt besitzt darauf ein Gebäude im Baurecht.</p>", unsafe_allow_html=True)

            f_df = df.copy()
            if f_mode == "Vollbesitz der Stadt (Gebäude und Land)":
                f_df = f_df[f_df['Filter_Kategorie'] == "Vollbesitz"]
            elif f_mode == "Bodenbesitz der Stadt (Land im Baurecht abgegeben)":
                f_df = f_df[f_df['Filter_Kategorie'] == "Bodenbesitz"]
            elif f_mode == "Gebäudebesitz der Stadt (Land im Baurecht erhalten)":
                f_df = f_df[f_df['Filter_Kategorie'] == "Gebäudebesitz"]

            if search:
                f_df = f_df[f_df['Adresse'].str.contains(search, case=False, na=False)]

            show_results = True
            if f_mode == "Alle Adressen" and search.strip() == "":
                show_results = False

            if "load_count" not in st.session_state: st.session_state.load_count = 20
            if "last_filter" not in st.session_state: st.session_state.last_filter = f_mode
            if "last_search" not in st.session_state: st.session_state.last_search = search

            if st.session_state.last_filter != f_mode or st.session_state.last_search != search:
                st.session_state.load_count = 20
                st.session_state.last_filter = f_mode
                st.session_state.last_search = search

            if show_results:
                if not f_df.empty:
                    st.markdown(f"<div style='margin-bottom:1rem; opacity:0.6; font-size:0.8rem;'>{len(f_df)} Treffer gefunden</div>", unsafe_allow_html=True)
                    display_df = f_df.iloc[:st.session_state.load_count]
                    
                    for _, r in display_df.iterrows():
                        with st.expander(f"{r['Adresse']}"):
                            st.markdown(f"<div class='info-text'>{generiere_besitz_text(r['Eigentumsverhältnis'], r['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
                            
                            # GOOGLE MAPS LINK
                            maps_query = urllib.parse.quote(f"{r['Adresse']}, Biel")
                            maps_url = f"[https://www.google.com/maps/search/?api=1&query=](https://www.google.com/maps/search/?api=1&query=){maps_query}"
                            st.markdown(f'<a href="{maps_url}" target="_blank" class="maps-link">📍 Auf Google Maps anzeigen</a>', unsafe_allow_html=True)
                            
                            st.write("---")
                            c1, c2, c3 = st.columns(3)
                            c1.markdown(f"<div class='label-text'>Parzelle</div>{r['Grundstücksnummer(n)']}", unsafe_allow_html=True)
                            c2.markdown(f"<div class='label-text'>Eigentum</div>{bereinige_eigentum_text(r['Eigentumsverhältnis'])}", unsafe_allow_html=True)
                            c3.markdown(f"<div class='label-text'>Fläche</div>{r['Fläche(n)']}", unsafe_allow_html=True)
                    
                    if len(f_df) > st.session_state.load_count:
                        if st.button("Weitere Treffer laden..."):
                            st.session_state.load_count += 30
                            st.rerun()
                else:
                    st.markdown("<p class='title-subtext' style='margin-top: 2rem;'>Keine Einträge gefunden.</p>", unsafe_allow_html=True)
            else:
                st.info("Bitte geben Sie eine Adresse ein oder wählen Sie einen Filter aus, um das Register zu durchsuchen.")

        # --- TAB 2: INTERAKTIVE KARTE ---
        with t2:
            st.write("")
            st.markdown("""
            <div class='legend-box'>
                <div class='legend-item'><div class='legend-color' style='background-color:#007AFF;'></div> Vollbesitz (Stadt)</div>
                <div class='legend-item'><div class='legend-color' style='background-color:#FF9500;'></div> Bodenbesitz (Baurecht abg.)</div>
                <div class='legend-item'><div class='legend-color' style='background-color:#34C759;'></div> Gebäudebesitz (Baurecht erh.)</div>
                <div class='legend-item'><div class='legend-color' style='background-color:#8E8E93;'></div> Privat / Andere</div>
            </div>
            """, unsafe_allow_html=True)
            
            with st.spinner("Kartendaten werden geladen..."):
                geo_features = load_geojson_map_data(df)
                
                if geo_features:
                    # PyDeck Karten-Ebene erstellen
                    layer = pdk.Layer(
                        "GeoJsonLayer",
                        data={"type": "FeatureCollection", "features": geo_features},
                        opacity=0.8,
                        stroked=True,
                        filled=True,
                        extruded=False,
                        get_fill_color="properties.fill_color",
                        get_line_color=[255, 255, 255, 100],
                        line_width_min_pixels=1,
                        pickable=True,
                    )
                    
                    # Karten-Ansicht zentriert auf Biel
                    view_state = pdk.ViewState(
                        latitude=47.1368,
                        longitude=7.2468,
                        zoom=13.5,
                        pitch=0,
                    )
                    
                    map_style = "mapbox://styles/mapbox/dark-v10" if dark_mode else "mapbox://styles/mapbox/light-v10"
                    
                    r = pdk.Deck(
                        layers=[layer], 
                        initial_view_state=view_state, 
                        map_style=map_style, 
                        tooltip={"text": "Parzelle {grst_nummer}\nKategorie: {kat_name}"}
                    )
                    
                    st.pydeck_chart(r)
                else:
                    st.warning("Die Datei 'Eigentum.md' wurde nicht gefunden oder konnte nicht gelesen werden. Bitte lade sie ins Repository hoch.")

        st.markdown("<div class='methodology-box'><strong>Methodik:</strong> Daten basieren auf dem WebGIS Biel (26.11.2025). Ersetzt kein amtliches Grundbuchdokument.</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"Ein Fehler ist aufgetreten: {e}")
