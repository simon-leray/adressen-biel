import streamlit as st
import pandas as pd
import os
import re
import urllib.parse
import json
import pydeck as pdk
import plotly.express as px

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
.methodology-box { margin-top: 4rem; padding: 2rem; border-radius: 12px; background-color: #F2F2F7; font-size: 0.9rem; color: #555555; line-height: 1.6; }
.stTabs [aria-selected="true"] { color: #111111 !important; border-bottom: 2px solid #111111 !important; }
div[role="radiogroup"] { gap: 8px !important; margin-top: 0.5rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
div[role="radiogroup"] > label { background-color: #FFFFFF !important; border: 1px solid #EAEAEA !important; padding: 8px 16px !important; border-radius: 30px !important; cursor: pointer; }
div[role="radiogroup"] > label:has(input:checked) { background-color: #111111 !important; border-color: #111111 !important; }
div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; }
.legend-box { padding: 15px; border-radius: 12px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 15px; display: flex; gap: 15px; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; font-size: 0.85rem; font-weight: 500; color: #111111; }
.legend-color { width: 14px; height: 14px; border-radius: 4px; margin-right: 8px; border: 1px solid rgba(0,0,0,0.1); }
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

            if kat == "Vollbesitz":
                f["properties"]["line_color"], f["properties"]["fill_color"], f["properties"]["kat_name"] = [0, 122, 255, 255], [0, 122, 255, 60], "Vollbesitz Stadt"
            elif kat == "Bodenbesitz":
                f["properties"]["line_color"], f["properties"]["fill_color"], f["properties"]["kat_name"] = [90, 200, 250, 255], [90, 200, 250, 60], "Bodenbesitz Stadt (Baurecht abgegeben)"
            elif kat == "Andere":
                f["properties"]["line_color"], f["properties"]["fill_color"], f["properties"]["kat_name"] = [255, 149, 0, 255], [255, 149, 0, 60], "Privat / Andere"
            elif kat == "Gebäudebesitz":
                f["properties"]["line_color"], f["properties"]["fill_color"], f["properties"]["kat_name"] = [255, 179, 64, 255], [255, 179, 64, 60], "Gebäudebesitz Stadt (Baurecht erhalten)"
                
        return features
    except Exception as e: 
        return None

def bereinige_eigentum_text(text): return re.sub(r'\d{2}:\s*', '', str(text))

# --- NEU: WIEDERHERGESTELLTER, DETAILLIERTER BESITZ-TEXT ---
def generiere_besitz_text(besitz_string, nummern_string):
    if not besitz_string: return "Keine detaillierten Daten verfügbar."
    
    def name_dativ(text):
        t = str(text).lower()
        if "01" in t: return "der Stadt Biel"
        if "03" in t: return "einer Privatperson oder privaten Firma"
        if "02" in t: return "einer öffentlichen Institution (Bund, Kanton, SBB oder Ähnliche)"
        return "einem unbekannten Eigentümer"

    def name_nominativ(text):
        t = str(text).lower()
        if "01" in t: return "die Stadt Biel"
        if "03" in t: return "eine Privatperson oder private Firma"
        if "02" in t: return "eine öffentliche Institution (Bund, Kanton, SBB oder Ähnliche)"
        return "ein unbekannter Eigentümer"

    besitzer_liste = str(besitz_string).split(" / ")
    info_liste = str(nummern_string).split(" / ")
    
    boden_besitzer, bau_besitzer, quelle_besitzer = [], [], []
    
    for i in range(len(besitzer_liste)):
        b = besitzer_liste[i]
        info = info_liste[i] if i < len(info_liste) else ""
        if "Quellenrecht" in info: quelle_besitzer.append(b)
        elif "Baurecht" in info: bau_besitzer.append(b)
        else: boden_besitzer.append(b)

    if quelle_besitzer:
        wer_quelle = name_nominativ(quelle_besitzer[0])
        if boden_besitzer:
            wer_boden = name_dativ(boden_besitzer[0])
            return f"<strong>QUELLENRECHT</strong><br><br>Der Grund und Boden dieser Parzelle gehört <strong>{wer_boden}</strong>. Jedoch besitzt <strong>{wer_quelle}</strong> hier ein Quellenrecht. Diese Partei darf auf diesem fremden Grundstück eine Wasserquelle fassen und nutzen."
        return f"<strong>QUELLENRECHT</strong><br><br>Sowohl der Boden als auch das Recht zur Wassernutzung gehören <strong>{name_dativ(quelle_besitzer[0])}</strong>."

    if bau_besitzer:
        txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
        txt_bau_nom = " sowie ".join(list(dict.fromkeys([name_nominativ(b) for b in bau_besitzer])))
        txt_bau_dat = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in bau_besitzer])))

        if txt_boden == txt_bau_dat:
            return f"<strong>BAURECHT</strong><br><br>Sowohl der Grund und Boden als auch das Gebäude gehören <strong>{txt_boden}</strong>. Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke, die im Register unabhängig voneinander behandelt werden."
        
        return f"<strong>BAURECHT</strong><br><br>Der Grund und Boden gehört <strong>{txt_boden}</strong>. Jedoch besitzt <strong>{txt_bau_nom}</strong> hier ein Baurecht. Das Gebäude gehört somit rechtlich <strong>{txt_bau_dat}</strong>, obwohl der Boden weiterhin <strong>{txt_boden}</strong> gehört."

    if len(boden_besitzer) > 1:
        txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
        return f"<strong>GRENZFALL / MITBESITZ / STOCKWERKEIGENTUM</strong><br><br>Dieses Objekt steht auf mehreren Parzellen oder gehört <strong>{txt_boden}</strong> gemeinsam. Dies ist zum Beispiel bei Stockwerkeigentum der Fall, wo verschiedene Parteien jeweils eigene Gebäudeteile besitzen, sich den Grund und Boden aber rechtlich teilen."

    txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl der Grund und Boden als auch das darauf stehende Gebäude gehören vollumfänglich <strong>{txt_boden}</strong>."


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
        
        # --- NEU: DREI TABS ---
        t1, t2, t3 = st.tabs(["🔍 Suche & Recherche", "📊 Facts & Diagramme", "🗺️ Interaktive Karte"])
        
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
            stadt_voll = df[df['Filter_Kategorie'] == "Vollbesitz"]
            stadt_boden = df[df['Filter_Kategorie'] == "Bodenbesitz"]
            total_flaeche_boden = stadt_voll['Fläche_Zahl'].sum() + stadt_boden['Fläche_Zahl'].sum()
            baurecht_ab_flaeche = stadt_boden['Fläche_Zahl'].sum()
            
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"<div class='fact-card'><span class='label-text'>Stadtbesitz Total (Boden)</span><br><div style='font-size:2.2rem;'>{int(total_flaeche_boden):,} m²</div><span>ca. {int(total_flaeche_boden/7140)} Fussballfelder.</span></div>", unsafe_allow_html=True)
                st.markdown(f"<div class='fact-card'><span class='label-text'>Strategisches Baurecht</span><br><div style='font-size:2.2rem;'>{int(baurecht_ab_flaeche):,} m²</div><span>Von der Stadt an Dritte abgegebene Baurechtsfläche.</span></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='fact-card'><span class='label-text'>Areal-Anteil</span><br><div style='font-size:2.2rem;'>{total_flaeche_boden/df['Fläche_Zahl'].sum()*100:.1f}%</div><span>am gesamten erfassten Register.</span></div>", unsafe_allow_html=True)
                st.markdown(f"<div class='fact-card'><span class='label-text'>Volleigentum</span><br><div style='font-size:2.2rem;'>{len(stadt_voll)}</div><span>Adressen ohne Fremdnutzung durch Baurecht.</span></div>", unsafe_allow_html=True)

            # --- NEU: INTERAKTIVES DIAGRAMM (PLOTLY) ---
            st.markdown("<div class='label-text' style='margin-top:2rem; text-align:center;'>Verteilung der erfassten Gesamtfläche</div>", unsafe_allow_html=True)
            
            cat_area = df.groupby('Filter_Kategorie')['Fläche_Zahl'].sum().reset_index()
            rename_map = {
                "Vollbesitz": "Vollbesitz (Stadt)",
                "Bodenbesitz": "Bodenbesitz (Baurecht abgegeben)",
                "Gebäudebesitz": "Gebäudebesitz (Baurecht erhalten)",
                "Andere": "Privat / Andere"
            }
            cat_area['Kategorie'] = cat_area['Filter_Kategorie'].map(rename_map)
            
            # Farben passend zur Karte
            color_map = {
                "Vollbesitz (Stadt)": "#007AFF", 
                "Bodenbesitz (Baurecht abgegeben)": "#5AC8FA", 
                "Privat / Andere": "#FF9500", 
                "Gebäudebesitz (Baurecht erhalten)": "#FFB340"
            }
            
            fig = px.pie(cat_area, values='Fläche_Zahl', names='Kategorie', hole=0.45, color='Kategorie', color_discrete_map=color_map)
            fig.update_traces(textposition='inside', textinfo='percent')
            fig.update_layout(
                showlegend=True, 
                legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
                margin=dict(t=20, b=0, l=0, r=0), 
                paper_bgcolor='rgba(0,0,0,0)', 
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color="#F5F5F7" if dark_mode else "#111111", size=13)
            )
            st.plotly_chart(fig, use_container_width=True)

        with t3:
            st.write("")
            st.markdown("""
            <div class='legend-box'>
                <div class='legend-item'><div class='legend-color' style='background-color:rgba(0,122,255,0.3); border-color:#007AFF;'></div> Vollbesitz (Stadt)</div>
                <div class='legend-item'><div class='legend-color' style='background-color:rgba(90,200,250,0.3); border-color:#5AC8FA;'></div> Bodenbesitz (Baurecht abg.)</div>
                <div class='legend-item'><div class='legend-color' style='background-color:rgba(255,149,0,0.3); border-color:#FF9500;'></div> Privat / Andere</div>
                <div class='legend-item'><div class='legend-color' style='background-color:rgba(255,179,64,0.3); border-color:#FFB340;'></div> Gebäudebesitz (Baurecht erh.)</div>
            </div>
            """, unsafe_allow_html=True)
            
            with st.spinner("Lade Karte..."):
                geo = load_geojson_map_data(df)
                if geo:
                    layer = pdk.Layer(
                        "GeoJsonLayer", data={"type": "FeatureCollection", "features": geo}, 
                        opacity=1.0, stroked=True, filled=True, 
                        get_fill_color="properties.fill_color", get_line_color="properties.line_color", 
                        line_width_min_pixels=2, pickable=True
                    )
                    st.pydeck_chart(pdk.Deck(
                        map_provider="carto", map_style="dark" if dark_mode else "light",
                        layers=[layer], initial_view_state=pdk.ViewState(latitude=47.1368, longitude=7.2468, zoom=14.0), 
                        tooltip={"html": "<b>Parzelle:</b> {grst_nummer}<br/><b>Kategorie:</b> {kat_name}", "style": {"backgroundColor": "steelblue", "color": "white"}}
                    ))
                else: st.warning("Kartendaten fehlen.")
                
        # --- NEU: AUSFÜHRLICHE METHODIK ---
        st.markdown("""
        <div class='methodology-box'>
            <strong>Methodik & Datenquellen:</strong><br><br>
            Die diesem Recherche-Tool zugrundeliegenden Daten basieren auf den öffentlich zugänglichen Geodaten des WebGIS der Stadt Biel (Stand: 26.11.2025). 
            Die Kategorisierung der Eigentumsverhältnisse (Aufschlüsselung nach Stadt, Institutionen und Privaten) erfolgt anhand der städtischen Codierungs-Struktur.<br><br>
            Die physischen Adressen und Grundstücksnummern wurden ergänzend mit den offiziellen Datensätzen des Bundes via <em>map.geo.admin.ch</em> abgeglichen, um eine möglichst hohe geografische Präzision zu gewährleisten.<br><br>
            <em><strong>Wichtiger Hinweis:</strong> Dieses Tool dient ausschliesslich der Orientierung und der journalistischen bzw. analytischen Recherche. Es bietet keine rechtsverbindliche Auskunft und ersetzt in keinem Fall ein amtliches Grundbuchdokument. Bei komplexen Grenz- oder Stockwerkeigentums-Fällen können vereinfachte Darstellungen auftreten.</em>
        </div>
        """, unsafe_allow_html=True)

except Exception as e: st.error(f"Fehler: {e}")
