import streamlit as st
import pandas as pd
import os
import re

# --- 1. SEITENKONFIGURATION ---
st.set_page_config(
    page_title="Immobilienregister Biel", 
    layout="wide"
)

# --- 2. THEME LOGIK ---
col_space, col_toggle = st.columns([6, 1.4])
with col_toggle:
    dark_mode = st.toggle("Dark Mode", value=False)

# --- 3. CSS DESIGN (APPLE MINIMALISM) ---
base_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

#MainMenu, footer, header {visibility: hidden;}

div[data-testid="stWidgetLabel"] p {
    white-space: nowrap !important;
    font-weight: 500;
}

[data-testid="stAppViewContainer"] {
    font-family: 'Inter', sans-serif !important;
    background-color: #FAFAFA !important;
    transition: background-color 0.3s ease;
}

.block-container {
    padding-top: 1rem;
    padding-bottom: 4rem;
    max-width: 850px;
}

.stTextInput > div > div > input {
    border-radius: 12px;
    padding: 1rem 1.5rem;
    font-size: 1.1rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    color: #111111 !important;
    box-shadow: 0 4px 20px rgba(0,0,0,0.02);
}

div[data-testid="stExpander"] {
    border-radius: 12px;
    margin-bottom: 1rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    box-shadow: 0 2px 15px rgba(0,0,0,0.02);
}

div[data-testid="stExpanderDetails"] { background-color: transparent !important; }

div[data-testid="stExpander"] summary {
    font-weight: 500; font-size: 1.1rem; padding: 1.2rem; color: #111111 !important;
}

.main-title {
    text-align: center; font-weight: 700; font-size: 2.8rem; letter-spacing: -0.03em; margin-top: 1rem; color: #111111 !important;
}

.title-subtext {
    text-align: center; color: #888888 !important; margin-bottom: 3rem; font-size: 1.05rem;
}

.label-text {
    font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: #86868B !important; margin-bottom: 0.5rem;
}

.info-text, .value-text { font-size: 0.95rem; color: #111111 !important; }

.fact-card {
    padding: 2rem; border-radius: 16px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 1rem;
}

.methodology-box {
    margin-top: 4rem; padding: 2rem; border-radius: 12px; background-color: #F2F2F7; font-size: 0.9rem; color: #555555; line-height: 1.5;
}

.stTabs [aria-selected="true"] { color: #111111 !important; border-bottom: 2px solid #111111 !important; }

/* Filter Buttons Styling */
.stRadio [data-testid="stWidgetLabel"] { display: none; }
div[data-testid="stHeader"] { background-color: transparent !important; }
</style>
"""

dark_css = """
<style>
[data-testid="stAppViewContainer"], .stApp { background-color: #000000 !important; }
div[data-testid="stWidgetLabel"] p { color: #FFFFFF !important; }
div[data-testid="stExpander"], div[data-testid="stExpander"] *, div[data-testid="stExpanderDetails"] { background-color: #1C1C1E !important; border-color: #333336 !important; }
div[data-testid="stExpander"] summary p, .main-title, div[data-testid="stMetricValue"], .info-text, .value-text, div[data-testid="stExpander"] strong { color: #F5F5F7 !important; }
.stTextInput > div > div > input { background-color: #1C1C1E !important; border-color: #333336 !important; color: #F5F5F7 !important; }
.fact-card { background-color: #1C1C1E !important; border-color: #333336 !important; }
.stTabs [aria-selected="true"] { color: #F5F5F7 !important; border-bottom-color: #F5F5F7 !important; }
.methodology-box { background-color: #1C1C1E !important; color: #A1A1A6 !important; }
hr { border-top-color: #333336 !important; }
</style>
"""

st.markdown(base_css + (dark_css if dark_mode else ""), unsafe_allow_html=True)

# --- 4. DATEN-LOGIK ---
@st.cache_data
def load_data():
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    df = df.fillna("")
    df['Fläche_Zahl'] = df['Fläche(n)'].str.extract(r'(\d+)').astype(float).fillna(0)
    return df

def bereinige_eigentum_text(text):
    return re.sub(r'\d{2}:\s*', '', str(text))

def generiere_besitz_text(besitz_string, nummern_string):
    if not besitz_string: return "Keine detaillierten Daten verfügbar."
    besitzer_liste = str(besitz_string).split(" / ")
    info_liste = str(nummern_string).split(" / ")
    boden_besitzer, bau_besitzer, quelle_besitzer = [], [], []
    
    for i in range(len(besitzer_liste)):
        b, info = besitzer_liste[i], info_liste[i] if i < len(info_liste) else ""
        if "Quellenrecht" in info: quelle_besitzer.append(b)
        elif "Baurecht" in info: bau_besitzer.append(b)
        else: boden_besitzer.append(b)

    def d(text):
        t = str(text).lower()
        if "01" in t: return "der Stadt Biel"
        if "03" in t: return "einer Privatperson oder privaten Firma"
        if "02" in t: return "einer öffentlichen Institution"
        return "einem unbekannten Eigentümer"

    if bau_besitzer:
        # Check ob Stadt Bodenbesitzerin ist
        stadt_boden = any("01" in str(b) for b in boden_besitzer)
        stadt_bau = any("01" in str(b) for b in bau_besitzer)
        
        if stadt_boden and not stadt_bau:
            return f"<strong>BAURECHT (ABGEGEBEN)</strong><br><br>Der Grund und Boden gehört **{d(boden_besitzer[0])}**. Jedoch besitzt eine andere Partei hier ein Baurecht. Das Gebäude gehört rechtlich nicht der Stadt."
        if stadt_bau and not stadt_boden:
            return f"<strong>BAURECHT (ÜBERNOMMEN)</strong><br><br>Der Boden gehört einem Dritten. Die **Stadt Biel** besitzt hier jedoch das Baurecht am Gebäude."
        return f"<strong>BAURECHT</strong><br><br>Hier liegt ein komplexes Baurechtsverhältnis vor."

    txt_boden = " sowie ".join(list(dict.fromkeys([d(b) for b in boden_besitzer])))
    return f"<strong>VOLLEIGENTUM</strong><br><br>Boden und Gebäude gehören vollumfänglich <strong>{txt_boden}</strong>."

# --- 5. APP STRUKTUR ---
try:
    df = load_data()

    # Logo & Header
    col_l1, col_logo, col_l3 = st.columns([1, 1.5, 1])
    with col_logo:
        logo_file = "logo_light.png" if dark_mode else "logo_dark.png"
        if os.path.exists(logo_file): st.image(logo_file, use_container_width=True)
            
    st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
    st.markdown("<div class='title-subtext'>Recherche-Tool für Bieler Eigentumsverhältnisse.</div>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["🔍 Adress-Suche & Filter", "📊 Stadt Biel: Facts"])

    with tab1:
        st.write("")
        # --- NEU: FILTER LOGIK ---
        st.markdown("<div class='label-text'>Filter nach Eigentumstyp</div>", unsafe_allow_html=True)
        filter_mode = st.radio(
            "Filter",
            ["Alle Adressen", "Vollbesitz Stadt Biel", "Baurecht abgegeben (Stadt besitzt nur Boden)", "Baurecht übernommen (Stadt besitzt nur Gebäude)"],
            horizontal=True, label_visibility="collapsed"
        )
        
        search_query = st.text_input("Suche nach Strasse oder Hausnummer", placeholder="Strasse eingeben...", label_visibility="collapsed")
        st.write("")
        
        # Daten filtern basierend auf Auswahl
        filtered_df = df.copy()
        
        if filter_mode == "Vollbesitz Stadt Biel":
            filtered_df = df[df['Eigentumsverhältnis'].str.contains("01", na=False) & ~df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
        elif filter_mode == "Baurecht abgegeben (Stadt besitzt nur Boden)":
            # 01 im Eigentum (Boden), aber Baurecht im Grundstücks-Feld
            filtered_df = df[df['Eigentumsverhältnis'].str.contains("01", na=False) & df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
        elif filter_mode == "Baurecht übernommen (Stadt besitzt nur Gebäude)":
            # Baurecht vorhanden, aber Stadt (01) nicht im Haupt-Eigentumstyp
            filtered_df = df[~df['Eigentumsverhältnis'].str.contains("01", na=False) & df['Grundstücksnummer(n)'].str.contains("01") & df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]

        if search_query:
            filtered_df = filtered_df[filtered_df['Adresse'].str.contains(search_query, case=False, na=False)]

        if not filtered_df.empty:
            st.markdown(f"<div style='margin-bottom:1rem; opacity:0.6; font-size:0.8rem;'>{len(filtered_df)} Treffer gefunden</div>", unsafe_allow_html=True)
            for _, row in filtered_df.iterrows():
                with st.expander(f"{row['Adresse']}", expanded=False):
                    st.markdown(f"<div class='info-text'>{generiere_besitz_text(row['Eigentumsverhältnis'], row['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
                    st.write("---")
                    c1, c2, c3 = st.columns(3)
                    c1.markdown(f"<div class='label-text'>Grundstück</div><div class='value-text'>{row['Grundstücksnummer(n)']}</div>", unsafe_allow_html=True)
                    c2.markdown(f"<div class='label-text'>Eigentum</div><div class='value-text'>{bereinige_eigentum_text(row['Eigentumsverhältnis'])}</div>", unsafe_allow_html=True)
                    c3.markdown(f"<div class='label-text'>Fläche</div><div class='value-text'>{row['Fläche(n)']}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<p class='title-subtext' style='margin-top: 2rem;'>Keine Einträge für diese Auswahl gefunden.</p>", unsafe_allow_html=True)

    with tab2:
        # --- FACTS BERECHNUNG ---
        stadt_df = df[df['Eigentumsverhältnis'].str.contains("01", na=False)]
        baurecht_abgegeben = stadt_df[stadt_df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
        reiner_besitz = stadt_df[~stadt_df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
        
        total_flaeche_stadt = stadt_df['Fläche_Zahl'].sum()
        fussballfelder = total_flaeche_stadt / 7140
        
        st.write("")
        st.markdown("<div class='label-text'>Die Macht der Stadt Biel am Boden</div>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"""<div class='fact-card'><span class='label-text'>Landbesitz Total</span><br><div style='font-size:2.2rem; font-weight:300;'>{int(total_flaeche_stadt):,} m²</div><span style='color:#86868B;'>Entspricht ca. <strong>{int(fussballfelder)} Fussballfeldern</strong>.</span></div>""", unsafe_allow_html=True)
            st.markdown(f"""<div class='fact-card'><span class='label-text'>Davon im Baurecht</span><br><div style='font-size:2.2rem; font-weight:300;'>{int(baurecht_abgegeben['Fläche_Zahl'].sum()):,} m²</div><span style='color:#86868B;'>Fläche, die die Stadt Biel Dritten zur Nutzung überlässt.</span></div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""<div class='fact-card'><span class='label-text'>Adressobjekte</span><br><div style='font-size:2.2rem; font-weight:300;'>{len(reiner_besitz)}</div><span style='color:#86868B;'>Einzeladressen im vollen städtischen Eigentum (ohne Baurechte).</span></div>""", unsafe_allow_html=True)
            st.markdown(f"""<div class='fact-card'><span class='label-text'>Areal-Anteil</span><br><div style='font-size:2.2rem; font-weight:300;'>{total_flaeche_stadt / df['Fläche_Zahl'].sum() * 100:.1f}%</div><span style='color:#86868B;'>Anteil der Stadt Biel an der gesamten Parzellenfläche im Register.</span></div>""", unsafe_allow_html=True)

    # --- METHODIK ---
    st.markdown(f"""
    <div class='methodology-box'>
        <strong>Methodik & Datenquellen</strong><br>
        Grundlage: WebGIS Biel (Stand 26.11.2025). Die Kategorisierung erfolgt nach der städtischen Datenstruktur (Stadt, Institutionen, Private). Verifiziert durch map.geo.admin.ch über amtliche Grundstücksnummern. 
    </div>
    """, unsafe_allow_html=True)

except Exception as e:
    st.error(f"Fehler: {e}")
