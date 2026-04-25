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
col_space, col_toggle = st.columns([5, 1.5])
with col_toggle:
    dark_mode = st.toggle("Dark Mode", value=False)

# --- 3. CSS DESIGN (MAXIMALE STABILITÄT) ---
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
    max-width: 850px;
}

/* Input Feld */
.stTextInput > div > div > input {
    border-radius: 12px;
    padding: 1rem 1.5rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    color: #111111 !important;
}

/* Expander Light */
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
    text-align: center; color: #888888 !important; margin-bottom: 3rem; font-size: 1.05rem;
}

.label-text {
    font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: #86868B !important; margin-bottom: 0.5rem;
}

.fact-card {
    padding: 2rem; border-radius: 16px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 1rem;
}

.methodology-box {
    margin-top: 4rem; padding: 2rem; border-radius: 12px; background-color: #F2F2F7; font-size: 0.9rem; color: #555555;
}

.stTabs [aria-selected="true"] { color: #111111 !important; border-bottom: 2px solid #111111 !important; }
</style>
"""

dark_css = """
<style>
[data-testid="stAppViewContainer"], .stApp { background-color: #000000 !important; }
div[data-testid="stWidgetLabel"] p { color: #FFFFFF !important; }

/* Expander Dark Mode Fix */
div[data-testid="stExpander"], 
div[data-testid="stExpander"] *, 
div[data-testid="stExpanderDetails"] { 
    background-color: #1C1C1E !important; 
    border-color: #333336 !important; 
}

div[data-testid="stExpander"] summary p, .main-title, div[data-testid="stMetricValue"], 
.info-text, .value-text, div[data-testid="stExpander"] strong { 
    color: #F5F5F7 !important; 
}

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
    if not os.path.exists('Biel_Adressregister_Final.xlsx'):
        return None
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    df = df.fillna("")
    df['Fläche_Zahl'] = df['Fläche(n)'].str.extract(r'(\d+)').astype(float).fillna(0)
    return df

def bereinige_eigentum_text(text):
    return re.sub(r'\d{2}:\s*', '', str(text))

def generiere_besitz_text(besitz_string, nummern_string):
    if not besitz_string: return "Keine Daten verfügbar."
    besitzer_liste = str(besitz_string).split(" / ")
    info_liste = str(nummern_string).split(" / ")
    boden, bau = [], []
    for i in range(len(besitzer_liste)):
        b, info = besitzer_liste[i], info_liste[i] if i < len(info_liste) else ""
        if "Baurecht" in info: bau.append(b)
        else: boden.append(b)

    def d(text):
        t = str(text).lower()
        if "01" in t: return "der Stadt Biel"
        if "03" in t: return "einer Privatperson/Firma"
        return "einer öff. Institution"

    if bau:
        s_boden = any("01" in str(b) for b in boden)
        s_bau = any("01" in str(b) for b in bau)
        if s_boden and not s_bau: return "<strong>BAURECHT (ABGEGEBEN)</strong><br><br>Die Stadt besitzt den Boden, das Gebäude gehört Dritten."
        if s_bau and not s_boden: return "<strong>BAURECHT (ÜBERNOMMEN)</strong><br><br>Die Stadt besitzt das Gebäude auf fremdem Boden."
        return "<strong>BAURECHT</strong><br><br>Komplexes Rechtsverhältnis."
    
    txt = " sowie ".join(list(dict.fromkeys([d(b) for b in boden])))
    return f"<strong>VOLLEIGENTUM</strong><br><br>Gehört vollumfänglich <strong>{txt}</strong>."

# --- 5. APP RENDERING ---
try:
    df = load_data()
    if df is None:
        st.error("Datei 'Biel_Adressregister_Final.xlsx' nicht gefunden!")
    else:
        # Header
        col_l, col_logo, col_r = st.columns([1, 1.5, 1])
        with col_logo:
            logo = "logo_light.png" if dark_mode else "logo_dark.png"
            if os.path.exists(logo): st.image(logo, use_container_width=True)
                
        st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
        st.markdown("<div class='title-subtext'>Immobilienregister Biel | Geodaten Analyse</div>", unsafe_allow_html=True)

        t1, t2 = st.tabs(["🔍 Suche & Filter", "📊 Stadt Biel: Facts"])

        with t1:
            st.write("")
            f_mode = st.radio("Filter", ["Alle", "Stadt: Vollbesitz", "Stadt: Boden (Baurecht abg.)", "Stadt: Gebäude (Baurecht übern.)"], horizontal=True)
            search = st.text_input("Adresse suchen", placeholder="z.B. Südstrasse...", label_visibility="collapsed")
            
            f_df = df.copy()
            if f_mode == "Stadt: Vollbesitz":
                f_df = df[df['Eigentumsverhältnis'].str.contains("01") & ~df['Grundstücksnummer(n)'].str.contains("Baurecht")]
            elif f_mode == "Stadt: Boden (Baurecht abg.)":
                f_df = df[df['Eigentumsverhältnis'].str.contains("01") & df['Grundstücksnummer(n)'].str.contains("Baurecht")]
            elif f_mode == "Stadt: Gebäude (Baurecht übern.)":
                f_df = df[~df['Eigentumsverhältnis'].str.contains("01") & df['Grundstücksnummer(n)'].str.contains("01") & df['Grundstücksnummer(n)'].str.contains("Baurecht")]

            if search:
                f_df = f_df[f_df['Adresse'].str.contains(search, case=False)]

            if not f_df.empty:
                for _, r in f_df.head(50).iterrows(): # Limit auf 50 für Performance
                    with st.expander(f"{r['Adresse']}"):
                        st.markdown(f"<div class='info-text'>{generiere_besitz_text(r['Eigentumsverhältnis'], r['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
                        st.write("---")
                        c1, c2, c3 = st.columns(3)
                        c1.markdown(f"<div class='label-text'>Parzelle</div>{r['Grundstücksnummer(n)']}", unsafe_allow_html=True)
                        c2.markdown(f"<div class='label-text'>Eigentum</div>{bereinige_eigentum_text(r['Eigentumsverhältnis'])}", unsafe_allow_html=True)
                        c3.markdown(f"<div class='label-text'>Fläche</div>{r['Fläche(n)']}", unsafe_allow_html=True)
            else:
                st.info("Keine Treffer.")

        with t2:
            st.write("")
            s_df = df[df['Eigentumsverhältnis'].str.contains("01")]
            total_s = s_df['Fläche_Zahl'].sum()
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"<div class='fact-card'><span class='label-text'>Stadtbesitz</span><br><div style='font-size:2.2rem;'>{int(total_s):,} m²</div><span>ca. {int(total_s/7140)} Fussballfelder.</span></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='fact-card'><span class='label-text'>Areal-Anteil</span><br><div style='font-size:2.2rem;'>{total_s/df['Fläche_Zahl'].sum()*100:.1f}%</div><span>am gesamten Register.</span></div>", unsafe_allow_html=True)

        st.markdown("<div class='methodology-box'><strong>Methodik:</strong> WebGIS Biel (26.11.2025). Abgleich mit map.geo.admin.ch.</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"Ein Fehler ist aufgetreten: {e}")
