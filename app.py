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

# --- 3. CSS DESIGN (DASHBOARD-STYLE & CHIPS) ---
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

.fact-card {
    padding: 2rem; border-radius: 16px; background-color: #FFFFFF; border: 1px solid #EAEAEA; margin-bottom: 1rem;
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

/* Aktiver Zustand Filter Light Mode */
div[role="radiogroup"] > label:has(input:checked) {
    background-color: #111111 !important;
    border-color: #111111 !important;
}
div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; }
</style>
"""

dark_css = """
<style>
[data-testid="stAppViewContainer"], .stApp { background-color: #000000 !important; }
div[data-testid="stWidgetLabel"] p { color: #FFFFFF !important; }

div[data-testid="stExpander"], 
div[data-testid="stExpander"] *, 
div[data-testid="stExpanderDetails"] { 
    background-color: #1C1C1E !important; 
    border-color: #333336 !important; 
}

div[data-testid="stExpander"] summary p, .main-title, div[data-testid="stMetricValue"], 
.info-text, .value-text, div[data-testid="stExpander"] strong { color: #F5F5F7 !important; }

.stTextInput > div > div > input { background-color: #1C1C1E !important; border-color: #333336 !important; color: #F5F5F7 !important; }
.fact-card { background-color: #1C1C1E !important; border-color: #333336 !important; }
.methodology-box { background-color: #1C1C1E !important; color: #A1A1A6 !important; }
hr { border-top-color: #333336 !important; }

/* Filter Buttons Dark Mode */
div[role="radiogroup"] > label {
    background-color: #1C1C1E !important;
    border-color: #333336 !important;
}
div[role="radiogroup"] > label p { color: #F5F5F7 !important; }

div[role="radiogroup"] > label:has(input:checked) {
    background-color: #FFFFFF !important;
    border-color: #FFFFFF !important;
}
div[role="radiogroup"] > label:has(input:checked) p { color: #111111 !important; }
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
            return f"<strong>BAURECHT (ABGEGEBEN)</strong><br><br>Der Boden gehört <strong>{d(boden[0])}</strong>. Die Stadt hat das Land jedoch an Dritte im Baurecht abgegeben."
        if s_bau and not s_boden:
            return f"<strong>GEBÄUDEBESITZ (BAURECHT ERHALTEN)</strong><br><br>Der Boden gehört einem Dritten. Die <strong>Stadt Biel</strong> besitzt hier jedoch das Gebäude (oder Teile davon, z.B. im Stockwerkeigentum) im Baurecht."
        return f"<strong>BAURECHT</strong><br><br>Hier besteht ein komplexes Baurechtsverhältnis zwischen mehreren Parteien."

    if len(boden) > 1:
        txt = " sowie ".join(list(dict.fromkeys([d(b) for b in boden])))
        return f"<strong>GRENZFALL / MITBESITZ</strong><br><br>Dieses Objekt steht auf mehreren Parzellen oder gehört <strong>{txt}</strong> gemeinsam."

    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl Boden als auch Gebäude gehören vollumfänglich <strong>{d(boden[0])}</strong>."

# --- 5. APP RENDERING ---
try:
    df = load_data()
    if df is not None:
        # Header
        col_l, col_logo, col_r = st.columns([1, 1.5, 1])
        with col_logo:
            logo = "logo_light.png" if dark_mode else "logo_dark.png"
            if os.path.exists(logo): st.image(logo, use_container_width=True)
                
        st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
        st.markdown("<div class='title-subtext'>Recherche-Portal für das Immobilienregister Biel</div>", unsafe_allow_html=True)

        t1, t2 = st.tabs(["🔍 Suche & Recherche", "📊 Stadt Biel: Facts"])

        with t1:
            st.write("")
            
            # --- DASHBOARD LAYOUT: SUCHFELD GANZ OBEN ---
            search = st.text_input("Suche", placeholder="Strasse und Hausnummer (z.B. Ring 16)...", label_visibility="collapsed")
            
            # --- DARUNTER DIE FILTER-CHIPS ---
            filter_options = [
                "Alle Adressen", 
                "Vollbesitz der Stadt (Gebäude und Land)", 
                "Bodenbesitz der Stadt (Land im Baurecht abgegeben)", 
                "Gebäudebesitz der Stadt (Land im Baurecht erhalten)"
            ]
            f_mode = st.radio("Eigentumstyp", filter_options, horizontal=True, label_visibility="collapsed")
            
            # --- NEUE ERKLÄRTEXTE ---
            if f_mode == "Alle Adressen":
                st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Zeigt das gesamte Immobilienregister. <strong>Bitte Suchbegriff eingeben.</strong></p>", unsafe_allow_html=True)
            elif f_mode == "Vollbesitz der Stadt (Gebäude und Land)":
                st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Adressen, bei denen Boden und Gebäude vollständig der Stadt Biel gehören.</p>", unsafe_allow_html=True)
            elif f_mode == "Bodenbesitz der Stadt (Land im Baurecht abgegeben)":
                st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Die Stadt besitzt das Land, hat es aber an Dritte im Baurecht abgegeben.</p>", unsafe_allow_html=True)
            elif f_mode == "Gebäudebesitz der Stadt (Land im Baurecht erhalten)":
                st.markdown("<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>💡 Der Boden gehört jemand anderem, aber die Stadt besitzt darauf ein Gebäude (oder Teile davon) im Baurecht.</p>", unsafe_allow_html=True)

            # --- BEWÄHRTE FILTER LOGIK ZURÜCKGEHOLT ---
            f_df = df.copy()
            if f_mode == "Vollbesitz der Stadt (Gebäude und Land)":
                f_df = f_df[f_df['Eigentumsverhältnis'].str.contains("01", na=False) & ~f_df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
            elif f_mode == "Bodenbesitz der Stadt (Land im Baurecht abgegeben)":
                f_df = f_df[f_df['Eigentumsverhältnis'].str.contains("01", na=False) & f_df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
            elif f_mode == "Gebäudebesitz der Stadt (Land im Baurecht erhalten)":
                f_df = f_df[~f_df['Eigentumsverhältnis'].str.contains("01", na=False) & f_df['Grundstücksnummer(n)'].str.contains("01", na=False) & f_df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]

            # --- EXAKTE SUCHLOGIK ---
            if search:
                f_df = f_df[f_df['Adresse'].str.contains(search, case=False, na=False)]

            # --- ANZEIGE-LOGIK ---
            show_results = True
            if f_mode == "Alle Adressen" and search.strip() == "":
                show_results = False

            # --- PAGINATION LOGIK ---
            if "load_count" not in st.session_state:
                st.session_state.load_count = 20
            if "last_filter" not in st.session_state:
                st.session_state.last_filter = f_mode
            if "last_search" not in st.session_state:
                st.session_state.last_search = search

            if st.session_state.last_filter != f_mode or st.session_state.last_search != search:
                st.session_state.load_count = 20
                st.session_state.last_filter = f_mode
                st.session_state.last_search = search

            # --- RENDER ADRESSEN ---
            if show_results:
                if not f_df.empty:
                    st.markdown(f"<div style='margin-bottom:1rem; opacity:0.6; font-size:0.8rem;'>{len(f_df)} Treffer gefunden</div>", unsafe_allow_html=True)
                    
                    display_df = f_df.iloc[:st.session_state.load_count]
                    
                    for _, r in display_df.iterrows():
                        with st.expander(f"{r['Adresse']}"):
                            st.markdown(f"<div class='info-text'>{generiere_besitz_text(r['Eigentumsverhältnis'], r['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
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
                    st.markdown("<p class='title-subtext' style='margin-top: 2rem;'>Keine Einträge gefunden. Versuchen Sie einen anderen Suchbegriff.</p>", unsafe_allow_html=True)
            else:
                st.info("Bitte geben Sie eine Adresse ein oder wählen Sie einen Filter aus, um das Register zu durchsuchen.")

        # --- TAB 2: FACTS & FIGURES ---
        with t2:
            st.write("")
            
            # Berechnung für die Statistik basierend auf der GLEICHEN bewährten Logik wie die Filter
            stadt_voll = df[df['Eigentumsverhältnis'].str.contains("01", na=False) & ~df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
            stadt_boden_abgegeben = df[df['Eigentumsverhältnis'].str.contains("01", na=False) & df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
            
            total_flaeche_boden = stadt_voll['Fläche_Zahl'].sum() + stadt_boden_abgegeben['Fläche_Zahl'].sum()
            baurecht_ab_flaeche = stadt_boden_abgegeben['Fläche_Zahl'].sum()
            
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"<div class='fact-card'><span class='label-text'>Stadtbesitz Total (Boden)</span><br><div style='font-size:2.2rem;'>{int(total_flaeche_boden):,} m²</div><span>ca. {int(total_flaeche_boden/7140)} Fussballfelder.</span></div>", unsafe_allow_html=True)
                st.markdown(f"<div class='fact-card'><span class='label-text'>Strategisches Baurecht</span><br><div style='font-size:2.2rem;'>{int(baurecht_ab_flaeche):,} m²</div><span>Von der Stadt an Dritte im Baurecht abgegeben.</span></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='fact-card'><span class='label-text'>Areal-Anteil</span><br><div style='font-size:2.2rem;'>{total_flaeche_boden/df['Fläche_Zahl'].sum()*100:.1f}%</div><span>am gesamten erfassten Register.</span></div>", unsafe_allow_html=True)
                st.markdown(f"<div class='fact-card'><span class='label-text'>Vollbesitz der Stadt</span><br><div style='font-size:2.2rem;'>{len(stadt_voll)}</div><span>Adressen, bei denen Gebäude und Land der Stadt gehören.</span></div>", unsafe_allow_html=True)

        st.markdown("<div class='methodology-box'><strong>Methodik:</strong> Daten basieren auf dem WebGIS Biel (26.11.2025). Abgleich mit map.geo.admin.ch über amtliche Grundstücksnummern. Ersetzt kein amtliches Grundbuchdokument.</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"Ein Fehler ist aufgetreten: {e}")
