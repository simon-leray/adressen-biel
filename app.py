import streamlit as st
import pandas as pd
import os
import re

# --- 1. SEITENKONFIGURATION ---
st.set_page_config(
    page_title="Immobilienregister Biel", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 2. THEME LOGIK & TOGGLE ---
col_space, col_toggle = st.columns([6, 1.4])
with col_toggle:
    dark_mode = st.toggle("Dark Mode", value=False)

# --- 3. CSS DESIGN (APPLE MINIMALISM & DARK MODE FIXES) ---
base_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

#MainMenu, footer, header {visibility: hidden;}

div[data-testid="stWidgetLabel"] p {
    white-space: nowrap !important;
    font-weight: 500;
}

[data-testid="stAppViewContainer"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
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
    box-shadow: 0 4px 20px rgba(0,0,0,0.03);
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
</style>
"""

dark_css = """
<style>
[data-testid="stAppViewContainer"], .stApp { background-color: #000000 !important; }
div[data-testid="stWidgetLabel"] p { color: #FFFFFF !important; }

/* Aggro-Fix für Expander im Dark Mode */
div[data-testid="stExpander"], 
div[data-testid="stExpander"] *, 
div[data-testid="stExpanderDetails"] { 
    background-color: #1C1C1E !important; 
    border-color: #333336 !important; 
}

div[data-testid="stExpander"] summary p, 
.main-title, div[data-testid="stMetricValue"], 
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

# --- 4. DATEN-LOGIK & FALLUNTERSCHEIDUNG ---
@st.cache_data
def load_data():
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    df = df.fillna("")
    # Flächen als numerische Werte extrahieren
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
        b = besitzer_liste[i]
        info = info_liste[i] if i < len(info_liste) else ""
        if "Quellenrecht" in info: quelle_besitzer.append(b)
        elif "Baurecht" in info: bau_besitzer.append(b)
        else: boden_besitzer.append(b)

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

    if quelle_besitzer:
        wer_quelle = name_nominativ(quelle_besitzer[0])
        if boden_besitzer:
            wer_boden = name_dativ(boden_besitzer[0])
            return f"<strong>QUELLENRECHT</strong><br><br>Der Grund und Boden gehört <strong>{wer_boden}</strong>. Jedoch besitzt <strong>{wer_quelle}</strong> hier ein Quellenrecht."
        return f"Sowohl der Boden als auch das Recht zur Wassernutzung gehören <strong>{name_dativ(quelle_besitzer[0])}</strong>."

    if bau_besitzer:
        txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
        txt_bau_nom = " sowie ".join(list(dict.fromkeys([name_nominativ(b) for b in bau_besitzer])))
        txt_bau_dat = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in bau_besitzer])))
        if txt_boden == txt_bau_dat:
            return f"<strong>BAURECHT</strong><br><br>Sowohl der Grund und Boden als auch das Gebäude gehören <strong>{txt_boden}</strong>. Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke."
        return f"<strong>BAURECHT</strong><br><br>Der Grund und Boden gehört <strong>{txt_boden}</strong>. Jedoch besitzt <strong>{txt_bau_nom}</strong> hier ein Baurecht. Das Gebäude gehört rechtlich <strong>{txt_bau_dat}</strong>, obwohl der Boden <strong>{txt_boden}</strong> gehört."

    if len(boden_besitzer) > 1:
        txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
        return f"<strong>GRENZFALL</strong><br><br>Dieses Gebäude steht auf mehreren Grundstücken gleichzeitig. Der gesamte Boden gehört <strong>{txt_boden}</strong>."

    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl der Boden als auch das Gebäude gehören vollumfänglich <strong>{name_dativ(boden_besitzer[0])}</strong>."

# --- 5. APP STRUKTUR & RENDERING ---
try:
    df = load_data()

    # Header & Logo
    col_l1, col_logo, col_l3 = st.columns([1, 1.5, 1])
    with col_logo:
        logo_file = "logo_light.png" if dark_mode else "logo_dark.png"
        if os.path.exists(logo_file): st.image(logo_file, use_container_width=True)
            
    st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
    st.markdown("<div class='title-subtext'>Analyse der Bieler Eigentumsverhältnisse auf Basis amtlicher Geodaten.</div>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["🔍 Adress-Suche", "📊 Stadt Biel: Facts"])

    # --- TAB 1: SUCHE ---
    with tab1:
        st.write("")
        search_query = st.text_input("Suche", placeholder="Strasse und Hausnummer eingeben...", label_visibility="collapsed")
        if search_query:
            results = df[df['Adresse'].str.contains(search_query, case=False, na=False)]
            if not results.empty:
                for _, row in results.iterrows():
                    with st.expander(f"{row['Adresse']}", expanded=True):
                        st.markdown(f"<div class='info-text'>{generiere_besitz_text(row['Eigentumsverhältnis'], row['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
                        st.markdown("<hr style='border-top: 1px solid rgba(134,134,139,0.1); margin: 1.5rem 0;'>", unsafe_allow_html=True)
                        c1, c2, c3 = st.columns(3)
                        c1.markdown(f"<div class='label-text'>Grundstück</div><div class='value-text'>{row['Grundstücksnummer(n)']}</div>", unsafe_allow_html=True)
                        c2.markdown(f"<div class='label-text'>Eigentum</div><div class='value-text'>{bereinige_eigentum_text(row['Eigentumsverhältnis'])}</div>", unsafe_allow_html=True)
                        c3.markdown(f"<div class='label-text'>Fläche</div><div class='value-text'>{row['Fläche(n)']}</div>", unsafe_allow_html=True)
            else:
                st.markdown("<p class='title-subtext' style='margin-top: 2rem;'>Keine Einträge unter dieser Adresse gefunden.</p>", unsafe_allow_html=True)

    # --- TAB 2: FACTS & FIGURES ---
    with tab2:
        st.write("")
        st.markdown("<div class='label-text'>Datenanalyse: Stadteigentum</div>", unsafe_allow_html=True)
        
        stadt_df = df[df['Eigentumsverhältnis'].str.contains("01", na=False)]
        baurecht_stadt_df = stadt_df[stadt_df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
        reines_stadt_df = stadt_df[~stadt_df['Grundstücksnummer(n)'].str.contains("Baurecht|Quellenrecht", na=False)]
        
        total_flaeche_stadt = stadt_df['Fläche_Zahl'].sum()
        baurecht_flaeche = baurecht_stadt_df['Fläche_Zahl'].sum()
        fussballfelder = total_flaeche_stadt / 7140
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"""<div class='fact-card'><span class='label-text'>Gesamtfläche Stadt Biel</span><br><div style='font-size:2.2rem; font-weight:300;'>{int(total_flaeche_stadt):,} m²</div><span style='color:#86868B;'>Entspricht ca. <strong>{int(fussballfelder)} Fussballfeldern</strong>.</span></div>""", unsafe_allow_html=True)
            st.markdown(f"""<div class='fact-card'><span class='label-text'>Strategisches Baurecht</span><br><div style='font-size:2.2rem; font-weight:300;'>{int(baurecht_flaeche):,} m²</div><span style='color:#86868B;'>Bodenbesitz der Stadt, der per Baurecht an Dritte abgegeben wurde.</span></div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""<div class='fact-card'><span class='label-text'>Reines Stadteigentum</span><br><div style='font-size:2.2rem; font-weight:300;'>{len(reines_stadt_df)}</div><span style='color:#86868B;'>Einzeladressen im vollen Besitz der Stadt (ohne Baurechte).</span></div>""", unsafe_allow_html=True)
            st.markdown(f"""<div class='fact-card'><span class='label-text'>Areal-Anteil</span><br><div style='font-size:2.2rem; font-weight:300;'>{total_flaeche_stadt / df['Fläche_Zahl'].sum() * 100:.1f}%</div><span style='color:#86868B;'>Anteil der Stadt Biel an der gesamten erfassten Parzellenfläche.</span></div>""", unsafe_allow_html=True)

        st.write("")
        st.markdown(f"""
        <div style='background-color: {"#1C1C1E" if dark_mode else "#FFFFFF"}; padding: 2rem; border-radius: 16px; border: 1px solid {"#333336" if dark_mode else "#EAEAEA"};'>
            <span class='label-text'>Fazit zur Landpolitik</span><br><br>
            <span class='info-text'>Die Stadt Biel nutzt ihr Eigentum gezielt als strategisches Instrument. Während Infrastrukturflächen 
            vollständig im Besitz verbleiben, ermöglicht das <strong>Baurecht</strong> der Stadt, die Kontrolle über den Boden 
            zu behalten, während private Investoren die bauliche Entwicklung übernehmen.</span>
        </div>
        """, unsafe_allow_html=True)

    # --- 6. METHODIK INFOBOX ---
    st.markdown(f"""
    <div class='methodology-box'>
        <strong>Methodik & Datenquellen</strong><br>
        Diese Anwendung basiert auf den öffentlichen Geodaten des WebGIS der Stadt Biel (Stand: 26.11.2025). 
        Aufgrund der städtischen Datenstruktur ist eine Unterteilung nur in die Kategorien Stadt Biel, öffentliche Institutionen 
        und Private möglich. Die Daten wurden mit map.geo.admin.ch abgeglichen und über amtliche Grundstücksnummern verifiziert. 
        Dies ersetzt kein amtliches Grundbuchdokument.
    </div>
    """, unsafe_allow_html=True)

except Exception as e:
    st.markdown(f"<p class='title-subtext'>Systemfehler: {e}</p>", unsafe_allow_html=True)
