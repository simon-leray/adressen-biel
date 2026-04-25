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
    max-width: 850px;
}

.stTextInput > div > div > input {
    border-radius: 12px;
    padding: 1rem 1.5rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    color: #111111 !important;
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
.stRadio [data-testid="stWidgetLabel"] { display: none; }
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
.info-text, .value-text, div[data-testid="stExpander"] strong { 
    color: #F5F5F7 !important; 
}

.stTextInput > div > div > input { background-color: #1C1C1E !important; border-color: #333336 !important; color: #F5F5F7 !important; }
.fact-card { background-color: #1C1C1E !important; border-color: #333336 !important; }
.methodology-box { background-color: #1C1C1E !important; color: #A1A1A6 !important; }
hr { border-top-color: #333336 !important; }
</style>
"""

st.markdown(base_css + (dark_css if dark_mode else ""), unsafe_allow_html=True)

# --- 4. DATEN-LOGIK & DETAILLIERTE BESCHREIBUNGEN ---
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
            return f"<strong>BAURECHT</strong><br><br>Sowohl der Grund und Boden als auch das Gebäude gehören <strong>{txt_boden}</strong>. Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke, die unabhängig voneinander behandelt werden."
        
        return f"<strong>BAURECHT</strong><br><br>Der Grund und Boden gehört <strong>{txt_boden}</strong>. Jedoch besitzt <strong>{txt_bau_nom}</strong> hier ein Baurecht. Das Gebäude gehört rechtlich <strong>{txt_bau_dat}</strong>, obwohl der Boden <strong>{txt_boden}</strong> gehört."

    if len(boden_besitzer) > 1:
        txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
        return f"<strong>GRENZFALL</strong><br><br>Dieses Gebäude steht auf mehreren Grundstücken gleichzeitig. Der gesamte Boden gehört <strong>{txt_boden}</strong>."

    txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl der Boden als auch das Gebäude gehören vollumfänglich <strong>{txt_boden}</strong>."

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
            st.markdown("<div class='label-text'>Eigentumstyp filtern</div>", unsafe_allow_html=True)
            
            filter_options = ["Alle Adressen", "Stadt: Vollbesitz", "Stadt: Boden (Baurecht abgegeben)", "Stadt: Gebäude (Baurecht übernommen)"]
            f_mode = st.radio("Filter", filter_options, horizontal=True, label_visibility="collapsed")
            
            # --- NEU: SICHTBARE FILTER-ERKLÄRUNGEN ---
            if f_mode == "Alle Adressen":
                st.markdown("<p style='color:#888888; font-size:0.9rem; margin-top:-10px; margin-bottom:20px;'>Zeigt das gesamte Immobilienregister unformatiert an.</p>", unsafe_allow_html=True)
            elif f_mode == "Stadt: Vollbesitz":
                st.markdown("<p style='color:#888888; font-size:0.9rem; margin-top:-10px; margin-bottom:20px;'>Zeigt nur Adressen, bei denen sowohl der Boden als auch das Gebäude vollständig der Stadt Biel gehören (keine Fremdnutzung durch Baurecht).</p>", unsafe_allow_html=True)
            elif f_mode == "Stadt: Boden (Baurecht abgegeben)":
                st.markdown("<p style='color:#888888; font-size:0.9rem; margin-top:-10px; margin-bottom:20px;'>Die Stadt Biel besitzt das Grundstück (Boden), hat aber Dritten das Recht eingeräumt, darauf ein eigenes Gebäude zu errichten und zu nutzen.</p>", unsafe_allow_html=True)
            elif f_mode == "Stadt: Gebäude (Baurecht übernommen)":
                st.markdown("<p style='color:#888888; font-size:0.9rem; margin-top:-10px; margin-bottom:20px;'>Der Boden gehört jemand anderem (z.B. Bund, SBB oder Privat), aber die Stadt Biel hat darauf ein eigenes Gebäude im Baurecht errichtet.</p>", unsafe_allow_html=True)

            search = st.text_input("Suche", placeholder="Strasse oder Hausnummer eingeben...", label_visibility="collapsed")
            
            # --- FILTER LOGIK ---
            f_df = df.copy()
            if f_mode == "Stadt: Vollbesitz":
                f_df = df[df['Eigentumsverhältnis'].str.contains("01") & ~df['Grundstücksnummer(n)'].str.contains("Baurecht")]
            elif f_mode == "Stadt: Boden (Baurecht abgegeben)":
                f_df = df[df['Eigentumsverhältnis'].str.contains("01") & df['Grundstücksnummer(n)'].str.contains("Baurecht")]
            elif f_mode == "Stadt: Gebäude (Baurecht übernommen)":
                f_df = df[~df['Eigentumsverhältnis'].str.contains("01") & df['Grundstücksnummer(n)'].str.contains("01") & df['Grundstücksnummer(n)'].str.contains("Baurecht")]

            if search:
                f_df = f_df[f_df['Adresse'].str.contains(search, case=False)]

            # --- PAGINATION LOGIK (KORRIGIERT) ---
            if "load_count" not in st.session_state:
                st.session_state.load_count = 20
            if "last_filter" not in st.session_state:
                st.session_state.last_filter = f_mode
            if "last_search" not in st.session_state:
                st.session_state.last_search = search

            # Reset Counter NUR wenn sich der Filter oder die Suche wirklich geändert hat
            if st.session_state.last_filter != f_mode or st.session_state.last_search != search:
                st.session_state.load_count = 20
                st.session_state.last_filter = f_mode
                st.session_state.last_search = search

            # --- RENDER ADRESSEN ---
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
                
                # "Mehr laden" Button
                if len(f_df) > st.session_state.load_count:
                    if st.button("Weitere Treffer laden..."):
                        st.session_state.load_count += 30
                        st.rerun()
            else:
                st.markdown("<p class='title-subtext' style='margin-top: 2rem;'>Keine Einträge für diese Auswahl gefunden.</p>", unsafe_allow_html=True)

        # --- TAB 2: FACTS & FIGURES ---
        with t2:
            st.write("")
            stadt_df = df[df['Eigentumsverhältnis'].str.contains("01", na=False)]
            baurecht_ab = stadt_df[stadt_df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
            reiner_besitz = stadt_df[~stadt_df['Grundstücksnummer(n)'].str.contains("Baurecht", na=False)]
            
            total_flaeche_stadt = stadt_df['Fläche_Zahl'].sum()
            
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"<div class='fact-card'><span class='label-text'>Gesamtfläche Stadt Biel</span><br><div style='font-size:2.2rem; font-weight:300;'>{int(total_flaeche_stadt):,} m²</div><span style='color:#86868B;'>Entspricht ca. <strong>{int(total_flaeche_stadt/7140)} Fussballfeldern</strong>.</span></div>", unsafe_allow_html=True)
                st.markdown(f"<div class='fact-card'><span class='label-text'>Strategisches Baurecht</span><br><div style='font-size:2.2rem; font-weight:300;'>{int(baurecht_ab['Fläche_Zahl'].sum()):,} m²</div><span style='color:#86868B;'>Von der Stadt Biel an Dritte abgegebene Baurechte.</span></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='fact-card'><span class='label-text'>Areal-Anteil</span><br><div style='font-size:2.2rem; font-weight:300;'>{total_flaeche_stadt/df['Fläche_Zahl'].sum()*100:.1f}%</div><span style='color:#86868B;'>Anteil der Stadt am gesamten erfassten Landbesitz.</span></div>", unsafe_allow_html=True)
                st.markdown(f"<div class='fact-card'><span class='label-text'>Volleigentum</span><br><div style='font-size:2.2rem; font-weight:300;'>{len(reiner_besitz)}</div><span style='color:#86868B;'>Adressen, die sich vollumfänglich im Stadtbesitz befinden.</span></div>", unsafe_allow_html=True)

        st.markdown("<div class='methodology-box'><strong>Methodik:</strong> Daten basieren auf dem WebGIS Biel (26.11.2025). Abgleich mit map.geo.admin.ch über amtliche Grundstücksnummern. Ersetzt kein amtliches Grundbuchdokument.</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"Ein Fehler ist aufgetreten: {e}")
