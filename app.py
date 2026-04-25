import streamlit as st
import pandas as pd
import os
import plotly.express as px
import re

# --- SEITENKONFIGURATION ---
st.set_page_config(
    page_title="Immobilienregister Biel", 
    layout="wide"
)

# --- THEME LOGIK & TOGGLE ---
col_space, col_toggle = st.columns([6, 1.4])
with col_toggle:
    dark_mode = st.toggle("Dark Mode", value=False)

# --- CSS DESIGN (APPLE MINIMALISM) ---
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
    max-width: 950px;
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

div[data-testid="stExpanderDetails"] {
    background-color: transparent !important;
}

div[data-testid="stExpander"] summary {
    font-weight: 500;
    font-size: 1.1rem;
    padding: 1.2rem;
    color: #111111 !important;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 2rem;
    border-bottom: 1px solid rgba(134, 134, 139, 0.3);
    margin-bottom: 2.5rem;
}

.stTabs [data-baseweb="tab"] {
    height: 50px;
    background-color: transparent;
    font-weight: 500;
    color: #999999 !important;
}

.stTabs [aria-selected="true"] {
    color: #111111 !important;
    border-bottom: 2px solid #111111 !important;
}

div[data-testid="stMetricValue"] {
    font-size: 3.2rem;
    font-weight: 300;
    color: #111111 !important;
}

.main-title {
    text-align: center;
    font-weight: 700;
    font-size: 2.8rem;
    letter-spacing: -0.03em;
    margin-top: 1rem;
    margin-bottom: 0.5rem;
    color: #111111 !important;
}

.title-subtext {
    text-align: center; 
    color: #888888 !important; 
    margin-bottom: 3rem;
    font-size: 1.05rem;
}

.label-text {
    font-size: 0.75rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #86868B !important;
    margin-bottom: 0.5rem;
}

.info-text, .value-text {
    font-size: 0.95rem;
    color: #111111 !important;
}

.methodology-box {
    margin-top: 4rem;
    padding: 2rem;
    border-radius: 12px;
    background-color: #F2F2F7;
    font-size: 0.9rem;
    color: #555555;
    line-height: 1.5;
}
</style>
"""

dark_css = """
<style>
[data-testid="stAppViewContainer"], .stApp {
    background-color: #000000 !important;
}

div[data-testid="stWidgetLabel"] p {
    color: #FFFFFF !important;
}

.stTextInput > div > div > input {
    background-color: #1C1C1E !important;
    border-color: #333336 !important;
    color: #F5F5F7 !important;
}

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

.stTabs [aria-selected="true"] {
    color: #F5F5F7 !important;
    border-bottom-color: #F5F5F7 !important;
}

.methodology-box {
    background-color: #1C1C1E !important;
    color: #A1A1A6 !important;
}

hr { border-top-color: #333336 !important; }
</style>
"""

st.markdown(base_css + (dark_css if dark_mode else ""), unsafe_allow_html=True)

# --- DATEN-LOGIK ---
@st.cache_data
def load_data():
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    df = df.fillna("")
    
    # Fläche extrahieren
    df['Fläche_Zahl'] = df['Fläche(n)'].str.extract(r'(\d+)').astype(float).fillna(0)
    
    # Rechtsform erkennen
    def check_recht(row):
        txt = str(row['Grundstücksnummer(n)'])
        if "Baurecht" in txt: return "Baurecht"
        if "Quellenrecht" in txt: return "Quellenrecht"
        return "Vollbesitz"
    df['Rechtsform'] = df.apply(check_recht, axis=1)
    
    # Kategorie Mapping
    def map_kat(val):
        t = str(val)
        if "01" in t: return "Stadt Biel"
        if "02" in t: return "Öffentl. Institutionen"
        if "03" in t: return "Privat"
        return "Sonstige"
    df['Kategorie'] = df['Eigentumsverhältnis'].apply(map_kat)
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
            return f"<strong>QUELLENRECHT</strong><br><br>Der Grund und Boden dieser Parzelle gehört <strong>{wer_boden}</strong>. Jedoch besitzt <strong>{wer_quelle}</strong> hier ein Quellenrecht."
        return f"Sowohl der Boden als auch das Recht zur Wassernutzung gehören <strong>{name_dativ(quelle_besitzer[0])}</strong>."

    if bau_besitzer:
        txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
        txt_bau_nom = " sowie ".join(list(dict.fromkeys([name_nominativ(b) for b in bau_besitzer])))
        txt_bau_dat = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in bau_besitzer])))
        
        if txt_boden == txt_bau_dat:
            return f"<strong>BAURECHT</strong><br><br>Sowohl der Grund und Boden als auch das Gebäude gehören <strong>{txt_boden}</strong>. Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke."
        
        return f"<strong>BAURECHT</strong><br><br>Der Grund und Boden gehört <strong>{txt_boden}</strong>. Jedoch besitzt <strong>{txt_bau_nom}</strong> hier ein Baurecht. Das Gebäude gehört rechtlich <strong>{txt_bau_dat}</strong>, obwohl der Boden <strong>{txt_boden}</strong> gehört."

    txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl der Boden als auch das Gebäude gehören vollumfänglich <strong>{txt_boden}</strong>."

# --- APP START ---
try:
    df = load_data()

    # Logo & Header
    col_l1, col_logo, col_l3 = st.columns([1, 1.5, 1])
    with col_logo:
        logo_file = "logo_light.png" if dark_mode else "logo_dark.png"
        if os.path.exists(logo_file): st.image(logo_file, use_container_width=True)
            
    st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
    st.markdown("<div class='title-subtext'>Analyse der Bieler Eigentumsverhältnisse auf Basis amtlicher Geodaten.</div>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["Adress-Suche", "Bestandesanalyse"])

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
                        with c1:
                            st.markdown("<div class='label-text'>Grundstück</div>", unsafe_allow_html=True)
                            st.markdown(f"<div class='value-text'>{str(row['Grundstücksnummer(n)']).replace(' / ', '<br>')}</div>", unsafe_allow_html=True)
                        with c2:
                            st.markdown("<div class='label-text'>Eigentum</div>", unsafe_allow_html=True)
                            st.markdown(f"<div class='value-text'>{bereinige_eigentum_text(row['Eigentumsverhältnis']).replace(' / ', '<br>')}</div>", unsafe_allow_html=True)
                        with c3:
                            st.markdown("<div class='label-text'>Fläche</div>", unsafe_allow_html=True)
                            st.markdown(f"<div class='value-text'>{str(row['Fläche(n)']).replace(' / ', '<br>')}</div>", unsafe_allow_html=True)
            else:
                st.markdown("<p class='title-subtext' style='margin-top: 2rem;'>Es wurden keine Einträge gefunden.</p>", unsafe_allow_html=True)

    with tab2:
        st.write("")
        st.markdown("<div class='label-text'>Interaktive Exploration</div>", unsafe_allow_html=True)
        
        # Filter
        f1, f2, f3 = st.columns([2, 1, 1])
        with f1:
            kat_f = st.multiselect("Kategorien", options=df['Kategorie'].unique(), default=df['Kategorie'].unique())
        with f2:
            recht_f = st.multiselect("Rechtsformen", options=df['Rechtsform'].unique(), default=df['Rechtsform'].unique())
        with f3:
            min_flaeche = st.number_input("Mindestfläche (m²)", 0, 100000, 0, step=100)
        
        filtered_df = df[(df['Kategorie'].isin(kat_f)) & (df['Rechtsform'].isin(recht_f)) & (df['Fläche_Zahl'] >= min_flaeche)]
        
        # Metriken
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Gefilterte Objekte", f"{len(filtered_df):,}")
        col_m2.metric("Gesamtfläche", f"{int(filtered_df['Fläche_Zahl'].sum()):,} m²")
        baurecht_quote = (len(filtered_df[filtered_df['Rechtsform'] == 'Baurecht']) / len(filtered_df) * 100) if len(filtered_df) > 0 else 0
        col_m3.metric("Baurechts-Quote", f"{baurecht_quote:.1f}%")
        
        st.write("---")
        
        # Treemap Visualisierung
        st.markdown("<div class='label-text'>Flächenverteilung (Treemap)</div>", unsafe_allow_html=True)
        if not filtered_df.empty:
            fig = px.treemap(filtered_df, 
                             path=['Kategorie', 'Rechtsform', 'Adresse'], 
                             values='Fläche_Zahl',
                             color='Kategorie',
                             color_discrete_sequence=["#000000", "#86868b", "#d2d2d7"] if not dark_mode else ["#FFFFFF", "#86868b", "#444444"])
            fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=500)
            fig.update_traces(textinfo="label+value")
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info("Keine Daten für diese Filterkombination vorhanden.")

    # --- METHODIK ---
    st.markdown(f"""
    <div class='methodology-box'>
        <strong>Methodik & Datenquellen</strong><br>
        Diese Anwendung basiert auf den öffentlichen Geodaten des WebGIS der Stadt Biel (Stand der letzten Aktualisierung: 26.11.2025). 
        Da die zugrundeliegenden Rohdaten der Stadt Biel eine Unterteilung in drei spezifische Eigentümergruppen (Stadt Biel, öffentliche 
        Institutionen und Private) vorgeben, ist eine detailliertere Aufschlüsselung innerhalb dieser Kategorien auf dieser Datenbasis 
        nicht möglich. Um eine höchstmögliche Genauigkeit zu gewährleisten, wurden die Kategorien mit den amtlichen Daten des 
        Bundes-Kartenportals (map.geo.admin.ch) abgeglichen und verifiziert. Die Zuordnung erfolgt über die amtlichen Grundstücksnummern. 
        Dies ersetzt kein amtliches Grundbuchdokument.
    </div>
    """, unsafe_allow_html=True)

except Exception as e:
    st.markdown(f"<p class='title-subtext'>Ein Fehler ist aufgetreten: {e}</p>", unsafe_allow_html=True)
