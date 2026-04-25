import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="Immobilienregister Biel", 
    layout="wide"
)

col_space, col_toggle = st.columns([6, 1.2])
with col_toggle:
    dark_mode = st.toggle("Dark Mode", value=False)

base_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

#MainMenu, footer, header {visibility: hidden;}

div[data-testid="stWidgetLabel"] p {
    white-space: nowrap !important;
    width: auto !important;
    min-width: 100px !important;
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
    font-weight: 400;
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
    border-radius: 0px;
    font-weight: 500;
    font-size: 1rem;
    color: #999999 !important;
}

.stTabs [aria-selected="true"] {
    color: #111111 !important;
    border-bottom: 2px solid #111111 !important;
}

div[data-testid="stMetricValue"] {
    font-size: 3.5rem;
    font-weight: 300;
    letter-spacing: -0.03em;
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

.info-text {
    font-size: 1.05rem;
    line-height: 1.6;
    font-weight: 300;
    color: #111111 !important;
}

.label-text {
    font-size: 0.7rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-bottom: 0.3rem;
    color: #86868B !important;
}

.value-text {
    font-size: 0.95rem;
    font-weight: 400;
    margin-bottom: 1.2rem;
    color: #111111 !important;
}

hr {
    border-top: 1px solid rgba(134, 134, 139, 0.2);
    margin: 1.5rem 0;
}
</style>
"""

dark_css = """
<style>
[data-testid="stAppViewContainer"] {
    background-color: #000000 !important;
}
[data-testid="stHeader"] {
    background-color: transparent !important;
}
.stTextInput > div > div > input {
    background-color: #1C1C1E !important;
    border-color: #333336 !important;
    color: #F5F5F7 !important;
}
div[data-testid="stExpander"] {
    background-color: #1C1C1E !important;
    border-color: #333336 !important;
}
div[data-testid="stExpanderDetails"] {
    background-color: #1C1C1E !important;
}
div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary p {
    color: #F5F5F7 !important;
}
div[data-testid="stExpander"] summary svg {
    color: #F5F5F7 !important;
    fill: #F5F5F7 !important;
}
.stTabs [aria-selected="true"] {
    color: #F5F5F7 !important;
    border-bottom-color: #F5F5F7 !important;
}
div[data-testid="stMetricValue"], .info-text, .value-text, .main-title {
    color: #F5F5F7 !important;
}
hr {
    border-top-color: #333336 !important;
}
</style>
"""

if dark_mode:
    st.markdown(base_css + dark_css, unsafe_allow_html=True)
    logo_file = "logo_light.png"
else:
    st.markdown(base_css, unsafe_allow_html=True)
    logo_file = "logo_dark.png"

@st.cache_data
def load_excel():
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    return df.fillna("")

def generiere_besitz_text(besitz_string, nummern_string):
    if not besitz_string: return "Keine detaillierten Eigentumsdaten verfügbar."
    
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
        txt_bau = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in bau_besitzer])))
        txt_bau_nom = " sowie ".join(list(dict.fromkeys([name_nominativ(b) for b in bau_besitzer])))

        if txt_boden == txt_bau:
            return f"<strong>BAURECHT</strong><br><br>Sowohl der Grund und Boden als auch das Gebäude gehören <strong>{txt_boden}</strong>. Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke, die unabhängig voneinander behandelt werden."
        
        return f"<strong>BAURECHT</strong><br><br>Der Grund und Boden gehört <strong>{txt_boden}</strong>. Jedoch besitzt <strong>{txt_bau_nom}</strong> hier ein Baurecht. Das Gebäude gehört rechtlich <strong>{txt_bau}</strong>, obwohl der Boden <strong>{txt_boden}</strong> gehört."

    if len(boden_besitzer) > 1:
        txt_boden = " sowie ".join(list(dict.fromkeys([name_dativ(b) for b in boden_besitzer])))
        return f"<strong>GRENZFALL</strong><br><br>Dieses Gebäude steht auf mehreren Grundstücken gleichzeitig. Der gesamte Boden gehört <strong>{txt_boden}</strong>."

    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl der Boden als auch das Gebäude gehören vollumfänglich <strong>{name_dativ(boden_besitzer[0])}</strong>."

try:
    df = load_excel()

    st.write("")
    col_l1, col_logo, col_l3 = st.columns([1, 1.5, 1])
    with col_logo:
        if os.path.exists(logo_file):
            st.image(logo_file, use_container_width=True)
            
    st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
    st.markdown("<div class='title-subtext'>Durchsuchen Sie das Immobilienregister auf Basis amtlicher Geodaten.</div>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["Adress-Suche", "Bestandesübersicht"])

    with tab1:
        st.write("")
        search_query = st.text_input("Suche", placeholder="Strasse und Hausnummer eingeben...", label_visibility="collapsed")
        st.write("")
        
        if search_query:
            results = df[df['Adresse'].str.contains(search_query, case=False, na=False)]
            if not results.empty:
                for _, row in results.iterrows():
                    with st.expander(f"{row['Adresse']}", expanded=True):
                        st.markdown(f"<div class='info-text'>{generiere_besitz_text(row['Eigentumsverhältnis'], row['Grundstücksnummer(n)'])}</div>", unsafe_allow_html=True)
                        st.markdown("<hr>", unsafe_allow_html=True)
                        
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.markdown("<div class='label-text'>Grundstücksnummer(n)</div>", unsafe_allow_html=True)
                            st.markdown(f"<div class='value-text'>{str(row['Grundstücksnummer(n)']).replace(' / ', '<br>')}</div>", unsafe_allow_html=True)
                        with c2:
                            st.markdown("<div class='label-text'>Eigentumsverhältnis</div>", unsafe_allow_html=True)
                            st.markdown(f"<div class='value-text'>{str(row['Eigentumsverhältnis']).replace(' / ', '<br>')}</div>", unsafe_allow_html=True)
                        with c3:
                            st.markdown("<div class='label-text'>Fläche(n)</div>", unsafe_allow_html=True)
                            st.markdown(f"<div class='value-text'>{str(row['Fläche(n)']).replace(' / ', '<br>')}</div>", unsafe_allow_html=True)
            else:
                st.markdown("<p class='title-subtext' style='margin-top: 2rem;'>Es wurden keine Einträge gefunden.</p>", unsafe_allow_html=True)

    with tab2:
        st.write("")
        col_m1, col_m2 = st.columns(2)
        stadt_besitz = df[df['Eigentumsverhältnis'].str.contains("01", na=False)]
        privat_besitz = df[df['Eigentumsverhältnis'].str.contains("03", na=False)]
        col_m1.metric("Mit Stadt-Beteiligung", len(stadt_besitz))
        col_m2.metric("In Privatbesitz", len(privat_besitz))
        
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("<div class='label-text' style='margin-bottom: 1rem;'>Top 10 der grössten städtischen Areale</div>", unsafe_allow_html=True)
        
        stadt_besitz_sort = stadt_besitz.copy()
        stadt_besitz_sort['Fläche_Zahl'] = stadt_besitz_sort['Fläche(n)'].str.extract(r'(\d+)').astype(float)
        top_10 = stadt_besitz_sort.sort_values(by='Fläche_Zahl', ascending=False).head(10)
        st.dataframe(top_10[['Adresse', 'Fläche(n)', 'Grundstücksnummer(n)']], use_container_width=True, hide_index=True)

except Exception as e:
    st.markdown("<p class='title-subtext'>Ein Systemfehler ist aufgetreten. Bitte laden Sie die Seite neu.</p>", unsafe_allow_html=True)
