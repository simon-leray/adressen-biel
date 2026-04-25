import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Immobilienregister Biel", 
    layout="wide"
)

st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 24px;
    border-bottom: 1px solid #e0e0e0;
}
.stTabs [data-baseweb="tab"] {
    height: 50px;
    white-space: pre-wrap;
    background-color: transparent;
    border-radius: 0px;
    color: #777;
    font-weight: 400;
}
.stTabs [aria-selected="true"] {
    border-bottom: 2px solid #000;
    color: #000;
    font-weight: 600;
}
div[data-testid="stMetricValue"] {
    font-size: 2.5rem;
    font-weight: 300;
    color: #000;
}
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_excel():
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    return df.fillna("")

def generiere_besitz_text(besitz_string, nummern_string):
    if not besitz_string: 
        return "Keine detaillierten Eigentumsdaten verfügbar."
    
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
    
    boden_besitzer = []
    bau_besitzer = []
    quelle_besitzer = []
    
    for i in range(len(besitzer_liste)):
        b = besitzer_liste[i]
        info = info_liste[i] if i < len(info_liste) else ""
        
        if "Quellenrecht" in info:
            quelle_besitzer.append(b)
        elif "Baurecht" in info:
            bau_besitzer.append(b)
        else:
            boden_besitzer.append(b)

    if quelle_besitzer:
        wer_quelle = name_nominativ(quelle_besitzer[0])
        if boden_besitzer:
            wer_boden = name_dativ(boden_besitzer[0])
            return (f"**QUELLENRECHT** — Der Grund und Boden dieser Parzelle gehört **{wer_boden}**. "
                    f"Jedoch besitzt **{wer_quelle}** hier ein Quellenrecht. Diese Partei "
                    f"darf auf diesem fremden Grundstück eine Wasserquelle fassen und nutzen.")
        else:
            return (f"**QUELLENRECHT** — Sowohl der Boden als auch das Recht zur Wassernutzung gehören **{name_dativ(quelle_besitzer[0])}**.")

    if bau_besitzer:
        einzig_boden = list(dict.fromkeys([name_dativ(b) for b in boden_besitzer]))
        einzig_bau = list(dict.fromkeys([name_dativ(b) for b in bau_besitzer]))
        einzig_bau_nom = list(dict.fromkeys([name_nominativ(b) for b in bau_besitzer]))
        
        txt_boden = " sowie ".join(einzig_boden)
        txt_bau = " sowie ".join(einzig_bau)
        txt_bau_nom = " sowie ".join(einzig_bau_nom)

        if txt_boden == txt_bau:
            return (f"**BAURECHT** — Sowohl der Grund und Boden als auch das Gebäude gehören **{txt_boden}**. "
                    f"Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke, die unabhängig voneinander behandelt werden.")
        
        return (f"**BAURECHT** — Der Grund und Boden gehört **{txt_boden}**. "
                f"Jedoch besitzt **{txt_bau_nom}** hier ein Baurecht. Das Gebäude gehört rechtlich **{txt_bau}**, obwohl der Boden **{txt_boden}** gehört.")

    if len(boden_besitzer) > 1:
        einzig_boden = list(dict.fromkeys([name_dativ(b) for b in boden_besitzer]))
        txt_boden = " sowie ".join(einzig_boden)
        return (f"**GRENZFALL** — Dieses Gebäude steht auf mehreren Grundstücken gleichzeitig. "
                f"Der gesamte Boden gehört **{txt_boden}**.")

    return f"**VOLLEIGENTUM** — Sowohl der Boden als auch das Gebäude gehören vollumfänglich **{name_dativ(boden_besitzer[0])}**."

try:
    df = load_excel()
    st.title("Immobilienregister Stadt Biel")
    st.markdown("<p style='font-size: 1.1rem; color: #555;'>Durchsuchen Sie die Eigentumsverhältnisse der Gebäude auf Basis amtlicher Geodaten.</p>", unsafe_allow_html=True)
    st.write("")

    tab1, tab2 = st.tabs(["Adress-Suche", "Bestandesübersicht"])

    with tab1:
        st.write("")
        search_query = st.text_input("Geben Sie eine Adresse ein", placeholder="Strasse und Hausnummer eingeben...", label_visibility="collapsed")
        st.write("")
        
        if search_query:
            results = df[df['Adresse'].str.contains(search_query, case=False, na=False)]
            if not results.empty:
                st.markdown(f"<span style='color: #888; font-size: 0.9rem;'>{len(results)} ERGEBNISSE</span>", unsafe_allow_html=True)
                st.write("")
                for _, row in results.iterrows():
                    with st.expander(f"{row['Adresse']}", expanded=True):
                        st.write("")
                        st.markdown(f"{generiere_besitz_text(row['Eigentumsverhältnis'], row['Grundstücksnummer(n)'])}")
                        st.write("")
                        st.write("")
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.markdown("<span style='font-size: 0.8rem; color: #888;'>GRUNDSTÜCKSNUMMER(N)</span>", unsafe_allow_html=True)
                            st.write(str(row['Grundstücksnummer(n)']).replace(" / ", "\n\n"))
                        with c2:
                            st.markdown("<span style='font-size: 0.8rem; color: #888;'>EIGENTUMSVERHÄLTNIS</span>", unsafe_allow_html=True)
                            st.write(str(row['Eigentumsverhältnis']).replace(" / ", "\n\n"))
                        with c3:
                            st.markdown("<span style='font-size: 0.8rem; color: #888;'>FLÄCHE(N)</span>", unsafe_allow_html=True)
                            st.write(str(row['Fläche(n)']).replace(" / ", "\n\n"))
                        st.write("")
            else:
                st.markdown("<span style='color: #888;'>Es wurden keine Einträge gefunden.</span>", unsafe_allow_html=True)
        else:
            st.markdown("<span style='color: #888;'>Bitte starten Sie die Suche über das Eingabefeld.</span>", unsafe_allow_html=True)

    with tab2:
        st.write("")
        col_m1, col_m2 = st.columns(2)
        stadt_besitz = df[df['Eigentumsverhältnis'].str.contains("01", na=False)]
        privat_besitz = df[df['Eigentumsverhältnis'].str.contains("03", na=False)]
        col_m1.metric("Adressen mit Stadt-Beteiligung", len(stadt_besitz))
        col_m2.metric("Adressen in Privatbesitz", len(privat_besitz))
        
        st.write("")
        st.write("")
        st.markdown("<span style='font-size: 0.8rem; color: #888;'>TOP 10 DER GRÖSSTEN STÄDTISCHEN AREALE</span>", unsafe_allow_html=True)
        stadt_besitz_sort = stadt_besitz.copy()
        stadt_besitz_sort['Fläche_Zahl'] = stadt_besitz_sort['Fläche(n)'].str.extract(r'(\d+)').astype(float)
        top_10 = stadt_besitz_sort.sort_values(by='Fläche_Zahl', ascending=False).head(10)
        st.dataframe(top_10[['Adresse', 'Fläche(n)', 'Grundstücksnummer(n)']], use_container_width=True, hide_index=True)

except Exception as e:
    st.markdown("<span style='color: #888;'>Ein Systemfehler ist aufgetreten. Bitte laden Sie die Seite neu.</span>", unsafe_allow_html=True)
