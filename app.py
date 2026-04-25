import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="Immobilien-Portal Biel", 
    layout="wide"
)

def local_css(file_name):
    if os.path.exists(file_name):
        with open(file_name) as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

local_css("style.css")

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
            return (f"<strong>QUELLENRECHT</strong><br><br>Der Grund und Boden dieser Parzelle gehört <strong>{wer_boden}</strong>. "
                    f"Jedoch besitzt <strong>{wer_quelle}</strong> hier ein Quellenrecht. Diese Partei "
                    f"darf auf diesem fremden Grundstück eine Wasserquelle fassen und nutzen.")
        else:
            return (f"<strong>QUELLENRECHT</strong><br><br>Sowohl der Boden als auch das Recht zur Wassernutzung gehören <strong>{name_dativ(quelle_besitzer[0])}</strong>.")

    if bau_besitzer:
        einzig_boden = list(dict.fromkeys([name_dativ(b) for b in boden_besitzer]))
        einzig_bau = list(dict.fromkeys([name_dativ(b) for b in bau_besitzer]))
        einzig_bau_nom = list(dict.fromkeys([name_nominativ(b) for b in bau_besitzer]))
        
        txt_boden = " sowie ".join(einzig_boden)
        txt_bau = " sowie ".join(einzig_bau)
        txt_bau_nom = " sowie ".join(einzig_bau_nom)

        if txt_boden == txt_bau:
            return (f"<strong>BAURECHT</strong><br><br>Sowohl der Grund und Boden als auch das Gebäude gehören <strong>{txt_boden}</strong>. "
                    f"Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke, die unabhängig voneinander behandelt werden.")
        
        return (f"<strong>BAURECHT</strong><br><br>Der Grund und Boden gehört <strong>{txt_boden}</strong>. "
                f"Jedoch besitzt <strong>{txt_bau_nom}</strong> hier ein Baurecht. Das Gebäude gehört rechtlich <strong>{txt_bau}</strong>, obwohl der Boden <strong>{txt_boden}</strong> gehört.")

    if len(boden_besitzer) > 1:
        einzig_boden = list(dict.fromkeys([name_dativ(b) for b in boden_besitzer]))
        txt_boden = " sowie ".join(einzig_boden)
        return (f"<strong>GRENZFALL</strong><br><br>Dieses Gebäude steht auf mehreren Grundstücken gleichzeitig. "
                f"Der gesamte Boden gehört <strong>{txt_boden}</strong>.")

    return f"<strong>VOLLEIGENTUM</strong><br><br>Sowohl der Boden als auch das Gebäude gehören vollumfänglich <strong>{name_dativ(boden_besitzer[0])}</strong>."

try:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=60)
        
    df = load_excel()
    st.title("Immobilienregister Biel")
    st.markdown("<p style='font-size: 1.2rem; color: #86868b; font-weight: 400;'>Durchsuchen Sie die Eigentumsverhältnisse der Gebäude auf Basis amtlicher Geodaten.</p>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["Adress-Suche", "Bestandesübersicht"])

    with tab1:
        st.write("")
        search_query = st.text_input("Adresse", placeholder="Strasse und Hausnummer eingeben...", label_visibility="collapsed")
        
        if search_query:
            results = df[df['Adresse'].str.contains(search_query, case=False, na=False)]
            if not results.empty:
                st.markdown(f"<p style='color: #86868b; font-size: 0.85rem; font-weight: 600; letter-spacing: 0.05em; margin-top: 1rem; margin-bottom: 1.5rem;'>{len(results)} ERGEBNISSE GEFUNDEN</p>", unsafe_allow_html=True)
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
                st.markdown("<p style='color: #86868b; margin-top: 2rem;'>Es wurden keine Einträge gefunden. Bitte überprüfen Sie Ihre Eingabe.</p>", unsafe_allow_html=True)

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
    st.markdown("<p style='color: #86868b;'>Ein Systemfehler ist aufgetreten. Bitte laden Sie die Seite neu.</p>", unsafe_allow_html=True)
