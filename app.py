import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Immobilien-Portal Biel", 
    page_icon="🇨🇭", 
    layout="wide"
)

@st.cache_data
def load_excel():
    df = pd.read_excel('Biel_Adressregister_Final.xlsx', sheet_name='Adress-Verzeichnis')
    return df.fillna("")

def generiere_besitz_text(besitz_string, nummern_string):
    if not besitz_string: 
        return "Zu diesem Objekt liegen leider keine detaillierten Eigentumsdaten vor."
    
    def name_dativ(text):
        t = str(text).lower()
        if "01" in t: return "der Stadt Biel"
        if "03" in t: return "einer Privatperson oder privaten Firma"
        if "02" in t: return "einer öffentlichen Institution (Bund, Kanton oder SBB)"
        return "einem unbekannten Eigentümer"

    def name_nominativ(text):
        t = str(text).lower()
        if "01" in t: return "die Stadt Biel"
        if "03" in t: return "eine Privatperson oder private Firma"
        if "02" in t: return "eine öffentliche Institution (Bund, Kanton oder SBB)"
        return "ein unbekannter Eigentümer"

    besitzer_liste = str(besitz_string).split(" / ")
    info_liste = str(nummern_string).split(" / ")
    
    # Daten sortieren nach Typ
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

    # FALL 1: Quellenrecht
    if quelle_besitzer:
        wer_quelle = name_nominativ(quelle_besitzer[0])
        if boden_besitzer:
            wer_boden = name_dativ(boden_besitzer[0])
            return (f"💧 **Quellenrecht:** Der Grund und Boden dieser Parzelle gehört **{wer_boden}**. "
                    f"Jedoch besitzt **{wer_quelle}** hier ein Quellenrecht. Das bedeutet: Diese Partei "
                    f"darf auf diesem fremden Grundstück eine Wasserquelle fassen und nutzen.")
        else:
            return (f"💧 **Quellenrecht:** Sowohl der Boden als auch das Recht zur Wassernutzung gehören **{name_dativ(quelle_besitzer[0])}**.")

    # FALL 2: Baurecht (Gebäude und Boden getrennt)
    if bau_besitzer:
        # Einzigartige Besitzer sammeln für die Grammatik
        einzig_boden = list(dict.fromkeys([name_dativ(b) for b in boden_besitzer]))
        einzig_bau = list(dict.fromkeys([name_dativ(b) for b in bau_besitzer]))
        einzig_bau_nom = list(dict.fromkeys([name_nominativ(b) for b in bau_besitzer]))
        
        txt_boden = " sowie ".join(einzig_boden)
        txt_bau = " sowie ".join(einzig_bau)
        txt_bau_nom = " sowie ".join(einzig_bau_nom)

        # Spezialfall: Gleiche Kategorie besitzt beides
        if txt_boden == txt_bau:
            return (f"🏢 **Besondere Situation (Baurecht):** Sowohl der Grund und Boden als auch das Gebäude gehören **{txt_boden}**. "
                    f"Rechtlich gesehen sind dies jedoch zwei getrennte Grundstücke, die unabhängig voneinander behandelt werden.")
        
        return (f"🏢 **Besondere Situation (Baurecht):** Der Grund und Boden gehört **{txt_boden}**. "
                f"Jedoch besitzt **{txt_bau_nom}** hier ein Baurecht. Das bedeutet: Das Gebäude gehört rechtlich **{txt_bau}**, obwohl der Boden jemand anderem gehört.")

    # FALL 3: Mehrere Böden (Grenz-Reiter)
    if len(boden_besitzer) > 1:
        einzig_boden = list(dict.fromkeys([name_dativ(b) for b in boden_besitzer]))
        txt_boden = " sowie ".join(einzig_boden)
        return (f"🏘️ **Grenzfall:** Dieses Gebäude steht auf mehreren Grundstücken gleichzeitig. "
                f"Der gesamte Boden gehört **{txt_boden}**.")

    # FALL 4: Normalfall
    return f"🏡 **Vollständiges Eigentum:** Sowohl der Boden als auch das Gebäude gehören vollumfänglich **{name_dativ(boden_besitzer[0])}**."

try:
    df = load_excel()
    st.title("🏛️ Immobilien-Register der Stadt Biel")
    st.markdown("Durchsuche die Eigentumsverhältnisse aller Gebäude auf Basis amtlicher Geodaten.")

    tab1, tab2 = st.tabs(["🔍 Adress-Suche", "📊 Bestandesübersicht"])

    with tab1:
        search_query = st.text_input("Adresse suchen (z.B. Südstrasse 82):", "")
        if search_query:
            results = df[df['Adresse'].str.contains(search_query, case=False, na=False)]
            if not results.empty:
                st.success(f"**{len(results)} Ergebnis(se) gefunden:**")
                for _, row in results.iterrows():
                    with st.expander(f"📍 {row['Adresse']}", expanded=True):
                        st.info(generiere_besitz_text(row['Eigentumsverhältnis'], row['Grundstücksnummer(n)']))
                        st.markdown("---")
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.caption("Grundstücksnummer(n)")
                            st.write(str(row['Grundstücksnummer(n)']).replace(" / ", "\n\n"))
                        with c2:
                            st.caption("Technisches Eigentumsverhältnis")
                            st.write(str(row['Eigentumsverhältnis']).replace(" / ", "\n\n"))
                        with c3:
                            st.caption("Fläche(n)")
                            st.write(str(row['Fläche(n)']).replace(" / ", "\n\n"))
            else:
                st.warning("Keine Treffer unter dieser Adresse gefunden.")

    with tab2:
        col_m1, col_m2 = st.columns(2)
        stadt_besitz = df[df['Eigentumsverhältnis'].str.contains("01", na=False)]
        privat_besitz = df[df['Eigentumsverhältnis'].str.contains("03", na=False)]
        col_m1.metric("Adressen mit Stadt-Beteiligung", len(stadt_besitz))
        col_m2.metric("Adressen in Privatbesitz", len(privat_besitz))
        
        st.markdown("### Top 10 der grössten städtischen Areale")
        stadt_besitz_sort = stadt_besitz.copy()
        stadt_besitz_sort['Fläche_Zahl'] = stadt_besitz_sort['Fläche(n)'].str.extract(r'(\d+)').astype(float)
        top_10 = stadt_besitz_sort.sort_values(by='Fläche_Zahl', ascending=False).head(10)
        st.dataframe(top_10[['Adresse', 'Fläche(n)', 'Grundstücksnummer(n)']], use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Fehler beim Laden der App. Details: {e}")
