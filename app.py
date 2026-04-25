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
    
    def name_finden(text):
        t = str(text).lower()
        if "01" in t: return "der Stadt Biel"
        if "03" in t: return "einer Privatperson oder privaten Firma"
        if "02" in t: return "einer öffentlichen Institution (Bund, Kanton oder SBB)"
        return "einem unbekannten Eigentümer"

    if "Quellenrecht" in str(nummern_string):
        if "/" in str(besitz_string):
            teile = str(besitz_string).split(" / ")
            boden_besitzer = name_finden(teile[0])
            quellen_besitzer = name_finden(teile[-1])
            
            return (f"💧 **Quellenrecht:** Der Grund und Boden dieser Parzelle gehört **{boden_besitzer}**. "
                    f"Jedoch besitzt **{quellen_besitzer}** hier ein sogenanntes Quellenrecht. "
                    f"Das bedeutet: Diese Partei hat das im Grundbuch verankerte Recht, eine Wasserquelle auf diesem fremden Grundstück zu fassen und das Wasser zu nutzen.")
        else:
            besitzer = name_finden(besitz_string)
            return (f"💧 **Quellenrecht:** Der Grund und Boden sowie das dazugehörige Quellenrecht gehören **{besitzer}**. "
                    f"Das bedeutet: Der Eigentümer hat hier das speziell im Grundbuch verankerte Recht, eine Wasserquelle auf dem Grundstück zu fassen und zu nutzen.")

    elif "/" in str(besitz_string):
        teile = str(besitz_string).split(" / ")
        if len(teile) >= 2:
            return (f"🏢 **Besondere Situation (Baurecht):** Der Grund und Boden gehört **{name_finden(teile[0])}**. "
                    f"Das Gebäude darauf gehört rechtlich jedoch **{name_finden(teile[1])}**.")
    
    return f"🏡 **Vollständiges Eigentum:** Sowohl der Boden als auch das Gebäude gehören vollumfänglich **{name_finden(besitz_string)}**."

try:
    df = load_excel()
    
    st.title("🏛️ Immobilien-Register der Stadt Biel")
    st.markdown("Durchsuche die Eigentumsverhältnisse aller Gebäude auf Basis amtlicher Geodaten.")

    tab1, tab2 = st.tabs(["🔍 Adress-Suche", "📊 Stadt-Portfolio & Statistik"])

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
                            st.caption("Eigentumsverhältnis")
                            st.write(str(row['Eigentumsverhältnis']).replace(" / ", "\n\n"))
                        with c3:
                            st.caption("Fläche(n)")
                            st.write(str(row['Fläche(n)']).replace(" / ", "\n\n"))
            else:
                st.warning("Keine Treffer unter dieser Adresse gefunden.")

    with tab2:
        st.header("Auswertung")
        col_m1, col_m2 = st.columns(2)
        
        stadt_besitz = df[df['Eigentumsverhältnis'].str.contains("01", na=False)]
        privat_besitz = df[df['Eigentumsverhältnis'].str.contains("03", na=False)]
        
        col_m1.metric("Adressen mit Stadt-Beteiligung", len(stadt_besitz))
        col_m2.metric("Adressen in Privatbesitz", len(privat_besitz))
        
        st.markdown("### Top 10 der grössten städtischen Areale (Bebaute Parzellen)")
        
        stadt_besitz_sort = stadt_besitz.copy()
        stadt_besitz_sort['Fläche_Zahl'] = stadt_besitz_sort['Fläche(n)'].str.extract(r'(\d+)').astype(float)
        top_10 = stadt_besitz_sort.sort_values(by='Fläche_Zahl', ascending=False).head(10)
        
        st.dataframe(
            top_10[['Adresse', 'Fläche(n)', 'Grundstücksnummer(n)']],
            use_container_width=True,
            hide_index=True
        )

except Exception as e:
    st.error(f"Konnte die App nicht laden. Ist die Datei 'Biel_Adressregister_Final.xlsx' im selben Ordner?")
    st.info(f"Details: {e}")
