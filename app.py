"""
Immobilienregister Biel – Streamlit-App
"""

import os
import re
import json
import urllib.parse

import pandas as pd
import pydeck as pdk
import streamlit as st
from streamlit_lottie import st_lottie

# ── 1. KONSTANTEN ────────────────────────────────────────────────────────────

EXCEL_FILE  = "Biel_Adressregister_Final.xlsx"
GEOJSON_FILE = "Eigentum.md"
LOTTIE_FILE = "ajour-logo.json"
SHEET_NAME  = "Adress-Verzeichnis"

# Eigentümer-Codes → (Dativ, Nominativ)
EIGENTUEMER = {
    "01": (
        "der Stadt Biel",
        "die Stadt Biel",
    ),
    "02": (
        "einer öffentlichen Institution (Bund, Kanton, SBB oder Ähnliche)",
        "eine öffentliche Institution (Bund, Kanton, SBB oder Ähnliche)",
    ),
    "03": (
        "einer Privatperson oder privaten Firma",
        "eine Privatperson oder private Firma",
    ),
}

# Farben pro Kategorie → (line_color RGBA, fill_color RGBA)
KATEGORIE_FARBEN = {
    "Vollbesitz":    ([0,   122, 255, 255], [0,   122, 255, 60]),
    "Bodenbesitz":   ([90,  200, 250, 255], [90,  200, 250, 60]),
    "Gebäudebesitz": ([255, 179,  64, 255], [255, 179,  64, 60]),
    "Andere":        ([255, 149,   0, 255], [255, 149,   0, 60]),
}
KATEGORIE_NAMEN = {
    "Vollbesitz":    "Vollbesitz Stadt",
    "Bodenbesitz":   "Bodenbesitz Stadt (Baurecht abgegeben)",
    "Gebäudebesitz": "Gebäudebesitz Stadt (Baurecht erhalten)",
    "Andere":        "Privat / Andere",
}

FILTER_OPTIONEN = [
    "Alle Adressen",
    "Vollbesitz der Stadt (Gebäude und Land)",
    "Bodenbesitz der Stadt (Land im Baurecht abgegeben)",
    "Gebäudebesitz der Stadt (Land im Baurecht erhalten)",
]
FILTER_HINWEISE = {
    "Alle Adressen":   "💡 Zeigt das gesamte Register. <strong>Bitte Suchbegriff eingeben.</strong>",
    "Vollbesitz":      "💡 Adressen, bei denen Boden und Gebäude vollständig der Stadt Biel gehören.",
    "Bodenbesitz":     "💡 Die Stadt besitzt das Land, hat es aber an Dritte im Baurecht abgegeben.",
    "Gebäudebesitz":   "💡 Der Boden gehört jemand anderem, aber die Stadt besitzt darauf ein Gebäude im Baurecht.",
}

# ── 2. SEITENKONFIGURATION ───────────────────────────────────────────────────

st.set_page_config(page_title="Immobilienregister Biel", layout="wide")

# Session-State initialisieren (immer ganz oben, vor dem ersten Rendering)
if "results_limit" not in st.session_state:
    st.session_state.results_limit = 20
if "filter_mode" not in st.session_state:
    st.session_state.filter_mode = FILTER_OPTIONEN[0]

# ── 3. CSS ───────────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

#MainMenu, footer, header { visibility: hidden; }

[data-testid="stAppViewContainer"] {
    font-family: 'Inter', sans-serif !important;
    background-color: #FAFAFA !important;
}
.block-container {
    padding-top: 0 !important;
    padding-bottom: 4rem;
    max-width: 900px;
}
/* Suchfeld */
.stTextInput,
.stTextInput > div,
.stTextInput > div > div {
    overflow: visible !important;
}
.stTextInput > div > div > input {
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    font-size: 1.2rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    color: #111111 !important;
    box-shadow: 0 8px 30px rgba(0,0,0,0.06);
    margin-bottom: 8px;
}
.stTextInput > div > div > input:focus {
    border-color: #AAAAAA !important;
    box-shadow: 0 8px 30px rgba(0,0,0,0.06) !important;
    outline: none !important;
}
div[data-testid="stExpander"] {
    border-radius: 12px;
    margin-bottom: 1rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
}
[data-testid="stExpanderDetails"] {
    padding-top: 0.5rem !important;
    padding-bottom: 1.25rem !important;
}
[data-testid="stExpanderDetails"] hr {
    margin-top: 0.4rem !important;
    margin-bottom: 0.75rem !important;
}
.main-title {
    text-align: center; font-weight: 700; font-size: 2.8rem;
    line-height: 1.15;
    letter-spacing: -0.03em; margin-top: 1rem; margin-bottom: 0.6rem; color: #111111 !important;
}
.title-subtext {
    text-align: center; color: #888888 !important;
    margin-top: 0; margin-bottom: 2rem; font-size: 1.05rem;
}
/* Lottie: Whitespace komprimieren */
[data-testid="stCustomComponentV1"] {
    margin-top: -0.5rem !important;
    margin-bottom: -1.5rem !important;
}
@media (max-width: 768px) {
    .main-title {
        font-size: 2rem;
        line-height: 1.1;
    }
    [data-testid="stCustomComponentV1"] {
        margin-top: -0.5rem !important;
        margin-bottom: -2.5rem !important;
    }
    /* Mobile: Selectbox-Filter zeigen, Radio-Pills verstecken */
    .stRadio { display: none !important; }
    [data-testid="stSelectbox"] { display: block !important; }
}
.label-text {
    font-size: 0.75rem; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.08em; color: #86868B !important;
    margin-top: 0.9rem;
    margin-bottom: 0.1rem;
    display: block;
}
.methodology-box {
    margin-top: 4rem; padding: 2rem; border-radius: 12px;
    background-color: #F2F2F7; font-size: 0.9rem; color: #555555; line-height: 1.6;
}
.stTabs [aria-selected="true"] {
    color: #111111 !important;
    border-bottom: 2px solid #111111 !important;
}
.stRadio [data-testid="stWidgetLabel"] { display: none; }
div[role="radiogroup"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: wrap !important;
    gap: 8px !important;
    margin-top: 0.5rem;
    margin-bottom: 1.5rem;
}
div[role="radiogroup"] > label {
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    padding: 8px 16px !important;
    border-radius: 30px !important;
    cursor: pointer;
}
div[role="radiogroup"] > label:has(input:checked) {
    background-color: #111111 !important;
    border-color: #111111 !important;
}
div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; }
.legend-box {
    padding: 15px; border-radius: 12px; background-color: #FFFFFF;
    border: 1px solid #EAEAEA; margin-bottom: 15px;
    display: flex; gap: 15px; flex-wrap: wrap;
}
.legend-item {
    display: flex; align-items: center;
    font-size: 0.85rem; font-weight: 500; color: #111111;
}
.legend-color {
    width: 14px; height: 14px; border-radius: 4px;
    margin-right: 8px; border: 1px solid rgba(0,0,0,0.1);
}
.maps-link {
    font-size: 0.85rem; color: #0066CC;
    text-decoration: none; font-weight: 500;
}
/* Desktop: Selectbox-Filter verstecken (nur Radio-Pills sichtbar) */
[data-testid="stSelectbox"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ── 4. METHODIK TEXT ─────────────────────────────────────────────────────────

METHODIK_TEXT = """
**Methodik & Datenquellen:**

Die Daten stammen aus dem öffentlichen WebGIS der Stadt Biel (Stand: 26.11.2025). Da die Rohdaten komplex sind, haben wir sie mit einer eigenen Logik neu aufbereitet:

	<strong>1. Besitz-Check:</strong> Wir haben für jede Parzelle automatisiert analysiert, wer involviert ist (z. B. wenn die Stadt den Boden besitzt, aber jemand anderes das Baurecht).
	<strong>2. Daten-Fusion:</strong> Wir haben die geografischen Pläne der Stadt mit dem Adressregister des Bundes [map.geo.admin.ch](https://map.geo.admin.ch) verknüpft, damit man Grundstücke einfach per Adresse finden kann.
	<strong>3. Einfachheit:</strong> Bei komplizierten Fällen (wie vielen verschiedenen Eigentümern in einem Haus) haben wir die Darstellung vereinfacht, um die Übersichtlichkeit zu wahren.

Dieses Tool dient ausschliesslich der Orientierung. Es bietet keine verbindliche Auskunft. Bei komplexen Grenz- oder Stockwerkeigentums-Fällen können vereinzelte Ungenauigkeiten auftreten.
"""

def methodik_als_html(text: str) -> str:
    """Konvertiert Markdown-Grundelemente zu HTML."""
    html = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)
    html = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2" target="_blank">\1</a>', html)
    return html.replace('\n', '<br>')


# ── 5. HILFSFUNKTIONEN ───────────────────────────────────────────────────────

def eigentuemer_dativ(code_str: str) -> str:
    code = next((k for k in EIGENTUEMER if k in str(code_str)), None)
    return EIGENTUEMER[code][0] if code else "einem unbekannten Eigentümer"

def eigentuemer_nominativ(code_str: str) -> str:
    code = next((k for k in EIGENTUEMER if k in str(code_str)), None)
    return EIGENTUEMER[code][1] if code else "ein unbekannter Eigentümer"

def lv95_to_wgs84(y: float, x: float) -> list[float]:
    y_p = (y - 2_600_000) / 1_000_000
    x_p = (x - 1_200_000) / 1_000_000
    lon = (2.6779094 + 4.728982*y_p + 0.791484*y_p*x_p
           + 0.1306*y_p*x_p**2 - 0.0436*y_p**3) * 100 / 36
    lat = (16.9023892 + 3.238272*x_p - 0.270978*y_p**2 - 0.002528*x_p**2
           - 0.0447*y_p**2*x_p - 0.0140*x_p**3) * 100 / 36
    return [lon, lat]

def recursive_convert(coords):
    if isinstance(coords[0], (int, float)):
        return lv95_to_wgs84(coords[0], coords[1])
    return [recursive_convert(c) for c in coords]

def bestimme_kategorie(row) -> str:
    besitz  = str(row['Eigentumsverhältnis']).split(" / ")
    nummern = str(row['Grundstücksnummer(n)']).split(" / ")
    boden, bau = [], []
    for i, b in enumerate(besitz):
        info = nummern[i] if i < len(nummern) else ""
        if "Baurecht" in info:
            bau.append(b)
        elif "Quellenrecht" not in info:
            boden.append(b)
    stadt_boden = any("01" in b for b in boden)
    stadt_bau   = any("01" in b for b in bau)
    if stadt_boden and not bau:               return "Vollbesitz"
    if stadt_boden and bau and not stadt_bau: return "Bodenbesitz"
    if not stadt_boden and stadt_bau:         return "Gebäudebesitz"
    if stadt_boden and stadt_bau:
        return "Bodenbesitz" if any("01" not in b for b in bau) else "Vollbesitz"
    return "Andere"

def generiere_besitz_text(besitz_string: str, nummern_string: str) -> str:
    if not besitz_string:
        return "Keine Daten verfügbar."

    b_list = str(besitz_string).split(" / ")
    n_list = str(nummern_string).split(" / ")
    boden, bau, quelle = [], [], []
    for i, b in enumerate(b_list):
        info = n_list[i] if i < len(n_list) else ""
        if "Quellenrecht" in info: quelle.append(b)
        elif "Baurecht" in info:   bau.append(b)
        else:                      boden.append(b)

    def unique_d(items): return " sowie ".join(dict.fromkeys(eigentuemer_dativ(b)    for b in items))
    def unique_n(items): return " sowie ".join(dict.fromkeys(eigentuemer_nominativ(b) for b in items))

    if quelle:
        return (
            f"<strong>QUELLENRECHT</strong><br><br>"
            f"Der Grund und Boden gehört <strong>{eigentuemer_dativ(boden[0])}</strong>. "
            f"Jedoch besitzt <strong>{eigentuemer_nominativ(quelle[0])}</strong> "
            f"hier ein Quellenrecht zur Wassernutzung."
        )
    if bau:
        txt_boden_d = unique_d(boden)
        txt_bau_n   = unique_n(bau)
        txt_bau_d   = unique_d(bau)
        if txt_boden_d == txt_bau_d:
            return (
                f"<strong>BAURECHT</strong><br><br>"
                f"Sowohl der Grund und Boden als auch das Gebäude gehören "
                f"<strong>{txt_boden_d}</strong>. Rechtlich sind dies jedoch zwei "
                f"getrennte Grundstücke, die im Register unabhängig behandelt werden."
            )
        return (
            f"<strong>BAURECHT</strong><br><br>"
            f"Der Grund gehört <strong>{txt_boden_d}</strong>. "
            f"Jedoch besitzt <strong>{txt_bau_n}</strong> hier ein Baurecht — "
            f"das Gebäude gehört somit <strong>{txt_bau_d}</strong>, "
            f"obwohl der Boden weiterhin <strong>{txt_boden_d}</strong> gehört."
        )
    if len(boden) > 1:
        return (
            f"<strong>GRENZFALL / MITBESITZ / STOCKWERKEIGENTUM</strong><br><br>"
            f"Dieses Objekt gehört <strong>{unique_d(boden)}</strong> gemeinsam."
        )
    return (
        f"<strong>VOLLEIGENTUM</strong><br><br>"
        f"Sowohl der Grund und Boden als auch das darauf stehende Gebäude gehören "
        f"vollumfänglich <strong>{eigentuemer_dativ(boden[0])}</strong>."
    )


# ── 6. DATEN LADEN ───────────────────────────────────────────────────────────

@st.cache_data
def load_data() -> pd.DataFrame | None:
    if not os.path.exists(EXCEL_FILE):
        st.error(f"Datei nicht gefunden: {EXCEL_FILE}")
        return None
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME).fillna("")
    df['Fläche_Zahl'] = df['Fläche(n)'].str.extract(r'(\d+)')[0].astype(float).fillna(0)
    df['Filter_Kategorie'] = df.apply(bestimme_kategorie, axis=1)
    return df

@st.cache_data
def load_geojson_map_data(df: pd.DataFrame) -> list | None:
    if not os.path.exists(GEOJSON_FILE):
        st.error(f"Datei nicht gefunden: {GEOJSON_FILE}")
        return None
    try:
        with open(GEOJSON_FILE, encoding="utf-8") as f:
            raw_data = json.load(f)
    except json.JSONDecodeError as e:
        st.error(f"GeoJSON konnte nicht gelesen werden: {e}")
        return None

    features: list = []
    if isinstance(raw_data, dict):
        if raw_data.get("type") == "FeatureCollection":
            features = raw_data.get("features", [])
        else:
            for val in raw_data.values():
                if isinstance(val, dict) and val.get("type") == "FeatureCollection":
                    features.extend(val.get("features", []))

    mapping: dict[str, str] = {}
    for _, row in df.iterrows():
        for num in re.findall(r'\d+', str(row['Grundstücksnummer(n)'])):
            mapping[num] = row['Filter_Kategorie']

    for feature in features:
        feature['geometry']['coordinates'] = recursive_convert(
            feature['geometry']['coordinates']
        )
        gn  = str(feature["properties"].get("grst_nummer", "")).strip()
        eg  = str(feature["properties"].get("eigentuemer_gr", ""))
        kat = mapping.get(gn, "Vollbesitz" if eg == "01" else "Andere")
        lc, fc = KATEGORIE_FARBEN.get(kat, KATEGORIE_FARBEN["Andere"])
        feature["properties"]["line_color"]     = lc
        feature["properties"]["fill_color"]     = fc
        feature["properties"]["kategorie_name"] = KATEGORIE_NAMEN.get(kat, "Unbekannt")

    return features


# ── 7. APP RENDERING ─────────────────────────────────────────────────────────

df = load_data()
if df is None:
    st.stop()

# Logo (Lottie-Animation – volle Breite wegen 1536x280 Canvas)
if os.path.exists(LOTTIE_FILE):
    with open(LOTTIE_FILE, encoding="utf-8") as f:
        lottie_data = json.load(f)
    st_lottie(lottie_data, height=100, loop=True, quality="high", key="logo")

st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
st.markdown("<div class='title-subtext'>Suchportal für den Immobilienbesitz der Stadt Biel</div>", unsafe_allow_html=True)

# Tab umbenannt in "Suche"
t1, t2 = st.tabs(["🔍 Suche", "Interaktive Karte"])

# ── Tab 1: Suche ─────────────────────────────────────────────────────────────
with t1:
    search = st.text_input(
        "Suche",
        placeholder="Strasse und Hausnummer",
        label_visibility="collapsed",
        key="search_input",
    )

    def clear_search():
        st.session_state.search_input = ""

    # Buttons: 50/50 → garantiert nebeneinander auf Desktop & Mobile
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        st.button("🔍 Suchen", use_container_width=True)
    with col_btn2:
        st.button("✕ Löschen", on_click=clear_search, use_container_width=True)

    # Filter: Desktop = Radio-Pills, Mobile = Selectbox (CSS show/hide)
    cur_idx = next(
        (i for i, o in enumerate(FILTER_OPTIONEN) if o == st.session_state.filter_mode), 0
    )

    def _sync_radio():
        st.session_state.filter_mode = st.session_state._f_radio

    def _sync_select():
        st.session_state.filter_mode = st.session_state._f_select

    # Desktop-Filter (Radio-Pills)
    st.radio(
        "Filter", FILTER_OPTIONEN, index=cur_idx, horizontal=True,
        label_visibility="collapsed", key="_f_radio", on_change=_sync_radio,
    )
    # Mobile-Filter (Selectbox – per CSS auf Desktop versteckt)
    st.selectbox(
        "Filter", FILTER_OPTIONEN, index=cur_idx,
        label_visibility="collapsed", key="_f_select", on_change=_sync_select,
    )

    f_mode = st.session_state.filter_mode
    hinweis_key = next((k for k in FILTER_HINWEISE if k in f_mode), "Alle Adressen")
    st.markdown(
        f"<p style='color:#888888; font-size:0.85rem; margin-top:-6px; margin-bottom:20px;'>"
        f"{FILTER_HINWEISE[hinweis_key]}</p>",
        unsafe_allow_html=True,
    )

    search = st.session_state.get("search_input", "")
    f_df = df.copy()
    if "Vollbesitz" in f_mode:      f_df = f_df[f_df['Filter_Kategorie'] == "Vollbesitz"]
    elif "Bodenbesitz" in f_mode:   f_df = f_df[f_df['Filter_Kategorie'] == "Bodenbesitz"]
    elif "Gebäudebesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Gebäudebesitz"]
    if search:
        f_df = f_df[f_df['Adresse'].str.contains(search, case=False, na=False)]

    if f_mode == "Alle Adressen" and not search:
        st.info("Bitte Adresse eingeben oder Filter wählen.")
    elif f_df.empty:
        st.info("Keine Treffer.")
    else:
        st.markdown(
            f"<div style='margin-bottom:1rem; opacity:0.6; font-size:0.8rem;'>{len(f_df)} Treffer</div>",
            unsafe_allow_html=True,
        )
        for _, r in f_df.iloc[:st.session_state.results_limit].iterrows():
            with st.expander(str(r['Adresse'])):
                st.markdown(
                    generiere_besitz_text(r['Eigentumsverhältnis'], r['Grundstücksnummer(n)']),
                    unsafe_allow_html=True,
                )
                maps_query = urllib.parse.quote(f"{r['Adresse']}, Biel")
                st.markdown(
                    f'<a href="https://www.google.com/maps/search/?api=1&query={maps_query}"'
                    f' target="_blank" class="maps-link">📍 Auf Google Maps anzeigen</a>',
                    unsafe_allow_html=True,
                )
                st.write("---")
                c1, c2, c3 = st.columns(3)
                eigentuem_clean = re.sub(r'\d{2}:\s*', '', str(r['Eigentumsverhältnis']))
                c1.markdown(f"<div class='label-text'>Parzelle</div>{r['Grundstücksnummer(n)']}", unsafe_allow_html=True)
                c2.markdown(f"<div class='label-text'>Eigentum</div>{eigentuem_clean}", unsafe_allow_html=True)
                c3.markdown(f"<div class='label-text'>Fläche</div>{r['Fläche(n)']}", unsafe_allow_html=True)

        if len(f_df) > st.session_state.results_limit:
            if st.button("Weitere laden"):
                st.session_state.results_limit += 30
                st.rerun()

# ── Tab 2: Karte ─────────────────────────────────────────────────────────────
@st.fragment
def render_karte():
    """Fragment: rendert nur neu wenn interne Widgets sich ändern,
    nicht bei Filter-/Sucheingaben in Tab 1."""
    st.markdown("""
    <div class='legend-box'>
        <div class='legend-item'>
            <div class='legend-color' style='background-color:rgba(0,122,255,0.3); border-color:#007AFF;'></div>
            Vollbesitz (Stadt)
        </div>
        <div class='legend-item'>
            <div class='legend-color' style='background-color:rgba(90,200,250,0.3); border-color:#5AC8FA;'></div>
            Bodenbesitz (Baurecht abg.)
        </div>
        <div class='legend-item'>
            <div class='legend-color' style='background-color:rgba(255,149,0,0.3); border-color:#FF9500;'></div>
            Privat / Andere
        </div>
        <div class='legend-item'>
            <div class='legend-color' style='background-color:rgba(255,179,64,0.3); border-color:#FFB340;'></div>
            Gebäudebesitz (Baurecht erh.)
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("Lade Karte..."):
        geo = load_geojson_map_data(df)

    if geo:
        st.pydeck_chart(pdk.Deck(
            map_provider="carto",
            map_style="light",
            initial_view_state=pdk.ViewState(latitude=47.1368, longitude=7.2468, zoom=14.0),
            layers=[pdk.Layer(
                "GeoJsonLayer",
                data={"type": "FeatureCollection", "features": geo},
                opacity=1.0,
                stroked=True,
                filled=True,
                get_fill_color="properties.fill_color",
                get_line_color="properties.line_color",
                line_width_min_pixels=2,
                pickable=True,
            )],
            tooltip={
                "html": "<b>Parzelle:</b> {grst_nummer}<br/><b>Kategorie:</b> {kategorie_name}",
                "style": {"backgroundColor": "steelblue", "color": "white"},
            },
        ))

with t2:
    render_karte()

# ── 8. FOOTER ────────────────────────────────────────────────────────────────
st.markdown(
    f"<div class='methodology-box'>{methodik_als_html(METHODIK_TEXT)}</div>",
    unsafe_allow_html=True,
)
