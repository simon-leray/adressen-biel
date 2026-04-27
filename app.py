"""
Immobilienregister Biel – Streamlit-App
"""

import os
import re
import time
import json
import html
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
        "einer öffentlich-rechtlichen Institution (z. B. Bund, Kanton, SBB, Landeskirchen)",
        "eine öffentlich-rechtliche Institution (z. B. Bund, Kanton, SBB, Landeskirchen)",
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
    "Alle",
    "Vollbesitz",
    "Bodenbesitz",
    "Gebäudebesitz",
]
FILTER_HINWEISE = {
    "Alle":          "💡 Zeigt das gesamte Register. <strong>Bitte Suchbegriff eingeben.</strong>",
    "Vollbesitz":    "💡 Adressen, bei denen Boden und Gebäude vollständig der Stadt Biel gehören.",
    "Bodenbesitz":   "💡 Die Stadt besitzt das Land, hat es aber an Dritte im Baurecht abgegeben.",
    "Gebäudebesitz": "💡 Der Boden gehört jemand anderem, aber die Stadt besitzt darauf ein Gebäude im Baurecht.",
}

# ── 2. SEITENKONFIGURATION ───────────────────────────────────────────────────

st.set_page_config(page_title="Immobilienregister Biel", layout="wide")

PAGE_SIZE = 20

# Session-State initialisieren (immer ganz oben, vor dem ersten Rendering)
if "page" not in st.session_state:
    st.session_state.page = 1
if "filter_mode" not in st.session_state:
    st.session_state.filter_mode = FILTER_OPTIONEN[0]
if "prev_search" not in st.session_state:
    st.session_state.prev_search = ""

# ── 3. CSS ───────────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

#MainMenu { display: none !important; }
footer { display: none !important; }
[data-testid="stHeader"] { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
[data-testid="stStatusWidget"] { display: none !important; }
[data-testid="stDeployButton"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
/* Letzten Tab (Info) rechtsbündig — CSS-Backup für JS */
[data-baseweb="tab-list"] [data-baseweb="tab"]:last-of-type {
    margin-left: auto !important;
}

[data-testid="stAppViewContainer"] {
    font-family: 'Inter', sans-serif !important;
    background-color: #FAFAFA !important;
}
.block-container {
    padding-top: 0 !important;
    padding-bottom: 4rem;
    max-width: 900px;
}
/* Suchfeld – Wrapper-Overflow freigeben damit box-shadow nicht abgeschnitten wird */
.stTextInput,
.stTextInput > div,
.stTextInput > div > div {
    overflow: visible !important;
}
/* BaseWeb-Container komplett transparent schalten */
[data-baseweb="input"],
[data-baseweb="base-input"] {
    border: none !important;
    box-shadow: none !important;
    outline: none !important;
    background: transparent !important;
}
[data-baseweb="input"]:focus-within,
[data-baseweb="base-input"]:focus-within,
.stTextInput > div > div:focus-within {
    border: none !important;
    box-shadow: none !important;
    outline: none !important;
}
/* Styling nur auf dem <input> selbst */
.stTextInput input {
    border-radius: 12px !important;
    padding: 1.2rem 1.5rem !important;
    font-size: 1.2rem !important;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
    color: #111111 !important;
    box-shadow: 0 8px 30px rgba(0,0,0,0.06) !important;
    margin-bottom: 8px !important;
    outline: none !important;
}
.stTextInput input:focus,
.stTextInput input:focus-visible {
    border-color: #AAAAAA !important;
    box-shadow: 0 8px 30px rgba(0,0,0,0.06) !important;
    outline: none !important;
}
[data-testid="InputInstructions"] { display: none !important; }
/* Resultate: native <details> statt st.expander */
details.ei {
    border-radius: 12px;
    margin-bottom: 0.75rem;
    background: #FFFFFF;
    border: 1px solid #EAEAEA;
    overflow: hidden;
}
details.ei summary {
    padding: 0.9rem 1.25rem;
    min-height: 56px;
    display: flex;
    align-items: center;
    cursor: pointer;
    font-weight: 500;
    font-size: 1rem;
    color: #111;
    list-style: none;
    user-select: none;
}
details.ei summary::-webkit-details-marker { display: none; }
details.ei summary::after {
    content: '›';
    margin-left: auto;
    font-size: 1.4rem;
    color: #AAAAAA;
    transition: transform 0.18s;
    display: inline-block;
}
details.ei[open] summary::after { transform: rotate(90deg); }
details.ei summary:hover { background: #F7F7F7; }
.ei-body {
    padding: 0.6rem 1.25rem 1.25rem;
    border-top: 1px solid #F0F0F0;
    font-size: 0.92rem;
    line-height: 1.6;
}
.ei-meta {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 1rem;
    margin-top: 0.75rem;
    padding-top: 0.75rem;
    border-top: 1px solid #F0F0F0;
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
/* Mobile-Anpassungen (Radio/Selectbox-Toggle macht JS) */
@media (max-width: 768px) {
    .main-title {
        font-size: 2rem;
        line-height: 1.1;
    }
    [data-testid="stCustomComponentV1"] {
        margin-top: -0.5rem !important;
        margin-bottom: -2.5rem !important;
    }
    /* Suchfeld-Buttons nebeneinander */
    [data-testid="column"] {
        width: calc(50% - 8px) !important;
        flex: 1 1 calc(50% - 8px) !important;
        min-width: 0 !important;
    }
    /* Selectbox: gleicher Abstand und gleiche Optik wie die Buttons */
    [data-testid="stSelectbox"] {
        margin-top: 0 !important;
        margin-bottom: 0 !important;
        padding-bottom: 0 !important;
    }
    [data-testid="stSelectbox"] > div {
        margin-top: 0 !important;
    }
    [data-testid="stSelectbox"] [data-baseweb="select"] > div:first-child {
        background-color: #FFFFFF !important;
        border: 1px solid rgba(49, 51, 63, 0.2) !important;
        border-radius: 0.5rem !important;
        min-height: 38px !important;
        box-shadow: none !important;
    }
    [data-testid="stSelectbox"] [data-baseweb="select"] > div:first-child:hover {
        border-color: rgba(49, 51, 63, 0.4) !important;
    }
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
button[data-baseweb="tab"] p,
button[data-baseweb="tab"] { font-size: 1.1rem !important; }
.stTabs [aria-selected="true"] {
    color: #111111 !important;
    border-bottom: 2px solid #111111 !important;
}
.stRadio [data-testid="stWidgetLabel"] { display: none; }
[data-testid="stRadio"],
[data-testid="stRadio"] > div,
[data-testid="stRadio"] > div > div {
    margin-bottom: -0.75rem !important;
    width: fit-content !important;
    max-width: fit-content !important;
}
/* Segmented Control */
div[role="radiogroup"] {
    display: inline-flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    width: fit-content !important;
    max-width: 475px !important;
    gap: 0 !important;
    margin-top: 0.5rem;
    margin-bottom: 0 !important;
    background-color: #EBEBEB !important;
    border-radius: 10px !important;
    padding: 3px !important;
}
div[role="radiogroup"] > label {
    background-color: transparent !important;
    border: none !important;
    padding: 7px 18px !important;
    border-radius: 8px !important;
    cursor: pointer;
    white-space: nowrap !important;
    transition: background 0.15s ease;
}
div[role="radiogroup"] > label > div:first-child { display: none !important; }
div[role="radiogroup"] > label:has(input:checked) {
    background-color: #FFFFFF !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.12) !important;
}
div[role="radiogroup"] > label:has(input:checked) p { color: #111111 !important; }
div[role="radiogroup"] > label p { color: #555555 !important; font-size: 0.9rem !important; }
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
</style>
""", unsafe_allow_html=True)

# ── 4. METHODIK TEXT ─────────────────────────────────────────────────────────

METHODIK_TEXT = """
**Info & Quellen**

Die Daten stammen aus dem <a href="https://biel-bienne.mapplus.ch/?lang=de&basemap=stadtplan&blop=1&x=2586000&y=1222000&zl=2&hl=0&layers=e322_grundeigentum_w%7Ce321_gemeindegrenze" target="_blank">öffentlichen Kartenportal (WebGIS) der Stadt Biel</a> (Stand: 26.11.2025). Da die Rohdaten aber rein nur auf geografischen Koordinaten und nicht auf konkreten Adressen basieren, haben wir sie mit einer eigenen Logik neu aufbereitet:

<strong>1. Besitz-Check:</strong> Wir haben für jede Parzelle automatisiert analysiert, wer involviert ist (z. B. wenn die Stadt den Boden besitzt, aber jemand anderes das Baurecht). Diese Daten sind im WebGIS der Stadt hinterlegt.

<strong>2. Daten-Fusion:</strong> Wir haben die geografischen Pläne der Stadt automatisiert mit dem <a href="https://map.geo.admin.ch" target="_blank">Adressregister des Bundes</a> abgeglichen.

<strong>3. Einfachheit:</strong> Bei komplizierten Fällen (wie etwa bei vielen verschiedenen Eigentümern in einem Haus) haben wir die Darstellung vereinfacht, um die Übersichtlichkeit zu wahren.

Aufgrund der Verknüpfung verschiedener Datenquellen sind vereinzelt Fehler möglich. Es ist weiter möglich, dass einzelne Hausnummern in der Datenbank des Bundes hinterlegt sind, an der betroffenen Stelle jedoch (noch) kein Gebäude steht.
"""

def methodik_als_html(text: str) -> str:
    """Konvertiert Markdown-Grundelemente zu HTML."""
    html = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)
    html = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2" target="_blank">\1</a>', html)
    return html.replace('\n', '<br>')


def normalize(s: str) -> str:
    """Ersetzt Umlaute für umlaut-insensitive Suche (Bozingen → Bözingen)."""
    return (str(s).lower()
            .replace('ä', 'a').replace('ö', 'o').replace('ü', 'u')
            .replace('ae', 'a').replace('oe', 'o').replace('ue', 'u'))

def natural_sort_key(s: str) -> str:
    """Zahlen im String auf 6 Stellen nullen-auffüllen → korrekte Sortierung.
    'Strasse 9' < 'Strasse 10' statt 'Strasse 10' < 'Strasse 9'."""
    return re.sub(r'(\d+)', lambda m: m.group(1).zfill(6), str(s).lower())

# Französische Strassentypen stehen fast immer am Wortanfang
_FR_PREFIX = re.compile(
    r'^\s*(rue|avenue|ave\.|place|chemin|route|voie|boulevard|allée|allee|'
    r'impasse|passage|quai|sentier|promenade|grand-rue|côte|cote|pont|'
    r'faubourg|esplanade|ruelle|traverse|rampe|montée|montee|terrasse|'
    r'rond-point|square|jardin|parc)\b',
    re.IGNORECASE,
)

def deutsch_zuerst(adresse: str) -> str:
    """Bei zweisprachigen Adressen (DE / FR) immer den deutschen Teil vorne."""
    if ' / ' not in str(adresse):
        return adresse
    a, b = str(adresse).split(' / ', 1)
    if _FR_PREFIX.match(a) and not _FR_PREFIX.match(b):
        return f"{b} / {a}"
    return adresse

# ── 5. HILFSFUNKTIONEN ───────────────────────────────────────────────────────

def _eigentuemer_code(code_str: str) -> str | None:
    """Findet den Eigentümer-Code (01/02/03) als exakten Token im String."""
    s = str(code_str)
    return next((k for k in EIGENTUEMER if re.search(rf'\b{k}\b', s)), None)

def eigentuemer_dativ(code_str: str) -> str:
    code = _eigentuemer_code(code_str)
    return EIGENTUEMER[code][0] if code else "einem unbekannten Eigentümer"

def eigentuemer_nominativ(code_str: str) -> str:
    code = _eigentuemer_code(code_str)
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
    stadt_boden = any(bool(re.search(r'\b01\b', b)) for b in boden)
    stadt_bau   = any(bool(re.search(r'\b01\b', b)) for b in bau)
    if stadt_boden and not bau:               return "Vollbesitz"
    if stadt_boden and bau and not stadt_bau: return "Bodenbesitz"
    if not stadt_boden and stadt_bau:         return "Gebäudebesitz"
    if stadt_boden and stadt_bau:
        return "Bodenbesitz" if any("01" not in b for b in bau) else "Vollbesitz"
    return "Andere"

def bestimme_kategorien(row) -> frozenset:
    """Returns all applicable filter categories for a row (may be more than one)."""
    besitz    = str(row['Eigentumsverhältnis']).split(" / ")
    nummern   = str(row['Grundstücksnummer(n)']).split(" / ")
    sonderfall = str(row.get('Sonderfall', '')).lower()
    boden, bau = [], []
    for i, b in enumerate(besitz):
        info = nummern[i] if i < len(nummern) else ""
        if "Baurecht" in info:
            bau.append(b)
        elif "Quellenrecht" not in info:
            boden.append(b)
    stadt_boden = any(re.search(r'\b01\b', b) for b in boden)
    stadt_bau   = any(re.search(r'\b01\b', b) for b in bau)
    other_bau   = any(not re.search(r'\b01\b', b) for b in bau)
    bau_split   = "aufgeteilt" in sonderfall
    if not stadt_boden and not stadt_bau:
        return frozenset({"Andere"})
    if not stadt_boden and stadt_bau:
        return frozenset({"Gebäudebesitz"})
    cats: set[str] = set()
    if not bau:
        cats.add("Vollbesitz")
    else:
        if stadt_bau:
            cats.add("Vollbesitz")
        if other_bau or (stadt_bau and bau_split):
            cats.add("Bodenbesitz")
        if not cats:
            cats.add("Bodenbesitz")
    return frozenset(cats)

def generiere_besitz_text(besitz_string: str, nummern_string: str, sonderfall: str = "") -> str:
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
        boden_dativ = eigentuemer_dativ(boden[0]) if boden else "einem unbekannten Eigentümer"
        return (
            f"<strong>QUELLENRECHT</strong><br><br>"
            f"Der Grund und Boden gehört <strong>{boden_dativ}</strong>. "
            f"Jedoch besitzt <strong>{eigentuemer_nominativ(quelle[0])}</strong> "
            f"hier ein Quellenrecht zur Wassernutzung."
        )
    if bau:
        txt_boden_d = unique_d(boden)
        txt_bau_n   = unique_n(bau)
        txt_bau_d   = unique_d(bau)

        boden_codes = {_eigentuemer_code(b) for b in boden} - {None}
        only_stadt_boden = boden_codes == {"01"}

        bau_dativ_list = list(dict.fromkeys(eigentuemer_dativ(b) for b in bau))
        multiple_bau_distinct = len(bau_dativ_list) > 1
        # Multiple bau records of the same code → "mehrere X" (e.g. 03+03 Baurecht)
        multiple_bau_same_code = len(bau) > 1 and not multiple_bau_distinct

        _BAU_PLURAL_D = {
            "01": "der Stadt Biel",
            "02": "mehreren öffentlich-rechtlichen Institutionen",
            "03": "mehreren Privatpersonen oder privaten Firmen",
        }

        if multiple_bau_same_code:
            bau_code = _eigentuemer_code(bau[0])
            plural_d = _BAU_PLURAL_D.get(bau_code, "mehreren Parteien")
            return (
                f"<strong>BAURECHT</strong><br><br>"
                f"Der Grund und Boden gehört <strong>{txt_boden_d}</strong>. "
                f"Das Baurecht (Gebäude) ist im Register <strong>{plural_d}</strong> zugewiesen."
            )

        if txt_boden_d == txt_bau_d:
            if sonderfall and "aufgeteilt" in sonderfall.lower():
                txt_boden_n = unique_n(boden)
                return (
                    f"<strong>BAURECHT</strong><br><br>"
                    f"Der Grund und Boden gehört <strong>{txt_boden_d}</strong>. "
                    f"Das Baurecht (Gebäude) ist aufgeteilt auf <strong>{txt_boden_n}</strong> und Private."
                )
            caveat = (
                " — wobei es sich nicht zwingend um dieselbe Institution handelt"
                if not only_stadt_boden else ""
            )
            return (
                f"<strong>BAURECHT</strong><br><br>"
                f"Der Grund gehört <strong>{txt_boden_d}</strong>. "
                f"Das Baurecht (Gebäude) gehört ebenfalls <strong>{txt_bau_d}</strong>{caveat}."
            )

        if multiple_bau_distinct:
            return (
                f"<strong>BAURECHT</strong><br><br>"
                f"Der Grund gehört <strong>{txt_boden_d}</strong>. "
                f"Das Baurecht (Gebäude) ist aufgeteilt: es gehört "
                f"<strong>{txt_bau_d}</strong>. "
                f"Der Boden bleibt weiterhin im Besitz von <strong>{txt_boden_d}</strong>."
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
            f"<strong>MEHRERE PARZELLEN</strong><br><br>"
            f"Diese Adresse umfasst mehrere separate Parzellen mit verschiedenen Eigentümern. "
            f"Beteiligt sind: <strong>{unique_d(boden)}</strong>."
        )
    if boden:
        return (
            f"<strong>VOLLEIGENTUM</strong><br><br>"
            f"Sowohl der Grund und Boden als auch das darauf stehende Gebäude gehören "
            f"vollumfänglich <strong>{eigentuemer_dativ(boden[0])}</strong>."
        )
    return "Keine Daten verfügbar."


# ── 6. DATEN LADEN ───────────────────────────────────────────────────────────

@st.cache_data
def load_data() -> pd.DataFrame | None:
    if not os.path.exists(EXCEL_FILE):
        st.error(f"Datei nicht gefunden: {EXCEL_FILE}")
        return None
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME).fillna("")
    df.rename(columns={"Unnamed: 4": "Sonderfall"}, inplace=True)
    if "Sonderfall" not in df.columns:
        df["Sonderfall"] = ""
    df['Adresse'] = df['Adresse'].apply(deutsch_zuerst)
    df['Filter_Kategorie']  = df.apply(bestimme_kategorie, axis=1)
    df['Filter_Kategorien'] = df.apply(bestimme_kategorien, axis=1)
    df['Adresse_norm'] = df['Adresse'].apply(normalize)
    return df

@st.cache_data
def load_lottie(path: str) -> dict | None:
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)

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
lottie_data = load_lottie(LOTTIE_FILE)
if lottie_data:
    st_lottie(lottie_data, height=100, loop=True, quality="high", key="logo")

st.markdown("<div class='main-title'>Wie viel Stadt besitzt die Stadt?</div>", unsafe_allow_html=True)
st.markdown("<div class='title-subtext'>Suchportal für den Immobilienbesitz der Stadt Biel</div>", unsafe_allow_html=True)

t1, t2, t3 = st.tabs(["🔍 Suche", "Interaktive Karte", "Info & Quellen"])

# Methodik-Tab rechtsbündig per JS
st.components.v1.html("""
<script>
(function() {
    function moveLastTab() {
        try {
            var doc = window.parent.document;
            var tabList = doc.querySelector('[data-baseweb="tab-list"]');
            if (!tabList) return;
            var tabs = tabList.querySelectorAll('[data-baseweb="tab"]');
            if (tabs.length === 0) return;
            tabs[tabs.length - 1].style.marginLeft = 'auto';
        } catch(e) {}
    }
    [0, 100, 300, 800].forEach(function(t) { setTimeout(moveLastTab, t); });
    try {
        new MutationObserver(moveLastTab)
            .observe(window.parent.document.body, { childList: true, subtree: true });
    } catch(e) {}
})();
</script>
""", height=0)

# ── Tab 1: Suche ─────────────────────────────────────────────────────────────

def clear_search():
    st.session_state.search_input = ""
    st.session_state.filter_mode = FILTER_OPTIONEN[0]
    st.session_state._f_radio = FILTER_OPTIONEN[0]
    st.session_state._f_select = FILTER_OPTIONEN[0]
    st.session_state.page = 1

with t1:
    st.text_input(
        "Suche",
        placeholder="Strasse und Hausnummer",
        label_visibility="collapsed",
        key="search_input",
    )

    col_btn1, col_btn2, _ = st.columns([1, 1, 3])
    with col_btn1:
        st.button("🔍 Suchen", use_container_width=True)
    with col_btn2:
        st.button("✕ Löschen", on_click=clear_search, use_container_width=True)

    # ── Filter: Radio-Pills auf Desktop, Dropdown auf Mobile ─────────────────
    # Beide Widgets vor dem Rendern auf den aktuellen filter_mode synchronisieren —
    # sonst driften sie auseinander wenn sich die Viewport-Breite ändert.
    st.session_state._f_radio  = st.session_state.filter_mode
    st.session_state._f_select = st.session_state.filter_mode

    def _sync_radio():
        st.session_state.filter_mode = st.session_state._f_radio
        st.session_state.page = 1

    def _sync_select():
        st.session_state.filter_mode = st.session_state._f_select
        st.session_state.page = 1

    st.radio(
        "Filter", FILTER_OPTIONEN, horizontal=True,
        label_visibility="collapsed", key="_f_radio", on_change=_sync_radio,
    )
    st.selectbox(
        "Filter", FILTER_OPTIONEN,
        label_visibility="collapsed", key="_f_select", on_change=_sync_select,
    )

    # JavaScript: Selectbox/Radio per Viewport-Breite ein-/ausblenden
    # (robuster als CSS-Media-Queries in Streamlit)
    st.components.v1.html("""
    <script>
    (function() {
        function apply() {
            try {
                var w = window.parent.innerWidth;
                var isMobile = w < 769;
                var doc = window.parent.document;
                doc.querySelectorAll('[data-testid="stRadio"]').forEach(function(el) {
                    el.style.setProperty('display', isMobile ? 'none' : 'block', 'important');
                });
                doc.querySelectorAll('[data-testid="stSelectbox"]').forEach(function(el) {
                    el.style.setProperty('display', isMobile ? 'block' : 'none', 'important');
                });
            } catch(e) {}
        }
        function hideBadge() {
            try {
                var doc = window.parent.document;
                var selectors = [
                    '[data-testid="stToolbar"]',
                    '[data-testid="stStatusWidget"]',
                    '[data-testid="stDeployButton"]',
                    '.stDeployButton',
                    '[class*="viewerBadge"]',
                    '[class*="badge"]'
                ];
                selectors.forEach(function(sel) {
                    doc.querySelectorAll(sel).forEach(function(el) {
                        el.style.setProperty('display', 'none', 'important');
                    });
                });
                // Fixed-position elements in bottom-right corner
                doc.querySelectorAll('body > div').forEach(function(el) {
                    var s = window.parent.getComputedStyle(el);
                    if (s.position === 'fixed' && s.bottom !== 'auto' && s.right !== 'auto') {
                        el.style.setProperty('display', 'none', 'important');
                    }
                });
            } catch(e) {}
        }
        apply();
        hideBadge();
        [100, 300, 700, 1500].forEach(function(t) { setTimeout(apply, t); setTimeout(hideBadge, t); });
        window.parent.addEventListener('resize', apply);

        new MutationObserver(function() {
            setTimeout(apply, 60);
            setTimeout(hideBadge, 60);
        }).observe(window.parent.document.body, {childList: true, subtree: true});
    })();
    </script>
    """, height=0)

    f_mode = st.session_state.filter_mode
    hinweis_key = f_mode if f_mode in FILTER_HINWEISE else "Alle"
    
    st.markdown(
        f"<p style='color:#888888; font-size:0.85rem; margin-top:0.4rem; margin-bottom:20px;'>"
        f"{FILTER_HINWEISE[hinweis_key]}</p>",
        unsafe_allow_html=True,
    )

    # 1. PLATZHALTER & SPINNER SOFORT STARTEN
    search = st.session_state.get("search_input", "")

    # Seite zurücksetzen wenn sich der Suchtext geändert hat
    if search != st.session_state.prev_search:
        st.session_state.page = 1
        st.session_state.prev_search = search

    _slot = st.empty()

    # Zeige den Spinner nur, wenn auch wirklich gefiltert/gesucht wird
    if f_mode != "Alle" or search:
        _slot.markdown("""
        <div style="position:fixed;top:0;left:0;right:0;bottom:0;z-index:99999;
        background:rgba(250,250,250,0.85);display:flex;align-items:center;justify-content:center;">
        <div style="width:48px;height:48px;border-radius:50%;
        border:4px solid #E0E0E0;border-top-color:#111;
        animation:_sp 0.75s linear infinite;"></div>
        <style>@keyframes _sp{to{transform:rotate(360deg)}}</style></div>
        """, unsafe_allow_html=True)
        time.sleep(0.1)  # Kurze Pause, damit Streamlit das HTML an den Browser sendet

    # 2. DATEN FILTERN & SORTIEREN
    f_df = df
    if f_mode != "Alle":
        f_df = f_df[f_df['Filter_Kategorien'].apply(lambda cats: f_mode in cats)]
    if search:
        norm_search = normalize(search)
        pattern = r'\b' + re.escape(norm_search)
        f_df = f_df[f_df['Adresse_norm'].str.contains(
            pattern, case=False, na=False, regex=True
        )]
    f_df = f_df.sort_values('Adresse', key=lambda col: col.map(natural_sort_key))

    # 3. RESULTATE RENDERN (Überschreibt den Spinner im _slot)
    if f_mode == "Alle" and not search:
        _slot.info("Bitte Adresse eingeben oder Filter wählen.")
    elif f_df.empty:
        _slot.info("Keine Treffer.")
    else:
        total       = len(f_df)
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        page        = max(1, min(st.session_state.page, total_pages))
        
        # Pagination-State synchronisieren
        if st.session_state.page != page:
            st.session_state.page = page
            
        start = (page - 1) * PAGE_SIZE
        end   = min(start + PAGE_SIZE, total)

        parts = [f"<div style='margin-bottom:1rem;opacity:0.6;font-size:0.8rem;'>"
                 f"Resultate {start + 1}–{end} von {total}</div>"]
                 
        for _, r in f_df.iloc[start:end].iterrows():
            try:
                eigentuem_clean = re.sub(r'\d{2}:\s*', '', str(r['Eigentumsverhältnis']))
                parzelle   = html.escape(str(r['Grundstücksnummer(n)']))
                eigentum   = html.escape(eigentuem_clean)
                flaeche    = html.escape(str(r['Fläche(n)']))
                adresse    = html.escape(str(r['Adresse']))
                maps_query = urllib.parse.quote(f"{r['Adresse']}, Biel")
                sonderfall = str(r.get('Sonderfall', '')).strip()
                besitz     = generiere_besitz_text(r['Eigentumsverhältnis'], r['Grundstücksnummer(n)'], sonderfall)
                sonderfall_html = (
                    f'<div style="background:#FFF8E1;border-left:3px solid #F59E0B;'
                    f'padding:8px 12px;margin:0 0 10px;border-radius:4px;font-size:0.85rem;line-height:1.5;">'
                    f'⚠️ <strong>Hinweis:</strong> {html.escape(sonderfall)}</div>'
                ) if sonderfall else ''
                parts.append(
                    f'<details class="ei"><summary>{adresse}</summary>'
                    f'<div class="ei-body">{sonderfall_html}{besitz}'
                    f'<a href="https://www.google.com/maps/search/?api=1&query={maps_query}"'
                    f' target="_blank" class="maps-link">📍 Auf Google Maps anzeigen</a>'
                    f'<div class="ei-meta">'
                    f'<div><span class="label-text">Parzelle</span>{parzelle}</div>'
                    f'<div><span class="label-text">Eigentum</span>{eigentum}</div>'
                    f'<div><span class="label-text">Fläche</span>{flaeche}</div>'
                    f'</div></div></details>'
                )
            except Exception:
                pass
                
        # Hier wird der Spinner durch die fertigen HTML-Resultate ersetzt
        _slot.markdown("".join(parts), unsafe_allow_html=True)

        # Pagination-Navigation
        if total_pages > 1:
            st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
            cols = st.columns([1, 2, 1])
            with cols[0]:
                if st.button("← Vorherige", disabled=(page == 1), use_container_width=True):
                    st.session_state.page = page - 1
                    st.rerun()
            with cols[1]:
                st.markdown(
                    f"<div style='text-align:center; padding-top:0.4rem; "
                    f"font-size:0.9rem; color:#555;'>Seite {page} von {total_pages}</div>",
                    unsafe_allow_html=True,
                )
            with cols[2]:
                if st.button("Nächste →", disabled=(page == total_pages), use_container_width=True):
                    st.session_state.page = page + 1
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

# ── Tab 3: Info & Quellen ────────────────────────────────────────────────────
with t3:
    st.markdown(
        f"<div class='methodology-box'>{methodik_als_html(METHODIK_TEXT)}</div>",
        unsafe_allow_html=True,
    )
