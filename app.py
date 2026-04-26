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

EXCEL_FILE   = "Biel_Adressregister_Final.xlsx"
GEOJSON_FILE = "Eigentum.md"
LOTTIE_FILE  = "ajour-logo.json"
GWR_FOLDER   = "GWR_Data"
SHEET_NAME   = "Adress-Verzeichnis"

# GWR: Gebäudekategorie-Labels kürzen
GKAT_KURZ = {
    "Gebäude mit ausschliesslicher Wohnnutzung":          "Wohngebäude",
    "Andere Wohngebäude (Wohngebäude mit Nebennutzung)":  "Wohngebäude (Nebennutzung)",
    "Gebäude mit teilweiser Wohnnutzung":                 "Mischnutzung",
    "Gebäude ohne Wohnnutzung":                           "Gewerbe / Industrie",
    "Sonderbau":                                          "Sonderbau",
    "Provisorische Unterkunft":                           "Provisorische Unterkunft",
}

# GWR: Energieträger-Kategorisierung
_FOSSIL     = {"Gas", "Heizöl"}
_ERNEUERBAR = {"Luft", "Erdwärmesonde", "Erdwärme (generisch)", "Erdregister",
               "Wasser (Grundwasser, Oberflächenwasser, Abwasser)",
               "Sonne (thermisch)", "Holz (generisch)", "Holz (Pellets)",
               "Holz (Stückholz)", "Holz (Schnitzel)", "Abwärme (innerhalb des Gebäudes)",
               "Elektrizität"}
_FERNWAERME = {"Fernwärme (generisch)", "Fernwärme (Hochtemperatur)", "Fernwärme (Niedertemperatur)"}

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
    "Alle Adressen",
    "Vollbesitz (Gebäude & Land)",
    "Bodenbesitz (Baurecht abgegeben)",
    "Gebäudebesitz (Baurecht erhalten)",
]
FILTER_HINWEISE = {
    "Alle Adressen":   "💡 Zeigt das gesamte Register. <strong>Bitte Suchbegriff eingeben.</strong>",
    "Vollbesitz":      "💡 Adressen, bei denen Boden und Gebäude vollständig der Stadt Biel gehören.",
    "Bodenbesitz":     "💡 Die Stadt besitzt das Land, hat es aber an Dritte im Baurecht abgegeben.",
    "Gebäudebesitz":   "💡 Der Boden gehört jemand anderem, aber die Stadt besitzt darauf ein Gebäude im Baurecht.",
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
[data-testid="InputInstructions"] { display: none !important; }
div[data-testid="stExpander"] {
    border-radius: 12px;
    margin-bottom: 1rem;
    background-color: #FFFFFF !important;
    border: 1px solid #EAEAEA !important;
}
div[data-testid="stExpander"] summary {
    padding: 1rem 1.25rem !important;
    min-height: 60px !important;
    display: flex !important;
    align-items: center !important;
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
/* Desktop: nur Radio-Pills, kein Dropdown */
@media (min-width: 769px) {
    [data-testid="stSelectbox"] { display: none !important; }
}
/* Mobile: nur Dropdown, keine Radio-Pills */
@media (max-width: 768px) {
    [data-testid="stRadio"] { display: none !important; }
    [data-testid="stSelectbox"] { display: block !important; }
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
/* Radio-Kreis ausblenden – der schwarze Button reicht als Indikator */
div[role="radiogroup"] > label > div:first-child { display: none !important; }
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


def normalize(s: str) -> str:
    """Ersetzt Umlaute für umlaut-insensitive Suche (Bozingen → Bözingen)."""
    return (str(s).lower()
            .replace('ä', 'a').replace('ö', 'o').replace('ü', 'u')
            .replace('ae', 'a').replace('oe', 'o').replace('ue', 'u'))

def natural_sort_key(s: str) -> str:
    """Zahlen im String auf 6 Stellen nullen-auffüllen → korrekte Sortierung."""
    return re.sub(r'(\d+)', lambda m: m.group(1).zfill(6), str(s).lower())

def _part_to_key(part: str) -> str:
    """Hilfsfunktion: einen Adressteil in einen GWR-Matching-Key umwandeln."""
    part = part.strip()
    tokens = part.rsplit(' ', 1)
    if len(tokens) == 2:
        return normalize(tokens[0]) + '_' + tokens[1].strip().lower()
    return normalize(part) + '_'

def adresse_keys(adresse: str) -> list[str]:
    """Gibt alle möglichen Matching-Keys zurück (DE- und FR-Teil bei zweisprachigen Adressen)."""
    parts = str(adresse).split(' / ')
    return [_part_to_key(p) for p in parts]

def format_baujahr(gbauj, gbaup_label: str) -> str:
    """Gibt Baujahr als lesbaren String zurück."""
    if pd.notna(gbauj) and gbauj > 0:
        return str(int(gbauj))
    label = str(gbaup_label)
    label = re.sub(r'Periode von (\d+) bis (\d+)', r'\1–\2', label)
    label = re.sub(r'Periode vor (\d+)', r'vor \1', label)
    label = re.sub(r'Periode nach (\d+)', r'nach \1', label)
    label = label.replace('Periode ', '')
    return label if label and label != 'nan' else '–'

def energie_icon(genh: str) -> str:
    """Gibt Emoji + Label für einen Energieträger zurück."""
    if genh in _FOSSIL:     return '🔥 ' + genh
    if genh in _FERNWAERME: return '🏭 ' + genh
    if genh in _ERNEUERBAR: return '🌿 ' + genh
    if genh in ('Keine', 'Unbestimmt', '', 'nan'): return '–'
    return '❓ ' + genh

def energie_typ(genh: str) -> str:
    """Kategorisiert einen Energieträger als fossil/erneuerbar/fernwärme/unbekannt."""
    if genh in _FOSSIL:     return 'fossil'
    if genh in _FERNWAERME: return 'fernwärme'
    if genh in _ERNEUERBAR: return 'erneuerbar'
    return 'unbekannt'

# Französische Strassentypen stehen fast immer am Wortanfang
_FR_PREFIX = re.compile(
    r'^\s*(rue|avenue|ave\.|place|chemin|route|voie|boulevard|allée|allee|'
    r'impasse|passage|quai|sentier|promenade|grand-rue|côte|cote)\b',
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
    df['Adresse'] = df['Adresse'].apply(deutsch_zuerst)
    df['Filter_Kategorie'] = df.apply(bestimme_kategorie, axis=1)
    df['Adresse_norm'] = df['Adresse'].apply(normalize)
    return df

@st.cache_data
def load_gwr() -> pd.DataFrame | None:
    """Lädt GWR-Daten aus Split-CSV-Dateien; eine Zeile pro Adresseingang."""
    geb_path  = os.path.join(GWR_FOLDER, "gebaeude_batiment_edificio.csv")
    eing_path = os.path.join(GWR_FOLDER, "eingang_entree_entrata.csv")
    if not os.path.exists(geb_path) or not os.path.exists(eing_path):
        return None

    # Code → Label Lookups (aus kodes_codes_codici.csv, hardcodiert)
    _GKAT_CODE = {
        1010: "Provisorische Unterkunft",
        1020: "Gebäude mit ausschliesslicher Wohnnutzung",
        1030: "Andere Wohngebäude (Wohngebäude mit Nebennutzung)",
        1040: "Gebäude mit teilweiser Wohnnutzung",
        1060: "Gebäude ohne Wohnnutzung",
        1080: "Sonderbau",
    }
    _GBAUP_CODE = {
        8011: "Periode vor 1919",
        8012: "Periode von 1919 bis 1945",
        8013: "Periode von 1946 bis 1960",
        8014: "Periode von 1961 bis 1970",
        8015: "Periode von 1971 bis 1980",
        8016: "Periode von 1981 bis 1985",
        8017: "Periode von 1986 bis 1990",
        8018: "Periode von 1991 bis 1995",
        8019: "Periode von 1996 bis 2000",
        8020: "Periode von 2001 bis 2005",
        8021: "Periode von 2006 bis 2010",
        8022: "Periode von 2011 bis 2015",
        8023: "Periode nach 2015",
    }
    _GENH_CODE = {
        7500: "Keine",
        7501: "Luft",
        7510: "Erdwärme (generisch)",
        7511: "Erdwärmesonde",
        7512: "Erdregister",
        7513: "Wasser (Grundwasser, Oberflächenwasser, Abwasser)",
        7520: "Gas",
        7530: "Heizöl",
        7540: "Holz (generisch)",
        7541: "Holz (Stückholz)",
        7542: "Holz (Pellets)",
        7543: "Holz (Schnitzel)",
        7550: "Abwärme (innerhalb des Gebäudes)",
        7560: "Elektrizität",
        7570: "Sonne (thermisch)",
        7580: "Fernwärme (generisch)",
        7581: "Fernwärme (Hochtemperatur)",
        7582: "Fernwärme (Niedertemperatur)",
        7598: "Unbestimmt",
        7599: "Andere",
    }

    # 1. Gebäudedaten (eine Zeile pro EGID, nur benötigte Spalten)
    geb = pd.read_csv(
        geb_path, sep='\t',
        usecols=['EGID', 'GKAT', 'GBAUP', 'GBAUJ', 'GASTW', 'GANZWHG', 'GENH1', 'GENW1'],
    )
    geb['GKAT_LABEL']  = geb['GKAT'].map(_GKAT_CODE).fillna('')
    geb['GBAUP_LABEL'] = geb['GBAUP'].map(_GBAUP_CODE).fillna('')
    geb['GENH1_LABEL'] = geb['GENH1'].map(_GENH_CODE).fillna('Unbestimmt')
    geb['GENW1_LABEL'] = geb['GENW1'].map(_GENH_CODE).fillna('Unbestimmt')
    geb['WOHN_ANZAHL'] = geb['GANZWHG'].fillna(0).astype(int)

    # 2. Eingänge – nur deutsche Strassennamen (STRSP = 9901)
    eingang = pd.read_csv(
        eing_path, sep='\t',
        usecols=['EGID', 'STRNAME', 'DEINR', 'STRSP'],
    )
    eingang = eingang[eingang['STRSP'] == 9901].drop(columns='STRSP')

    # 3. Join: Eingänge mit Gebäudeinfos verknüpfen (eine Zeile pro Adresseingang)
    merged = eingang.merge(
        geb[['EGID', 'GKAT_LABEL', 'GBAUP_LABEL', 'GBAUJ', 'GASTW',
             'GENH1_LABEL', 'GENW1_LABEL', 'WOHN_ANZAHL']],
        on='EGID', how='left'
    )

    # 4. Berechnete Spalten
    merged['Baujahr']   = merged.apply(
        lambda r: format_baujahr(r['GBAUJ'], r['GBAUP_LABEL']), axis=1
    )
    merged['Kategorie'] = merged['GKAT_LABEL'].map(GKAT_KURZ).fillna(merged['GKAT_LABEL'])
    merged['_key']      = (merged['STRNAME'].apply(normalize) + '_' +
                           merged['DEINR'].astype(str).str.strip().str.lower())
    return merged

@st.cache_data
def prepare_energie_data(df: pd.DataFrame, gwr_df: pd.DataFrame) -> pd.DataFrame:
    """Verknüpft Adressregister (nur Stadtliegenschaften) mit GWR.
    Probiert bei zweisprachigen Adressen beide Teile als Matching-Key."""
    stadt = df[df['Filter_Kategorie'] != 'Andere'].copy()
    gwr_keys = gwr_df[['_key', 'Baujahr', 'Kategorie', 'GENH1_LABEL',
                        'GENW1_LABEL', 'GASTW', 'WOHN_ANZAHL']].copy()

    # Key für den ersten Adressteil (meistens Deutsch)
    stadt['_key1'] = stadt['Adresse'].apply(lambda a: adresse_keys(a)[0])
    # Key für den zweiten Adressteil (meistens Französisch bei umgekehrter Reihenfolge)
    stadt['_key2'] = stadt['Adresse'].apply(
        lambda a: adresse_keys(a)[1] if len(adresse_keys(a)) > 1 else adresse_keys(a)[0]
    )

    # Erster Versuch: Key1
    m1 = stadt.merge(gwr_keys, left_on='_key1', right_on='_key', how='inner')
    # Zweiter Versuch: Key2 für noch nicht gematchte Adressen
    matched = set(m1['Adresse'])
    rest = stadt[~stadt['Adresse'].isin(matched)]
    m2 = rest.merge(gwr_keys, left_on='_key2', right_on='_key', how='inner')

    return pd.concat([m1, m2], ignore_index=True)

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

gwr_df = load_gwr()

if "energie_auth" not in st.session_state:
    st.session_state.energie_auth = False

t1, t2, t3, t4 = st.tabs(["🔍 Suche", "Interaktive Karte", "🧪 Experimentell", "ℹ️ Methodik"])

# ── Tab 1: Suche ─────────────────────────────────────────────────────────────

def clear_search():
    st.session_state.search_input = ""
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
    cur_idx = next(
        (i for i, o in enumerate(FILTER_OPTIONEN) if o == st.session_state.filter_mode), 0
    )

    def _sync_radio():
        st.session_state.filter_mode = st.session_state._f_radio
        st.session_state.page = 1

    def _sync_select():
        st.session_state.filter_mode = st.session_state._f_select
        st.session_state.page = 1

    st.radio(
        "Filter", FILTER_OPTIONEN, index=cur_idx, horizontal=True,
        label_visibility="collapsed", key="_f_radio", on_change=_sync_radio,
    )
    st.selectbox(
        "Filter", FILTER_OPTIONEN, index=cur_idx,
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
        apply();
        [100, 300, 700, 1500].forEach(function(t) { setTimeout(apply, t); });
        window.parent.addEventListener('resize', apply);
        try {
            new MutationObserver(function() { setTimeout(apply, 60); })
                .observe(window.parent.document.body, { childList: true, subtree: true });
        } catch(e) {}
    })();
    </script>
    """, height=1)

    f_mode = st.session_state.filter_mode
    hinweis_key = next((k for k in FILTER_HINWEISE if k in f_mode), "Alle Adressen")
    st.markdown(
        f"<p style='color:#888888; font-size:0.85rem; margin-top:-10px; margin-bottom:20px;'>"
        f"{FILTER_HINWEISE[hinweis_key]}</p>",
        unsafe_allow_html=True,
    )

    search = st.session_state.get("search_input", "")

    # Seite zurücksetzen wenn sich der Suchtext geändert hat
    if search != st.session_state.prev_search:
        st.session_state.page = 1
        st.session_state.prev_search = search

    f_df = df.copy()
    if "Vollbesitz" in f_mode:      f_df = f_df[f_df['Filter_Kategorie'] == "Vollbesitz"]
    elif "Bodenbesitz" in f_mode:   f_df = f_df[f_df['Filter_Kategorie'] == "Bodenbesitz"]
    elif "Gebäudebesitz" in f_mode: f_df = f_df[f_df['Filter_Kategorie'] == "Gebäudebesitz"]
    if search:
        norm_search = normalize(search)
        pattern = r'\b' + re.escape(norm_search)
        f_df = f_df[f_df['Adresse_norm'].str.contains(
            pattern, case=False, na=False, regex=True
        )]
    f_df = f_df.sort_values('Adresse', key=lambda col: col.map(natural_sort_key))

    if f_mode == "Alle Adressen" and not search:
        st.info("Bitte Adresse eingeben oder Filter wählen.")
    elif f_df.empty:
        st.info("Keine Treffer.")
    else:
        total      = len(f_df)
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        page       = max(1, min(st.session_state.page, total_pages))
        start      = (page - 1) * PAGE_SIZE
        end        = min(start + PAGE_SIZE, total)

        # Trefferanzeige
        st.markdown(
            f"<div style='margin-bottom:1rem; opacity:0.6; font-size:0.8rem;'>"
            f"Resultate {start + 1}–{end} von {total}</div>",
            unsafe_allow_html=True,
        )

        # Ergebnisse der aktuellen Seite
        for _, r in f_df.iloc[start:end].iterrows():
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

# ── Tab 3: Experimentell (passwortgeschützt) ──────────────────────────────────
with t3:
    if not st.session_state.energie_auth:
        pw = st.text_input("Passwort", type="password", key="energie_pw")
        if pw == "1312":
            st.session_state.energie_auth = True
            st.rerun()
        elif pw:
            st.error("Falsches Passwort.")
    else:
        if gwr_df is None:
            st.warning("GWR-Daten nicht gefunden (Ordner GWR_Data/ fehlt).")
        else:
            # Gecachter Merge – läuft nur einmal
            merged = prepare_energie_data(df, gwr_df)

            total_geb = len(merged)
            n_fossil  = merged['GENH1_LABEL'].isin(_FOSSIL).sum()
            n_erneu   = merged['GENH1_LABEL'].isin(_ERNEUERBAR).sum()
            n_fern    = merged['GENH1_LABEL'].isin(_FERNWAERME).sum()

            pct = lambda n: f"{n*100//total_geb}%" if total_geb else "–"

            # ── Metriken als HTML-Karten (kein Abschneiden) ──────────────
            st.markdown(f"""
            <div style='display:flex; gap:0.75rem; flex-wrap:wrap; margin-bottom:1.25rem;'>
              <div style='flex:1; min-width:120px; background:#f5f5f5; border-radius:10px; padding:0.9rem 1rem;'>
                <div style='font-size:0.75rem; color:#666; margin-bottom:0.25rem;'>Gebäude erfasst</div>
                <div style='font-size:1.8rem; font-weight:700; line-height:1.1;'>{total_geb}</div>
              </div>
              <div style='flex:1; min-width:120px; background:#f5f5f5; border-radius:10px; padding:0.9rem 1rem;'>
                <div style='font-size:0.75rem; color:#666; margin-bottom:0.25rem;'>🔥 Fossil (Gas/Öl)</div>
                <div style='font-size:1.8rem; font-weight:700; line-height:1.1;'>{n_fossil}</div>
                <div style='font-size:0.85rem; color:#888;'>{pct(n_fossil)}</div>
              </div>
              <div style='flex:1; min-width:120px; background:#f5f5f5; border-radius:10px; padding:0.9rem 1rem;'>
                <div style='font-size:0.75rem; color:#666; margin-bottom:0.25rem;'>🌿 Erneuerbar</div>
                <div style='font-size:1.8rem; font-weight:700; line-height:1.1;'>{n_erneu}</div>
                <div style='font-size:0.85rem; color:#888;'>{pct(n_erneu)}</div>
              </div>
              <div style='flex:1; min-width:120px; background:#f5f5f5; border-radius:10px; padding:0.9rem 1rem;'>
                <div style='font-size:0.75rem; color:#666; margin-bottom:0.25rem;'>🏭 Fernwärme</div>
                <div style='font-size:1.8rem; font-weight:700; line-height:1.1;'>{n_fern}</div>
                <div style='font-size:0.85rem; color:#888;'>{pct(n_fern)}</div>
              </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Filter-Radio ──────────────────────────────────────────────
            e_filter = st.radio(
                "Anzeigen",
                ["Alle", "🔥 Fossil", "🌿 Erneuerbar", "🏭 Fernwärme", "❓ Unbekannt"],
                horizontal=True, label_visibility="collapsed", key="e_filter",
            )

            e_df = merged.copy()
            if e_filter == "🔥 Fossil":
                e_df = e_df[e_df['GENH1_LABEL'].isin(_FOSSIL)]
            elif e_filter == "🌿 Erneuerbar":
                e_df = e_df[e_df['GENH1_LABEL'].isin(_ERNEUERBAR)]
            elif e_filter == "🏭 Fernwärme":
                e_df = e_df[e_df['GENH1_LABEL'].isin(_FERNWAERME)]
            elif e_filter == "❓ Unbekannt":
                e_df = e_df[~e_df['GENH1_LABEL'].isin(_FOSSIL | _ERNEUERBAR | _FERNWAERME)]

            # ── Suchfeld ──────────────────────────────────────────────────
            e_search = st.text_input(
                "Adresse suchen",
                placeholder="Strasse oder Hausnummer...",
                label_visibility="collapsed",
                key="e_search",
            )
            if e_search:
                norm_s = normalize(e_search)
                e_df = e_df[e_df['Adresse'].apply(normalize).str.contains(
                    re.escape(norm_s), case=False, na=False
                )]

            # Sortierung: fossil zuerst, dann älteste Gebäude
            e_df = e_df.sort_values(
                ['GENH1_LABEL', 'Baujahr'],
                key=lambda c: c.map(lambda v: (
                    energie_typ(str(v)) if c.name == 'GENH1_LABEL'
                    else natural_sort_key(str(v))
                ))
            )

            st.markdown(
                f"<div style='margin-bottom:1rem;opacity:0.6;font-size:0.8rem;'>"
                f"{len(e_df)} Liegenschaften</div>",
                unsafe_allow_html=True,
            )

            # ── Tabelle ───────────────────────────────────────────────────
            display = e_df[['Adresse','Filter_Kategorie','Baujahr','GENH1_LABEL','GENW1_LABEL']].copy()
            display.columns = ['Adresse','Eigentum','Baujahr','Heizung','Warmwasser']
            display['Heizung']    = display['Heizung'].apply(energie_icon)
            display['Warmwasser'] = display['Warmwasser'].apply(energie_icon)
            display['Eigentum']   = display['Eigentum'].map({
                'Vollbesitz':    'Vollbesitz',
                'Bodenbesitz':   'Bodenbesitz',
                'Gebäudebesitz': 'Gebäudebesitz',
            })
            st.dataframe(display.reset_index(drop=True), use_container_width=True, hide_index=True)

# ── Tab 4: Methodik ───────────────────────────────────────────────────────────
with t4:
    st.markdown(
        f"<div class='methodology-box'>{methodik_als_html(METHODIK_TEXT)}</div>",
        unsafe_allow_html=True,
    )
