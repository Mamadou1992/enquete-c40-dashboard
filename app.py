# -*- coding: utf-8 -*-
"""
Plateforme de suivi - Enquête C40 GREEN+ · Travailleurs informels des déchets, Dakar
Connexion KoboToolbox (enquête individuelle + grille qualitative) et visualisation.

Lancement :  streamlit run app.py
"""
import io
import json
import os
import random
import re
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import requests
import streamlit as st

# ============================================================================
# Utilitaires Kobo / XLSForm (anciennement kobo_utils.py, fusionné ici pour
# n'avoir qu'un seul fichier Python à déployer)
# ============================================================================

# Version du module : doit correspondre à celle attendue par app.py
VERSION = "2026-09-07"

# --------------------------------------------------------------- Formulaires ---

FORM_QUANTI = {
    "titre": "Enquête individuelle",
    "uid": "ahVcrcrzvdoNYwkFSNjUpj",
    "xlsform": "form_quanti.xlsx",
    # noms alternatifs acceptés (ancien nommage avec accents et espaces)
    "alias": ["Évaluation des besoins des travailleurs informels du secteur des déchets - Dakar.xlsx"],
    "motif": "valuation des besoins",
}
FORM_QUALI = {
    "titre": "Entretiens & focus groups",
    "uid": "aotuV7h5fdZe8LDS3KF6KP",
    "xlsform": "form_quali.xlsx",
    "alias": ["Grille de saisie qualitative (entretiens et focus groups).xlsx"],
    "motif": "qualitative",
}


def trouver_xlsform(dossier: str, config: dict) -> str:
    """Retourne le chemin du XLSForm : nom principal, alias, puis recherche par motif.
    Robuste aux différences d'encodage des noms de fichiers (Windows / Linux)."""
    candidats = [config["xlsform"]] + list(config.get("alias", []))
    for nom in candidats:
        chemin = os.path.join(dossier, nom)
        if os.path.exists(chemin):
            return chemin
    motif = config.get("motif", "").lower()
    if motif:
        for f in sorted(os.listdir(dossier)):
            if f.lower().endswith((".xlsx", ".xls")) and motif in f.lower():
                return os.path.join(dossier, f)
    raise FileNotFoundError(
        f"XLSForm introuvable pour « {config['titre']} ». Attendu : {config['xlsform']} "
        f"dans {dossier}. Fichiers présents : "
        + ", ".join(f for f in sorted(os.listdir(dossier)) if f.lower().endswith(".xlsx"))
    )

# Colonnes clés du formulaire quantitatif
COL = {
    "commune": "commune",
    "enqueteur": "enqueteur",
    "gps": "gps",
    "consent": "consent",
    "sexe": "sexe",
    "age": "age",
    "migratoire": "migratoire",
    "fonction": "fonction",
}

# Champs exclus de l'analyse (identifiants / données personnelles)
CHAMPS_EXCLUS = {"contact", "enqueteur", "animateur", "quartier", "gps",
                 "org_nom", "commune_lieu"}


# ------------------------------------------------------------------ XLSForm ---

def norm_key(s: str) -> str:
    """Normalise un nom de commune (sans accents, majuscules, sans ponctuation)."""
    s = str(s).replace("œ", "oe").replace("Œ", "OE")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", " ", s).strip().upper()
    return re.sub(r"\s+", " ", s)

# Correspondance label du formulaire -> clé du GeoJSON (cas non triviaux)
COMMUNE_ALIAS = {"DAKAR PLATEAU": "PLATEAU"}


def commune_to_key(label: str) -> str:
    k = norm_key(label)
    return COMMUNE_ALIAS.get(k, k)


def _col_label(df: pd.DataFrame) -> str:
    """Nom de la colonne des libellés ('label' ou 'label::Français (fr)')."""
    for c in df.columns:
        if str(c).strip().lower() == "label":
            return c
    for c in df.columns:
        if str(c).strip().lower().startswith("label"):
            return c
    return ""


def load_xlsform(path: str) -> dict:
    """Lit un XLSForm et retourne types, libellés, groupes (imbriqués) et choix."""
    survey = pd.read_excel(path, sheet_name="survey").fillna("")
    choices = pd.read_excel(path, sheet_name="choices").fillna("")
    lab_s, lab_c = _col_label(survey), _col_label(choices)
    if lab_s and lab_s != "label":
        survey = survey.rename(columns={lab_s: "label"})
    if lab_c and lab_c != "label":
        choices = choices.rename(columns={lab_c: "label"})
    if "label" not in survey:
        survey["label"] = ""
    if "label" not in choices:
        choices["label"] = ""

    qtype, label, listname, group_of, groups, ordre = {}, {}, {}, {}, {}, []
    stack = []
    for _, r in survey.iterrows():
        t = str(r["type"]).strip()
        name = str(r["name"]).strip()
        lab = str(r.get("label", "")).strip()
        if not t:
            continue
        if t == "begin_group":
            stack.append(name)
            if lab:  # les groupes techniques sans libellé ne sont pas des sections
                groups[name] = lab
                ordre.append(name)
            continue
        if t == "end_group":
            if stack:
                stack.pop()
            continue
        if not name:
            continue
        qtype[name] = t
        label[name] = lab or name
        # section = groupe libellé le plus proche dans la pile
        section = ""
        for g in reversed(stack):
            if g in groups:
                section = g
                break
        group_of[name] = section
        m = re.match(r"select_(one|multiple)\s+(\S+)", t)
        if m:
            listname[name] = m.group(2)

    choice_map = {}
    for _, r in choices.iterrows():
        ln = str(r["list_name"]).strip()
        code = str(r["name"]).strip()
        lab = str(r["label"]).strip()
        if ln:
            choice_map.setdefault(ln, {})[code] = lab.strip(" ,")

    return {"qtype": qtype, "label": label, "listname": listname,
            "group_of": group_of, "groups": groups, "ordre_groupes": ordre,
            "choices": choice_map}


# -------------------------------------------------------------------- Kobo ----

def kobo_headers(token: str) -> dict:
    return {"Authorization": f"Token {token.strip()}"}


def list_assets(base_url: str, token: str) -> list:
    """Liste les formulaires déployés du compte (uid, nom, nb soumissions)."""
    url = f"{base_url.rstrip('/')}/api/v2/assets.json?q=asset_type:survey&limit=100"
    r = requests.get(url, headers=kobo_headers(token), timeout=60)
    r.raise_for_status()
    out = []
    for a in r.json().get("results", []):
        if a.get("has_deployment"):
            out.append({"uid": a["uid"], "name": a["name"],
                        "count": a.get("deployment__submission_count", 0)})
    return out


def fetch_submissions(base_url: str, token: str, asset_uid: str) -> list:
    """Télécharge toutes les soumissions d'un formulaire (paginé)."""
    url = f"{base_url.rstrip('/')}/api/v2/assets/{asset_uid}/data.json?limit=1000"
    records = []
    while url:
        r = requests.get(url, headers=kobo_headers(token), timeout=120)
        r.raise_for_status()
        js = r.json()
        records.extend(js.get("results", []))
        url = js.get("next")
    return records


# ---------------------------------------------------------- Mise en forme -----

def to_dataframe(records: list, form: dict) -> pd.DataFrame:
    """Aplatit les soumissions Kobo : retire les préfixes de groupes, décode les
    select_one en libellés, extrait GPS et dates."""
    if not records:
        return pd.DataFrame()
    rows = []
    for rec in records:
        row = {}
        for k, v in rec.items():
            row[k.split("/")[-1]] = v
        rows.append(row)
    df = pd.DataFrame(rows)

    # Dates
    if "_submission_time" in df:
        df["_submission_time"] = pd.to_datetime(df["_submission_time"], errors="coerce")
        df["date_soumission"] = df["_submission_time"].dt.date

    # GPS : champ geopoint "lat lon alt acc", sinon _geolocation
    lat, lon = [], []
    for _, r in df.iterrows():
        la = lo = None
        v = r.get(COL["gps"])
        if isinstance(v, str) and v.strip():
            p = v.split()
            if len(p) >= 2:
                try:
                    la, lo = float(p[0]), float(p[1])
                except ValueError:
                    pass
        if la is None and isinstance(r.get("_geolocation"), list):
            g = r["_geolocation"]
            if len(g) == 2 and g[0] is not None:
                la, lo = g[0], g[1]
        lat.append(la)
        lon.append(lo)
    df["latitude"], df["longitude"] = lat, lon

    # Décodage des select_one (les select_multiple restent en codes)
    for name, t in form["qtype"].items():
        if name in df.columns and t.startswith("select_one"):
            cmap = form["choices"].get(form["listname"].get(name, ""), {})
            df[name] = df[name].map(lambda x: cmap.get(x, x) if isinstance(x, str) else x)

    # Numériques
    for name, t in form["qtype"].items():
        if name in df.columns and t in ("integer", "decimal"):
            df[name] = pd.to_numeric(df[name], errors="coerce")

    return df


def dataframe_vide(form: dict) -> pd.DataFrame:
    """DataFrame sans aucune ligne mais avec toutes les colonnes du formulaire.
    Permet d'afficher le tableau de bord (compteurs à zéro) avant la collecte."""
    cols = list(form["qtype"]) + ["_id", "_submission_time", "date_soumission",
                                  "latitude", "longitude"]
    df = pd.DataFrame({c: pd.Series(dtype="object") for c in dict.fromkeys(cols)})
    for name, t in form["qtype"].items():
        if t in ("integer", "decimal"):
            df[name] = pd.Series(dtype="float")
    df["latitude"] = pd.Series(dtype="float")
    df["longitude"] = pd.Series(dtype="float")
    return df


def explode_multiple(df: pd.DataFrame, col: str, form: dict) -> pd.Series:
    """Compte les réponses d'un select_multiple (codes séparés par des espaces)."""
    cmap = form["choices"].get(form["listname"].get(col, ""), {})
    vals = []
    for v in df[col].dropna():
        if isinstance(v, list):
            v = " ".join(str(x) for x in v)
        if isinstance(v, str):
            for code in v.split():
                lab = str(cmap.get(code, code)).strip()
                if lab:
                    vals.append(lab)
    return pd.Series(vals).value_counts()


# ------------------------------------------- Plan d'échantillonnage C40 -------
# Source : Plan d'échantillonnage stratifié par commune (ANSD 2023, 400 enquêtes)
PLAN_TOTAL = 400
PLAN_QUOTAS = {  # clé GeoJSON -> (population ANSD 2023, quota d'enquêtes)
    "GRAND YOFF": (186775, 58),
    "PARCELLES ASSAINIES": (167671, 52),
    "YOFF": (119351, 37),
    "OUAKAM": (100541, 31),
    "HANN BEL AIR": (86908, 27),
    "MEDINA": (82544, 26),
    "BISCUITERIE": (74025, 23),
    "CAMBERENE": (71432, 22),
    "GUEULE TAPEE FASS COLOBANE": (59227, 19),
    "GRAND DAKAR": (47334, 15),
    "PATTE D OIE": (46821, 15),
    "HLM": (42975, 13),
    "SICAP LIBERTE": (41079, 13),
    "DIEUPPEUL DERKLE": (38725, 12),
    "MERMOZ SACRE COEUR": (38598, 12),
    "PLATEAU": (34951, 11),
    "FANN POINT E AMITIE": (20115, 6),
    "NGOR": (17706, 6),
    "GOREE": (1691, 2),
}
# ------------------------------------------------ Plan qualitatif prévu ------
# Nombre de séances / entretiens attendus pour chaque item.
# Modifiez ces valeurs si le plan de collecte qualitative évolue.
PLAN_QUALI = {
    "groupe_fg": 1,        # 1 séance par focus group (FG1 à FG6)
    "profil_ind": 2,       # 2 entretiens par profil de répondant
    "structure_inst": 1,   # 1 entretien par type de structure
}
PLAN_QUALI_LIBELLES = {
    "groupe_fg": "Focus groups",
    "profil_ind": "Entretiens individuels (par profil)",
    "structure_inst": "Entretiens institutionnels (par structure)",
}


def suivi_plan_quali(dfq: pd.DataFrame, form_q: dict, champ: str) -> pd.DataFrame:
    """Compare le réalisé à la cible pour un champ du plan qualitatif."""
    cibles = form_q["choices"].get(form_q["listname"].get(champ, ""), {})
    attendu = PLAN_QUALI.get(champ, 1)
    realise = (dfq[champ].astype(str).value_counts()
               if (champ in dfq.columns and not dfq.empty) else pd.Series(dtype=int))
    lignes = []
    for libelle in cibles.values():
        lib = str(libelle).strip()
        fait = int(realise.get(lib, 0))
        lignes.append({"Item": lib, "Réalisé": fait, "Prévu": attendu,
                       "Reste": max(attendu - fait, 0),
                       "Statut": "✅ Fait" if fait >= attendu else
                                 ("🟡 En cours" if fait else "⬜ À faire")})
    return pd.DataFrame(lignes)


# Quotas transversaux (part minimale attendue) et par fonction
PLAN_TRANSVERSAL = {"femmes": 0.40, "jeunes": 0.25, "migrants": 0.15}
PLAN_FONCTIONS = {"Charretier de pré-collecte": 0.60, "Récupérateur": 0.20,
                  "Trieur": 0.15, "Autre": 0.05}
AGE_JEUNE_MAX = 35  # « jeunes » = moins de 35 ans


# ------------------------------------------------------- Données de démo ------

def generate_demo_data(form: dict, n: int = 150, seed: int = 42,
                       quali: bool = False) -> list:
    """Simule des soumissions Kobo pour tester l'app sans connexion."""
    rng = random.Random(seed)
    ch, ln = form["choices"], form["listname"]
    centres = {
        "c1": (14.667, -17.433), "c2": (14.678, -17.451), "c3": (14.685, -17.448),
        "c4": (14.693, -17.465), "c5": (14.667, -17.398), "c6": (14.717, -17.425),
        "c7": (14.703, -17.452), "c8": (14.717, -17.455), "c9": (14.708, -17.447),
        "c10": (14.713, -17.443), "c11": (14.712, -17.459), "c12": (14.712, -17.472),
        "c13": (14.735, -17.452), "c14": (14.735, -17.44), "c15": (14.755, -17.435),
        "c16": (14.765, -17.43), "c17": (14.75, -17.47), "c18": (14.748, -17.512),
        "c19": (14.722, -17.49),
    }
    codes_communes = list(ch.get("communes", {}))
    recs = []
    t0 = datetime.now() - timedelta(days=21)
    for i in range(n):
        ts = t0 + timedelta(days=rng.uniform(0, 21))
        rec = {"_id": i + 1, "_submission_time": ts.isoformat()}
        if not quali and codes_communes:
            idx = rng.randrange(len(codes_communes))
            la, lo = list(centres.values())[idx % len(centres)]
            la += rng.uniform(-0.008, 0.008)
            lo += rng.uniform(-0.008, 0.008)
            rec["_geolocation"] = [la, lo]
            rec["gps"] = f"{la} {lo} 0 5"
            rec["commune"] = codes_communes[idx]
        for name, t in form["qtype"].items():
            if name in rec:
                continue
            codes = list(ch.get(ln.get(name, ""), {}))
            if t.startswith("select_one") and codes:
                rec[name] = rng.choice(codes)
            elif t.startswith("select_multiple") and codes:
                rec[name] = " ".join(rng.sample(codes, k=min(len(codes), rng.randint(1, 3))))
            elif t == "integer":
                if name == "age":
                    rec[name] = rng.randint(18, 65)
                elif name == "revenu":
                    rec[name] = max(15000, int(rng.gauss(85000, 35000)))
                elif name.startswith("nb_"):
                    rec[name] = rng.randint(0, 12)
                else:
                    rec[name] = rng.randint(1, 30)
            elif t == "text" and rng.random() < 0.5:
                rec[name] = f"Réponse libre {i + 1}"
        recs.append(rec)
    return recs


# Palette verte de la plateforme
VERTS = ["#1b5e20", "#66bb6a", "#9ccc65", "#2e7d32", "#a5d6a7", "#33691e"]
px.defaults.color_discrete_sequence = VERTS
px.defaults.color_continuous_scale = "Greens"

APP_DIR = os.path.dirname(os.path.abspath(__file__))
GEOJSON = os.path.join(APP_DIR, "communes_dakar.geojson")
LOGO = os.path.join(APP_DIR, "logo_sgp.jpeg")
DEFAULT_URL = "https://kf.kobotoolbox.org"

st.set_page_config(page_title="Enquête C40 GREEN+ Dakar",
                   page_icon=LOGO if os.path.exists(LOGO) else "♻️", layout="wide")

# Habillage vert des composants (puces, encadrés, onglets, indicateurs)
st.markdown("""
<style>
[data-baseweb="tag"] { background-color: #1b5e20 !important; }
[data-baseweb="tag"] span { color: #ffffff !important; }
[data-baseweb="tag"] svg { fill: #ffffff !important; }

[data-baseweb="popover"] li:hover,
[data-baseweb="popover"] li[aria-selected="true"] { background-color: #f1f8e9 !important; }

div[data-testid="stAlert"] {
  background-color: #f1f8e9 !important;
  border-left: 5px solid #9ccc65 !important;
  border-radius: 8px;
}
div[data-testid="stAlert"] p { color: #12291a !important; }

button[data-baseweb="tab"][aria-selected="true"] { color: #1b5e20 !important; font-weight: 600; }
div[data-baseweb="tab-highlight"] { background-color: #1b5e20 !important; }

div[data-testid="stMetric"] {
  background-color: #f1f8e9;
  border-left: 5px solid #9ccc65;
  border-radius: 10px;
  padding: 12px 16px;
}

details summary { background-color: #f1f8e9 !important; border-radius: 8px; }

div[data-testid="stDownloadButton"] button {
  background-color: #1b5e20 !important; color: #ffffff !important; border: none;
}
div[data-testid="stDownloadButton"] button:hover { background-color: #0d3f14 !important; }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------- Authentification ----

MOT_DE_PASSE_DEFAUT = "senegalgreen+"


def verifier_acces():
    """Affiche un écran de connexion tant que le mot de passe n'est pas saisi."""
    attendu = st.secrets.get("acces", {}).get("mot_de_passe", MOT_DE_PASSE_DEFAUT)
    if st.session_state.get("authentifie"):
        return

    _, centre, _ = st.columns([1, 2, 1])
    with centre:
        if os.path.exists(LOGO):
            g, d = st.columns([1, 3])
            g.image(LOGO, width=110)
            d.markdown("### Enquête C40 GREEN+")
        else:
            st.markdown("### ♻️ Enquête C40 GREEN+")
        st.caption("Plateforme de suivi - Travailleurs informels des déchets, Dakar")
        saisie = st.text_input("Mot de passe", type="password",
                               placeholder="Saisissez le mot de passe d'accès")
        if st.button("Se connecter", width="stretch"):
            if saisie == attendu:
                st.session_state["authentifie"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
    st.stop()


verifier_acces()

# ----------------------------------------------------------- Chargements ----

@st.cache_resource
def get_form(cle):
    config = FORM_QUANTI if cle == "quanti" else FORM_QUALI
    return load_xlsform(trouver_xlsform(APP_DIR, config))


@st.cache_resource
def get_geojson():
    with open(GEOJSON, encoding="utf-8") as f:
        return json.load(f)


@st.cache_data(ttl=300, show_spinner="Téléchargement des données Kobo…")
def get_data(base_url, token, uid):
    return fetch_submissions(base_url, token, uid)


try:
    form = get_form("quanti")
    form_q = get_form("quali")
    geojson = get_geojson()
except FileNotFoundError as e:
    st.error(f"Fichier manquant : {e}")
    st.info("Vérifiez que `form_quanti.xlsx`, `form_quali.xlsx` et "
            "`communes_dakar.geojson` sont bien présents à côté de `app.py`.")
    st.stop()

# --------------------------------------------------------------- Sidebar ----

if os.path.exists(LOGO):
    st.sidebar.image(LOGO, width="stretch")
st.sidebar.title("Enquête C40 GREEN+")
st.sidebar.caption("Travailleurs informels des déchets - Département de Dakar")

base_url = st.secrets.get("kobo", {}).get("base_url", DEFAULT_URL)
token = st.secrets.get("kobo", {}).get("token", "")
uid = st.secrets.get("kobo", {}).get("asset_uid", FORM_QUANTI["uid"])
uid_q = st.secrets.get("kobo", {}).get("asset_uid_quali", FORM_QUALI["uid"])

with st.sidebar.expander("⚙️ Connexion Kobo", expanded=not token):
    base_url = st.text_input("Serveur", base_url)
    token = st.text_input("Jeton API", token, type="password",
                          help="Compte Kobo → Account Settings → Security → API Key")
    st.caption(f"Enquête individuelle : `{uid}`")
    st.caption(f"Qualitatif : `{uid_q}`")

if st.sidebar.button("🔄 Actualiser les données"):
    get_data.clear()
    st.rerun()

if st.sidebar.button("🔒 Se déconnecter"):
    st.session_state["authentifie"] = False
    st.rerun()

# ------------------------------------------------------------- Données ------

if not token:
    st.warning("Renseignez le jeton API Kobo dans la barre latérale.")
    st.stop()
try:
    records = get_data(base_url, token, uid)
except Exception as e:
    st.error(f"Erreur API Kobo (enquête individuelle) : {e}")
    st.stop()

try:
    records_q = get_data(base_url, token, uid_q)
except Exception:
    records_q = []

df = to_dataframe(records, form)
dfq = to_dataframe(records_q, form_q)

# Avant la collecte : tableau de bord complet avec des compteurs à zéro
collecte_vide = df.empty
if collecte_vide:
    df = dataframe_vide(form)
    st.info("**La collecte n'a pas encore démarré.** Le tableau de bord ci-dessous est "
            "prêt : compteurs à zéro, plan d'échantillonnage affiché. Il se remplira "
            "automatiquement dès les premières soumissions Kobo (bouton « 🔄 Actualiser » "
            "dans la barre latérale).")

C = COL

# --------------------------------------------------------------- Filtres ----

st.sidebar.header("Filtres")
fdf = df.copy()

if "date_soumission" in fdf and fdf["date_soumission"].notna().any():
    dmin, dmax = fdf["date_soumission"].min(), fdf["date_soumission"].max()
    d1, d2 = st.sidebar.date_input("Période", (dmin, dmax), min_value=dmin, max_value=dmax)
    fdf = fdf[(fdf["date_soumission"] >= d1) & (fdf["date_soumission"] <= d2)]


def sidebar_filter(col, titre):
    global fdf
    if col in fdf and fdf[col].notna().any():
        sel = st.sidebar.multiselect(titre, sorted(fdf[col].dropna().astype(str).unique()))
        if sel:
            fdf = fdf[fdf[col].astype(str).isin(sel)]


sidebar_filter(C["commune"], "Commune d'enquête")
sidebar_filter(C["enqueteur"], "Enquêteur")
sidebar_filter(C["fonction"], "Fonction")
sidebar_filter(C["sexe"], "Sexe")

st.sidebar.metric("Enquêtes affichées", f"{len(fdf)} / {len(df)}")

# ----------------------------------------------------------------- Corps ----

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["📊 Suivi des enquêtes", "🗺️ Carte", "📈 Analyse thématique",
     "🗣️ Qualitatif", "📥 Données & export"])

# ============================================== 1. SUIVI DES ENQUÊTES ========
with tab1:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Enquêtes réalisées", len(fdf))
    nb_com = fdf[C["commune"]].nunique() if C["commune"] in fdf else 0
    c2.metric("Communes couvertes", f"{nb_com} / 19")
    nb_enq = fdf[C["enqueteur"]].nunique() if C["enqueteur"] in fdf else 0
    c3.metric("Enquêteurs actifs", nb_enq)
    if C["consent"] in fdf:
        cons = (fdf[C["consent"]].astype(str).str.lower().str.startswith("oui").mean()
                if len(fdf) else 0.0)
        c4.metric("Taux de consentement", f"{cons:.0%}")

    prog = len(fdf) / PLAN_TOTAL
    st.progress(min(prog, 1.0), text=f"Progression du plan d'échantillonnage : "
                f"{len(fdf)} / {PLAN_TOTAL} enquêtes ({prog:.0%})")

    st.markdown("**Quotas transversaux du plan C40**")
    q1, q2, q3 = st.columns(3)
    p = (fdf[C["sexe"]].astype(str) == "Femme").mean() if (C["sexe"] in fdf and len(fdf)) else 0.0
    q1.metric("Femmes (cible >= 40 %)", f"{p:.0%}", delta=f"{(p - 0.40) * 100:+.0f} pts")
    ages = pd.to_numeric(fdf[C["age"]], errors="coerce").dropna() if C["age"] in fdf else pd.Series(dtype=float)
    p = (ages < AGE_JEUNE_MAX).mean() if len(ages) else 0.0
    q2.metric(f"Jeunes < {AGE_JEUNE_MAX} ans (cible >= 25 %)", f"{p:.0%}",
              delta=f"{(p - 0.25) * 100:+.0f} pts")
    p = (fdf[C["migratoire"]].astype(str).str.contains("migrant", case=False).mean()
         if (C["migratoire"] in fdf and len(fdf)) else 0.0)
    q3.metric("Migrants (cible >= 15 %)", f"{p:.0%}", delta=f"{(p - 0.15) * 100:+.0f} pts")

    g1, g2 = st.columns(2)
    with g1:
        if "date_soumission" in fdf and fdf["date_soumission"].notna().any():
            par_jour = fdf.groupby("date_soumission").size().reset_index(name="Enquêtes")
            fig = px.area(par_jour, x="date_soumission", y="Enquêtes", markers=True,
                          title="Évolution des soumissions par jour",
                          labels={"date_soumission": "Date"})
            st.plotly_chart(fig, width="stretch")
    with g2:
        # Quotas par fonction (plan : 60 % charretiers, 20 % récupérateurs, 15 % trieurs)
        if C["fonction"] in fdf:
            obs = (fdf[C["fonction"]].value_counts(normalize=True) * 100
                   if fdf[C["fonction"]].notna().any() else pd.Series(dtype=float))
            lignes = []
            for fonction, part in PLAN_FONCTIONS.items():
                lignes.append({"Fonction": fonction, "Part": obs.get(fonction, 0.0),
                               "Type": "Réalisé"})
                lignes.append({"Fonction": fonction, "Part": part * 100, "Type": "Cible"})
            d = pd.DataFrame(lignes)
            fig = px.bar(d, x="Part", y="Fonction", color="Type", barmode="group",
                         orientation="h", title="Quotas par fonction (% des enquêtes)",
                         color_discrete_sequence=["#1b5e20", "#9ccc65"])
            st.plotly_chart(fig, width="stretch")

    g3, g4 = st.columns(2)
    with g3:
        if C["commune"] in fdf and fdf[C["commune"]].notna().any():
            par_com = fdf[C["commune"]].value_counts().reset_index()
            par_com.columns = ["Commune", "Enquêtes"]
            fig = px.bar(par_com.sort_values("Enquêtes"), x="Enquêtes", y="Commune",
                         orientation="h", title="Enquêtes par commune", text_auto=True)
            fig.update_layout(height=520, yaxis_title="")
            st.plotly_chart(fig, width="stretch")
    with g4:
        if C["enqueteur"] in fdf and fdf[C["enqueteur"]].notna().any():
            par_enq = fdf[C["enqueteur"]].value_counts().reset_index()
            par_enq.columns = ["Enquêteur", "Enquêtes"]
            fig = px.bar(par_enq, x="Enquêteur", y="Enquêtes", text_auto=True,
                         title="Enquêtes par enquêteur")
            fig.update_layout(height=520)
            st.plotly_chart(fig, width="stretch")

# ================================================================ 2. CARTE ===
with tab2:
    left, right = st.columns([3, 1])
    with right:
        fond = st.radio("Affichage", ["Points GPS", "Densité par commune", "Les deux"], index=2)
        choix_ind = ["Enquêtes réalisées", "Taux de réalisation (%)",
                     "Quota du plan", "Population ANSD 2023"]
        indicateur = st.radio("Couleur des communes", choix_ind,
                              index=2 if collecte_vide else 0)

    counts = pd.DataFrame({"key": [], "Enquêtes réalisées": []})
    if C["commune"] in fdf:
        cc = fdf[C["commune"]].dropna().map(commune_to_key).value_counts()
        counts = cc.reset_index()
        counts.columns = ["key", "Enquêtes réalisées"]
    all_keys = pd.DataFrame({"key": [f["properties"]["key"] for f in geojson["features"]],
                             "Commune": [f["properties"]["commune"] for f in geojson["features"]]})
    counts = all_keys.merge(counts, on="key", how="left").fillna({"Enquêtes réalisées": 0})
    counts["Population ANSD 2023"] = counts["key"].map(lambda k: PLAN_QUOTAS.get(k, (0, 0))[0])
    counts["Quota du plan"] = counts["key"].map(lambda k: PLAN_QUOTAS.get(k, (0, 0))[1])
    counts["Taux de réalisation (%)"] = (counts["Enquêtes réalisées"]
                                         / counts["Quota du plan"].replace(0, pd.NA) * 100).round(0)
    counts["Restant"] = (counts["Quota du plan"] - counts["Enquêtes réalisées"]).clip(lower=0)

    with left:
        try:
            if fond in ("Densité par commune", "Les deux"):
                fig = px.choropleth_map(counts, geojson=geojson, locations="key",
                                        featureidkey="properties.key", color=indicateur,
                                        hover_name="Commune", color_continuous_scale="Greens",
                                        hover_data={"key": False, "Quota du plan": True,
                                                    "Enquêtes réalisées": True,
                                                    "Taux de réalisation (%)": True},
                                        center={"lat": 14.716, "lon": -17.45}, zoom=10.6,
                                        opacity=0.6, height=650)
            else:
                fig = px.choropleth_map(counts.assign(z=0), geojson=geojson, locations="key",
                                        featureidkey="properties.key", color="z",
                                        hover_name="Commune",
                                        color_continuous_scale=["#eee", "#eee"],
                                        center={"lat": 14.716, "lon": -17.45}, zoom=10.6,
                                        opacity=0.25, height=650)
                fig.update_coloraxes(showscale=False)
            gps = fdf.dropna(subset=["latitude", "longitude"])
            if fond in ("Points GPS", "Les deux") and len(gps):
                hover = gps.get(C["commune"], pd.Series([""] * len(gps), index=gps.index))
                fig.add_scattermap(lat=gps["latitude"], lon=gps["longitude"], mode="markers",
                                   marker={"size": 9, "color": "#d62728"}, text=hover,
                                   name="Enquêtes", hovertemplate="%{text}<extra></extra>")
            fig.update_layout(margin={"l": 0, "r": 0, "t": 0, "b": 0}, map_style="carto-positron")
            st.plotly_chart(fig, width="stretch")
        except Exception as e:
            st.error(f"Erreur d'affichage de la carte : {e}")

    st.caption("Suivi du plan d'échantillonnage (400 enquêtes réparties au poids "
               "démographique ANSD 2023)")
    st.dataframe(
        counts[["Commune", "Population ANSD 2023", "Quota du plan", "Enquêtes réalisées",
                "Taux de réalisation (%)", "Restant"]].sort_values("Quota du plan",
                                                                   ascending=False),
        width="stretch", hide_index=True)

# ================================================== 3. ANALYSE THÉMATIQUE ====


def afficher_questions(data, formulaire, questions, cle):
    """Affiche le graphique adapté à chaque question (anneau, barres, treemap,
    histogramme) ou la liste des réponses libres."""
    rendu = False
    for q in questions:
        t = formulaire["qtype"][q]
        titre = formulaire["label"].get(q, q)
        reponses = data[q].dropna()
        reponses = reponses[reponses.astype(str).str.strip() != ""]
        if reponses.empty:
            st.caption(f"◽ {titre} : aucune réponse exploitable pour le moment")
            continue
        rendu = True
        if t == "text":
            with st.expander(f"📝 {titre} - {len(reponses)} réponse(s)"):
                d = reponses.astype(str).reset_index(drop=True)
                d.index += 1
                st.dataframe(d.rename("Réponse"), width="stretch")
            continue
        if t.startswith("select_multiple"):
            s = explode_multiple(data, q, formulaire)
            if s.empty:
                st.caption(f"◽ {titre} : aucune réponse exploitable pour le moment")
                continue
            d = s.reset_index()
            d.columns = ["Réponse", "Nombre"]
            fig = px.treemap(d, path=["Réponse"], values="Nombre",
                             title=f"{titre} (choix multiples, n={len(reponses)})")
            fig.update_traces(textinfo="label+value+percent root")
            fig.update_layout(height=420, margin={"t": 60, "l": 10, "r": 10, "b": 10})
        elif t.startswith("select_one"):
            s = reponses.astype(str).str.strip().value_counts()
            d = s.reset_index()
            d.columns = ["Réponse", "Nombre"]
            n = int(d["Nombre"].sum())
            if len(d) <= 5:
                fig = px.pie(d, names="Réponse", values="Nombre", hole=0.45,
                             title=f"{titre} (n={n})")
                fig.update_traces(textinfo="percent+value")
                fig.update_layout(height=380)
            else:
                d = d.sort_values("Nombre")
                d["pct"] = (d["Nombre"] / n * 100).map(lambda v: f"{v:.0f} %")
                fig = px.bar(d, x="Nombre", y="Réponse", orientation="h",
                             title=f"{titre} (n={n})", text="pct")
                fig.update_traces(textposition="outside", cliponaxis=False)
                fig.update_layout(height=max(320, 28 * min(len(d), 20)), yaxis_title="")
        else:
            serie = pd.to_numeric(data[q], errors="coerce").dropna()
            if serie.empty:
                st.caption(f"◽ {titre} : aucune réponse exploitable pour le moment")
                continue
            st.markdown(f"**{titre}** (n={len(serie)}) - moyenne : {serie.mean():,.0f} · "
                        f"médiane : {serie.median():,.0f} · min : {serie.min():,.0f} · "
                        f"max : {serie.max():,.0f}")
            fig = px.histogram(serie, nbins=30, title=f"{titre} (n={len(serie)})",
                               labels={"value": titre}, marginal="box")
            fig.update_layout(showlegend=False, height=400)
        st.plotly_chart(fig, width="stretch", key=f"{cle}_{q}")
    if not rendu:
        st.info("Aucune réponse exploitable dans cette section pour le moment.")


def questions_de(formulaire, data, gname):
    return [q for q, g in formulaire["group_of"].items()
            if g == gname and q in data.columns and q not in CHAMPS_EXCLUS
            and formulaire["qtype"][q].split()[0] in
            ("select_one", "select_multiple", "integer", "decimal", "text")]


with tab3:
    libelles = {form["groups"][g]: g for g in form["ordre_groupes"]}
    gsel = st.selectbox("Module du questionnaire", list(libelles))
    afficher_questions(fdf, form, questions_de(form, fdf, libelles[gsel]), "quanti")

# ========================================================== 4. QUALITATIF ====
def suivi_plan_qualitatif(donnees):
    """Ce qui est réalisé et ce qui reste, pour les focus groups, les profils
    d'entretiens individuels et les structures institutionnelles."""
    st.subheader("Suivi du plan qualitatif")
    total_prevu = total_fait = 0
    tables = {}
    for champ in ("groupe_fg", "profil_ind", "structure_inst"):
        t = suivi_plan_quali(donnees, form_q, champ)
        if not t.empty:
            tables[champ] = t
            total_prevu += int(t["Prévu"].sum())
            total_fait += int(t["Réalisé"].sum())
    if not tables:
        return
    avance = total_fait / total_prevu if total_prevu else 0
    st.progress(min(avance, 1.0),
                text=f"Avancement global : {total_fait} / {total_prevu} "
                     f"séances et entretiens ({avance:.0%})")
    cols = st.columns(len(tables))
    for col, (champ, t) in zip(cols, tables.items()):
        with col:
            fait = int((t["Réalisé"] >= t["Prévu"]).sum())
            st.markdown(f"**{PLAN_QUALI_LIBELLES[champ]}** - {fait}/{len(t)} complétés")
            st.dataframe(t[["Item", "Réalisé", "Prévu", "Statut"]], width="stretch",
                         hide_index=True, height=min(38 * len(t) + 40, 300))
    manque = []
    for champ, t in tables.items():
        manque += list(t.loc[t["Réalisé"] == 0, "Item"])
    if manque:
        st.warning("**Pas encore réalisé :** " + " · ".join(manque))
    else:
        st.success("Toutes les séances et tous les entretiens prévus sont couverts.")


with tab4:
    if dfq.empty:
        st.info("Aucun entretien ni focus group saisi pour le moment - le plan ci-dessous "
                "se cochera au fur et à mesure de la collecte qualitative.")
        suivi_plan_qualitatif(dfq)
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Fiches saisies", len(dfq))
        if "type_outil" in dfq:
            vc = dfq["type_outil"].value_counts()
            c2.metric("Entretiens individuels", int(vc.get("Entretien individuel", 0)))
            c3.metric("Entretiens institutionnels", int(vc.get("Entretien institutionnel", 0)))
            c4.metric("Focus groups", int(vc.get("Focus group", 0)))

        p1, p2 = st.columns(2)
        with p1:
            if "type_outil" in dfq and dfq["type_outil"].notna().any():
                d = dfq["type_outil"].value_counts().reset_index()
                d.columns = ["Type", "Nombre"]
                fig = px.pie(d, names="Type", values="Nombre", hole=0.45,
                             title="Répartition par type d'outil")
                fig.update_traces(textinfo="percent+value")
                st.plotly_chart(fig, width="stretch")
        with p2:
            if "groupe_fg" in dfq and dfq["groupe_fg"].notna().any():
                d = dfq["groupe_fg"].value_counts().reset_index()
                d.columns = ["Focus group", "Séances"]
                fig = px.bar(d.sort_values("Séances"), x="Séances", y="Focus group",
                             orientation="h", title="Séances par focus group", text_auto=True)
                fig.update_layout(yaxis_title="")
                st.plotly_chart(fig, width="stretch")

        cols_part = [c for c in ["nb_participants", "nb_femmes", "nb_jeunes", "nb_migrants"]
                     if c in dfq]
        if cols_part:
            tot = {form_q["label"].get(c, c): int(pd.to_numeric(dfq[c], errors="coerce").sum())
                   for c in cols_part}
            st.markdown("**Participants aux focus groups**")
            cs = st.columns(len(tot))
            for col, (k, v) in zip(cs, tot.items()):
                col.metric(k.replace("dont ", "Dont "), v)

        st.divider()
        suivi_plan_qualitatif(dfq)

        st.divider()
        libelles_q = {form_q["groups"][g]: g for g in form_q["ordre_groupes"]}
        gq = st.selectbox("Section de la grille qualitative", list(libelles_q))
        sous = dfq
        if "type_outil" in dfq:
            corresp = {"Entretien individuel": "Entretien individuel",
                       "Entretien institutionnel": "Entretien institutionnel",
                       "Focus group": "Focus group"}
            if gq in corresp:
                sous = dfq[dfq["type_outil"] == corresp[gq]]
                st.caption(f"{len(sous)} fiche(s) de ce type")
        afficher_questions(sous, form_q, questions_de(form_q, sous, libelles_q[gq]), "quali")

# ==================================================== 5. DONNÉES & EXPORT ====
with tab5:
    jeu = st.radio("Jeu de données", ["Enquête individuelle", "Entretiens & focus groups"],
                   horizontal=True)
    source, formulaire, nom = (fdf, form, "enquete_individuelle") if jeu == "Enquête individuelle" \
        else (dfq, form_q, "qualitatif")

    if source.empty:
        st.info("Aucune donnée à afficher.")
    else:
        ren, vus = {}, {}
        for c in source.columns:
            lab = str(formulaire["label"].get(c, c)).strip()
            n = vus.get(lab, 0)
            vus[lab] = n + 1
            ren[c] = lab if n == 0 else f"{lab} ({n + 1})"
        show = source.rename(columns=ren)
        st.dataframe(show, width="stretch", height=500)

        c1, c2 = st.columns(2)
        csv = show.to_csv(index=False).encode("utf-8-sig")
        c1.download_button("⬇️ Télécharger CSV", csv, f"C40_{nom}.csv", "text/csv")

        buf = io.BytesIO()
        exp = show.copy()
        for col in exp.columns:
            if exp[col].map(lambda x: isinstance(x, (list, dict))).any():
                exp[col] = exp[col].astype(str)
            if isinstance(exp[col].dtype, pd.DatetimeTZDtype):
                exp[col] = exp[col].dt.tz_localize(None)
        exp.to_excel(buf, index=False)
        c2.download_button("⬇️ Télécharger Excel", buf.getvalue(), f"C40_{nom}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
