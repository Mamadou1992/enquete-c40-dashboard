"""Suivi de la campagne MODECOM Pole Nord - version en un seul fichier.

Tout est regroupe ici : configuration, lecture des XLSForm, sources de donnees,
controles qualite et les cinq vues, presentees en onglets.

Aucun dossier n'est necessaire pour demarrer. Les fichiers de donnees sont
cherches dans `donnees/` puis a la racine ; s'ils manquent, l'application
s'ouvre quand meme et signale ce qui est indisponible.
"""
import hashlib
import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Suivi de la caractérisation - Pôle Nord", page_icon="♻️", layout="wide")

# =====================================================================
# CONFIGURATION
# =====================================================================
RACINE = Path(__file__).resolve().parent
DOSSIERS_DONNEES = [RACINE / "donnees", RACINE]

FORM_SOCIO = "Enquête ménage MODECOM Pôle Nord.xlsx"
FORM_CARAC = "Fiche de caractérisation MODECOM (13 catégories).xlsx"
FORM_SITES_AGRO = "Identification et cartographie des sites — Agro-pastoral (Pôle Nord).xlsx"
FORM_TRI_AGRO = "Caractérisation agro-pastorale (A1–A6) — MODECOM Pôle Nord.xlsx"
FORM_PHOTOS = "Photo caractérisation.xlsx"
FICHIER_GPKG = "Caracterisation_communes.gpkg"

KOBO_URL = "https://kf.kobotoolbox.org"

# Roles des formulaires. Les identifiants connus sont declares ici ;
# ceux qui ne le sont pas sont reconnus a leur contenu.
ROLE_MENAGES, ROLE_TRI = "menages", "tri"
ROLE_SITES_AGRO, ROLE_TRI_AGRO = "sites_agro", "tri_agro"
ROLE_PHOTOS = "photos"

UID_ROLES = {
    "aRuKgN8hXyDTyLRbSTLdNH": ROLE_MENAGES,      # enquête ménage, version en service
    "aG2BGSMEib9xWfRzoeXcXm": ROLE_MENAGES,      # enquête socio-démographique, 1re version
    "aFVV4abQHn6H8p7hYPen44": ROLE_TRI,          # fiche de caractérisation, 13 catégories
    "aSnArkcnDbH4pDqq9uDFTy": ROLE_SITES_AGRO,   # identification des sites agro-pastoraux
    "aCvQFPEqY9sPUGjXFLPh6q": ROLE_TRI_AGRO,     # caractérisation agro-pastorale A1-A6
    "aabjLvJkGkwgDA3gaUzJtj": ROLE_PHOTOS,       # reportage photo par commune
}
UIDS_DEFAUT = list(UID_ROLES)

FORMULAIRES = {
    ROLE_MENAGES:    {"titre": "Enquête ménage", "fichier": FORM_SOCIO},
    ROLE_TRI:        {"titre": "Caractérisation des déchets", "fichier": FORM_CARAC},
    ROLE_SITES_AGRO: {"titre": "Sites agro-pastoraux", "fichier": FORM_SITES_AGRO},
    ROLE_TRI_AGRO:   {"titre": "Caractérisation agro-pastorale", "fichier": FORM_TRI_AGRO},
    ROLE_PHOTOS:     {"titre": "Reportage photo", "fichier": FORM_PHOTOS},
}

CATEGORIES_AGRO = {
    "a1": "A1 déjections animales", "a2": "A2 résidus de cultures",
    "a3": "A3 déchets d'abattage", "a4": "A4 emballages agricoles",
    "a5": "A5 cuirs et peaux", "a6": "A6 inertes agro-pastoraux",
}

CV_HYPOTHESE = 0.40
MARGE_CIBLE = 0.10
CIBLE_MENAGES = 62
# Cibles particulieres : Ranerou est en regime allege
CIBLES_PARTICULIERES = {"Ranérou": 30}
# Ces valeurs sont des plafonds : le terrain peut remonter moins de sacs.

COMMUNES = {
    "Dagana":       {"infra": "CTT Dagana",  "region": "Saint-Louis", "methode": "Sur sites de collecte"},
    "Richard Toll": {"infra": "CTT Dagana",  "region": "Saint-Louis", "methode": "MODECOM à la source"},
    "Bokhol":       {"infra": "CIVD Bokhol", "region": "Saint-Louis", "methode": "Sur sites de collecte"},
    "Fanaye":       {"infra": "CIVD Bokhol", "region": "Saint-Louis", "methode": "MODECOM à la source"},
    "Ndioum":       {"infra": "CIVD Ndioum", "region": "Saint-Louis", "methode": "MODECOM à la source et sur sites"},
    "Podor":        {"infra": "CIVD Ndioum", "region": "Saint-Louis", "methode": "MODECOM à la source"},
    "Golléré":      {"infra": "CIVD Ndioum", "region": "Saint-Louis", "methode": "MODECOM à la source"},
    "Ogo":          {"infra": "CIVD Ogo",    "region": "Matam",
                     "methode": "MODECOM adapté agro-pastoral et sur sites"},
    "Matam":        {"infra": "CIVD Ogo",    "region": "Matam",       "methode": "MODECOM à la source et sur sites"},
    "Ourossogui":   {"infra": "CIVD Ogo",    "region": "Matam",       "methode": "MODECOM à la source et sur sites"},
    "Bokidiawé":    {"infra": "CIVD Ogo",    "region": "Matam",       "methode": "MODECOM adapté agro-pastoral"},
    "Ranérou":      {"infra": "CET Ranérou", "region": "Matam",       "methode": "MODECOM à la source, régime allégé"},
}
COMMUNES_SOURCE = [c for c, v in COMMUNES.items() if v["methode"].startswith("MODECOM")]
COMMUNES_SITE = [c for c, v in COMMUNES.items() if not v["methode"].startswith("MODECOM")]


def cible_commune(commune):
    """Nombre de menages a equiper en sacs, 0 si prelevement sur site de collecte."""
    if commune not in COMMUNES_SOURCE:
        return 0
    return CIBLES_PARTICULIERES.get(commune, CIBLE_MENAGES)


CIBLE_TOTALE = sum(cible_commune(c) for c in COMMUNES)

# Regroupement des communes sur la methode declaree, sans famille
# intermediaire : chaque methode du dictionnaire COMMUNES forme un groupe,
# dans l'ordre ou elle apparait.
LIBELLES_COURTS = {
    "MODECOM à la source": "À la source",
    "MODECOM à la source et sur sites": "À la source et sur site de collecte",
    "MODECOM à la source, régime allégé": "À la source, régime allégé",
    "MODECOM adapté agro-pastoral": "Adapté agro-pastoral",
    "MODECOM adapté agro-pastoral et sur sites":
        "Agro-pastoral et site de collecte",
    "Sur sites de collecte": "Sur site de collecte (décharge)",
}

COMMUNES_PAR_METHODE = {}
for _c, _v in COMMUNES.items():
    COMMUNES_PAR_METHODE.setdefault(_v["methode"], []).append(_c)

COULEURS = {
    "MODECOM à la source": "#1B7F4B",
    "MODECOM à la source, régime allégé": "#4FB477",
    "MODECOM à la source et sur sites": "#2E9B6B",
    "MODECOM adapté agro-pastoral": "#8DC63F",
    "MODECOM adapté agro-pastoral et sur sites": "#6FA32F",
    "Sur sites de collecte": "#F4A93B",
    "Sur sites régime allégé": "#B01B2E",
}
VERT = "#1B5E36"
ENCRE = "#1a1a1a"

NOMS_LOGO = ["logo_sgp.png", "logo_sgp.jpg", "logo_sgp.jpeg", "logo_SGP.png",
             "logo_SGP.jpg", "logo_SGP.jpeg"]


def chemin(nom):
    for d in DOSSIERS_DONNEES:
        p = d / nom
        if p.exists():
            return p
    return None


def logo():
    for d in [RACINE / "assets", RACINE]:
        for n in NOMS_LOGO:
            p = d / n
            if p.exists():
                return p
    return None


# =====================================================================
# ACCES PAR MOT DE PASSE
# =====================================================================
CLE_AUTH = "authentifie"


def _empreinte(t):
    return hashlib.sha256(str(t).encode("utf-8")).hexdigest()


CLES_MDP = ["mot_de_passe", "motdepasse", "password", "mdp"]
CLES_EMPREINTE = ["empreinte_mot_de_passe", "empreinte", "hash"]


def _lire_secrets():
    """Renvoie (dictionnaire des secrets, message d'erreur eventuel)."""
    try:
        return dict(st.secrets), None
    except Exception as e:
        return {}, str(e)


def _attendue():
    """Empreinte attendue. Cherche dans [app], puis a la racine des secrets."""
    secrets, _ = _lire_secrets()
    sources = []
    for nom in ("app", "App", "APP", "general"):
        bloc = secrets.get(nom)
        if hasattr(bloc, "keys"):
            sources.append(bloc)
    sources.append(secrets)          # cles posees sans section
    for bloc in sources:
        for c in CLES_EMPREINTE:
            v = bloc.get(c) if hasattr(bloc, "get") else None
            if v:
                return str(v).strip().lower()
        for c in CLES_MDP:
            v = bloc.get(c) if hasattr(bloc, "get") else None
            if v:
                return _empreinte(v)
    return None


def _diagnostic_secrets():
    """Decrit ce que Streamlit voit, sans jamais afficher de valeur."""
    secrets, err = _lire_secrets()
    if err:
        return f"Streamlit ne trouve aucun secret. Message : {err}"
    if not secrets:
        return "Streamlit lit bien les secrets, mais ils sont vides."
    lignes = []
    for cle, val in secrets.items():
        if hasattr(val, "keys"):
            lignes.append(f"- section `[{cle}]` contenant : "
                          + ", ".join(f"`{k}`" for k in val.keys()))
        else:
            lignes.append(f"- clé `{cle}` posée hors section")
    return "Voici ce que Streamlit voit dans les secrets :\n\n" + "\n".join(lignes)


VERT_CLAIR = "#8DC63F"
FOND = "#10151A"        # fond general
SURFACE = "#19212A"     # cartes et panneaux
BORDURE = "#2C3742"
TEXTE = "#E9ECE8"
TEXTE_DOUX = "#9BA6A0"
SABLE = SURFACE
GRIS = TEXTE_DOUX

CSS = """
<style>
  #MainMenu, footer {visibility:hidden;}
  .block-container {padding-top:1.4rem; padding-bottom:2rem; max-width:1500px;}

  .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"]
      {background:#10151A; color:#E9ECE8;}
  .stApp p, .stApp li, .stApp label, .stApp span, .stApp h1, .stApp h2, .stApp h3,
  .stApp h4, .stApp h5, .stApp h6, .stApp strong {color:#E9ECE8;}
  [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p
      {color:#9BA6A0 !important;}

  .entete {background:linear-gradient(100deg,#0E3D24 0%,#17663A 55%,#1F8A4C 100%);
           color:#fff; padding:18px 24px; border-radius:10px; margin-bottom:20px;
           border:1px solid #2B6B45;}
  .entete h1 {margin:0; font-size:27px; font-weight:700; letter-spacing:-.2px; color:#fff;}
  .entete p  {margin:6px 0 0 0; font-size:14px; color:#D6E7DC; max-width:900px;}

  .fiche {background:#19212A; border:1px solid #2C3742; border-left:4px solid #35C46F;
          border-radius:8px; padding:12px 16px; height:100%;}
  .fiche .lib  {font-size:12.5px; color:#9BA6A0; text-transform:uppercase;
                letter-spacing:.4px; margin:0 0 6px 0; line-height:1.25;}
  .fiche .val  {font-size:27px; font-weight:700; color:#F2F5F1; line-height:1;}
  .fiche .note {font-size:12px; color:#8D978F; margin-top:5px;}
  .fiche.alerte {border-left-color:#E4576B;}
  .fiche.veille {border-left-color:#F4A93B;}
  .fiche.neutre {border-left-color:#4A5763;}

  .rubrique {font-size:17px; font-weight:700; color:#7ED694; margin:22px 0 2px 0;}
  .rubrique + hr {margin:6px 0 14px 0; border:none; border-top:1px solid #2C3742;}

  .stTabs [data-baseweb="tab-list"] {gap:2px; border-bottom:1px solid #2C3742;}
  .stTabs [data-baseweb="tab"] {padding:9px 18px; font-size:14.5px; font-weight:600;
                               color:#9BA6A0; background:transparent;}
  .stTabs [aria-selected="true"] {color:#7ED694 !important;
                                  border-bottom:3px solid #35C46F !important;}
  .stTabs [data-baseweb="tab"] p {color:inherit !important;}

  section[data-testid="stSidebar"] {background:#141B22; border-right:1px solid #2C3742;}
  section[data-testid="stSidebar"] img {margin-bottom:6px; border-radius:4px;}
  section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2,
  section[data-testid="stSidebar"] h3, section[data-testid="stSidebar"] p,
  section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] li,
  section[data-testid="stSidebar"] .stMarkdown,
  section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p
      {color:#E9ECE8 !important;}
  section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
  section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p
      {color:#9BA6A0 !important;}
  section[data-testid="stSidebar"] h2 {font-size:16px; font-weight:700;
      letter-spacing:.3px; text-transform:uppercase; color:#7ED694 !important;}
  section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] textarea
      {background:#0D1217 !important; color:#E9ECE8 !important;
       border:1px solid #2C3742 !important;}
  section[data-testid="stSidebar"] .stButton button
      {background:#1F8A4C; color:#fff !important; border:none; font-weight:600;}
  section[data-testid="stSidebar"] .stButton button:hover {background:#166B3A;}
  section[data-testid="stSidebar"] .stButton button p {color:#fff !important;}

  [data-testid="stAlert"] {background:#19212A; border:1px solid #2C3742;
                           border-radius:8px;}
  [data-testid="stAlert"] p {color:#E9ECE8 !important;}

  div[data-testid="stMetricValue"] {font-size:25px; font-weight:700; color:#F2F5F1;}
  div[data-testid="stMetricLabel"] {color:#9BA6A0;}
  .stDataFrame {border:1px solid #2C3742; border-radius:8px;}
  .stDownloadButton button {background:#19212A; color:#E9ECE8 !important;
                            border:1px solid #2C3742;}
  .stDownloadButton button:hover {border-color:#35C46F; color:#7ED694 !important;}
  .pied {color:#7E8A83; font-size:12px; border-top:1px solid #2C3742;
         margin-top:26px; padding-top:12px;}
</style>
"""


def bandeau(titre, sous_titre):
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(f"<div class='entete'><h1>{titre}</h1><p>{sous_titre}</p></div>",
                unsafe_allow_html=True)


def fiche(colonne, libelle, valeur, note=None, ton="normal"):
    """Carte d'indicateur. ton : normal, alerte, veille, neutre."""
    classe = "fiche" if ton == "normal" else f"fiche {ton}"
    html = (f"<div class='{classe}'><p class='lib'>{libelle}</p>"
            f"<div class='val'>{valeur}</div>")
    if note:
        html += f"<div class='note'>{note}</div>"
    colonne.markdown(html + "</div>", unsafe_allow_html=True)


def rubrique(texte):
    st.markdown(f"<div class='rubrique'>{texte}</div><hr>", unsafe_allow_html=True)


def habiller(fig, hauteur=None):
    """Mise en forme commune des graphiques Plotly."""
    fig.update_layout(
        font=dict(family="Source Sans Pro, Segoe UI, sans-serif", size=13, color="#E9ECE8"),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0, r=40, t=28, b=0),
        xaxis=dict(gridcolor="#2C3742", zeroline=False, linecolor="#2C3742"),
        yaxis=dict(gridcolor="#2C3742", zeroline=False, linecolor="#2C3742"),
        legend=dict(font=dict(color="#E9ECE8")),
        title_font=dict(size=14, color="#7ED694"),
    )
    if hauteur:
        fig.update_layout(height=hauteur)
    return fig


def attente(message):
    st.markdown(
        f"<div style='background:{SURFACE};border:1px dashed {BORDURE};"
        f"border-radius:10px;padding:34px 24px;text-align:center;color:{TEXTE_DOUX};"
        "margin-top:8px'>"
        f"<div style='font-size:30px;margin-bottom:8px'>⏳</div>{message}</div>",
        unsafe_allow_html=True)


def exiger_mot_de_passe():
    if st.session_state.get(CLE_AUTH):
        return True
    attendue = _attendue()
    if attendue is None:
        bandeau("Suivi de la campagne de Caractérisation Pôle Nord",
                "Accès réservé à l'équipe du projet")
        st.error("Aucun mot de passe n'est configuré. Ajouter dans les secrets :\n\n"
                 "```toml\n[app]\nmot_de_passe = \"votre_mot_de_passe\"\n```")
        st.warning(_diagnostic_secrets())
        st.caption(
            "Sur Streamlit Cloud : Manage app, menu à trois points, Settings, onglet "
            "Secrets. Vérifier que les guillemets sont droits et qu'aucun espace ne "
            "précède la ligne [app]. Après enregistrement, redémarrer avec Reboot app."
        )
        st.stop()
    bandeau("Suivi de la campagne de Caractérisation Pôle Nord",
                "Accès réservé à l'équipe du projet")
    with st.form("connexion"):
        saisi = st.text_input("Mot de passe", type="password")
        valider = st.form_submit_button("Entrer")
    if valider:
        if _empreinte(saisi) == attendue:
            st.session_state[CLE_AUTH] = True
            st.rerun()
        else:
            st.error("Mot de passe incorrect.")
    st.caption("Les données d'enquête portent sur des ménages identifiés. "
               "Ne pas diffuser le mot de passe hors de l'équipe.")
    st.stop()


# =====================================================================
# LECTURE DES XLSFORM
# =====================================================================
LABEL = "label::Français (fr)"
META = {"start", "end", "today", "deviceid", "note", "begin_group", "end_group",
        "begin_repeat", "end_repeat", "nan"}
GRANULOS = ["hétéroclites", "grossiers", "moyens"]
MASSES_GLOBALES = ["masse totale brute", "masse après quartage", "masse nette totale"]


@st.cache_data(show_spinner=False)
def charger_form(chemin_fichier):
    survey = pd.read_excel(chemin_fichier, sheet_name="survey")
    choices = pd.read_excel(chemin_fichier, sheet_name="choices")
    survey.columns = [str(c).strip() for c in survey.columns]
    choices.columns = [str(c).strip() for c in choices.columns]
    survey["type"] = survey["type"].astype(str).str.strip()
    survey["name"] = survey["name"].astype(str).str.strip()
    if LABEL not in survey.columns:
        cand = [c for c in survey.columns if c.startswith("label")]
        survey[LABEL] = survey[cand[0]] if cand else ""
    groupes, courant, lignes = [], None, []
    for _, r in survey.iterrows():
        t = r["type"]
        if t == "begin_group":
            courant = str(r[LABEL])
            groupes.append({"name": r["name"], "label": courant})
            continue
        if t == "end_group":
            courant = None
            continue
        base = t.split()[0] if t and t != "nan" else t
        if base in META:
            continue
        lignes.append({"groupe": courant or "(hors groupe)", "type": t, "type_base": base,
                       "name": r["name"], "label": str(r[LABEL])})
    return {"survey": survey, "choices": choices, "groupes": groupes,
            "questions": pd.DataFrame(lignes), "col_liste": choices.columns[0]}


def options(form, liste):
    ch = form["choices"]
    col = form["col_liste"]
    lab = LABEL if LABEL in ch.columns else ch.columns[2]
    sub = ch[ch[col].astype(str).str.strip() == str(liste)]
    return [str(x).strip() for x in sub[lab].dropna().tolist()]


def decoupe_categorie(label):
    txt = re.sub(r"^\s*Masse\s*\(\s*Kg\s*\)?\s*", "", str(label)).strip()
    txt = re.sub(r"^\s*Masse\s*\(\s*Kg\s*", "", txt).strip()
    bas = txt.lower()
    for g in GRANULOS:
        if bas.endswith(g):
            return txt[: -len(g)].strip(), g
    return txt, "global"


def table_pesees(form):
    q = form["questions"]
    d = q[q["type_base"] == "decimal"].copy()
    dec = d["label"].apply(decoupe_categorie)
    d["categorie"] = [x[0] for x in dec]
    d["granulometrie"] = [x[1] for x in dec]
    d["globale"] = d["categorie"].str.lower().isin(MASSES_GLOBALES)
    return d.reset_index(drop=True)


def resume(form):
    q = form["questions"]
    return {"questions": len(q), "groupes": len(form["groupes"]),
            "listes": form["choices"][form["col_liste"]].nunique(),
            "types": q["type_base"].value_counts().to_dict()}


# =====================================================================
# SOURCES DE DONNEES
# =====================================================================
def sans_accent(s):
    return "".join(c for c in unicodedata.normalize("NFKD", str(s))
                   if not unicodedata.combining(c))


def rapproche_commune(v):
    if pd.isna(v):
        return None
    x = sans_accent(v).strip().lower().replace("_", " ").replace("-", " ")
    for off in COMMUNES:
        if sans_accent(off).lower().replace("-", " ") == x:
            return off
    return str(v).strip()


def charger_kobo(base_url, token, uid, timeout=60):
    import requests
    url = f"{base_url.rstrip('/')}/api/v2/assets/{uid}/data.json"
    entetes = {"Authorization": f"Token {token}"}
    lignes, page = [], url
    while page:
        r = requests.get(page, headers=entetes, timeout=timeout)
        r.raise_for_status()
        js = r.json()
        lignes.extend(js.get("results", []))
        page = js.get("next")
    return pd.json_normalize(lignes)


DUREE_CACHE = 300  # secondes : au-dela, les soumissions sont relues sur Kobo

SIGNES = {
    ROLE_MENAGES: {"menage_id", "nb_total", "nb_present", "consentement",
                   "tri_source", "dispose_trier"},
    ROLE_TRI: {"masse_echantillon", "masse_quart", "masse_totale_triee",
               "total_het", "total_g100", "total_m20", "s1_1_het"},
    ROLE_SITES_AGRO: {"categories_presentes", "accessibilite", "type_activite", "site"},
    ROLE_TRI_AGRO: {"masse_a1", "masse_a3", "te_a1", "momov_a1", "categorie"},
    ROLE_PHOTOS: {"Point_and_shoot_Use_mera_to_take_a_photo",
                  "Point_and_shoot_Use_mera_to_take_a_photo_001",
                  "Point_and_shoot_Use_mera_to_take_a_photo_002",
                  "Point_and_shoot_Use_mera_to_take_a_photo_003"},
}


def deviner_role(df):
    """Role d'un jeu de soumissions, d'apres ses colonnes."""
    if df is None or getattr(df, "empty", True):
        return None
    cols = {str(c).split("/")[-1] for c in df.columns}
    scores = {r: len(cols & sig) for r, sig in SIGNES.items()}
    meilleur = max(scores, key=scores.get)
    return meilleur if scores[meilleur] > 0 else None


@st.cache_data(ttl=DUREE_CACHE, show_spinner="Récupération des soumissions sur Kobo...")
def recuperer_kobo(url, token, uids):
    """Recupere tous les formulaires et les range par role.

    Un identifiant declare dans UID_ROLES prend ce role ; les autres sont
    reconnus a leur contenu. Plusieurs formulaires d'un meme role sont
    empiles, ce qui permet de garder l'ancienne et la nouvelle version
    d'une meme fiche.
    """
    par_role, journal = {}, []
    for uid in uids:
        if not uid:
            continue
        df = charger_kobo(url, token, uid)
        role = UID_ROLES.get(uid) or deviner_role(df)
        journal.append({"Formulaire": uid,
                        "Rôle": FORMULAIRES.get(role, {}).get("titre", "non reconnu"),
                        "Soumissions": 0 if df is None else len(df)})
        if role is None:
            continue
        par_role.setdefault(role, []).append(df)
    jeux = {r: pd.concat(v, ignore_index=True) if len(v) > 1 else v[0]
            for r, v in par_role.items()}
    return jeux, pd.DataFrame(journal), pd.Timestamp.now()


def secrets_kobo():
    """Adresse, jeton et liste des identifiants declares dans les secrets."""
    try:
        sec = st.secrets.get("kobo", {})
    except Exception:
        return {"url": "", "token": "", "uids": []}
    uids = []
    for cle in sorted(sec.keys()):
        if str(cle).lower().startswith("uid"):
            v = str(sec[cle]).strip()
            if v and v not in uids:
                uids.append(v)
    return {"url": sec.get("url", ""), "token": sec.get("token", ""), "uids": uids}


def charger_fichier(f):
    nom = getattr(f, "name", str(f)).lower()
    if nom.endswith((".xlsx", ".xls")):
        return pd.read_excel(f)
    donnees = f.read() if hasattr(f, "read") else open(f, "rb").read()
    for sep in (";", ",", "\t"):
        try:
            df = pd.read_csv(io.BytesIO(donnees), sep=sep)
            if df.shape[1] > 1:
                return df
        except Exception:
            continue
    return pd.read_csv(io.BytesIO(donnees))


SIGNES_SOCIO = {"menage_id", "nb_total", "nb_present", "consentement"}
SIGNES_CARAC = {"masse_totale_brute", "masse_apres_quartage", "masse_nette_totale",
                "code_echantillon", "Infrastructure_concern_e"}


def identifier_formulaire(df):
    if df is None or getattr(df, "empty", True):
        return None
    cols = {str(c).split("/")[-1] for c in df.columns}
    ns, nc = len(cols & SIGNES_SOCIO), len(cols & SIGNES_CARAC)
    return "carac" if nc > ns else ("socio" if ns > nc else None)


def repartir(jeux):
    socio = carac = None
    restants = []
    for df in jeux:
        g = identifier_formulaire(df)
        if g == "socio" and socio is None:
            socio = df
        elif g == "carac" and carac is None:
            carac = df
        elif df is not None:
            restants.append(df)
    for df in restants:
        if socio is None:
            socio = df
        elif carac is None:
            carac = df
    return socio, carac


COLS_SOCIO = {
    "menage_id": ["menage_id"], "commune": ["commune"], "region": ["region"],
    "quartier": ["quartier"], "strate": ["strate"], "logement": ["logement"],
    "revenu": ["revenu"], "enqueteur": ["enqueteur"],
    "consentement": ["consentement"],
    "nb_total": ["nb_total", "a1_total"], "nb_present": ["nb_present", "a2_present"],
    "recipient": ["recipient", "c1_recipient"], "service": ["service", "d1_service"],
    "operateur": ["operateur", "d2_operateur"],
    "frequence": ["frequence", "d4_frequence"],
    "pratiques": ["pratiques", "e1_pratiques"],
    "tri_source": ["tri_source", "e2_tri"], "materiaux": ["materiaux", "e2_materiaux"],
    "dispose_trier": ["dispose_trier", "f2_tri"],
    "elevage": ["elevage", "g1_elevage"], "especes": ["especes"],
    "effectif": ["effectif"], "fumier": ["fumier", "g3_fumier"],
    "residus_agri": ["residus_agri", "g4_residus"],
    "saison_var": ["saison_var", "g5_saison"],
    "observations": ["observations", "i_observations"], "gps": ["gps"],
    "date": ["today", "date", "_submission_time", "end"],
}
# Nomenclature MODECOM du formulaire de caracterisation : 13 categories,
# 50 sous-categories. Chaque sous-categorie est pesee en trois fractions
# granulometriques, sauf les fines qui portent une pesee unique.
NOMENCLATURE = {
    "1": {"titre": "D\u00e9chets putrescibles", "sous": {
        "s1_1": "D\u00e9chets alimentaires (restes de cuisine non consommable)",
        "s1_2": "Produits alimentaires non consomm\u00e9s sans emballages",
        "s1_3": "Produits alimentaires non consomm\u00e9s sous emballages",
        "s1_4": "Autres putrescibles",
        "s1_5": "D\u00e9chets de jardin",
    }},
    "2": {"titre": "Papiers", "sous": {
        "s2_1": "Emballages papiers",
        "s2_2": "Journaux, magazines et revues",
        "s2_3": "Imprim\u00e9s publicitaires",
        "s2_4": "Papiers bureautiques",
        "s2_5": "Autres papiers",
    }},
    "3": {"titre": "Cartons", "sous": {
        "s3_1": "Emballages cartons plats",
        "s3_2": "Emballages cartons ondul\u00e9s",
        "s3_3": "Autres cartons",
    }},
    "4": {"titre": "Composites", "sous": {
        "s4_1": "Emballages de Liquide Alimentaire (ELA)",
        "s4_2": "Autres emballages composites",
        "s4_3": "Petits Appareils \u00c9lectrom\u00e9nagers (PAM)",
    }},
    "5": {"titre": "Textiles", "sous": {
        "s5": "Textiles",
    }},
    "6": {"titre": "Textiles sanitaires", "sous": {
        "s6_1": "Couches b\u00e9b\u00e9s",
        "s6_2": "Autre fraction hygi\u00e9nique",
    }},
    "7": {"titre": "Plastiques", "sous": {
        "s7_1": "Sacs poubelle",
        "s7_2": "Autres sacs plastiques",
        "s7_3": "Autres films plastiques d'emballage",
        "s7_4": "Bouteilles et flacons en PET",
        "s7_5": "Bouteilles et flacons polyol\u00e9fines",
        "s7_6": "Autres emballages plastiques",
        "s7_7": "Autres plastiques",
    }},
    "8": {"titre": "Combustibles non class\u00e9s", "sous": {
        "s8_1": "Emballages en bois",
        "s8_2": "Chaussures",
        "s8_3": "Maroquinerie",
        "s8_4": "Autres combustibles",
    }},
    "9": {"titre": "Verre", "sous": {
        "s9_1": "Emballages en verre incolore",
        "s9_2": "Emballages en verre de couleur",
        "s9_3": "Autres verres",
    }},
    "10": {"titre": "M\u00e9taux", "sous": {
        "s10_1": "Emballages m\u00e9taux ferreux",
        "s10_2": "Emballages aluminium",
        "s10_3": "Autres m\u00e9taux ferreux",
        "s10_4": "Autres m\u00e9taux",
    }},
    "11": {"titre": "Incombustibles non class\u00e9s", "sous": {
        "s11_1": "Emballages incombustibles",
        "s11_2": "Autres incombustibles",
    }},
    "12": {"titre": "D\u00e9chets dangereux", "sous": {
        "s12_1": "D\u00e9chets diffus sp\u00e9cifiques",
        "s12_2": "Tubes fluorescents et lampes basses consommation",
        "s12_3": "Piles et accumulateurs",
        "s12_4": "D\u00e9chets d'activit\u00e9s de soins perforants",
        "s12_5": "Huiles min\u00e9rales",
        "s12_6": "Cartouche d'impression",
        "s12_7": "Bouteilles de gaz",
        "s12_8": "M\u00e9dicaments non utilis\u00e9s",
        "s12_9": "Autres d\u00e9chets m\u00e9nagers sp\u00e9ciaux",
    }},
    "13": {"titre": "Fines", "sous": {
        "s13_1": "\u00c9l\u00e9ments fins entre 8 et 20 mm",
        "s13_2": "\u00c9l\u00e9ments fins < 8 mm",
    }},
}

FRACTIONS = {"het": "Hétéroclites", "g100": "> 100 mm", "m20": "100 à 20 mm",
             "poids": "Fines"}


# Colonnes de masse du formulaire : sous-categorie et fraction.
def _colonnes_masse():
    cols = {}
    for num, v in NOMENCLATURE.items():
        for cle in v["sous"]:
            suffixes = ("poids",) if num == "13" else ("het", "g100", "m20")
            for f in suffixes:
                cols[f"{cle}_{f}"] = (num, cle, f)
    return cols


COLONNES_MASSE = _colonnes_masse()

COLS_CARAC = {
    "date": ["date_carac", "_submission_time"],
    "methode": ["mode_caracterisation"], "region": ["region"],
    "infrastructure": ["infrastructure"], "commune": ["commune"],
    "site": ["site_precis"], "code_echantillon": ["code_echantillon"],
    "enqueteur": ["enqueteur"], "gps": ["gps"],
    "strate": ["strate"], "nb_menages": ["nombre_menages"],
    "sacs_distribues": ["nombre_sachets_distribues"],
    "sacs_collectes": ["nombre_sachets_collectes"],
    "date_distribution": ["date_distribution"],
    "date_recuperation": ["date_recuperation"],
    "masse_echantillon": ["masse_echantillon"], "masse_quart": ["masse_quart"],
    "total_het": ["total_het"], "total_g100": ["total_g100"],
    "total_m20": ["total_m20"], "total_fines": ["total_fines"],
    "masse_triee": ["masse_totale_triee"],
    "notes": ["notes"],
}

STRATES = {"modeste": "Modeste", "moyen": "Moyen", "aise": "Aisé",
           "mixte": "Mixte résidentiel / commercial"}
MODES_CARAC = {"sur_sites": "MODECOM sur site", "a_la_source": "MODECOM à la source"}


def _renomme(df, corr):
    ren = {}
    for cible, cands in corr.items():
        for c in cands:
            if c in df.columns:
                ren[c] = cible
                break
    return df.rename(columns=ren)


def normalise_socio(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).split("/")[-1] for c in df.columns]
    df = _renomme(df, COLS_SOCIO)
    if "commune" in df:
        df["commune"] = df["commune"].map(rapproche_commune)
    for c in ("nb_total", "nb_present", "effectif"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    if "date" in df:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df


def normalise_carac(df):
    """Fiche de caracterisation MODECOM a 13 categories.

    Ajoute une colonne de masse par categorie, `cat_1` a `cat_13`, egale a la
    somme des sous-categories et des fractions granulometriques.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).split("/")[-1] for c in df.columns]
    df = _renomme(df, COLS_CARAC)
    if "commune" in df:
        df["commune"] = df["commune"].map(rapproche_commune)
    if "strate" in df:
        df["strate"] = df["strate"].map(lambda v: STRATES.get(str(v), v))
    if "methode" in df:
        df["methode"] = df["methode"].map(lambda v: MODES_CARAC.get(str(v), v))

    for c in ("masse_echantillon", "masse_quart", "total_het", "total_g100",
              "total_m20", "total_fines", "masse_triee", "nb_menages",
              "sacs_distribues", "sacs_collectes"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    presentes = [c for c in COLONNES_MASSE if c in df.columns]
    for c in presentes:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    totaux = {}
    for num in NOMENCLATURE:
        cols = [c for c in presentes if COLONNES_MASSE[c][0] == num]
        if cols:
            totaux[f"cat_{num}"] = df[cols].sum(axis=1, min_count=1)
    if totaux:
        df = pd.concat([df, pd.DataFrame(totaux, index=df.index)], axis=1)
        df["masse_triee_calc"] = df[list(totaux)].sum(axis=1, min_count=1)
    else:
        df["masse_triee_calc"] = pd.NA
    if "masse_triee" not in df:
        df["masse_triee"] = df["masse_triee_calc"]

    for c in ("date", "date_distribution", "date_recuperation"):
        if c in df:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return _coordonnees(df)


COLS_SITES_AGRO = {
    "date": ["date_releve", "_submission_time"], "technicien": ["technicien"],
    "superviseur": ["superviseur"], "commune": ["commune"], "site": ["site"],
    "type_activite": ["type_activite"], "categories": ["categories_presentes"],
    "accessibilite": ["accessibilite"], "contact": ["contact_local"],
    "observations": ["observations"], "gps": ["gps"],
}

COLS_TRI_AGRO = {
    "date": ["date_collecte", "_submission_time"], "technicien": ["technicien"],
    "superviseur": ["superviseur"], "commune": ["commune"], "site": ["site"],
    "type_activite": ["type_activite"], "categorie": ["categorie"],
    "code_echantillon": ["code_echantillon"], "meteo": ["meteo"], "saison": ["saison"],
    "methode": ["methode"], "masse_brute": ["masse_brute"],
    "masse_quartage": ["masse_quartage"], "masse_nette": ["masse_nette"],
    "masse_totale_caract": ["masse_totale_caract"],
}


COLS_PHOTOS = {
    "date": ["_submission_time", "start"], "commune": ["Commune", "commune"],
}


def normalise_photos(df):
    """Reportage photo : une soumission, une commune, jusqu'a quatre images."""
    if df is None or getattr(df, "empty", True):
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).split("/")[-1] for c in df.columns]
    df = _renomme(df, COLS_PHOTOS)
    if "commune" in df:
        df["commune"] = df["commune"].map(rapproche_commune)
    if "date" in df:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.sort_values("date", ascending=False)
    return df


def nb_photos_commune(t, commune):
    """Nombre d'images rattachees a une commune, tous formulaires confondus."""
    if t is None or getattr(t, "empty", True) or "commune" not in t:
        return 0
    return int((t["commune"] == commune).sum())


def _coordonnees(df, colonne="gps"):
    """Extrait latitude et longitude d'un champ geopoint de Kobo."""
    if colonne not in df:
        return df
    brut = df[colonne].astype(str).str.strip()
    parts = brut.str.split(r"\s+", expand=True)
    if parts.shape[1] >= 2:
        df["lat"] = pd.to_numeric(parts[0], errors="coerce")
        df["lon"] = pd.to_numeric(parts[1], errors="coerce")
    for a, b in (("_gps_latitude", "lat"), ("_gps_longitude", "lon")):
        if a in df.columns and df.get(b) is None:
            df[b] = pd.to_numeric(df[a], errors="coerce")
    return df


def normalise_sites_agro(df):
    if df is None or getattr(df, "empty", True):
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).split("/")[-1] for c in df.columns]
    df = _renomme(df, COLS_SITES_AGRO)
    if "commune" in df:
        df["commune"] = df["commune"].map(rapproche_commune)
    if "date" in df:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return _coordonnees(df)


def normalise_tri_agro(df):
    if df is None or getattr(df, "empty", True):
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).split("/")[-1] for c in df.columns]
    df = _renomme(df, COLS_TRI_AGRO)
    if "commune" in df:
        df["commune"] = df["commune"].map(rapproche_commune)
    for c in ["masse_brute", "masse_quartage", "masse_nette", "masse_totale_caract"]:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for cle in CATEGORIES_AGRO:
        for prefixe in ("masse", "te", "momov", "pci", "da"):
            col = f"{prefixe}_{cle}"
            if col in df:
                df[col] = pd.to_numeric(df[col], errors="coerce")
    if "date" in df:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df


# ---------------------------------------------------------------------
# PHOTOS JOINTES AUX SOUMISSIONS
# ---------------------------------------------------------------------
@st.cache_data(ttl=1800, show_spinner=False)
def telecharger_image(url, token):
    """Recupere une piece jointe Kobo, qui exige le jeton d'API."""
    import requests
    r = requests.get(url, headers={"Authorization": f"Token {token}"}, timeout=60)
    r.raise_for_status()
    return r.content


CLES_LEGENDE = ["commune", "site", "menage_id", "code_echantillon", "quartier"]


def table_photos(sources):
    """Toutes les images jointes, quel que soit le formulaire d'origine.

    `sources` associe un libelle a un jeu normalise. Une soumission peut
    porter plusieurs images : chacune donne une ligne.
    """
    lignes = []
    for libelle, df in sources.items():
        if df is None or getattr(df, "empty", True):
            continue
        if "_attachments" not in df.columns:
            continue
        for _, r in df.iterrows():
            jointes = r.get("_attachments")
            if not isinstance(jointes, list):
                continue
            for a in jointes:
                if not isinstance(a, dict):
                    continue
                if "image" not in str(a.get("mimetype", "")).lower():
                    continue
                url = a.get("download_small_url") or a.get("download_url")
                if not url:
                    continue
                bouts = [str(r.get(c)) for c in CLES_LEGENDE
                         if r.get(c) is not None and str(r.get(c)) not in ("nan", "")]
                lignes.append({
                    "url": url, "source": libelle,
                    "commune": r.get("commune"), "date": r.get("date"),
                    "legende": " · ".join(bouts) or "sans identifiant"})
    t = pd.DataFrame(lignes)
    if not t.empty:
        t["date"] = pd.to_datetime(t.get("date"), errors="coerce")
        t = t.sort_values("date", ascending=False, na_position="last")
    return t


def galerie_table(t, token, colonnes=4):
    """Affiche une table d'images produite par table_photos."""
    if t.empty:
        return 0
    cols = st.columns(colonnes)
    affichees = 0
    for i, (_, r) in enumerate(t.iterrows()):
        cible = cols[i % colonnes]
        legende = f"{r['commune'] or 'commune non renseignée'} · {r['source']}"
        try:
            cible.image(telecharger_image(r["url"], token), caption=legende,
                        width="stretch")
            affichees += 1
        except Exception:
            cible.caption(f"Image indisponible · {legende}")
    return affichees


def liens_photos(df, limite=24):
    """Liste (url, legende) des images jointes aux soumissions."""
    liens = []
    if df is None or getattr(df, "empty", True) or "_attachments" not in df.columns:
        return liens
    for _, r in df.iterrows():
        jointes = r.get("_attachments")
        if not isinstance(jointes, list):
            continue
        for a in jointes:
            if not isinstance(a, dict):
                continue
            if "image" not in str(a.get("mimetype", "")).lower():
                continue
            url = a.get("download_small_url") or a.get("download_url")
            if not url:
                continue
            bouts = [str(r.get(c)) for c in CLES_LEGENDE
                     if r.get(c) is not None and str(r.get(c)) not in ("nan", "")]
            liens.append((url, " · ".join(bouts) or "sans identifiant"))
            if len(liens) >= limite:
                return liens
    return liens


def galerie(df, token, titre="Photos", colonnes=4, limite=24):
    """Affiche les photos jointes, si le formulaire en contient."""
    liens = liens_photos(df, limite)
    if not liens:
        return
    if titre:
        rubrique(titre)
    cols = st.columns(colonnes)
    affichees = 0
    for i, (url, legende) in enumerate(liens):
        cible = cols[i % colonnes]
        try:
            cible.image(telecharger_image(url, token), caption=legende,
                        width="stretch")
            affichees += 1
        except Exception:
            cible.caption(f"Image indisponible · {legende}")
    st.caption(f"{affichees} photos affichées, les plus récentes d'abord. "
               "Les images restent hébergées sur Kobo, rien n'est copié ici.")


# =====================================================================
# CONTROLES QUALITE
# =====================================================================
def _ajoute(a, gravite, quest, controle, sous, cle, detail):
    for _, r in sous.iterrows():
        a.append({"gravité": gravite, "questionnaire": quest, "contrôle": controle,
                  "identifiant": r.get(cle, ""), "commune": r.get("commune", ""),
                  "détail": detail(r) if callable(detail) else detail})


def controler_socio(df):
    a = []
    if df is None or df.empty:
        return pd.DataFrame(a)
    Q = "Socio-démographique"
    if "menage_id" in df:
        dup = df[df.duplicated("menage_id", keep=False) & df["menage_id"].notna()]
        _ajoute(a, "Bloquant", Q, "Identifiant de ménage en doublon", dup, "menage_id",
                "Le même numéro de sac apparaît plusieurs fois")
    if {"nb_total", "nb_present"} <= set(df.columns):
        _ajoute(a, "Bloquant", Q, "Présents supérieurs à l'effectif",
                df[df["nb_present"] > df["nb_total"]], "menage_id",
                lambda r: f"{r['nb_present']:.0f} présents pour {r['nb_total']:.0f} résidents")
        _ajoute(a, "À vérifier", Q, "Ménage de taille inhabituelle",
                df[df["nb_total"] > 40], "menage_id",
                lambda r: f"{r['nb_total']:.0f} personnes déclarées")
    if "consentement" in df:
        _ajoute(a, "Information", Q, "Refus de participation",
                df[df["consentement"].astype(str).str.lower().str.startswith("non")],
                "menage_id", "Le ménage a refusé, le sac ne doit pas être compté")
    if "commune" in df:
        _ajoute(a, "Bloquant", Q, "Commune hors périmètre",
                df[~df["commune"].isin(COMMUNES)], "menage_id",
                lambda r: f"Commune non reconnue : {r['commune']}")
        _ajoute(a, "À vérifier", Q, "Commune sans dépôt de sacs",
                df[df["commune"].isin(COMMUNES_SITE)], "menage_id",
                lambda r: f"{r['commune']} est en prélèvement sur site de collecte")
    return pd.DataFrame(a)


def controler_carac(df, pesees=None):
    """Controles de la fiche de caracterisation a 13 categories."""
    a = []
    if df is None or df.empty:
        return pd.DataFrame(a)
    Q = "Caractérisation"
    if {"masse_echantillon", "masse_quart"} <= set(df.columns):
        _ajoute(a, "Bloquant", Q, "Quart supérieur à l'échantillon",
                df[df["masse_quart"] > df["masse_echantillon"] * 1.001],
                "code_echantillon",
                lambda r: f"{r['masse_quart']:.1f} kg pour un échantillon de "
                          f"{r['masse_echantillon']:.1f} kg")
    if {"masse_triee", "masse_quart"} <= set(df.columns):
        t = df.assign(_e=(df["masse_triee"] - df["masse_quart"])
                      / df["masse_quart"].replace(0, pd.NA))
        _ajoute(a, "Bloquant", Q, "Masse triée hors tolérance",
                t[t["_e"].abs() > 0.05], "code_echantillon",
                lambda r: f"écart de {r['_e'] * 100:+.1f} % avec la masse du quart")
    if {"masse_triee", "masse_triee_calc"} <= set(df.columns):
        d = (df["masse_triee"] - df["masse_triee_calc"]).abs()
        _ajoute(a, "Information", Q, "Total calculé différent du total transmis",
                df[d > 0.05], "code_echantillon",
                lambda r: f"{r['masse_triee']:.2f} kg transmis contre "
                          f"{r['masse_triee_calc']:.2f} kg recalculés")
    if {"sacs_distribues", "sacs_collectes"} <= set(df.columns):
        _ajoute(a, "Bloquant", Q, "Plus de sacs collectés que distribués",
                df[df["sacs_collectes"] > df["sacs_distribues"]], "code_echantillon",
                lambda r: f"{r['sacs_collectes']:.0f} collectés pour "
                          f"{r['sacs_distribues']:.0f} distribués")
        t = df.assign(_t=df["sacs_collectes"]
                      / df["sacs_distribues"].replace(0, pd.NA))
        _ajoute(a, "À vérifier", Q, "Récupération des sacs faible",
                t[t["_t"] < 0.7], "code_echantillon",
                lambda r: f"{r['_t'] * 100:.0f} % des sacs récupérés")
    if {"date_distribution", "date_recuperation"} <= set(df.columns):
        _ajoute(a, "Bloquant", Q, "Récupération avant distribution",
                df[df["date_recuperation"] < df["date_distribution"]],
                "code_echantillon", "Les deux dates sont incohérentes")
    if "masse_echantillon" in df:
        _ajoute(a, "À vérifier", Q, "Échantillon de faible masse",
                df[df["masse_echantillon"] < 20], "code_echantillon",
                lambda r: f"{r['masse_echantillon']:.1f} kg, en dessous de 20 kg")
    if "cat_13" in df and "masse_triee" in df:
        t = df.assign(_p=df["cat_13"] / df["masse_triee"].replace(0, pd.NA))
        _ajoute(a, "À vérifier", Q, "Part de fines élevée",
                t[t["_p"] > 0.3], "code_echantillon",
                lambda r: f"les fines représentent {r['_p'] * 100:.0f} % "
                          "de la masse triée")
    if "commune" in df:
        _ajoute(a, "Bloquant", Q, "Commune hors périmètre",
                df[~df["commune"].isin(COMMUNES)], "code_echantillon",
                lambda r: f"Commune non reconnue : {r['commune']}")
    if "code_echantillon" in df:
        dup = df[df.duplicated("code_echantillon", keep=False)
                 & df["code_echantillon"].notna()]
        _ajoute(a, "Bloquant", Q, "Code d'échantillon en doublon", dup,
                "code_echantillon", "Le même code apparaît plusieurs fois")
    return pd.DataFrame(a)


# =====================================================================
# INDICATEURS THEMATIQUES
# =====================================================================
def _txt(serie):
    return serie.astype(str).str.strip().str.lower()


def _est_oui(df, col):
    if col not in df:
        return None
    return _txt(df[col]).isin(["oui", "1", "true", "yes"])


def _contient(df, col, *motifs):
    """Vrai si une reponse a choix multiples contient l'un des motifs."""
    if col not in df:
        return None
    s = _txt(df[col])
    m = pd.Series(False, index=df.index)
    for mot in motifs:
        m |= s.str.contains(mot, regex=False, na=False)
    return m


def _part(masque):
    if masque is None or len(masque) == 0:
        return None
    return float(masque.mean() * 100)


def _moyenne(df, col):
    if col not in df:
        return None
    v = pd.to_numeric(df[col], errors="coerce")
    return float(v.mean()) if v.notna().any() else None


# famille, libelle, unite, sens (True = plus c'est haut mieux c'est)
INDICATEURS = [
    ("Service de collecte", "Ménages desservis", "%", True),
    ("Service de collecte", "Collecte assurée par la SONAGED", "%", True),
    ("Service de collecte", "Collecte par un opérateur informel ou aucun", "%", False),
    ("Service de collecte", "Collecte irrégulière", "%", False),
    ("Stockage", "Poubelle avec couvercle", "%", True),
    ("Stockage", "Sac plastique ou aucun récipient", "%", False),
    ("Pratiques d'élimination", "Brûlage des déchets", "%", False),
    ("Pratiques d'élimination", "Enfouissement", "%", False),
    ("Pratiques d'élimination", "Dépôt sauvage", "%", False),
    ("Pratiques d'élimination", "Compostage ou alimentation du bétail", "%", True),
    ("Pratiques d'élimination", "Réutilisation, don ou vente", "%", True),
    ("Tri et disposition", "Tri déjà pratiqué à la source", "%", True),
    ("Tri et disposition", "Disposés à trier", "%", True),
    ("Tri et disposition", "Disposés à trier sans condition", "%", True),
    ("Agro-pastoral", "Ménages pratiquant l'élevage", "%", None),
    ("Agro-pastoral", "Cheptel moyen des éleveurs", "têtes", None),
    ("Agro-pastoral", "Fumier rejeté ou non géré", "%", False),
    ("Agro-pastoral", "Production variant selon la saison", "%", None),
    ("Ménage", "Taille moyenne du ménage", "pers.", None),
    ("Ménage", "Personnes présentes pendant l'enquête", "pers.", None),
]


def calcul_indicateurs(df):
    """Dictionnaire indicateur -> valeur, pour un sous-ensemble de fiches."""
    if df is None or df.empty:
        return {}
    eleveurs = df[_est_oui(df, "elevage")] if "elevage" in df else df.iloc[0:0]
    return {
        "Ménages desservis": _part(_est_oui(df, "service")),
        "Collecte assurée par la SONAGED": _part(_contient(df, "operateur", "sonaged")),
        "Collecte par un opérateur informel ou aucun":
            _part(_contient(df, "operateur", "informel", "aucun")),
        "Collecte irrégulière": _part(_contient(df, "frequence", "irregul", "irrégul")),
        "Poubelle avec couvercle": _part(_contient(df, "recipient", "poubelle")),
        "Sac plastique ou aucun récipient":
            _part(_contient(df, "recipient", "sac", "autre")),
        "Brûlage des déchets": _part(_contient(df, "pratiques", "brulage", "brûlage")),
        "Enfouissement": _part(_contient(df, "pratiques", "enfouissement")),
        "Dépôt sauvage": _part(_contient(df, "pratiques", "depot_sauvage",
                                         "dépôt sauvage")),
        "Compostage ou alimentation du bétail":
            _part(_contient(df, "pratiques", "compostage", "betail", "bétail")),
        "Réutilisation, don ou vente":
            _part(_contient(df, "pratiques", "reutilisation", "don", "vente")),
        "Tri déjà pratiqué à la source": _part(_est_oui(df, "tri_source")),
        "Disposés à trier": _part(_contient(df, "dispose_trier", "oui",
                                            "sous_conditions", "sous conditions")),
        "Disposés à trier sans condition": _part(_est_oui(df, "dispose_trier")),
        "Ménages pratiquant l'élevage": _part(_est_oui(df, "elevage")),
        "Cheptel moyen des éleveurs": _moyenne(eleveurs, "effectif"),
        "Fumier rejeté ou non géré": _part(_contient(df, "fumier", "rejete", "rejeté",
                                                     "non géré")),
        "Production variant selon la saison": _part(_est_oui(df, "saison_var")),
        "Taille moyenne du ménage": _moyenne(df, "nb_total"),
        "Personnes présentes pendant l'enquête": _moyenne(df, "nb_present"),
    }


def tableau_thematique(df, dimension):
    """Indicateurs calcules par modalite de la dimension choisie."""
    if df is None or df.empty or dimension not in df:
        return pd.DataFrame()
    lignes = []
    for modalite, sous in df.groupby(dimension):
        d = calcul_indicateurs(sous)
        d = {"Groupe": modalite, "Ménages enquêtés": len(sous), **d}
        lignes.append(d)
    ens = calcul_indicateurs(df)
    lignes.append({"Groupe": "Ensemble", "Ménages enquêtés": len(df), **ens})
    cols = ["Groupe", "Ménages enquêtés"] + [n for _, n, _, _ in INDICATEURS]
    return pd.DataFrame(lignes)[[c for c in cols if c in pd.DataFrame(lignes).columns]]


# =====================================================================
# INTERFACE
# =====================================================================
exiger_mot_de_passe()

p_logo = logo()
if p_logo:
    st.sidebar.image(str(p_logo), width="stretch")
else:
    st.sidebar.markdown(
        f"<div style='background:{VERT};color:#fff;padding:10px 12px;border-radius:6px;"
        "font-weight:700;letter-spacing:1px;text-align:center'>SGP</div>",
        unsafe_allow_html=True)
st.sidebar.caption("Caractérisation des déchets ménagers, Pôle Nord")
if st.sidebar.button("Se déconnecter", width="stretch"):
    st.session_state[CLE_AUTH] = False
    st.rerun()
st.sidebar.divider()

bandeau("Suivi de la campagne de Caractérisation Pôle Nord",
        "12 communes des régions de Saint-Louis et de Matam - avancement, "
        "qualité des données et composition des déchets")

# --- questionnaires, facultatifs au demarrage
f_socio = f_carac = pesees = None
f_sites_agro = f_tri_agro = None
p1, p2 = chemin(FORM_SOCIO), chemin(FORM_CARAC)
manquants = [n for n, p in [(FORM_SOCIO, p1), (FORM_CARAC, p2)] if p is None]
if p1:
    try:
        f_socio = charger_form(p1)
    except Exception as e:
        st.warning(f"Lecture du questionnaire socio-démographique impossible : {e}")
if p2:
    try:
        f_carac = charger_form(p2)
        pesees = table_pesees(f_carac)
    except Exception as e:
        st.warning(f"Lecture du questionnaire de caractérisation impossible : {e}")
f_sites_agro = f_tri_agro = None
for nom, cible in ((FORM_SITES_AGRO, "f_sites_agro"), (FORM_TRI_AGRO, "f_tri_agro")):
    pc = chemin(nom)
    if pc:
        try:
            globals()[cible] = charger_form(pc)
        except Exception as e:
            st.warning(f"Lecture de « {nom} » impossible : {e}")
if manquants:
    st.info("Fichiers absents du dépôt : " + ", ".join(f"`{m}`" for m in manquants) +
            ". Les déposer dans un dossier `donnees` pour activer les pesées et la "
            "structure des questionnaires. Le reste fonctionne sans eux.")

# --- source des donnees : KoboToolbox en direct, sans autre option
sec = secrets_kobo()
jeux, journal, derniere = {}, pd.DataFrame(), None

st.sidebar.header("KoboToolbox")

if sec.get("token"):
    url = sec.get("url") or KOBO_URL
    token = sec["token"]
    uids = sec.get("uids") or UIDS_DEFAUT
    st.sidebar.caption(f"Connexion directe à {url}")
    st.sidebar.caption(f"{len(uids)} formulaires interrogés")
else:
    st.sidebar.error(
        "Le jeton d'API n'est pas dans les secrets. Ajouter ce bloc dans "
        "Manage app, Settings, Secrets :\n\n"
        "```toml\n[kobo]\nurl = \"https://kf.kobotoolbox.org\"\n"
        "token = \"votre_jeton\"\nuid_1 = \"...\"\nuid_2 = \"...\"\n```")
    st.sidebar.caption("En attendant, saisir les paramètres ici.")
    url = st.sidebar.text_input("Serveur Kobo", KOBO_URL)
    token = st.sidebar.text_input("Jeton d'API", "", type="password")
    saisis = st.sidebar.text_area("Identifiants des formulaires, un par ligne",
                                  "\n".join(UIDS_DEFAUT), height=140)
    uids = [x.strip() for x in saisis.splitlines() if x.strip()]

if st.sidebar.button("Actualiser maintenant", width="stretch",
                     help="Force une nouvelle lecture, sans attendre les 5 minutes"):
    recuperer_kobo.clear()
    st.rerun()

if token:
    try:
        jeux, journal, derniere = recuperer_kobo(url, token, tuple(uids))
    except Exception as e:
        st.sidebar.error(f"Échec de la récupération : {e}")
        st.sidebar.caption("Vérifier le jeton d'API et les identifiants de "
                           "formulaire dans les secrets.")

socio = normalise_socio(jeux.get(ROLE_MENAGES))
carac = normalise_carac(jeux.get(ROLE_TRI))
sites_agro = normalise_sites_agro(jeux.get(ROLE_SITES_AGRO))
tri_agro = normalise_tri_agro(jeux.get(ROLE_TRI_AGRO))
photos = normalise_photos(jeux.get(ROLE_PHOTOS))

# Toutes les images, quel que soit le formulaire qui les porte.
images = table_photos({
    "Reportage photo": photos,
    "Enquête ménage": socio,
    "Caractérisation": carac,
    "Sites agro-pastoraux": sites_agro,
    "Tri agro-pastoral": tri_agro,
})

if derniere is not None:
    age = int((pd.Timestamp.now() - derniere).total_seconds())
    st.sidebar.success(
        f"{len(socio)} fiches ménage, {len(carac)} fiches de tri\n\n"
        f"{len(sites_agro)} sites agro, {len(tri_agro)} tris agro\n\n"
        f"Lecture du {derniere:%H:%M:%S}"
        + (f", il y a {age // 60} min" if age >= 60 else ", à l'instant"))
    st.sidebar.caption(f"Relecture de Kobo toutes les {DUREE_CACHE // 60} minutes, "
                       "à la première interaction avec la page.")
    with st.sidebar.expander("Formulaires interrogés"):
        st.dataframe(journal, hide_index=True, width="stretch")

if all(x.empty for x in (socio, carac, sites_agro, tri_agro)):
    st.sidebar.info("Aucune soumission sur Kobo pour l'instant.")

# --- indicateurs generaux
cible_totale = CIBLE_TOTALE
c1, c2, c3, c4 = st.columns(4)
part = len(socio) / cible_totale * 100 if cible_totale else 0
fiche(c1, "Objectif de sacs à distribuer", f"{cible_totale}",
      f"{CIBLE_MENAGES} par commune sur {len(COMMUNES_SOURCE)} communes à la source",
      "neutre")
fiche(c2, "Fiches ménage reçues", f"{len(socio)}", f"{part:.0f} % de l'objectif",
      "normal" if part >= 50 else "veille")
fiche(c3, "Échantillons triés", f"{len(carac)}", "questionnaire de caractérisation")
vues = set(socio["commune"].dropna()) if not socio.empty else set()
vues |= set(carac["commune"].dropna()) if not carac.empty else set()
nb_com = len(vues & set(COMMUNES))
fiche(c4, "Communes couvertes", f"{nb_com} / {len(COMMUNES)}",
      "au moins une fiche ménage ou de tri", "normal" if nb_com else "neutre")
st.write("")

onglets = st.tabs(["Avancement", "Carte", "Analyse thématique", "Qualité",
                   "Composition", "Agro-pastoral", "Photos", "Questionnaires"])

# ---------------------------------------------------------------- AVANCEMENT
with onglets[0]:
    lignes = []
    for commune, info in COMMUNES.items():
        recu = int((socio["commune"] == commune).sum()) if not socio.empty else 0
        trie = int((carac["commune"] == commune).sum()) if not carac.empty else 0
        cible = cible_commune(commune)
        lignes.append({"Commune": commune, "Région": info["region"],
                       "Infrastructure": info["infra"], "Méthode": info["methode"],
                       "Objectif sacs": cible, "Fiches reçues": recu,
                       "Marge restante": max(cible - recu, 0),
                       "Avancement": (recu / cible * 100) if cible else None,
                       "Échantillons triés": trie})
    suivi = pd.DataFrame(lignes)
    tot_cible = suivi["Objectif sacs"].sum()
    tot_recu = suivi.loc[suivi["Objectif sacs"] > 0, "Fiches reçues"].sum()

    a, b, c = st.columns(3)
    pc = tot_recu / tot_cible * 100 if tot_cible else 0
    fiche(a, "Objectif total", f"{tot_cible}", "sacs au maximum", "neutre")
    fiche(b, "Reçu", f"{tot_recu}", f"{pc:.0f} % de l'objectif",
          "normal" if pc >= 50 else "veille")
    fiche(c, "Marge restante", f"{max(tot_cible - tot_recu, 0)}", "sacs possibles",
          "veille" if tot_cible - tot_recu > 0 else "normal")
    st.write("")

    try:
        import plotly.graph_objects as go
        src = suivi[suivi["Objectif sacs"] > 0].sort_values("Avancement")
        fig = go.Figure()
        fig.add_bar(y=src["Commune"], x=src["Objectif sacs"], orientation="h",
                    marker_color="#DCDCD4", hoverinfo="skip")
        fig.add_bar(y=src["Commune"], x=src["Fiches reçues"], orientation="h",
                    marker_color=[COULEURS[m] for m in src["Méthode"]],
                    hovertemplate="%{y} : %{x} fiches<extra></extra>")
        fig.add_vline(x=CIBLE_MENAGES, line_dash="dot", line_color="#B01B2E",
                      annotation_text=f"cible {CIBLE_MENAGES}")
        fig.update_layout(barmode="overlay", height=420, showlegend=False,
                          margin=dict(l=0, r=0, t=10, b=0), xaxis_title="Ménages",
                          plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, width="stretch")
    except Exception:
        st.bar_chart(suivi.set_index("Commune")["Fiches reçues"])

    st.dataframe(suivi, hide_index=True, width="stretch",
                 column_config={"Avancement": st.column_config.ProgressColumn(
                     "Avancement", format="%.0f %%", min_value=0, max_value=100)})
    st.caption(
        f"{CIBLE_MENAGES} sacs par commune est un maximum, pas un quota à atteindre. "
        "Un effectif de 60 ou moins reste conforme : il traduit la réalité du terrain "
        "et non un incident de collecte. Ranérou est en régime allégé, maximum "
        f"{CIBLES_PARTICULIERES['Ranérou']}. Dagana et Bokhol sont en prélèvement sur "
        "site de collecte, la décharge, sans sacs. Ndioum combine les deux "
        "dispositifs.")
    st.download_button("Télécharger le tableau de suivi",
                       suivi.to_csv(index=False).encode("utf-8"),
                       "suivi_avancement.csv", "text/csv")

# ---------------------------------------------------------------- CARTE
FONDS = {
    "Plan OpenStreetMap": {
        "tuiles": "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
        "credit": "© OpenStreetMap contributors"},
    "Image satellite": {
        "tuiles": "https://server.arcgisonline.com/ArcGIS/rest/services/"
                  "World_Imagery/MapServer/tile/{z}/{y}/{x}",
        "credit": "Esri, Maxar, Earthstar Geographics"},
    "Relief clair": {
        "tuiles": "https://{s}.basemaps.cartocdn.com/rastertiles/voyager/"
                  "{z}/{x}/{y}{r}.png",
        "credit": "© OpenStreetMap contributors, © CARTO"},
    "Fond sombre": {
        "tuiles": "https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png",
        "credit": "© OpenStreetMap contributors, © CARTO"},
}


with onglets[1]:
    p_gpkg = chemin(FICHIER_GPKG)
    if p_gpkg is None:
        attente(f"Le fond de carte `{FICHIER_GPKG}` n'est pas dans le dépôt.<br>"
                "Le déposer dans un dossier `donnees` pour activer cette vue.")
    else:
        try:
            import geopandas as gpd

            @st.cache_data(show_spinner=False)
            def charger_couches(p):
                com = gpd.read_file(p, layer="communes")
                try:
                    sit = gpd.read_file(p, layer="sites")
                except Exception:
                    sit = None
                return com, sit

            communes_geo, sites = charger_couches(str(p_gpkg))
            recus = socio.groupby("commune").size() if not socio.empty else pd.Series(dtype=int)
            communes_geo["recus"] = communes_geo["commune"].map(recus).fillna(0).astype(int)
            communes_geo["cible"] = communes_geo["commune"].apply(cible_commune)
            communes_geo["avancement"] = [
                round(r / c * 100) if c else None
                for r, c in zip(communes_geo["recus"], communes_geo["cible"])]

            c1, c2 = st.columns([2, 3])
            fond = c1.selectbox("Fond de carte", list(FONDS), index=0)
            variable = c2.radio("Colorer les communes par",
                                ["Méthode", "Avancement", "Population 2023"],
                                horizontal=True)

            def couleur(r):
                if variable == "Méthode":
                    return COULEURS.get(r["methode"], "#888888")
                if variable == "Avancement":
                    a = r["avancement"]
                    if a is None:
                        return "#5A6672"
                    return "#B01B2E" if a < 34 else ("#F4A93B" if a < 67 else "#1B7F4B")
                pop = r.get("pop_2023") or 0
                bornes = [10000, 25000, 50000, 70000]
                palette = ["#DCEBD2", "#A9D18E", "#6FBF73", "#2E9B57", "#146B3A"]
                for i, b in enumerate(bornes):
                    if pop < b:
                        return palette[i]
                return palette[-1]

            try:
                import folium
                from streamlit_folium import st_folium

                b = communes_geo.total_bounds
                carte = folium.Map(
                    location=[(b[1] + b[3]) / 2, (b[0] + b[2]) / 2],
                    zoom_start=8, tiles=None, control_scale=True)
                for nom, f in FONDS.items():
                    folium.TileLayer(f["tuiles"], name=nom, attr=f["credit"],
                                     overlay=False,
                                     show=(nom == fond)).add_to(carte)

                geo = communes_geo.to_crs(4326).copy()
                geo["couleur"] = [couleur(r) for _, r in geo.iterrows()]
                geo["_av"] = geo["avancement"].map(
                    lambda v: "sans cible" if v is None else f"{v} %")
                folium.GeoJson(
                    geo.__geo_interface__, name="Communes",
                    style_function=lambda x: {
                        "fillColor": x["properties"]["couleur"], "color": "#FFFFFF",
                        "weight": 1.2, "fillOpacity": 0.55},
                    highlight_function=lambda x: {"weight": 3, "fillOpacity": 0.75},
                    tooltip=folium.GeoJsonTooltip(
                        fields=[c for c in ["commune", "methode", "pop_2023", "menages",
                                            "recus", "_av"] if c in geo.columns],
                        aliases=["Commune", "Méthode", "Population", "Ménages",
                                 "Fiches reçues", "Avancement"][:len(
                            [c for c in ["commune", "methode", "pop_2023", "menages",
                                         "recus", "_av"] if c in geo.columns])],
                        sticky=True),
                ).add_to(carte)

                if sites is not None and len(sites):
                    grp = folium.FeatureGroup(name="Sites de tri", show=True)
                    for _, r in sites.to_crs(4326).iterrows():
                        folium.CircleMarker(
                            [r.geometry.y, r.geometry.x], radius=7,
                            color="#FFFFFF", weight=2,
                            fill_color=COULEURS.get(r["methode"], "#888"),
                            fill_opacity=1,
                            popup=f"<b>{r['commune']}</b><br>{r['methode']}",
                            tooltip=f"Site de {r['commune']}").add_to(grp)
                    grp.add_to(carte)

                if not sites_agro.empty and "lat" in sites_agro.columns:
                    grp = folium.FeatureGroup(name="Sites agro-pastoraux", show=True)
                    pts = sites_agro.dropna(subset=["lat", "lon"])
                    for _, r in pts.iterrows():
                        texte = f"<b>{r.get('site', 'site')}</b><br>{r.get('commune', '')}"
                        if r.get("type_activite"):
                            texte += f"<br>{r['type_activite']}"
                        folium.Marker(
                            [r["lat"], r["lon"]], popup=texte,
                            tooltip=str(r.get("site", "site agro-pastoral")),
                            icon=folium.Icon(color="green", icon="leaf")).add_to(grp)
                    grp.add_to(carte)
                    st.caption(f"{len(pts)} sites agro-pastoraux géolocalisés.")

                folium.LayerControl(collapsed=False).add_to(carte)
                st_folium(carte, width=None, height=560,
                          returned_objects=[], key="carte_suivi")
                st.caption(FONDS[fond]["credit"] +
                           ". Limites communales : OpenStreetMap. "
                           "Population : ANSD, RGPH-5 2023.")
            except ImportError:
                st.warning("Les paquets `folium` et `streamlit-folium` ne sont pas "
                           "installés. Les ajouter à requirements.txt pour la carte "
                           "interactive.")

            rubrique("Détail par commune")
            cols = [c for c in ["commune", "region", "departement", "methode",
                                "logist_cat", "pop_2023", "menages", "recus", "cible",
                                "avancement"] if c in communes_geo.columns]
            st.dataframe(communes_geo[cols].sort_values("commune"),
                         hide_index=True, width="stretch")
        except Exception as e:
            st.error(f"Rendu de la carte impossible : {e}")


# ---------------------------------------------------------------- THEMATIQUE
with onglets[2]:
    if socio.empty:
        attente("Cette vue s'appuie sur les fiches ménage du questionnaire "
                "socio-démographique.<br>Aucune n'est encore arrivée de Kobo.")
    else:
        dims = {"Commune": "commune", "Région": "region",
                "Strate (standing)": "strate", "Type de logement": "logement",
                "Revenu du ménage": "revenu"}
        dispo = {k: v for k, v in dims.items() if v in socio.columns}
        if not dispo:
            st.warning("Les colonnes de comparaison sont absentes des fiches.")
        else:
            dim_label = st.radio("Comparer par", list(dispo), horizontal=True)
            dim = dispo[dim_label]
            tab = tableau_thematique(socio, dim)

            ens = calcul_indicateurs(socio)
            k = st.columns(4)
            resume_cles = [("Ménages desservis", "normal"),
                           ("Brûlage des déchets", "alerte"),
                           ("Tri déjà pratiqué à la source", "normal"),
                           ("Disposés à trier", "normal")]
            for i, (lib, ton) in enumerate(resume_cles):
                val = ens.get(lib)
                fiche(k[i], lib, "n.d." if val is None else f"{val:.0f} %",
                      f"sur {len(socio)} fiches", ton if val is not None else "neutre")
            st.write("")
            st.divider()

            familles = []
            for fam, nom, unite, sens in INDICATEURS:
                if fam not in familles:
                    familles.append(fam)
            choix_fam = st.multiselect("Thèmes", familles, default=familles)
            noms = [n for f, n, u, s in INDICATEURS if f in choix_fam]

            for fam in choix_fam:
                rubrique(fam)
                sous = [(n, u, s) for f, n, u, s in INDICATEURS if f == fam]
                for nom, unite, sens in sous:
                    if nom not in tab.columns:
                        continue
                    serie = tab[tab["Groupe"] != "Ensemble"][["Groupe", nom]].dropna()
                    if serie.empty:
                        continue
                    valeur_ens = ens.get(nom)
                    g, d = st.columns([3, 1])
                    with g:
                        try:
                            import plotly.express as px
                            fig = px.bar(serie.sort_values(nom), x=nom, y="Groupe",
                                         orientation="h", text=serie.sort_values(nom)[nom]
                                         .map(lambda v: f"{v:,.0f}".replace(",", " ")))
                            coul = VERT if sens is not False else "#B01B2E"
                            fig.update_traces(marker_color=coul, textposition="outside",
                                              cliponaxis=False)
                            if valeur_ens is not None:
                                fig.add_vline(x=valeur_ens, line_dash="dot",
                                              line_color="#666",
                                              annotation_text="ensemble")
                            fig.update_layout(height=max(220, 34 * len(serie)),
                                              margin=dict(l=0, r=50, t=26, b=0),
                                              title=f"{nom} ({unite})",
                                              xaxis_title="", yaxis_title="",
                                              plot_bgcolor="rgba(0,0,0,0)")
                            st.plotly_chart(fig, width="stretch")
                        except Exception:
                            st.write(f"**{nom}** ({unite})")
                            st.bar_chart(serie.set_index("Groupe")[nom])
                    with d:
                        st.metric("Ensemble",
                                  "n.d." if valeur_ens is None else f"{valeur_ens:,.0f}"
                                  .replace(",", " "))
                        ecart = serie[nom].max() - serie[nom].min()
                        st.caption(f"Écart entre groupes : {ecart:,.0f} {unite}"
                                   .replace(",", " "))

            st.divider()
            rubrique("Tableau complet")
            st.dataframe(tab.round(1), hide_index=True, width="stretch")
            st.download_button("Télécharger les indicateurs",
                               tab.to_csv(index=False).encode("utf-8"),
                               f"indicateurs_{dim}.csv", "text/csv")
            st.caption(
                "Les pourcentages portent sur les fiches où la question a été posée. "
                "Un indicateur reste vide tant que la question correspondante n'a reçu "
                "aucune réponse exploitable."
            )
            galerie(socio, token, "Photos jointes aux fiches ménage")


# ---------------------------------------------------------------- QUALITE
with onglets[3]:
    if socio.empty and carac.empty:
        attente("Aucune soumission reçue de Kobo pour l'instant.<br>"
                "Les contrôles s'activeront dès les premières fiches.")
    else:
        a1 = controler_socio(socio)
        a2 = controler_carac(carac, pesees)
        anomalies = pd.concat([a1, a2], ignore_index=True) if len(a1) or len(a2) else pd.DataFrame()
        nb = lambda g: int((anomalies["gravité"] == g).sum()) if not anomalies.empty else 0  # noqa: E731
        q1, q2, q3, q4 = st.columns(4)
        fiche(q1, "Soumissions contrôlées", f"{len(socio) + len(carac)}",
              "fiches ménage et fiches de tri", "neutre")
        fiche(q2, "Bloquantes", f"{nb('Bloquant')}", "à corriger avant analyse",
              "alerte" if nb("Bloquant") else "normal")
        fiche(q3, "À vérifier", f"{nb('À vérifier')}", "à confirmer sur le terrain",
              "veille" if nb("À vérifier") else "normal")
        fiche(q4, "Informations", f"{nb('Information')}", "sans action requise", "neutre")
        st.write("")

        if nb("Bloquant"):
            st.error(f"{nb('Bloquant')} anomalies bloquantes à corriger avant analyse.")
        elif not anomalies.empty:
            st.warning("Aucune anomalie bloquante. Des points restent à vérifier.")
        else:
            st.success("Aucune anomalie détectée.")

        if not socio.empty and not carac.empty:
            st.caption("La fiche de caractérisation ne porte pas d'identifiant de "
                       "ménage : le croisement avec l'enquête ménage se fait au "
                       "niveau de la commune et de la strate.")

        if not anomalies.empty:
            g, d = st.columns([2, 3])
            with g:
                rubrique("Par contrôle")
                st.dataframe(anomalies.groupby(["gravité", "contrôle"]).size()
                             .reset_index(name="Nombre").sort_values("Nombre", ascending=False),
                             hide_index=True, width="stretch")
            with d:
                rubrique("Par commune")
                st.dataframe(anomalies.groupby(["commune", "gravité"]).size()
                             .unstack(fill_value=0), width="stretch")
            rubrique("Détail")
            st.dataframe(anomalies, hide_index=True, width="stretch", height=380)
            st.download_button("Télécharger les anomalies",
                               anomalies.to_csv(index=False).encode("utf-8"),
                               "anomalies_modecom.csv", "text/csv")

# ---------------------------------------------------------------- COMPOSITION
with onglets[4]:
    if carac.empty:
        attente("Cette vue s'appuie sur les fiches de caractérisation.<br>"
                "Aucune n'est encore arrivée de Kobo.")
    else:
        f1, f2, f3 = st.columns([3, 2, 2])
        dispo = sorted(carac["commune"].dropna().unique())
        choix = f1.multiselect("Communes", dispo, default=dispo, key="comp_communes")
        strates = sorted(carac["strate"].dropna().unique()) if "strate" in carac else []
        ch_str = f2.multiselect("Strates", strates, default=strates,
                                key="comp_strates") if strates else []
        modes = sorted(carac["methode"].dropna().unique()) if "methode" in carac else []
        ch_mod = f3.multiselect("Méthodes", modes, default=modes,
                                key="comp_modes") if modes else []

        sel = carac[carac["commune"].isin(choix)]
        if ch_str:
            sel = sel[sel["strate"].isin(ch_str)]
        if ch_mod:
            sel = sel[sel["methode"].isin(ch_mod)]

        if sel.empty:
            st.info("Aucune fiche ne correspond à cette sélection.")
        else:
            cols_cat = [f"cat_{n}" for n in NOMENCLATURE if f"cat_{n}" in sel.columns]
            par_cat = sel[cols_cat].sum(min_count=1)
            total = float(par_cat.sum())

            m1, m2, m3 = st.columns(3)
            fiche(m1, "Échantillons", f"{len(sel)}", "fiches retenues", "neutre")
            if "masse_echantillon" in sel:
                fiche(m2, "Masse d'échantillon",
                      f"{sel['masse_echantillon'].sum():,.0f}".replace(",", " ") + " kg",
                      "avant quartage")
            fiche(m3, "Masse triée", f"{total:,.1f}".replace(",", " ") + " kg",
                  "somme des 13 catégories")
            st.write("")

            comp = pd.DataFrame({
                "Catégorie": [f"{n}. {NOMENCLATURE[n]['titre']}"
                              for n in NOMENCLATURE if f"cat_{n}" in sel.columns],
                "Masse (kg)": par_cat.values,
            })
            comp["Part (%)"] = comp["Masse (kg)"] / total * 100 if total else 0
            comp = comp.sort_values("Part (%)", ascending=False)

            try:
                import plotly.express as px
                ordre = comp.sort_values("Part (%)")
                fig = px.bar(ordre, x="Part (%)", y="Catégorie", orientation="h",
                             text=ordre["Part (%)"].map(lambda v: f"{v:.1f} %"))
                fig.update_traces(marker_color=VERT, textposition="outside",
                                  cliponaxis=False)
                fig.update_layout(height=520, margin=dict(l=0, r=50, t=10, b=0),
                                  plot_bgcolor="rgba(0,0,0,0)", yaxis_title="")
                st.plotly_chart(fig, width="stretch")
            except Exception:
                st.bar_chart(comp.set_index("Catégorie")["Part (%)"])

            st.dataframe(comp.round(2), hide_index=True, width="stretch", height=340)

            rubrique("Par fraction granulométrique")
            frac = {}
            for cle, lib in (("total_het", "Hétéroclites"),
                             ("total_g100", "Plus de 100 mm"),
                             ("total_m20", "100 à 20 mm"),
                             ("total_fines", "Fines, moins de 20 mm")):
                if cle in sel:
                    frac[lib] = float(sel[cle].sum())
            if frac:
                tf = pd.DataFrame({"Fraction": list(frac), "Masse (kg)": list(frac.values())})
                sf = tf["Masse (kg)"].sum()
                tf["Part (%)"] = tf["Masse (kg)"] / sf * 100 if sf else 0
                st.dataframe(tf.round(2), hide_index=True, width="stretch")
            else:
                st.caption("Les totaux par fraction ne figurent pas dans cet export.")

            rubrique("Détail des sous-catégories")
            lignes = []
            for num, v in NOMENCLATURE.items():
                for cle, libelle in v["sous"].items():
                    suffixes = ("poids",) if num == "13" else ("het", "g100", "m20")
                    cs = [f"{cle}_{f}" for f in suffixes if f"{cle}_{f}" in sel.columns]
                    if not cs:
                        continue
                    kg = float(sel[cs].sum().sum())
                    lignes.append({"Catégorie": f"{num}. {v['titre']}",
                                   "Sous-catégorie": libelle, "Masse (kg)": round(kg, 3),
                                   "Part (%)": round(kg / total * 100, 2) if total else 0})
            detail = pd.DataFrame(lignes).sort_values("Masse (kg)", ascending=False)
            st.dataframe(detail, hide_index=True, width="stretch", height=380)

            if strates and len(ch_str) > 1:
                rubrique("Composition par strate")
                lignes = []
                for st_nom, g in sel.groupby("strate"):
                    t = float(g[cols_cat].sum().sum())
                    for n in NOMENCLATURE:
                        c = f"cat_{n}"
                        if c in g.columns:
                            lignes.append({
                                "Strate": st_nom,
                                "Catégorie": f"{n}. {NOMENCLATURE[n]['titre']}",
                                "Part (%)": round(float(g[c].sum()) / t * 100, 2)
                                if t else 0})
                st.dataframe(pd.DataFrame(lignes).pivot(index="Catégorie",
                                                        columns="Strate",
                                                        values="Part (%)"),
                             width="stretch")
                st.caption("Comparaison en part de masse, les strates n'ayant pas "
                           "le même nombre d'échantillons.")

            st.download_button("Télécharger la composition",
                               detail.to_csv(index=False).encode("utf-8"),
                               "composition_modecom_13_categories.csv", "text/csv")

# ---------------------------------------------------------------- AGRO-PASTORAL
with onglets[5]:
    if sites_agro.empty and tri_agro.empty:
        attente("Aucun site agro-pastoral ni tri A1 à A6 reçu de Kobo.<br>"
                "Ces deux formulaires couvrent Ogo et Bokidiawé.")
    else:
        a1, a2, a3, a4 = st.columns(4)
        fiche(a1, "Sites recensés", f"{len(sites_agro)}", "zones de parcage, marchés…",
              "normal" if len(sites_agro) else "neutre")
        n_gps = int(sites_agro["lat"].notna().sum()) if "lat" in sites_agro else 0
        fiche(a2, "Sites géolocalisés", f"{n_gps}",
              "points GPS exploitables",
              "normal" if n_gps == len(sites_agro) and n_gps else "veille")
        fiche(a3, "Échantillons triés", f"{len(tri_agro)}", "fiches A1 à A6",
              "normal" if len(tri_agro) else "neutre")
        masse = tri_agro["masse_nette"].sum() if "masse_nette" in tri_agro else 0
        fiche(a4, "Masse nette triée", f"{masse:,.0f}".replace(",", " ") + " kg",
              "cumul des campagnes")
        st.write("")

        if not sites_agro.empty:
            rubrique("Sites recensés")
            g, d = st.columns([3, 2])
            with g:
                if "type_activite" in sites_agro:
                    compte = {}
                    for lib in ["parcage", "abattage", "marche", "abreuvement", "autre"]:
                        m = _contient(sites_agro, "type_activite", lib)
                        if m is not None and m.any():
                            compte[lib] = int(m.sum())
                    if compte:
                        libelles = {"parcage": "Zone de parcage",
                                    "abattage": "Abattage informel",
                                    "marche": "Marché à bétail",
                                    "abreuvement": "Point d'abreuvement",
                                    "autre": "Autre"}
                        act = pd.DataFrame({"Activité": [libelles[k] for k in compte],
                                            "Sites": list(compte.values())})
                        try:
                            import plotly.express as px
                            f = px.bar(act.sort_values("Sites"), x="Sites", y="Activité",
                                       orientation="h")
                            f.update_traces(marker_color=VERT_CLAIR)
                            st.plotly_chart(habiller(f, 300), width="stretch")
                        except Exception:
                            st.bar_chart(act.set_index("Activité"))
            with d:
                if "commune" in sites_agro:
                    st.dataframe(
                        sites_agro.groupby("commune").size().reset_index(name="Sites"),
                        hide_index=True, width="stretch")
                if "accessibilite" in sites_agro:
                    st.dataframe(
                        sites_agro.groupby("accessibilite").size()
                        .reset_index(name="Sites"), hide_index=True, width="stretch")

            cols = [c for c in ["commune", "site", "type_activite", "categories",
                                "accessibilite", "contact", "lat", "lon", "date"]
                    if c in sites_agro.columns]
            st.dataframe(sites_agro[cols], hide_index=True, width="stretch", height=300)
            st.download_button("Télécharger les sites",
                               sites_agro[cols].to_csv(index=False).encode("utf-8"),
                               "sites_agro_pastoraux.csv", "text/csv")

        if not tri_agro.empty:
            rubrique("Composition agro-pastorale, A1 à A6")
            lignes = []
            for cle, lib in CATEGORIES_AGRO.items():
                col = f"masse_{cle}"
                if col in tri_agro:
                    lignes.append({"Catégorie": lib,
                                   "Masse (kg)": float(tri_agro[col].sum()),
                                   "Teneur en eau (%)": tri_agro.get(f"te_{cle}",
                                                                     pd.Series(dtype=float)).mean()})
            if lignes:
                comp = pd.DataFrame(lignes)
                total = comp["Masse (kg)"].sum()
                comp["Part (%)"] = comp["Masse (kg)"] / total * 100 if total else 0
                g, d = st.columns([3, 2])
                with g:
                    try:
                        import plotly.express as px
                        f = px.bar(comp.sort_values("Masse (kg)"), x="Part (%)",
                                   y="Catégorie", orientation="h",
                                   text=comp.sort_values("Masse (kg)")["Part (%)"]
                                   .map(lambda v: f"{v:.1f} %"))
                        f.update_traces(marker_color=VERT_CLAIR, textposition="outside",
                                        cliponaxis=False)
                        st.plotly_chart(habiller(f, 340), width="stretch")
                    except Exception:
                        st.bar_chart(comp.set_index("Catégorie")["Part (%)"])
                with d:
                    st.dataframe(comp.round(1), hide_index=True, width="stretch")
                st.download_button("Télécharger la composition agro-pastorale",
                                   comp.to_csv(index=False).encode("utf-8"),
                                   "composition_agro.csv", "text/csv")
                st.caption("La teneur en eau conditionne le potentiel de compostage "
                           "et de méthanisation ; elle n'est renseignée que pour les "
                           "catégories qui la prévoient.")

        galerie(sites_agro, token, "Photos des sites agro-pastoraux")
        galerie(tri_agro, token, "Photos des campagnes de tri agro-pastoral")


# ---------------------------------------------------------------- QUESTIONNAIRES
with onglets[6]:
    rubrique("Photos de terrain")
    if images.empty:
        st.info("Aucune photo reçue. Les images apparaissent ici dès la première "
                "soumission portant une pièce jointe, quel que soit le formulaire.")
    else:
        a, b, c = st.columns(3)
        fiche(a, "Photos reçues", f"{len(images)}", "toutes sources confondues",
              "neutre")
        vues_ph = images["commune"].dropna().nunique()
        fiche(b, "Communes documentées", f"{vues_ph} / {len(COMMUNES)}",
              "au moins une photo", "normal" if vues_ph else "neutre")
        fiche(c, "Formulaires porteurs", f"{images['source'].nunique()}",
              "questionnaires avec images", "neutre")

        st.write("")
        f1, f2 = st.columns([3, 2])
        dispo = [x for x in COMMUNES if nb_photos_commune(images, x)]
        # Images dont la commune est vide ou hors des douze communes suivies.
        autres = int((~images["commune"].isin(list(COMMUNES))).sum())
        choix = f1.multiselect("Communes", dispo, default=dispo,
                               key="filtre_photos")
        srcs = sorted(images["source"].unique())
        choix_src = f2.multiselect("Formulaires", srcs, default=srcs,
                                   key="filtre_photos_source")

        sel = images[images["commune"].isin(choix) & images["source"].isin(choix_src)]
        st.caption(f"{len(sel)} photos correspondent aux filtres, sur "
                   f"{len(images)} reçues."
                   + (f" {autres} photos sans commune reconnue ne sont pas "
                      "affichées." if autres else ""))
        if sel.empty:
            st.info("Aucune photo pour cette sélection.")
        else:
            par_page = st.select_slider(
                "Photos affichées", options=[12, 24, 48, 96, 200],
                value=48 if len(sel) > 48 else 24, key="photos_par_page")
            pages = max(1, -(-len(sel) // par_page))
            page = 1
            if pages > 1:
                page = st.number_input("Page", 1, pages, 1, key="photos_page")
            debut = (int(page) - 1) * par_page
            n = galerie_table(sel.iloc[debut:debut + par_page], token)
            st.caption(f"{n} photos affichées, page {int(page)} sur {pages}, "
                       "les plus récentes d'abord. Les images restent hébergées "
                       "sur Kobo, rien n'est copié ici.")

        rubrique("Répartition par commune")
        recap = (images.assign(commune=images["commune"].fillna("non renseignée"))
                 .groupby(["commune", "source"]).size().unstack(fill_value=0))
        recap["Total"] = recap.sum(axis=1)
        st.dataframe(recap.sort_values("Total", ascending=False), width="stretch")
    st.caption("Le formulaire *Photo caractérisation* ne relève pas de coordonnées "
               "GPS : ses images sont rattachées à une commune, pas à un point de "
               "la carte.")


# ------------------------------------------------------------- QUESTIONNAIRES
with onglets[7]:
    if all(f is None for f in (f_socio, f_carac, f_sites_agro, f_tri_agro)):
        attente("Les deux questionnaires XLSForm ne sont pas dans le dépôt.")
    else:
        dispo = [n for n, f in [("Enquête socio-démographique", f_socio),
                                ("Caractérisation des déchets", f_carac),
                                ("Sites agro-pastoraux", f_sites_agro),
                                ("Caractérisation agro-pastorale", f_tri_agro)]
                 if f is not None]
        choix = st.radio("Questionnaire", dispo, horizontal=True)
        formulaires_dispo = {"Enquête socio-démographique": f_socio,
                             "Caractérisation des déchets": f_carac,
                             "Sites agro-pastoraux": f_sites_agro,
                             "Caractérisation agro-pastorale": f_tri_agro}
        form = formulaires_dispo[choix]
        r = resume(form)
        k1, k2, k3 = st.columns(3)
        fiche(k1, "Questions", f"{r['questions']}", "hors métadonnées", "neutre")
        fiche(k2, "Sections", f"{r['groupes']}", "groupes du formulaire", "neutre")
        fiche(k3, "Listes de choix", f"{r['listes']}", "modalités prédéfinies", "neutre")
        st.write("")
        q = form["questions"]
        st.dataframe(q.groupby("groupe").size().reset_index(name="Questions"),
                     hide_index=True, width="stretch")
        st.dataframe(q[["groupe", "type", "name", "label"]], hide_index=True,
                     width="stretch", height=380)
        if choix == "Caractérisation des déchets":
            n_sous = sum(len(v["sous"]) for v in NOMENCLATURE.values())
            st.write(f"{len(NOMENCLATURE)} catégories MODECOM, {n_sous} "
                     f"sous-catégories, {len(COLONNES_MASSE)} pesées au total. "
                     "Les fines portent une pesée unique, les autres "
                     "sous-catégories trois fractions granulométriques.")

st.markdown(
    "<div class='pied'>Suivi de la campagne de Caractérisation, Pôle Nord. "
    "La collecte se fait dans ODK Collect ou KoboCollect, hors connexion ; cette "
    "application lit les soumissions une fois synchronisées.<br>"
    f"Cible de {CIBLE_MENAGES} ménages par commune sur {len(COMMUNES_SOURCE)} communes "
    "en MODECOM à la source. Population de référence : ANSD, RGPH-5 2023.</div>",
    unsafe_allow_html=True)
