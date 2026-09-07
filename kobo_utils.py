# -*- coding: utf-8 -*-
"""Utilitaires : connexion API KoboToolbox + lecture des XLSForm de l'étude C40 GREEN+.

Deux formulaires :
  - quantitatif : "Évaluation des besoins des travailleurs informels ... Dakar"
  - qualitatif  : "Grille de saisie qualitative (entretiens et focus groups)"
"""
import os
import re
import random
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import requests

# --------------------------------------------------------------- Formulaires ---

FORM_QUANTI = {
    "titre": "Enquête individuelle",
    "uid": "ahVcrcrzvdoNYwkFSNjUpj",
    "xlsform": "Évaluation des besoins des travailleurs informels du secteur des déchets - Dakar.xlsx",
}
FORM_QUALI = {
    "titre": "Entretiens & focus groups",
    "uid": "aotuV7h5fdZe8LDS3KF6KP",
    "xlsform": "Grille de saisie qualitative (entretiens et focus groups).xlsx",
}

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
