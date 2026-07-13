# -*- coding: utf-8 -*-
"""Utilitaires : connexion API KoboToolbox + lecture du XLSForm enquete_C40.xlsx."""
import re
import unicodedata
import random
from datetime import datetime, timedelta

import pandas as pd
import requests

# ---------------------------------------------------------------- XLSForm ---

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


def load_xlsform(path: str) -> dict:
    """Lit le XLSForm et retourne types, labels, groupes et choix."""
    survey = pd.read_excel(path, sheet_name="survey").fillna("")
    choices = pd.read_excel(path, sheet_name="choices").fillna("")

    qtype, label, listname, group_of, groups = {}, {}, {}, {}, {}
    stack = []
    for _, r in survey.iterrows():
        t = str(r["type"]).strip()
        name = str(r["name"]).strip()
        lab = str(r.get("label", "")).strip()
        if not t:
            continue
        if t == "begin_group":
            stack.append(name)
            groups[name] = lab or name
            continue
        if t == "end_group":
            if stack:
                stack.pop()
            continue
        if not name:
            continue
        qtype[name] = t
        label[name] = lab or name
        group_of[name] = stack[-1] if stack else ""
        m = re.match(r"select_(one|multiple)\s+(\S+)", t)
        if m:
            listname[name] = m.group(2)

    choice_map = {}
    for _, r in choices.iterrows():
        ln, code, lab = str(r["list_name"]).strip(), str(r["name"]).strip(), str(r["label"]).strip()
        if ln:
            choice_map.setdefault(ln, {})[code] = lab.strip(" ,")

    return {"qtype": qtype, "label": label, "listname": listname,
            "group_of": group_of, "groups": groups, "choices": choice_map}


# ------------------------------------------------------------- API Kobo -----

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
    """Télécharge toutes les soumissions (paginé)."""
    url = f"{base_url.rstrip('/')}/api/v2/assets/{asset_uid}/data.json?limit=1000"
    records = []
    while url:
        r = requests.get(url, headers=kobo_headers(token), timeout=120)
        r.raise_for_status()
        js = r.json()
        records.extend(js.get("results", []))
        url = js.get("next")
    return records


# --------------------------------------------------------- Mise en forme ----

def to_dataframe(records: list, form: dict) -> pd.DataFrame:
    """Aplati les soumissions Kobo : retire les préfixes de groupes, décode les
    choix (select_one -> label), extrait GPS et dates."""
    if not records:
        return pd.DataFrame()
    rows = []
    for rec in records:
        row = {}
        for k, v in rec.items():
            short = k.split("/")[-1]
            row[short] = v
        rows.append(row)
    df = pd.DataFrame(rows)

    # Dates
    if "_submission_time" in df:
        df["_submission_time"] = pd.to_datetime(df["_submission_time"], errors="coerce")
        df["date_soumission"] = df["_submission_time"].dt.date

    # GPS : champ geopoint "lat lon alt acc", sinon _geolocation
    gps_col = "Localisation_GPS_du_point_d_enqu_te"
    lat, lon = [], []
    for _, r in df.iterrows():
        la = lo = None
        v = r.get(gps_col)
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
    """Compte les réponses d'un select_multiple (codes séparés par espaces)."""
    cmap = form["choices"].get(form["listname"].get(col, ""), {})
    vals = []
    for v in df[col].dropna():
        if isinstance(v, list):  # certains exports renvoient des listes
            v = " ".join(str(x) for x in v)
        if isinstance(v, str):
            for code in v.split():
                lab = str(cmap.get(code, code)).strip()
                if lab:
                    vals.append(lab)
    return pd.Series(vals).value_counts()


# ------------------------------------------------------- Données de démo ----

def generate_demo_data(form: dict, n: int = 180, seed: int = 42) -> list:
    """Simule des soumissions Kobo réalistes pour tester l'app sans connexion."""
    rng = random.Random(seed)
    ch = form["choices"]
    ln = form["listname"]
    # centres approximatifs par commune (lat, lon)
    centres = {
        "dakar_plateau": (14.667, -17.433), "m_dina": (14.678, -17.451),
        "gor_e__m_dina__gueule_tap_e_fass_coloban": (14.667, -17.398),
        "gueule_tap_e_fass_colobane": (14.685, -17.448),
        "fann_point_e_amiti": (14.693, -17.465), "biscuiterie": (14.703, -17.452),
        "grand_dakar": (14.708, -17.447), "hlm": (14.713, -17.443),
        "sicap_libert": (14.712, -17.459), "dieuppeul_derkl": (14.717, -17.455),
        "hann_bel_air": (14.717, -17.425), "golf_sud__gu_diawaye": (14.712, -17.472),
        "ouakam": (14.722, -17.49), "ngor": (14.748, -17.512), "yoff": (14.75, -17.47),
        "grand_yoff": (14.735, -17.452), "patte_d_oie": (14.735, -17.44),
        "parcelles_assainies": (14.755, -17.435), "camb_r_ne": (14.765, -17.43),
    }
    enqueteurs = ["A. Ndiaye", "F. Sarr", "M. Fall", "K. Diop", "S. Ba"]
    recs = []
    t0 = datetime.now() - timedelta(days=21)
    for i in range(n):
        commune = rng.choice(list(centres))
        la, lo = centres[commune]
        la += rng.uniform(-0.008, 0.008)
        lo += rng.uniform(-0.008, 0.008)
        ts = t0 + timedelta(days=rng.uniform(0, 21))
        rec = {"_id": i + 1, "_submission_time": ts.isoformat(),
               "_geolocation": [la, lo],
               "group_nx6ww37/Nom_de_l_enqu_teur": rng.choice(enqueteurs),
               "group_nx6ww37/Commune_d_enqu_te": commune,
               "group_nx6ww37/Localisation_GPS_du_point_d_enqu_te": f"{la} {lo} 0 5",
               "group_nx6ww37/Consentement_clair_obtenu": rng.choices(["oui", "non"], [95, 5])[0],
               "group_ha3cx90/Revenu_moyen_estim_FCFA": int(rng.gauss(85000, 35000)),
               "group_sd37u39/Volume_approximatif_coul_en_kg": max(1, int(rng.gauss(120, 60))),
               }
        # remplit les select_one / select_multiple avec des codes aléatoires
        for name, t in form["qtype"].items():
            key = name if not form["group_of"].get(name) else f"{form['group_of'][name]}/{name}"
            if key in rec or name in ("Commune_d_enqu_te", "Consentement_clair_obtenu"):
                continue
            codes = list(ch.get(ln.get(name, ""), {}))
            if not codes:
                continue
            if t.startswith("select_one"):
                rec[key] = rng.choice(codes)
            elif t.startswith("select_multiple"):
                rec[key] = " ".join(rng.sample(codes, k=min(len(codes), rng.randint(1, 3))))
        recs.append(rec)
    return recs


# ------------------------------------------- Plan d'échantillonnage C40 -----
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
# Quotas transversaux : part minimale attendue
PLAN_TRANSVERSAL = {"femmes": 0.40, "jeunes": 0.25, "migrants": 0.15}
