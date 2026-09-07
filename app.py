# -*- coding: utf-8 -*-
"""
Plateforme de suivi - Enquête C40 GREEN+ · Travailleurs informels des déchets, Dakar
Connexion KoboToolbox (enquête individuelle + grille qualitative) et visualisation.

Lancement :  streamlit run app.py
"""
import io
import json
import os

import pandas as pd
import plotly.express as px
import streamlit as st

import kobo_utils as ku

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
def get_form(xlsform):
    return ku.load_xlsform(os.path.join(APP_DIR, xlsform))


@st.cache_resource
def get_geojson():
    with open(GEOJSON, encoding="utf-8") as f:
        return json.load(f)


@st.cache_data(ttl=300, show_spinner="Téléchargement des données Kobo…")
def get_data(base_url, token, uid):
    return ku.fetch_submissions(base_url, token, uid)


form = get_form(ku.FORM_QUANTI["xlsform"])
form_q = get_form(ku.FORM_QUALI["xlsform"])
geojson = get_geojson()

# --------------------------------------------------------------- Sidebar ----

if os.path.exists(LOGO):
    st.sidebar.image(LOGO, width="stretch")
st.sidebar.title("Enquête C40 GREEN+")
st.sidebar.caption("Travailleurs informels des déchets - Département de Dakar")

base_url = st.secrets.get("kobo", {}).get("base_url", DEFAULT_URL)
token = st.secrets.get("kobo", {}).get("token", "")
uid = st.secrets.get("kobo", {}).get("asset_uid", ku.FORM_QUANTI["uid"])
uid_q = st.secrets.get("kobo", {}).get("asset_uid_quali", ku.FORM_QUALI["uid"])

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

df = ku.to_dataframe(records, form)
dfq = ku.to_dataframe(records_q, form_q)

if df.empty:
    st.info("Aucune soumission de l'enquête individuelle pour le moment.")
    st.stop()

C = ku.COL

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
    if C["consent"] in fdf and len(fdf):
        cons = fdf[C["consent"]].astype(str).str.lower().str.startswith("oui").mean()
        c4.metric("Taux de consentement", f"{cons:.0%}")

    prog = len(fdf) / ku.PLAN_TOTAL
    st.progress(min(prog, 1.0), text=f"Progression du plan d'échantillonnage : "
                f"{len(fdf)} / {ku.PLAN_TOTAL} enquêtes ({prog:.0%})")

    st.markdown("**Quotas transversaux du plan C40**")
    q1, q2, q3 = st.columns(3)
    if C["sexe"] in fdf and len(fdf):
        p = (fdf[C["sexe"]].astype(str) == "Femme").mean()
        q1.metric("Femmes (cible >= 40 %)", f"{p:.0%}", delta=f"{(p - 0.40) * 100:+.0f} pts")
    if C["age"] in fdf and fdf[C["age"]].notna().any():
        p = (fdf[C["age"]] < ku.AGE_JEUNE_MAX).mean()
        q2.metric(f"Jeunes < {ku.AGE_JEUNE_MAX} ans (cible >= 25 %)", f"{p:.0%}",
                  delta=f"{(p - 0.25) * 100:+.0f} pts")
    if C["migratoire"] in fdf and len(fdf):
        p = fdf[C["migratoire"]].astype(str).str.contains("migrant", case=False).mean()
        q3.metric("Migrants (cible >= 15 %)", f"{p:.0%}", delta=f"{(p - 0.15) * 100:+.0f} pts")

    g1, g2 = st.columns(2)
    with g1:
        if "date_soumission" in fdf:
            par_jour = fdf.groupby("date_soumission").size().reset_index(name="Enquêtes")
            fig = px.area(par_jour, x="date_soumission", y="Enquêtes", markers=True,
                          title="Évolution des soumissions par jour",
                          labels={"date_soumission": "Date"})
            st.plotly_chart(fig, width="stretch")
    with g2:
        # Quotas par fonction (plan : 60 % charretiers, 20 % récupérateurs, 15 % trieurs)
        if C["fonction"] in fdf and fdf[C["fonction"]].notna().any():
            obs = fdf[C["fonction"]].value_counts(normalize=True) * 100
            lignes = []
            for fonction, part in ku.PLAN_FONCTIONS.items():
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
        if C["commune"] in fdf:
            par_com = fdf[C["commune"]].value_counts().reset_index()
            par_com.columns = ["Commune", "Enquêtes"]
            fig = px.bar(par_com.sort_values("Enquêtes"), x="Enquêtes", y="Commune",
                         orientation="h", title="Enquêtes par commune", text_auto=True)
            fig.update_layout(height=520, yaxis_title="")
            st.plotly_chart(fig, width="stretch")
    with g4:
        if C["enqueteur"] in fdf:
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
        indicateur = st.radio("Couleur des communes",
                              ["Enquêtes réalisées", "Taux de réalisation (%)",
                               "Quota du plan", "Population ANSD 2023"])

    counts = pd.DataFrame({"key": [], "Enquêtes réalisées": []})
    if C["commune"] in fdf:
        cc = fdf[C["commune"]].dropna().map(ku.commune_to_key).value_counts()
        counts = cc.reset_index()
        counts.columns = ["key", "Enquêtes réalisées"]
    all_keys = pd.DataFrame({"key": [f["properties"]["key"] for f in geojson["features"]],
                             "Commune": [f["properties"]["commune"] for f in geojson["features"]]})
    counts = all_keys.merge(counts, on="key", how="left").fillna({"Enquêtes réalisées": 0})
    counts["Population ANSD 2023"] = counts["key"].map(lambda k: ku.PLAN_QUOTAS.get(k, (0, 0))[0])
    counts["Quota du plan"] = counts["key"].map(lambda k: ku.PLAN_QUOTAS.get(k, (0, 0))[1])
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
            s = ku.explode_multiple(data, q, formulaire)
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
            if g == gname and q in data.columns and q not in ku.CHAMPS_EXCLUS
            and formulaire["qtype"][q].split()[0] in
            ("select_one", "select_multiple", "integer", "decimal", "text")]


with tab3:
    libelles = {form["groups"][g]: g for g in form["ordre_groupes"]}
    gsel = st.selectbox("Module du questionnaire", list(libelles))
    afficher_questions(fdf, form, questions_de(form, fdf, libelles[gsel]), "quanti")

# ========================================================== 4. QUALITATIF ====
with tab4:
    if dfq.empty:
        st.info("Aucun entretien ni focus group saisi pour le moment. "
                f"Formulaire : {ku.FORM_QUALI['titre']}.")
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
