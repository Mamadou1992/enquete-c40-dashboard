# -*- coding: utf-8 -*-
"""
Tableau de bord Enquête C40 GREEN+ - Acteurs du secteur des déchets, Dakar
Connexion KoboToolbox (kf.kobotoolbox.org) + visualisation Streamlit.

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
XLSFORM = os.path.join(APP_DIR, "enquete_C40.xlsx")
GEOJSON = os.path.join(APP_DIR, "communes_dakar.geojson")
DEFAULT_URL = "https://kf.kobotoolbox.org"

st.set_page_config(page_title="Enquête C40 GREEN+ Dakar", page_icon="♻️", layout="wide")

# Habillage vert des composants (puces, encadrés, onglets, indicateurs)
st.markdown("""
<style>
/* Puces des filtres (multiselect) : vert plein, texte blanc */
[data-baseweb="tag"] { background-color: #1b5e20 !important; }
[data-baseweb="tag"] span { color: #ffffff !important; }
[data-baseweb="tag"] svg { fill: #ffffff !important; }

/* Option survolée / sélectionnée dans les listes déroulantes */
[data-baseweb="popover"] li:hover,
[data-baseweb="popover"] li[aria-selected="true"] { background-color: #f1f8e9 !important; }

/* Encadrés info / avertissement : fond vert pâle, liseré vert */
div[data-testid="stAlert"] {
  background-color: #f1f8e9 !important;
  border-left: 5px solid #9ccc65 !important;
  border-radius: 8px;
}
div[data-testid="stAlert"] p { color: #12291a !important; }

/* Onglet actif souligné en vert */
button[data-baseweb="tab"][aria-selected="true"] {
  color: #1b5e20 !important;
  font-weight: 600;
}
div[data-baseweb="tab-highlight"] { background-color: #1b5e20 !important; }

/* Indicateurs (metrics) : cartes vert pâle */
div[data-testid="stMetric"] {
  background-color: #f1f8e9;
  border-left: 5px solid #9ccc65;
  border-radius: 10px;
  padding: 12px 16px;
}

/* En-têtes des volets dépliants (expanders) */
details summary {
  background-color: #f1f8e9 !important;
  border-radius: 8px;
}

/* Boutons de téléchargement */
div[data-testid="stDownloadButton"] button {
  background-color: #1b5e20 !important;
  color: #ffffff !important;
  border: none;
}
div[data-testid="stDownloadButton"] button:hover { background-color: #0d3f14 !important; }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------- Chargements ----

@st.cache_resource
def get_form():
    return ku.load_xlsform(XLSFORM)

@st.cache_resource
def get_geojson():
    with open(GEOJSON, encoding="utf-8") as f:
        return json.load(f)

@st.cache_data(ttl=300, show_spinner="Téléchargement des données Kobo…")
def get_data(base_url, token, uid):
    return ku.fetch_submissions(base_url, token, uid)

@st.cache_data(ttl=300)
def get_assets(base_url, token):
    return ku.list_assets(base_url, token)

form = get_form()
geojson = get_geojson()

# --------------------------------------------------------------- Sidebar ----

st.sidebar.title("♻️ Enquête C40 GREEN+")
st.sidebar.caption("Acteurs du secteur des déchets - Département de Dakar")

base_url = st.secrets.get("kobo", {}).get("base_url", DEFAULT_URL)
token = st.secrets.get("kobo", {}).get("token", "")
uid = st.secrets.get("kobo", {}).get("asset_uid", "")

with st.sidebar.expander("⚙️ Connexion Kobo", expanded=not token):
    base_url = st.text_input("Serveur", base_url)
    token = st.text_input("Jeton API", token, type="password",
                          help="Compte Kobo → Paramètres → Sécurité → Clé API")
    if token and not uid:
        try:
            assets = get_assets(base_url, token)
            if assets:
                opts = {f"{a['name']} ({a['count']} soumissions)": a["uid"] for a in assets}
                uid = opts[st.selectbox("Formulaire", list(opts))]
        except Exception as e:
            st.error(f"Connexion impossible : {e}")
    elif uid:
        st.caption(f"Formulaire : `{uid}`")

if st.sidebar.button("🔄 Actualiser les données"):
    get_data.clear()
    st.rerun()

# ------------------------------------------------------------- Données ------

if not (token and uid):
    st.warning("Renseignez le jeton API et choisissez un formulaire dans la barre latérale.")
    st.stop()
try:
    records = get_data(base_url, token, uid)
except Exception as e:
    st.error(f"Erreur API Kobo : {e}")
    st.stop()

df = ku.to_dataframe(records, form)
if df.empty:
    st.info("Aucune soumission pour le moment.")
    st.stop()

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
        opts = sorted(fdf[col].dropna().unique())
        sel = st.sidebar.multiselect(titre, opts)
        if sel:
            fdf = fdf[fdf[col].isin(sel)]

sidebar_filter("Commune_d_enqu_te", "Commune d'enquête")
sidebar_filter("Nom_de_l_enqu_teur", "Enquêteur")
sidebar_filter("Sexe", "Sexe")

st.sidebar.metric("Enquêtes affichées", f"{len(fdf)} / {len(df)}")

# ----------------------------------------------------------------- Corps ----

tab1, tab2, tab3, tab4 = st.tabs(
    ["📊 Suivi des enquêtes", "🗺️ Carte", "📈 Analyse thématique", "📥 Données & export"])

COL_COMMUNE = "Commune_d_enqu_te"

# ============ 1. SUIVI ============
with tab1:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Enquêtes réalisées", len(fdf))
    nb_com = fdf[COL_COMMUNE].nunique() if COL_COMMUNE in fdf else 0
    c2.metric("Communes couvertes", f"{nb_com} / 19")
    nb_enq = fdf["Nom_de_l_enqu_teur"].nunique() if "Nom_de_l_enqu_teur" in fdf else 0
    c3.metric("Enquêteurs actifs", nb_enq)
    if "Consentement_clair_obtenu" in fdf and len(fdf):
        cons = fdf["Consentement_clair_obtenu"].astype(str).str.lower().str.startswith("oui").mean()
        c4.metric("Taux de consentement", f"{cons:.0%}")

    # Progression vs plan d'échantillonnage (400 enquêtes)
    prog = len(fdf) / ku.PLAN_TOTAL
    st.progress(min(prog, 1.0), text=f"Progression du plan d'échantillonnage : "
                f"{len(fdf)} / {ku.PLAN_TOTAL} enquêtes ({prog:.0%})")

    # Quotas transversaux (plan C40 : >= 40 % femmes, >= 25 % jeunes, >= 15 % migrants)
    q1, q2, q3 = st.columns(3)
    if "Sexe" in fdf and len(fdf):
        p = (fdf["Sexe"] == "Femme").mean()
        q1.metric("Femmes (cible >= 40 %)", f"{p:.0%}",
                  delta=f"{(p - 0.40) * 100:+.0f} pts", delta_color="normal")
    if "_ge" in fdf and len(fdf):
        p = fdf["_ge"].isin(["<18", "18-24", "25-34"]).mean()
        q2.metric("Jeunes < 35 ans (cible >= 25 %)", f"{p:.0%}",
                  delta=f"{(p - 0.25) * 100:+.0f} pts", delta_color="normal")
    if "Statut_migratoire" in fdf and len(fdf):
        p = fdf["Statut_migratoire"].astype(str).str.contains("migrant", case=False).mean()
        q3.metric("Migrants (cible >= 15 %)", f"{p:.0%}",
                  delta=f"{(p - 0.15) * 100:+.0f} pts", delta_color="normal")

    g1, g2 = st.columns(2)
    with g1:
        if "date_soumission" in fdf:
            par_jour = fdf.groupby("date_soumission").size().reset_index(name="Enquêtes")
            fig = px.area(par_jour, x="date_soumission", y="Enquêtes",
                          title="Évolution des soumissions par jour",
                          labels={"date_soumission": "Date"}, markers=True)
            st.plotly_chart(fig, width="stretch")
    with g2:
        if COL_COMMUNE in fdf:
            par_com = fdf[COL_COMMUNE].value_counts().reset_index()
            par_com.columns = ["Commune", "Enquêtes"]
            fig = px.bar(par_com.sort_values("Enquêtes"), x="Enquêtes", y="Commune",
                         orientation="h", title="Enquêtes par commune", text_auto=True)
            st.plotly_chart(fig, width="stretch")

    if "Nom_de_l_enqu_teur" in fdf:
        par_enq = fdf["Nom_de_l_enqu_teur"].value_counts().reset_index()
        par_enq.columns = ["Enquêteur", "Enquêtes"]
        fig = px.bar(par_enq, x="Enquêteur", y="Enquêtes", title="Enquêtes par enquêteur",
                     text_auto=True)
        st.plotly_chart(fig, width="stretch")

# ============ 2. CARTE ============
with tab2:
    left, right = st.columns([3, 1])
    with right:
        fond = st.radio("Affichage", ["Points GPS", "Densité par commune", "Les deux"], index=2)
        indicateur = st.radio("Couleur des communes",
                              ["Enquêtes réalisées", "Taux de réalisation (%)",
                               "Quota du plan", "Population ANSD 2023"])

    counts = pd.DataFrame({"key": [], "Enquêtes réalisées": []})
    if COL_COMMUNE in fdf:
        cc = fdf[COL_COMMUNE].dropna().map(ku.commune_to_key).value_counts()
        counts = cc.reset_index()
        counts.columns = ["key", "Enquêtes réalisées"]
    # toutes les communes, même à 0, enrichies du plan d'échantillonnage
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
                                        hover_name="Commune", color_continuous_scale=["#eee", "#eee"],
                                        center={"lat": 14.716, "lon": -17.45}, zoom=10.6,
                                        opacity=0.25, height=650)
                fig.update_coloraxes(showscale=False)
            gps = fdf.dropna(subset=["latitude", "longitude"])
            if fond in ("Points GPS", "Les deux") and len(gps):
                hover = gps.get(COL_COMMUNE, pd.Series([""] * len(gps), index=gps.index))
                fig.add_scattermap(lat=gps["latitude"], lon=gps["longitude"],
                                   mode="markers", marker={"size": 9, "color": "#d62728"},
                                   text=hover, name="Enquêtes", hovertemplate="%{text}<extra></extra>")
            fig.update_layout(margin={"l": 0, "r": 0, "t": 0, "b": 0}, map_style="carto-positron")
            st.plotly_chart(fig, width="stretch")
        except Exception as e:
            st.error(f"Erreur d'affichage de la carte : {e}")

    st.caption("Suivi du plan d'échantillonnage (400 enquêtes réparties au poids "
               "démographique ANSD 2023)")
    st.dataframe(
        counts[["Commune", "Population ANSD 2023", "Quota du plan", "Enquêtes réalisées",
                "Taux de réalisation (%)", "Restant"]]
        .sort_values("Quota du plan", ascending=False),
        width="stretch", hide_index=True)

# ============ 3. ANALYSE ============
with tab3:
    groupes = {v: k for k, v in form["groups"].items()}
    gsel = st.selectbox("Section du questionnaire", list(groupes))
    gname = groupes[gsel]
    # Champs exclus de l'analyse (identifiants personnels)
    EXCLUS = {"Num_ro_de_t_l_phone", "Nom_de_l_enqu_teur", "Nom_de_l_enqu_t"}
    questions = [q for q, g in form["group_of"].items()
                 if g == gname and q in fdf.columns and q not in EXCLUS
                 and form["qtype"][q].split()[0] in ("select_one", "select_multiple",
                                                     "integer", "decimal", "text")]
    if not questions:
        st.info("Aucune question analysable dans cette section.")
    for q in questions:
        t = form["qtype"][q]
        titre = form["label"].get(q, q)
        reponses = fdf[q].dropna()
        reponses = reponses[reponses.astype(str).str.strip() != ""]
        if reponses.empty:
            continue
        if t == "text":  # réponse libre : liste déroulante des réponses saisies
            with st.expander(f"📝 {titre} - {len(reponses)} réponse(s) saisie(s)"):
                d = reponses.astype(str).reset_index(drop=True)
                d.index += 1
                st.dataframe(d.rename("Réponse"), width="stretch")
            continue
        if t.startswith("select_multiple"):
            # Choix multiples : treemap (pavés proportionnels)
            s = ku.explode_multiple(fdf, q, form)
            if s.empty or s.sum() == 0:
                st.caption(f"◽ {titre} : aucune réponse exploitable pour le moment")
                continue
            d = s.reset_index()
            d.columns = ["Réponse", "Nombre"]
            fig = px.treemap(d, path=["Réponse"], values="Nombre",
                             title=f"{titre} (choix multiples, n={len(reponses)})")
            fig.update_traces(textinfo="label+value+percent root")
            fig.update_layout(height=420, margin={"t": 60, "l": 10, "r": 10, "b": 10})
        elif t.startswith("select_one"):
            s = reponses.astype(str).str.strip()
            s = s[s != ""].value_counts()
            if s.empty:
                st.caption(f"◽ {titre} : aucune réponse exploitable pour le moment")
                continue
            d = s.reset_index()
            d.columns = ["Réponse", "Nombre"]
            if len(d) <= 5:
                # Peu de modalités : anneau (donut)
                fig = px.pie(d, names="Réponse", values="Nombre", hole=0.45,
                             title=f"{titre} (n={int(d['Nombre'].sum())})")
                fig.update_traces(textinfo="percent+value")
                fig.update_layout(height=380)
            else:
                # Longue liste : barres triées avec pourcentages
                d = d.sort_values("Nombre")
                d["pct"] = (d["Nombre"] / d["Nombre"].sum() * 100).map(lambda v: f"{v:.0f} %")
                fig = px.bar(d, x="Nombre", y="Réponse", orientation="h",
                             title=f"{titre} (n={int(d['Nombre'].sum())})", text="pct")
                fig.update_traces(textposition="outside", cliponaxis=False)
                fig.update_layout(height=max(320, 28 * min(len(d), 20)), yaxis_title="")
        else:
            # Numérique : histogramme + boîte à moustaches
            serie = pd.to_numeric(fdf[q], errors="coerce").dropna()
            if serie.empty:
                st.caption(f"◽ {titre} : aucune réponse exploitable pour le moment")
                continue
            st.markdown(f"**{titre}** (n={len(serie)}) - moyenne : {serie.mean():,.0f} · médiane : "
                        f"{serie.median():,.0f} · min : {serie.min():,.0f} · max : {serie.max():,.0f}")
            fig = px.histogram(serie, nbins=30, title=f"{titre} (n={len(serie)})",
                               labels={"value": titre}, marginal="box")
            fig.update_layout(showlegend=False, height=400)
        st.plotly_chart(fig, width="stretch")

# ============ 4. DONNÉES ============
with tab4:
    # colonnes renommées avec les libellés du formulaire
    ren, vus = {}, {}
    for c in fdf.columns:
        lab = str(form["label"].get(c, c)).strip()
        n = vus.get(lab, 0)
        vus[lab] = n + 1
        ren[c] = lab if n == 0 else f"{lab} ({n + 1})"
    show = fdf.rename(columns=ren)
    st.dataframe(show, width="stretch", height=500)

    c1, c2 = st.columns(2)
    csv = show.to_csv(index=False).encode("utf-8-sig")
    c1.download_button("⬇️ Télécharger CSV", csv, "enquete_C40_donnees.csv", "text/csv")
    buf = io.BytesIO()
    exp = show.copy()
    for col in exp.columns:  # Excel n'accepte pas les tz ni listes
        if exp[col].map(lambda x: isinstance(x, (list, dict))).any():
            exp[col] = exp[col].astype(str)
        if isinstance(exp[col].dtype, pd.DatetimeTZDtype):
            exp[col] = exp[col].dt.tz_localize(None)
    exp.to_excel(buf, index=False)
    c2.download_button("⬇️ Télécharger Excel", buf.getvalue(), "enquete_C40_donnees.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
