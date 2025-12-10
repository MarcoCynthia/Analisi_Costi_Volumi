import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
from pathlib import Path

# -------------------------------------------------------
# CONFIGURAZIONE STREAMLIT
# -------------------------------------------------------
st.set_page_config(page_title="Analisi Costi & Volumi – Fornitore", layout="wide")
st.title("📊 Analisi Fatturazione – Fornitore")
st.write("Il file Excel viene caricato automaticamente dalla cartella dedicata sul Desktop.")

# -------------------------------------------------------
# 1) Caricamento automatico dataset
# -------------------------------------------------------
cartella = Path(r"C:\Users\A762431\OneDrive - OPENFIBER SPA\Desktop\Analisi_Costi_Volumi")

if not cartella.exists():
    st.error(f"❌ La cartella non esiste: {cartella}")
    st.stop()

file_list = list(cartella.glob("*.xlsx"))

if len(file_list) == 0:
    st.error(f"❌ Nessun file Excel trovato in {cartella}")
    st.stop()

file_scelto = file_list[0]
df = pd.read_excel(file_scelto, engine="openpyxl")

# -------------------------------------------------------
# 2) Preparazione Dati
# -------------------------------------------------------
cols_vol = ["vol_AB", "vol_CD", "vol_AGF"]
cols_cost = ["cost_AB", "cost_CD", "cost_AGF"]

df[cols_vol + cols_cost] = df[cols_vol + cols_cost].fillna(0)

df["vol_totale"] = df[cols_vol].sum(axis=1)
df["costo_totale"] = df[cols_cost].sum(axis=1)

def safe_div(num, den):
    return np.where(den > 0, num / den, np.nan)

df["€/unità_AB"] = safe_div(df["cost_AB"], df["vol_AB"])
df["€/unità_CD"] = safe_div(df["cost_CD"], df["vol_CD"])
df["€/unità_AGF"] = safe_div(df["cost_AGF"], df["vol_AGF"])

# DF long
df_long = pd.DataFrame()
for clus in ["AB", "CD", "AGF"]:
    vol_col = f"vol_{clus}"
    cost_col = f"cost_{clus}"
    df_sub = df.copy()
    df_sub["cluster"] = clus
    df_sub["volume"] = df_sub[vol_col]
    df_sub["costo"] = df_sub[cost_col]
    df_sub["costo_unitario"] = safe_div(df_sub[cost_col], df_sub[vol_col])
    df_long = pd.concat([df_long, df_sub], ignore_index=True)

if 'mese' in df_long.columns:
    df_long['mese_num'] = df_long['mese'].astype(str).str[-2:].astype(int)
else:
    df_long['mese_num'] = 0

# -------------------------------------------------------
# 3) FILTRI GLOBALI (sidebar)
# -------------------------------------------------------
st.sidebar.header("Filtri")

cluster_unici = ["Tutti"] + sorted(df_long["cluster"].dropna().unique().tolist())
cluster_sel = st.sidebar.selectbox("Seleziona Cluster", cluster_unici)

linea_unici = ["Tutti"]
if "linea" in df_long.columns:
    valori_linea = df_long["linea"].dropna().unique()
    opzioni_linea = ["Delivery", "Assurance"]
    presenti = [x for x in opzioni_linea if x in valori_linea]
    if len(presenti) == 0:
        presenti = sorted(valori_linea.tolist())
    linea_unici += presenti
linea_sel = st.sidebar.selectbox("Seleziona Linea", linea_unici)

mesi_unici = sorted(df_long["mese_num"].dropna().unique())
mesi_sel = st.sidebar.multiselect(
    "Seleziona Mese",
    options=[m for m in mesi_unici if m != 12],
    default=[m for m in mesi_unici if m != 12]
)

df_filtered = df_long.copy()
if cluster_sel != "Tutti":
    df_filtered = df_filtered[df_filtered["cluster"] == cluster_sel]
if linea_sel != "Tutti":
    df_filtered = df_filtered[df_filtered["linea"] == linea_sel]
if mesi_sel:
    df_filtered = df_filtered[df_filtered["mese_num"].isin(mesi_sel)]

# -------------------------------------------------------
# 4) Vista Macro: confronto fornitori
# -------------------------------------------------------
st.subheader("📊 Confronto Fornitori – Visione Macro")

fornitori = sorted(df_filtered["fornitore"].dropna().unique())
if len(fornitori) < 2:
    st.warning("❗️Sono necessari almeno due fornitori per il confronto.")
else:
    df_forn_costo = df_filtered.groupby("fornitore")["costo"].sum().reset_index()
    df_forn_vol = df_filtered.groupby("fornitore")["volume"].sum().reset_index()

    col1, col2 = st.columns(2)
    with col1:
        fig1 = px.pie(df_forn_costo, values="costo", names="fornitore",
                      title="Peso Costo Totale per Fornitore")
        fig1.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        fig2 = px.pie(df_forn_vol, values="volume", names="fornitore",
                      title="Peso Volume Totale per Fornitore")
        fig2.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig2, use_container_width=True)

# -------------------------------------------------------
# 5) Analisi Dettagliata per Fornitore
# -------------------------------------------------------
st.subheader("🔎 Analisi Dettagliata per Fornitore")
lista_fornitori = sorted(df_filtered["fornitore"].dropna().unique())
fornitore_sel = st.selectbox("Scegli il fornitore:", lista_fornitori)

df_forn = df_filtered[df_filtered["fornitore"] == fornitore_sel]

costo_tot = df_forn["costo"].sum()
vol_tot = df_forn["volume"].sum()
cu_ponderato = costo_tot / vol_tot if vol_tot > 0 else np.nan
incidenza = 100 * costo_tot / df_filtered["costo"].sum() if df_filtered["costo"].sum() > 0 else np.nan

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Costo Totale", f"{costo_tot:,.0f} €")
col2.metric("📦 Volume Totale", f"{vol_tot:,.0f}")
col3.metric("⚖️ Costo Unitario Medio", f"{cu_ponderato:.3f} €")
col4.metric("📊 Incidenza sul Totale", f"{incidenza:.2f} %")

if "categoria" in df_forn.columns:
    st.markdown("### 📂 Distribuzione Costi per Categoria")
    df_cat = df_forn.groupby("categoria")["costo"].sum().sort_values(ascending=False)
    st.bar_chart(df_cat)

# -------------------------------------------------------
# 6) Trend Mensile Costo/Volume (Gen–Nov)
# -------------------------------------------------------
st.subheader("📈 Trend Mensile Costo e Volume per Fornitore (Gen-Nov)")

df_trend = df_long[df_long["mese_num"].between(1, 11)]
df_trend_agg = df_trend.groupby(["fornitore", "mese_num"])[["costo", "volume"]].sum().reset_index()

fig_costo = px.line(df_trend_agg, x="mese_num", y="costo", color="fornitore", markers=True)
st.plotly_chart(fig_costo, use_container_width=True)

fig_volume = px.line(df_trend_agg, x="mese_num", y="volume", color="fornitore", markers=True)
st.plotly_chart(fig_volume, use_container_width=True)

# -------------------------------------------------------
# 7) Trend Mensile per Categoria – Filtri Fornitore + Cluster + Linea (Gen–Ott)
# -------------------------------------------------------
st.subheader("📈 Trend Mensile per Categoria – Filtra Fornitore, Cluster e Linea (Gen-Ott)")

# Liste valori unici
fornitori_all = sorted(df_long["fornitore"].dropna().unique())
cluster_all = ["Tutti"] + sorted(df_long["cluster"].dropna().unique())
linee_all = ["Tutte"] + sorted(df_long["linea"].dropna().unique())

# Selettori
col_f1, col_f2, col_f3 = st.columns(3)
fornitore_trend_sel = col_f1.selectbox("Seleziona Fornitore", fornitori_all)
cluster_trend_sel = col_f2.selectbox("Seleziona Cluster", cluster_all)
linea_trend_sel = col_f3.selectbox("Seleziona Linea", linee_all)

# Filtri
df_tc = df_long[
    (df_long["fornitore"] == fornitore_trend_sel) &
    (df_long["mese_num"].between(1, 10))
]

if cluster_trend_sel != "Tutti":
    df_tc = df_tc[df_tc["cluster"] == cluster_trend_sel]

if linea_trend_sel != "Tutte":
    df_tc = df_tc[df_tc["linea"] == linea_trend_sel]

# Aggregazione
df_tc_agg = df_tc.groupby(["categoria", "mese_num"]).agg(
    costo=("costo", "sum"),
    volume=("volume", "sum")
).reset_index()

df_tc_agg["costo_unitario"] = df_tc_agg["costo"] / df_tc_agg["volume"].replace(0, np.nan)

# ---------------------------
# Grafici: Costo + Volume affiancati
# ---------------------------
col_g1, col_g2 = st.columns(2)

fig_c = px.line(
    df_tc_agg, x="mese_num", y="costo", color="categoria",
    markers=True, title="Costo"
)
fig_c.update_layout(legend=dict(orientation="h", y=-0.4))
col_g1.plotly_chart(fig_c, use_container_width=True)

fig_v = px.line(
    df_tc_agg, x="mese_num", y="volume", color="categoria",
    markers=True, title="Volume"
)
fig_v.update_layout(legend=dict(orientation="h", y=-0.4))
col_g2.plotly_chart(fig_v, use_container_width=True)

# ---------------------------
# Grafico sotto: Costo Unitario (full width)
# ---------------------------
fig_cu = px.line(
    df_tc_agg, x="mese_num", y="costo_unitario", color="categoria",
    markers=True, title="Costo Unitario"
)
fig_cu.update_layout(legend=dict(orientation="h", y=-0.3))

st.plotly_chart(fig_cu, use_container_width=True)



# -------------------------------------------------------
# 8) Trend Invertito: Categoria → Confronto Fornitori (Gen–Ott)
# -------------------------------------------------------
st.subheader("📊 Confronto Trend Categoria fra Fornitori (Gen-Ott)")

# FILTRI FORNITORE, CLUSTER E LINEA
fornitori_all = sorted(df_long["fornitore"].dropna().unique())
fornitore_sel_trend = st.selectbox("Seleziona Fornitore (Trend Categoria)", ["Tutti"] + fornitori_all)

cluster_all = sorted(df_long["cluster"].dropna().unique())
cluster_sel_trend = st.selectbox("Seleziona Cluster (Trend Categoria)", ["Tutti"] + cluster_all)

linea_all = ["Tutti"]
if "linea" in df_long.columns:
    valori_linea = df_long["linea"].dropna().unique()
    opzioni_linea = ["Delivery", "Assurance"]
    presenti = [x for x in opzioni_linea if x in valori_linea]
    if len(presenti) == 0:
        presenti = sorted(valori_linea.tolist())
    linea_all += presenti
linea_sel_trend = st.selectbox("Seleziona Linea (Trend Categoria)", linea_all)

# APPLICO I FILTRI
df_trend_filtrato = df_long[df_long["mese_num"].between(1, 10)]
if fornitore_sel_trend != "Tutti":
    df_trend_filtrato = df_trend_filtrato[df_trend_filtrato["fornitore"] == fornitore_sel_trend]
if cluster_sel_trend != "Tutti":
    df_trend_filtrato = df_trend_filtrato[df_trend_filtrato["cluster"] == cluster_sel_trend]
if linea_sel_trend != "Tutti":
    df_trend_filtrato = df_trend_filtrato[df_trend_filtrato["linea"] == linea_sel_trend]

# Scegli categoria
categorie_all = sorted(df_trend_filtrato["categoria"].dropna().unique())
categoria_sel = st.selectbox("Seleziona Categoria (confronto fornitori)", categorie_all)

df_comp = df_trend_filtrato[df_trend_filtrato["categoria"] == categoria_sel]

df_comp_agg = df_comp.groupby(["fornitore", "mese_num"]).agg(
    costo=("costo", "sum"),
    volume=("volume", "sum")
).reset_index()

df_comp_agg["costo_unitario"] = df_comp_agg["costo"] / df_comp_agg["volume"].replace(0, np.nan)

# Layout grafici: costo e volume affiancati, costo unitario sotto
col_cf1, col_cf2 = st.columns(2)
col_cf3 = st.container()  # per costo unitario sotto

fig_cc = px.line(df_comp_agg, x="mese_num", y="costo", color="fornitore", markers=True, title="Costo")
fig_cc.update_layout(legend=dict(orientation="h", y=-0.3))
col_cf1.plotly_chart(fig_cc, use_container_width=True)

fig_cv = px.line(df_comp_agg, x="mese_num", y="volume", color="fornitore", markers=True, title="Volume")
fig_cv.update_layout(legend=dict(orientation="h", y=-0.3))
col_cf2.plotly_chart(fig_cv, use_container_width=True)

fig_ccu = px.line(df_comp_agg, x="mese_num", y="costo_unitario", color="fornitore", markers=True, title="Costo Unitario")
fig_ccu.update_layout(legend=dict(orientation="h", y=-0.3))
col_cf3.plotly_chart(fig_ccu, use_container_width=True)

# -------------------------------------------------------
# 8b) Vista Stacked – Peso Categorie Mese per Mese (Filtrata)
# -------------------------------------------------------
st.subheader("📊 Composizione Mensile delle Categorie (Stacked Bar – Gen-Ott)")

# Applico gli stessi filtri già definiti
df_stack = df_long[df_long["mese_num"].between(1, 10)]
if fornitore_sel_trend != "Tutti":
    df_stack = df_stack[df_stack["fornitore"] == fornitore_sel_trend]
if cluster_sel_trend != "Tutti":
    df_stack = df_stack[df_stack["cluster"] == cluster_sel_trend]
if linea_sel_trend != "Tutti":
    df_stack = df_stack[df_stack["linea"] == linea_sel_trend]

# Calcolo costo totale per mese
df_stack_mese = df_stack.groupby("mese_num")["costo"].sum().reset_index()
df_stack = df_stack.merge(df_stack_mese, on="mese_num", suffixes=("", "_tot"))

# Calcolo peso percentuale di ciascuna categoria
df_stack["peso_categoria"] = df_stack["costo"] / df_stack["costo_tot"]

# Grafico Stacked Bar
fig_stack = px.bar(
    df_stack,
    x="mese_num",
    y="peso_categoria",
    color="categoria",
    title="Peso % delle Categorie per Mese (Filtrato)",
    labels={"peso_categoria": "Peso %", "mese_num": "Mese"}
)
fig_stack.update_layout(
    barmode="stack",
    legend=dict(orientation="h", y=-0.3),
)

st.plotly_chart(fig_stack, use_container_width=True)


# -------------------------------------------------------
# 9) Matrix di Contributo alla Variazione Mese/mese (Potenziata)
# -------------------------------------------------------
st.subheader("🔍 Matrix Contributo Categorie → Variazione Mese/Mese (Gen-Ott)")

# --------------------------
# FILTRI: Fornitore, Cluster, Linea
# --------------------------
fornitori_all = sorted(df_long["fornitore"].dropna().unique())
fornitore_sel = st.selectbox("Seleziona Fornitore (Matrix Contributo)", ["Tutti"] + fornitori_all)

cluster_all = sorted(df_long["cluster"].dropna().unique())
cluster_sel = st.selectbox("Seleziona Cluster (Matrix Contributo)", ["Tutti"] + cluster_all)

linea_all = ["Tutti"]
if "linea" in df_long.columns:
    valori_linea = df_long["linea"].dropna().unique()
    opzioni_linea = ["Delivery", "Assurance"]
    presenti = [x for x in opzioni_linea if x in valori_linea]
    if len(presenti) == 0:
        presenti = sorted(valori_linea.tolist())
    linea_all += presenti
linea_sel = st.selectbox("Seleziona Linea (Matrix Contributo)", linea_all)

# --------------------------
# APPLICO FILTRI
# --------------------------
df_trend_filtrato = df_long[df_long["mese_num"].between(1, 10)]
if fornitore_sel != "Tutti":
    df_trend_filtrato = df_trend_filtrato[df_trend_filtrato["fornitore"] == fornitore_sel]
if cluster_sel != "Tutti":
    df_trend_filtrato = df_trend_filtrato[df_trend_filtrato["cluster"] == cluster_sel]
if linea_sel != "Tutti":
    df_trend_filtrato = df_trend_filtrato[df_trend_filtrato["linea"] == linea_sel]

# --------------------------
# Tooltip esplicativo
# --------------------------
st.info("""
ℹ️ **Come leggere la matrice:**
- Ogni colonna rappresenta la variazione dei costi da un mese al successivo (es. 7→8).
- I valori indicano **quanto ogni categoria ha contribuito** alla variazione totale.
- **Rosso** = categoria che ha fatto **aumentare** maggiormente il costo.
- **Verde** = categoria che ha **ridotto** il costo compensando gli aumenti.
- Il totale della colonna è sempre **100%** (o -100% se il costo totale è sceso).
""")

# --------------------------
# CALCOLO MATRICE SICURA
# --------------------------
df_var = df_trend_filtrato.pivot_table(
    index="categoria",
    columns="mese_num",
    values="costo",
    aggfunc="sum",   # somma i valori duplicati
    fill_value=0
)

mesi = sorted(df_var.columns)
contrib_matrix = pd.DataFrame()
alert_over_50 = []  # per alert automatici

for i in range(1, len(mesi)):
    m_prev = mesi[i-1]
    m_curr = mesi[i]
    
    diff_tot = df_var[m_curr].sum() - df_var[m_prev].sum()
    diff_cat = df_var[m_curr] - df_var[m_prev]

    contrib_pct = diff_cat / diff_tot if diff_tot != 0 else diff_cat * 0
    contrib_matrix[f"{m_prev}→{m_curr}"] = contrib_pct

    # ALERT AUTOMATICO > 50%
    for categoria, pct in contrib_pct.items():
        if abs(pct) > 0.50:
            alert_over_50.append(
                f"🔴 **{categoria}** contribuisce per **{pct:.1%}** alla variazione **{m_prev}→{m_curr}**."
            )

# Mostra matrice
st.dataframe(contrib_matrix.style.background_gradient(cmap="RdYlGn_r").format("{:.1%}"))

# --------------------------
# SPIEGAZIONE AUTOMATICA
# --------------------------
st.markdown("### 🧠 Spiegazione Automatica della Variazione")

explanations = []
for col in contrib_matrix.columns:
    top_cat = contrib_matrix[col].idxmax()
    top_value = contrib_matrix[col].max()
    bottom_cat = contrib_matrix[col].idxmin()
    bottom_value = contrib_matrix[col].min()

    explanations.append(
        f"➡️ Nel passaggio **{col}**, la categoria **{top_cat}** ha guidato l’andamento con un contributo del **{top_value:.1%}**."
    )
    if bottom_value < 0:
        explanations.append(
            f"   La categoria **{bottom_cat}** ha compensato la variazione con **{bottom_value:.1%}**."
        )

for e in explanations:
    st.write(e)


