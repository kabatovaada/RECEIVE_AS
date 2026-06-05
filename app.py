import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os, shutil

st.set_page_config(page_title="RECEIVE AS Report", page_icon="📥", layout="wide")

# ─── Custom CSS ──────────────────────────────────────
st.markdown("""
<style>
    .block-container {padding-top: 1.5rem; padding-bottom: 1rem;}
    div[data-testid="stMetric"] {background: #161B22; border: 1px solid #30363D; border-radius: 8px; padding: 12px 16px;}
    div[data-testid="stMetric"] label {color: #8B949E !important; font-size: 0.8rem !important;}
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {color: #C9D1D9 !important;}
    .green-card {background: #1A7F37; border: 1px solid #238636; border-radius: 10px; padding: 24px; margin-bottom: 16px;}
    .green-card h1 {color: white; margin: 0; font-size: 3rem;}
    .green-card p {color: #C5E8D0; margin: 4px 0 0 0;}
    .info-bar {background: #262730; border: 1px solid #30363D; border-radius: 6px; padding: 8px 14px; margin-bottom: 12px; color: #8B949E; font-size: 0.85rem;}
    .section-title {color: #484F58; letter-spacing: 2px; font-size: 0.75rem; font-weight: 700; border-bottom: 1px solid #30363D; padding-bottom: 6px; margin: 20px 0 10px 0;}
    .card-metric {background: #161B22; border: 1px solid #30363D; border-radius: 8px; padding: 16px; text-align: center;}
    .card-metric .value {font-size: 1.8rem; font-weight: 700; color: #C9D1D9;}
    .card-metric .label {font-size: 0.75rem; color: #8B949E;}
    .card-metric .sub {font-size: 0.7rem; color: #484F58;}
</style>
""", unsafe_allow_html=True)

GEOSIZE_COLORS = {"SPO": "#2EA043", "BPO": "#6366F1"}
ABC_COLORS = {"A": "#2EA043", "B": "#D29922", "C": "#FF4B4B"}
PLOTLY_TEMPLATE = "plotly_dark"

# ─── Trvalé úložisko súboru ───────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
DATA_FILE = os.path.join(DATA_DIR, "RECEIVE_AS_REPORT.xlsx")
os.makedirs(DATA_DIR, exist_ok=True)

# ─── Load data ───────────────────────────────────────
@st.cache_data
def load_data(path):
    df = pd.read_excel(path, engine="openpyxl")
    df["Čas vzniku"] = pd.to_datetime(df["Čas vzniku"], errors="coerce")
    df["den"] = df["Čas vzniku"].dt.date
    df["hodina"] = df["Čas vzniku"].dt.hour
    df["den_tyzdna_sk"] = df["Čas vzniku"].dt.dayofweek.map({
        0: "Pondelok", 1: "Utorok", 2: "Streda", 3: "Štvrtok",
        4: "Piatok", 5: "Sobota", 6: "Nedeľa"
    })
    df["den_num"] = df["Čas vzniku"].dt.dayofweek
    return df

# ─── Sidebar ──────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📥 RECEIVE AS")
    st.markdown("---")

    # Zobrazíme stav nahratého súboru
    if os.path.exists(DATA_FILE):
        mtime = os.path.getmtime(DATA_FILE)
        mtime_str = pd.Timestamp(mtime, unit="s").strftime("%d.%m.%Y %H:%M")
        st.success(f"✅ Súbor nahraný\n\n`{mtime_str}`")
    else:
        st.warning("⚠️ Žiadny súbor na serveri")

    uploaded = st.file_uploader("Nahrať / aktualizovať XLSX", type=["xlsx"])

    if uploaded is not None:
        with open(DATA_FILE, "wb") as f:
            f.write(uploaded.read())
        st.cache_data.clear()
        st.success("✅ Súbor uložený na server!")
        st.rerun()

    st.markdown("---")
    st.markdown("**Filtre**")

# ─── Načítanie dát ────────────────────────────────────
if not os.path.exists(DATA_FILE):
    st.info("👈 Nahrajte súbor **RECEIVE_AS_REPORT.xlsx** v ľavom paneli. Súbor zostane uložený na serveri.")
    st.stop()

df_raw = load_data(DATA_FILE)

# ─── Filtre ───────────────────────────────────────────
with st.sidebar:
    geosize_opts = sorted(df_raw["GeoSize"].dropna().unique())
    sel_geosize = st.multiselect("GeoSize", geosize_opts, default=geosize_opts)

    abc_opts = sorted(df_raw["ABC rank"].dropna().unique())
    sel_abc = st.multiselect("ABC rank", abc_opts, default=abc_opts)

    ops = sorted(df_raw["Naskladnil"].dropna().unique())
    sel_ops = st.multiselect("Operátor (Naskladnil)", ops, default=ops)

df = df_raw.copy()
if sel_geosize:
    df = df[df["GeoSize"].isin(sel_geosize)]
if sel_abc:
    df = df[df["ABC rank"].isin(sel_abc)]
if sel_ops:
    df = df[df["Naskladnil"].isin(sel_ops)]

# ═══════════════════════════════════════════════════════
# SEKCIA 1 – HLAVNÝ PREHĽAD
# ═══════════════════════════════════════════════════════
st.markdown("## 📥 RECEIVE AS — Nepotvrdené príjmy do AutoStore")
st.markdown(f"""
<div class="info-bar">
Store: 162 SKLC3 · Nepotvrdené záznamy · {df['den'].min()} → {df['den'].max()}
</div>
""", unsafe_allow_html=True)

col_hero, col_kpi = st.columns([1, 1])

with col_hero:
    total_lines = len(df)
    total_kusy = int(df["Ks na lokačním pohybu"].sum())
    total_sku = df["Produkt"].nunique()
    st.markdown(f"""
    <div class="green-card">
        <p style="color:#C5E8D0;font-size:0.85rem;letter-spacing:2px;margin-bottom:4px;">CELKOVÉ ZÁZNAMY · RECEIVE AS REPORT</p>
        <h1>{total_lines:,}</h1>
        <p>{total_kusy:,} kusov celkovo</p>
        <div style="display:inline-block;background:#0E1117;border:1px solid #238636;border-radius:4px;padding:2px 12px;margin-top:6px;">
            <span style="color:#3FB950;font-weight:700;font-size:0.9rem;">{total_sku} SKU</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_kpi:
    k1, k2 = st.columns(2)
    with k1:
        st.metric("Unikátnych dokladov", f"{df['Skladový doklad'].nunique():,}")
        st.metric("Unikátnych operátorov", f"{df['Naskladnil'].nunique():,}")
    with k2:
        st.metric("Priemerné kusy / záznam", f"{df['Ks na lokačním pohybu'].mean():.1f}")
        st.metric("Max kusy / záznam", f"{int(df['Ks na lokačním pohybu'].max()):,}")

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 2 – GEOSIZE & ABC RANK
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">ROZDELENIE PODĽA GEOSIZE A ABC RANK</p>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

with col1:
    geo_counts = df["GeoSize"].value_counts().reset_index()
    geo_counts.columns = ["GeoSize", "Záznamy"]
    fig = px.pie(geo_counts, names="GeoSize", values="Záznamy",
                 color="GeoSize", color_discrete_map=GEOSIZE_COLORS,
                 template=PLOTLY_TEMPLATE, hole=0.5,
                 title="Záznamy podľa GeoSize")
    fig.update_traces(textinfo="percent+label")
    fig.update_layout(showlegend=False, height=300, margin=dict(t=40, b=10, l=10, r=10))
    st.plotly_chart(fig, use_container_width=True)

with col2:
    abc_counts = df["ABC rank"].value_counts().reset_index()
    abc_counts.columns = ["ABC rank", "Záznamy"]
    fig2 = px.bar(abc_counts, x="ABC rank", y="Záznamy",
                  color="ABC rank", color_discrete_map=ABC_COLORS,
                  template=PLOTLY_TEMPLATE, title="Záznamy podľa ABC rank")
    fig2.update_layout(showlegend=False, height=300, margin=dict(t=40, b=10, l=10, r=10))
    st.plotly_chart(fig2, use_container_width=True)

with col3:
    geo_abc = df.groupby(["GeoSize", "ABC rank"]).size().reset_index(name="Záznamy")
    fig3 = px.bar(geo_abc, x="GeoSize", y="Záznamy", color="ABC rank",
                  color_discrete_map=ABC_COLORS, barmode="stack",
                  template=PLOTLY_TEMPLATE, title="GeoSize × ABC rank")
    fig3.update_layout(height=300, margin=dict(t=40, b=10, l=10, r=10))
    st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 3 – ČASOVÝ TREND
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">PRÍJEM V ČASE — DENNÝ TREND</p>', unsafe_allow_html=True)

daily = df.groupby("den").agg(
    Záznamy=("Produkt", "count"),
    Kusy=("Ks na lokačním pohybu", "sum")
).reset_index()
daily["den"] = daily["den"].astype(str)

col_d1, col_d2 = st.columns(2)
with col_d1:
    fig_d1 = px.bar(daily, x="den", y="Záznamy", template=PLOTLY_TEMPLATE,
                    title="Záznamy príjmu podľa dňa", color_discrete_sequence=["#2EA043"])
    fig_d1.update_layout(height=300, margin=dict(t=40, b=10, l=10, r=10))
    st.plotly_chart(fig_d1, use_container_width=True)
with col_d2:
    fig_d2 = px.bar(daily, x="den", y="Kusy", template=PLOTLY_TEMPLATE,
                    title="Kusy príjmu podľa dňa", color_discrete_sequence=["#6366F1"])
    fig_d2.update_layout(height=300, margin=dict(t=40, b=10, l=10, r=10))
    st.plotly_chart(fig_d2, use_container_width=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 4 – HODINOVÁ DISTRIBÚCIA
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">PRÍJEM PODĽA HODINY</p>', unsafe_allow_html=True)

hourly = df.groupby("hodina").agg(
    Záznamy=("Produkt", "count"),
    Kusy=("Ks na lokačním pohybu", "sum")
).reset_index().sort_values("hodina")

col_h1, col_h2 = st.columns(2)
with col_h1:
    fig_h1 = px.bar(hourly, x="hodina", y="Záznamy", template=PLOTLY_TEMPLATE,
                    title="Záznamy príjmu podľa hodiny", color_discrete_sequence=["#D29922"])
    fig_h1.update_layout(height=300, margin=dict(t=40, b=10, l=10, r=10))
    fig_h1.update_xaxes(dtick=1)
    st.plotly_chart(fig_h1, use_container_width=True)
with col_h2:
    fig_h2 = px.bar(hourly, x="hodina", y="Kusy", template=PLOTLY_TEMPLATE,
                    title="Kusy príjmu podľa hodiny", color_discrete_sequence=["#FF4B4B"])
    fig_h2.update_layout(height=300, margin=dict(t=40, b=10, l=10, r=10))
    fig_h2.update_xaxes(dtick=1)
    st.plotly_chart(fig_h2, use_container_width=True)

st.markdown('<p class="section-title">PEAKOVÝ DEŇ TÝŽDŇA</p>', unsafe_allow_html=True)
day_order = ["Pondelok", "Utorok", "Streda", "Štvrtok", "Piatok", "Sobota", "Nedeľa"]
weekday = df.groupby(["den_num", "den_tyzdna_sk"]).agg(
    Záznamy=("Produkt", "count"),
    Kusy=("Ks na lokačním pohybu", "sum")
).reset_index().sort_values("den_num")
fig_wd = px.bar(weekday, x="den_tyzdna_sk", y="Záznamy",
                template=PLOTLY_TEMPLATE, title="Záznamy podľa dňa v týždni",
                color_discrete_sequence=["#3FB950"],
                category_orders={"den_tyzdna_sk": day_order})
fig_wd.update_layout(height=300, margin=dict(t=40, b=10, l=10, r=10))
st.plotly_chart(fig_wd, use_container_width=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 5 – OPERÁTORI
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">OPERÁTORI — NASKLADNIL</p>', unsafe_allow_html=True)

ops_df = df.groupby("Naskladnil").agg(
    Záznamy=("Produkt", "count"),
    Kusy=("Ks na lokačním pohybu", "sum"),
    Produkty=("Produkt", "nunique")
).reset_index().sort_values("Záznamy", ascending=False).head(15)

col_o1, col_o2 = st.columns([2, 1])
with col_o1:
    fig_ops = px.bar(ops_df, x="Záznamy", y="Naskladnil", orientation="h",
                     template=PLOTLY_TEMPLATE, title="Top operátori podľa počtu záznamov",
                     color="Záznamy", color_continuous_scale="Greens")
    fig_ops.update_layout(height=400, margin=dict(t=40, b=10, l=10, r=10), yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig_ops, use_container_width=True)
with col_o2:
    fig_ops2 = px.bar(ops_df.head(10), x="Kusy", y="Naskladnil", orientation="h",
                      template=PLOTLY_TEMPLATE, title="Top 10 operátori — kusy",
                      color_discrete_sequence=["#6366F1"])
    fig_ops2.update_layout(height=400, margin=dict(t=40, b=10, l=10, r=10), yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig_ops2, use_container_width=True)

with st.expander("📋 Detail všetkých operátorov"):
    ops_full = df.groupby("Naskladnil").agg(
        Záznamy=("Produkt", "count"),
        Kusy=("Ks na lokačním pohybu", "sum"),
        Unikátnych_SKU=("Produkt", "nunique"),
        Unikátnych_dokladov=("Skladový doklad", "nunique")
    ).reset_index().sort_values("Záznamy", ascending=False)
    st.dataframe(ops_full, use_container_width=True, hide_index=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 6 – TOP PRODUKTY
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">TOP PRODUKTY</p>', unsafe_allow_html=True)

prod_df = df.groupby(["Produkt", "Popis", "ABC rank", "GeoSize"]).agg(
    Záznamy=("Produkt", "count"),
    Kusy=("Ks na lokačním pohybu", "sum")
).reset_index().sort_values("Kusy", ascending=False).head(20)

col_p1, col_p2 = st.columns([2, 1])
with col_p1:
    fig_prod = px.bar(prod_df.head(15), x="Kusy", y="Produkt", orientation="h",
                      template=PLOTLY_TEMPLATE, title="Top 15 produktov podľa kusov",
                      color="ABC rank", color_discrete_map=ABC_COLORS)
    fig_prod.update_layout(height=450, margin=dict(t=40, b=10, l=10, r=10), yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig_prod, use_container_width=True)
with col_p2:
    st.markdown("**Top 10 podľa kusov**")
    for _, row in prod_df.head(10).iterrows():
        abc_color = ABC_COLORS.get(row["ABC rank"], "#8B949E")
        geo_color = GEOSIZE_COLORS.get(row["GeoSize"], "#8B949E")
        popis_short = str(row["Popis"])[:50] + "…" if len(str(row["Popis"])) > 50 else str(row["Popis"])
        st.markdown(f"""
        <div style="background:#161B22;border:1px solid #30363D;border-radius:6px;padding:8px 12px;margin-bottom:6px;">
            <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="color:#C9D1D9;font-weight:700;font-size:0.85rem;">{row['Produkt']}</span>
                <span style="color:{abc_color};font-weight:700;">{int(row['Kusy'])} ks</span>
            </div>
            <div style="color:#8B949E;font-size:0.72rem;margin-top:2px;">{popis_short}</div>
            <div style="margin-top:4px;">
                <span style="background:{abc_color};color:white;border-radius:3px;padding:1px 7px;font-size:0.7rem;font-weight:700;">ABC: {row['ABC rank']}</span>
                <span style="background:{geo_color};color:white;border-radius:3px;padding:1px 7px;font-size:0.7rem;font-weight:700;margin-left:4px;">{row['GeoSize']}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 7 – DOKLADY
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">TOP DOKLADY</p>', unsafe_allow_html=True)

doc_df = df.groupby("Skladový doklad").agg(
    Záznamy=("Produkt", "count"),
    Kusy=("Ks na lokačním pohybu", "sum"),
    SKU=("Produkt", "nunique")
).reset_index().sort_values("Kusy", ascending=False).head(15)

fig_doc = px.bar(doc_df, x="Skladový doklad", y="Kusy",
                 template=PLOTLY_TEMPLATE, title="Top 15 dokladov podľa kusov",
                 color="Záznamy", color_continuous_scale="Blues",
                 hover_data=["SKU"])
fig_doc.update_layout(height=350, margin=dict(t=40, b=10, l=10, r=10))
st.plotly_chart(fig_doc, use_container_width=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 8 – SUROVÉ DÁTA
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">SUROVÉ DÁTA</p>', unsafe_allow_html=True)

with st.expander("📋 Zobraziť/exportovať surové dáta"):
    show_cols = ["Produkt", "Popis", "GeoSize", "ABC rank", "Ks na lokačním pohybu",
                 "Skladový doklad", "Naskladnil", "Čas vzniku", "Bedna", "Lokace"]
    st.dataframe(df[[c for c in show_cols if c in df.columns]], use_container_width=True, hide_index=True)
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Stiahnuť CSV", csv, "receive_as_report.csv", "text/csv")
