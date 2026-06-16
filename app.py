import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os, base64, io, requests

st.set_page_config(page_title="RECEIVE AS Report", page_icon="📥", layout="wide")

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
</style>
""", unsafe_allow_html=True)

GEOSIZE_COLORS = {"SPO": "#2EA043", "BPO": "#6366F1"}
ABC_COLORS     = {"A": "#2EA043", "B": "#D29922", "C": "#FF4B4B"}
PT             = "plotly_dark"

# ═══════════════════════════════════════════════════════
# GITHUB STORAGE
# ═══════════════════════════════════════════════════════
GH_TOKEN  = st.secrets.get("GH_TOKEN", "")
GH_REPO   = st.secrets.get("GH_REPO", "")
GH_PATH   = st.secrets.get("GH_PATH", "data/RECEIVE_AS_REPORT.xlsx")
GH_BRANCH = st.secrets.get("GH_BRANCH", "main")
GH_API    = f"https://api.github.com/repos/{GH_REPO}/contents/{GH_PATH}"
GH_HDR    = {"Authorization": f"token {GH_TOKEN}", "Accept": "application/vnd.github.v3+json"}

def gh_get():
    try:
        r = requests.get(GH_API, headers=GH_HDR, params={"ref": GH_BRANCH}, timeout=15)
        if r.status_code == 200:
            d = r.json()
            return base64.b64decode(d["content"]), d["sha"]
    except Exception:
        pass
    return None, None

def gh_put(data, sha=None):
    try:
        payload = {
            "message": "update RECEIVE_AS_REPORT.xlsx via Streamlit",
            "content": base64.b64encode(data).decode(),
            "branch": GH_BRANCH,
        }
        if sha:
            payload["sha"] = sha
        r = requests.put(GH_API, headers=GH_HDR, json=payload, timeout=30)
        return r.status_code in (200, 201)
    except Exception:
        return False

def gh_last_update():
    try:
        r = requests.get(f"https://api.github.com/repos/{GH_REPO}/commits",
                         headers=GH_HDR, params={"path": GH_PATH, "per_page": 1}, timeout=10)
        if r.status_code == 200 and r.json():
            return pd.Timestamp(r.json()[0]["commit"]["committer"]["date"]).strftime("%d.%m.%Y %H:%M")
    except Exception:
        pass
    return "neznámy"

# ═══════════════════════════════════════════════════════
# LOAD DATA
# ═══════════════════════════════════════════════════════
@st.cache_data(ttl=300)
def load_data():
    content, sha = gh_get()
    if content is None:
        return None, None
    df = pd.read_excel(io.BytesIO(content), engine="openpyxl")
    df["Čas vzniku"] = pd.to_datetime(df["Čas vzniku"], errors="coerce")
    df = df.dropna(subset=["Čas vzniku"]).copy()
    df["den"]        = pd.to_datetime(df["Čas vzniku"].dt.date)
    df["hodina"]     = df["Čas vzniku"].dt.hour
    df["den_num"]    = df["Čas vzniku"].dt.dayofweek
    df["den_sk"]     = df["den_num"].map({
        0:"Pondelok",1:"Utorok",2:"Streda",3:"Štvrtok",4:"Piatok",5:"Sobota",6:"Nedeľa"
    })
    return df, sha

# ═══════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📥 RECEIVE AS")
    st.markdown("---")

    if not GH_TOKEN or not GH_REPO:
        st.error("⚠️ Chýbajú Secrets:\n`GH_TOKEN` a `GH_REPO`")
        st.markdown("""
**Nastav v Streamlit Cloud → Settings → Secrets:**
```
GH_TOKEN  = "ghp_xxxx"
GH_REPO   = "user/repo"
GH_PATH   = "data/RECEIVE_AS_REPORT.xlsx"
GH_BRANCH = "main"
```
        """)
    else:
        st.success(f"✅ GitHub: `{GH_REPO}`\n\nAktualizácia: `{gh_last_update()}`")

    uploaded = st.file_uploader("Nahrať / aktualizovať XLSX", type=["xlsx"])
    if uploaded:
        with st.spinner("Ukladám na GitHub..."):
            _, cur_sha = gh_get()
            ok = gh_put(uploaded.read(), sha=cur_sha)
        if ok:
            st.cache_data.clear()
            st.success("✅ Uložené na GitHub!")
            st.rerun()
        else:
            st.error("❌ Chyba — skontroluj token a oprávnenia.")

    st.markdown("---")
    st.markdown("**Filtre**")

# ═══════════════════════════════════════════════════════
# STOP AK CHÝBA CONFIG ALEBO DÁTA
# ═══════════════════════════════════════════════════════
if not GH_TOKEN or not GH_REPO:
    st.warning("⚙️ Najprv nastav GitHub Secrets v Streamlit Cloud.")
    st.stop()

with st.spinner("Načítavam dáta z GitHub..."):
    df_raw, file_sha = load_data()

if df_raw is None:
    st.info("👈 Nahrajte súbor **RECEIVE_AS_REPORT.xlsx** v ľavom paneli.")
    st.stop()

# ═══════════════════════════════════════════════════════
# FILTRE
# ═══════════════════════════════════════════════════════
with st.sidebar:
    sel_geo = st.multiselect("GeoSize",
        sorted(df_raw["GeoSize"].dropna().unique()),
        default=sorted(df_raw["GeoSize"].dropna().unique()))
    sel_abc = st.multiselect("ABC rank",
        sorted(df_raw["ABC rank"].dropna().unique()),
        default=sorted(df_raw["ABC rank"].dropna().unique()))
    sel_ops = st.multiselect("Operátor",
        sorted(df_raw["Naskladnil"].dropna().unique()),
        default=sorted(df_raw["Naskladnil"].dropna().unique()))

df = df_raw.copy()
if sel_geo: df = df[df["GeoSize"].isin(sel_geo)]
if sel_abc: df = df[df["ABC rank"].isin(sel_abc)]
if sel_ops: df = df[df["Naskladnil"].isin(sel_ops)]

# ═══════════════════════════════════════════════════════
# HLAVIČKA
# ═══════════════════════════════════════════════════════
if df_raw.empty:
    st.warning("⚠️ V súbore nie sú žiadne riadky s platným dátumom v stĺpci 'Čas vzniku'.")
    st.stop()
den_sorted = df_raw["den"].dropna().sort_values()
if den_sorted.empty:
    st.warning("⚠️ V súbore nie sú žiadne riadky s platným dátumom v stĺpci 'Čas vzniku'.")
    st.stop()

prvy_den    = den_sorted.iloc[0]
posledny_den = den_sorted.iloc[-1]
datum_od  = prvy_den.strftime("%d.%m.%Y")
datum_do  = posledny_den.strftime("%d.%m.%Y")
pocet_dni = (posledny_den - prvy_den).days + 1

st.markdown("## 📥 RECEIVE AS — Nepotvrdené príjmy do AutoStore")
st.markdown(f"""
<div class="info-bar">
📅 Obdobie reportu: <strong style="color:#C9D1D9;">{datum_od} — {datum_do}</strong> &nbsp;·&nbsp; {pocet_dni} dní
&nbsp;&nbsp;|&nbsp;&nbsp; Store: 162 SKLC3 &nbsp;·&nbsp; Nepotvrdené záznamy &nbsp;·&nbsp; GitHub: {GH_REPO}/{GH_PATH}
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
# SEKCIA 1 – KPI
# ═══════════════════════════════════════════════════════
col_hero, col_kpi = st.columns(2)

with col_hero:
    st.markdown(f"""
    <div class="green-card">
        <p style="color:#C5E8D0;font-size:0.85rem;letter-spacing:2px;margin-bottom:4px;">CELKOVÉ ZÁZNAMY · RECEIVE AS REPORT</p>
        <h1>{len(df):,}</h1>
        <p>{int(df["Ks na lokačním pohybu"].sum()):,} kusov celkovo</p>
        <div style="display:inline-block;background:#0E1117;border:1px solid #238636;border-radius:4px;padding:2px 12px;margin-top:6px;">
            <span style="color:#3FB950;font-weight:700;font-size:0.9rem;">{df["Produkt"].nunique()} SKU</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_kpi:
    k1, k2 = st.columns(2)
    with k1:
        st.metric("Unikátnych dokladov",  f"{df['Skladový doklad'].nunique():,}")
        st.metric("Unikátnych operátorov",f"{df['Naskladnil'].nunique():,}")
    with k2:
        st.metric("Priemerné kusy / záznam", f"{df['Ks na lokačním pohybu'].mean():.1f}")
        st.metric("Max kusy / záznam",        f"{int(df['Ks na lokačním pohybu'].max()):,}")

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 2 – GEOSIZE & ABC
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">ROZDELENIE PODĽA GEOSIZE A ABC RANK</p>', unsafe_allow_html=True)
c1, c2, c3 = st.columns(3)

with c1:
    g = df["GeoSize"].value_counts().reset_index()
    g.columns = ["GeoSize","Záznamy"]
    fig = px.pie(g, names="GeoSize", values="Záznamy", color="GeoSize",
                 color_discrete_map=GEOSIZE_COLORS, template=PT, hole=0.5,
                 title="Záznamy podľa GeoSize")
    fig.update_traces(textinfo="percent+label")
    fig.update_layout(showlegend=False, height=300, margin=dict(t=40,b=10,l=10,r=10))
    st.plotly_chart(fig, use_container_width=True)

with c2:
    a = df["ABC rank"].value_counts().reset_index()
    a.columns = ["ABC rank","Záznamy"]
    fig2 = px.bar(a, x="ABC rank", y="Záznamy", color="ABC rank",
                  color_discrete_map=ABC_COLORS, template=PT, title="Záznamy podľa ABC rank")
    fig2.update_layout(showlegend=False, height=300, margin=dict(t=40,b=10,l=10,r=10))
    st.plotly_chart(fig2, use_container_width=True)

with c3:
    ga = df.groupby(["GeoSize","ABC rank"]).size().reset_index(name="Záznamy")
    fig3 = px.bar(ga, x="GeoSize", y="Záznamy", color="ABC rank",
                  color_discrete_map=ABC_COLORS, barmode="stack",
                  template=PT, title="GeoSize × ABC rank")
    fig3.update_layout(height=300, margin=dict(t=40,b=10,l=10,r=10))
    st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 3 – DENNÝ TREND
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">PRÍJEM V ČASE — DENNÝ TREND</p>', unsafe_allow_html=True)
daily = df.groupby("den").agg(Záznamy=("Produkt","count"), Kusy=("Ks na lokačním pohybu","sum")).reset_index()
daily["den"] = daily["den"].astype(str)

d1, d2 = st.columns(2)
with d1:
    fig = px.bar(daily, x="den", y="Záznamy", template=PT,
                 title="Záznamy podľa dňa", color_discrete_sequence=["#2EA043"])
    fig.update_layout(height=300, margin=dict(t=40,b=10,l=10,r=10))
    st.plotly_chart(fig, use_container_width=True)
with d2:
    fig = px.bar(daily, x="den", y="Kusy", template=PT,
                 title="Kusy podľa dňa", color_discrete_sequence=["#6366F1"])
    fig.update_layout(height=300, margin=dict(t=40,b=10,l=10,r=10))
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 4 – HODINOVÁ DISTRIBÚCIA
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">PRÍJEM PODĽA HODINY</p>', unsafe_allow_html=True)
hourly = df.groupby("hodina").agg(Záznamy=("Produkt","count"), Kusy=("Ks na lokačním pohybu","sum")).reset_index()

h1, h2 = st.columns(2)
with h1:
    fig = px.bar(hourly, x="hodina", y="Záznamy", template=PT,
                 title="Záznamy podľa hodiny", color_discrete_sequence=["#D29922"])
    fig.update_layout(height=300, margin=dict(t=40,b=10,l=10,r=10))
    fig.update_xaxes(dtick=1)
    st.plotly_chart(fig, use_container_width=True)
with h2:
    fig = px.bar(hourly, x="hodina", y="Kusy", template=PT,
                 title="Kusy podľa hodiny", color_discrete_sequence=["#FF4B4B"])
    fig.update_layout(height=300, margin=dict(t=40,b=10,l=10,r=10))
    fig.update_xaxes(dtick=1)
    st.plotly_chart(fig, use_container_width=True)

st.markdown('<p class="section-title">PEAKOVÝ DEŇ TÝŽDŇA</p>', unsafe_allow_html=True)
day_order = ["Pondelok","Utorok","Streda","Štvrtok","Piatok","Sobota","Nedeľa"]
wd = df.groupby(["den_num","den_sk"]).agg(Záznamy=("Produkt","count")).reset_index().sort_values("den_num")
fig = px.bar(wd, x="den_sk", y="Záznamy", template=PT,
             title="Záznamy podľa dňa v týždni",
             color_discrete_sequence=["#3FB950"],
             category_orders={"den_sk": day_order})
fig.update_layout(height=300, margin=dict(t=40,b=10,l=10,r=10))
st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 5 – OPERÁTORI
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">OPERÁTORI — NASKLADNIL</p>', unsafe_allow_html=True)
ops_df = df.groupby("Naskladnil").agg(
    Záznamy=("Produkt","count"),
    Kusy=("Ks na lokačním pohybu","sum"),
    SKU=("Produkt","nunique")
).reset_index().sort_values("Záznamy", ascending=False).head(15)

o1, o2 = st.columns([2,1])
with o1:
    fig = px.bar(ops_df, x="Záznamy", y="Naskladnil", orientation="h",
                 template=PT, title="Top operátori podľa záznamov",
                 color="Záznamy", color_continuous_scale="Greens")
    fig.update_layout(height=400, margin=dict(t=40,b=10,l=10,r=10), yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig, use_container_width=True)
with o2:
    fig = px.bar(ops_df.head(10), x="Kusy", y="Naskladnil", orientation="h",
                 template=PT, title="Top 10 — kusy",
                 color_discrete_sequence=["#6366F1"])
    fig.update_layout(height=400, margin=dict(t=40,b=10,l=10,r=10), yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig, use_container_width=True)

with st.expander("📋 Detail všetkých operátorov"):
    ops_full = df.groupby("Naskladnil").agg(
        Záznamy=("Produkt","count"),
        Kusy=("Ks na lokačním pohybu","sum"),
        Unikátnych_SKU=("Produkt","nunique"),
        Unikátnych_dokladov=("Skladový doklad","nunique")
    ).reset_index().sort_values("Záznamy", ascending=False)
    st.dataframe(ops_full, use_container_width=True, hide_index=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 6 – TOP PRODUKTY
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">TOP PRODUKTY</p>', unsafe_allow_html=True)
prod_df = df.groupby(["Produkt","Popis","ABC rank","GeoSize"]).agg(
    Záznamy=("Produkt","count"),
    Kusy=("Ks na lokačním pohybu","sum")
).reset_index().sort_values("Kusy", ascending=False).head(20)

p1, p2 = st.columns([2,1])
with p1:
    fig = px.bar(prod_df.head(15), x="Kusy", y="Produkt", orientation="h",
                 template=PT, title="Top 15 produktov podľa kusov",
                 color="ABC rank", color_discrete_map=ABC_COLORS)
    fig.update_layout(height=450, margin=dict(t=40,b=10,l=10,r=10), yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig, use_container_width=True)
with p2:
    st.markdown("**Top 10 podľa kusov**")
    for _, row in prod_df.head(10).iterrows():
        ac = ABC_COLORS.get(row["ABC rank"], "#8B949E")
        gc = GEOSIZE_COLORS.get(row["GeoSize"], "#8B949E")
        popis = str(row["Popis"])[:50] + "…" if len(str(row["Popis"])) > 50 else str(row["Popis"])
        st.markdown(f"""
        <div style="background:#161B22;border:1px solid #30363D;border-radius:6px;padding:8px 12px;margin-bottom:6px;">
            <div style="display:flex;justify-content:space-between;">
                <span style="color:#C9D1D9;font-weight:700;font-size:0.85rem;">{row['Produkt']}</span>
                <span style="color:{ac};font-weight:700;">{int(row['Kusy'])} ks</span>
            </div>
            <div style="color:#8B949E;font-size:0.72rem;margin-top:2px;">{popis}</div>
            <div style="margin-top:4px;">
                <span style="background:{ac};color:white;border-radius:3px;padding:1px 7px;font-size:0.7rem;font-weight:700;">ABC: {row['ABC rank']}</span>
                <span style="background:{gc};color:white;border-radius:3px;padding:1px 7px;font-size:0.7rem;font-weight:700;margin-left:4px;">{row['GeoSize']}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 7 – DOKLADY
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">TOP DOKLADY</p>', unsafe_allow_html=True)
doc_df = df.groupby("Skladový doklad").agg(
    Záznamy=("Produkt","count"),
    Kusy=("Ks na lokačním pohybu","sum"),
    SKU=("Produkt","nunique")
).reset_index().sort_values("Kusy", ascending=False).head(15)

fig = px.bar(doc_df, x="Skladový doklad", y="Kusy", template=PT,
             title="Top 15 dokladov podľa kusov",
             color="Záznamy", color_continuous_scale="Blues",
             hover_data=["SKU"])
fig.update_layout(height=350, margin=dict(t=40,b=10,l=10,r=10))
st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════
# SEKCIA 8 – SUROVÉ DÁTA
# ═══════════════════════════════════════════════════════
st.markdown('<p class="section-title">SUROVÉ DÁTA</p>', unsafe_allow_html=True)
with st.expander("📋 Zobraziť/exportovať surové dáta"):
    show_cols = ["Produkt","Popis","GeoSize","ABC rank","Ks na lokačním pohybu",
                 "Skladový doklad","Naskladnil","Čas vzniku","Bedna","Lokace"]
    st.dataframe(df[[c for c in show_cols if c in df.columns]], use_container_width=True, hide_index=True)
    st.download_button("⬇️ Stiahnuť CSV",
                       df.to_csv(index=False).encode("utf-8"),
                       "receive_as_report.csv", "text/csv")
