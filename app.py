import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from io import BytesIO

st.set_page_config(page_title="STOW AS Report", page_icon="📦", layout="wide")

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
    .type-badge {display: inline-block; padding: 2px 10px; border-radius: 4px; font-weight: 700; font-size: 0.75rem; color: white;}
    .card-metric {background: #161B22; border: 1px solid #30363D; border-radius: 8px; padding: 16px; text-align: center;}
    .card-metric .value {font-size: 1.8rem; font-weight: 700; color: #C9D1D9;}
    .card-metric .label {font-size: 0.75rem; color: #8B949E;}
    .card-metric .sub {font-size: 0.7rem; color: #484F58;}
</style>
""", unsafe_allow_html=True)

# ─── Color map ───────────────────────────────────────
TYPE_COLORS = {"VGP": "#2EA043", "SP": "#6366F1", "VV": "#D29922", "REP": "#FF4B4B", "SKL": "#3FB950"}
TYPE_ORDER = ["VGP", "VV", "REP", "SP", "SKL"]

# ─── Persistence – uloží na disk + commitne do GitHub ─
import os, glob, base64, requests

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)

def get_saved_file():
    files = glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
    return files[0] if files else None

def save_file(name, raw_bytes):
    for old in glob.glob(os.path.join(DATA_DIR, "*.xlsx")):
        os.remove(old)
    path = os.path.join(DATA_DIR, name)
    with open(path, "wb") as f:
        f.write(raw_bytes)
    return path

def delete_saved_file():
    for old in glob.glob(os.path.join(DATA_DIR, "*.xlsx")):
        os.remove(old)

def commit_to_github(filename, raw_bytes):
    """Commitne súbor do GitHub repa → prežije redeploy."""
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo  = st.secrets["GITHUB_REPO"]   # napr. "user/stow_as"
    except Exception:
        return False, "GITHUB_TOKEN alebo GITHUB_REPO chýba v secrets"

    api = f"https://api.github.com/repos/{repo}/contents/data/{filename}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github+json"}

    # Zisti sha ak súbor už existuje (pre update)
    sha = None
    r = requests.get(api, headers=headers)
    if r.status_code == 200:
        sha = r.json().get("sha")

    payload = {
        "message": f"update {filename}",
        "content": base64.b64encode(raw_bytes).decode("utf-8"),
    }
    if sha:
        payload["sha"] = sha

    r = requests.put(api, headers=headers, json=payload)
    if r.status_code in (200, 201):
        return True, "Commitnuté do GitHub"
    else:
        return False, f"GitHub API chyba: {r.status_code}"

def parse_excel(path_or_bytes):
    if isinstance(path_or_bytes, (str, os.PathLike)):
        df = pd.read_excel(path_or_bytes, engine='openpyxl')
    else:
        df = pd.read_excel(BytesIO(path_or_bytes), engine='openpyxl')
    df['doklad_type'] = df['Doklad'].astype(str).str.extract(r'^([A-Z]+)')

    # Operátor
    df['_op'] = df['Spustil'] if 'Spustil' in df.columns else df.get('Potvrzovač', 'N/A')

    # Sekcia / poschodie
    if 'Zdroj.lokace' in df.columns:
        df['section'] = df['Zdroj.lokace'].astype(str).str.extract(r'^(\d+[A-Z]+)')
        df['floor'] = df['section'].astype(str).str[0]
    elif 'Stanice lokace' in df.columns:
        df['section'] = df['Stanice lokace'].astype(str)
        df['floor'] = df['section'].str.extract(r'^([A-Z]+)')
    else:
        df['section'] = 'N/A'
        df['floor'] = 'N/A'

    # Dátum
    for col in ['Konec Zpracování na Pob', 'Datum vytvoření', 'Čas vzniku']:
        if col in df.columns:
            df['_date'] = pd.to_datetime(df[col], errors='coerce')
            break
    else:
        df['_date'] = pd.NaT

    return df

# ─── Sidebar ─────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙ Parametre")

    uploaded = st.file_uploader("📂 Nahraj STOW AS Report (.xlsx)", type=["xlsx"])

    if uploaded is not None:
        raw = uploaded.getvalue()
        save_file(uploaded.name, raw)
        ok, msg = commit_to_github(uploaded.name, raw)
        if ok:
            st.success(f"✅ Uložené + {msg}")
        else:
            st.warning(f"💾 Uložené lokálne · ⚠ {msg}")

    saved = get_saved_file()
    if saved is None:
        st.info("Nahraj STOW_AS_REPORT.xlsx pre analýzu")
        st.stop()

    df = parse_excel(saved)
    fname = os.path.basename(saved)
    st.success(f"✅ {fname} · {len(df):,} JBL")

    all_types = sorted(df['doklad_type'].dropna().unique())
    sel_types = st.multiselect("Typ dokladu", all_types, default=all_types)

    all_ops = sorted(df['_op'].dropna().unique())
    sel_ops = st.multiselect("Operátor", all_ops, default=all_ops)

    all_floors = sorted(df['floor'].dropna().unique())
    sel_floors = st.multiselect("Poschodie", all_floors, default=all_floors)

    st.divider()
    if st.button("🗑 Odstrániť súbor"):
        delete_saved_file()
        st.rerun()

# ─── Apply filters ───────────────────────────────────
mask = (
    df['doklad_type'].isin(sel_types) &
    df['_op'].isin(sel_ops) &
    df['floor'].isin(sel_floors)
)
fdf = df[mask].copy()

# ═══════════════════════════════════════════════════════
# HEADER
# ═══════════════════════════════════════════════════════
st.markdown("# 📦 STOW AS Report")
date_min = fdf['_date'].min()
date_max = fdf['_date'].max()
obdobie = f" · Obdobie: <b>{date_min.strftime('%d.%m.%Y')}</b> – <b>{date_max.strftime('%d.%m.%Y')}</b>" if pd.notna(date_min) and pd.notna(date_max) else ""
st.markdown(f'<div class="info-bar">📂 <b>{fname}</b> · {len(df):,} JBL · Filter: {len(fdf):,} JBL{obdobie}</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
# HERO CARD + KPIs
# ═══════════════════════════════════════════════════════
c1, c2 = st.columns([5, 5])

with c1:
    total_jbl = len(fdf)
    total_qty = int(fdf['Množství'].sum())
    st.markdown(f"""
    <div class="green-card">
        <p style="font-size:0.8rem; letter-spacing:1px;">CELKOVÉ JBL · FILTROVANÉ</p>
        <h1>{total_jbl:,}</h1>
        <p>{total_qty:,} kusov celkovo</p>
    </div>
    """, unsafe_allow_html=True)

with c2:
    m1, m2, m3 = st.columns(3)
    m1.metric("Unikátne doklady", f"{fdf['Doklad'].nunique():,}")
    m2.metric("Unikátne produkty", f"{fdf['Produkt'].nunique():,}")
    m3.metric("Operátori", f"{fdf['_op'].nunique()}")
    
    m4, m5, m6 = st.columns(3)
    m4.metric("Priemer ks/JBL", f"{fdf['Množství'].mean():.1f}")
    m5.metric("Medián ks/JBL", f"{fdf['Množství'].median():.0f}")
    m6.metric("Max ks", f"{fdf['Množství'].max():,}")

# ═══════════════════════════════════════════════════════
# ROZDELENIE PODĽA DOKLADU
# ═══════════════════════════════════════════════════════
st.markdown('<div class="section-title">ROZDELENIE PODĽA TYPU DOKLADU</div>', unsafe_allow_html=True)

# Share bar
type_counts = fdf['doklad_type'].value_counts()
type_pcts = (type_counts / len(fdf) * 100).round(1)

fig_bar = go.Figure()
cumx = 0
for t in TYPE_ORDER:
    if t in type_counts.index:
        v = type_counts[t]
        pct = type_pcts[t]
        fig_bar.add_trace(go.Bar(
            x=[pct], y=[""], orientation='h', name=t,
            marker_color=TYPE_COLORS.get(t, "#666"),
            text=f"{t} {pct}%", textposition='inside',
            textfont=dict(color='white', size=12),
            hovertemplate=f"{t}: {v:,} JBL ({pct}%)<extra></extra>"
        ))
fig_bar.update_layout(
    barmode='stack', height=50, margin=dict(l=0,r=0,t=0,b=0),
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
    showlegend=False, xaxis=dict(visible=False), yaxis=dict(visible=False),
)
st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})

# Type cards
cols = st.columns(len(TYPE_ORDER))
for i, t in enumerate(TYPE_ORDER):
    with cols[i]:
        grp = fdf[fdf['doklad_type'] == t]
        if len(grp) == 0:
            st.caption(f"{t}: 0 JBL")
            continue
        pct = len(grp) / len(fdf) * 100
        color = TYPE_COLORS.get(t, "#666")
        
        st.markdown(f"""
        <div style="background:#161B22; border:1px solid #30363D; border-top:3px solid {color}; border-radius:8px; padding:12px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <span style="color:{color}; font-size:1.2rem; font-weight:700;">● {t}</span>
                <span style="color:#8B949E; font-size:0.85rem;">{pct:.1f}%</span>
            </div>
            <div style="border-top:1px solid #30363D; padding-top:8px;">
                <div style="display:flex; justify-content:space-between;"><span style="color:#8B949E; font-size:0.7rem;">JBL</span><span style="color:#C9D1D9; font-weight:700;">{len(grp):,}</span></div>
                <div style="display:flex; justify-content:space-between;"><span style="color:#8B949E; font-size:0.7rem;">Doklady</span><span style="color:#C9D1D9; font-weight:700;">{grp['Doklad'].nunique():,}</span></div>
                <div style="display:flex; justify-content:space-between;"><span style="color:#8B949E; font-size:0.7rem;">Množstvo</span><span style="color:#C9D1D9; font-weight:700;">{int(grp['Množství'].sum()):,}</span></div>
                <div style="display:flex; justify-content:space-between;"><span style="color:#8B949E; font-size:0.7rem;">Produkty</span><span style="color:#C9D1D9; font-weight:700;">{grp['Produkt'].nunique():,}</span></div>
                <div style="display:flex; justify-content:space-between;"><span style="color:#8B949E; font-size:0.7rem;">Operátori</span><span style="color:#C9D1D9; font-weight:700;">{grp['_op'].nunique()}</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
# OPERÁTORI
# ═══════════════════════════════════════════════════════
st.markdown('<div class="section-title">OPERÁTORI – TOP 15</div>', unsafe_allow_html=True)

top_ops = fdf.groupby('_op').agg(
    jbl=('Doklad', 'count'),
    qty=('Množství', 'sum'),
    dok=('Doklad', 'nunique'),
    prod=('Produkt', 'nunique')
).sort_values('jbl', ascending=False).head(15).reset_index()
top_ops['ratio'] = (top_ops['qty'] / top_ops['jbl']).round(1)

# Dominant type per operator
dom_type = fdf.groupby(['_op', 'doklad_type']).size().reset_index(name='cnt')
dom_type = dom_type.loc[dom_type.groupby('_op')['cnt'].idxmax()][['_op', 'doklad_type', 'cnt']]
top_ops = top_ops.merge(dom_type, on='_op', how='left')

c1, c2 = st.columns([6, 4])

with c1:
    fig_ops = go.Figure()
    top_ops_rev = top_ops.iloc[::-1]
    fig_ops.add_trace(go.Bar(
        x=top_ops_rev['jbl'], y=top_ops_rev['_op'].str.split(' ').str[0],
        orientation='h', marker_color='#2EA043',
        text=top_ops_rev['jbl'].apply(lambda x: f"{x:,}"),
        textposition='outside', textfont=dict(color='#3FB950', size=11),
        hovertemplate='%{y}: %{x:,} JBL<extra></extra>'
    ))
    fig_ops.update_layout(
        height=450, margin=dict(l=0,r=40,t=10,b=10),
        paper_bgcolor='#161B22', plot_bgcolor='#161B22',
        xaxis=dict(showgrid=False, visible=False, range=[0, top_ops['jbl'].max()*1.2]),
        yaxis=dict(showgrid=False, tickfont=dict(color='#C9D1D9', size=11)),
        font=dict(color='#C9D1D9'),
    )
    st.plotly_chart(fig_ops, use_container_width=True, config={'displayModeBar': False})

with c2:
    st.markdown("**JBL vs Množstvo – Top 5**")
    for _, row in top_ops.head(5).iterrows():
        short = row['_op'].split(' ')[0]
        st.markdown(f"""
        <div style="background:#161B22; border:1px solid #30363D; border-radius:6px; padding:8px 12px; margin-bottom:6px;">
            <div style="display:flex; justify-content:space-between;">
                <span style="color:#C9D1D9; font-weight:600;">{short}</span>
                <span style="color:#3FB950;">{row['jbl']:,} JBL</span>
            </div>
            <div style="display:flex; justify-content:space-between;">
                <span style="color:#484F58;">⌀ {row['ratio']} ks/JBL</span>
                <span style="color:#8B949E;">{int(row['qty']):,} ks</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("**Dominantný typ dokladu**")
    for _, row in top_ops.head(10).iterrows():
        short = row['_op'].split(' ')[0]
        dt = row['doklad_type']
        dc = TYPE_COLORS.get(dt, '#666')
        st.markdown(f"""
        <div style="display:flex; align-items:center; gap:8px; padding:3px 0;">
            <span style="color:#C9D1D9; width:100px; font-size:0.85rem;">{short}</span>
            <span class="type-badge" style="background:{dc};">{dt}</span>
            <span style="color:#8B949E; font-size:0.85rem;">{int(row['cnt']):,}</span>
        </div>
        """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
# ANALÝZA MNOŽSTVA
# ═══════════════════════════════════════════════════════
st.markdown('<div class="section-title">ANALÝZA MNOŽSTVA</div>', unsafe_allow_html=True)

c1, c2 = st.columns([5, 5])

with c1:
    bins = [0, 1, 5, 20, 100, float('inf')]
    labels = ['1 ks', '2–5 ks', '6–20 ks', '21–100 ks', '100+ ks']
    colors = ['#2EA043', '#6366F1', '#D29922', '#FF4B4B', '#F85149']
    fdf_qty = fdf.copy()
    fdf_qty['bucket'] = pd.cut(fdf_qty['Množství'], bins=bins, labels=labels, right=True)
    bucket_counts = fdf_qty['bucket'].value_counts().reindex(labels)
    
    fig_dist = go.Figure()
    for i, (lbl, cnt) in enumerate(bucket_counts.items()):
        pct = cnt / len(fdf) * 100
        fig_dist.add_trace(go.Bar(
            x=[pct], y=[lbl], orientation='h',
            marker_color=colors[i], name=lbl,
            text=f"{pct:.1f}% · {cnt:,} JBL", textposition='inside',
            textfont=dict(color='white', size=11),
            hovertemplate=f"{lbl}: {cnt:,} JBL ({pct:.1f}%)<extra></extra>"
        ))
    fig_dist.update_layout(
        height=300, margin=dict(l=0,r=0,t=10,b=10), barmode='stack',
        paper_bgcolor='#161B22', plot_bgcolor='#161B22',
        showlegend=False,
        xaxis=dict(visible=False), yaxis=dict(tickfont=dict(color='#C9D1D9', size=12), autorange='reversed'),
        font=dict(color='#C9D1D9'),
    )
    st.plotly_chart(fig_dist, use_container_width=True, config={'displayModeBar': False})

with c2:
    st.markdown("**Množstvo podľa typu dokladu**")
    qty_stats = fdf.groupby('doklad_type')['Množství'].agg(['mean', 'median', 'max', 'sum']).round(1)
    qty_stats.columns = ['Priemer', 'Medián', 'Max', 'Celkom']
    qty_stats = qty_stats.reindex([t for t in TYPE_ORDER if t in qty_stats.index])
    qty_stats['Celkom'] = qty_stats['Celkom'].astype(int)
    qty_stats['Max'] = qty_stats['Max'].astype(int)
    qty_stats['Medián'] = qty_stats['Medián'].astype(int)
    st.dataframe(qty_stats, use_container_width=True)

# ═══════════════════════════════════════════════════════
# SKLADOVÉ SEKCIE
# ═══════════════════════════════════════════════════════
st.markdown('<div class="section-title">SKLADOVÉ SEKCIE</div>', unsafe_allow_html=True)

# Floor bar
floor_counts = fdf['floor'].value_counts()
f2 = floor_counts.get('2', 0)
f3 = floor_counts.get('3', 0)
f_total = f2 + f3 if (f2 + f3) > 0 else 1

progress_pct = f2 / f_total * 100
st.markdown(f"""
<div style="background:#161B22; border:1px solid #30363D; border-radius:8px; padding:12px 16px; margin-bottom:12px;">
    <div style="display:flex; justify-content:space-between; margin-bottom:6px;">
        <span style="color:#818CF8; font-weight:600;">2. poschodie: {f2:,} ({f2/f_total*100:.1f}%)</span>
        <span style="color:#3FB950; font-weight:600;">3. poschodie: {f3:,} ({f3/f_total*100:.1f}%)</span>
    </div>
    <div style="background:#21262D; border-radius:4px; height:20px; overflow:hidden;">
        <div style="background:#6366F1; width:{progress_pct:.1f}%; height:100%; border-radius:4px;"></div>
    </div>
</div>
""", unsafe_allow_html=True)

c1, c2 = st.columns([5, 5])

with c1:
    sec_counts = fdf['section'].value_counts().head(15)
    sec_df = sec_counts.reset_index()
    sec_df.columns = ['section', 'count']
    sec_df['color'] = sec_df['section'].apply(lambda x: '#6366F1' if str(x).startswith('2') else '#2EA043')
    
    fig_sec = go.Figure()
    sec_rev = sec_df.iloc[::-1]
    fig_sec.add_trace(go.Bar(
        x=sec_rev['count'], y=sec_rev['section'],
        orientation='h', marker_color=sec_rev['color'],
        text=sec_rev['count'].apply(lambda x: f"{x:,}"),
        textposition='outside', textfont=dict(size=10, color='#C9D1D9'),
        hovertemplate='%{y}: %{x:,} JBL<extra></extra>'
    ))
    fig_sec.update_layout(
        height=400, margin=dict(l=0,r=40,t=10,b=10),
        paper_bgcolor='#161B22', plot_bgcolor='#161B22',
        xaxis=dict(showgrid=False, visible=False, range=[0, sec_counts.max()*1.15]),
        yaxis=dict(tickfont=dict(color='#C9D1D9', size=11)),
        font=dict(color='#C9D1D9'),
    )
    st.plotly_chart(fig_sec, use_container_width=True, config={'displayModeBar': False})

with c2:
    st.markdown("**Doklad × Poschodie**")
    heat_data = []
    for t in TYPE_ORDER:
        grp = fdf[fdf['doklad_type'] == t]
        if len(grp) == 0:
            continue
        f2t = (grp['floor'] == '2').sum()
        f3t = (grp['floor'] == '3').sum()
        total_t = f2t + f3t if (f2t + f3t) > 0 else 1
        heat_data.append({
            'Doklad': t, '2. posch.': f2t, '3. posch.': f3t,
            '% 2.p': f"{f2t/total_t*100:.0f}%"
        })
    heat_df = pd.DataFrame(heat_data)
    
    for _, row in heat_df.iterrows():
        dc = TYPE_COLORS.get(row['Doklad'], '#666')
        total_r = row['2. posch.'] + row['3. posch.']
        f2pct = row['2. posch.'] / total_r * 100 if total_r > 0 else 0
        is_low = f2pct < 55
        badge_bg = '#FF4B4B' if is_low else '#21262D'
        st.markdown(f"""
        <div style="display:flex; align-items:center; gap:10px; padding:8px 0; border-bottom:1px solid #21262D;">
            <span style="color:{dc}; font-weight:700; width:50px; font-size:1rem;">{row['Doklad']}</span>
            <span style="color:#818CF8; width:65px; font-weight:600;">{row['2. posch.']:,}</span>
            <span style="color:#3FB950; width:55px;">{row['3. posch.']:,}</span>
            <span style="background:{badge_bg}; color:white; padding:2px 8px; border-radius:4px; font-weight:700; font-size:0.85rem;">{row['% 2.p']}</span>
            <div style="flex:1; background:#21262D; height:10px; border-radius:4px; overflow:hidden;">
                <div style="background:#6366F1; width:{f2pct:.0f}%; height:100%;"></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
# PRODUKTOVÁ ANALÝZA
# ═══════════════════════════════════════════════════════
st.markdown('<div class="section-title">PRODUKTOVÁ ANALÝZA</div>', unsafe_allow_html=True)

c1, c2 = st.columns([6, 4])

with c1:
    top_prod = fdf['Produkt'].value_counts().head(15).reset_index()
    top_prod.columns = ['Produkt', 'count']
    
    # Get dominant doklad type for each product
    prod_type = fdf.groupby(['Produkt', 'doklad_type']).size().reset_index(name='cnt')
    prod_type = prod_type.loc[prod_type.groupby('Produkt')['cnt'].idxmax()][['Produkt', 'doklad_type']]
    top_prod = top_prod.merge(prod_type, on='Produkt', how='left')
    top_prod['color'] = top_prod['doklad_type'].map(TYPE_COLORS).fillna('#666')
    
    fig_prod = go.Figure()
    tp_rev = top_prod.iloc[::-1]
    fig_prod.add_trace(go.Bar(
        x=tp_rev['count'], y=tp_rev['Produkt'],
        orientation='h', marker_color=tp_rev['color'],
        text=tp_rev['count'].apply(lambda x: f"{x:,}"),
        textposition='outside', textfont=dict(color='#D29922', size=11),
        hovertemplate='%{y}: %{x:,} JBL<extra></extra>'
    ))
    fig_prod.update_layout(
        height=450, margin=dict(l=0,r=40,t=10,b=10),
        paper_bgcolor='#161B22', plot_bgcolor='#161B22',
        xaxis=dict(showgrid=False, visible=False, range=[0, top_prod['count'].max()*1.2]),
        yaxis=dict(tickfont=dict(color='#C9D1D9', size=10)),
        font=dict(color='#C9D1D9'),
    )
    st.plotly_chart(fig_prod, use_container_width=True, config={'displayModeBar': False})

with c2:
    st.markdown("**Top 3 produkty podľa typu**")
    for t in TYPE_ORDER:
        grp = fdf[fdf['doklad_type'] == t]
        if len(grp) == 0:
            continue
        top3 = grp['Produkt'].value_counts().head(3)
        dc = TYPE_COLORS.get(t, '#666')
        items = " · ".join([f"{p} ({c})" for p, c in top3.items()])
        st.markdown(f"""
        <div style="background:#161B22; border-left:3px solid {dc}; border-radius:0 6px 6px 0; padding:8px 12px; margin-bottom:6px;">
            <span style="color:{dc}; font-weight:700; font-size:0.95rem;">● {t}</span>
            <div style="color:#8B949E; font-size:0.8rem; margin-top:4px;">{items}</div>
        </div>
        """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
# RAW DATA
# ═══════════════════════════════════════════════════════
with st.expander("📋 Zobraziť surové dáta"):
    st.dataframe(fdf, use_container_width=True, height=400)
    st.download_button("⬇ Stiahnuť filtrované dáta (CSV)", fdf.to_csv(index=False).encode('utf-8'), "stow_filtered.csv", "text/csv")
