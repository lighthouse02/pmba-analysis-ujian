import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import dash
from dash import dcc, html
import dash_bootstrap_components as dbc

# ── Load & prepare data ──────────────────────────────────────────────────────
df = pd.read_excel("all-daftar-pertandingan.xlsx")
df["Kayıt Tarihi"] = pd.to_datetime(df["Kayıt Tarihi"], dayfirst=True)

# Deduplicate: keep first registration per email for unique-person counts
df_sorted = df.sort_values("Kayıt Tarihi")
df_sorted["_ek"] = df_sorted["E-posta"].str.lower().str.strip()
_first_mask = ~df_sorted.duplicated(subset=["_ek"], keep="first")
df_unique   = df_sorted[_first_mask]

daily = df_unique.groupby("Kayıt Tarihi").size().reset_index(name="Count")
daily = daily.sort_values("Kayıt Tarihi")
daily["Cumulative"] = daily["Count"].cumsum()

# ── Story milestones ─────────────────────────────────────────────────────────
milestones = [
    {"date": "2025-12-08", "emoji": "📅", "label": "Kumpulan Dicipta"},
    {"date": "2025-12-25", "emoji": "⏸️", "label": "Tangguh Pertandingan"},
    {"date": "2026-03-25", "emoji": "📝", "label": "Open Public Quiz"},
    {"date": "2026-02-27", "emoji": "🏆", "label": "Pengiktirafan PAJSK"},
    {"date": "2026-03-07", "emoji": "📖", "label": "Live Bedah Buku"},
    {"date": "2026-03-14", "emoji": "✏️", "label": "Ujian Percubaan"},
    {"date": "2026-03-27", "emoji": "🎓", "label": "Ujian Sebenar"},
]

# ── Negeri normalisation ──────────────────────────────────────────────────────
NEGERI_NORM = {
    "johor":                        "Johor",
    "negeri johor":                 "Johor",
    "kedah":                        "Kedah",
    "kelantan":                     "Kelantan",
    "melaka":                       "Melaka",
    "malacca":                      "Melaka",
    "negeri sembilan":              "Negeri Sembilan",
    "sembilan":                     "Negeri Sembilan",
    "pahang":                       "Pahang",
    "perak":                        "Perak",
    "perlis":                       "Perlis",
    "pulau pinang":                 "Pulau Pinang",
    "penang":                       "Pulau Pinang",
    "sabah":                        "Sabah",
    "sarawak":                      "Sarawak",
    "selangor":                     "Selangor",
    "terengganu":                   "Terengganu",
    "kuala lumpur":                 "Kuala Lumpur",
    "w.p. kuala lumpur":            "Kuala Lumpur",
    "wilayah persekutuan kuala lumpur": "Kuala Lumpur",
    "labuan":                       "Labuan",
    "w.p. labuan":                  "Labuan",
    "putrajaya":                    "Putrajaya",
    "w.p. putrajaya":               "Putrajaya",
}

def _norm_negeri(val):
    if pd.isna(val):
        return "Tidak Diketahui"
    return NEGERI_NORM.get(str(val).strip().lower(), str(val).strip().title())

df["Negeri"] = df["İl"].apply(_norm_negeri)

# ── Gender detection ──────────────────────────────────────────────────────────
def _detect_gender(name):
    if pd.isna(name):
        return "Tidak Diketahui"
    n = str(name).upper()
    if " BIN " in n or " BIN@" in n or n.endswith(" BIN"):
        return "Lelaki"
    if (" BINTI " in n or " BINTI@" in n or n.endswith(" BINTI")
            or " BT " in n or " BT@" in n or n.endswith(" BT")
            or " BTE " in n or " BINTE " in n):
        return "Perempuan"
    return "Tidak Diketahui"

df["Gender"] = df["Ad"].apply(_detect_gender)

# ── Aggregated counts ─────────────────────────────────────────────────────────
negeri_counts = (
    df[df["Negeri"] != "Tidak Diketahui"]["Negeri"]
    .value_counts()
    .reset_index()
)
negeri_counts.columns = ["Negeri", "Count"]
negeri_counts = negeri_counts.sort_values("Count", ascending=True)   # for horizontal bar

gender_counts = df["Gender"].value_counts()

# ── Timeline by negeri ────────────────────────────────────────────────────────
# Top 8 states by registration (rest grouped as "Lain-lain")
_TOP_N = 8
_top_negeri = (
    negeri_counts.sort_values("Count", ascending=False)
    .head(_TOP_N)["Negeri"].tolist()
)

df["NegeriGroup"] = df["Negeri"].apply(
    lambda n: n if n in _top_negeri else "Lain-lain"
)

_by_dn = (
    df[df["Negeri"] != "Tidak Diketahui"]
    .groupby(["Kayıt Tarihi", "NegeriGroup"])
    .size().reset_index(name="Count")
    .sort_values("Kayıt Tarihi")
)

# Full date range index so every state has every date (fill 0 for missing)
_all_dates  = pd.date_range(df["Kayıt Tarihi"].min(), df["Kayıt Tarihi"].max())
_all_groups = _top_negeri + ["Lain-lain"]
_idx        = pd.MultiIndex.from_product([_all_dates, _all_groups],
                                          names=["Kayıt Tarihi", "NegeriGroup"])
_by_dn_full = (
    _by_dn.set_index(["Kayıt Tarihi", "NegeriGroup"])
    .reindex(_idx, fill_value=0)
    .reset_index()
)

# Compute cumulative per group
_cum_parts = []
for grp in _all_groups:
    sub = _by_dn_full[_by_dn_full["NegeriGroup"] == grp].copy()
    sub["Cumulative"] = sub["Count"].cumsum()
    _cum_parts.append(sub)
negeri_timeline_df = pd.concat(_cum_parts, ignore_index=True)

# Weekly heatmap: resample daily counts to weekly, pivot negeri × week
_hm_raw = (
    df[df["Negeri"] != "Tidak Diketahui"]
    .groupby(["Kayıt Tarihi", "Negeri"])
    .size().reset_index(name="Count")
)
_hm_raw["Week"] = _hm_raw["Kayıt Tarihi"].dt.to_period("W").apply(
    lambda p: p.start_time.strftime("%d %b '%y")
)
negeri_heatmap_df = (
    _hm_raw.groupby(["Negeri", "Week"])["Count"].sum()
    .unstack(fill_value=0)
)  # rows=Negeri, cols=Week (sorted chronologically)
# sort rows by total desc
negeri_heatmap_df = negeri_heatmap_df.loc[
    negeri_heatmap_df.sum(axis=1).sort_values(ascending=False).index
]

# ── Malaysia state centroids (for bubble map) ─────────────────────────────────
STATE_CENTROIDS = {
    "Johor":           (1.90,  103.50),
    "Kedah":           (6.05,  100.40),
    "Kelantan":        (5.30,  102.00),
    "Melaka":          (2.20,  102.25),
    "Negeri Sembilan": (2.75,  102.00),
    "Pahang":          (3.80,  103.00),
    "Perak":           (4.60,  101.05),
    "Perlis":          (6.30,  100.20),
    "Pulau Pinang":    (5.35,  100.40),
    "Sabah":           (5.20,  117.00),
    "Sarawak":         (2.50,  113.00),
    "Selangor":        (3.30,  101.50),
    "Terengganu":      (5.20,  103.00),
    "Kuala Lumpur":    (3.15,  101.70),
    "Labuan":          (5.28,  115.22),
    "Putrajaya":       (2.93,  101.69),
}

# ── Exam data ─────────────────────────────────────────────────────────────────
dp_raw = pd.read_excel("result ujian percubaan.xlsx")
ds_raw = pd.read_excel("result ujian sebenar.xlsx")
for _ex in (dp_raw, ds_raw):
    _ex["Total_Jawab"] = _ex["Doğru"] + _ex["Yanlış"] + _ex["Boş"]
dp_full = dp_raw[dp_raw["Total_Jawab"] == 40].copy()
ds_full = ds_raw[ds_raw["Total_Jawab"] == 40].copy()
PASS_MARK = 40.0  # Puan ≥ 40 = lulus (≥ 40 % of 100)

dp_full["email_key"] = dp_full["E-posta"].str.lower().str.strip()
ds_full["email_key"] = ds_full["E-posta"].str.lower().str.strip()
dp_raw["email_key"]  = dp_raw["E-posta"].str.lower().str.strip()
dr_emails       = set(df["E-posta"].dropna().str.lower().str.strip())
dp_full_emails  = set(dp_full["email_key"])
ds_full_emails  = set(ds_full["email_key"])
dp_any_emails   = set(dp_raw["email_key"].dropna())   # anyone who appeared in percubaan

# ── Analysis 1 : Top-30 consistency (join by External ID) ───────────────────
dp_top30 = (dp_full.nlargest(30, "Puan")[["İsim", "Puan", "External ID"]]
            .reset_index(drop=True))
ds_top30 = (ds_full.nlargest(30, "Puan")[["İsim", "Puan", "External ID"]]
            .reset_index(drop=True))
top30_both = (set(dp_top30["External ID"].dropna().astype(int))
              & set(ds_top30["External ID"].dropna().astype(int)))

# ── Analysis 2 : Attendance funnel ───────────────────────────────────────────
# Use dp_full / ds_full (completed all 40 questions) — consistent definition
f_both       = len(dr_emails & dp_full_emails & ds_full_emails)
f_perc_only  = len(dr_emails & (dp_full_emails - ds_full_emails))
f_seb_only   = len(dr_emails & (ds_full_emails - dp_full_emails))
f_none       = len(dr_emails - dp_full_emails - ds_full_emails)

# ── Analysis 3 : Score improvement (join by External ID) ────────────────────
score_merged = (
    dp_full[["External ID", "İsim", "Puan"]]
    .rename(columns={"Puan": "Puan_P", "İsim": "Nama"})
    .merge(ds_full[["External ID", "Puan"]].rename(columns={"Puan": "Puan_S"}),
           on="External ID", how="inner")
)
score_merged["Delta"] = score_merged["Puan_S"] - score_merged["Puan_P"]

# ── Analysis 4 : Registration date vs Sebenar score ──────────────────────────
reg_score_df = (
    df_unique[["_ek", "Kayıt Tarihi"]]
    .merge(ds_full[["email_key", "Puan"]].rename(columns={"email_key": "_ek"}),
           on="_ek", how="inner")
    .sort_values("Kayıt Tarihi")
    .reset_index(drop=True)
)
_REG_PERIODS = [
    ("Okt–Nov 2025",   "2025-10-01", "2025-11-30"),
    ("Dis 2025",       "2025-12-01", "2025-12-31"),
    ("Jan–Feb 2026",   "2026-01-01", "2026-02-28"),
    ("1–7 Mac 2026",   "2026-03-01", "2026-03-07"),
    ("8–14 Mac 2026",  "2026-03-08", "2026-03-14"),
    ("15–27 Mac 2026", "2026-03-15", "2026-03-27"),
]
_period_rows = []
for label, start, end in _REG_PERIODS:
    mask = (reg_score_df["Kayıt Tarihi"] >= start) & (reg_score_df["Kayıt Tarihi"] <= end)
    sub = reg_score_df[mask]["Puan"]
    if len(sub):
        _period_rows.append({
            "Period": label, "n": len(sub),
            "mean": round(sub.mean(), 1),
            "median": sub.median(),
            "pass_pct": round((sub >= PASS_MARK).mean() * 100, 1),
        })
reg_period_df = pd.DataFrame(_period_rows)

# ── Analysis 5 : Bell curve prep ─────────────────────────────────────────────
dp_pass_n   = int((dp_full["Puan"] >= PASS_MARK).sum())
dp_fail_n   = int((dp_full["Puan"] < PASS_MARK).sum())
dp_pass_pct = round(dp_pass_n / len(dp_full) * 100, 1)
ds_pass_n   = int((ds_full["Puan"] >= PASS_MARK).sum())
ds_fail_n   = int((ds_full["Puan"] < PASS_MARK).sum())
ds_pass_pct = round(ds_pass_n / len(ds_full) * 100, 1)

# ── Analysis 6 : Anomaly detection ───────────────────────────────────────────
anomaly_s = ds_raw[(ds_raw["Total_Jawab"] == 0) & ds_raw["Bitiş"].notna()].copy()
anomaly_s["Punca"] = "Hantar tanpa jawab (masalah teknikal)"

# ── Appendix : Registration duplicate anomalies ───────────────────────────────
import unicodedata as _ud, re as _re

def _norm_name(s):
    if pd.isna(s): return ""
    s = str(s).strip().lower()
    s = _ud.normalize("NFKD", s)
    s = "".join(c for c in s if not _ud.combining(c))
    return _re.sub(r"\s+", " ", s).strip()

df["_name_key"] = df["Ad"].apply(_norm_name)
df["_email_key"] = df["E-posta"].str.lower().str.strip()

# Type A : same email, different names (parent registers multiple children)
email_multi = (df.groupby("_email_key")
               .filter(lambda x: x["_name_key"].nunique() > 1)
               .sort_values(["_email_key", "Kayıt Tarihi"]))
_n_email_affected = email_multi["_email_key"].nunique()
_n_email_rows     = len(email_multi)

# Type B : same name, different emails (double-registered / genuine namesakes)
name_multi = (df.groupby("_name_key")
              .filter(lambda x: x["_email_key"].nunique() > 1)
              .sort_values(["_name_key", "Kayıt Tarihi"]))
_n_name_affected = name_multi["_name_key"].nunique()
_n_name_rows     = len(name_multi)

# Sebenar participation lookup for anomaly tables
_ds_any_emails = set(ds_raw["E-posta"].dropna().str.lower().str.strip())
def _seb_status(email):
    if pd.isna(email): return "—"
    e = str(email).lower().strip()
    if e in ds_full_emails: return "✔ Lengkap"
    if e in _ds_any_emails: return "⚠ Hadir (0 jawapan)"
    return "✘ Tidak hadir"
email_multi["Status Sebenar"] = email_multi["_email_key"].apply(_seb_status)
name_multi["Status Sebenar"]  = name_multi["_email_key"].apply(_seb_status)

# Negeri breakdown for anomalies
_anm_email_negeri = email_multi.groupby("Negeri").size().reset_index(name="Count_A")
_anm_name_negeri  = name_multi.groupby("Negeri").size().reset_index(name="Count_B")
_anm_negeri = (
    _anm_email_negeri
    .merge(_anm_name_negeri, on="Negeri", how="outer")
    .fillna(0)
)
_anm_negeri["Count_A"] = _anm_negeri["Count_A"].astype(int)
_anm_negeri["Count_B"] = _anm_negeri["Count_B"].astype(int)
_anm_negeri["Total"]   = _anm_negeri["Count_A"] + _anm_negeri["Count_B"]
_anm_negeri = _anm_negeri.sort_values("Total", ascending=True)

# ── KPI values ────────────────────────────────────────────────────────────────
total     = len(df_unique)
peak_row  = daily.loc[daily["Count"].idxmax()]
peak_day  = peak_row["Kayıt Tarihi"].strftime("%d %b %Y")
peak_val  = int(peak_row["Count"])
_n_days   = (daily["Kayıt Tarihi"].max() - daily["Kayıt Tarihi"].min()).days + 1
avg_daily = round(total / _n_days, 1)

# ── Shared styles ─────────────────────────────────────────────────────────────
BG       = "#0b0b16"
CARD_BG  = "#13131f"
CARD_BG2 = "#1a1a2e"
BORDER   = "rgba(255,255,255,0.07)"
SECTION_LABEL_STYLE = {
    "color": "#555", "fontSize": "0.7rem", "letterSpacing": "0.12em",
    "fontWeight": "700", "textTransform": "uppercase", "marginBottom": "10px",
}

# ── KPI card component ────────────────────────────────────────────────────────
def kpi_card(title, value, sub, color):
    return html.Div(
        style={
            "backgroundColor": CARD_BG,
            "borderRadius": "14px",
            "padding": "0",
            "overflow": "hidden",
            "boxShadow": f"0 4px 24px rgba(0,0,0,0.4), 0 0 0 1px rgba(255,255,255,0.05)",
            "height": "100%",
            "position": "relative",
        },
        children=[
            # Colored top strip
            html.Div(style={
                "height": "4px",
                "background": f"linear-gradient(90deg, {color}, {color}88)",
            }),
            # Card body
            html.Div(
                style={"padding": "20px 24px 22px 24px"},
                children=[
                    html.P(title, style={
                        "color": "#666", "fontSize": "0.72rem", "letterSpacing": "0.1em",
                        "textTransform": "uppercase", "fontWeight": "600", "marginBottom": "10px",
                    }),
                    html.H2(str(value), style={
                        "color": color, "fontWeight": "800", "fontSize": "2.2rem",
                        "margin": "0 0 6px 0", "lineHeight": "1",
                    }),
                    html.P(sub, style={"color": "#555", "fontSize": "0.8rem", "margin": "0"}),
                ],
            ),
        ],
    )

# ── Chart: Trend ──────────────────────────────────────────────────────────────
def build_main_chart():
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(go.Scatter(
        x=daily["Kayıt Tarihi"], y=daily["Cumulative"],
        name="Kumulatif",
        fill="tozeroy",
        line=dict(color="#4E9AF1", width=3),
        fillcolor="rgba(78,154,241,0.10)",
        hovertemplate="<b>%{x|%d %b %Y}</b><br>Kumulatif: <b>%{y:,}</b><extra></extra>",
    ), secondary_y=False)

    fig.add_trace(go.Bar(
        x=daily["Kayıt Tarihi"], y=daily["Count"],
        name="Harian",
        marker_color="rgba(244,162,97,0.45)",
        marker_line_width=0,
        hovertemplate="<b>%{x|%d %b %Y}</b><br>Hari ini: <b>%{y}</b><extra></extra>",
    ), secondary_y=True)

    # Stagger heights (paper y) so labels of nearby events never overlap.
    # Three tiers: top (0.97), mid (0.80), low (0.63).
    # Assign each milestone a tier based on proximity to its neighbours.
    label_y = [0.97, 0.80, 0.63, 0.97, 0.80, 0.63, 0.97]   # one per milestone

    for i, m in enumerate(milestones):
        ms_date = pd.to_datetime(m["date"])

        # Subtle vertical rule — no arrow, no visual clutter
        fig.add_shape(
            type="line",
            x0=ms_date, x1=ms_date,
            y0=0, y1=1,
            xref="x", yref="paper",
            line=dict(color="rgba(244,162,97,0.30)", width=1.2, dash="dot"),
        )

        # Label pinned to paper-space at the staggered height; no arrows
        fig.add_annotation(
            x=ms_date,
            y=label_y[i],
            xref="x",
            yref="paper",
            text=f"{m['emoji']} <b>{m['label']}</b>",
            showarrow=False,
            xanchor="center",
            yanchor="middle",
            bgcolor="rgba(14,14,24,0.90)",
            bordercolor="rgba(244,162,97,0.65)",
            borderwidth=1,
            borderpad=7,
            font=dict(size=11, color="#f4c261"),
        )

    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        legend=dict(orientation="h", y=1.06, x=0,
                    font=dict(color="#888", size=12), bgcolor="rgba(0,0,0,0)"),
        hovermode="x unified",
        margin=dict(l=70, r=70, t=40, b=50),
        xaxis=dict(gridcolor="rgba(255,255,255,0.04)", tickfont=dict(color="#666", size=11),
                   showline=False, tickformat="%b '%y"),
        yaxis=dict(title="Kumulatif Pendaftaran", title_font=dict(color="#4E9AF1", size=11),
                   gridcolor="rgba(255,255,255,0.04)", tickfont=dict(color="#4E9AF1", size=11),
                   tickformat=","),
        yaxis2=dict(title="Pendaftaran Harian", title_font=dict(color="#F4A261", size=11),
                    tickfont=dict(color="#F4A261", size=11), showgrid=False),
    )
    return fig

# ── Chart: Negeri horizontal bar ──────────────────────────────────────────────
def build_negeri_bar():
    # colour gradient: fewer → dark blue, more → bright blue
    max_c = negeri_counts["Count"].max()
    colors = [
        f"rgba(78,{int(100 + 155 * (c / max_c))},{int(200 + 55 * (c / max_c))},0.85)"
        for c in negeri_counts["Count"]
    ]

    fig = go.Figure(go.Bar(
        x=negeri_counts["Count"],
        y=negeri_counts["Negeri"],
        orientation="h",
        marker=dict(color=colors, line_width=0),
        text=negeri_counts["Count"].apply(lambda v: f"{v:,}"),
        textposition="outside",
        textfont=dict(color="#aaa", size=11),
        hovertemplate="<b>%{y}</b><br>Pendaftar: <b>%{x:,}</b><extra></extra>",
    ))

    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        margin=dict(l=10, r=80, t=20, b=20),
        xaxis=dict(showgrid=False, showticklabels=False, showline=False),
        yaxis=dict(tickfont=dict(color="#ccc", size=12), showgrid=False),
        height=480,
        bargap=0.25,
    )
    return fig

# ── Chart: Malaysia bubble map ────────────────────────────────────────────────
def build_malaysia_map():
    rows = []
    for _, row in negeri_counts.iterrows():
        coords = STATE_CENTROIDS.get(row["Negeri"])
        if coords:
            rows.append({
                "Negeri": row["Negeri"],
                "Count": row["Count"],
                "lat": coords[0],
                "lon": coords[1],
            })
    map_df = pd.DataFrame(rows)

    max_c  = map_df["Count"].max()
    # Bubble size: scale between 10 and 55
    map_df["size"] = 10 + 45 * (map_df["Count"] / max_c)

    fig = go.Figure(go.Scattermap(
        lat=map_df["lat"],
        lon=map_df["lon"],
        mode="markers+text",
        marker=dict(
            size=map_df["size"],
            color=map_df["Count"],
            colorscale=[[0, "#0d2a4a"], [0.4, "#1e5fa0"], [0.7, "#3b82f6"], [1, "#93c5fd"]],
            showscale=True,
            colorbar=dict(
                title=dict(text="Pendaftar", font=dict(color="#888", size=11)),
                tickfont=dict(color="#888", size=10),
                bgcolor="rgba(0,0,0,0)",
                thickness=12,
                len=0.6,
            ),
            opacity=0.85,
            sizemode="diameter",
        ),
        text=map_df["Negeri"],
        textfont=dict(size=9, color="rgba(255,255,255,0.75)"),
        textposition="top center",
        customdata=map_df[["Count"]].values,
        hovertemplate="<b>%{text}</b><br>Pendaftar: <b>%{customdata[0]:,}</b><extra></extra>",
    ))

    fig.update_layout(
        map=dict(
            style="carto-darkmatter",
            center={"lat": 4.0, "lon": 109.5},
            zoom=3.8,
        ),
        paper_bgcolor=CARD_BG,
        margin=dict(l=0, r=0, t=0, b=0),
        height=480,
    )
    return fig

# ── Chart: Gender donut ───────────────────────────────────────────────────────
def build_gender_donut():
    labels = list(gender_counts.index)
    values = list(gender_counts.values)
    colors = {
        "Lelaki":           "#4E9AF1",
        "Perempuan":        "#F472B6",
        "Tidak Diketahui":  "#3a3a52",
    }
    color_list = [colors.get(l, "#888") for l in labels]

    fig = go.Figure(go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        marker=dict(colors=color_list, line=dict(color=CARD_BG, width=3)),
        textinfo="percent",
        textfont=dict(size=13, color="white"),
        hovertemplate="<b>%{label}</b><br>%{value:,} peserta<br>%{percent}<extra></extra>",
        sort=False,
    ))

    fig.update_layout(
        paper_bgcolor=CARD_BG,
        plot_bgcolor=CARD_BG,
        margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(
            orientation="v", x=0.5, y=-0.08, xanchor="center",
            font=dict(color="#aaa", size=12), bgcolor="rgba(0,0,0,0)",
        ),
        annotations=[dict(
            text=f"<b>{total:,}</b><br><span style='font-size:11px;color:#666'>peserta</span>",
            showarrow=False, font=dict(size=16, color="white"),
            x=0.5, y=0.5,
        )],
    )
    return fig

# ── Milestone table rows (HTML) ───────────────────────────────────────────────
def milestone_table_rows():
    rows = []
    for i, m in enumerate(milestones):
        ms_date   = pd.to_datetime(m["date"])
        past      = daily[daily["Kayıt Tarihi"] <= ms_date]
        cum       = int(past["Cumulative"].iloc[-1]) if not past.empty else 0
        row_bg    = CARD_BG if i % 2 == 0 else CARD_BG2
        cell      = {"padding": "14px 24px", "verticalAlign": "middle"}

        rows.append(html.Tr([
            html.Td(ms_date.strftime("%d %b %Y"),
                    style={**cell, "color": "#666", "fontSize": "0.82rem",
                           "whiteSpace": "nowrap"}),
            html.Td([
                html.Span(m["emoji"] + "\u2009"),
                html.Span(m["label"], style={"color": "#dde", "fontWeight": "600",
                                              "fontSize": "0.93rem"}),
            ], style=cell),
            html.Td(f"{cum:,}",
                    style={**cell, "color": "#4E9AF1", "fontWeight": "700",
                           "fontSize": "1rem", "textAlign": "right",
                           "whiteSpace": "nowrap"}),
        ], style={"backgroundColor": row_bg, "borderBottom": f"1px solid {BORDER}"}))
    return rows

# ── Gender stat cards ──────────────────────────────────────────────────────────
def gender_stat(label, count, pct, color):
    return html.Div(
        style={
            "backgroundColor": CARD_BG2, "borderRadius": "12px",
            "padding": "20px 24px", "borderLeft": f"3px solid {color}",
            "marginBottom": "12px",
        },
        children=[
            html.P(label, style={"color": "#666", "fontSize": "0.72rem",
                                  "letterSpacing": "0.1em", "textTransform": "uppercase",
                                  "fontWeight": "600", "marginBottom": "6px"}),
            html.Div([
                html.Span(f"{count:,}", style={"color": color, "fontWeight": "800",
                                                "fontSize": "1.8rem", "lineHeight": "1"}),
                html.Span(f"  {pct:.1f}%", style={"color": "#555", "fontSize": "0.9rem",
                                                    "marginLeft": "8px"}),
            ]),
        ],
    )

# ── Chart: Cumulative negeri stacked area ────────────────────────────────────
def build_negeri_stacked_area():
    palette = [
        "#4E9AF1", "#F4A261", "#2A9D8F", "#E76F51", "#9b8fe8",
        "#F472B6", "#34D399", "#FBBF24", "#94A3B8",
    ]
    fig = go.Figure()
    ordered = _top_negeri + ["Lain-lain"]
    for i, grp in enumerate(reversed(ordered)):
        sub = negeri_timeline_df[negeri_timeline_df["NegeriGroup"] == grp]
        col = palette[i % len(palette)]
        fig.add_trace(go.Scatter(
            x=sub["Kayıt Tarihi"], y=sub["Cumulative"],
            name=grp,
            stackgroup="one",
            mode="lines",
            line=dict(width=0.5, color=col),
            hovertemplate=f"<b>{grp}</b><br>%{{x|%d %b %Y}}<br>Kumulatif: <b>%{{y:,}}</b><extra></extra>",
        ))

    for m in milestones:
        fig.add_vline(
            x=pd.to_datetime(m["date"]), line_width=1,
            line_dash="dot", line_color="rgba(255,255,255,0.18)",
        )
        fig.add_annotation(
            x=pd.to_datetime(m["date"]), y=1, yref="paper",
            text=f"<b>{m['emoji']}</b>", showarrow=False,
            font=dict(size=13), yanchor="bottom",
            bgcolor="rgba(0,0,0,0)", borderwidth=0,
        )

    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        legend=dict(
            orientation="v", x=1.01, y=1, xanchor="left",
            font=dict(color="#ccc", size=11), bgcolor="rgba(0,0,0,0)",
            traceorder="normal",
        ),
        hovermode="x unified",
        margin=dict(l=60, r=140, t=30, b=50),
        xaxis=dict(
            gridcolor="rgba(255,255,255,0.04)",
            tickfont=dict(color="#666", size=11),
            tickformat="%b '%y", showline=False,
        ),
        yaxis=dict(
            title="Kumulatif",
            title_font=dict(color="#888", size=11),
            gridcolor="rgba(255,255,255,0.04)",
            tickfont=dict(color="#888", size=11),
            tickformat=",",
        ),
    )
    return fig


# ── Chart: Weekly heatmap negeri × week ──────────────────────────────────────
def build_negeri_heatmap():
    z      = negeri_heatmap_df.values.tolist()
    x_labs = list(negeri_heatmap_df.columns)
    y_labs = list(negeri_heatmap_df.index)

    text_vals = [
        [f"<b>{y_labs[r]}</b><br>Minggu: {x_labs[c]}<br>Pendaftar: <b>{z[r][c]}</b>"
         for c in range(len(x_labs))]
        for r in range(len(y_labs))
    ]

    fig = go.Figure(go.Heatmap(
        z=z, x=x_labs, y=y_labs,
        colorscale=[[0, "#0d0d1a"], [0.1, "#0d2a4a"],
                    [0.4, "#1e5fa0"], [0.7, "#3b82f6"], [1, "#93c5fd"]],
        showscale=True,
        colorbar=dict(
            title=dict(text="Pendaftar", font=dict(color="#666", size=10)),
            tickfont=dict(color="#666", size=9),
            thickness=10, len=0.8,
        ),
        text=text_vals, hoverinfo="text",
        xgap=2, ygap=2,
    ))

    for m in milestones:
        ms_str = pd.to_datetime(m["date"]).to_period("W").start_time.strftime("%d %b '%y")
        if ms_str in x_labs:
            fig.add_vline(
                x=x_labs.index(ms_str) - 0.5,
                line_width=1.5, line_dash="dot",
                line_color="rgba(244,194,97,0.6)",
            )

    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        margin=dict(l=10, r=60, t=20, b=80),
        xaxis=dict(
            tickfont=dict(color="#666", size=9),
            tickangle=-45, showgrid=False, nticks=18,
        ),
        yaxis=dict(
            tickfont=dict(color="#ccc", size=11),
            showgrid=False, autorange="reversed",
        ),
    )
    return fig


# ── Chart: Gender × Negeri stacked bar ───────────────────────────────────────
def _build_gender_negeri_chart():
    top10 = negeri_counts.sort_values("Count", ascending=False).head(10)["Negeri"].tolist()
    sub   = df[df["Negeri"].isin(top10) & (df["Gender"] != "Tidak Diketahui")]
    gn    = (sub.groupby(["Negeri", "Gender"])
               .size().reset_index(name="Count"))
    order = (gn.groupby("Negeri")["Count"].sum()
               .sort_values(ascending=True).index.tolist())

    fig = go.Figure()
    g_colors = {"Lelaki": "#4E9AF1", "Perempuan": "#F472B6"}

    for gender in ["Lelaki", "Perempuan"]:
        sub2 = gn[gn["Gender"] == gender]
        sub2 = sub2.set_index("Negeri").reindex(order).fillna(0).reset_index()
        fig.add_trace(go.Bar(
            y=sub2["Negeri"],
            x=sub2["Count"],
            name=gender,
            orientation="h",
            marker_color=g_colors[gender],
            hovertemplate=f"<b>%{{y}}</b><br>{gender}: <b>%{{x:,}}</b><extra></extra>",
        ))

    fig.update_layout(
        barmode="stack",
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        legend=dict(orientation="h", y=1.06, x=0,
                    font=dict(color="#aaa", size=12), bgcolor="rgba(0,0,0,0)"),
        margin=dict(l=10, r=50, t=20, b=20),
        xaxis=dict(showgrid=False, showticklabels=False, showline=False),
        yaxis=dict(tickfont=dict(color="#ccc", size=11), showgrid=False),
        bargap=0.25,
    )
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# Exam analysis charts
# ══════════════════════════════════════════════════════════════════════════════

# ── Analysis 1 : Top-30 table ─────────────────────────────────────────────────
def build_top30_table():
    rows = []
    for i, r in dp_top30.iterrows():
        eid = int(r["External ID"]) if pd.notna(r["External ID"]) else None
        in_both = eid in top30_both
        color_p = "#22c55e" if in_both else "#e2e8f0"
        rows.append(
            html.Tr([
                html.Td(i + 1, style={"color": "#9ca3af", "width": "36px", "textAlign": "center"}),
                html.Td(r["İsim"], style={"color": color_p, "fontWeight": "600" if in_both else "normal"}),
                html.Td(f"{r['Puan']:.1f}", style={"color": color_p, "textAlign": "right",
                                                    "fontWeight": "600"}),
            ], style={"borderBottom": "1px solid rgba(255,255,255,0.05)"})
        )
    table_p = html.Table([
        html.Thead(html.Tr([
            html.Th("#",     style={"color": "#6b7280", "width": "36px", "textAlign": "center"}),
            html.Th("Nama",  style={"color": "#6b7280"}),
            html.Th("Puan",  style={"color": "#6b7280", "textAlign": "right"}),
        ])),
        html.Tbody(rows),
    ], style={"width": "100%", "fontSize": "0.82rem", "borderCollapse": "collapse"})

    rows2 = []
    for i, r in ds_top30.iterrows():
        eid = int(r["External ID"]) if pd.notna(r["External ID"]) else None
        in_both = eid in top30_both
        color_s = "#22c55e" if in_both else "#e2e8f0"
        rows2.append(
            html.Tr([
                html.Td(i + 1, style={"color": "#9ca3af", "width": "36px", "textAlign": "center"}),
                html.Td(r["İsim"], style={"color": color_s, "fontWeight": "600" if in_both else "normal"}),
                html.Td(f"{r['Puan']:.1f}", style={"color": color_s, "textAlign": "right",
                                                    "fontWeight": "600"}),
            ], style={"borderBottom": "1px solid rgba(255,255,255,0.05)"})
        )
    table_s = html.Table([
        html.Thead(html.Tr([
            html.Th("#",     style={"color": "#6b7280", "width": "36px", "textAlign": "center"}),
            html.Th("Nama",  style={"color": "#6b7280"}),
            html.Th("Puan",  style={"color": "#6b7280", "textAlign": "right"}),
        ])),
        html.Tbody(rows2),
    ], style={"width": "100%", "fontSize": "0.82rem", "borderCollapse": "collapse"})

    return table_p, table_s


# ── Analysis 2 : Attendance funnel ───────────────────────────────────────────
def build_funnel_chart():
    labels = ["Daftar Sahaja\n(Tidak Hadir)", "Ujian Percubaan\nSahaja",
              "Ujian Sebenar\nSahaja", "Kedua-dua Ujian"]
    values = [f_none, f_perc_only, f_seb_only, f_both]
    colors = ["#ef4444", "#f97316", "#3b82f6", "#22c55e"]
    fig = go.Figure(go.Bar(
        x=labels, y=values,
        marker_color=colors,
        text=[f"{v:,}<br>({v/len(dr_emails)*100:.1f}%)" for v in values],
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>%{y:,} peserta<extra></extra>",
    ))
    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        margin=dict(l=20, r=20, t=40, b=20),
        font=dict(color="#ccc"),
        showlegend=False,
        yaxis=dict(showgrid=False, showticklabels=False),
        xaxis=dict(tickfont=dict(color="#bbb", size=11)),
        uniformtext_minsize=10,
    )
    return fig


# ── Analysis 3 : Score scatter + histogram ───────────────────────────────────
def build_score_scatter():
    improved = score_merged[score_merged["Delta"] > 0]
    dropped  = score_merged[score_merged["Delta"] < 0]
    same_df  = score_merged[score_merged["Delta"] == 0]
    max_p = max(score_merged["Puan_P"].max(), score_merged["Puan_S"].max()) + 5

    fig = go.Figure()
    fig.add_shape(type="line", x0=0, y0=0, x1=max_p, y1=max_p,
                  line=dict(color="rgba(255,255,255,0.15)", dash="dot"))
    for subset, col, name in [
        (improved, "#22c55e", f"Naik ({len(improved)})"),
        (dropped,  "#ef4444", f"Turun ({len(dropped)})"),
        (same_df,  "#94a3b8", f"Sama ({len(same_df)})"),
    ]:
        fig.add_trace(go.Scatter(
            x=subset["Puan_P"], y=subset["Puan_S"],
            mode="markers",
            name=name,
            marker=dict(color=col, size=7, opacity=0.8,
                        line=dict(color="rgba(0,0,0,0.3)", width=0.5)),
            hovertemplate="<b>%{customdata}</b><br>Percubaan: %{x:.1f}<br>Sebenar: %{y:.1f}<extra></extra>",
            customdata=subset["Nama"],
        ))
    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(title="Puan Percubaan", color="#aaa",
                   gridcolor="rgba(255,255,255,0.05)", zeroline=False),
        yaxis=dict(title="Puan Sebenar", color="#aaa",
                   gridcolor="rgba(255,255,255,0.05)", zeroline=False),
        legend=dict(font=dict(color="#ccc", size=11), bgcolor="rgba(0,0,0,0)",
                    orientation="h", y=1.06),
        margin=dict(l=10, r=10, t=30, b=10),
        font=dict(color="#ccc"),
    )
    return fig


def build_score_histogram():
    fig = go.Figure(go.Histogram(
        x=score_merged["Delta"],
        nbinsx=20,
        marker_color="#6366f1",
        marker_line=dict(color="#0b0b16", width=0.5),
        hovertemplate="Delta: %{x}<br>Bilangan: %{y}<extra></extra>",
    ))
    fig.add_vline(x=0, line_color="rgba(255,255,255,0.4)", line_dash="dash")
    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(title="Perubahan Markah (Sebenar − Percubaan)", color="#aaa",
                   gridcolor="rgba(255,255,255,0.05)", zeroline=False),
        yaxis=dict(title="Bilangan Peserta", color="#aaa",
                   gridcolor="rgba(255,255,255,0.05)"),
        margin=dict(l=10, r=10, t=30, b=10),
        font=dict(color="#ccc"),
    )
    return fig


# ── Analysis 4 : Registration date vs Sebenar score ──────────────────────────
def build_reg_score_scatter():
    import numpy as np
    df_s = reg_score_df.copy()
    # 14-day rolling mean on date-aggregated averages
    daily_mean = (df_s.groupby("Kayıt Tarihi")["Puan"]
                  .mean().reset_index().sort_values("Kayıt Tarihi"))
    daily_mean["Rolling"] = daily_mean["Puan"].rolling(7, min_periods=1).mean()

    fig = go.Figure()
    # Scatter: individual dots coloured by score band
    for lo, hi, col, name in [
        (80, 101, "#22c55e", "≥ 80"),
        (60,  80, "#f59e0b", "60–79"),
        (40,  60, "#f97316", "40–59"),
        ( 0,  40, "#ef4444", "< 40"),
    ]:
        sub = df_s[(df_s["Puan"] >= lo) & (df_s["Puan"] < hi)]
        fig.add_trace(go.Scatter(
            x=sub["Kayıt Tarihi"], y=sub["Puan"],
            mode="markers", name=name,
            marker=dict(color=col, size=4, opacity=0.45,
                        line=dict(width=0)),
            hovertemplate="<b>%{x|%d %b %Y}</b><br>Puan: %{y}<extra></extra>",
        ))
    # Rolling mean line
    fig.add_trace(go.Scatter(
        x=daily_mean["Kayıt Tarihi"], y=daily_mean["Rolling"],
        mode="lines", name="Rolling avg (7-hari)",
        line=dict(color="#ffffff", width=2, dash="solid"),
        hovertemplate="<b>%{x|%d %b %Y}</b><br>Avg: %{y:.1f}<extra></extra>",
    ))
    # Milestone vertical lines
    for m in milestones:
        fig.add_vline(x=pd.to_datetime(m["date"]).timestamp() * 1000,
                      line_width=1, line_dash="dot",
                      line_color="rgba(255,255,100,0.35)",
                      annotation_text=m["emoji"],
                      annotation_position="top",
                      annotation_font=dict(size=11))

    # ── Pass-mark reference ──────────────────────────────────────────────
    fig.add_hrect(y0=-5, y1=40,
                  fillcolor="rgba(239,68,68,0.05)",
                  line_width=0, layer="below")
    fig.add_hline(y=40,
                  line_width=1.5, line_dash="dash",
                  line_color="rgba(239,68,68,0.55)",
                  annotation_text="  Had Lulus (40)",
                  annotation_position="right",
                  annotation_font=dict(color="#ef4444", size=10))

    # ── Reading-guide annotation (top-left) ─────────────────────────────
    fig.add_annotation(
        x=0.01, y=0.97, xref="paper", yref="paper",
        text="💡 Daftar lebih awal = titik di sebelah kiri<br>"
             "   Markah tinggi = titik di bahagian atas<br>"
             "   Garisan putih = purata bergerak (7 hari)<br>"
             "   Kawasan merah = bawah had lulus",
        showarrow=False,
        align="left",
        bgcolor="rgba(0,0,0,0.55)",
        bordercolor="rgba(255,255,255,0.08)",
        borderwidth=1,
        borderpad=8,
        font=dict(color="#9ca3af", size=10),
    )

    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(title="Tarikh Daftar", color="#aaa",
                   gridcolor="rgba(255,255,255,0.04)", zeroline=False),
        yaxis=dict(title="Puan Sebenar", color="#aaa", range=[-5, 105],
                   gridcolor="rgba(255,255,255,0.04)", zeroline=False),
        legend=dict(font=dict(color="#ccc", size=10), bgcolor="rgba(0,0,0,0)",
                    orientation="h", y=1.05),
        margin=dict(l=10, r=80, t=40, b=10),
        font=dict(color="#ccc"),
    )
    return fig


def build_reg_period_bar():
    df_p = reg_period_df.copy()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Min Puan",
        x=df_p["Period"],
        y=df_p["mean"],
        marker_color="#6366f1",
        text=[f"{v:.1f}" for v in df_p["mean"]],
        textposition="outside",
        yaxis="y",
        hovertemplate="<b>%{x}</b><br>Min Puan: %{y:.1f}<br>n=%{customdata}<extra></extra>",
        customdata=df_p["n"],
    ))
    fig.add_trace(go.Scatter(
        name="Kadar Lulus (%)",
        x=df_p["Period"],
        y=df_p["pass_pct"],
        mode="lines+markers",
        line=dict(color="#22c55e", width=2),
        marker=dict(size=8, color="#22c55e"),
        yaxis="y2",
        hovertemplate="<b>%{x}</b><br>Lulus: %{y:.1f}%<extra></extra>",
    ))

    # ── Overall avg pass rate reference line ────────────────────────────
    overall_pass = df_p["pass_pct"].mean()
    fig.add_hline(y=overall_pass,
                  line_width=1, line_dash="dot",
                  line_color="rgba(34,197,94,0.4)",
                  yref="y2",
                  annotation_text=f"  Purata {overall_pass:.1f}%",
                  annotation_position="top right",
                  annotation_font=dict(color="#22c55e", size=9))

    # ── Reading-guide annotation ─────────────────────────────────────────
    fig.add_annotation(
        x=0.01, y=0.97, xref="paper", yref="paper",
        text="💡 Bar ungu = min markah (skala kiri)<br>"
             "   Garisan hijau = kadar lulus % (skala kanan)<br>"
             "   Garisan bertitik = purata kadar lulus keseluruhan<br>"
             "   Skala kanan bermula dari 75%, bukan 0%",
        showarrow=False,
        align="left",
        bgcolor="rgba(0,0,0,0.55)",
        bordercolor="rgba(255,255,255,0.08)",
        borderwidth=1,
        borderpad=8,
        font=dict(color="#9ca3af", size=10),
    )

    fig.update_layout(
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        barmode="group",
        xaxis=dict(color="#aaa", tickfont=dict(size=10),
                   gridcolor="rgba(255,255,255,0.04)"),
        yaxis=dict(title="Min Puan", color="#6366f1", range=[0, 110],
                   gridcolor="rgba(255,255,255,0.04)", zeroline=False),
        yaxis2=dict(title="Kadar Lulus (%)", color="#22c55e",
                    overlaying="y", side="right", range=[75, 105],
                    showgrid=False, zeroline=False),
        legend=dict(font=dict(color="#ccc", size=10), bgcolor="rgba(0,0,0,0)",
                    orientation="h", y=1.05),
        margin=dict(l=10, r=60, t=40, b=10),
        font=dict(color="#ccc"),
    )
    return fig


# ── Analysis 5 : Bell curve ───────────────────────────────────────────────────
def build_bell_curve():
    import numpy as np
    fig = go.Figure()
    for data, col, name, pass_n, total_n in [
        (dp_full["Puan"], "rgba(249,115,22,0.7)", "Percubaan", dp_pass_n, len(dp_full)),
        (ds_full["Puan"], "rgba(59,130,246,0.7)", "Sebenar",   ds_pass_n, len(ds_full)),
    ]:
        fig.add_trace(go.Histogram(
            x=data, nbinsx=20, name=name,
            marker_color=col,
            marker_line=dict(color="#0b0b16", width=0.5),
            opacity=0.85,
            hovertemplate="Markah: %{x}<br>Bilangan: %{y}<extra></extra>",
        ))
    # Pass/fail line
    fig.add_vline(x=PASS_MARK, line_color="#fbbf24", line_width=2, line_dash="dash",
                  annotation_text="Lulus (40%)", annotation_font_color="#fbbf24",
                  annotation_position="top right")
    fig.update_layout(
        barmode="overlay",
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(title="Markah (Puan)", color="#aaa",
                   gridcolor="rgba(255,255,255,0.05)", zeroline=False,
                   range=[0, 105]),
        yaxis=dict(title="Bilangan Peserta", color="#aaa",
                   gridcolor="rgba(255,255,255,0.05)"),
        legend=dict(font=dict(color="#ccc"), bgcolor="rgba(0,0,0,0)",
                    orientation="h", y=1.06),
        margin=dict(l=10, r=10, t=30, b=10),
        font=dict(color="#ccc"),
    )
    return fig


# ── Analysis 6 : Anomaly ──────────────────────────────────────────────────────
def build_anomaly_table():
    tbl = anomaly_s[["İsim", "E-posta", "Bitiş", "Doğru", "Yanlış", "Boş"]].head(20)
    head_style = {"color": "#6b7280", "padding": "6px 8px",
                  "borderBottom": "1px solid rgba(255,255,255,0.12)",
                  "background": CARD_BG2, "fontSize": "0.78rem"}
    cell_style = {"color": "#d1d5db", "padding": "5px 8px",
                  "borderBottom": "1px solid rgba(255,255,255,0.05)",
                  "fontSize": "0.78rem"}
    header = html.Thead(html.Tr([
        html.Th("Nama",    style=head_style),
        html.Th("E-mel",   style=head_style),
        html.Th("Bitiş",   style=head_style),
        html.Th("Betul",   style={**head_style, "textAlign": "center"}),
        html.Th("Salah",   style={**head_style, "textAlign": "center"}),
        html.Th("Kosong",  style={**head_style, "textAlign": "center"}),
    ]))
    rows = [
        html.Tr([
            html.Td(r["İsim"],   style=cell_style),
            html.Td(str(r["E-posta"]),  style={**cell_style, "fontSize": "0.72rem",
                                               "color": "#6b7280"}),
            html.Td(str(r["Bitiş"]),    style={**cell_style, "color": "#fbbf24"}),
            html.Td(int(r["Doğru"]),    style={**cell_style, "textAlign": "center",
                                               "color": "#22c55e"}),
            html.Td(int(r["Yanlış"]),   style={**cell_style, "textAlign": "center",
                                               "color": "#ef4444"}),
            html.Td(int(r["Boş"]),      style={**cell_style, "textAlign": "center"}),
        ])
        for _, r in tbl.iterrows()
    ]
    return html.Table([header, html.Tbody(rows)],
                      style={"width": "100%", "borderCollapse": "collapse"})


# ── Appendix table builders ──────────────────────────────────────────────────
def _reg_table(frame, group_col, show_cols, col_labels, highlight_col=None):
    """Generic grouped registration anomaly table."""
    HEAD = {"color": "#6b7280", "padding": "6px 10px",
            "borderBottom": "1px solid rgba(255,255,255,0.12)",
            "background": CARD_BG2, "fontSize": "0.77rem", "whiteSpace": "nowrap"}
    CELL = {"color": "#d1d5db", "padding": "5px 10px",
            "borderBottom": "1px solid rgba(255,255,255,0.04)",
            "fontSize": "0.77rem", "verticalAlign": "top"}
    rows = []
    prev_group = None
    for _, r in frame.iterrows():
        group_val = r[group_col]
        is_new_group = group_val != prev_group
        prev_group = group_val
        cells = []
        for col in show_cols:
            val = str(r[col]) if not pd.isna(r[col]) else ""
            style = dict(CELL)
            if col == highlight_col:
                style["color"] = "#fbbf24"
                style["fontSize"] = "0.72rem"
            if col == "Status Sebenar":
                if "Lengkap" in val:
                    style["color"] = "#22c55e"
                elif "Hadir" in val:
                    style["color"] = "#f97316"
                elif "Tidak" in val:
                    style["color"] = "#ef4444"
                style["fontSize"] = "0.72rem"
                style["whiteSpace"] = "nowrap"
                style["fontSize"] = "0.72rem"
            if is_new_group and col == show_cols[0]:
                style["borderTop"] = "1px solid rgba(255,255,255,0.10)"
                style["paddingTop"] = "10px"
            cells.append(html.Td(val, style=style))
        rows.append(html.Tr(cells))
    header = html.Thead(html.Tr([
        html.Th(label, style=HEAD) for label in col_labels
    ]))
    return html.Table([header, html.Tbody(rows)],
                      style={"width": "100%", "borderCollapse": "collapse"})


def build_type_a_table(n=300):
    """Same email, different names."""
    tbl = email_multi.head(n)
    return _reg_table(
        tbl,
        group_col="_email_key",
        show_cols=["Ad", "E-posta", "Kayıt Tarihi", "Status Sebenar"],
        col_labels=["Nama", "E-mel (dikongsi)", "Tarikh Daftar", "Ujian Sebenar"],
        highlight_col="E-posta",
    )


def build_type_b_table(n=300):
    """Same name, different emails."""
    tbl = name_multi.head(n)
    return _reg_table(
        tbl,
        group_col="_name_key",
        show_cols=["Ad", "E-posta", "Kayıt Tarihi", "Status Sebenar"],
        col_labels=["Nama (sama)", "E-mel", "Tarikh Daftar", "Ujian Sebenar"],
        highlight_col="Ad",
    )


# ── Chart: Anomaly by negeri ──────────────────────────────────────────────────
def build_anomaly_negeri_chart():
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=_anm_negeri["Negeri"],
        x=_anm_negeri["Count_A"],
        name="Jenis A (E-mel Dikongsi)",
        orientation="h",
        marker_color="rgba(245,158,11,0.82)",
        hovertemplate="<b>%{y}</b><br>Jenis A: <b>%{x:,}</b><extra></extra>",
    ))
    fig.add_trace(go.Bar(
        y=_anm_negeri["Negeri"],
        x=_anm_negeri["Count_B"],
        name="Jenis B (Nama Sama)",
        orientation="h",
        marker_color="rgba(167,139,250,0.82)",
        hovertemplate="<b>%{y}</b><br>Jenis B: <b>%{x:,}</b><extra></extra>",
    ))
    bar_height = max(380, len(_anm_negeri) * 30)
    fig.update_layout(
        barmode="stack",
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        legend=dict(orientation="h", y=1.06, x=0,
                    font=dict(color="#aaa", size=12), bgcolor="rgba(0,0,0,0)"),
        margin=dict(l=10, r=70, t=30, b=20),
        xaxis=dict(showgrid=False, showticklabels=False, showline=False),
        yaxis=dict(tickfont=dict(color="#ccc", size=12), showgrid=False),
        bargap=0.25,
        height=bar_height,
    )
    # Add total labels at end of each bar
    for _, row in _anm_negeri.iterrows():
        fig.add_annotation(
            x=row["Total"], y=row["Negeri"],
            text=f"<b>{row['Total']:,}</b>",
            showarrow=False, xanchor="left",
            font=dict(color="#999", size=11), xshift=5,
        )
    return fig


# ── Conclusion box helper ────────────────────────────────────────────────────
def _conclusion(text, accent="#4E9AF1"):
    return html.Div(
        [html.Span("📌 ", style={"fontSize": "0.9rem"}),
         html.Span("Kesimpulan: ", style={"color": accent, "fontWeight": "700",
                                          "fontSize": "0.85rem"}),
         html.Span(text, style={"color": "#9ca3af"})],
        style={
            "background": CARD_BG2,
            "borderLeft": f"3px solid {accent}",
            "borderRadius": "0 8px 8px 0",
            "padding": "12px 20px",
            "fontSize": "0.85rem",
            "lineHeight": "1.7",
            "marginBottom": "44px",
        },
    )


# ── App ───────────────────────────────────────────────────────────────────────
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])
server = app.server  # required for gunicorn / Render
app.title = "PMBA Dashboard"

lelaki_n   = int(gender_counts.get("Lelaki", 0))
perempuan_n = int(gender_counts.get("Perempuan", 0))
unknown_n  = int(gender_counts.get("Tidak Diketahui", 0))
_best_reg  = reg_period_df.sort_values("mean", ascending=False).iloc[0]
_worst_reg = reg_period_df.sort_values("mean").iloc[0]

app.layout = html.Div(
    style={"backgroundColor": BG, "minHeight": "100vh",
           "fontFamily": "'Inter', 'Segoe UI', sans-serif"},
    children=[

        # ── Top nav bar ───────────────────────────────────────────────────
        html.Div(
            style={
                "backgroundColor": CARD_BG, "borderBottom": f"1px solid {BORDER}",
                "padding": "14px 48px", "display": "flex",
                "justifyContent": "space-between", "alignItems": "center",
            },
            children=[
                html.Div([
                    html.Span("📊 ", style={"fontSize": "1.3rem"}),
                    html.Span("PMBA", style={"color": "white", "fontWeight": "800",
                                              "fontSize": "1.15rem", "letterSpacing": "-0.01em"}),
                    html.Span(" · Analisis Pendaftaran Pertandingan",
                              style={"color": "#555", "fontSize": "0.88rem", "marginLeft": "8px"}),
                ]),
                html.Div("19 Okt 2025 – 27 Mac 2026",
                         style={"color": "#444", "fontSize": "0.78rem", "letterSpacing": "0.04em"}),
            ],
        ),

        # ── Page body ─────────────────────────────────────────────────────
        html.Div(
            style={"padding": "36px 48px", "maxWidth": "1600px", "margin": "0 auto"},
            children=[

                # ── KPI row ───────────────────────────────────────────────
                html.P("Ringkasan", style=SECTION_LABEL_STYLE),
                dbc.Row([
                    dbc.Col(kpi_card("Jumlah Peserta",    f"{total:,}",
                                     "Keseluruhan pendaftar", "#4E9AF1"), md=3),
                    dbc.Col(kpi_card("Puncak Harian",      f"{peak_val}",
                                     f"Pada {peak_day}",      "#F4A261"), md=3),
                    dbc.Col(kpi_card("Purata Harian",      f"{avg_daily}",
                                     "Pendaftar sehari",       "#2A9D8F"), md=3),
                    dbc.Col(kpi_card("Tempoh Pendaftaran", "161 hari",
                                     "Tempoh terbuka",         "#9b8fe8"), md=3),
                ], className="g-3 mb-5"),

                # ── Trend chart ───────────────────────────────────────────
                html.P("Trend Pendaftaran & Peristiwa Penting", style=SECTION_LABEL_STYLE),
                html.Div(
                    style={"backgroundColor": CARD_BG, "borderRadius": "14px",
                           "border": f"1px solid {BORDER}", "padding": "6px",
                           "marginBottom": "44px"},
                    children=[dcc.Graph(
                        figure=build_main_chart(),
                        config={"displayModeBar": "hover",
                                "modeBarButtonsToRemove": ["select2d", "lasso2d"],
                                "toImageButtonOptions": {"format": "png", "scale": 2}},
                        style={"height": "540px"},
                    )],
                ),

                _conclusion(
                    f"Pendaftaran kumulatif mencapai {total:,} peserta dalam {_n_days} hari. "
                    f"Separuh daripada pendaftaran berlaku dalam tempoh 20 hari terakhir "
                    f"(Mac 2026) sahaja. Lonjakan harian tertinggi ialah {peak_val} peserta "
                    f"pada {peak_day}.",
                    accent="#4E9AF1",
                ),

                # ── Milestone table ───────────────────────────────────────
                html.P("Garis Masa Peristiwa Penting", style=SECTION_LABEL_STYLE),
                html.Div(
                    style={"backgroundColor": CARD_BG, "borderRadius": "14px",
                           "border": f"1px solid {BORDER}", "overflow": "hidden",
                           "marginBottom": "44px"},
                    children=[
                        html.Div(
                            "Jumlah pendaftar terkumpul semasa setiap peristiwa berlaku",
                            style={"padding": "16px 24px",
                                   "borderBottom": f"1px solid {BORDER}",
                                   "color": "#666", "fontSize": "0.82rem"},
                        ),
                        html.Table(
                            style={"width": "100%", "borderCollapse": "collapse"},
                            children=[
                                html.Thead(html.Tr([
                                    html.Th(h, style={
                                        "color": "#555", "fontSize": "0.68rem",
                                        "letterSpacing": "0.1em", "textTransform": "uppercase",
                                        "padding": "10px 24px", "fontWeight": "600",
                                        "borderBottom": f"1px solid {BORDER}",
                                        "textAlign": "right" if h == "Kumulatif Ketika Itu" else "left",
                                        "width": w,
                                    })
                                    for h, w in [
                                        ("Tarikh", "120px"), ("Peristiwa", "auto"),
                                        ("Kumulatif Ketika Itu", "180px"),
                                    ]
                                ])),
                                html.Tbody(milestone_table_rows()),
                            ],
                        ),
                    ],
                ),

                _conclusion(
                    "Pengiktirafan PAJSK (27 Feb 2026) menjadi titik perubahan paling kritikal — "
                    "ia mengesahkan pertandingan ini sebagai aktiviti yang diiktiraf secara rasmi "
                    "dan sah di bawah sistem pendidikan, mencetuskan gelombang pendaftaran hampir "
                    "2× ganda. Ujian Percubaan yang dibuka kepada awam 2 hari sebelum Ujian Sebenar "
                    "turut membantu peserta membuat persediaan terakhir dan mendorong pendaftaran "
                    "saat-saat akhir.",
                    accent="#9b8fe8",
                ),

                # ── Negeri analysis ───────────────────────────────────────
                html.P("Analisis Mengikut Negeri", style=SECTION_LABEL_STYLE),
                dbc.Row([
                    # Horizontal bar
                    dbc.Col(
                        html.Div(
                            style={"backgroundColor": CARD_BG, "borderRadius": "14px",
                                   "border": f"1px solid {BORDER}", "padding": "20px 16px 12px 16px"},
                            children=[
                                html.P("Bilangan Pendaftar Mengikut Negeri",
                                       style={"color": "#666", "fontSize": "0.82rem",
                                              "marginBottom": "4px", "paddingLeft": "8px"}),
                                dcc.Graph(
                                    figure=build_negeri_bar(),
                                    config={"displayModeBar": False},
                                    style={"height": "480px"},
                                ),
                            ],
                        ),
                        md=5,
                    ),
                    # Malaysia bubble map
                    dbc.Col(
                        html.Div(
                            style={"backgroundColor": CARD_BG, "borderRadius": "14px",
                                   "border": f"1px solid {BORDER}", "padding": "20px 16px 12px 16px"},
                            children=[
                                html.P("Peta Malaysia — Saiz Bulatan Mengikut Pendaftar",
                                       style={"color": "#666", "fontSize": "0.82rem",
                                              "marginBottom": "4px", "paddingLeft": "8px"}),
                                dcc.Graph(
                                    figure=build_malaysia_map(),
                                    config={"displayModeBar": False},
                                    style={"height": "480px"},
                                ),
                            ],
                        ),
                        md=7,
                    ),
                ], className="g-3 mb-5"),

                # ── Negeri timeline breakdown ─────────────────────────────
                html.P("Garis Masa Pendaftaran Mengikut Negeri", style=SECTION_LABEL_STYLE),

                # Stacked area chart
                html.Div(
                    style={
                        "backgroundColor": CARD_BG, "borderRadius": "14px",
                        "border": f"1px solid {BORDER}", "padding": "20px 16px 12px 16px",
                        "marginBottom": "16px",
                    },
                    children=[
                        html.P("Kumulatif Pendaftaran — Pecahan Mengikut Negeri (Top 8 + Lain-lain)",
                               style={"color": "#666", "fontSize": "0.82rem",
                                      "marginBottom": "4px", "paddingLeft": "8px"}),
                        dcc.Graph(
                            figure=build_negeri_stacked_area(),
                            config={"displayModeBar": "hover",
                                    "modeBarButtonsToRemove": ["select2d", "lasso2d"],
                                    "toImageButtonOptions": {"format": "png", "scale": 2}},
                            style={"height": "460px"},
                        ),
                    ],
                ),

                # Weekly heatmap
                html.Div(
                    style={
                        "backgroundColor": CARD_BG, "borderRadius": "14px",
                        "border": f"1px solid {BORDER}", "padding": "20px 16px 16px 16px",
                        "marginBottom": "44px",
                    },
                    children=[
                        html.P("Bilangan Pendaftaran Mingguan Mengikut Negeri (Heatmap)",
                               style={"color": "#666", "fontSize": "0.82rem",
                                      "marginBottom": "4px", "paddingLeft": "8px"}),
                        html.P("Garisan kuning bertitik = peristiwa penting",
                               style={"color": "#444", "fontSize": "0.72rem",
                                      "marginBottom": "0", "paddingLeft": "8px"}),
                        dcc.Graph(
                            figure=build_negeri_heatmap(),
                            config={"displayModeBar": False},
                            style={"height": "400px"},
                        ),
                    ],
                ),

                # ── Gender analysis ───────────────────────────────────────
                html.P("Analisis Mengikut Jantina", style=SECTION_LABEL_STYLE),
                dbc.Row([
                    # Stat cards
                    dbc.Col(
                        html.Div(
                            style={"paddingTop": "30px"},
                            children=[
                                gender_stat("Lelaki",   lelaki_n,
                                            100 * lelaki_n / total,    "#4E9AF1"),
                                gender_stat("Perempuan", perempuan_n,
                                            100 * perempuan_n / total, "#F472B6"),
                                gender_stat("Tidak Diketahui", unknown_n,
                                            100 * unknown_n / total,  "#555"),
                                html.P(
                                    "* Dikesan berdasarkan 'bin' (lelaki) dan 'binti/bt' (perempuan) dalam nama.",
                                    style={"color": "#444", "fontSize": "0.72rem",
                                           "marginTop": "16px", "lineHeight": "1.5"},
                                ),
                            ],
                        ),
                        md=3,
                    ),
                    # Gender donut
                    dbc.Col(
                        html.Div(
                            style={"backgroundColor": CARD_BG, "borderRadius": "14px",
                                   "border": f"1px solid {BORDER}", "padding": "20px 16px"},
                            children=[
                                html.P("Pecahan Jantina",
                                       style={"color": "#666", "fontSize": "0.82rem",
                                              "marginBottom": "4px", "paddingLeft": "8px"}),
                                dcc.Graph(
                                    figure=build_gender_donut(),
                                    config={"displayModeBar": False},
                                    style={"height": "380px"},
                                ),
                            ],
                        ),
                        md=4,
                    ),
                    # Gender × Negeri top-10 stacked bar
                    dbc.Col(
                        html.Div(
                            style={"backgroundColor": CARD_BG, "borderRadius": "14px",
                                   "border": f"1px solid {BORDER}", "padding": "20px 16px"},
                            children=[
                                html.P("Jantina Mengikut Negeri (Top 10)",
                                       style={"color": "#666", "fontSize": "0.82rem",
                                              "marginBottom": "4px", "paddingLeft": "8px"}),
                                dcc.Graph(
                                    figure=_build_gender_negeri_chart(),
                                    config={"displayModeBar": False},
                                    style={"height": "380px"},
                                ),
                            ],
                        ),
                        md=5,
                    ),
                ], className="g-3 mb-5"),

                _conclusion(
                    f"Peserta perempuan ({round(100*perempuan_n/total,1)}%) mendominasi berbanding "
                    f"lelaki ({round(100*lelaki_n/total,1)}%). Corak ini konsisten merentasi hampir "
                    f"semua negeri. Walau bagaimanapun, {unknown_n:,} peserta "
                    f"({round(100*unknown_n/total,1)}%) tidak dapat ditentukan jantina berdasarkan "
                    f"nama — ini memberi had kepada ketepatan analisis jantina.",
                    accent="#F472B6",
                ),

                # ══════════════════════════════════════════════════════════
                # BAHAGIAN 2 : Analisis Ujian Percubaan & Sebenar
                # ══════════════════════════════════════════════════════════
                html.Hr(style={"borderColor": "rgba(255,255,255,0.06)", "margin": "8px 0 28px"}),

                html.P("ANALISIS UJIAN PERCUBAAN & SEBENAR",
                       style={"color": "#6b7280", "fontSize": "0.7rem", "letterSpacing": "0.12em",
                              "fontWeight": "700", "marginBottom": "22px",
                              "borderLeft": "3px solid #6366f1", "paddingLeft": "10px"}),

                # ── Analysis 2 : Attendance Funnel ────────────────────────
                html.P("Kehadiran Peserta ke Ujian",
                       style={"color": "#e2e8f0", "fontWeight": "700", "fontSize": "1.05rem",
                              "marginBottom": "6px"}),
                html.P(
                    f"Daripada {len(dr_emails):,} peserta yang daftar, "
                    f"{f_both + f_seb_only:,} ({(f_both + f_seb_only)/len(dr_emails)*100:.1f}%) "
                    f"berjaya menamatkan Ujian Sebenar (40 soalan). "
                    f"Hanya {f_both:,} peserta hadir kedua-dua ujian dan siap 40 soalan.",
                    style={"color": "#9ca3af", "fontSize": "0.9rem", "marginBottom": "16px"},
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                dcc.Graph(
                                    figure=build_funnel_chart(),
                                    config={"displayModeBar": False},
                                    style={"height": "320px"},
                                ),
                            ], style={"padding": "12px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=8),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P("Ringkasan Kehadiran", style={"color": "#9ca3af",
                                       "fontSize": "0.78rem", "marginBottom": "14px",
                                       "fontWeight": "700"}),
                                *[
                                    html.Div([
                                        html.Span(label, style={"color": "#9ca3af",
                                                                 "fontSize": "0.82rem"}),
                                        html.Span(f"{val:,}", style={"color": col,
                                                                      "fontWeight": "700",
                                                                      "fontSize": "1.1rem",
                                                                      "float": "right"}),
                                    ], style={"marginBottom": "14px", "overflow": "hidden"})
                                    for label, val, col in [
                                        ("Daftar sahaja (tiada ujian)", f_none, "#ef4444"),
                                        ("Percubaan sahaja", f_perc_only, "#f97316"),
                                        ("Sebenar sahaja", f_seb_only, "#3b82f6"),
                                        ("Kedua-dua ujian", f_both, "#22c55e"),
                                    ]
                                ],
                            ], style={"padding": "18px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px", "height": "100%"}),
                    ], md=4),
                ], className="g-3 mb-5"),

                _conclusion(
                    f"Kadar kehadiran ujian secara keseluruhan adalah "
                    f"{round((f_both+f_perc_only+f_seb_only)/len(dr_emails)*100,1)}% "
                    f"({f_both+f_perc_only+f_seb_only:,} daripada {len(dr_emails):,} pendaftar). "
                    f"Seramai {f_none:,} ({round(f_none/len(dr_emails)*100,1)}%) sama sekali tidak "
                    f"menduduki mana-mana ujian. Ini mencadangkan keperluan sistem peringatan yang "
                    f"lebih aktif antara fasa pendaftaran dan fasa ujian.",
                    accent="#3b82f6",
                ),

                # ── Analysis 1 : Top-30 Consistency ───────────────────────
                html.P("Top 30 Konsisten: Percubaan → Sebenar",
                       style={"color": "#e2e8f0", "fontWeight": "700", "fontSize": "1.05rem",
                              "marginBottom": "6px"}),
                html.P(
                    f"Hanya {len(top30_both)} peserta kekal dalam kedudukan Top 30 di kedua-dua ujian. "
                    f"Peserta yang kekal diserlahkan dalam warna hijau.",
                    style={"color": "#9ca3af", "fontSize": "0.9rem", "marginBottom": "16px"},
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P("Top 30 — Ujian Percubaan",
                                       style={"color": "#f97316", "fontSize": "0.85rem",
                                              "fontWeight": "700", "marginBottom": "10px"}),
                                html.Div(build_top30_table()[0],
                                         style={"maxHeight": "520px", "overflowY": "auto"}),
                            ], style={"padding": "16px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=6),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P("Top 30 — Ujian Sebenar",
                                       style={"color": "#3b82f6", "fontSize": "0.85rem",
                                              "fontWeight": "700", "marginBottom": "10px"}),
                                html.Div(build_top30_table()[1],
                                         style={"maxHeight": "520px", "overflowY": "auto"}),
                            ], style={"padding": "16px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=6),
                ], className="g-3 mb-5"),

                # ── Analysis 3 : Score Improvement ────────────────────────
                html.P("Perubahan Markah: Percubaan vs Sebenar",
                       style={"color": "#e2e8f0", "fontWeight": "700", "fontSize": "1.05rem",
                              "marginBottom": "6px"}),
                html.P(
                    f"Daripada {len(score_merged)} peserta yang hadir kedua-dua ujian, "
                    f"{(score_merged['Delta']>0).sum()} ({round((score_merged['Delta']>0).mean()*100,1)}%) "
                    f"meningkat markah dalam Ujian Sebenar.",
                    style={"color": "#9ca3af", "fontSize": "0.9rem", "marginBottom": "16px"},
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P("Scatter: Puan Percubaan vs Sebenar",
                                       style={"color": "#6b7280", "fontSize": "0.78rem",
                                              "marginBottom": "6px"}),
                                dcc.Graph(
                                    figure=build_score_scatter(),
                                    config={"displayModeBar": False},
                                    style={"height": "360px"},
                                ),
                            ], style={"padding": "12px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=7),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P("Taburan Perubahan Markah (Delta)",
                                       style={"color": "#6b7280", "fontSize": "0.78rem",
                                              "marginBottom": "6px"}),
                                dcc.Graph(
                                    figure=build_score_histogram(),
                                    config={"displayModeBar": False},
                                    style={"height": "360px"},
                                ),
                            ], style={"padding": "12px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=5),
                ], className="g-3 mb-5"),

                _conclusion(
                    f"Sebanyak {round((score_merged['Delta']>0).mean()*100,1)}% peserta meningkat "
                    f"markah dalam Ujian Sebenar berbanding Percubaan, dengan min delta "
                    f"+{round(score_merged['Delta'].mean(),1)} mata. Ini mengesahkan bahawa Ujian "
                    f"Percubaan berfungsi sebagai latihan berkesan — peserta lebih bersedia untuk "
                    f"ujian muktamad.",
                    accent="#22c55e",
                ),

                # ── Analysis 4 : Registration Date vs Score ────────────────
                html.P("Tarikh Daftar vs Prestasi Ujian Sebenar",
                       style={"color": "#e2e8f0", "fontWeight": "700", "fontSize": "1.05rem",
                              "marginBottom": "6px"}),
                html.P(
                    f"Analisis {len(reg_score_df):,} peserta yang hadir Ujian Sebenar — "
                    "adakah mereka yang daftar awal mendapat markah lebih tinggi?",
                    style={"color": "#9ca3af", "fontSize": "0.9rem", "marginBottom": "16px"},
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P("Setiap titik = 1 peserta | Garisan putih = purata bergerak 7-hari "
                                       "| Garisan kuning bertitik = peristiwa penting",
                                       style={"color": "#6b7280", "fontSize": "0.78rem",
                                              "marginBottom": "6px"}),
                                dcc.Graph(
                                    figure=build_reg_score_scatter(),
                                    config={"displayModeBar": False},
                                    style={"height": "380px"},
                                ),
                            ], style={"padding": "12px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=8),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P("Min Puan & Kadar Lulus mengikut Tempoh Pendaftaran",
                                       style={"color": "#6b7280", "fontSize": "0.78rem",
                                              "marginBottom": "6px"}),
                                dcc.Graph(
                                    figure=build_reg_period_bar(),
                                    config={"displayModeBar": False},
                                    style={"height": "380px"},
                                ),
                            ], style={"padding": "12px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=4),
                ], className="g-3 mb-5"),

                _conclusion(
                    f"Peserta dari tempoh '{_best_reg['Period']}' mendapat min markah tertinggi "
                    f"({_best_reg['mean']:.1f}) dengan kadar lulus {_best_reg['pass_pct']:.1f}%. "
                    f"Sebaliknya, tempoh '{_worst_reg['Period']}' mencatat min markah paling rendah "
                    f"({_worst_reg['mean']:.1f}, lulus {_worst_reg['pass_pct']:.1f}%). "
                    f"Tempoh persediaan yang lebih panjang berkorelasi dengan prestasi lebih baik.",
                    accent="#6366f1",
                ),

                # ── Analysis 5 : Bell Curve ────────────────────────────────
                html.P("Taburan Markah — Bell Curve",
                       style={"color": "#e2e8f0", "fontWeight": "700", "fontSize": "1.05rem",
                              "marginBottom": "6px"}),
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                dcc.Graph(
                                    figure=build_bell_curve(),
                                    config={"displayModeBar": False},
                                    style={"height": "360px"},
                                ),
                            ], style={"padding": "12px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=8),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P("Ringkasan Lulus/Gagal", style={"color": "#9ca3af",
                                       "fontSize": "0.78rem", "marginBottom": "14px",
                                       "fontWeight": "700"}),
                                *[
                                    html.Div([
                                        html.P(exam, style={"color": exam_col,
                                               "fontWeight": "700", "fontSize": "0.85rem",
                                               "marginBottom": "6px"}),
                                        html.Div([
                                            html.Span("Lulus ", style={"color": "#9ca3af",
                                                                        "fontSize": "0.82rem"}),
                                            html.Span(f"{pn:,} ({pp}%)",
                                                      style={"color": "#22c55e",
                                                             "fontWeight": "700",
                                                             "fontSize": "1rem"}),
                                        ], style={"marginBottom": "4px"}),
                                        html.Div([
                                            html.Span("Gagal ", style={"color": "#9ca3af",
                                                                        "fontSize": "0.82rem"}),
                                            html.Span(f"{fn:,} ({round(fn/(pn+fn)*100,1)}%)",
                                                      style={"color": "#ef4444",
                                                             "fontWeight": "700",
                                                             "fontSize": "1rem"}),
                                        ], style={"marginBottom": "16px"}),
                                    ])
                                    for exam, exam_col, pn, fn, pp in [
                                        ("Percubaan", "#f97316", dp_pass_n, dp_fail_n, dp_pass_pct),
                                        ("Sebenar",   "#3b82f6", ds_pass_n, ds_fail_n, ds_pass_pct),
                                    ]
                                ],
                            ], style={"padding": "18px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px", "height": "100%"}),
                    ], md=4),
                ], className="g-3 mb-5"),

                _conclusion(
                    f"Ujian Sebenar mencatat kadar lulus {ds_pass_pct}% ({ds_pass_n:,} peserta) — "
                    f"lebih tinggi berbanding Ujian Percubaan ({dp_pass_pct}%, {dp_pass_n:,} peserta). "
                    f"Taburan markah kedua-dua ujian menghampiri bentuk loceng normal. "
                    f"Peningkatan kadar lulus ini mengesahkan keberkesanan Ujian Percubaan sebagai persediaan.",
                    accent="#3b82f6",
                ),

                # ── Analysis 6 : Anomaly ───────────────────────────────────
                html.P("Anomali: Peserta Technikal Bermasalah",
                       style={"color": "#e2e8f0", "fontWeight": "700", "fontSize": "1.05rem",
                              "marginBottom": "6px"}),
                html.P(
                    f"{len(anomaly_s):,} peserta dalam Ujian Sebenar merekod masa tamat (Bitiş) "
                    f"tetapi tiada jawapan tersimpan — kemungkinan masalah teknikal atau server.",
                    style={"color": "#9ca3af", "fontSize": "0.9rem", "marginBottom": "16px"},
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.P(f"Senarai Peserta Anomali (menunjukkan 20 daripada {len(anomaly_s):,})",
                                       style={"color": "#6b7280", "fontSize": "0.78rem",
                                              "marginBottom": "10px"}),
                                html.Div(build_anomaly_table(),
                                         style={"overflowX": "auto"}),
                            ], style={"padding": "16px"}),
                        ], style={"background": CARD_BG, "border": f"1px solid {BORDER}",
                                  "borderRadius": "10px"}),
                    ], md=12),
                ], className="g-3 mb-5"),

                # ══════════════════════════════════════════════════════════
                # APPENDIX : Anomali Pendaftaran
                # ══════════════════════════════════════════════════════════
                html.Hr(style={"borderColor": "rgba(255,255,255,0.06)",
                               "margin": "8px 0 28px"}),
                html.P("APPENDIX — ANOMALI PENDAFTARAN",
                       style={"color": "#6b7280", "fontSize": "0.7rem",
                              "letterSpacing": "0.12em", "fontWeight": "700",
                              "marginBottom": "22px",
                              "borderLeft": "3px solid #f59e0b",
                              "paddingLeft": "10px"}),

                # KPI row for appendix
                dbc.Row([
                    dbc.Col([
                        dbc.Card(dbc.CardBody([
                            html.P("E-mel Dikongsi (nama berbeza)",
                                   style={"color": "#9ca3af", "fontSize": "0.78rem",
                                          "marginBottom": "4px"}),
                            html.P(f"{_n_email_affected:,} e-mel",
                                   style={"color": "#f59e0b", "fontWeight": "700",
                                          "fontSize": "1.5rem", "margin": "0"}),
                            html.P(f"{_n_email_rows:,} baris rekod",
                                   style={"color": "#6b7280", "fontSize": "0.75rem",
                                          "marginTop": "2px"}),
                        ]), style={"background": CARD_BG,
                                   "border": "1px solid rgba(245,158,11,0.3)",
                                   "borderRadius": "10px"}),
                    ], md=3),
                    dbc.Col([
                        dbc.Card(dbc.CardBody([
                            html.P("Nama Sama Berbeza E-mel",
                                   style={"color": "#9ca3af", "fontSize": "0.78rem",
                                          "marginBottom": "4px"}),
                            html.P(f"{_n_name_affected:,} nama",
                                   style={"color": "#a78bfa", "fontWeight": "700",
                                          "fontSize": "1.5rem", "margin": "0"}),
                            html.P(f"{_n_name_rows:,} baris rekod",
                                   style={"color": "#6b7280", "fontSize": "0.75rem",
                                          "marginTop": "2px"}),
                        ]), style={"background": CARD_BG,
                                   "border": "1px solid rgba(167,139,250,0.3)",
                                   "borderRadius": "10px"}),
                    ], md=3),
                    dbc.Col([
                        dbc.Card(dbc.CardBody([
                            html.P("Jumlah Rekod Bermasalah",
                                   style={"color": "#9ca3af", "fontSize": "0.78rem",
                                          "marginBottom": "4px"}),
                            html.P(f"{_n_email_rows + _n_name_rows:,} baris",
                                   style={"color": "#ef4444", "fontWeight": "700",
                                          "fontSize": "1.5rem", "margin": "0"}),
                            html.P(f"daripada {len(df):,} pendaftaran",
                                   style={"color": "#6b7280", "fontSize": "0.75rem",
                                          "marginTop": "2px"}),
                        ]), style={"background": CARD_BG,
                                   "border": "1px solid rgba(239,68,68,0.3)",
                                   "borderRadius": "10px"}),
                    ], md=3),
                ], className="g-3 mb-4"),

                # ── Anomaly by negeri chart ─────────────────────────────
                html.P("Taburan Anomali Mengikut Negeri",
                       style={"color": "#e2e8f0", "fontWeight": "700",
                              "fontSize": "1.0rem", "marginBottom": "4px"}),
                html.P(
                    "Bilangan rekod anomali (Jenis A + Jenis B) mengikut negeri.",
                    style={"color": "#9ca3af", "fontSize": "0.88rem",
                           "marginBottom": "14px"},
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card(dbc.CardBody([
                            dcc.Graph(
                                figure=build_anomaly_negeri_chart(),
                                config={"displayModeBar": False},
                            ),
                        ], style={"padding": "12px"}),
                        style={"background": CARD_BG,
                               "border": "1px solid rgba(245,158,11,0.2)",
                               "borderRadius": "10px"}),
                    ], md=12),
                ], className="g-3 mb-5"),

                # ── Type A : Same email ─────────────────────────────────
                html.P("Jenis A — E-mel Dikongsi, Nama Berbeza",
                       style={"color": "#e2e8f0", "fontWeight": "700",
                              "fontSize": "1.0rem", "marginBottom": "4px"}),
                html.P(
                    f"Kemungkinan ibu bapa / penjaga mendaftarkan {_n_email_rows:,} anak "
                    f"menggunakan {_n_email_affected:,} e-mel yang sama. "
                    "Rekod dikumpulkan mengikut e-mel (garisan pemisah).",
                    style={"color": "#9ca3af", "fontSize": "0.88rem",
                           "marginBottom": "14px"},
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card(dbc.CardBody([
                            html.Div(
                                build_type_a_table(),
                                style={"maxHeight": "460px", "overflowY": "auto",
                                       "overflowX": "auto"},
                            ),
                        ], style={"padding": "12px"}),
                        style={"background": CARD_BG,
                               "border": "1px solid rgba(245,158,11,0.2)",
                               "borderRadius": "10px"}),
                    ], md=12),
                ], className="g-3 mb-5"),

                # ── Type B : Same name ──────────────────────────────────
                html.P("Jenis B — Nama Sama, E-mel Berbeza",
                       style={"color": "#e2e8f0", "fontWeight": "700",
                              "fontSize": "1.0rem", "marginBottom": "4px"}),
                html.P(
                    f"{_n_name_affected:,} nama unik muncul dengan {_n_name_rows:,} e-mel berbeza. "
                    "Kemungkinan pendaftaran berganda oleh peserta yang sama, "
                    "atau kebetulan nama sepunya.",
                    style={"color": "#9ca3af", "fontSize": "0.88rem",
                           "marginBottom": "14px"},
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card(dbc.CardBody([
                            html.Div(
                                build_type_b_table(),
                                style={"maxHeight": "460px", "overflowY": "auto",
                                       "overflowX": "auto"},
                            ),
                        ], style={"padding": "12px"}),
                        style={"background": CARD_BG,
                               "border": "1px solid rgba(167,139,250,0.2)",
                               "borderRadius": "10px"}),
                    ], md=12),
                ], className="g-3 mb-5"),

                # ── Footer ────────────────────────────────────────────────
                html.P(
                    "Sumber data: all-daftar-pertandingan.xlsx  ·  Dibina dengan Dash & Plotly",
                    style={"color": "#2a2a3a", "fontSize": "0.73rem",
                           "textAlign": "center", "paddingBottom": "16px"},
                ),
            ],
        ),
    ],
)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8050, debug=False)
