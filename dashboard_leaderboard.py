import pandas as pd
import re
import plotly.graph_objects as go
from dash import dcc, html
import dash
import dash_bootstrap_components as dbc

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


# ── School name normalisation ─────────────────────────────────────────────────
_SCHOOL_ABBR = [
    (r'^Smka\b',  'Sekolah Menengah Kebangsaan Agama'),
    (r'^Smk A\b', 'Sekolah Menengah Kebangsaan Agama'),
    (r'^Smk\b',   'Sekolah Menengah Kebangsaan'),
    (r'^Sma\b',   'Sekolah Menengah Agama'),
    (r'^Smss$',   'Sekolah Menengah Sains Selangor'),
    (r'^Sm Sains\b', 'Sekolah Menengah Sains'),
    (r'^Sms\b',   'Sekolah Menengah Sains'),
    (r'^Sm Agama\b', 'Sekolah Menengah Agama'),
    (r'^Sm Teknik\b', 'Sekolah Menengah Teknik'),
    (r'^Sm Imtiaz\b', 'Sekolah Menengah Imtiaz'),
    (r'^Sm\b',    'Sekolah Menengah'),
    (r'^Sbpi\b',  'Sekolah Berasrama Penuh Integrasi'),
    (r'^Sbp\b',   'Sekolah Berasrama Penuh'),
    (r'^Sek Men\b', 'Sekolah Menengah'),
    (r'^Sek\b',   'Sekolah'),
    (r'^Mrsm\b',  'Mrsm'),
    (r'^Snk\b',   'Sekolah Menengah Kebangsaan'),
]

_TYPO_FIX = {
    'Menegah': 'Menengah', 'Menangah': 'Menengah', 'Mencengah': 'Menengah',
    'Kenengah': 'Menengah', 'Mengengah': 'Menengah', 'Menenggah': 'Menengah',
    'Kebanggsaan': 'Kebangsaan', 'Kebangasaan': 'Kebangsaan',
    'Kebangsan': 'Kebangsaan', 'Kebanggasaan': 'Kebangsaan',
    'Kebabgsaan': 'Kebangsaan',
    'Intergrasi': 'Integrasi',
    'Idlam': 'Islam',
}

# Schools where the state/region name is part of the actual school name.
# Matched (case-insensitive) AFTER abbreviation expansion and typo fixes,
# BEFORE state-suffix stripping.  If matched → return canonical name.
_CANONICAL_SCHOOL = [
    (re.compile(r'\bSains\s+Selangor\b', re.I), 'Sm Sains Selangor'),
    (re.compile(r'\bSains\s+Kuala\s+Selangor\b', re.I), 'Sekolah Menengah Sains Kuala Selangor'),
    (re.compile(r'\bSains\s+Kuala\s+Te[rt]e?n?g{1,2}an?u{1,2}\b', re.I), 'Sekolah Menengah Sains Kuala Terengganu'),
    (re.compile(r'\bJalan\s+Apas\b', re.I), 'Sekolah Menengah Kebangsaan Jalan Apas'),
    (re.compile(r'^Sekolah\s+(?:Menengah\s+)?Kebangsaan\s+Tawau', re.I), 'Sekolah Menengah Kebangsaan Tawau'),
    (re.compile(r'^Sekolah\s+Berasrama\s+Penuh\s+Integrasi\s+(?:\([^)]*\)\s*)?Rawang', re.I), 'Sekolah Berasrama Penuh Integrasi Rawang'),
]

# State / location suffixes to strip
_STATE_SUFFIXES = [
    'Johor', 'Kedah', 'Kelantan', 'Melaka', 'Negeri Sembilan', 'Pahang',
    'Perak', 'Perlis', 'Pulau Pinang', 'Sabah', 'Sarawak', 'Selangor',
    'Terengganu', 'Kuala Lumpur', 'Labuan', 'Putrajaya',
    'W.P. Kuala Lumpur', 'W.P. Labuan', 'W.P. Putrajaya', 'W.P.K.L',
]

def _norm_school(val):
    if pd.isna(val):
        return "Tidak Diketahui"
    s = str(val).strip()
    # Remove dots in abbreviations like S.M.K, S.K
    s = re.sub(r'\b([A-Za-z])\.([A-Za-z])\.([A-Za-z])\b', r'\1\2\3', s)
    s = re.sub(r'\b([A-Za-z])\.([A-Za-z])\b', r'\1\2', s)
    # Split stuck-together abbreviations like "Smktawau" → "Smk Tawau"
    s = re.sub(r'^(Smka|Smk|Sma|Sms|Sbpi|Sbp|Mrsm|Snk)([A-Z])', lambda m: m.group(1) + ' ' + m.group(2), s, flags=re.IGNORECASE)
    # Also handle "Smsselangor" → "Sms Selangor" (lowercase after prefix)
    s = re.sub(r'^(Smka|Smk|Sma|Sms|Sbpi|Sbp|Mrsm)([a-z])', lambda m: m.group(1) + ' ' + m.group(2), s, flags=re.IGNORECASE)
    # Title case
    s = s.title()
    # Expand abbreviations
    for pat, repl in _SCHOOL_ABBR:
        s, n = re.subn(pat, repl, s, count=1, flags=re.IGNORECASE)
        if n:
            break
    # Handle "Sekolah Smka ..." / "Sekolah Smk ..." (abbr not at start)
    s = re.sub(r'^Sekolah Smka\b', 'Sekolah Menengah Kebangsaan Agama', s, flags=re.IGNORECASE)
    s = re.sub(r'^Sekolah Smk\b', 'Sekolah Menengah Kebangsaan', s, flags=re.IGNORECASE)
    # Normalise "Sekolah Berasrama Penuh I <place>" → "... Integrasi <place>"
    s = re.sub(r'^(Sekolah Berasrama Penuh) I\b', r'\1 Integrasi', s)
    # Fix typos
    for wrong, right in _TYPO_FIX.items():
        s = s.replace(wrong, right)
    # Check canonical school names (before state stripping)
    for cpat, cname in _CANONICAL_SCHOOL:
        if cpat.search(s):
            return cname
    # Strip trailing state names, commas, dots, whitespace
    s = re.sub(r'[,.]\s*$', '', s)
    for st in sorted(_STATE_SUFFIXES, key=len, reverse=True):
        s = re.sub(r'[,\s]+' + re.escape(st) + r'[.,\s]*$', '', s, flags=re.IGNORECASE)
    # Strip parenthetical suffixes like (Smss), (Sabk), (Saputra)
    s = re.sub(r'\s*\([^)]*\)\s*$', '', s)
    # Collapse multiple spaces
    s = re.sub(r'\s+', ' ', s).strip()
    return s


# ── School → negeri overrides ─────────────────────────────────────────────────
# Schools physically in a different state than most students registered under.
_SCHOOL_NEGERI_OVERRIDE = {
    'Sm Sains Selangor': 'Kuala Lumpur',
    'Sekolah Menengah Sains Kuala Selangor': 'Kuala Lumpur',
}

# ── Load & prepare data ──────────────────────────────────────────────────────
lb = pd.read_excel("all_leaderboard_export.xlsx")
lb["Negeri"]  = lb["Negeri"].astype(str).str.strip().apply(_norm_negeri)
lb["Sekolah"] = lb["Sekolah"].apply(_norm_school)
lb["Markah"]  = pd.to_numeric(lb["Markah"], errors="coerce").fillna(0)

# Override negeri for specific schools
for _sch, _target_neg in _SCHOOL_NEGERI_OVERRIDE.items():
    lb.loc[lb["Sekolah"] == _sch, "Negeri"] = _target_neg

# Top negeri by participant count
lb_neg_vc   = lb["Negeri"].value_counts()
lb_neg_all  = lb_neg_vc.reset_index()
lb_neg_all.columns = ["Negeri", "Count"]
lb_top3_neg = lb_neg_vc.head(3).index.tolist()

# Negeri list for breakdowns: top-3 + Kuala Lumpur (if not already in top 3)
lb_breakdown_neg = list(lb_top3_neg)
if "Kuala Lumpur" not in lb_breakdown_neg:
    lb_breakdown_neg.append("Kuala Lumpur")

# Per breakdown negeri: top-3 schools (score > 0)
lb_pos = lb[lb["Markah"] > 0].copy()
lb_school_data = {}
for _neg in lb_breakdown_neg:
    _sub = lb_pos[lb_pos["Negeri"] == _neg]
    _top = (
        _sub.groupby("Sekolah")
        .agg(Peserta=("Markah", "count"),
             Min=("Markah", "min"),
             Median=("Markah", "median"),
             Mean=("Markah", "mean"),
             Max=("Markah", "max"))
        .sort_values("Peserta", ascending=False)
        .head(3)
        .reset_index()
    )
    _top["Mean"] = _top["Mean"].round(1)
    lb_school_data[_neg] = _top

# Per school: skewness of Markah
lb_skew_data = {}
for _neg in lb_breakdown_neg:
    if _neg not in lb_school_data:
        continue
    for _sch in lb_school_data[_neg]["Sekolah"].tolist():
        _sub = lb_pos[(lb_pos["Negeri"] == _neg) & (lb_pos["Sekolah"] == _sch)]
        if len(_sub) < 3:
            continue
        lb_skew_data[(_neg, _sch)] = round(_sub["Markah"].skew(), 2)

# ── Styles ────────────────────────────────────────────────────────────────────
BG       = "#0b0b16"
CARD_BG  = "#13131f"
CARD_BG2 = "#1a1a2e"
BORDER   = "rgba(255,255,255,0.07)"
SECTION_LABEL_STYLE = {
    "color": "#555", "fontSize": "0.7rem", "letterSpacing": "0.12em",
    "fontWeight": "700", "textTransform": "uppercase", "marginBottom": "10px",
}

# ── Chart builders ────────────────────────────────────────────────────────────
def build_lb_negeri_bar():
    df_plot = lb_neg_all.sort_values("Count", ascending=True).tail(14)
    colors = ["#38bdf8" if n in lb_top3_neg else "#334155" for n in df_plot["Negeri"]]
    fig = go.Figure(go.Bar(
        x=df_plot["Count"], y=df_plot["Negeri"], orientation="h",
        marker_color=colors, text=df_plot["Count"], textposition="outside",
        textfont=dict(color="#ccc", size=11),
    ))
    fig.update_layout(
        template="plotly_dark", paper_bgcolor=BG, plot_bgcolor=BG,
        margin=dict(l=10, r=40, t=30, b=10), height=420,
        xaxis=dict(showgrid=False, visible=False),
        yaxis=dict(showgrid=False, tickfont=dict(size=11)),
    )
    return fig


def build_lb_school_table(negeri):
    sdf = lb_school_data.get(negeri, pd.DataFrame())
    if sdf.empty:
        return html.P("Tiada data", style={"color": "#666"})
    header = html.Thead(html.Tr([
        html.Th(c, style={"fontSize": "0.75rem", "color": "#888",
                          "borderBottom": "1px solid #333", "padding": "6px 10px"})
        for c in ["Sekolah", "Peserta", "Min", "Median", "Mean", "Max"]
    ]))
    rows = []
    for _, r in sdf.iterrows():
        rows.append(html.Tr([
            html.Td(r["Sekolah"], style={"fontSize": "0.78rem", "color": "#ccc",
                                          "maxWidth": "260px", "padding": "6px 10px"}),
            html.Td(str(int(r["Peserta"])), style={"fontSize": "0.78rem", "color": "#38bdf8",
                                                     "textAlign": "center", "padding": "6px 10px"}),
            html.Td(str(r["Min"]), style={"fontSize": "0.78rem", "color": "#aaa",
                                           "textAlign": "center", "padding": "6px 10px"}),
            html.Td(str(r["Median"]), style={"fontSize": "0.78rem", "color": "#aaa",
                                              "textAlign": "center", "padding": "6px 10px"}),
            html.Td(str(r["Mean"]), style={"fontSize": "0.78rem", "color": "#aaa",
                                            "textAlign": "center", "padding": "6px 10px"}),
            html.Td(str(r["Max"]), style={"fontSize": "0.78rem", "color": "#aaa",
                                           "textAlign": "center", "padding": "6px 10px"}),
        ]))
    body = html.Tbody(rows)
    return dbc.Table([header, body], bordered=False,
                     style={"marginBottom": "0", "background": "transparent",
                            "color": "#ccc"})


def build_lb_skew_chart(negeri, school):
    skew_val = lb_skew_data.get((negeri, school))
    if skew_val is None:
        return html.P("Tidak cukup data untuk skewness",
                       style={"color": "#666", "fontSize": "0.78rem"})
    _sub = lb_pos[(lb_pos["Negeri"] == negeri) & (lb_pos["Sekolah"] == school)]
    fig = go.Figure()
    fig.add_trace(go.Histogram(
        x=_sub["Markah"], nbinsx=12,
        marker_color="#38bdf8", opacity=0.85,
    ))
    color = "#34d399" if skew_val >= 0 else "#f87171"
    fig.add_annotation(
        text=f"Skewness: <b>{skew_val:+.2f}</b>",
        xref="paper", yref="paper", x=0.97, y=0.95,
        showarrow=False, font=dict(color=color, size=12),
        bgcolor="rgba(0,0,0,0.5)", borderpad=4,
    )
    fig.update_layout(
        template="plotly_dark", paper_bgcolor=BG, plot_bgcolor=BG,
        margin=dict(l=10, r=10, t=10, b=30), height=200,
        xaxis=dict(showgrid=False, title=dict(text="Markah", font=dict(size=10))),
        yaxis=dict(showgrid=False, title=dict(text="Peserta", font=dict(size=10))),
        bargap=0.08,
    )
    return dcc.Graph(figure=fig, config={"displayModeBar": False})


def _conclusion(text, accent="#4E9AF1"):
    return html.Div(
        [html.Span("\U0001f4cc ", style={"fontSize": "0.9rem"}),
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


# ── KPIs ──────────────────────────────────────────────────────────────────────
total_peserta = len(lb)
total_scoring = len(lb_pos)
avg_markah    = round(lb_pos["Markah"].mean(), 1)
max_markah    = lb_pos["Markah"].max()

def kpi_card(title, value, sub, color):
    return html.Div(
        style={
            "backgroundColor": CARD_BG,
            "borderRadius": "14px",
            "padding": "0",
            "overflow": "hidden",
            "boxShadow": f"0 4px 24px rgba(0,0,0,0.4), 0 0 0 1px rgba(255,255,255,0.05)",
        },
        children=[
            html.Div(style={"height": "3px", "background": color}),
            html.Div(
                style={"padding": "22px 24px 18px"},
                children=[
                    html.P(title, style={
                        "color": "#666", "marginBottom": "6px",
                        "fontSize": "0.73rem", "letterSpacing": "0.08em",
                        "fontWeight": "600", "textTransform": "uppercase",
                    }),
                    html.H3(value, style={
                        "color": "#e2e8f0", "fontWeight": "800",
                        "marginBottom": "4px", "fontSize": "1.75rem",
                    }),
                    html.P(sub, style={
                        "color": "#555", "fontSize": "0.7rem", "marginBottom": "0",
                    }),
                ],
            ),
        ],
    )


# ── App ───────────────────────────────────────────────────────────────────────
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])
server = app.server
app.title = "PMBA Leaderboard Insights"

app.layout = html.Div(
    style={"backgroundColor": BG, "minHeight": "100vh", "fontFamily": "'Inter', sans-serif"},
    children=[
        html.Div(
            style={"maxWidth": "1100px", "margin": "0 auto", "padding": "40px 24px"},
            children=[
                # ── Header ────────────────────────────────────────────
                html.H2("PMBA Leaderboard Insights",
                        style={"color": "#e2e8f0", "fontWeight": "800",
                               "fontSize": "1.6rem", "marginBottom": "4px"}),
                html.P("Analisis penyertaan mengikut negeri, sekolah & taburan markah.",
                       style={"color": "#666", "fontSize": "0.85rem",
                              "marginBottom": "32px"}),

                # ── KPI cards ─────────────────────────────────────────
                dbc.Row([
                    dbc.Col(kpi_card("Jumlah Peserta", f"{total_peserta:,}",
                                     "Semua entri leaderboard", "#38bdf8"), md=3),
                    dbc.Col(kpi_card("Markah > 0", f"{total_scoring:,}",
                                     f"{round(total_scoring/total_peserta*100,1)}% daripada jumlah",
                                     "#34d399"), md=3),
                    dbc.Col(kpi_card("Purata Markah", str(avg_markah),
                                     "Peserta yang scoring sahaja", "#a78bfa"), md=3),
                    dbc.Col(kpi_card("Markah Tertinggi", str(int(max_markah)),
                                     "Skor maksimum", "#f59e0b"), md=3),
                ], className="g-3 mb-5"),

                # ── Section 1: Negeri bar chart ───────────────────────
                html.P("PENYERTAAN MENGIKUT NEGERI", style=SECTION_LABEL_STYLE),
                html.P("Jumlah peserta leaderboard mengikut negeri (Top 3 diwarnakan biru).",
                       style={"color": "#666", "fontSize": "0.8rem",
                              "marginBottom": "16px"}),
                dbc.Row([
                    dbc.Col([
                        html.Div(
                            dcc.Graph(figure=build_lb_negeri_bar(),
                                      config={"displayModeBar": False}),
                            style={"background": CARD_BG,
                                   "borderRadius": "12px", "padding": "16px",
                                   "border": f"1px solid {BORDER}"},
                        ),
                    ], md=12),
                ], className="g-3 mb-4"),

                _conclusion(
                    f"Top 3 negeri dengan penyertaan tertinggi: "
                    f"{', '.join(lb_top3_neg)}. "
                    f"Jumlah keseluruhan peserta leaderboard: {total_peserta:,}."
                ),

                # ── Section 2: Top 3 schools per negeri ───────────────
                html.P("TOP 3 SEKOLAH PER NEGERI (MARKAH > 0)", style=SECTION_LABEL_STYLE),
                html.P("Sekolah dengan peserta scoring terbanyak bagi setiap negeri utama.",
                       style={"color": "#666", "fontSize": "0.8rem",
                              "marginBottom": "16px"}),

                *[html.Div([
                    html.H5(f"\U0001f4cd {neg}",
                            style={"color": "#38bdf8", "fontSize": "0.95rem",
                                   "fontWeight": "600", "marginBottom": "10px",
                                   "marginTop": "20px"}),
                    html.Div(
                        build_lb_school_table(neg),
                        style={"background": CARD_BG, "borderRadius": "12px",
                               "padding": "14px", "border": f"1px solid {BORDER}",
                               "overflowX": "auto", "marginBottom": "10px"},
                    ),
                    # ── Section 3: Skewness per school ────────────────
                    html.P("Taburan Markah & Skewness",
                           style={"color": "#888", "fontSize": "0.75rem",
                                  "marginTop": "12px", "marginBottom": "6px"}),
                    dbc.Row([
                        dbc.Col([
                            html.P(sch, style={"color": "#ccc", "fontSize": "0.75rem",
                                               "marginBottom": "2px",
                                               "textAlign": "center"}),
                            html.Div(
                                build_lb_skew_chart(neg, sch),
                                style={"background": CARD_BG, "borderRadius": "10px",
                                       "padding": "10px",
                                       "border": f"1px solid {BORDER}"},
                            ),
                        ], md=4)
                        for sch in lb_school_data.get(neg, pd.DataFrame()).get("Sekolah", pd.Series()).tolist()
                    ], className="g-3 mb-4"),
                ]) for neg in lb_breakdown_neg],

                _conclusion(
                    "Skewness positif menunjukkan majoriti peserta mendapat markah rendah "
                    "(ekor panjang ke kanan). Skewness negatif bermakna kebanyakan peserta "
                    "skor tinggi. Nilai hampir 0 = taburan sekata."
                ),

                # ── Footer ────────────────────────────────────────────
                html.P(
                    "Sumber data: all_leaderboard_export.xlsx  \u00b7  Dibina dengan Dash & Plotly",
                    style={"color": "#2a2a3a", "fontSize": "0.73rem",
                           "textAlign": "center", "paddingBottom": "16px"},
                ),
            ],
        ),
    ],
)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8052, debug=False)
