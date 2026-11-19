"""
Honda in Mexico, 2015-2025: a consulting breakdown.
Source: INEGI RAIAVL (light-vehicle retail sales), master file INEGI_Auto_Ventas_Maestro_2015_2025.xlsx

Structure (issue tree):
    Key question -> 01 Financials  |  02 Automotive industry  |  03 Customer preferences -> Conclusions & 2030 plan

Run:  python -m streamlit run honda_dashboard.py
"""
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# Page + design tokens
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Honda Mexico 2015–2025",
    page_icon=":material/directions_car:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA = Path(__file__).parent / "INEGI_Auto_Ventas_Maestro_2015_2025.xlsx"

# Story colors: every color has one job and keeps it across the whole page.
HONDA = "#d62839"     # the protagonist
TOYOTA = "#2a78d6"    # the benchmark
CHINA = "#eda100"     # the disruptors
GREY = "#c9c7bf"      # everyone else (context, never competes for attention)
GREY_D = "#898781"    # market / reference lines
PEER = "#52514e"      # Japanese peer group (secondary emphasis)
EV_HEV, EV_PHEV, EV_BEV = "#4fbf8f", "#16936a", "#0b5a40"  # electrified (ordinal ramp)

# Segment colors (used only when a chart is split by segment; kept away from the brand hues above)
SEG_ORDER = ["SUVs", "Pickups", "Subcompact", "Compact", "Minivans", "Luxury", "Sports"]
SEG_COL = {
    "SUVs": "#3b3486", "Compact": "#7c74c9", "Subcompact": "#b9b3e6",
    "Pickups": "#4f5b64", "Minivans": "#97a2aa", "Luxury": "#c2b48a", "Sports": "#e0d9c6",
}

INK, INK2, MUTED = "#0b0b0b", "#52514e", "#898781"
GRID, AXIS, SURFACE = "#ebeae4", "#c3c2b7", "#ffffff"
FONT = 'system-ui, -apple-system, "Segoe UI", Roboto, sans-serif'
BAND = "rgba(137,135,129,0.15)"
BAND2 = "rgba(137,135,129,0.08)"

# Bucket identity colors (each matches the bucket's main protagonist)
B1, B2, B3 = HONDA, TOYOTA, CHINA

SEG_EN = {
    "SUV's": "SUVs", "Compactos": "Compact", "Subcompactos": "Subcompact",
    "Pick Ups": "Pickups", "Minivans": "Minivans", "De Lujo": "Luxury", "Deportivos": "Sports",
}

# Brand-of-origin groups (where the brand is from, not where the car is built)
ORIGIN = {
    **dict.fromkeys(["Nissan", "Toyota", "Mazda", "Honda", "Suzuki", "Mitsubishi", "Subaru", "Lexus",
                     "Isuzu", "Infiniti", "Acura"], "Japanese"),
    **dict.fromkeys(["General Motors", "Chrysler", "Ford Motor", "Lincoln"], "American"),
    **dict.fromkeys(["KIA", "Hyundai"], "Korean"),
}

CSS = """
<style>
.block-container {padding-top: 1.2rem; max-width: 1240px;}
h1, h2, h3, h4, p, li, div {font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;}
.hero {border-radius: 22px; padding: 40px 44px 34px; color: #fff;
  background: radial-gradient(1200px 400px at 85% 0%, rgba(214,40,57,.55), transparent 60%),
              linear-gradient(135deg, #141413 0%, #24231f 65%, #2e1215 100%);}
.hero .eyebrow {letter-spacing: .18em; text-transform: uppercase; font-size: 12px; color: #f3a6ae; font-weight: 600;}
.hero h1 {font-size: clamp(30px, 4.4vw, 50px); line-height: 1.05; margin: 10px 0 12px; color: #fff; font-weight: 750;}
.hero p.lede {font-size: 18px; color: #d9d7d0; max-width: 780px; margin: 0;}
.hero-stats {display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 14px; margin-top: 28px;}
.hero-stat {background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.10); border-radius: 14px; padding: 16px 18px;}
.hero-stat .lbl {font-size: 12px; color: #b9b7ae; text-transform: uppercase; letter-spacing: .08em;}
.hero-stat .val {font-size: 30px; font-weight: 700; margin-top: 4px; color: #fff;}
.hero-stat .val .from {color: #9b998f; font-weight: 500;}
.hero-stat .val .to {color: #ff6b78;}
.hero-stat .sub {font-size: 13px; color: #b9b7ae; margin-top: 2px;}

/* Issue tree */
.tree {margin: 40px 0 8px;}
.tree-q {background: #141413; color: #fff; border-radius: 16px; padding: 20px 26px; text-align: center; max-width: 900px; margin: 0 auto;}
.tree-q .eyebrow {letter-spacing: .16em; text-transform: uppercase; font-size: 11.5px; color: #b9b7ae; font-weight: 700;}
.tree-q h3 {color: #fff; font-size: clamp(18px, 2.2vw, 23px); line-height: 1.3; margin: 6px 0 0; font-weight: 700;}
.rail {height: 24px; margin: 24px 16.67% 0; border: 2px solid #c3c2b7; border-bottom: none; border-radius: 10px 10px 0 0; position: relative;}
.rail::before {content: ""; position: absolute; left: 50%; top: -26px; height: 24px; border-left: 2px solid #c3c2b7;}
.rail::after {content: ""; position: absolute; left: 50%; top: 0; height: 22px; border-left: 2px solid #c3c2b7;}
.buckets {display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px;}
a.bk {display: block; text-decoration: none !important; color: inherit !important; background: #fff; border: 1px solid rgba(11,11,11,.08);
  border-radius: 16px; padding: 18px 20px 16px; border-top: 6px solid; transition: transform .15s ease, box-shadow .15s ease;}
a.bk:hover {transform: translateY(-2px); box-shadow: 0 10px 24px rgba(11,11,11,.08);}
.bk .n {font-size: 12px; font-weight: 800; letter-spacing: .14em; text-transform: uppercase;}
.bk h3 {font-size: 23px; margin: 4px 0 4px; color: #0b0b0b; font-weight: 750;}
.bk .q {font-size: 14.5px; color: #52514e; margin: 0 0 10px; min-height: 42px;}
.bk .f {display: flex; justify-content: space-between; gap: 10px; border-top: 1px solid #efeee9; padding: 7px 0; font-size: 13.5px; color: #52514e;}
.bk .f b {color: #0b0b0b; white-space: nowrap;}
.bk .subs {margin: 8px 0 0; padding: 0; list-style: none; font-size: 12.5px; color: #898781;}
.bk .subs li {padding: 2px 0;}
.bk .go {font-size: 13px; font-weight: 750; margin-top: 10px;}
.down {text-align: center; color: #c3c2b7; font-size: 22px; line-height: 1; margin: 10px 0 6px;}
a.sowhat {display: flex; align-items: center; justify-content: space-between; gap: 14px; text-decoration: none !important;
  background: #141413; border-radius: 14px; padding: 15px 22px; color: #e8e6df !important;}
a.sowhat b {color: #fff;} a.sowhat .arr {color: #ff6b78; font-weight: 800; white-space: nowrap;}
@media (max-width: 860px) {.buckets {grid-template-columns: 1fr;} .rail {display: none;} .tree-q {margin-bottom: 14px;}}

/* Bucket headers */
.bhd {margin: 80px 0 14px; border-radius: 18px; padding: 24px 28px; background: #fff; border: 1px solid rgba(11,11,11,.08);
  border-left: 10px solid; display: grid; grid-template-columns: auto 1fr; gap: 24px; align-items: center;}
.bhd .big {font-size: 64px; font-weight: 800; line-height: .9; letter-spacing: -.02em;}
.bhd .eyebrow {letter-spacing: .16em; text-transform: uppercase; font-size: 12px; font-weight: 800;}
.bhd h2 {font-size: clamp(24px, 3vw, 32px); line-height: 1.15; margin: 4px 0 8px; font-weight: 740; color: #0b0b0b;}
.bhd p {font-size: 16px; color: #52514e; margin: 0; max-width: 860px;}
.bhd .chips {display: flex; gap: 8px; flex-wrap: wrap; margin-top: 12px;}
.bhd .chip {font-size: 12.5px; border-radius: 999px; padding: 4px 11px; background: #f3f2ee; color: #2b2a27;}
@media (max-width: 720px) {.bhd {grid-template-columns: 1fr; gap: 6px;} .bhd .big {font-size: 44px;}}
.subh {margin: 34px 0 8px;}
.subh .num {font-size: 12px; font-weight: 800; letter-spacing: .12em;}
.subh h3 {font-size: 21px; margin: 2px 0 4px; color: #0b0b0b; font-weight: 720;}
.subh p {font-size: 15px; color: #52514e; margin: 0; max-width: 860px;}
.take {margin: 22px 0 0; background: #141413; color: #d9d7d0; border-radius: 14px; padding: 16px 20px; font-size: 15.5px;
  display: grid; grid-template-columns: auto 1fr; gap: 16px; align-items: start;}
.take .t {font-size: 11.5px; font-weight: 800; letter-spacing: .14em; text-transform: uppercase; padding-top: 3px; white-space: nowrap;}
.take b {color: #fff;}

.chapter {margin: 80px 0 14px;}
.chapter .eyebrow {letter-spacing: .16em; text-transform: uppercase; font-size: 12px; font-weight: 700;}
.chapter h2 {font-size: clamp(24px, 3vw, 34px); line-height: 1.15; margin: 6px 0 10px; font-weight: 720; color: #0b0b0b;}
.chapter p {font-size: 17px; color: #52514e; max-width: 820px; margin: 0;}
.callout {border-left: 4px solid; border-radius: 10px; padding: 14px 18px; margin: 10px 0 6px; font-size: 15.5px; color: #2b2a27;}
.callout b {color: #0b0b0b;}
.kpis {display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 14px; margin: 8px 0 4px;}
.kpi {background: #fff; border: 1px solid rgba(11,11,11,.08); border-radius: 14px; padding: 16px 18px; border-top: 4px solid #c9c7bf;}
.kpi .lbl {font-size: 12.5px; color: #898781; text-transform: uppercase; letter-spacing: .06em; font-weight: 600;}
.kpi .val {font-size: 28px; font-weight: 720; color: #0b0b0b; margin-top: 4px; line-height: 1.1;}
.kpi .val .from {color: #898781; font-weight: 550;}
.kpi .sub {font-size: 13.5px; color: #52514e; margin-top: 4px;}
.vs {display: grid; grid-template-columns: 1fr auto 1fr; gap: 18px; align-items: stretch; margin: 6px 0 8px;}
.vs .side {background: #fff; border: 1px solid rgba(11,11,11,.08); border-radius: 16px; padding: 20px 22px;}
.vs .brand {font-size: 13px; font-weight: 800; letter-spacing: .14em; text-transform: uppercase;}
.vs .units {font-size: clamp(30px, 4vw, 44px); font-weight: 760; color: #0b0b0b; line-height: 1.05; margin-top: 6px;}
.vs .meta {font-size: 14px; color: #52514e; margin-top: 6px;}
.vs .mid {display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; min-width: 120px;}
.vs .mid .x {font-size: 13px; color: #898781; text-transform: uppercase; letter-spacing: .12em;}
.vs .mid .ratio {font-size: 30px; font-weight: 760; color: #0b0b0b;}
.vs .mid .cap {font-size: 12.5px; color: #52514e; max-width: 140px;}
@media (max-width: 720px) {.vs {grid-template-columns: 1fr;} .hero {padding: 28px 22px;}}
.score {display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 14px; margin-bottom: 14px;}
.sc {background: #fff; border: 1px solid rgba(11,11,11,.08); border-radius: 14px; padding: 16px 18px;}
.sc .lbl {font-size: 12.5px; color: #898781; text-transform: uppercase; letter-spacing: .06em; font-weight: 600;}
.sc .row {display: flex; align-items: baseline; gap: 10px; margin-top: 8px; flex-wrap: wrap;}
.sc .y {font-size: 11.5px; color: #898781; font-weight: 600;}
.sc .a {font-size: 22px; color: #898781; font-weight: 600;}
.sc .arrow {color: #c3c2b7; font-size: 18px;}
.sc .b {font-size: 28px; color: #0b0b0b; font-weight: 750;}
.badge {display: inline-block; font-size: 12.5px; font-weight: 700; padding: 3px 9px; border-radius: 999px; margin-top: 8px;}
.badge.bad {background: #fbe7e9; color: #a3142a;}
.badge.good {background: #e3f3e6; color: #0a5c1c;}
.badge.neu {background: #efeee9; color: #52514e;}
.concl {display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 14px;}
.cc {background: #141413; color: #e8e6df; border-radius: 16px; padding: 20px 22px; border-top: 5px solid;}
.cc .n {font-weight: 800; font-size: 12px; letter-spacing: .12em; text-transform: uppercase;}
.cc h4 {color: #fff; margin: 6px 0 8px; font-size: 19px; line-height: 1.25;}
.cc p {color: #c9c7bf; font-size: 14.5px; margin: 0;}
.rec {background: #fff; border: 1px solid rgba(11,11,11,.08); border-radius: 16px; padding: 18px 22px; margin-bottom: 12px;
  display: grid; grid-template-columns: 52px 1fr; gap: 14px;}
.rec .num {width: 42px; height: 42px; border-radius: 50%; background: #d62839; color: #fff; display: flex; align-items: center;
  justify-content: center; font-weight: 800; font-size: 18px;}
.rec h4 {margin: 2px 0 6px; font-size: 18px; color: #0b0b0b;}
.rec p {margin: 0 0 8px; color: #52514e; font-size: 15px;}
.rec .tags {display: flex; gap: 8px; flex-wrap: wrap;}
.tag {font-size: 12.5px; border-radius: 8px; padding: 4px 10px; background: #f3f2ee; color: #2b2a27;}
.tag b {color: #0b0b0b;}
.tag.bk1 {background: #fbe7e9;} .tag.bk2 {background: #e6effb;} .tag.bk3 {background: #fdf3dc;}
.road {display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 14px; margin-top: 6px;}
.rd {border-radius: 14px; padding: 16px 18px; background: #fff; border: 1px solid rgba(11,11,11,.08); border-top: 4px solid #d62839;}
.rd .when {font-size: 12px; font-weight: 800; letter-spacing: .12em; color: #d62839; text-transform: uppercase;}
.rd ul {margin: 8px 0 0 18px; padding: 0; color: #2b2a27; font-size: 14.5px;}
.foot {color: #898781; font-size: 12.5px; margin-top: 48px; border-top: 1px solid #e1e0d9; padding-top: 14px;}
.legend-inline span {display: inline-flex; align-items: center; gap: 6px; margin-right: 16px; font-size: 13.5px; color: #52514e;}
.legend-inline i {width: 12px; height: 12px; border-radius: 3px; display: inline-block;}
</style>
"""


def html(s: str) -> None:
    """Render HTML without Markdown turning indented lines into code blocks."""
    st.markdown("\n".join(ln.strip() for ln in s.splitlines() if ln.strip()), unsafe_allow_html=True)


def bucket_header(num: str, name: str, color: str, title: str, lede: str, chips: list[str], anchor: str) -> None:
    chip_html = "".join(f'<span class="chip">{c}</span>' for c in chips)
    html(f"""
    <div class="bhd" id="{anchor}" style="border-left-color:{color}">
      <div class="big" style="color:{color}">{num}</div>
      <div><div class="eyebrow" style="color:{color}">Bucket {num} · {name}</div>
        <h2>{title}</h2><p>{lede}</p><div class="chips">{chip_html}</div></div>
    </div>""")


def subhead(num: str, title: str, lede: str = "", color: str = INK2) -> None:
    html(f'<div class="subh"><div class="num" style="color:{color}">{num}</div><h3>{title}</h3>'
         + (f"<p>{lede}</p>" if lede else "") + "</div>")


def chapter(num: str, title: str, lede: str, color: str = HONDA, anchor: str = "") -> None:
    html(f"""
    <div class="chapter" id="{anchor}">
      <div class="eyebrow" style="color:{color}">{num}</div>
      <h2>{title}</h2>
      <p>{lede}</p>
    </div>""")


def callout(text: str, color: str = HONDA, tint: str = "#fdf0f1") -> None:
    html(f'<div class="callout" style="border-color:{color};background:{tint}">{text}</div>')


def takeaway(text: str, color: str) -> None:
    html(f'<div class="take"><div class="t" style="color:{color}">So what</div><div>{text}</div></div>')


def style(fig: go.Figure, height: int = 380, title: str = "", legend: bool = True) -> go.Figure:
    fig.update_layout(
        height=height,
        title=dict(text=f"<b>{title}</b>" if title else "", font=dict(size=15, color=INK), x=0, xanchor="left", y=0.98),
        margin=dict(l=8, r=16, t=64 if title else 24, b=8),
        paper_bgcolor=SURFACE, plot_bgcolor=SURFACE,
        font=dict(family=FONT, color=INK2, size=13),
        hoverlabel=dict(bgcolor="#ffffff", bordercolor=GRID, font=dict(color=INK, family=FONT, size=13)),
        legend=dict(orientation="h", x=0, y=1.02, yanchor="bottom", font=dict(size=12.5, color=INK2), bgcolor="rgba(0,0,0,0)"),
        showlegend=legend,
        bargap=0.28,
    )
    fig.update_xaxes(showgrid=False, zeroline=False, linecolor=AXIS, ticks="", tickfont=dict(color=MUTED))
    fig.update_yaxes(gridcolor=GRID, gridwidth=1, zeroline=False, linecolor=AXIS, tickfont=dict(color=MUTED))
    return fig


def show(fig: go.Figure) -> None:
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})


def color_for(brand: str) -> str:
    if brand == "Honda":
        return HONDA
    if brand == "Toyota":
        return TOYOTA
    if brand in CHINESE or brand.startswith("Chinese"):
        return CHINA
    return GREY


def shock_bands(fig: go.Figure, labels: bool = True) -> None:
    """COVID-19 (2020) and chip shortage (2021-2022) on an annual numeric x-axis."""
    fig.add_vrect(x0=2019.5, x1=2020.5, fillcolor=BAND, line_width=0, layer="below")
    fig.add_vrect(x0=2020.5, x1=2022.5, fillcolor=BAND2, line_width=0, layer="below")
    if labels:
        for x, t in ((2020, "COVID-19"), (2021.5, "Chip shortage")):
            fig.add_annotation(x=x, y=1, yref="paper", text=t, showarrow=False, yanchor="top", yshift=-2,
                               font=dict(size=11, color=MUTED))


def year_axis(fig: go.Figure, a: int, b: int, compact: bool = False) -> None:
    """Year ticks, with the two compared years in bold. compact=True thins ticks for half-width charts."""
    ys = sorted({Y0, Y1, a, b, *range(Y0, Y1 + 1, 2)}) if compact else YEARS
    if compact:  # drop odd ticks that sit right next to a selected year
        ys = [y for y in ys if y in (a, b, Y0, Y1) or all(abs(y - s) > 1 for s in (a, b))]
    fig.update_xaxes(tickvals=ys, ticktext=[f"<b>{y}</b>" if y in (a, b) else str(y) for y in ys],
                     tickfont=dict(color=INK2))


def year_range(key: str, color: str) -> tuple[int, int]:
    """Range slider that picks the two years to compare. Returns (a, b) with a < b."""
    with st.container(border=True):
        c1, c2 = st.columns([3.2, 1], vertical_alignment="center")
        with c1:
            a, b = st.select_slider(
                "Compare years", options=YEARS, value=(Y0, Y1), key=key,
                help="Drag either handle. Every comparison in this bucket updates to the two years you pick.",
            )
        with c2:
            html(f'<div style="font-size:13px;color:{MUTED};line-height:1.35">Comparing<br>'
                 f'<span style="font-size:22px;font-weight:750;color:{INK}">{a} <span style="color:{AXIS}">→</span> '
                 f'<span style="color:{color}">{b}</span></span></div>')
    if a == b:
        st.caption("Both handles are on the same year, so the comparison uses the next available year.")
        b = a + 1 if a < Y1 else a
        a = a if a < Y1 else a - 1
    return a, b


def pct(new: float, old: float) -> float:
    return (new / old - 1) * 100 if old else float("nan")


def cagr(new: float, old: float, n: int) -> float:
    return ((new / old) ** (1 / n) - 1) * 100 if old and n else float("nan")


def fmt_k(v: float) -> str:
    return f"{v/1e6:.2f}M" if v >= 1e6 else f"{v/1e3:.1f}k"


# ─────────────────────────────────────────────────────────────────────────────
# Data
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading INEGI RAIAVL data…")
def load_data():
    d = pd.read_excel(
        DATA, sheet_name="Datos",
        usecols=["ANIO", "ID_MES", "FECHA", "MARCA", "MODELO", "TIPO", "SEGMENTO", "ORIGEN", "UNI_VEH", "Categoria"],
    )
    d["SEG"] = d["SEGMENTO"].map(SEG_EN).fillna(d["SEGMENTO"])
    d["FECHA"] = pd.to_datetime(d["FECHA"])
    e = pd.read_excel(
        DATA, sheet_name="Datos_HEV",
        usecols=["ANIO", "VEH_ELECTR", "VEH_HIBRIDAS_PLUGIN", "VEH_HIBRIDAS", "TOTAL_ELECTRIFICADOS"],
    )
    e = e[e["ANIO"] <= 2025].groupby("ANIO").sum()
    return d, e


df, ev = load_data()
YEARS = [int(y) for y in sorted(df["ANIO"].unique())]
Y0, Y1 = YEARS[0], YEARS[-1]
CHINESE = sorted(df.loc[df["Categoria"] == "China", "MARCA"].unique())

tot = df.groupby("ANIO")["UNI_VEH"].sum()
brand = df.pivot_table(index="ANIO", columns="MARCA", values="UNI_VEH", aggfunc="sum", fill_value=0)
share = brand.div(tot, axis=0) * 100
rank = brand.rank(axis=1, ascending=False, method="min").astype(int)
china_u = brand[CHINESE].sum(axis=1)
china_s = china_u / tot * 100

H, T = brand["Honda"], brand["Toyota"]
hs, ts = share["Honda"], share["Toyota"]

# Segment views
mseg = df.pivot_table(index="ANIO", columns="SEG", values="UNI_VEH", aggfunc="sum", fill_value=0)
mseg = mseg[[s for s in SEG_ORDER if s in mseg.columns]]
mmix = mseg.div(mseg.sum(axis=1), axis=0) * 100
hseg = (df[df["MARCA"] == "Honda"].pivot_table(index="ANIO", columns="SEG", values="UNI_VEH", aggfunc="sum", fill_value=0)
        .reindex(index=YEARS, columns=mseg.columns, fill_value=0))
HSEGS = [s for s in hseg.columns if hseg[s].max() > 100]
hmix = hseg.div(hseg.sum(axis=1), axis=0) * 100

# Brand-of-origin groups
grp_of = {b: ("Chinese" if b in CHINESE else ORIGIN.get(b, "European & other")) for b in brand.columns}
grp = brand.T.groupby(grp_of).sum().T
grp_s = grp.div(tot, axis=0) * 100
jap_ex = grp_s["Japanese"] - hs
honda_in_jp = H / grp["Japanese"] * 100


def seg_view(year: int) -> pd.DataFrame:
    s = df[df["ANIO"] == year].pivot_table(index="SEG", columns="MARCA", values="UNI_VEH", aggfunc="sum", fill_value=0)
    out = pd.DataFrame({"total": s.sum(axis=1)})
    out["mkt_pct"] = out["total"] / out["total"].sum() * 100
    for b in ("Honda", "Toyota"):
        out[f"{b}_u"] = s[b] if b in s else 0
        out[f"{b}_s"] = out[f"{b}_u"] / out["total"] * 100
    return out.sort_values("total", ascending=False)


def nameplates(year: int) -> pd.Series:
    x = df[(df["ANIO"] == year) & (df["MARCA"] == "Honda") & (df["UNI_VEH"] > 0)].copy()
    x["NP"] = x["MODELO"].str.replace("Honda ", "", regex=False).str.split().str[0].str.strip("-")
    return x.groupby("NP")["UNI_VEH"].sum().sort_values(ascending=False)


def origin_share(year: int, b: str = "Honda") -> float:
    x = df[(df["ANIO"] == year) & (df["MARCA"] == b)].groupby("ORIGEN")["UNI_VEH"].sum()
    return x.get("NACIONAL", 0) / x.sum() * 100


seg0, seg1 = seg_view(Y0), seg_view(Y1)
_suv0 = df[(df["ANIO"] == Y0) & (df["SEG"] == "SUVs")].groupby("MARCA")["UNI_VEH"].sum().rank(ascending=False, method="min")
suv_rank0 = int(_suv0["Honda"])
np0, np1 = nameplates(Y0), nameplates(Y1)
e0, e1 = int(ev.index.min()), int(ev.index.max())
pen = ev["TOTAL_ELECTRIFICADOS"] / tot.reindex(ev.index) * 100
pen0, pen1 = pen[e0], pen[e1]

# ─────────────────────────────────────────────────────────────────────────────
# Sidebar: navigation follows the issue tree
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### The breakdown")
    for a_, label in [
        ("overview", "Key question & issue tree"), ("b1", "01 · Financials"), ("b2", "02 · Automotive industry"),
        ("b3", "03 · Customer preferences"), ("concl", "Conclusions & 2030 plan"),
    ]:
        st.markdown(f"[{label}](#{a_})")
    st.caption("Source: INEGI, RAIAVL — retail sales of light vehicles, Mexico. Units sold to the public.")

st.markdown(CSS, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# HERO
# ─────────────────────────────────────────────────────────────────────────────
one_in_0, one_in_1 = 100 / hs[Y0], 100 / hs[Y1]
html(f"""
<div class="hero">
  <div class="eyebrow">Honda Mexico · consulting breakdown · {Y0}–{Y1}</div>
  <h1>Honda didn't lose a crisis.<br>It lost a decade.</h1>
  <p class="lede">In {Y0}, one in every {one_in_0:.0f} new vehicles sold in Mexico was a Honda. By {Y1}, it was one in {one_in_1:.0f}.
  We break the problem into three buckets: Honda's own volume, the industry around it, and what customers now buy.</p>
  <div class="hero-stats">
    <div class="hero-stat"><div class="lbl">Market share</div>
      <div class="val"><span class="from">{hs[Y0]:.1f}%</span> → <span class="to">{hs[Y1]:.1f}%</span></div>
      <div class="sub">{hs[Y1]-hs[Y0]:+.1f} pp in ten years</div></div>
    <div class="hero-stat"><div class="lbl">Units sold</div>
      <div class="val"><span class="from">{H[Y0]/1e3:.1f}k</span> → <span class="to">{H[Y1]/1e3:.1f}k</span></div>
      <div class="sub">{pct(H[Y1], H[Y0]):+.0f}% while the market grew {pct(tot[Y1], tot[Y0]):+.0f}%</div></div>
    <div class="hero-stat"><div class="lbl">Brand ranking</div>
      <div class="val"><span class="from">#{rank.loc[Y0,'Honda']}</span> → <span class="to">#{rank.loc[Y1,'Honda']}</span></div>
      <div class="sub">Toyota went #{rank.loc[Y0,'Toyota']} → #{rank.loc[Y1,'Toyota']}</div></div>
    <div class="hero-stat"><div class="lbl">Recovery vs 2019</div>
      <div class="val"><span class="to">{H[Y1]/H[2019]*100:.0f}%</span></div>
      <div class="sub">Market is at {tot[Y1]/tot[2019]*100:.0f}% of its 2019 level</div></div>
  </div>
</div>""")

# ─────────────────────────────────────────────────────────────────────────────
# OVERVIEW — issue tree
# ─────────────────────────────────────────────────────────────────────────────
suv_mix0, suv_mix1 = mmix.loc[Y0, "SUVs"], mmix.loc[Y1, "SUVs"]
html(f"""
<div class="tree" id="overview">
  <div class="tree-q"><div class="eyebrow">Key question</div>
    <h3>Why did Honda's share of Mexico's market fall from {hs[Y0]:.1f}% to {hs[Y1]:.1f}%,
    and what would it take to get back above 4% by 2030?</h3></div>
  <div class="rail"></div>
  <div class="buckets">
    <a class="bk" href="#b1" style="border-top-color:{B1}">
      <div class="n" style="color:{B1}">Bucket 01</div><h3>Financials</h3>
      <p class="q">How much volume did Honda lose, in which segments, and when?</p>
      <div class="f"><span>Units {Y0} → {Y1}</span><b>{H[Y0]/1e3:.1f}k → {H[Y1]/1e3:.1f}k</b></div>
      <div class="f"><span>Change in units</span><b style="color:{HONDA}">{pct(H[Y1], H[Y0]):+.0f}%</b></div>
      <div class="f"><span>{Y1} volume vs pre-COVID 2019</span><b>{H[Y1]/H[2019]*100:.0f}%</b></div>
      <ul class="subs"><li>1.1 Units by year and segment</li><li>1.2 Segment and nameplate comparison</li><li>1.3 COVID-19 and chip shortage lens</li></ul>
      <div class="go" style="color:{B1}">Open bucket →</div></a>
    <a class="bk" href="#b2" style="border-top-color:{B2}">
      <div class="n" style="color:{B2}">Bucket 02</div><h3>Automotive industry</h3>
      <p class="q">Did the market shrink, or did it grow somewhere Honda isn't?</p>
      <div class="f"><span>Market units {Y0} → {Y1}</span><b>{pct(tot[Y1], tot[Y0]):+.0f}%</b></div>
      <div class="f"><span>Electrified share of sales</span><b>{pen0:.1f}% → {pen1:.1f}%</b></div>
      <div class="f"><span>Toyota sells × Honda</span><b style="color:{TOYOTA}">{T[Y0]/H[Y0]:.1f}× → {T[Y1]/H[Y1]:.1f}×</b></div>
      <ul class="subs"><li>2.1 Market size by year and segment</li><li>2.2 Electrified vehicles</li><li>2.3 Closest competitor: Toyota</li></ul>
      <div class="go" style="color:{B2}">Open bucket →</div></a>
    <a class="bk" href="#b3" style="border-top-color:{B3}">
      <div class="n" style="color:{B3}">Bucket 03</div><h3>Customer preferences</h3>
      <p class="q">Have buyers shifted segments, origins and brands?</p>
      <div class="f"><span>SUVs, share of purchases</span><b>{suv_mix0:.0f}% → {suv_mix1:.0f}%</b></div>
      <div class="f"><span>Chinese brands' share</span><b style="color:#a87200">{china_s[Y0]:.1f}% → {china_s[Y1]:.1f}%</b></div>
      <div class="f"><span>Honda's slice of Japanese-brand sales</span><b>{honda_in_jp[Y0]:.0f}% → {honda_in_jp[Y1]:.0f}%</b></div>
      <ul class="subs"><li>3.1 Segment shift</li><li>3.2 Brand and origin shift</li><li>3.3 The Chinese wave</li></ul>
      <div class="go" style="color:#a87200">Open bucket →</div></a>
  </div>
  <div class="down">↓</div>
  <a class="sowhat" href="#concl"><span><b>So what?</b> Three findings, one per bucket, add up to a structural decline, and a five-move plan to 2030.</span>
    <span class="arr">Conclusions & recommendations →</span></a>
</div>""")

html(f"""<div class="legend-inline" style="margin-top:18px">
<span><i style="background:{HONDA}"></i>Honda</span><span><i style="background:{TOYOTA}"></i>Toyota (benchmark)</span>
<span><i style="background:{CHINA}"></i>Chinese brands</span><span><i style="background:{EV_PHEV}"></i>Electrified</span>
<span><i style="background:{GREY}"></i>Rest of market</span>
<span><i style="background:{BAND};border:1px solid {AXIS}"></i>COVID-19 / chip shortage</span></div>""")

# ═════════════════════════════════════════════════════════════════════════════
# BUCKET 01 — FINANCIALS
# ═════════════════════════════════════════════════════════════════════════════
bucket_header(
    "01", "Financials", B1, "Honda's volume: how much was lost, where, and when",
    "INEGI doesn't publish revenue, so units sold are our top-line proxy. Use the slider to compare any two years; "
    "shaded columns mark the COVID-19 shock (2020) and the chip shortage (2021–2022).",
    ["Units by year", "Units by segment", "Nameplates", "Shock periods"], "b1",
)
a, b = year_range("fin_years", B1)
n = b - a
seg_chg = (hseg.loc[b, HSEGS] - hseg.loc[a, HSEGS])
html(f"""
<div class="kpis">
  <div class="kpi" style="border-top-color:{HONDA}"><div class="lbl">Honda units</div>
    <div class="val"><span class="from">{fmt_k(H[a])}</span> → {fmt_k(H[b])}</div>
    <div class="sub">{pct(H[b], H[a]):+.0f}% · market {pct(tot[b], tot[a]):+.0f}%</div></div>
  <div class="kpi" style="border-top-color:{HONDA}"><div class="lbl">Annual growth (CAGR)</div>
    <div class="val">{cagr(H[b], H[a], n):+.1f}%</div><div class="sub">market {cagr(tot[b], tot[a], n):+.1f}% per year · {n} yr</div></div>
  <div class="kpi" style="border-top-color:{HONDA}"><div class="lbl">Market share</div>
    <div class="val"><span class="from">{hs[a]:.1f}%</span> → {hs[b]:.1f}%</div><div class="sub">{hs[b]-hs[a]:+.2f} pp</div></div>
  <div class="kpi" style="border-top-color:{HONDA}"><div class="lbl">Biggest segment swing</div>
    <div class="val">{seg_chg.abs().idxmax()}</div><div class="sub">{seg_chg[seg_chg.abs().idxmax()]:+,.0f} units, {a} → {b}</div></div>
</div>""")

subhead("1.1", "Units sold by year and segment",
        "Each column is Honda's total for the year, split by segment. The two years you selected are at full color.", B1)
op = [1 if y in (a, b) else 0.32 for y in YEARS]
fig = go.Figure()
shock_bands(fig)
for s in HSEGS:
    fig.add_bar(x=YEARS, y=hseg[s].values, name=s, marker=dict(color=SEG_COL[s], opacity=op, line=dict(color="#fff", width=0.8)),
                hovertemplate=f"<b>{s}</b> %{{x}}<br>%{{y:,.0f}} units<extra></extra>")
htot = hseg[HSEGS].sum(axis=1)
for y in YEARS:
    sel = y in (a, b)
    fig.add_annotation(x=y, y=htot[y], text=f"<b>{htot[y]/1e3:.1f}k</b>" if sel else f"{htot[y]/1e3:.0f}k", showarrow=False,
                       yanchor="bottom", yshift=3, font=dict(size=12.5 if sel else 11, color=INK if sel else MUTED))
fig.update_layout(barmode="stack", bargap=0.22)
year_axis(fig, a, b)
fig.update_yaxes(tickformat="~s", range=[0, htot.max() * 1.2])
show(style(fig, 440, f"Honda units sold by segment, {Y0}–{Y1}"))

subhead("1.2", f"{a} vs {b}: which segments and nameplates moved", color=B1)
c1, c2 = st.columns(2)
with c1:
    cmp = pd.DataFrame({"a": hseg.loc[a, HSEGS], "b": hseg.loc[b, HSEGS]})
    cmp = cmp[cmp.max(axis=1) > 0].sort_values("a")
    lab = []
    for s_ in cmp.index:
        va, vb = cmp.loc[s_, "a"], cmp.loc[s_, "b"]
        lab.append("exited" if vb == 0 and va > 0 else ("new" if va == 0 else f"{fmt_k(vb)} ({pct(vb, va):+.0f}%)"))
    fig = go.Figure()
    fig.add_bar(y=cmp.index, x=cmp["a"], orientation="h", name=str(a), marker=dict(color=GREY, cornerradius=4),
                hovertemplate=f"<b>%{{y}}</b> {a}<br>%{{x:,.0f}} units<extra></extra>")
    fig.add_bar(y=cmp.index, x=cmp["b"], orientation="h", name=str(b), marker=dict(color=HONDA, cornerradius=4),
                text=lab, textposition="outside", textfont=dict(color=INK2, size=12), cliponaxis=False,
                hovertemplate=f"<b>%{{y}}</b> {b}<br>%{{x:,.0f}} units<extra></extra>")
    fig.update_layout(barmode="group", bargap=0.25, bargroupgap=0.06)
    fig.update_xaxes(tickformat="~s", showgrid=True, gridcolor=GRID, range=[0, cmp.values.max() * 1.35])
    fig.update_yaxes(showgrid=False, tickfont=dict(color=INK2))
    show(style(fig, 400, f"Honda units by segment, {a} vs {b}"))
with c2:
    npa, npb = nameplates(a), nameplates(b)
    allnp = pd.DataFrame({"a": npa, "b": npb}).fillna(0)
    allnp = allnp[allnp.max(axis=1) >= 300].sort_values("a")
    fig = go.Figure()
    fig.add_bar(y=allnp.index, x=allnp["a"], orientation="h", name=str(a), marker=dict(color=GREY, cornerradius=4),
                hovertemplate=f"<b>%{{y}}</b> {a}<br>%{{x:,.0f}} units<extra></extra>")
    fig.add_bar(y=allnp.index, x=allnp["b"], orientation="h", name=str(b), marker=dict(color=HONDA, cornerradius=4),
                text=["discontinued" if v == 0 else ("new" if allnp.loc[i, "a"] == 0 else "") for i, v in allnp["b"].items()],
                textposition="outside", textfont=dict(color=MUTED, size=11.5), cliponaxis=False,
                hovertemplate=f"<b>%{{y}}</b> {b}<br>%{{x:,.0f}} units<extra></extra>")
    fig.update_layout(barmode="group", bargap=0.25, bargroupgap=0.06)
    fig.update_xaxes(tickformat="~s", showgrid=True, gridcolor=GRID, range=[0, allnp.values.max() * 1.3])
    fig.update_yaxes(showgrid=False, tickfont=dict(color=INK2))
    show(style(fig, 400, f"Honda units by nameplate, {a} vs {b} (≥300 units in either year)"))

gone = [s_ for s_ in HSEGS if hseg.loc[a, s_] > 100 and hseg.loc[b, s_] == 0]
callout(f"<b>{a} → {b}: Honda moved {pct(H[b], H[a]):+.0f}% while the market moved {pct(tot[b], tot[a]):+.0f}%.</b> "
        f"The largest swing came from <b>{seg_chg.abs().idxmax()}</b> ({seg_chg[seg_chg.abs().idxmax()]:+,.0f} units)"
        + (f", and Honda left <b>{', '.join(gone)}</b> entirely" if gone else "")
        + f". Nameplates on sale went from {len(npa)} to {len(npb)}.")

subhead("1.3", "The shock lens: COVID-19 and the chip shortage",
        "Rolling 12-month sales, indexed so that calendar 2019 = 100. Every brand fell; the question is who came back.", B1)
mon = df.groupby(["FECHA", "MARCA"])["UNI_VEH"].sum().unstack(fill_value=0).sort_index()
mon["Market"] = mon.sum(axis=1)
roll = mon[["Market", "Toyota", "Honda"]].rolling(12).sum().dropna()
idx = roll / roll.loc[pd.Timestamp("2019-12-01")] * 100
fig = go.Figure()
fig.add_vrect(x0="2020-03-15", x1="2020-12-31", fillcolor=BAND, line_width=0, layer="below")
fig.add_vrect(x0="2021-01-01", x1="2022-12-31", fillcolor=BAND2, line_width=0, layer="below")
fig.add_annotation(x="2020-08-01", y=1.0, yref="paper", text="<b>COVID-19</b><br>lockdowns", showarrow=False, font=dict(size=11.5, color=MUTED), yanchor="bottom")
fig.add_annotation(x="2022-01-01", y=1.0, yref="paper", text="<b>Chip shortage</b><br>supply rationed", showarrow=False, font=dict(size=11.5, color=MUTED), yanchor="bottom")
fig.add_hline(y=100, line=dict(color=AXIS, width=1))
for col, c, w, nm in (("Market", GREY_D, 2.5, "Total market"), ("Toyota", TOYOTA, 2.5, "Toyota"), ("Honda", HONDA, 3.5, "Honda")):
    fig.add_scatter(x=idx.index, y=idx[col], name=nm, mode="lines", line=dict(color=c, width=w),
                    hovertemplate=f"<b>{nm}</b> · %{{x|%b %Y}}<br>Index %{{y:.0f}} (2019 = 100)<extra></extra>")
    fig.add_annotation(x=idx.index[-1], y=idx[col].iloc[-1], text=f"<b>{nm} {idx[col].iloc[-1]:.0f}</b>", showarrow=False,
                       xanchor="left", xshift=8, font=dict(color=INK, size=12.5))
fig.update_xaxes(range=[idx.index[0], idx.index[-1] + pd.DateOffset(months=16)])
fig.update_layout(hovermode="x unified")
show(style(fig, 430, "Rolling 12-month sales, indexed (calendar 2019 = 100)"))

c1, c2 = st.columns(2)
with c1:
    yy = [y for y in YEARS if y >= 2019]
    g = pd.DataFrame({"Market": tot.pct_change() * 100, "Toyota": T.pct_change() * 100, "Honda": H.pct_change() * 100}).loc[yy]
    fig = go.Figure()
    fig.add_vrect(x0=2019.5, x1=2020.5, fillcolor=BAND, line_width=0, layer="below")
    fig.add_vrect(x0=2020.5, x1=2022.5, fillcolor=BAND2, line_width=0, layer="below")
    for col, c in (("Market", GREY), ("Toyota", TOYOTA), ("Honda", HONDA)):
        fig.add_bar(x=yy, y=g[col], name=col if col != "Market" else "Total market", marker=dict(color=c, cornerradius=4),
                    hovertemplate=f"<b>{col}</b> %{{x}}<br>%{{y:+.1f}}% vs prior year<extra></extra>")
    fig.add_hline(y=0, line=dict(color=AXIS, width=1))
    fig.update_layout(barmode="group", bargap=0.25, bargroupgap=0.06)
    fig.update_xaxes(tickvals=yy)
    fig.update_yaxes(ticksuffix="%")
    show(style(fig, 390, "Year-over-year change in units sold"))
with c2:
    top12 = brand.loc[Y0].sort_values(ascending=False).head(12).index
    rec = (brand.loc[Y1, top12] / brand.loc[2019, top12] * 100).sort_values()
    mk = tot[Y1] / tot[2019] * 100
    fig = go.Figure(go.Bar(
        x=rec.values, y=rec.index, orientation="h", marker=dict(color=[color_for(x) for x in rec.index], cornerradius=4),
        text=[f"<b>{v:.0f}%</b>" if x == "Honda" else f"{v:.0f}%" for x, v in rec.items()], textposition="outside",
        textfont=dict(color=INK2, size=12), cliponaxis=False,
        hovertemplate=f"<b>%{{y}}</b><br>{Y1} sales = %{{x:.0f}}% of 2019<extra></extra>",
    ))
    fig.add_vline(x=100, line=dict(color=AXIS, width=1))
    fig.add_vline(x=mk, line=dict(color=GREY_D, width=1.5, dash="dot"))
    fig.add_annotation(x=mk, y=1.0, yref="paper", text=f"Market {mk:.0f}%", showarrow=False, yanchor="bottom", font=dict(color=MUTED, size=11.5))
    fig.update_xaxes(ticksuffix="%", showgrid=True, gridcolor=GRID, range=[0, rec.max() * 1.15])
    fig.update_yaxes(showgrid=False, tickfont=dict(color=INK2))
    show(style(fig, 390, f"Recovery: {Y1} sales as % of 2019 (top-12 brands of {Y0})", legend=False))

worst = [y for y in YEARS if y >= 2021 and g.loc[y, "Honda"] < 0 and g.loc[y, "Market"] > 0]
peak = idx["Honda"].idxmax()
takeaway(f"<b>Honda's volume problem is structural, not cyclical.</b> Its 12-month sales peaked in {peak:%B %Y}, before the pandemic. "
         f"In 2020 Honda fell {abs(g.loc[2020,'Honda']):.0f}% vs the market's {abs(g.loc[2020,'Market']):.0f}%, but in "
         f"<b>{', '.join(map(str, worst))}</b> the market grew while Honda kept shrinking. Honda is at <b>{rec['Honda']:.0f}%</b> of its 2019 "
         f"volume, the weakest recovery among the 12 largest brands of {Y0} (market: {mk:.0f}%).", B1)

# ═════════════════════════════════════════════════════════════════════════════
# BUCKET 02 — AUTOMOTIVE INDUSTRY
# ═════════════════════════════════════════════════════════════════════════════
bucket_header(
    "02", "Automotive industry", B2, "The market didn't shrink. It moved.",
    "How the Mexican light-vehicle market grew or shrank by year and segment, how fast electrified vehicles took off, "
    "and how Honda stacks up against its closest competitor, Toyota.",
    ["Market size", "Segment growth", "EVs & hybrids", "Toyota head-to-head"], "b2",
)
a, b = year_range("ind_years", B2)
n = b - a
ea, eb = max(a, e0), max(b, e0)  # electrified data starts in e0
pen_a, pen_b = f"{pen[ea]:.1f}%", f"{pen[eb]:.1f}%"
pen_note = f"{ea} → {eb}" + (f" (data starts {e0})" if a < e0 else "")
html(f"""
<div class="kpis">
  <div class="kpi" style="border-top-color:{GREY_D}"><div class="lbl">Market units</div>
    <div class="val"><span class="from">{fmt_k(tot[a])}</span> → {fmt_k(tot[b])}</div>
    <div class="sub">{pct(tot[b], tot[a]):+.0f}% · {cagr(tot[b], tot[a], n):+.1f}% per year</div></div>
  <div class="kpi" style="border-top-color:{SEG_COL['SUVs']}"><div class="lbl">SUVs, share of market</div>
    <div class="val"><span class="from">{mmix.loc[a,'SUVs']:.0f}%</span> → {mmix.loc[b,'SUVs']:.0f}%</div>
    <div class="sub">SUV units {pct(mseg.loc[b,'SUVs'], mseg.loc[a,'SUVs']):+.0f}%</div></div>
  <div class="kpi" style="border-top-color:{EV_PHEV}"><div class="lbl">Electrified share of sales</div>
    <div class="val"><span class="from">{pen_a}</span> → {pen_b}</div>
    <div class="sub">hybrid + plug-in + battery EV · {pen_note}</div></div>
  <div class="kpi" style="border-top-color:{TOYOTA}"><div class="lbl">Toyota sells × Honda</div>
    <div class="val"><span class="from">{T[a]/H[a]:.1f}×</span> → {T[b]/H[b]:.1f}×</div>
    <div class="sub">gap {fmt_k(T[a]-H[a])} → {fmt_k(T[b]-H[b])} units</div></div>
</div>""")

subhead("2.1", "Market size by year and segment",
        "Total light-vehicle units sold in Mexico, split by segment. The selected years are at full color.", B2)
op = [1 if y in (a, b) else 0.32 for y in YEARS]
fig = go.Figure()
shock_bands(fig)
for s in mseg.columns:
    fig.add_bar(x=YEARS, y=mseg[s].values, name=s, marker=dict(color=SEG_COL[s], opacity=op, line=dict(color="#fff", width=0.8)),
                hovertemplate=f"<b>{s}</b> %{{x}}<br>%{{y:,.0f}} units<extra></extra>")
for y in YEARS:
    sel = y in (a, b)
    fig.add_annotation(x=y, y=tot[y], text=f"<b>{tot[y]/1e6:.2f}M</b>" if sel else f"{tot[y]/1e6:.2f}M", showarrow=False,
                       yanchor="bottom", yshift=3, font=dict(size=12.5 if sel else 11, color=INK if sel else MUTED))
fig.update_layout(barmode="stack", bargap=0.22)
year_axis(fig, a, b)
fig.update_yaxes(tickformat="~s", range=[0, tot.max() * 1.2])
show(style(fig, 440, f"Mexico light-vehicle sales by segment, {Y0}–{Y1}"))

c1, c2 = st.columns([1.15, 1])
with c1:
    gr = pd.DataFrame({"a": mseg.loc[a], "b": mseg.loc[b]})
    gr["chg"] = (gr["b"] / gr["a"] - 1) * 100
    gr = gr.sort_values("chg")
    hp = hseg.loc[b].reindex(gr.index).fillna(0) > 100
    ylab = [f"{s_}{'' if hp[s_] else '  ✕'}" for s_ in gr.index]
    fig = go.Figure(go.Bar(
        x=gr["chg"], y=ylab, orientation="h",
        marker=dict(color=[SEG_COL[s_] for s_ in gr.index], cornerradius=4),
        hovertemplate="<b>%{y}</b><br>%{x:+.1f}% units<extra></extra>",
    ))
    # labels always sit to the right of the bar end or the zero line, so they never hit the axis labels
    for yl, (s_, v) in zip(ylab, gr["chg"].items()):
        fig.add_annotation(x=max(v, 0), y=yl, text=f"<b>{v:+.0f}%</b>  {fmt_k(gr.loc[s_,'a'])} → {fmt_k(gr.loc[s_,'b'])}",
                           showarrow=False, xanchor="left", xshift=6, font=dict(color=INK2, size=12))
    fig.add_vline(x=0, line=dict(color=AXIS, width=1))
    lo, hi = min(gr["chg"].min(), 0), max(gr["chg"].max(), 0)
    fig.update_xaxes(ticksuffix="%", showgrid=True, gridcolor=GRID, range=[lo * 1.15 - 5, hi * 1.9 + 60])
    fig.update_yaxes(showgrid=False, tickfont=dict(color=INK2))
    show(style(fig, 380, f"Segment growth, {a} → {b} (✕ = Honda absent in {b})", legend=False))
with c2:
    grow = gr[gr["chg"] > 0]
    absent = [s_ for s_ in gr.index if not hp[s_]]
    absent_pct = mmix.loc[b, absent].sum() if absent else 0
    callout(f"<b>Where the market grew, {a} → {b}:</b> " + (", ".join(f"{s_} ({v:+.0f}%)" for s_, v in grow["chg"].sort_values(ascending=False).items()) or "no segment grew")
            + f". SUVs went from {mmix.loc[a,'SUVs']:.0f}% to <b>{mmix.loc[b,'SUVs']:.0f}%</b> of all sales.", SEG_COL["SUVs"], "#efeefa")
    callout(f"<b>Honda's blind spots.</b> In {b}, Honda sold nothing in <b>{', '.join(absent) if absent else 'no segment'}</b>, "
            f"which together are <b>{absent_pct:.0f}%</b> of the market. Every unit of growth there went to someone else.", HONDA)

subhead("2.2", "Electrified vehicles: the fastest-growing pool of demand",
        f"Hybrid, plug-in hybrid and battery-electric sales ({e0}–{e1}). Labels show electrified share of the total market.", EV_PHEV)
html(f"""
<div class="kpis">
  <div class="kpi" style="border-top-color:{EV_PHEV}"><div class="lbl">Electrified units {e1}</div>
    <div class="val">{ev.loc[e1,'TOTAL_ELECTRIFICADOS']:,.0f}</div><div class="sub">{ev.loc[e1,'TOTAL_ELECTRIFICADOS']/ev.loc[e0,'TOTAL_ELECTRIFICADOS']:.0f}× vs {e0} · {pct(ev.loc[e1,'TOTAL_ELECTRIFICADOS'], ev.loc[e1-1,'TOTAL_ELECTRIFICADOS']):+.0f}% vs {e1-1}</div></div>
  <div class="kpi" style="border-top-color:{EV_HEV}"><div class="lbl">Hybrids (HEV) {e1}</div>
    <div class="val">{ev.loc[e1,'VEH_HIBRIDAS']:,.0f}</div><div class="sub">{ev.loc[e1,'VEH_HIBRIDAS']/ev.loc[e1,'TOTAL_ELECTRIFICADOS']*100:.0f}% of electrified sales</div></div>
  <div class="kpi" style="border-top-color:{EV_BEV}"><div class="lbl">Battery-electric (BEV) {e1}</div>
    <div class="val">{ev.loc[e1,'VEH_ELECTR']:,.0f}</div><div class="sub">from {ev.loc[e0,'VEH_ELECTR']:,.0f} in {e0}</div></div>
  <div class="kpi" style="border-top-color:{HONDA}"><div class="lbl">Electrified market vs Honda</div>
    <div class="val">{ev.loc[e1,'TOTAL_ELECTRIFICADOS']/H[e1]:.1f}×</div><div class="sub">Electrified sales exceed all of Honda's sales</div></div>
</div>""")
c1, c2 = st.columns([1.35, 1])
with c1:
    evy = [int(y) for y in ev.index]
    ev_op = [1 if (y in (a, b) or not ({a, b} & set(evy))) else 0.4 for y in evy]
    fig = go.Figure()
    shock_bands(fig, labels=False)
    for col, nm, c in (("VEH_HIBRIDAS", "Hybrid (HEV)", EV_HEV), ("VEH_HIBRIDAS_PLUGIN", "Plug-in hybrid (PHEV)", EV_PHEV), ("VEH_ELECTR", "Battery electric (BEV)", EV_BEV)):
        fig.add_bar(x=evy, y=ev[col], name=nm, marker=dict(color=c, opacity=ev_op, line=dict(color="#fff", width=1.2)),
                    hovertemplate=f"<b>{nm}</b> %{{x}}<br>%{{y:,.0f}} units<extra></extra>")
    for y in evy:
        fig.add_annotation(x=y, y=ev.loc[y, "TOTAL_ELECTRIFICADOS"], text=f"{pen[y]:.1f}%", showarrow=False, yanchor="bottom", yshift=3,
                           font=dict(size=11.5, color=INK if y in (a, b) else INK2))
    fig.update_layout(barmode="stack", bargap=0.3)
    fig.update_xaxes(tickvals=evy, ticktext=[f"<b>{y}</b>" if y in (a, b) else str(y) for y in evy], tickfont=dict(color=INK2))
    fig.update_yaxes(tickformat="~s", range=[0, ev["TOTAL_ELECTRIFICADOS"].max() * 1.15])
    show(style(fig, 420, "Electrified vehicle sales by powertrain"))
with c2:
    hyb = df[(df["MARCA"] == "Honda") & df["MODELO"].str.contains("hybrid|híbrid|hev", case=False, regex=True)]
    hyb_u, hyb_last = int(hyb["UNI_VEH"].sum()), int(hyb.loc[hyb["UNI_VEH"] > 0, "ANIO"].max())
    callout(f"<b>Where is Honda?</b> In INEGI's model-level registry, Honda's only nameplate explicitly reported as electrified is the "
            f"<b>Insight Hybrid: {hyb_u:,} units in total</b>, last sold in {hyb_last}. Hybrids alone sold "
            f"<b>{ev.loc[e1,'VEH_HIBRIDAS']:,.0f}</b> units in {e1} ({pct(ev.loc[e1,'VEH_HIBRIDAS'], ev.loc[e1-1,'VEH_HIBRIDAS']):+.0f}% YoY). "
            f"Hybrids are Honda's core global competence (e:HEV), and the brand hasn't turned that into volume in Mexico.", EV_BEV, "#eaf6f0")
    callout("<b>Hybrids lead, BEVs follow.</b> Mexican buyers adopt electrification through hybrids first. That's the bridge Honda "
            "can cross quickly, without waiting for charging infrastructure.", EV_PHEV, "#eaf6f0")

subhead("2.3", "Closest competitor: Toyota",
        f"Two Japanese brands with similar reputations for quality. Switch the snapshot between the two years selected above ({a} and {b}).", TOYOTA)
pick = st.segmented_control("Snapshot year", [str(a), str(b)], default=str(b), key=f"snap_{a}_{b}")
yr = int(pick) if pick else b
sv = seg_view(yr)
html(f"""
<div class="vs">
  <div class="side" style="border-top:5px solid {HONDA}">
    <div class="brand" style="color:{HONDA}">Honda · {yr}</div>
    <div class="units">{H[yr]:,.0f}</div>
    <div class="meta">units sold · <b>{hs[yr]:.1f}%</b> market share · rank #{rank.loc[yr,'Honda']}</div>
  </div>
  <div class="mid"><div class="x">Toyota sells</div><div class="ratio">{T[yr]/H[yr]:.1f}×</div>
    <div class="cap">Gap: {(T[yr]-H[yr]):,.0f} units</div></div>
  <div class="side" style="border-top:5px solid {TOYOTA}">
    <div class="brand" style="color:{TOYOTA}">Toyota · {yr}</div>
    <div class="units">{T[yr]:,.0f}</div>
    <div class="meta">units sold · <b>{ts[yr]:.1f}%</b> market share · rank #{rank.loc[yr,'Toyota']}</div>
  </div>
</div>""")

segs = sv[(sv["Honda_u"] + sv["Toyota_u"]) > 50]
labels = [f"{s_}  ·  {sv.loc[s_,'mkt_pct']:.0f}% of mkt" for s_ in segs.index]
c1, c2 = st.columns(2)
with c1:
    fig = go.Figure()
    fig.add_bar(y=labels, x=-segs["Honda_u"], orientation="h", name="Honda", marker=dict(color=HONDA, cornerradius=4),
                customdata=segs["Honda_u"], text=[f"{v/1e3:.1f}k" if v >= 50 else "none" for v in segs["Honda_u"]], textposition="outside",
                textfont=dict(color=INK2, size=12), cliponaxis=False,
                hovertemplate="<b>Honda · %{y}</b><br>%{customdata:,.0f} units<extra></extra>")
    fig.add_bar(y=labels, x=segs["Toyota_u"], orientation="h", name="Toyota", marker=dict(color=TOYOTA, cornerradius=4),
                text=[f"{v/1e3:.1f}k" if v >= 50 else "none" for v in segs["Toyota_u"]], textposition="outside",
                textfont=dict(color=INK2, size=12), cliponaxis=False,
                hovertemplate="<b>Toyota · %{y}</b><br>%{x:,.0f} units<extra></extra>")
    m = max(segs["Honda_u"].max(), segs["Toyota_u"].max()) * 1.3
    fig.update_layout(barmode="relative", bargap=0.35)
    fig.update_xaxes(range=[-m, m], showticklabels=False, showgrid=False, zeroline=True, zerolinecolor=AXIS)
    fig.update_yaxes(autorange="reversed", showgrid=False, tickfont=dict(color=INK2))
    show(style(fig, 400, f"Units sold by segment, {yr}"))
with c2:
    fig = go.Figure()
    for b_, c in (("Honda", HONDA), ("Toyota", TOYOTA)):
        fig.add_bar(y=list(segs.index), x=segs[f"{b_}_s"], orientation="h", name=b_, marker=dict(color=c, cornerradius=4),
                    text=[f"{v:.1f}%" for v in segs[f"{b_}_s"]], textposition="outside", textfont=dict(color=INK2, size=12),
                    cliponaxis=False, hovertemplate=f"<b>{b_} · %{{y}}</b><br>%{{x:.1f}}% of the segment<extra></extra>")
    fig.update_layout(barmode="group", bargap=0.3, bargroupgap=0.08)
    fig.update_xaxes(ticksuffix="%", showgrid=True, gridcolor=GRID, range=[0, max(segs["Honda_s"].max(), segs["Toyota_s"].max()) * 1.25])
    fig.update_yaxes(autorange="reversed", showgrid=False, tickfont=dict(color=INK2))
    show(style(fig, 400, f"Share of each segment, {yr} (brand units ÷ segment units)"))

both = segs[(segs["Honda_u"] > 50) | (segs["Toyota_u"] > 50)]
t_leads = int((both["Toyota_u"] > both["Honda_u"]).sum())
h_leads = [s_ for s_ in both.index if both.loc[s_, "Honda_u"] > both.loc[s_, "Toyota_u"]]
callout(f"<b>{yr}: Toyota out-sells Honda in {t_leads} of {len(both)} segments where either competes"
        + (f"; Honda still leads in {', '.join(h_leads)}" if h_leads else "") + ".</b> "
        f"In SUVs, the largest segment ({sv.loc['SUVs','mkt_pct']:.0f}% of the market), Honda holds {sv.loc['SUVs','Honda_s']:.1f}% "
        f"vs Toyota's {sv.loc['SUVs','Toyota_s']:.1f}%. Toyota also owns a pickup franchise (Hilux, Tacoma) where Honda has no product.",
        TOYOTA, "#eef4fc")

fig = go.Figure()
shock_bands(fig)
for b_, s_, c, w in (("Toyota", ts, TOYOTA, 3), ("Honda", hs, HONDA, 3.5)):
    fig.add_scatter(x=YEARS, y=s_.values, name=b_, mode="lines+markers", line=dict(color=c, width=w),
                    marker=dict(size=[11 if y in (a, b) else 7 for y in YEARS], line=dict(color="#fff", width=2)),
                    hovertemplate=f"<b>{b_}</b> %{{x}}<br>%{{y:.2f}}% share<extra></extra>")
    fig.add_annotation(x=Y1, y=s_[Y1], text=f"<b>{b_} {s_[Y1]:.1f}%</b>", showarrow=False, xanchor="left", xshift=10, font=dict(color=INK, size=13))
    fig.add_annotation(x=Y0, y=s_[Y0], text=f"{s_[Y0]:.1f}%", showarrow=False, xanchor="right", xshift=-10, font=dict(color=INK2, size=12.5))
year_axis(fig, a, b)
fig.update_xaxes(range=[Y0 - 0.6, Y1 + 1.1])
fig.update_yaxes(ticksuffix="%", rangemode="tozero", range=[0, ts.max() * 1.2])
show(style(fig, 380, "The scissors: market share, Honda vs Toyota"))

takeaway(f"<b>The industry grew {pct(tot[Y1], tot[Y0]):+.0f}% over the decade, but in places Honda doesn't play.</b> SUVs went from "
         f"{mmix.loc[Y0,'SUVs']:.0f}% to {mmix.loc[Y1,'SUVs']:.0f}% of the market, pickups and hybrids grew, and electrified sales reached "
         f"{pen1:.1f}% of the market. Toyota, with full-line coverage and hybrids, went from {T[Y0]/H[Y0]:.1f}× to "
         f"<b>{T[Y1]/H[Y1]:.1f}×</b> Honda's volume.", B2)

# ═════════════════════════════════════════════════════════════════════════════
# BUCKET 03 — CUSTOMER PREFERENCES
# ═════════════════════════════════════════════════════════════════════════════
bucket_header(
    "03", "Customer preferences", B3, "Buyers changed what, and whom, they buy",
    "Three lenses on the Mexican buyer: which segments they choose, which brand origins they trust, and which brands won "
    "or lost their wallet. Honda's slice of each tells us whether its offer still matches demand.",
    ["Segment mix", "Brand origin", "Brand winners & losers", "Chinese brands"], "b3",
)
a, b = year_range("cus_years", "#a87200")
html(f"""
<div class="kpis">
  <div class="kpi" style="border-top-color:{SEG_COL['SUVs']}"><div class="lbl">SUVs, share of purchases</div>
    <div class="val"><span class="from">{mmix.loc[a,'SUVs']:.0f}%</span> → {mmix.loc[b,'SUVs']:.0f}%</div>
    <div class="sub">{mmix.loc[b,'SUVs']-mmix.loc[a,'SUVs']:+.1f} pp</div></div>
  <div class="kpi" style="border-top-color:{CHINA}"><div class="lbl">Chinese brands' share</div>
    <div class="val"><span class="from">{china_s[a]:.1f}%</span> → {china_s[b]:.1f}%</div>
    <div class="sub">{fmt_k(china_u[a]) if china_u[a] else '0'} → {fmt_k(china_u[b])} units</div></div>
  <div class="kpi" style="border-top-color:{PEER}"><div class="lbl">Other Japanese brands' share</div>
    <div class="val"><span class="from">{jap_ex[a]:.1f}%</span> → {jap_ex[b]:.1f}%</div>
    <div class="sub">{jap_ex[b]-jap_ex[a]:+.1f} pp (Japanese brands excluding Honda)</div></div>
  <div class="kpi" style="border-top-color:{HONDA}"><div class="lbl">Honda's slice of Japanese sales</div>
    <div class="val"><span class="from">{honda_in_jp[a]:.1f}%</span> → {honda_in_jp[b]:.1f}%</div>
    <div class="sub">1 in {100/honda_in_jp[a]:.0f} → 1 in {100/honda_in_jp[b]:.0f} Japanese-brand buyers</div></div>
</div>""")

subhead("3.1", "Segment shift: buyers moved to SUVs",
        "Market mix by segment (left) and how Honda's own mix compares to where demand is (right).", SEG_COL["SUVs"])
c1, c2 = st.columns([1.3, 1])
with c1:
    fig = go.Figure()
    for s_ in mmix.columns:  # SUVs first, so they form the bottom band
        fig.add_scatter(x=YEARS, y=mmix[s_].values, name=s_, mode="lines", stackgroup="one", line=dict(width=0.6, color="#fff"),
                        fillcolor=SEG_COL[s_], hovertemplate=f"<b>{s_}</b> %{{x}}<br>%{{y:.1f}}% of units<extra></extra>")
    for y in (a, b):
        fig.add_vline(x=y, line=dict(color=INK, width=1.5, dash="dot"))
        fig.add_annotation(x=y, y=mmix.loc[y, "SUVs"] / 2, text=f"<b>SUVs<br>{mmix.loc[y,'SUVs']:.0f}%</b>", showarrow=False,
                           font=dict(color="#fff", size=12), xanchor="left" if y == a else "right", xshift=5 if y == a else -5)
    year_axis(fig, a, b, compact=True)
    fig.update_xaxes(range=[Y0, Y1])
    fig.update_yaxes(ticksuffix="%", range=[0, 100])
    style(fig, 440, "Market mix by segment (% of units)")
    fig.update_layout(legend=dict(orientation="h", x=0, y=-0.08, yanchor="top"), margin=dict(t=44, b=8))
    show(fig)
with c2:
    mx = pd.DataFrame({"Market": mmix.loc[b], "Honda": hmix.loc[b]}).sort_values("Market")
    fig = go.Figure()
    fig.add_bar(y=mx.index, x=mx["Market"], orientation="h", name=f"Market {b}", marker=dict(color=GREY, cornerradius=4),
                text=[f"{v:.0f}%" for v in mx["Market"]], textposition="outside", textfont=dict(color=MUTED, size=11.5), cliponaxis=False,
                hovertemplate="<b>%{y}</b><br>%{x:.1f}% of market units<extra></extra>")
    fig.add_bar(y=mx.index, x=mx["Honda"], orientation="h", name=f"Honda {b}", marker=dict(color=HONDA, cornerradius=4),
                text=[f"<b>{v:.0f}%</b>" if v >= 0.5 else "0%" for v in mx["Honda"]], textposition="outside",
                textfont=dict(color=INK, size=11.5), cliponaxis=False,
                hovertemplate="<b>%{y}</b><br>%{x:.1f}% of Honda units<extra></extra>")
    fig.update_layout(barmode="group", bargap=0.25, bargroupgap=0.06)
    fig.update_xaxes(ticksuffix="%", showgrid=True, gridcolor=GRID, range=[0, mx.values.max() * 1.2])
    fig.update_yaxes(showgrid=False, tickfont=dict(color=INK2))
    show(style(fig, 440, f"Where demand is vs where Honda sells, {b}"))
zero = [s_ for s_ in mx.index if mx.loc[s_, "Honda"] < 0.5 and mx.loc[s_, "Market"] >= 5]
callout(f"<b>In {b}, {mmix.loc[b,'SUVs']:.0f}% of buyers chose an SUV</b> (vs {mmix.loc[a,'SUVs']:.0f}% in {a}). "
        f"Honda's mix is concentrated in SUVs ({hmix.loc[b,'SUVs']:.0f}%) and compacts ({hmix.loc[b,'Compact']:.0f}%)"
        + (f", and it offers nothing to the <b>{mx.loc[zero,'Market'].sum():.0f}%</b> of buyers who chose {' or '.join(zero).lower()}." if zero else "."),
        SEG_COL["SUVs"], "#efeefa")

subhead("3.2", "Brand and origin shift: who won the buyer's wallet",
        "Market share by brand origin (where the brand is from, not where the car is built), with Honda split out of the Japanese group.", B3)
c1, c2 = st.columns(2)
with c1:
    lines = pd.DataFrame({
        "Japanese (ex-Honda)": jap_ex, "American": grp_s["American"], "European & other": grp_s["European & other"],
        "Korean": grp_s["Korean"], "Chinese": grp_s["Chinese"], "Honda": hs,
    })
    lc = {"Japanese (ex-Honda)": PEER, "American": GREY, "European & other": GREY, "Korean": GREY, "Chinese": CHINA, "Honda": HONDA}
    lw = {"Honda": 3.5, "Chinese": 3.5, "Japanese (ex-Honda)": 2.5}
    fig = go.Figure()
    shock_bands(fig, labels=False)
    for col in lines.columns:
        fig.add_scatter(x=YEARS, y=lines[col].values, name=col, mode="lines", line=dict(color=lc[col], width=lw.get(col, 2)),
                        hovertemplate=f"<b>{col}</b> %{{x}}<br>%{{y:.1f}}% share<extra></extra>", showlegend=False)
    # direct labels, nudged apart so they don't collide
    ends = lines.loc[Y1].sort_values()
    placed = []
    for col, v in ends.items():
        yv = v
        for p in placed:
            if abs(yv - p) < 2.2:
                yv = p + 2.2
        placed.append(yv)
        short = {"Japanese (ex-Honda)": "Japanese ex-Honda", "European & other": "European+"}.get(col, col)
        fig.add_annotation(x=Y1, y=yv, text=f"<b>{short}</b> {v:.1f}%", showarrow=False, xanchor="left", xshift=8,
                           font=dict(color=lc[col] if col in ("Honda", "Chinese", "Japanese (ex-Honda)") else MUTED, size=11.5))
    for y in (a, b):
        fig.add_vline(x=y, line=dict(color=AXIS, width=1, dash="dot"))
    year_axis(fig, a, b, compact=True)
    fig.update_xaxes(range=[Y0 - 0.3, Y1 + 5])
    fig.update_yaxes(ticksuffix="%", rangemode="tozero")
    show(style(fig, 430, "Market share by brand origin", legend=False))
with c2:
    delta = (share.loc[b] - share.loc[a]).drop(CHINESE)
    delta["Chinese brands (all)"] = china_s[b] - china_s[a]
    dd = delta[delta.abs() >= 0.5]
    if len(dd) < 5:
        dd = delta.reindex(delta.abs().sort_values(ascending=False).index[:10])
    if "Honda" not in dd.index:
        dd["Honda"] = delta["Honda"]
    dd = dd.sort_values()
    fig = go.Figure(go.Bar(
        x=dd.values, y=dd.index, orientation="h", marker=dict(color=[color_for(x) for x in dd.index], cornerradius=4),
        text=[f"<b>{v:+.1f} pp</b>" if x in ("Honda", "Toyota", "Chinese brands (all)") else f"{v:+.1f} pp" for x, v in dd.items()],
        textposition="outside", textfont=dict(color=INK2, size=12), cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>%{x:+.2f} pp share change<extra></extra>",
    ))
    fig.add_vline(x=0, line=dict(color=AXIS, width=1))
    rng = max(abs(dd.min()), abs(dd.max()))
    fig.update_xaxes(ticksuffix=" pp", showgrid=True, gridcolor=GRID, range=[min(dd.min(), 0) - rng * 0.35, max(dd.max(), 0) + rng * 0.35])
    fig.update_yaxes(showgrid=False, tickfont=dict(color=INK2))
    show(style(fig, 430, f"Who won and lost share, {a} → {b}", legend=False))

big = brand.loc[a].sort_values(ascending=False).head(12).index
rel = (share.loc[b, big] / share.loc[a, big] - 1) * 100
h_rank_rel = int(rel.rank(method="min").get("Honda", 0)) if "Honda" in rel else None
jd, hd = jap_ex[b] - jap_ex[a], hs[b] - hs[a]
head = ("Buyers didn't abandon Japanese brands; they abandoned Honda." if hd < 0 and abs(jd) < abs(hd)
        else f"Honda's share moved {hd:+.1f} pp vs {jd:+.1f} pp for the other Japanese brands.")
callout(f"<b>{head}</b> Between {a} and {b}, the other Japanese brands moved {jap_ex[a]:.1f}% → {jap_ex[b]:.1f}% of the market "
        f"({jd:+.1f} pp) while Honda moved {hs[a]:.1f}% → {hs[b]:.1f}% ({hd:+.1f} pp). Honda's slice of Japanese-brand sales went from "
        f"{honda_in_jp[a]:.1f}% to <b>{honda_in_jp[b]:.1f}%</b>. Among the 12 largest brands of {a}, Honda's relative share change "
        f"({rel.get('Honda', float('nan')):+.0f}%) ranks #{h_rank_rel} worst.", PEER, "#f3f2ee")

subhead("3.3", "The Chinese wave", "From zero to a double-digit share of the market in a few years.", B3)
c1, c2 = st.columns(2)
with c1:
    fig = go.Figure()
    shock_bands(fig, labels=False)
    fig.add_scatter(x=YEARS, y=china_u.values, name="All Chinese brands", mode="lines+markers", line=dict(color=CHINA, width=3.5),
                    marker=dict(size=8, line=dict(color="#fff", width=2)), hovertemplate="<b>Chinese brands</b> %{x}<br>%{y:,.0f} units<extra></extra>")
    fig.add_scatter(x=YEARS, y=H.values, name="Honda", mode="lines+markers", line=dict(color=HONDA, width=3.5),
                    marker=dict(size=8, line=dict(color="#fff", width=2)), hovertemplate="<b>Honda</b> %{x}<br>%{y:,.0f} units<extra></extra>")
    cross = next((y for y in YEARS if china_u[y] > H[y]), None)
    if cross:
        fig.add_annotation(x=cross, y=china_u[cross], text=f"<b>{cross}</b>: Chinese brands<br>overtake Honda", ax=-80, ay=-46,
                           arrowcolor=MUTED, arrowwidth=1, font=dict(size=12, color=INK))
    fig.add_annotation(x=Y1, y=china_u[Y1], text=f"<b>{china_u[Y1]/1e3:.0f}k</b>", showarrow=False, xanchor="left", xshift=8, font=dict(color=INK))
    fig.add_annotation(x=Y1, y=H[Y1], text=f"<b>{H[Y1]/1e3:.0f}k</b>", showarrow=False, xanchor="left", xshift=8, font=dict(color=INK))
    year_axis(fig, a, b, compact=True)
    fig.update_xaxes(range=[Y0 - 0.4, Y1 + 0.8])
    fig.update_yaxes(tickformat="~s")
    show(style(fig, 400, "Units sold: Chinese brands (combined) vs Honda"))
with c2:
    cb = brand.loc[b, CHINESE].sort_values(ascending=False)
    cb = cb[cb > 0].head(7)
    if len(cb):
        comp = pd.concat([cb, pd.Series({"Honda": H[b]})]).sort_values()
        fig = go.Figure(go.Bar(
            x=comp.values, y=comp.index, orientation="h", marker=dict(color=[color_for(x) for x in comp.index], cornerradius=4),
            text=[f"<b>{v/1e3:.1f}k</b>" if x == "Honda" else f"{v/1e3:.1f}k" for x, v in comp.items()],
            textposition="outside", textfont=dict(color=INK2, size=12), cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>%{x:,.0f} units<extra></extra>",
        ))
        fig.update_xaxes(tickformat="~s", showgrid=True, gridcolor=GRID, range=[0, comp.max() * 1.2])
        fig.update_yaxes(showgrid=False, tickfont=dict(color=INK2))
        show(style(fig, 400, f"{b}: Honda vs the largest Chinese brands", legend=False))
    else:
        callout(f"<b>No Chinese brand reported sales in {b}.</b> Move the right-hand slider handle to a later year to see the wave.", CHINA, "#fdf5e3")
mg_over = next((y for y in YEARS if brand.loc[y, "MG Motor"] > H[y]), None)
st.caption("Note: INEGI's RAIAVL only covers brands that report to it. Some players (e.g., BYD, Tesla) don't report, so Chinese-brand "
           "and EV totals here are likely understated. Electrified data is reported by state, not by brand.")

takeaway(f"<b>The buyer Honda used to win now has better-matched options.</b> {mmix.loc[Y1,'SUVs']:.0f}% of buyers want an SUV, "
         f"and Chinese brands went from 0% to <b>{china_s[Y1]:.1f}%</b> share by offering more equipment for less money"
         + (f"; MG Motor alone out-sold Honda by {mg_over}" if mg_over else "")
         + f". The other Japanese brands held their ground ({jap_ex[Y0]:.1f}% → {jap_ex[Y1]:.1f}%), while Honda's slice of "
         f"Japanese-brand sales halved, from {honda_in_jp[Y0]:.0f}% to {honda_in_jp[Y1]:.0f}%.", B3)

# ═════════════════════════════════════════════════════════════════════════════
# CONCLUSIONS & RECOMMENDATIONS
# ═════════════════════════════════════════════════════════════════════════════
chapter(
    "Conclusions & recommendations", "The road back: a 2030 plan for Honda Mexico",
    "Each bucket points to the same diagnosis: Honda's decline is structural, not cyclical, and the fix is about portfolio, "
    "powertrain and value, not just marketing. The good news is that the demand Honda lost is growing, and Honda already has the technology.",
    anchor="concl",
)

subhead("Synthesis", f"Honda {Y0} vs Honda {Y1} at a glance", color=HONDA)


def sc(label, a_, b_, badge, kind="bad"):
    return (f'<div class="sc"><div class="lbl">{label}</div><div class="row"><span class="y">{Y0}</span><span class="a">{a_}</span>'
            f'<span class="arrow">→</span><span class="y">{Y1}</span><span class="b">{b_}</span></div>'
            f'<span class="badge {kind}">{badge}</span></div>')


segs_present = lambda s_: int((s_["Honda_u"] > 100).sum())
cards = [
    sc("Units sold", f"{H[Y0]/1e3:.1f}k", f"{H[Y1]/1e3:.1f}k", f"{pct(H[Y1], H[Y0]):+.0f}% · market {pct(tot[Y1], tot[Y0]):+.0f}%"),
    sc("Market share", f"{hs[Y0]:.1f}%", f"{hs[Y1]:.1f}%", f"{hs[Y1]-hs[Y0]:+.1f} pp"),
    sc("National ranking", f"#{rank.loc[Y0,'Honda']}", f"#{rank.loc[Y1,'Honda']}", f"{rank.loc[Y1,'Honda']-rank.loc[Y0,'Honda']} places lost"),
    sc("Gap to Toyota (units)", f"{(T[Y0]-H[Y0])/1e3:.0f}k", f"{(T[Y1]-H[Y1])/1e3:.0f}k", f"{(T[Y1]-H[Y1])/(T[Y0]-H[Y0]):.1f}× wider"),
    sc("Share of SUV segment", f"{seg0.loc['SUVs','Honda_s']:.1f}%", f"{seg1.loc['SUVs','Honda_s']:.1f}%",
       f"{seg1.loc['SUVs','Honda_s']-seg0.loc['SUVs','Honda_s']:+.1f} pp in the largest segment"),
    sc("Nameplates on sale", f"{len(np0)}", f"{len(np1)}", f"{len(np1)-len(np0)} models"),
    sc("Segments with presence", f"{segs_present(seg0)}", f"{segs_present(seg1)}", "exited Subcompact"),
    sc("Made in Mexico (% of sales)", f"{origin_share(Y0):.0f}%", f"{origin_share(Y1):.0f}%", f"{origin_share(Y1)-origin_share(Y0):+.0f} pp"),
    sc("Best-selling nameplate", f"{np0.index[0]} {np0.iloc[0]/1e3:.1f}k", f"{np1.index[0]} {np1.iloc[0]/1e3:.1f}k",
       f"{pct(np1.iloc[0], np0.iloc[0]):+.0f}%"),
]
html('<div class="score">' + "".join(cards) + "</div>")

with st.expander("View the underlying data"):
    t = pd.DataFrame({
        "Market units": tot, "Honda units": H, "Toyota units": T, "Chinese brands units": china_u,
        "Honda share %": hs.round(2), "Toyota share %": ts.round(2), "Chinese share %": china_s.round(2),
        "Honda rank": rank["Honda"], "Toyota rank": rank["Toyota"], "SUV % of market": mmix["SUVs"].round(1),
        "Honda % of Japanese-brand sales": honda_in_jp.round(1),
    })
    st.dataframe(t, width="stretch")

subhead("Conclusions", "One finding per bucket", color=HONDA)
html(f"""
<div class="concl">
  <div class="cc" style="border-top-color:{B1}"><div class="n" style="color:#ff6b78">01 · Financials</div><h4>Volume halved, and it's structural</h4>
    <p>Honda went from {H[Y0]/1e3:.0f}k to {H[Y1]/1e3:.0f}k units. The market recovered to {tot[Y1]/tot[2019]*100:.0f}% of 2019 and Toyota to
    {T[Y1]/T[2019]*100:.0f}%; Honda is at {H[Y1]/H[2019]*100:.0f}%. COVID and chips were the trigger, not the cause.</p></div>
  <div class="cc" style="border-top-color:{B2}"><div class="n" style="color:#7fb1f0">02 · Automotive industry</div><h4>The market grew where Honda isn't</h4>
    <p>SUVs rose to {mmix.loc[Y1,'SUVs']:.0f}% of sales, electrified vehicles to {pen1:.1f}%, and pickups and subcompacts remain
    {seg1.loc['Subcompact','mkt_pct']+seg1.loc['Pickups','mkt_pct']:.0f}% of the market with zero Honda presence. Toyota now sells {T[Y1]/H[Y1]:.1f}× Honda.</p></div>
  <div class="cc" style="border-top-color:{B3}"><div class="n" style="color:{CHINA}">03 · Customer preferences</div><h4>Squeezed from both sides</h4>
    <p>Above Honda, Toyota offers hybrids and full-line coverage. Below it, Chinese brands ({china_s[Y1]:.1f}% share) win on price and features.
    Other Japanese brands held {jap_ex[Y1]:.0f}% of the market, but Honda's slice of Japanese sales fell from {honda_in_jp[Y0]:.0f}% to {honda_in_jp[Y1]:.0f}%.</p></div>
</div>""")

subhead("Ambition", "Where this could go by 2030", color=HONDA)
mkt2030 = tot[Y1] * (1.01 ** 5)
target_s = 4.5
sq = hs[Y1] * (hs[Y1] / hs[Y1 - 2]) ** (5 / 2)  # status quo: extend 2023–25 share trend
c1, c2 = st.columns([1.35, 1])
with c1:
    fut = list(range(Y1, 2031))
    tgt = [hs[Y1] + (target_s - hs[Y1]) * ((y - Y1) / 5) ** 1.15 for y in fut]
    sqp = [hs[Y1] * (sq / hs[Y1]) ** ((y - Y1) / 5) for y in fut]
    fig = go.Figure()
    fig.add_hrect(y0=4.5, y1=5.5, fillcolor="rgba(214,40,57,0.06)", line_width=0, layer="below")
    fig.add_annotation(x=2030, y=5.5, text=f"{Y0} level ({hs[Y0]:.1f}%)", showarrow=False, yanchor="bottom", xanchor="right",
                       font=dict(color=MUTED, size=11.5))
    fig.add_scatter(x=YEARS, y=hs.values, name="Honda actual", mode="lines+markers", line=dict(color=HONDA, width=3.5),
                    marker=dict(size=8, line=dict(color="#fff", width=2)), hovertemplate="<b>Honda</b> %{x}<br>%{y:.2f}%<extra></extra>")
    fig.add_scatter(x=fut, y=tgt, name="2030 plan", mode="lines", line=dict(color=HONDA, width=3, dash="dot"),
                    hovertemplate="<b>Plan</b> %{x}<br>%{y:.1f}%<extra></extra>")
    fig.add_scatter(x=fut, y=sqp, name="Status quo (2023–25 trend)", mode="lines", line=dict(color=GREY_D, width=2, dash="dot"),
                    hovertemplate="<b>Status quo</b> %{x}<br>%{y:.1f}%<extra></extra>")
    fig.add_annotation(x=2030, y=target_s, text=f"<b>{target_s:.1f}%</b>", showarrow=False, xanchor="left", xshift=8, font=dict(color=INK))
    fig.add_annotation(x=2030, y=sq, text=f"{sq:.1f}%", showarrow=False, xanchor="left", xshift=8, font=dict(color=INK2))
    fig.add_annotation(x=Y1, y=hs[Y1], text=f"<b>{hs[Y1]:.1f}%</b>", showarrow=False, yanchor="top", yshift=-10, font=dict(color=INK))
    fig.update_xaxes(tickvals=[2015, 2017, 2019, 2021, 2023, 2025, 2027, 2030], range=[Y0 - 0.4, 2031.2])
    fig.update_yaxes(ticksuffix="%", range=[0, 6.6])
    show(style(fig, 400, "Market share path: actual, status quo, and 2030 ambition (illustrative)"))
with c2:
    html(f"""
    <div class="kpis" style="grid-template-columns:1fr">
      <div class="kpi" style="border-top-color:{HONDA}"><div class="lbl">2030 ambition</div>
        <div class="val">{target_s:.1f}% share</div><div class="sub">≈ {target_s/100*mkt2030/1e3:.0f}k units on a ~{mkt2030/1e6:.2f}M market (1% annual growth assumed): back to 2019 volume</div></div>
      <div class="kpi" style="border-top-color:{GREY_D}"><div class="lbl">Cost of doing nothing</div>
        <div class="val">{sq:.1f}% share</div><div class="sub">≈ {sq/100*mkt2030/1e3:.0f}k units. Honda drifts toward niche status.</div></div>
      <div class="kpi" style="border-top-color:{EV_PHEV}"><div class="lbl">Electrified mix target</div>
        <div class="val">≥ 40%</div><div class="sub">of Honda Mexico sales hybrid or plug-in by 2030</div></div>
    </div>""")

subhead("Recommendations", "Five moves, each tied to the bucket it fixes", color=HONDA)
recs = [
    ("Go hybrid-first, now", f"Electrified sales reached {ev.loc[e1,'TOTAL_ELECTRIFICADOS']/1e3:.0f}k units ({pen1:.1f}% of the market) and hybrids grew "
     f"{pct(ev.loc[e1,'VEH_HIBRIDAS'], ev.loc[e1-1,'VEH_HIBRIDAS']):.0f}% last year. Honda's e:HEV system is a proven global asset. Make it available on CR-V, HR-V and City, "
     "priced within ~10% of the gasoline version, and lead the marketing with fuel savings.",
     "Hybrids are the fastest-growing pool of demand", "≥ 40% electrified mix · 25k+ hybrid units", [("bk2", "02 Industry")]),
    ("Win back the SUV core", f"SUVs are now {seg1.loc['SUVs','mkt_pct']:.0f}% of the market. Honda went from the #{suv_rank0} SUV brand "
     f"({seg0.loc['SUVs','Honda_s']:.1f}% share) to {seg1.loc['SUVs','Honda_s']:.1f}%. Sharpen HR-V/BR-V entry trims against Chinese C-SUVs and add a family 3-row SUV between CR-V and Pilot.",
     f"Honda SUV share {seg0.loc['SUVs','Honda_s']:.1f}% → {seg1.loc['SUVs','Honda_s']:.1f}%", "8% SUV segment share",
     [("bk3", "03 Customer"), ("bk1", "01 Financials")]),
    ("Re-enter the entry segment, made in Celaya", f"Subcompacts are still {seg1.loc['Subcompact','mkt_pct']:.0f}% of the market and Honda has had zero presence since the Fit's exit. "
     f"Locally built share of Honda sales fell from {origin_share(Y0):.0f}% to {origin_share(Y1):.0f}%. Build an affordable, USMCA-qualified model in Celaya "
     "to win on cost and stay protected from tariffs on non-FTA imports.",
     f"0% share in a segment worth {seg1.loc['Subcompact','total']/1e3:.0f}k units", "15k entry units · ≥ 50% made-in-Mexico",
     [("bk1", "01 Financials"), ("bk2", "02 Industry")]),
    ("Out-value the Chinese challengers", f"Chinese brands took {china_s[Y1]:.1f}% share in a few years"
     + (" and MG alone out-sells Honda" if mg_over else "") + ". Don't fight on sticker price; fight on "
     "total cost of ownership: longer warranties, guaranteed resale value, Honda Financial Services rates, connected-car features, and fixed-price maintenance.",
     f"{china_u[Y1]/1e3:.0f}k Chinese units vs {H[Y1]/1e3:.0f}k Honda", "+3 pp conquest rate from Chinese owners", [("bk3", "03 Customer")]),
    ("Plug the pickup gap", f"Pickups are {seg1.loc['Pickups','mkt_pct']:.0f}% of the market and a core Toyota profit pool. Honda has no offering. Test a targeted, "
     "USMCA-sourced Ridgeline launch with fleet and agribusiness buyers before committing to a full program.",
     f"{seg1.loc['Pickups','total']/1e3:.0f}k pickup units, Honda share 0%", "5k units · 2% pickup share", [("bk2", "02 Industry")]),
]
for i, (title, body, signal, kpi, bks) in enumerate(recs, 1):
    bk_tags = "".join(f'<span class="tag {c}"><b>Fixes:</b> {t_}</span>' for c, t_ in bks)
    html(f"""<div class="rec"><div class="num">{i}</div><div><h4>{title}</h4><p>{body}</p>
    <div class="tags">{bk_tags}<span class="tag"><b>Data signal:</b> {signal}</span><span class="tag"><b>2030 KPI:</b> {kpi}</span></div></div></div>""")

html("""
<div class="road">
  <div class="rd"><div class="when">2026 · Stop the bleeding</div><ul>
    <li>Total-cost-of-ownership offer: warranty, financing, resale guarantee</li><li>Reprice HR-V/BR-V entry trims</li><li>Pilot Ridgeline with fleets</li></ul></div>
  <div class="rd"><div class="when">2027–2028 · Electrify the core</div><ul>
    <li>e:HEV on CR-V, HR-V, City</li><li>Hybrid-led brand campaign</li><li>Dealer upgrades for EV service</li></ul></div>
  <div class="rd"><div class="when">2029–2030 · Scale & localize</div><ul>
    <li>Celaya-built entry model launch</li><li>3-row SUV added</li><li>Target 4.5–5% share, ≥ 40% electrified</li></ul></div>
</div>""")

html(f"""<div class="foot">Source: INEGI, Registro Administrativo de la Industria Automotriz de Vehículos Ligeros (RAIAVL), monthly retail sales
{Y0}–{Y1} by brand and model; electrified sales by state {e0}–{e1}. Market share = brand units / total reported units.
Units sold are used as the financial proxy because INEGI does not report revenue. Brand origin = the brand's home country, not the plant's.
Segments follow INEGI's classification. 2030 projections are illustrative scenarios, not forecasts.</div>""")
