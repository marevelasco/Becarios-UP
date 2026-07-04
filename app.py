"""Dashboard de servicio becario — Universidad Panamericana."""

import hashlib
import html
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from calculations import compute_dashboard, summary_metrics, upcoming_events
from data_loader import normalize_text, parse_sheet, sheet_scores

SAMPLE_FILE = Path(__file__).parent / "sample_data" / "ejemplo_becarios.xlsx"

st.set_page_config(page_title="Becarios UP", page_icon="🎓", layout="wide")

# ---------------------------------------------------------------------------
# Paleta y tokens visuales
# ---------------------------------------------------------------------------

STATUS_STYLES = {
    "Cumplido": {"fg": "#1A7F4E", "bg": "#E6F4EC", "fill": "#1A7F4E", "track": "#E6F4EC", "icon": "dot"},
    "En curso": {"fg": "#B7791F", "bg": "#FEF3E2", "fill": "#B7791F", "track": "#FEF3E2", "icon": "dot"},
    "Atrasado": {"fg": "#C0392B", "bg": "#FCE8E6", "fill": "#C0392B", "track": "#FCE8E6", "icon": "dot"},
    "Sin meta": {"fg": "#8A8A85", "bg": "#F1F1EF", "fill": "#8A8A85", "track": "#F1F1EF", "icon": "dot"},
}
EVAL_STYLES = {
    "Completada": {"fg": "#1A7F4E", "bg": "#E6F4EC", "icon": "dot"},
    "Pendiente": {"fg": "#B7791F", "bg": "#FEF3E2", "icon": "⏳"},
}
# Pasteles para avatares: (fondo, texto). Asignación determinística por nombre.
PASTELS = [
    ("#FDE8E8", "#B03E3E"), ("#FEF0DC", "#A86A1C"), ("#FBF3CF", "#8F7B1F"),
    ("#E2F5E9", "#2F7D4F"), ("#E1F0FA", "#2E6E9E"), ("#ECE8FA", "#5D4FB0"),
    ("#FAE6F2", "#A94E86"), ("#EDEBE3", "#79684A"),
]
MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root {
  --ink: #111110; --ink-2: #52514E; --ink-3: #8A8A85;
  --line: #EAEAE8; --line-soft: #F0F0EE; --card: #FFFFFF; --page: #F7F7F5;
}

/* Chrome de Streamlit fuera */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header[data-testid="stHeader"] { display: none; }

.stApp { background: var(--page); }
.block-container { max-width: 1400px; margin: 0 auto; padding-top: 2.2rem; padding-bottom: 3rem; }

/* Inter en todo (incluidos modales y popovers, que se montan fuera del
   contenedor principal), sin romper los iconos material de Streamlit */
:is([data-testid="stAppViewContainer"], div[role="dialog"], [data-baseweb="popover"]) :is(p, span, div, button, input, textarea, label, small, h1, h2, h3, h4, h5, h6):not([data-testid="stIconMaterial"]):not([class*="material-symbols"]) {
  font-family: 'Inter', -apple-system, 'Segoe UI', sans-serif;
}

/* Tarjetas (contenedores con borde) */
div[data-testid="stVerticalBlockBorderWrapper"] {
  background: var(--card); border: 1px solid var(--line); border-radius: 12px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.04); padding: 10px 8px;
}

/* Header propio */
.logo-badge { width: 44px; height: 44px; border-radius: 11px; background: var(--ink);
  color: #fff; font-weight: 700; font-size: 15px; letter-spacing: .02em;
  display: flex; align-items: center; justify-content: center; }
.app-title { font-size: 22px; font-weight: 700; color: var(--ink); letter-spacing: -.02em; line-height: 1.15; }
.app-sub { font-size: 13px; color: var(--ink-3); margin-top: 2px; }
.field-label { font-size: 11px; font-weight: 600; color: var(--ink-3); text-transform: uppercase;
  letter-spacing: .06em; margin-bottom: 6px; }

/* Radios como control segmentado tipo pastillas */
div[data-testid="stRadio"] > div[role="radiogroup"] { flex-direction: row; flex-wrap: wrap; gap: 6px; }
div[data-testid="stRadio"] label[data-baseweb="radio"] {
  background: #F1F1EF; border-radius: 999px; padding: 6px 14px; margin: 0;
  border: 1px solid transparent; cursor: pointer; transition: background .15s ease;
}
div[data-testid="stRadio"] label[data-baseweb="radio"] > div:first-of-type { display: none; }
div[data-testid="stRadio"] label[data-baseweb="radio"] p { font-size: 13px; font-weight: 500; color: var(--ink-2); }
div[data-testid="stRadio"] label[data-baseweb="radio"]:hover { background: #E8E8E4; }
div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) { background: var(--ink); }
div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) p { color: #fff; font-weight: 600; }

/* Botones */
[data-testid="stBaseButton-secondary"], [data-testid="stPopover"] button {
  background: #fff; border: 1px solid var(--line); border-radius: 10px;
  color: var(--ink); font-weight: 500; box-shadow: 0 1px 2px rgba(0,0,0,0.03);
}
[data-testid="stBaseButton-secondary"]:hover, [data-testid="stPopover"] button:hover {
  border-color: #D5D5D0; color: var(--ink);
}
[data-testid="stBaseButton-tertiary"] { color: var(--ink) !important; font-weight: 600 !important;
  padding: 0 !important; min-height: 0 !important; font-size: 14px !important; }
[data-testid="stBaseButton-tertiary"]:hover { text-decoration: underline; }

/* Tarjetas KPI */
.kpi-card { background: var(--card); border: 1px solid var(--line); border-radius: 12px;
  padding: 20px 22px 18px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
.kpi-label { font-size: 12.5px; font-weight: 500; color: var(--ink-3); margin-bottom: 10px;
  display: flex; align-items: center; gap: 7px; }
.kpi-dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; flex: none; }
.kpi-value { font-size: 30px; font-weight: 700; color: var(--ink); line-height: 1.05; letter-spacing: -.02em; }
.kpi-sub { font-size: 12.5px; color: var(--ink-3); margin-top: 7px; }
.kpi-meter { height: 5px; border-radius: 999px; background: #EEEEEB; margin-top: 14px; overflow: hidden; }
.kpi-meter > div { height: 100%; border-radius: 999px; background: var(--ink); }

/* Secciones */
.sec-title { font-size: 15.5px; font-weight: 700; color: var(--ink); letter-spacing: -.01em; }
.sec-sub { font-size: 12.5px; color: var(--ink-3); margin: 3px 0 6px; }

/* Tabla de becarios */
.tbl-head { font-size: 11px; font-weight: 600; color: var(--ink-3); text-transform: uppercase;
  letter-spacing: .06em; white-space: nowrap; }
.row-sep { border: none; border-top: 1px solid var(--line-soft); margin: 9px 0; }
.avatar { border-radius: 50%; display: flex; align-items: center; justify-content: center;
  font-weight: 600; letter-spacing: .02em; flex: none; }
.pct-chip { font-size: 13px; font-weight: 500; color: var(--ink-2); font-variant-numeric: tabular-nums; }

.meter-row { display: flex; align-items: center; gap: 12px; }
.meter { height: 8px; border-radius: 999px; overflow: hidden; flex: 1; }
.meter > div { height: 100%; border-radius: 999px; }
.meter-text { font-size: 12.5px; font-weight: 500; color: var(--ink-2); white-space: nowrap;
  font-variant-numeric: tabular-nums; }

/* Pills de estatus / eval */
.pill { display: inline-flex; align-items: center; gap: 6px; border-radius: 999px;
  padding: 4px 11px; font-size: 12px; font-weight: 600; white-space: nowrap; }
.pill-dot { width: 6px; height: 6px; border-radius: 50%; background: currentColor; display: inline-block; flex: none; }

/* Próximos eventos */
.ev-item { display: flex; gap: 13px; padding: 12px 2px; border-bottom: 1px solid var(--line-soft); }
.ev-item:last-child { border-bottom: none; }
.ev-date { min-width: 48px; height: 52px; border-radius: 10px; background: #F1F1EF; flex: none;
  display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 1px; }
.ev-mon { font-size: 10px; font-weight: 600; color: var(--ink-3); letter-spacing: .09em; }
.ev-day { font-size: 17px; font-weight: 700; color: var(--ink); line-height: 1; }
.ev-name { font-size: 13.5px; font-weight: 600; color: var(--ink); line-height: 1.3; }
.ev-place { font-size: 12px; color: var(--ink-3); margin-top: 2px; }
.ev-who { display: flex; align-items: center; gap: 6px; margin-top: 6px; font-size: 12px; color: var(--ink-2); }
.ev-unassigned { font-style: italic; color: var(--ink-3); }

/* Detalle (modal) */
div[role="dialog"] { border-radius: 16px; }
.dlg-metric { flex: 1; background: #FAFAF8; border: 1px solid var(--line-soft); border-radius: 10px; padding: 14px 16px; }
.dlg-metric-label { font-size: 12px; color: var(--ink-3); margin-bottom: 5px; }
.dlg-metric-value { font-size: 22px; font-weight: 700; color: var(--ink); letter-spacing: -.01em; }
.dlg-ev { display: flex; align-items: baseline; gap: 12px; padding: 9px 2px; border-bottom: 1px solid var(--line-soft); }
.dlg-ev:last-child { border-bottom: none; }
.dlg-ev-date { min-width: 90px; font-size: 12px; color: var(--ink-3); font-variant-numeric: tabular-nums; }
.dlg-ev-name { flex: 1; font-size: 13.5px; font-weight: 500; color: var(--ink); }
.dlg-ev-hours { font-size: 13px; font-weight: 600; color: var(--ink); white-space: nowrap; }
.mini-tag { font-size: 10.5px; font-weight: 500; color: var(--ink-3); background: #F1F1EF;
  border-radius: 999px; padding: 2px 8px; margin-left: 6px; }

/* Uploader y expander */
[data-testid="stFileUploaderDropzone"] { background: #FAFAF8; border: 1px dashed #D8D8D3; border-radius: 12px; }
[data-testid="stExpander"] details { border: 1px solid var(--line); border-radius: 12px; background: #fff; }

.footer-note { font-size: 12px; color: var(--ink-3); margin-top: 6px; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Carga de datos (session_state + caché)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def cached_scores(file_bytes: bytes):
    return sheet_scores(file_bytes)


@st.cache_data(show_spinner="Leyendo Excel…")
def cached_parse(file_bytes: bytes, sheet: str):
    return parse_sheet(file_bytes, sheet)


def set_data(file_bytes: bytes, name: str):
    st.session_state["data"] = file_bytes
    st.session_state["data_name"] = name
    st.session_state["data_md5"] = hashlib.md5(file_bytes).hexdigest()


def handle_upload(uploaded):
    """Aplica un archivo del uploader solo si es distinto al ya cargado."""
    if uploaded is None:
        return
    file_bytes = uploaded.getvalue()
    md5 = hashlib.md5(file_bytes).hexdigest()
    if st.session_state.get("last_upload_md5") != md5:
        st.session_state["last_upload_md5"] = md5
        set_data(file_bytes, uploaded.name)
        st.rerun()


def load_sample():
    set_data(SAMPLE_FILE.read_bytes(), SAMPLE_FILE.name)


# ---------------------------------------------------------------------------
# Piezas de UI
# ---------------------------------------------------------------------------

def fmt_h(value) -> str:
    return f"{value:g}"


def brand_html() -> str:
    return (
        '<div style="display:flex;align-items:center;gap:14px">'
        '<div class="logo-badge">UP</div>'
        '<div><div class="app-title">Becarios UP</div>'
        '<div class="app-sub">Servicio becario · Dirección</div></div></div>'
    )


def pastel_for(nombre: str):
    digest = hashlib.md5(normalize_text(nombre).encode()).hexdigest()
    return PASTELS[int(digest[:8], 16) % len(PASTELS)]


def avatar_html(nombre: str, size: int = 36) -> str:
    words = nombre.split()
    initials = (words[0][0] + (words[1][0] if len(words) > 1 else "")).upper()
    bg, fg = pastel_for(nombre)
    style = f"background:{bg};color:{fg};width:{size}px;height:{size}px;font-size:{size * .36:.0f}px"
    return f'<div class="avatar" style="{style}">{html.escape(initials)}</div>'


def pill(text: str, style: dict) -> str:
    icon = '<span class="pill-dot"></span>' if style["icon"] == "dot" else f'<span>{style["icon"]}</span>'
    return f'<span class="pill" style="background:{style["bg"]};color:{style["fg"]}">{icon}{html.escape(text)}</span>'


def meter_html(pct: float, fill: str, track: str, text: str) -> str:
    return (
        f'<div class="meter-row"><div class="meter" style="background:{track}">'
        f'<div style="width:{pct:.1%};background:{fill}"></div></div>'
        f'<span class="meter-text">{text}</span></div>'
    )


# ---------------------------------------------------------------------------
# Pantalla de carga inicial
# ---------------------------------------------------------------------------

if "data" not in st.session_state:
    st.markdown(brand_html(), unsafe_allow_html=True)
    st.write("")
    _, center, _ = st.columns([1, 1.6, 1])
    with center, st.container(border=True):
        st.markdown(
            '<div class="sec-title" style="margin-bottom:2px">Sube el Excel de eventos y becarios</div>'
            '<div class="sec-sub">El archivo debe tener la hoja del semestre (ej. «Propuesta de 2027-1») '
            'con la lista de eventos, la tabla de becas y los becarios activos. Puedes volver a subir '
            'una versión actualizada en cualquier momento.</div>',
            unsafe_allow_html=True,
        )
        handle_upload(st.file_uploader("Archivo Excel", type=["xlsx", "xlsm"], label_visibility="collapsed"))
        if SAMPLE_FILE.exists() and st.button("Usar archivo de ejemplo", type="secondary"):
            load_sample()
            st.rerun()
    st.stop()

data: bytes = st.session_state["data"]
data_name: str = st.session_state["data_name"]

# ---------------------------------------------------------------------------
# Encabezado: marca + selector de semestre + cambio de archivo
# ---------------------------------------------------------------------------

scores = cached_scores(data)
# Solo hojas que realmente parecen de eventos; máximo 6 para que el control
# segmentado no se desborde con hojas históricas.
sheet_options = [name for name, score in scores if score >= 5][:6] or [scores[0][0]]

head_left, _, head_mid, head_right = st.columns([1.45, 0.85, 1.6, 0.55], vertical_alignment="center")
with head_left:
    st.markdown(brand_html(), unsafe_allow_html=True)
with head_mid:
    st.markdown('<div class="field-label">Semestre (hoja del Excel)</div>', unsafe_allow_html=True)
    sheet = st.radio("Semestre", sheet_options, horizontal=True, label_visibility="collapsed",
                     key=f"sheet_{st.session_state['data_md5']}")
with head_right:
    with st.popover("📁 Archivo", width="stretch"):
        st.caption(f"Archivo actual: **{data_name}**")
        handle_upload(st.file_uploader("Subir otro Excel", type=["xlsx", "xlsm"], key="re_upload"))
        if SAMPLE_FILE.exists() and st.button("Usar archivo de ejemplo"):
            load_sample()
            st.rerun()

parsed = cached_parse(data, sheet)
result = compute_dashboard(parsed)
becarios = result["becarios"]
warnings = result["warnings"]
metrics = summary_metrics(becarios)

st.write("")

# ---------------------------------------------------------------------------
# Sección 1 — Tarjetas resumen
# ---------------------------------------------------------------------------

jefes = int((becarios["rol"].str.contains("jefe", case=False)).sum()) if len(becarios) else 0
meta_total = metrics["horas_meta_total"]
pct_horas = min(metrics["horas_hechas"] / meta_total, 1.0) if meta_total else 0.0
con_meta = becarios[becarios["meta"].notna()] if len(becarios) else becarios
cumplidos = int((con_meta["estatus"] == "Cumplido").sum()) if len(becarios) else 0
evals_dot = "#B7791F" if metrics["evals_pendientes"] else "#1A7F4E"

kpis = st.columns(4)
kpi_data = [
    ("Becarios activos", f"{metrics['activos']}",
     f"{jefes} jefe{'s' if jefes != 1 else ''} de becarios · {metrics['activos'] - jefes} becarios",
     "#8A8A85", None),
    ("Meta de horas cumplida", f"{metrics['pct_meta_cumplida']:.0%}",
     f"{cumplidos} de {len(con_meta)} becarios en meta", "#1A7F4E", None),
    ("Horas acumuladas", f"{fmt_h(metrics['horas_hechas'])} h",
     f"de {fmt_h(meta_total)} h meta del semestre", "#8A8A85", pct_horas),
    ("Evaluaciones 360 pendientes", f"{metrics['evals_pendientes']}",
     f"de {metrics['activos']} evaluaciones", evals_dot, None),
]
for col, (label, value, sub, dot, meter) in zip(kpis, kpi_data):
    meter_div = f'<div class="kpi-meter"><div style="width:{meter:.0%}"></div></div>' if meter is not None else ""
    col.markdown(
        f'<div class="kpi-card"><div class="kpi-label"><span class="kpi-dot" style="background:{dot}"></span>'
        f'{label}</div><div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div>{meter_div}</div>',
        unsafe_allow_html=True,
    )

if warnings:
    st.write("")
    with st.expander(f"⚠️ Advertencias del archivo ({len(warnings)})"):
        for w in warnings:
            st.markdown(f"- {w}")

st.write("")

# ---------------------------------------------------------------------------
# Detalle de becario (modal)
# ---------------------------------------------------------------------------

@st.dialog("Detalle del becario", width="large")
def show_detail(row):
    ss = STATUS_STYLES[row["estatus"]]
    ev_style = EVAL_STYLES[row["eval_360"]]
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px">{avatar_html(row["nombre"], 48)}'
        f'<div><div style="font-size:18px;font-weight:700;color:#111110">{html.escape(row["nombre"])}</div>'
        f'<div style="color:#8A8A85;font-size:13px;margin-top:2px">{html.escape(row["rol"] or "Becario")} · '
        f'Beca {row["pct_beca"]:g}% · Semestre {html.escape(str(row["semestre"]))}</div></div>'
        f'<div style="margin-left:auto">{pill(row["estatus"], ss)}</div></div>',
        unsafe_allow_html=True,
    )
    tiles = [
        ("Meta", f"{fmt_h(row['meta'])} h" if row["meta"] else "—"),
        ("Horas hechas", f"{fmt_h(row['hechas'])} h"),
        ("Restantes", f"{fmt_h(row['restantes'])} h" if row["restantes"] is not None else "—"),
    ]
    tiles_html = "".join(
        f'<div class="dlg-metric"><div class="dlg-metric-label">{label}</div>'
        f'<div class="dlg-metric-value">{value}</div></div>' for label, value in tiles
    )
    tiles_html += (
        f'<div class="dlg-metric"><div class="dlg-metric-label">Eval. 360</div>'
        f'<div style="margin-top:4px">{pill(row["eval_360"], ev_style)}</div></div>'
    )
    st.markdown(f'<div style="display:flex;gap:10px;margin-bottom:14px">{tiles_html}</div>',
                unsafe_allow_html=True)

    if row["meta"]:
        pct = min(row["hechas"] / row["meta"], 1.0)
        st.markdown(
            f'<div style="margin-bottom:18px">{meter_html(pct, ss["fill"], ss["track"], f"{pct:.0%} de avance")}</div>',
            unsafe_allow_html=True,
        )

    st.markdown('<div class="sec-title" style="font-size:14px;margin-bottom:4px">Eventos cubiertos</div>',
                unsafe_allow_html=True)
    if not row["eventos"]:
        st.markdown('<div class="sec-sub">Aún no tiene eventos registrados en este semestre.</div>',
                    unsafe_allow_html=True)
        return

    eventos = sorted(row["eventos"], key=lambda e: (e["fecha"] is None, e["fecha"] or date.max))
    rows_html = []
    for ev in eventos:
        if ev["fecha"]:
            fecha_txt = f'{ev["fecha"].day:02d} {MESES[ev["fecha"].month - 1].lower()} {ev["fecha"].year}'
        else:
            fecha_txt = "Sin fecha"
        tag = ""
        if ev["horas_pendientes"]:
            tag = '<span class="mini-tag">Pendiente de horas</span>'
        elif ev["horas_estimadas"]:
            tag = '<span class="mini-tag">Estimada del horario</span>'
        rows_html.append(
            f'<div class="dlg-ev"><div class="dlg-ev-date">{fecha_txt}</div>'
            f'<div class="dlg-ev-name">{html.escape(ev["evento"])}{tag}</div>'
            f'<div class="dlg-ev-hours">{fmt_h(ev["horas"])} h</div></div>'
        )
    st.markdown("".join(rows_html), unsafe_allow_html=True)
    st.markdown(
        f'<div class="footer-note">Total: <b style="color:#111110">{fmt_h(row["hechas"])} h</b> '
        f'en {len(eventos)} eventos.</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Sección 2 + 3 — Tabla de cumplimiento y próximos eventos
# ---------------------------------------------------------------------------

main_col, side_col = st.columns([2, 1], gap="medium")

with main_col, st.container(border=True):
    st.markdown(
        '<div class="sec-title">Cumplimiento de horas de servicio</div>'
        '<div class="sec-sub">Haz clic en un nombre para ver el detalle de sus eventos.</div>',
        unsafe_allow_html=True,
    )

    if becarios.empty:
        st.markdown('<div class="sec-sub">Esta hoja no tiene tabla de «Becarios activos». '
                    'Selecciona otra hoja o revisa el archivo.</div>', unsafe_allow_html=True)
    else:
        ordered = becarios.sort_values(["avance", "nombre"], ascending=[False, True], na_position="last")
        counts = becarios["estatus"].value_counts()
        filter_labels = [f"Todos ({len(becarios)})"] + [
            f"{status} ({int(counts.get(status, 0))})" for status in ("Cumplido", "En curso", "Atrasado")
        ]
        selected = st.radio("Filtro", filter_labels, horizontal=True, label_visibility="collapsed",
                            key=f"filter_{sheet}")
        status_filter = selected.rsplit(" (", 1)[0]
        subset = ordered if status_filter == "Todos" else ordered[ordered["estatus"] == status_filter]

        st.markdown('<hr class="row-sep">', unsafe_allow_html=True)
        if subset.empty:
            st.markdown('<div class="sec-sub">No hay becarios en esta categoría.</div>', unsafe_allow_html=True)
        else:
            head = st.columns([0.55, 2.5, 0.8, 3.4, 1.5, 1.4], vertical_alignment="center")
            for col, label in zip(head, ["", "Becario", "% Beca", "Horas de servicio", "Eval. 360", "Estatus"]):
                col.markdown(f'<div class="tbl-head">{label}</div>', unsafe_allow_html=True)
            st.markdown('<hr class="row-sep">', unsafe_allow_html=True)

            for _, row in subset.iterrows():
                ss = STATUS_STYLES[row["estatus"]]
                ev_style = EVAL_STYLES[row["eval_360"]]
                cols = st.columns([0.55, 2.5, 0.8, 3.4, 1.5, 1.4], vertical_alignment="center")
                cols[0].markdown(avatar_html(row["nombre"]), unsafe_allow_html=True)
                if cols[1].button(row["nombre"], key=f"b_{row['nombre']}", type="tertiary",
                                  help=f"{row['rol'] or 'Becario'} · ver detalle"):
                    show_detail(row)
                cols[2].markdown(f'<span class="pct-chip">{row["pct_beca"]:g}%</span>', unsafe_allow_html=True)
                if row["meta"]:
                    pct = min(row["hechas"] / row["meta"], 1.0)
                    cols[3].markdown(
                        meter_html(pct, ss["fill"], ss["track"], f'{fmt_h(row["hechas"])} / {fmt_h(row["meta"])} h'),
                        unsafe_allow_html=True,
                    )
                else:
                    cols[3].markdown(
                        f'<span class="meter-text">{fmt_h(row["hechas"])} h · sin meta definida</span>',
                        unsafe_allow_html=True,
                    )
                cols[4].markdown(pill(row["eval_360"], ev_style), unsafe_allow_html=True)
                cols[5].markdown(pill(row["estatus"], ss), unsafe_allow_html=True)
                st.markdown('<hr class="row-sep">', unsafe_allow_html=True)

with side_col, st.container(border=True):
    st.markdown('<div class="sec-title" style="margin-bottom:4px">Próximos eventos</div>', unsafe_allow_html=True)
    upcoming = upcoming_events(parsed.events, today=date.today(), limit=9)
    if upcoming.empty:
        st.markdown('<div class="sec-sub">No hay eventos con fecha en esta hoja.</div>', unsafe_allow_html=True)
    else:
        items = []
        for _, ev_row in upcoming.iterrows():
            place = f'<div class="ev-place">{html.escape(ev_row["lugar"])}</div>' if ev_row["lugar"] else ""
            hours_note = f' · {fmt_h(ev_row["horas"])} h' if ev_row["horas"] else ""

            raw = ev_row["becarios_raw"].strip()
            encargado = ev_row["encargado"].strip()
            if raw and normalize_text(raw) not in ("pendiente", "na", "-"):
                first = raw.replace("/", ",").split(",")[0].strip()
                who = (f'{avatar_html(first, 18)}<span>{html.escape(raw)}{hours_note}</span>')
            elif encargado and normalize_text(encargado) not in ("pendiente", "na", "-"):
                who = (f'{avatar_html(encargado, 18)}'
                       f'<span>Encargado: {html.escape(encargado)}{hours_note}</span>')
            else:
                who = f'<span class="ev-unassigned">Sin asignar</span>'

            items.append(
                f'<div class="ev-item"><div class="ev-date">'
                f'<div class="ev-mon">{MESES[ev_row["fecha"].month - 1]}</div>'
                f'<div class="ev-day">{ev_row["fecha"].day}</div></div>'
                f'<div style="min-width:0"><div class="ev-name">{html.escape(ev_row["evento"])}</div>'
                f'{place}<div class="ev-who">{who}</div></div></div>'
            )
        st.markdown("".join(items), unsafe_allow_html=True)

    undated = parsed.events[parsed.events["fecha"].isna()] if not parsed.events.empty else pd.DataFrame()
    if len(undated):
        st.markdown(f'<div class="footer-note">+ {len(undated)} eventos pendientes de fecha en el Excel.</div>',
                    unsafe_allow_html=True)

st.markdown(
    f'<div class="footer-note" style="margin-top:14px">Datos: <b style="color:#52514E">{html.escape(data_name)}</b> '
    f'· hoja «{html.escape(sheet)}» · {len(parsed.events)} eventos leídos.</div>',
    unsafe_allow_html=True,
)
