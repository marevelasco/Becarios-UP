"""Dashboard de servicio becario — Universidad Panamericana."""

import base64
import contextlib
import hashlib
import html
import re
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import auth
import config_store
import excel_store
import reportes
import verificaciones as verif_store
from calculations import compute_dashboard, events_for_jefe, split_becarios, summary_metrics, upcoming_events
from data_loader import normalize_text, parse_sheet, sheet_scores

SAMPLE_FILE = Path(__file__).parent / "sample_data" / "ejemplo_becarios.xlsx"
ASSETS_DIR = Path(__file__).parent / "assets"


@st.cache_data(show_spinner=False)
def logo_data_uri(filename: str) -> str | None:
    """Logo oficial de la UP como data URI, o None si el archivo no está (no truena el dashboard)."""
    path = ASSETS_DIR / filename
    if not path.exists():
        return None
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode()

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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Cormorant+Garamond:wght@600;700&display=swap');

:root {
  --ink: #111110; --ink-2: #52514E; --ink-3: #8A8A85;
  --line: #EAEAE8; --line-soft: #F0F0EE; --card: #FFFFFF; --page: #F7F7F5;
  /* Identidad Universidad Panamericana (extraída de up.edu.mx) */
  --dorado: #C8A568; --dorado-soft: #F4EEE1;
  --rojo: #8C1F3D; --rojo-hover: #6E1830; --rojo-soft: #F6E8ED;
  --font-serif: 'Cormorant Garamond', Georgia, 'Times New Roman', serif;
}

/* Chrome de Streamlit fuera — pero sin ocultar el header completo: ahí vive el
   botón para volver a abrir el sidebar cuando alguien lo colapsa (si se oculta
   con display:none, ese botón desaparece con él y ya no hay forma de reabrirlo). */
footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; box-shadow: none; }
[data-testid="stAppDeployButton"], [data-testid="stMainMenu"] { display: none; }

.stApp { background: var(--page); }
.block-container { max-width: 1400px; margin: 0 auto; padding-top: 1.5rem; padding-bottom: 2.4rem; }

/* Inter en todo (incluidos modales y popovers, que se montan fuera del
   contenedor principal), sin romper los iconos material de Streamlit */
:is([data-testid="stAppViewContainer"], div[role="dialog"], [data-baseweb="popover"]) :is(p, span, div, button, input, textarea, label, small, h1, h2, h3, h4, h5, h6):not([data-testid="stIconMaterial"]):not([class*="material-symbols"]) {
  font-family: 'Inter', -apple-system, 'Segoe UI', sans-serif;
}

/* Tarjetas (contenedores con borde) */
div[data-testid="stVerticalBlockBorderWrapper"] {
  background: var(--card); border: 1px solid var(--line); border-radius: 14px;
  box-shadow: 0 2px 10px rgba(140,31,61,0.045); padding: 10px 8px;
}

/* Header propio */
.logo-badge { height: 44px; width: 44px; display: flex; align-items: center; justify-content: center; }
.logo-badge img { height: 100%; width: auto; object-fit: contain; }
.logo-badge-fallback { width: 44px; height: 44px; border-radius: 11px; background: var(--rojo);
  color: var(--dorado); font-weight: 700; font-size: 15px; letter-spacing: .02em;
  display: flex; align-items: center; justify-content: center; }
.app-title { font-family: var(--font-serif) !important; font-size: 26px; font-weight: 700; color: var(--ink);
  letter-spacing: -.01em; line-height: 1.1; }
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
div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) { background: var(--dorado); }
div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) p { color: var(--ink); font-weight: 700; }

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
[data-testid="stBaseButton-primary"], [data-testid="stBaseButton-secondaryFormSubmit"] {
  background: var(--rojo); border: 1px solid var(--rojo); color: #fff; font-weight: 600;
  border-radius: 10px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);
}
[data-testid="stBaseButton-primary"]:hover, [data-testid="stBaseButton-secondaryFormSubmit"]:hover {
  background: var(--rojo-hover); border-color: var(--rojo-hover); color: #fff;
}

/* Pestañas (Dashboard / Verificar horas) */
button[data-baseweb="tab"] p { font-weight: 500; }
button[data-baseweb="tab"][aria-selected="true"] p { color: var(--rojo); font-weight: 700; }
div[data-baseweb="tab-highlight"] { background-color: var(--rojo) !important; height: 2.5px !important; }
div[data-baseweb="tab-border"] { background-color: var(--line) !important; }

/* Foco de marca en inputs */
[data-baseweb="input"]:focus-within, [data-baseweb="base-input"]:focus-within {
  border-color: var(--dorado) !important; box-shadow: 0 0 0 1px var(--dorado) !important;
}

/* Tarjetas KPI */
.kpi-card { background: var(--card); border: 1px solid var(--line); border-radius: 14px;
  padding: 15px 18px 13px; box-shadow: 0 2px 10px rgba(140,31,61,0.045); }
.kpi-label { font-size: 12px; font-weight: 500; color: var(--ink-3); margin-bottom: 7px;
  display: flex; align-items: center; gap: 7px; }
.kpi-dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; flex: none; }
.kpi-value { font-size: 25px; font-weight: 700; color: var(--ink); line-height: 1.05; letter-spacing: -.02em; }
.kpi-sub { font-size: 12px; color: var(--ink-3); margin-top: 5px; }
.kpi-meter { height: 5px; border-radius: 999px; background: #EEEEEB; margin-top: 9px; overflow: hidden; }
.kpi-meter > div { height: 100%; border-radius: 999px; background: var(--dorado); }

/* Secciones */
.sec-title { font-family: var(--font-serif) !important; font-size: 19px; font-weight: 700; color: var(--ink);
  letter-spacing: -.005em; }
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

jefe_nombre, jefe_username = auth.render_login_sidebar()


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


def persist_excel(file_bytes: bytes, name: str) -> None:
    """Guarda el Excel en Postgres para que no se pierda al refrescar. No falla si no hay base."""
    try:
        conn = excel_store.get_connection()
        excel_store.guardar_excel(conn, name, file_bytes)
    except Exception:
        pass


def try_autoload_from_db() -> bool:
    """Si hay un Excel guardado en Postgres y todavía no hay uno en la sesión, lo carga."""
    if "data" in st.session_state:
        return True
    try:
        conn = excel_store.get_connection()
        guardado = excel_store.cargar_excel(conn)
    except Exception:
        return False
    if guardado is None:
        return False
    set_data(guardado[1], guardado[0])
    return True


def handle_upload(uploaded):
    """Aplica un archivo del uploader solo si es distinto al ya cargado, y lo guarda en la base."""
    if uploaded is None:
        return
    file_bytes = uploaded.getvalue()
    md5 = hashlib.md5(file_bytes).hexdigest()
    if st.session_state.get("last_upload_md5") != md5:
        st.session_state["last_upload_md5"] = md5
        set_data(file_bytes, uploaded.name)
        persist_excel(file_bytes, uploaded.name)
        st.rerun()


def load_sample():
    # Solo para pruebas: nunca se guarda en la base (no debe pisar el Excel real).
    set_data(SAMPLE_FILE.read_bytes(), SAMPLE_FILE.name)


def fecha_actualizacion_excel():
    """Cuándo se guardó el Excel actual en la base, o None si no aplica."""
    try:
        conn = excel_store.get_connection()
        return excel_store.fecha_actualizacion(conn)
    except Exception:
        return None


def get_fecha_limite(sheet: str):
    """Fecha límite configurada para este semestre, o None si no hay base/no se ha puesto."""
    try:
        conn = config_store.get_connection()
        return config_store.cargar_fecha_limite(conn, sheet)
    except Exception:
        return None


def set_fecha_limite(sheet: str, fecha) -> None:
    try:
        conn = config_store.get_connection()
        config_store.guardar_fecha_limite(conn, sheet, fecha)
    except Exception:
        pass


def urgencia_html(restantes, fecha_limite, hoy) -> str:
    """Pill chica de urgencia si hay fecha límite configurada y todavía faltan horas."""
    if not fecha_limite or restantes is None or restantes <= 0:
        return ""
    dias = (fecha_limite - hoy).days
    if dias < 0:
        return pill(f"Venció hace {abs(dias)} d", {"fg": "#C0392B", "bg": "#FCE8E6", "icon": "🔴"})
    if dias <= 14:
        return pill(f"Quedan {dias} d", {"fg": "#B7791F", "bg": "#FEF3E2", "icon": "🟠"})
    return ""


def get_verificaciones(sheet: str) -> dict:
    """Horas verificadas por jefes para esta hoja, o {} si no hay base configurada."""
    try:
        conn = verif_store.get_connection()
        return verif_store.cargar_verificaciones(conn, sheet)
    except Exception:
        return {}


def encargados_sin_cuenta(events: pd.DataFrame) -> list[str]:
    """ENCARGADOs del Excel que no tienen ningún jefe configurado en Secrets con ese nombre."""
    ignorar = {"", "pendiente", "na", "n/a", "-"}
    vistos: dict[str, str] = {}
    for raw in events["encargado"]:
        for token in re.split(r"[,/;\n]+", str(raw)):
            nombre = token.strip()
            norm = normalize_text(nombre)
            if norm in ignorar:
                continue
            vistos.setdefault(norm, nombre)
    configurados = auth.configured_jefe_names()
    return [nombre for norm, nombre in vistos.items() if norm not in configurados]


# ---------------------------------------------------------------------------
# Piezas de UI
# ---------------------------------------------------------------------------

def fmt_h(value) -> str:
    return f"{value:g}"


def fmt_names(names: list[str], limit: int = 3) -> str:
    """Lista de nombres separados por coma, recortada para que no se haga infinita."""
    if not names:
        return ""
    if len(names) <= limit:
        return ", ".join(names)
    return f"{', '.join(names[:limit])} +{len(names) - limit} más"


def _logo_badge_html() -> str:
    uri = logo_data_uri("logo_tree.png")
    if uri:
        return f'<div class="logo-badge"><img src="{uri}" alt="Universidad Panamericana"></div>'
    return '<div class="logo-badge-fallback">UP</div>'  # respaldo si falta assets/logo_tree.png


def welcome_brand_html() -> str:
    """Logo completo (con el wordmark oficial) para la pantalla de bienvenida."""
    uri = logo_data_uri("logo_full.png")
    if uri:
        return f'<img src="{uri}" alt="Universidad Panamericana" style="height:52px">'
    return brand_html()


def brand_html() -> str:
    return (
        '<div style="display:flex;align-items:center;gap:14px">'
        f'{_logo_badge_html()}'
        '<div><div class="app-title">Becarios UP</div>'
        '<div class="app-sub">Servicio becario · Dirección</div></div></div>'
    )


def session_badge_html(nombre: str) -> str:
    """Insignia siempre visible (fuera del sidebar) con el jefe que tiene la sesión abierta."""
    return (
        '<div style="display:flex;align-items:center;gap:9px;justify-content:flex-end">'
        f'<div style="border-radius:50%;padding:2px;border:1.5px solid #C8A568">{avatar_html(nombre, 30)}</div>'
        '<div style="line-height:1.25">'
        f'<div style="font-size:13px;font-weight:600;color:#111110">{html.escape(nombre)}</div>'
        '<div style="font-size:11px;color:#8A8A85">Jefe de becarios · sesión activa</div>'
        '</div></div>'
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


def render_verification_panel(jefe_nombre: str, parsed, becarios: pd.DataFrame, sheet: str):
    """Panel donde un jefe confirma o corrige las horas de sus becarios, un expander por evento."""
    st.markdown(
        '<div class="sec-sub" style="margin-bottom:10px">Confirma o corrige las horas de los '
        'becarios que te apoyaron en cada actividad que organizaste. Lo que guardes aquí es lo '
        'que cuenta en su total. Los eventos ya verificados aparecen cerrados — ábrelos si '
        'necesitas cambiar algo.</div>',
        unsafe_allow_html=True,
    )

    try:
        conn = verif_store.get_connection()
    except Exception:
        st.warning(
            "Aún no hay una base de datos conectada (falta configurar Secrets). Puedes ver "
            "tus eventos, pero por ahora no se puede guardar lo que confirmes."
        )
        conn = None

    eventos_jefe = events_for_jefe(parsed.events, becarios, jefe_nombre)
    if not eventos_jefe:
        st.markdown(
            '<div class="sec-sub">No apareces como ENCARGADO en ningún evento de esta hoja.</div>',
            unsafe_allow_html=True,
        )
        return

    registros = verif_store.cargar_verificaciones(conn, sheet) if conn else {}

    for ev in eventos_jefe:
        fecha_key = ev["fecha"].isoformat() if ev["fecha"] else ""
        if ev["fecha"]:
            fecha_txt = f'{ev["fecha"].day:02d} {MESES[ev["fecha"].month - 1].lower()} {ev["fecha"].year}'
        else:
            fecha_txt = "sin fecha"
        lugar_txt = f' · {ev["lugar"]}' if ev["lugar"] else ""

        registros_ev = {
            becario: registros.get((normalize_text(ev["evento"]), fecha_key, normalize_text(becario)))
            for becario in ev["becarios"]
        }
        verificados = sum(1 for r in registros_ev.values() if r)
        total = len(ev["becarios"])
        discrepancias_ev = [
            becario for becario, r in registros_ev.items()
            if r and abs(float(r["horas_verificadas"]) - float(ev["horas"])) > 1e-9
        ]
        resumen = f"{verificados}/{total} verificadas" if verificados else f"{total} sin verificar"
        if discrepancias_ev:
            resumen += " · ⚠️ Excel actualizado"
        icono = "✅" if verificados == total and not discrepancias_ev else "⏳"
        label = f"{icono} {ev['evento']} · {fecha_txt}{lugar_txt} · {resumen}"

        with st.expander(label, expanded=(verificados < total or bool(discrepancias_ev))):
            for becario in ev["becarios"]:
                registro = registros_ev[becario]
                default_horas = float(registro["horas_verificadas"]) if registro else float(ev["horas"])
                input_key = f"verif_{sheet}_{normalize_text(ev['evento'])}_{fecha_key}_{normalize_text(becario)}"

                cols = st.columns([2.2, 1.1, 1.6, 1.1], vertical_alignment="center")
                cols[0].markdown(
                    f'<div style="display:flex;align-items:center;gap:8px">{avatar_html(becario, 24)}'
                    f'<span style="font-size:13.5px">{html.escape(becario)}</span></div>',
                    unsafe_allow_html=True,
                )
                horas_input = cols[1].number_input(
                    "Horas", min_value=0.0, step=0.5, value=default_horas,
                    key=input_key, label_visibility="collapsed",
                )
                if registro:
                    cols[2].markdown(
                        f'<span class="meter-text">✅ {registro["estatus"]} por '
                        f'{html.escape(registro["verificado_por"])}</span>',
                        unsafe_allow_html=True,
                    )
                else:
                    cols[2].markdown(
                        '<span class="meter-text" style="color:#B7791F">⏳ Sin verificar</span>',
                        unsafe_allow_html=True,
                    )
                guardar = cols[3].button(
                    "Guardar", key=f"save_{input_key}", type="primary", disabled=(conn is None),
                )
                if becario in discrepancias_ev:
                    st.markdown(
                        f'<div style="font-size:12px;color:#B7791F;margin:-4px 0 6px">⚠️ El Excel '
                        f'ahora dice {fmt_h(float(ev["horas"]))} h para este becario, pero tú '
                        f'verificaste {fmt_h(float(registro["horas_verificadas"]))} h. Revisa y '
                        f'guarda de nuevo si el cambio es correcto.</div>',
                        unsafe_allow_html=True,
                    )
                if registro and conn:
                    historial = verif_store.cargar_historial(
                        conn, sheet, normalize_text(ev["evento"]), fecha_key, normalize_text(becario)
                    )
                    if len(historial) > 1:
                        # st.expander no se puede anidar dentro de otro expander — usamos
                        # popover, que sí es válido aquí.
                        with st.popover(f"🕓 Ver historial ({len(historial)} cambios)"):
                            for h in historial:
                                anterior = (
                                    fmt_h(float(h["horas_anteriores"]))
                                    if h["horas_anteriores"] is not None else "—"
                                )
                                nuevo = fmt_h(float(h["horas_nuevas"]))
                                fecha_h = h["verificado_en"].strftime("%d %b %Y, %H:%M")
                                st.markdown(
                                    f'<div style="font-size:12px;color:var(--ink-3)">{fecha_h} · '
                                    f'{html.escape(h["verificado_por"])}: {anterior} h → {nuevo} h</div>',
                                    unsafe_allow_html=True,
                                )
                if guardar and conn:
                    verif_store.guardar_verificacion(
                        conn, semestre=sheet, evento=ev["evento"], fecha=ev["fecha"],
                        becario=becario, encargado=jefe_nombre,
                        horas_originales=float(ev["horas"]), horas_verificadas=float(horas_input),
                        verificado_por=jefe_nombre,
                    )
                    st.rerun()


# ---------------------------------------------------------------------------
# Pantalla de carga inicial
# ---------------------------------------------------------------------------

try_autoload_from_db()

if "data" not in st.session_state:
    if jefe_nombre:
        b_left, b_right = st.columns([2, 1], vertical_alignment="center")
        b_left.markdown(welcome_brand_html(), unsafe_allow_html=True)
        b_right.markdown(session_badge_html(jefe_nombre), unsafe_allow_html=True)
    else:
        st.markdown(welcome_brand_html(), unsafe_allow_html=True)
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

head_left, _, head_mid, head_session, head_right = st.columns(
    [1.3, 0.55, 1.6, 1.0, 0.55], vertical_alignment="center"
)
with head_left:
    st.markdown(brand_html(), unsafe_allow_html=True)
with head_mid:
    st.markdown('<div class="field-label">Semestre (hoja del Excel)</div>', unsafe_allow_html=True)
    sheet = st.radio("Semestre", sheet_options, horizontal=True, label_visibility="collapsed",
                     key=f"sheet_{st.session_state['data_md5']}")
with head_session:
    if jefe_nombre:
        st.markdown(session_badge_html(jefe_nombre), unsafe_allow_html=True)
with head_right:
    with st.popover("📁 Archivo", width="stretch"):
        st.caption(f"Archivo actual: **{data_name}**")
        actualizado = fecha_actualizacion_excel()
        if actualizado:
            st.caption(f"Última actualización guardada: {actualizado.strftime('%d %b %Y, %H:%M')}")
        handle_upload(st.file_uploader("Subir otro Excel", type=["xlsx", "xlsm"], key="re_upload"))
        if SAMPLE_FILE.exists() and st.button("Usar archivo de ejemplo"):
            load_sample()
            st.rerun()

        st.markdown("---")
        st.caption("Fecha límite de este semestre (opcional, para el aviso de urgencia)")
        fecha_limite_actual = get_fecha_limite(sheet)
        nueva_fecha = st.date_input(
            "Fecha límite", value=fecha_limite_actual, key=f"fecha_limite_{sheet}",
            label_visibility="collapsed",
        )
        if nueva_fecha != fecha_limite_actual:
            set_fecha_limite(sheet, nueva_fecha)
            st.rerun()

parsed = cached_parse(data, sheet)
verificaciones = get_verificaciones(sheet)
result = compute_dashboard(parsed, verificaciones)
becarios = result["becarios"]
warnings = result["warnings"]
metrics = summary_metrics(becarios)

st.write("")

if jefe_nombre:
    tab_dashboard, tab_verificar = st.tabs(["📊 Dashboard", "✅ Verificar horas"])
else:
    tab_dashboard, tab_verificar = contextlib.nullcontext(), None

with tab_dashboard:
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
         "#C8A568", None),
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

    if verificaciones:
        st.markdown(
            f'<div class="footer-note">✅ {result["pct_verificado"]:.0%} de las horas mostradas ya '
            'fueron confirmadas por un jefe de becarios.</div>',
            unsafe_allow_html=True,
        )
        if result["discrepancias"]:
            st.markdown(
                f'<div class="footer-note" style="color:#B7791F">⚠️ {result["discrepancias"]} hora(s) '
                'verificada(s) ya no coinciden con el Excel actual — el jefe correspondiente puede '
                'revisarlas en su pestaña "Verificar horas".</div>',
                unsafe_allow_html=True,
            )

    if warnings:
        with st.expander(f"⚠️ Advertencias del archivo ({len(warnings)})"):
            for w in warnings:
                st.markdown(f"- {w}")

    pendientes_cuenta = encargados_sin_cuenta(parsed.events) if st.secrets.get("jefes") else []
    if pendientes_cuenta:
        with st.expander(f"ℹ️ Encargados sin cuenta de jefe todavía ({len(pendientes_cuenta)})"):
            st.markdown(
                "Estos nombres aparecen como `ENCARGADO` en el Excel pero no tienen usuario "
                "configurado en Secrets, así que no pueden verificar horas todavía:"
            )
            for nombre in pendientes_cuenta:
                st.markdown(f"- {nombre}")
            st.caption(
                "Dales de alta con `scripts/generar_password.py` y agrégalos a Secrets, sección "
                "`[jefes.<usuario>]`, con `name` escrito igual que aquí."
            )

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
            f'<div style="margin-left:auto">{pill(row["estatus"], ss)} '
            f'{urgencia_html(row["restantes"], fecha_limite, hoy)}</div></div>',
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
        title_col, download_col = st.columns([3, 1], vertical_alignment="center")
        title_col.markdown(
            '<div class="sec-title">Cumplimiento de horas de servicio</div>'
            '<div class="sec-sub">Haz clic en un nombre para ver el detalle de sus eventos.</div>',
            unsafe_allow_html=True,
        )
        if not becarios.empty:
            nombre_base = f"reporte_{normalize_text(sheet).replace(' ', '_')}"
            with download_col.popover("📥 Descargar reporte", width="stretch"):
                st.download_button(
                    "Excel (.xlsx)",
                    data=reportes.generar_reporte_excel(becarios, sheet),
                    file_name=f"{nombre_base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )
                st.download_button(
                    "PDF",
                    data=reportes.generar_reporte_pdf(becarios, sheet),
                    file_name=f"{nombre_base}.pdf",
                    mime="application/pdf",
                    width="stretch",
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
            fecha_limite = get_fecha_limite(sheet)
            hoy = date.today()

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
                    urgencia = urgencia_html(row["restantes"], fecha_limite, hoy)
                    if urgencia:
                        cols[3].markdown(f'<div style="margin-top:4px">{urgencia}</div>', unsafe_allow_html=True)
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
                hora_txt = f' · {html.escape(ev_row["hora_texto"])}' if ev_row["hora_texto"] else ""
                place = f'<div class="ev-place">{html.escape(ev_row["lugar"])}</div>' if ev_row["lugar"] else ""
                hours_note = f' · {fmt_h(ev_row["horas"])} h' if ev_row["horas"] else ""

                encargados = [
                    n.strip() for n in re.split(r"[,/;\n]+", ev_row["encargado"])
                    if n.strip() and normalize_text(n) not in ("pendiente", "na", "-")
                ]
                encargado_txt = fmt_names(encargados) if encargados else "Sin asignar"
                encargado_avatar = avatar_html(encargados[0], 18) if encargados else ""

                is_all, becario_names = split_becarios(ev_row["becarios_raw"])
                if is_all and not becarios.empty:
                    becario_names = list(becarios["nombre"])
                becarios_txt = fmt_names(becario_names) if becario_names else "Sin becarios asignados todavía"
                becario_avatar = avatar_html(becario_names[0], 18) if becario_names else ""

                who = (
                    f'<div class="ev-who">{encargado_avatar}'
                    f'<span>Encargado: {html.escape(encargado_txt)}</span></div>'
                    f'<div class="ev-who">{becario_avatar}'
                    f'<span>Becarios: {html.escape(becarios_txt)}{hours_note}</span></div>'
                )

                items.append(
                    f'<div class="ev-item"><div class="ev-date">'
                    f'<div class="ev-mon">{MESES[ev_row["fecha"].month - 1]}</div>'
                    f'<div class="ev-day">{ev_row["fecha"].day}</div></div>'
                    f'<div style="min-width:0"><div class="ev-name">{html.escape(ev_row["evento"])}'
                    f'<span style="font-weight:400;color:var(--ink-3)">{hora_txt}</span></div>'
                    f'{place}{who}</div></div>'
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

if tab_verificar is not None:
    with tab_verificar:
        render_verification_panel(jefe_nombre, parsed, becarios, sheet)
