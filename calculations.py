"""Lógica de cálculo de horas de servicio becario."""

from __future__ import annotations

import re
from datetime import date

import pandas as pd

from data_loader import ParsedSheet, normalize_text

# Tokens en la columna Becario(s) que significan "todos los becarios activos".
_ALL_TOKENS = ("todos", "team", "todas")
# Tokens que no son nombres y no deben generar advertencia.
_IGNORE_TOKENS = {"", "-", "na", "n/a", "pendiente", "pendientes", "x"}


def split_becarios(raw: str) -> tuple[bool, list[str]]:
    """(participan_todos, [nombres]) a partir de la celda Becario(s)."""
    norm = normalize_text(raw)
    if not norm:
        return False, []
    if any(re.search(rf"\b{token}\b", norm) for token in _ALL_TOKENS):
        return True, []
    tokens = re.split(r"[,/;\n]+", str(raw))
    names = []
    for token in tokens:
        cleaned = re.sub(r"\s+", " ", token).strip()
        if normalize_text(cleaned) not in _IGNORE_TOKENS:
            names.append(cleaned)
    return False, names


def compute_dashboard(parsed: ParsedSheet) -> dict:
    """Cruza eventos × becarios × tabla de becas.

    Devuelve:
      - becarios: DataFrame con meta, hechas, restantes, avance, estatus y la
        lista de eventos cubiertos por cada becario.
      - warnings: nombres sin match, becarios sin meta, horas pendientes.
    """
    warnings = list(parsed.warnings)
    becarios = parsed.becarios.copy()
    events = parsed.events

    # % beca -> horas requeridas
    metas = {}
    if not parsed.tabla_becas.empty:
        metas = dict(zip(parsed.tabla_becas["pct_beca"], parsed.tabla_becas["horas_requeridas"]))

    if becarios.empty:
        return {"becarios": becarios, "warnings": warnings}

    name_map = {normalize_text(n): n for n in becarios["nombre"]}
    hours_done = {n: 0.0 for n in becarios["nombre"]}
    events_by_becario = {n: [] for n in becarios["nombre"]}
    unmatched = []

    for _, ev in events.iterrows():
        is_all, names = split_becarios(ev["becarios_raw"])
        if is_all:
            assigned = list(becarios["nombre"])
        else:
            assigned = []
            for name in names:
                canonical = name_map.get(normalize_text(name))
                if canonical:
                    assigned.append(canonical)
                else:
                    unmatched.append((name, ev["evento"], ev["fila"]))
        for name in assigned:
            hours_done[name] += ev["horas"]
            events_by_becario[name].append({
                "evento": ev["evento"],
                "fecha": ev["fecha"],
                "horas": ev["horas"],
                "horas_pendientes": ev["horas_pendientes"],
                "horas_estimadas": ev["horas_estimadas"],
            })

    for name, evento, fila in unmatched:
        warnings.append(
            f"'{name}' (evento \"{evento}\", fila {fila}) no coincide con ningún becario activo — "
            "revisa la ortografía en el Excel; sus horas no se están contando."
        )

    pendientes = events[events["horas_pendientes"] & (events["becarios_raw"] != "")]
    for _, ev in pendientes.iterrows():
        warnings.append(
            f"El evento \"{ev['evento']}\" (fila {ev['fila']}) tiene becarios asignados pero sin "
            "'Horas Contabilizables' ni horario interpretable: se contó como 0 h."
        )

    def build_row(row):
        meta = metas.get(row["pct_beca"])
        hechas = round(hours_done[row["nombre"]], 2)
        if meta is None or meta == 0:
            warnings.append(
                f"{row['nombre']}: su % de beca ({_fmt_pct(row['pct_beca'])}) no aparece en la "
                "Tabla de becas; no se puede calcular su meta."
            )
            return pd.Series({"meta": None, "hechas": hechas, "restantes": None,
                              "avance": None, "estatus": "Sin meta"})
        avance = hechas / meta
        if hechas >= meta:
            estatus = "Cumplido"
        elif avance >= 0.5:
            estatus = "En curso"
        else:
            estatus = "Atrasado"
        return pd.Series({
            "meta": meta,
            "hechas": hechas,
            "restantes": max(meta - hechas, 0.0),
            "avance": min(avance, 1.0),
            "estatus": estatus,
        })

    becarios = pd.concat([becarios, becarios.apply(build_row, axis=1)], axis=1)
    becarios["eventos"] = becarios["nombre"].map(events_by_becario)
    return {"becarios": becarios, "warnings": warnings}


def _fmt_pct(value) -> str:
    if value is None:
        return "¿?"
    return f"{value:g}%"


def summary_metrics(becarios: pd.DataFrame) -> dict:
    """Números para las tarjetas resumen."""
    total = len(becarios)
    con_meta = becarios[becarios["meta"].notna()] if total else becarios
    cumplidos = int((con_meta["estatus"] == "Cumplido").sum()) if total else 0
    return {
        "activos": total,
        "pct_meta_cumplida": (cumplidos / len(con_meta)) if len(con_meta) else 0.0,
        "horas_hechas": float(con_meta["hechas"].sum()) if total else 0.0,
        "horas_meta_total": float(con_meta["meta"].sum()) if total else 0.0,
        "evals_pendientes": int((becarios["eval_360"] == "Pendiente").sum()) if total else 0,
    }


def upcoming_events(events: pd.DataFrame, today: date | None = None, limit: int = 8) -> pd.DataFrame:
    """Eventos con fecha >= hoy ordenados por fecha; si no hay, los últimos."""
    if events.empty:
        return events
    today = today or date.today()
    dated = events[events["fecha"].notna()]
    future = dated[dated["fecha"] >= today].sort_values("fecha")
    if future.empty:
        return dated.sort_values("fecha", ascending=False).head(limit)
    return future.head(limit)
