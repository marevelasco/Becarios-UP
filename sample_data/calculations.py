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


def _resolve_assigned(ev: pd.Series, becarios: pd.DataFrame, name_map: dict) -> tuple[list[str], list[str]]:
    """(nombres canónicos asignados, nombres sin match) para un evento."""
    is_all, names = split_becarios(ev["becarios_raw"])
    if is_all:
        return list(becarios["nombre"]), []
    assigned, unmatched = [], []
    for name in names:
        canonical = name_map.get(normalize_text(name))
        if canonical:
            assigned.append(canonical)
        else:
            unmatched.append(name)
    return assigned, unmatched


def _horas_evento(ev: pd.Series, becario: str, verificaciones: dict | None) -> tuple[float, bool, bool]:
    """Horas a contar para (evento, becario): la verificada por su jefe si existe, si no la del Excel.

    Devuelve (horas, verificado, discrepancia). `discrepancia` es True cuando ya existe una
    verificación pero la hora cruda del Excel cambió desde entonces (alguien corrigió el
    archivo después de que el jefe firmó) — la hora verificada se sigue usando igual, esto
    es solo para poder avisar que convendría revisarla de nuevo.
    """
    if not verificaciones:
        return ev["horas"], False, False
    fecha_key = ev["fecha"].isoformat() if ev["fecha"] else ""
    registro = verificaciones.get((normalize_text(ev["evento"]), fecha_key, normalize_text(becario)))
    if registro is None:
        return ev["horas"], False, False
    horas_verificadas = float(registro["horas_verificadas"])
    discrepancia = abs(horas_verificadas - float(ev["horas"])) > 1e-9
    return horas_verificadas, True, discrepancia


def split_encargados(raw: str) -> list[str]:
    """Nombres normalizados en la celda ENCARGADO (acepta varios separados por coma)."""
    return [normalize_text(t) for t in re.split(r"[,/;\n]+", str(raw)) if normalize_text(t)]


def events_for_jefe(events: pd.DataFrame, becarios: pd.DataFrame, jefe_nombre: str) -> list[dict]:
    """Eventos donde `jefe_nombre` aparece como ENCARGADO, con sus becarios asignados."""
    if events.empty or becarios.empty:
        return []
    name_map = {normalize_text(n): n for n in becarios["nombre"]}
    jefe_norm = normalize_text(jefe_nombre)
    resultado = []
    for _, ev in events.iterrows():
        if jefe_norm not in split_encargados(ev["encargado"]):
            continue
        assigned, _sin_match = _resolve_assigned(ev, becarios, name_map)
        if not assigned:
            continue
        resultado.append({
            "evento": ev["evento"],
            "fecha": ev["fecha"],
            "lugar": ev["lugar"],
            "horas": ev["horas"],
            "becarios": assigned,
        })
    resultado.sort(key=lambda e: (e["fecha"] is None, e["fecha"] or date.max), reverse=True)
    return resultado


def compute_dashboard(parsed: ParsedSheet, verificaciones: dict | None = None) -> dict:
    """Cruza eventos × becarios × tabla de becas.

    `verificaciones`: dict (evento_norm, fecha_key, becario_norm) -> registro,
    tal como lo devuelve `verificaciones.cargar_verificaciones`. Cuando existe
    un registro para un evento×becario, su hora verificada reemplaza a la del
    Excel en la suma del becario.

    Devuelve:
      - becarios: DataFrame con meta, hechas, restantes, avance, estatus y la
        lista de eventos cubiertos por cada becario.
      - warnings: nombres sin match, becarios sin meta, horas pendientes.
      - pct_verificado: proporción de horas contadas que ya fueron confirmadas
        por un jefe.
    """
    warnings = list(parsed.warnings)
    becarios = parsed.becarios.copy()
    events = parsed.events

    # % beca -> horas requeridas
    metas = {}
    if not parsed.tabla_becas.empty:
        metas = dict(zip(parsed.tabla_becas["pct_beca"], parsed.tabla_becas["horas_requeridas"]))

    if becarios.empty:
        return {"becarios": becarios, "warnings": warnings, "pct_verificado": 0.0, "discrepancias": 0}

    name_map = {normalize_text(n): n for n in becarios["nombre"]}
    hours_done = {n: 0.0 for n in becarios["nombre"]}
    events_by_becario = {n: [] for n in becarios["nombre"]}
    unmatched = []
    horas_totales = 0.0
    horas_verificadas_total = 0.0
    discrepancias = 0

    for _, ev in events.iterrows():
        assigned, sin_match = _resolve_assigned(ev, becarios, name_map)
        for name in sin_match:
            unmatched.append((name, ev["evento"], ev["fila"]))
        for name in assigned:
            horas, verificado, discrepancia = _horas_evento(ev, name, verificaciones)
            hours_done[name] += horas
            horas_totales += horas
            if verificado:
                horas_verificadas_total += horas
            if discrepancia:
                discrepancias += 1
            events_by_becario[name].append({
                "evento": ev["evento"],
                "fecha": ev["fecha"],
                "horas": horas,
                "horas_pendientes": ev["horas_pendientes"],
                "horas_estimadas": ev["horas_estimadas"],
                "verificado": verificado,
                "discrepancia": discrepancia,
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
    pct_verificado = (horas_verificadas_total / horas_totales) if horas_totales else 0.0
    return {
        "becarios": becarios,
        "warnings": warnings,
        "pct_verificado": pct_verificado,
        "discrepancias": discrepancias,
    }


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
