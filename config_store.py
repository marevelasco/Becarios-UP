"""Configuración chica por semestre — hoy solo la fecha límite para el indicador de urgencia.

Independiente del Excel y de las verificaciones. Misma conexión Postgres que ya usan
`verificaciones.py` y `excel_store.py`.
"""

from __future__ import annotations

from datetime import date

import streamlit as st
from sqlalchemy import text

_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS config_semestre (
    semestre TEXT PRIMARY KEY,
    fecha_limite DATE,
    actualizado_en TIMESTAMPTZ NOT NULL DEFAULT now()
);
"""


def _ensure_schema(conn) -> None:
    with conn.session as session:
        session.execute(text(_SCHEMA_SQL))
        session.commit()


@st.cache_resource(show_spinner=False)
def get_connection():
    """Conexión cacheada a Postgres. Lanza si `st.secrets` no la tiene configurada."""
    conn = st.connection("postgresql", type="sql")
    _ensure_schema(conn)
    return conn


def cargar_fecha_limite(conn, semestre: str) -> date | None:
    with conn.session as session:
        row = session.execute(
            text("SELECT fecha_limite FROM config_semestre WHERE semestre = :semestre"),
            {"semestre": semestre},
        ).fetchone()
    return row[0] if row and row[0] else None


def guardar_fecha_limite(conn, semestre: str, fecha_limite: date) -> None:
    with conn.session as session:
        session.execute(
            text(
                """
                INSERT INTO config_semestre (semestre, fecha_limite, actualizado_en)
                VALUES (:semestre, :fecha_limite, now())
                ON CONFLICT (semestre) DO UPDATE SET
                    fecha_limite = EXCLUDED.fecha_limite, actualizado_en = now()
                """
            ),
            {"semestre": semestre, "fecha_limite": fecha_limite},
        )
        session.commit()
