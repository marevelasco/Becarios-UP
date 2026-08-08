"""Guarda el último Excel subido en Postgres, para que no se pierda al refrescar.

Independiente de las verificaciones de horas: una sola fila que siempre se
sobrescribe (sin historial) con el Excel más reciente que alguien haya subido
desde el dashboard. Usa la misma conexión que `verificaciones.py`.
"""

from __future__ import annotations

import streamlit as st
from sqlalchemy import text

_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS excel_actual (
    id SMALLINT PRIMARY KEY DEFAULT 1 CHECK (id = 1),
    nombre TEXT NOT NULL,
    contenido BYTEA NOT NULL,
    subido_en TIMESTAMPTZ NOT NULL DEFAULT now()
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


def cargar_excel(conn) -> tuple[str, bytes] | None:
    """(nombre, contenido) del último Excel guardado, o None si no hay ninguno todavía.

    Usa `conn.session` (no `conn.query`) a propósito: `.query()` pasa el resultado por el
    caché de Streamlit, que intenta *picklear* el DataFrame — y falla con columnas BYTEA
    grandes (memoryview no es picklable ahí), lanzando una excepción que quedaba silenciada
    más arriba. Con `.session` leemos directo, sin caché.
    """
    with conn.session as session:
        row = session.execute(text("SELECT nombre, contenido FROM excel_actual WHERE id = 1")).fetchone()
    if row is None:
        return None
    nombre, contenido = row
    return nombre, bytes(contenido)


def fecha_actualizacion(conn):
    """Fecha/hora de la última vez que se guardó un Excel, o None."""
    with conn.session as session:
        row = session.execute(text("SELECT subido_en FROM excel_actual WHERE id = 1")).fetchone()
    return None if row is None else row[0]


def guardar_excel(conn, nombre: str, contenido: bytes) -> None:
    """Sobrescribe el Excel guardado (siempre hay una sola fila, la más reciente)."""
    with conn.session as session:
        session.execute(
            text(
                """
                INSERT INTO excel_actual (id, nombre, contenido, subido_en)
                VALUES (1, :nombre, :contenido, now())
                ON CONFLICT (id) DO UPDATE SET
                    nombre = EXCLUDED.nombre,
                    contenido = EXCLUDED.contenido,
                    subido_en = now()
                """
            ),
            {"nombre": nombre, "contenido": contenido},
        )
        session.commit()
