"""Persistencia de las horas verificadas por cada jefe de becarios.

Independiente del Excel: guarda solo la corrección/confirmación que hace un
jefe sobre las horas de un evento×becario, en una base Postgres gratuita
(ej. Neon). El connection string vive en `st.secrets`:

    [connections.postgresql]
    url = "postgresql+psycopg2://usuario:password@host/basedatos?sslmode=require"

La tabla se crea sola la primera vez que la app corre (no hace falta correr
SQL a mano).
"""

from __future__ import annotations

import streamlit as st
from sqlalchemy import text

from data_loader import normalize_text

_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS verificaciones (
    id SERIAL PRIMARY KEY,
    semestre TEXT NOT NULL,
    evento_norm TEXT NOT NULL,
    fecha DATE,
    fecha_key TEXT NOT NULL,
    becario_norm TEXT NOT NULL,
    encargado_norm TEXT NOT NULL,
    horas_originales NUMERIC NOT NULL,
    horas_verificadas NUMERIC NOT NULL,
    estatus TEXT NOT NULL,
    verificado_por TEXT NOT NULL,
    verificado_en TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (semestre, evento_norm, fecha_key, becario_norm)
);
"""

# Tabla de solo inserción (nunca se actualiza ni se borra): deja rastro de cada
# corrección, aunque `verificaciones` solo guarde el último estado.
_SCHEMA_HISTORIAL_SQL = """
CREATE TABLE IF NOT EXISTS verificaciones_historial (
    id SERIAL PRIMARY KEY,
    semestre TEXT NOT NULL,
    evento_norm TEXT NOT NULL,
    fecha_key TEXT NOT NULL,
    becario_norm TEXT NOT NULL,
    horas_anteriores NUMERIC,
    horas_nuevas NUMERIC NOT NULL,
    estatus TEXT NOT NULL,
    verificado_por TEXT NOT NULL,
    verificado_en TIMESTAMPTZ NOT NULL DEFAULT now()
);
"""


def _ensure_schema(conn) -> None:
    with conn.session as session:
        session.execute(text(_SCHEMA_SQL))
        session.execute(text(_SCHEMA_HISTORIAL_SQL))
        session.commit()


@st.cache_resource(show_spinner=False)
def get_connection():
    """Conexión cacheada a Postgres. Lanza si `st.secrets` no la tiene configurada."""
    conn = st.connection("postgresql", type="sql")
    _ensure_schema(conn)
    return conn


def cargar_verificaciones(conn, semestre: str) -> dict[tuple[str, str, str], dict]:
    """(evento_norm, fecha_key, becario_norm) -> registro, para una hoja/semestre."""
    df = conn.query(
        "SELECT evento_norm, fecha_key, becario_norm, horas_originales, horas_verificadas, "
        "estatus, verificado_por, verificado_en FROM verificaciones WHERE semestre = :semestre",
        params={"semestre": semestre},
        ttl=0,
    )
    return {
        (row["evento_norm"], row["fecha_key"], row["becario_norm"]): row
        for row in df.to_dict("records")
    }


def guardar_verificacion(
    conn,
    *,
    semestre: str,
    evento: str,
    fecha,
    becario: str,
    encargado: str,
    horas_originales: float,
    horas_verificadas: float,
    verificado_por: str,
) -> None:
    """Inserta o actualiza la verificación de un becario en un evento (upsert), y deja
    un rastro en `verificaciones_historial` de qué cambió, quién y cuándo."""
    estatus = "confirmado" if abs(horas_verificadas - horas_originales) < 1e-9 else "corregido"
    evento_norm = normalize_text(evento)
    fecha_key = fecha.isoformat() if fecha else ""
    becario_norm = normalize_text(becario)
    with conn.session as session:
        anterior = session.execute(
            text(
                "SELECT horas_verificadas FROM verificaciones WHERE semestre = :semestre AND "
                "evento_norm = :evento_norm AND fecha_key = :fecha_key AND becario_norm = :becario_norm"
            ),
            {"semestre": semestre, "evento_norm": evento_norm, "fecha_key": fecha_key, "becario_norm": becario_norm},
        ).fetchone()
        horas_anteriores = float(anterior[0]) if anterior else None

        session.execute(
            text(
                """
                INSERT INTO verificaciones
                    (semestre, evento_norm, fecha, fecha_key, becario_norm, encargado_norm,
                     horas_originales, horas_verificadas, estatus, verificado_por, verificado_en)
                VALUES
                    (:semestre, :evento_norm, :fecha, :fecha_key, :becario_norm, :encargado_norm,
                     :horas_originales, :horas_verificadas, :estatus, :verificado_por, now())
                ON CONFLICT (semestre, evento_norm, fecha_key, becario_norm) DO UPDATE SET
                    horas_verificadas = EXCLUDED.horas_verificadas,
                    estatus = EXCLUDED.estatus,
                    verificado_por = EXCLUDED.verificado_por,
                    verificado_en = now()
                """
            ),
            {
                "semestre": semestre,
                "evento_norm": evento_norm,
                "fecha": fecha,
                "fecha_key": fecha_key,
                "becario_norm": becario_norm,
                "encargado_norm": normalize_text(encargado),
                "horas_originales": horas_originales,
                "horas_verificadas": horas_verificadas,
                "estatus": estatus,
                "verificado_por": verificado_por,
            },
        )
        session.execute(
            text(
                """
                INSERT INTO verificaciones_historial
                    (semestre, evento_norm, fecha_key, becario_norm, horas_anteriores,
                     horas_nuevas, estatus, verificado_por, verificado_en)
                VALUES
                    (:semestre, :evento_norm, :fecha_key, :becario_norm, :horas_anteriores,
                     :horas_nuevas, :estatus, :verificado_por, now())
                """
            ),
            {
                "semestre": semestre,
                "evento_norm": evento_norm,
                "fecha_key": fecha_key,
                "becario_norm": becario_norm,
                "horas_anteriores": horas_anteriores,
                "horas_nuevas": horas_verificadas,
                "estatus": estatus,
                "verificado_por": verificado_por,
            },
        )
        session.commit()


def cargar_historial(conn, semestre: str, evento_norm: str, fecha_key: str, becario_norm: str) -> list[dict]:
    """Cambios registrados para un evento×becario, del más antiguo al más reciente."""
    df = conn.query(
        "SELECT horas_anteriores, horas_nuevas, estatus, verificado_por, verificado_en "
        "FROM verificaciones_historial WHERE semestre = :semestre AND evento_norm = :evento_norm "
        "AND fecha_key = :fecha_key AND becario_norm = :becario_norm ORDER BY verificado_en ASC",
        params={
            "semestre": semestre, "evento_norm": evento_norm,
            "fecha_key": fecha_key, "becario_norm": becario_norm,
        },
        ttl=0,
    )
    return df.to_dict("records")
