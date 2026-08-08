"""Genera el reporte descargable (Excel o PDF) del cumplimiento de horas de un semestre."""

from __future__ import annotations

import io

import pandas as pd
from fpdf import FPDF

_COLUMNAS = {
    "nombre": "Nombre",
    "rol": "Rol",
    "pct_beca": "% Beca",
    "meta": "Meta (h)",
    "hechas": "Horas hechas",
    "restantes": "Restantes (h)",
    "avance": "Avance %",
    "estatus": "Estatus",
    "eval_360": "Eval. 360",
}

# Ancho de cada columna en el PDF (mm) — suman ~198mm, caben cómodas en A4 horizontal (297mm).
_ANCHOS_PDF = [48, 24, 15, 18, 22, 22, 18, 22, 22]


def _tabla_reporte(becarios: pd.DataFrame) -> pd.DataFrame:
    tabla = becarios.copy()
    if "avance" in tabla:
        tabla["avance"] = (tabla["avance"] * 100).round(1)
    tabla = tabla[list(_COLUMNAS.keys())].rename(columns=_COLUMNAS)
    return tabla.sort_values("Nombre")


def generar_reporte_excel(becarios: pd.DataFrame, semestre: str) -> bytes:
    """Tabla de cumplimiento (la misma que se ve en pantalla) como .xlsx en memoria."""
    tabla = _tabla_reporte(becarios)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        hoja = (semestre or "Reporte")[:31]  # límite de Excel para nombres de hoja
        tabla.to_excel(writer, sheet_name=hoja, index=False)
    return buffer.getvalue()


def _latin1(texto: str) -> str:
    """Las fuentes base de fpdf2 (Helvetica) solo soportan latin-1 — sustituye lo que no entre
    (ej. rayas largas "—") en vez de tronar."""
    return str(texto).encode("latin-1", "replace").decode("latin-1")


def generar_reporte_pdf(becarios: pd.DataFrame, semestre: str) -> bytes:
    """Tabla de cumplimiento como PDF (una hoja, horizontal), lista para imprimir o compartir."""
    tabla = _tabla_reporte(becarios)

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 9, _latin1("Becarios UP - Cumplimiento de horas de servicio"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 7, _latin1(f"Semestre: {semestre}"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    def fila(valores, negrita):
        pdf.set_font("Helvetica", "B" if negrita else "", 8)
        for valor, ancho in zip(valores, _ANCHOS_PDF):
            pdf.cell(ancho, 7, _latin1(valor), border=1)
        pdf.ln()

    fila(tabla.columns.tolist(), negrita=True)
    for _, row in tabla.iterrows():
        fila(row.tolist(), negrita=False)

    return bytes(pdf.output())
