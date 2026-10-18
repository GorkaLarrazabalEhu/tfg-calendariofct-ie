"""
Interfaz web (Streamlit).

Ejecutar desde la raíz del proyecto:

    streamlit run src/ui/web.py
"""

import sys

from io import BytesIO
from pathlib import Path

import streamlit as st

# Streamlit ejecuta este fichero como script suelto: se añade la raíz
# del proyecto al path para poder importar el paquete src.
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from src.analisis import Tabla, analizar_horario
from src.registros import generar_registros


TITULO = "Análisis de horario docente"

# Cómo se muestran las celdas sin valor
SIN_VALOR = 0


def main():
    st.set_page_config(page_title=TITULO, layout="wide")
    st.title(TITULO)

    pestana_horas, pestana_registros = st.tabs(
        ["Análisis de horas", "Registros GAUR"]
    )

    with pestana_horas:
        seccion_analisis()

    with pestana_registros:
        seccion_registros()


def seccion_analisis():
    """
    Sube un horario y muestra las horas por asignatura y modalidad.
    """

    archivo = st.file_uploader("Seleccionar horario", type="xlsx", key="horario")

    if archivo is None:
        return

    try:
        resultado = analizar_horario(archivo)

    except Exception as error:
        st.error(f"No se ha podido analizar el archivo.\n\n{error}")
        return

    mostrar_resultado(resultado)


def seccion_registros():
    """
    Genera la lista de registros (hoja registrosGAUR)
    a partir de uno o varios horarios, p. ej. castellano y euskera.
    """

    archivos = st.file_uploader(
        "Seleccionar horarios (castellano y/o euskera)",
        type="xlsx",
        accept_multiple_files=True,
        key="horarios",
    )

    cod_plan = st.text_input("Código del plan (COD_PLAN), p. ej. GELECT30").strip()

    if not archivos or not cod_plan:
        return

    destino = BytesIO()

    try:
        total = generar_registros(archivos, cod_plan, destino)

    except Exception as error:
        st.error(f"No se han podido generar los registros.\n\n{error}")
        return

    st.success(f"Se han generado {total} registros.")

    st.download_button(
        "Descargar registrosGAUR.xlsx",
        data=destino.getvalue(),
        file_name="registrosGAUR.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def mostrar_resultado(resultado: Tabla):
    """
    Muestra cualquier resultado tabular.

    La tabla de Streamlit permite copiar celdas y descargarla como CSV.
    """

    filas = [
        {
            columna: SIN_VALOR if valor is None else valor
            for columna, valor in zip(resultado.columnas, fila)
        }
        for fila in resultado.filas
    ]

    st.dataframe(filas, hide_index=True, width='stretch')


if __name__ == "__main__":
    main()
