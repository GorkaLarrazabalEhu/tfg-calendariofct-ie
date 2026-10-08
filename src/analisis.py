from dataclasses import dataclass

from src.excel import leer_horario
from src.horas import ResumenHoras, calcular_horas


@dataclass
class Tabla:
    """
    Resultado listo para mostrar.

    Las celdas sin valor son None; la interfaz decide
    cómo representarlas.
    """

    columnas: list[str]
    filas: list[list]


def analizar_horario(ruta) -> Tabla:
    """
    Punto de entrada del análisis.

    La interfaz solo necesita llamar a esta función.
    """

    return construir_tabla(calcular_horas(leer_horario(ruta)))


def construir_tabla(resumen: ResumenHoras) -> Tabla:
    """
    Convierte los cálculos en una tabla lista para mostrar,
    con las asignaturas ordenadas por abreviatura.
    """

    columnas = [
        "Código",
        "Asignatura",
        *(modalidad.etiqueta for modalidad in resumen.modalidades),
        "Total",
    ]

    filas = []

    asignaturas = sorted(
        resumen.horas.items(),
        key=lambda elemento: elemento[0].abreviatura,
    )

    for asignatura, horas in asignaturas:

        filas.append(
            [
                asignatura.codigo,
                asignatura.abreviatura,
                *(horas.get(modalidad.clave) for modalidad in resumen.modalidades),
                sum(horas.values()),
            ]
        )

    return Tabla(columnas, filas)
