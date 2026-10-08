import re

from collections import Counter, defaultdict
from dataclasses import dataclass

from src.modelos import Asignatura, Clase, Modalidad


ORDEN_MODALIDADES = ("T", "GA", "S", "GL", "GO")

POSICION_MODALIDAD = {
    modalidad: indice
    for indice, modalidad in enumerate(ORDEN_MODALIDADES)
}


@dataclass
class ResumenHoras:
    """
    modalidades: modalidades presentes en el horario, ya ordenadas.
    horas:       horas de cada asignatura por clave de modalidad.
    """

    modalidades: list[Modalidad]
    horas: dict[Asignatura, dict[str, int]]


def separar_tipo(tipo):
    """
    Separa la modalidad del grupo.

    Ejemplos:

        T   -> ("T", "T")
        GA1 -> ("GA", "GA1")
        S1  -> ("S", "S1")
        GL2 -> ("GL", "GL2")
    """

    tipo = tipo.strip().upper()

    coincidencia = re.fullmatch(r"([A-Z]+)(\d+)?", tipo)

    if coincidencia is None:
        return tipo, tipo

    return coincidencia.group(1), tipo


def ordenar_modalidades(modalidades):
    """
    Orden deseado: T, GA, S, GL, GO.

    Si en el futuro aparece otra modalidad,
    se coloca después automáticamente.
    """

    return sorted(
        modalidades,
        key=lambda modalidad: (
            POSICION_MODALIDAD.get(modalidad, len(ORDEN_MODALIDADES)),
            modalidad,
        ),
    )


def calcular_horas(clases: list[Clase]) -> ResumenHoras:
    """
    Calcula las horas por asignatura y modalidad.

    Los grupos paralelos NO se suman:

        GL1 = 13
        GL2 = 13

        GL = 13

    Lo mismo ocurre con GA1/GA2, GO1/GO2, etc.
    """

    # asignatura -> modalidad -> grupo -> número de sesiones
    sesiones = defaultdict(lambda: defaultdict(Counter))

    # modalidad -> grupos que aparecen en todo el horario
    grupos_modalidad = defaultdict(set)

    for clase in clases:

        modalidad, grupo = separar_tipo(clase.tipo)

        sesiones[clase.asignatura][modalidad][grupo] += 1
        grupos_modalidad[modalidad].add(grupo)

    horas = {
        asignatura: {
            modalidad: max(grupos.values())
            for modalidad, grupos in modalidades.items()
        }
        for asignatura, modalidades in sesiones.items()
    }

    return ResumenHoras(
        modalidades=[
            Modalidad(modalidad, etiqueta_modalidad(modalidad, grupos_modalidad[modalidad]))
            for modalidad in ordenar_modalidades(grupos_modalidad)
        ],
        horas=horas,
    )


def etiqueta_modalidad(modalidad, grupos):
    """
    Decide cómo se mostrará una modalidad.

        Solo existe GA1      -> GA1
        Existen GA1 y GA2    -> GA
    """

    if len(grupos) == 1:
        return next(iter(grupos))

    return modalidad
