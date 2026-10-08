from dataclasses import dataclass


@dataclass(frozen=True)
class Asignatura:
    """
    Fila de la tabla de asignaturas del horario.

    Ejemplo:

        EM.II | 26643 | Elektromagnetismoa II
    """

    codigo: str
    abreviatura: str
    nombre: str


@dataclass(frozen=True)
class Clase:
    """
    Una sesión del horario.

    Ejemplo: la celda «EM.II (GA1) <1.9>» del lunes a las 8:40
    de la semana 4.

    curso y grupo salen de la cabecera de cada semana
    («Kurtsoa/Curso: 3   Taldea/Grupo: 31»).
    """

    asignatura: Asignatura
    tipo: str
    aula: str | None
    semana: int
    dia: str
    franja: str
    curso: str | None = None
    grupo: str | None = None


@dataclass(frozen=True)
class Modalidad:
    """
    Modalidad docente (T, GA, S, GL, GO...).

    clave:    modalidad sin número de grupo (GA).
    etiqueta: texto que se muestra en la tabla (GA, o GA1 si
              solamente existe ese grupo).
    """

    clave: str
    etiqueta: str
