import re

from collections import defaultdict
from dataclasses import dataclass
from datetime import timedelta
from functools import lru_cache

from openpyxl import Workbook
from openpyxl.styles import Font

from src.excel import PATRON_HORA, leer_horario
from src.modelos import Clase


# --------------------------------------------------
# FORMATO DE LA HOJA
# --------------------------------------------------

NOMBRE_HOJA = "registrosGAUR"

COLUMNAS = [
    "COD_PLAN",
    "CURSO",
    "COD_ASIGNATURA",
    "DESC_ASIG",
    "DESC_ASIG_EUSK",
    "DESC_GRUPO_ASIG",
    "TIPO_GRUPO",
    "SECUENCIA_TIPO_GRUPO",
    None,  # columna vacía
    "SEMANA_INICIO",
    "SEMANA_FIN",
    "FASE",
    "PERIODO",
    "DIA_SEMANA",
    "HORA_INICIO",
    "HORA_FIN",
]

ANCHOS_COLUMNA = {"D": 38, "E": 24, "H": 7}

NUMERO_DIA = {"Lunes": 1, "Martes": 2, "Miércoles": 3, "Jueves": 4, "Viernes": 5}

# Dos sesiones seguidas se juntan en un bloque si entre ellas
# hay como mucho este descanso (incluye el de 11:30 a 12:00).
DESCANSO_MAXIMO = timedelta(minutes=30)

# GAUR no usa ciclos de más de cinco semanas
PERIODO_MAXIMO = 5


@dataclass(frozen=True)
class Registro:
    """
    Una fila de la hoja registrosGAUR.

    Las semanas en las que se imparte son las de
    [semana_inicio, semana_fin] cuya posición dentro del
    ciclo de «periodo» semanas coincide con «fase»:

        (semana - semana_inicio) % periodo == fase - 1
    """

    cod_plan: str
    curso: str
    cod_asignatura: str
    desc_asig: str | None
    desc_asig_eusk: str | None
    grupo: str
    tipo_grupo: str
    secuencia: int
    semana_inicio: int
    semana_fin: int
    fase: int
    periodo: int
    dia: int
    hora_inicio: str
    hora_fin: str


# --------------------------------------------------
# PUNTO DE ENTRADA
# --------------------------------------------------

def generar_registros(rutas, cod_plan, destino):
    """
    Lee uno o varios horarios (p. ej. castellano y euskera)
    y guarda la lista de registros en formato GAUR.

    Devuelve el número de registros escritos.
    """

    clases = [clase for ruta in rutas for clase in leer_horario(ruta)]

    registros = construir_registros(clases, cod_plan)

    guardar_registros(registros, destino)

    return len(registros)


def construir_registros(clases: list[Clase], cod_plan) -> list[Registro]:
    """
    Agrupa las clases por grupo, día y franja, y comprime
    las semanas en filas con fase y periodo.
    """

    nombres = nombres_asignaturas(clases)

    # (grupo, día, hora_inicio, hora_fin) -> semana -> actividades
    huecos = defaultdict(lambda: defaultdict(set))

    # grupo -> curso
    cursos = {}

    for bloque in juntar_sesiones(clases):

        grupo, semana, dia, inicio, fin, actividad = bloque

        huecos[(grupo, dia, inicio, fin)][semana].add(actividad)

    # grupo -> días con alguna clase
    dias_con_clase = defaultdict(set)

    for clase in clases:
        cursos[clase.grupo] = clase.curso
        dias_con_clase[clase.grupo].add((clase.semana, clase.dia))

    ultima_semana = max((clase.semana for clase in clases), default=0)

    registros = []

    for (grupo, dia, inicio, fin), semanas in huecos.items():

        # Un día sin ninguna clase es festivo: no rompe
        # la repetición de las semanas (None = cualquier cosa).
        secuencia = tuple(
            frozenset(semanas.get(semana, ()))
            if (semana, dia) in dias_con_clase[grupo] else None
            for semana in range(1, ultima_semana + 1)
        )

        for actividad, semana_inicio, semana_fin, fase, periodo in comprimir_semanas(secuencia):

            codigo, tipo_grupo, numero = actividad
            desc_asig, desc_asig_eusk = nombres[codigo]

            registros.append(
                Registro(
                    cod_plan=cod_plan,
                    curso=cursos[grupo],
                    cod_asignatura=codigo,
                    desc_asig=desc_asig,
                    desc_asig_eusk=desc_asig_eusk,
                    grupo=grupo,
                    tipo_grupo=tipo_grupo,
                    secuencia=numero,
                    semana_inicio=semana_inicio,
                    semana_fin=semana_fin,
                    fase=fase,
                    periodo=periodo,
                    dia=NUMERO_DIA[dia],
                    hora_inicio=formatear_hora(inicio),
                    hora_fin=formatear_hora(fin),
                )
            )

    registros.sort(
        key=lambda registro: (
            registro.curso or "",
            registro.cod_asignatura,
            registro.grupo,
            registro.tipo_grupo,
            registro.secuencia,
            registro.semana_inicio,
            registro.fase,
            registro.dia,
            registro.hora_inicio,
        )
    )

    return registros


# --------------------------------------------------
# NOMBRES DE ASIGNATURA
# --------------------------------------------------

def es_grupo_euskera(grupo):
    """
    En la UPV/EHU los grupos 31-60 son de euskera.
    """

    return grupo is not None and 31 <= int(grupo) <= 60


def nombres_asignaturas(clases):
    """
    Devuelve código -> (nombre en castellano, nombre en euskera).

    Cada nombre se toma del horario del idioma correspondiente;
    si no se ha cargado ese horario, queda vacío.
    """

    castellano = {}
    euskera = {}

    for clase in clases:

        asignatura = clase.asignatura

        if es_grupo_euskera(clase.grupo):
            euskera.setdefault(asignatura.codigo, asignatura.nombre)
        else:
            castellano.setdefault(asignatura.codigo, asignatura.nombre)

    return {
        codigo: (castellano.get(codigo), euskera.get(codigo))
        for codigo in castellano.keys() | euskera.keys()
    }


# --------------------------------------------------
# BLOQUES HORARIOS
# --------------------------------------------------

def horas_franja(franja):
    """
    "8:40\\n9:30" -> (8:40, 9:30) como timedelta.
    """

    inicio, fin = PATRON_HORA.findall(franja)[:2]

    return convertir_hora(inicio), convertir_hora(fin)


def convertir_hora(texto):
    horas, minutos = texto.split(":")
    return timedelta(hours=int(horas), minutes=int(minutos))


def hora_gaur(hora):
    """
    GAUR usa franjas de una hora que empiezan en punto
    o a y media: 8:40 -> 8:30, 15:55 -> 16:00, 18:50 -> 19:00.
    """

    media_hora = timedelta(minutes=30)

    return round(hora / media_hora) * media_hora


def formatear_hora(hora):
    minutos = int(hora.total_seconds()) // 60
    return f"{minutos // 60:02d}:{minutos % 60:02d}:00"


def separar_grupo(tipo):
    """
    Separa el tipo de grupo de su secuencia.

        T   -> ("T", 0)
        GA1 -> ("GA", 1)
        S   -> ("S", 1)
    """

    coincidencia = re.fullmatch(r"([A-Za-z]+)(\d+)?", tipo.strip())

    if coincidencia is None:
        return tipo.strip(), 1

    tipo_grupo, numero = coincidencia.groups()
    tipo_grupo = tipo_grupo.upper()

    if numero is not None:
        return tipo_grupo, int(numero)

    return tipo_grupo, 0 if tipo_grupo == "T" else 1


def juntar_sesiones(clases):
    """
    Junta las sesiones seguidas de la misma actividad en un día
    en un único bloque.

    Ejemplo: GL1 de 15:00, 15:55 y 17:00  ->  15:00-18:00.

    Devuelve tuplas (grupo, semana, día, inicio, fin, actividad),
    con las horas ya en franjas de GAUR.
    """

    sesiones = defaultdict(list)

    for clase in clases:

        actividad = (clase.asignatura.codigo, *separar_grupo(clase.tipo))

        sesiones[(clase.grupo, clase.semana, clase.dia, actividad)].append(
            horas_franja(clase.franja)
        )

    bloques = []

    for (grupo, semana, dia, actividad), franjas in sesiones.items():

        franjas.sort()

        inicio, fin = franjas[0]
        ultimo_inicio = inicio

        for siguiente_inicio, siguiente_fin in franjas[1:]:

            if siguiente_inicio - fin > DESCANSO_MAXIMO:
                bloques.append((grupo, semana, dia, inicio, ultimo_inicio, actividad))
                inicio = siguiente_inicio

            ultimo_inicio, fin = siguiente_inicio, siguiente_fin

        bloques.append((grupo, semana, dia, inicio, ultimo_inicio, actividad))

    # En GAUR el bloque acaba una hora después
    # de que empiece su última franja.
    return [
        (grupo, semana, dia, hora_gaur(inicio), hora_gaur(ultimo_inicio) + timedelta(hours=1), actividad)
        for grupo, semana, dia, inicio, ultimo_inicio, actividad in bloques
    ]


# --------------------------------------------------
# COMPRESIÓN DE SEMANAS
# --------------------------------------------------

def comprimir_semanas(secuencia):
    """
    Recibe, para un mismo hueco (grupo, día y horas), las
    actividades de cada semana:

        secuencia[0] -> actividades de la semana 1, ...

    (None si ese día es festivo) y la divide en tramos que
    se repiten cada «periodo» semanas, buscando el menor
    número de filas posible.

    Ejemplo (TAP, jueves 10:30):

        semanas 1-5:  T T T T T          -> T   1-5   fase 1 periodo 1
        semanas 6-15: S T S T S T ...    -> S   6-15  fase 1 periodo 2
                                            T   6-15  fase 2 periodo 2

    Devuelve tuplas (actividad, semana_inicio, semana_fin, fase, periodo).
    """

    n = len(secuencia)

    @lru_cache(maxsize=None)
    def mejor(desde):
        """
        Mejor forma de cubrir las semanas desde «desde» hasta el final.

        Coste, por orden de importancia: filas, suma de periodos,
        número de tramos y, en empate, alargar el tramo actual.
        """

        if desde == n:
            return (0, 0, 0), ()

        # Semana vacía o festiva: se salta sin coste
        if not secuencia[desde]:
            return mejor(desde + 1)

        opciones = []

        for hasta in range(desde, n):

            # Los tramos empiezan y terminan en semanas con clase
            if not secuencia[hasta]:
                continue

            longitud = hasta - desde + 1

            for periodo in range(1, min(PERIODO_MAXIMO, longitud) + 1):

                # Un periodo mayor que 1 debe verse casi dos veces
                # (p. ej. GO - GL - GO con periodo 2)
                if periodo > 1 and longitud < 2 * periodo - 1:
                    continue

                fases = valores_por_fase(desde, hasta, periodo)

                if fases is None:
                    continue

                filas = tuple(
                    fila
                    for fase, (actividades, semanas_reales) in enumerate(fases)
                    for fila in filas_de_fase(
                        actividades, semanas_reales, desde, hasta, fase, periodo
                    )
                )

                (coste_filas, suma_periodos, tramos), resto = mejor(hasta + 1)

                coste = (coste_filas + len(filas), suma_periodos + periodo, tramos + 1)

                opciones.append(((coste, -longitud), (coste, filas + resto)))

        return min(opciones, key=lambda opcion: opcion[0])[1]

    def filas_de_fase(actividades, semanas_reales, desde, hasta, fase, periodo):
        """
        Filas de una fase del tramo.

        Si en un tramo periódico la fase solo tiene una semana
        real, se escribe como esa semana suelta:

            semanas 7-9:  GO  GL  GO  ->  GO  7-9  fase 1 periodo 2
                                          GL  8-8  fase 1 periodo 1
        """

        for actividad in sorted(actividades):

            if periodo > 1 and len(semanas_reales) == 1:
                semana = semanas_reales[0] + 1
                yield (actividad, semana, semana, 1, 1)

            else:
                yield (actividad, desde + 1, hasta + 1, fase + 1, periodo)

    def valores_por_fase(desde, hasta, periodo):
        """
        Para cada fase, (actividades, semanas no festivas), si todas
        las semanas de una misma fase coinciden (sin contar festivos);
        si no, None.
        """

        fases = [None] * periodo
        semanas_reales = [[] for _ in range(periodo)]

        for semana in range(desde, hasta + 1):

            valor = secuencia[semana]

            if valor is None:
                continue

            fase = (semana - desde) % periodo
            semanas_reales[fase].append(semana)

            if fases[fase] is None:
                fases[fase] = valor

            elif fases[fase] != valor:
                return None

        return [
            (valor or frozenset(), reales)
            for valor, reales in zip(fases, semanas_reales)
        ]

    _, filas = mejor(0)

    return list(filas)


# --------------------------------------------------
# ESCRITURA
# --------------------------------------------------

def guardar_registros(registros: list[Registro], destino):

    libro = Workbook()
    hoja = libro.active
    hoja.title = NOMBRE_HOJA

    hoja.append(COLUMNAS)

    for celda in hoja[1]:
        celda.font = Font(bold=True)

    for fila, registro in enumerate(registros, start=2):

        hoja.append(
            [
                registro.cod_plan,
                registro.curso,
                registro.cod_asignatura,
                registro.desc_asig,
                registro.desc_asig_eusk,
                registro.grupo,
                registro.tipo_grupo,
                registro.secuencia,
                f"=G{fila}&H{fila}",
                registro.semana_inicio,
                registro.semana_fin,
                registro.fase,
                registro.periodo,
                registro.dia,
                registro.hora_inicio,
                registro.hora_fin,
            ]
        )

    for columna, ancho in ANCHOS_COLUMNA.items():
        hoja.column_dimensions[columna].width = ancho

    libro.save(destino)
