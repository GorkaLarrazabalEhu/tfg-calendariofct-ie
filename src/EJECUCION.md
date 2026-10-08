# Línea de ejecución del programa

Este documento recorre el programa desde que se arranca hasta que muestra
un resultado o escribe un fichero, explicando cada módulo, clase y función
en el orden en que se ejecutan.

## Índice

1. [Visión general](#1-visión-general)
2. [Arranque: `main.py`](#2-arranque-mainpy)
3. [Interfaz: `ui/interfaz.py`](#3-interfaz-uiinterfazpy)
4. [Modelos de datos: `modelos.py`](#4-modelos-de-datos-modelospy)
5. [Lectura del Excel: `excel.py`](#5-lectura-del-excel-excelpy)
6. [Flujo A — Análisis de horas](#6-flujo-a--análisis-de-horas)
   - [`analisis.py`](#61-analisispy)
   - [`horas.py`](#62-horaspy)
7. [Flujo B — Registros GAUR: `registros.py`](#7-flujo-b--registros-gaur-registrospy)
8. [Resumen de llamadas](#8-resumen-de-llamadas)

---

## 1. Visión general

El programa es una aplicación de escritorio (Tkinter) que lee horarios de la
Facultad de Ciencia y Tecnología en formato Excel (`.xlsx`) y ofrece dos
funciones:

| Botón | Qué hace | Resultado |
|---|---|---|
| **Seleccionar horario** | Cuenta las horas de cada asignatura por modalidad docente (T, GA, S, GL, GO). | Tabla en pantalla (copiable a Excel con `Ctrl+C`). |
| **Generar registros GAUR** | Convierte uno o varios horarios en la lista de registros que usa GAUR. | Fichero `registrosGAUR.xlsx`. |

Estructura de módulos y dependencias:

```
main.py
  └── ui/interfaz.py
        ├── analisis.py ──┬── excel.py ── modelos.py
        │                 └── horas.py ── modelos.py
        └── registros.py ─┬── excel.py
                          └── modelos.py
```

Ambos flujos comparten la misma lectura del Excel (`excel.leer_horario`), que
convierte el fichero en una lista de objetos `Clase`. A partir de ahí cada
flujo procesa esa lista a su manera.

Única dependencia externa: `openpyxl` (lectura y escritura de `.xlsx`).

**Cómo se ejecuta** (desde la raíz del repositorio, porque los imports usan el
prefijo `src.`):

```
python -m src.main
```

---

## 2. Arranque: `main.py`

### `main()`

Único punto de entrada. Llama a `iniciar_interfaz()` y no hace nada más.
El bloque `if __name__ == "__main__"` permite ejecutarlo como script o módulo.

---

## 3. Interfaz: `ui/interfaz.py`

La interfaz solo se encarga de pedir datos al usuario (ficheros, código de
plan) y de mostrar resultados. No sabe cómo se calculan las horas ni cómo se
generan los registros: delega en `analisis.py` y `registros.py`.

### Constantes

| Constante | Valor | Uso |
|---|---|---|
| `TITULO` | `"Análisis de horario docente"` | Título de la ventana y cabecera. |
| `SIN_VALOR` | `"0"` | Cómo se muestra una celda `None` en la tabla. |
| `ANCHO_MINIMO` / `ANCHO_MAXIMO` | `70` / `250` | Límites en píxeles del ancho de cada columna. |

### `iniciar_interfaz()`

1. Crea la ventana raíz `tk.Tk()`.
2. Construye la `VentanaPrincipal` sobre ella.
3. Entra en `root.mainloop()`: a partir de aquí el programa queda **esperando
   eventos** (clics en botones, `Ctrl+C`). Todo lo demás se ejecuta como
   respuesta a esos eventos.

### Clase `VentanaPrincipal`

#### `__init__(self, root)`

Monta la ventana (1000×600):

1. Título y etiqueta grande con `TITULO`.
2. `self.etiqueta_archivo`: muestra la ruta del horario cargado
   (al principio, *"Ningún archivo seleccionado"*).
3. `self.tabla`: la tabla de resultados, creada con `_crear_tabla`. Se asocia
   `Ctrl+C` a `copiar_tabla`.
4. Dos botones:
   - **Seleccionar horario** → `cargar_horario` (flujo A).
   - **Generar registros GAUR** → `generar_registros` (flujo B).
5. `self.cod_plan = ""`: recuerda el último código de plan introducido para
   proponerlo la siguiente vez.

#### `_crear_tabla(self, root)`

Crea un `ttk.Treeview` (en modo solo cabeceras) dentro de un marco con barras
de desplazamiento vertical y horizontal, colocado con `grid` para que se
estire al redimensionar la ventana. Devuelve el `Treeview`.

#### `cargar_horario(self)` — inicio del flujo A

1. Abre un diálogo para elegir **un** `.xlsx`. Si se cancela, termina.
2. Llama a `analizar_horario(ruta)` (ver [§6](#6-flujo-a--análisis-de-horas)).
3. Si salta cualquier excepción (formato inesperado, asignatura desconocida…),
   la muestra en un `messagebox` de error y termina.
4. Si todo va bien, pone la ruta en `etiqueta_archivo` y llama a
   `mostrar_resultado`.

#### `generar_registros(self)` — inicio del flujo B

1. Abre un diálogo para elegir **uno o varios** `.xlsx` (normalmente el
   horario en castellano y el de euskera).
2. Pide el código de plan (`COD_PLAN`, p. ej. `GELECT30`), proponiendo el
   último usado. Si se cancela o queda vacío, termina.
3. Pide dónde guardar el resultado (por defecto `registrosGAUR.xlsx`).
4. Llama a `registros.generar_registros(rutas, cod_plan, destino)`
   (ver [§7](#7-flujo-b--registros-gaur-registrospy)).
5. Muestra un error si falla, o un aviso con el número de registros escritos.

#### `mostrar_resultado(self, resultado: Tabla)`

Pinta cualquier `Tabla` (columnas + filas) sin conocer su significado:

1. Borra las filas anteriores.
2. Convierte cada valor a texto; los `None` pasan a `SIN_VALOR` (`"0"`).
3. Define las columnas. El ancho de cada una se calcula a partir del texto más
   largo de esa columna (`longitud * 9 + 20` px), limitado entre
   `ANCHO_MINIMO` y `ANCHO_MAXIMO`.
4. Inserta las filas.

#### `copiar_tabla(self)`

Se ejecuta con `Ctrl+C` sobre la tabla. Construye un texto separado por
tabuladores (formato que Excel pega en celdas):

- Primera línea: cabeceras.
- Después, las filas seleccionadas o, si no hay ninguna seleccionada, todas.
- Los `"0"` que representan "sin valor" se copian como celda vacía.

Lo deja en el portapapeles.

---

## 4. Modelos de datos: `modelos.py`

Tres `dataclass` inmutables (`frozen=True`), por lo que se pueden usar como
claves de diccionario y en conjuntos.

### `Asignatura`

Una fila de la tabla de asignaturas que hay al pie de cada hoja del horario.

| Campo | Ejemplo |
|---|---|
| `codigo` | `"26643"` |
| `abreviatura` | `"EM.II"` |
| `nombre` | `"Elektromagnetismoa II"` |

### `Clase`

Una sesión concreta del horario: una celda de la cuadrícula en una semana,
un día y una franja.

| Campo | Ejemplo | Origen |
|---|---|---|
| `asignatura` | `Asignatura(...)` | Abreviatura de la celda buscada en la tabla de asignaturas. |
| `tipo` | `"GA1"` | Texto entre paréntesis de la celda. |
| `aula` | `"1.9"` o `None` | Texto entre `< >` de la celda. |
| `semana` | `4` | Cabecera `Astea/Semana: 4`. |
| `dia` | `"Lunes"` | Columna de la celda. |
| `franja` | `"8:40\n9:30"` | Primera celda de la fila. |
| `curso` | `"3"` | Cabecera `Kurtsoa/Curso: 3`. |
| `grupo` | `"31"` | Cabecera `Taldea/Grupo: 31` (con dos cifras). |

### `Modalidad`

Modalidad docente tal como se muestra en la tabla de horas.

- `clave`: modalidad sin número de grupo (`"GA"`).
- `etiqueta`: texto de la columna (`"GA"`, o `"GA1"` si solo existe ese grupo).

---

## 5. Lectura del Excel: `excel.py`

Convierte un fichero de horario en una `list[Clase]`. Lo usan los dos flujos.

### Formato esperado de cada hoja

```
Kurtsoa/Curso: 3  Taldea/Grupo: 31 ... Astea/Semana: 1   ← cabecera de semana
            | Lunes          | Martes | Miércoles | ...  ← fila de días
8:40 9:30   | EM.II (T) <1.9>| ...                       ← franjas horarias
9:40 10:30  | ...
...
Kurtsoa/Curso: 3  Taldea/Grupo: 31 ... Astea/Semana: 2   ← siguiente semana
...
ASIGNATURA/IRAKASGAIA ...                                ← tabla de asignaturas
EM.II       | 26643 | Elektromagnetismoa II
```

### Patrones (expresiones regulares)

| Patrón | Reconoce |
|---|---|
| `PATRON_CLASE` | `ABREV (TIPO) <AULA>`; el aula es opcional. |
| `PATRON_SEMANA` | `Astea/Semana: N` |
| `PATRON_CURSO` | `Kurtsoa/Curso: N` |
| `PATRON_GRUPO` | `Taldea/Grupo: N` |
| `PATRON_HORA` | `H:MM` o `HH:MM` (también lo usa `registros.py`). |

`INICIO_TABLA_ASIGNATURAS = "ASIGNATURA/IRAKASGAIA"` marca el final de la
cuadrícula. `DIAS` relaciona el texto en minúsculas con el nombre normalizado
del día.

### `leer_horario(ruta) -> list[Clase]` — punto de entrada

1. Abre el libro con `load_workbook(ruta, data_only=True)` (`data_only` hace
   que las fórmulas se lean como su valor calculado).
2. Llama a `leer_hoja` para **cada hoja** y concatena todas las clases.
3. Cierra el libro siempre (`finally`), incluso si hay un error.

### `leer_hoja(hoja) -> list[Clase]`

Recorre la hoja fila a fila como una pequeña máquina de estados:

1. Primero lee la tabla de asignaturas completa con `leer_asignaturas(hoja)`,
   para poder traducir abreviaturas a `Asignatura`.
2. Estado inicial: `semana = None` (fuera de cualquier bloque semanal).
3. Para cada fila:
   - **Si contiene `Astea/Semana`** (`obtener_semana`): empieza un bloque.
     Guarda el número de semana, extrae curso y grupo
     (`obtener_curso_grupo`) y **consume la fila siguiente** con
     `next(filas)` para localizar la columna de cada día
     (`obtener_columnas_dias`).
   - **Si la primera celda es `ASIGNATURA/IRAKASGAIA`**: se ha llegado a la
     tabla de asignaturas; `semana = None` para no leer nada más como clases.
   - **Si estamos dentro de una semana y la primera celda es una franja
     horaria** (`es_franja_horaria`): lee las clases de esa fila con
     `leer_franja`.
   - En cualquier otro caso, la fila se ignora.

### `leer_franja(fila, columnas_dias, semana, asignaturas, curso, grupo)`

Para una fila horaria, mira la celda de cada día:

1. `interpretar_celda` separa las clases que contiene (puede haber varias).
2. Si una abreviatura no está en la tabla de asignaturas, lanza
   `ValueError` (la interfaz lo mostrará como error).
3. Crea un `Clase` por cada una con la semana, día, franja, curso y grupo.

### `interpretar_celda(valor)`

Aplica `PATRON_CLASE.findall` al texto de la celda y devuelve una lista de
tuplas `(abreviatura, tipo, aula)`. Ejemplos:

```
"EM.II (T) <1.9>"              -> [("EM.II", "T", "1.9")]
"ELEK.D (GL1)"                 -> [("ELEK.D", "GL1", None)]
"MICR (GL4) <0.17>\nREGU (GL2)" -> [("MICR", "GL4", "0.17"), ("REGU", "GL2", None)]
None                           -> []
```

`findall` devuelve `""` cuando falta el aula; se convierte a `None`.

### `obtener_semana(fila)`

Busca `Astea/Semana: N` en cualquier celda de la fila. Devuelve `N` como
entero o `None` si la fila no es una cabecera de semana.

### `obtener_curso_grupo(fila)`

Busca en la cabecera el curso y el grupo. El grupo se rellena a dos cifras
(`"1"` → `"01"`), como lo espera GAUR. Devuelve `(curso, grupo)`; cualquiera
de los dos puede ser `None` si no aparece.

### `obtener_columnas_dias(fila, nombre_hoja, numero_fila)`

Devuelve `{"Lunes": índice, "Martes": índice, ...}` buscando el nombre de cada
día dentro del texto de las celdas (da igual que haya columnas vacías entre
medias). Si no encuentra los cinco días lanza `ValueError` indicando hoja y
fila, para que el usuario sepa dónde está el problema.

### `es_franja_horaria(valor)`

`True` si la celda contiene al menos dos horas (`8:40 … 9:30`).

### `es_inicio_tabla_asignaturas(valor)`

`True` si la celda contiene `ASIGNATURA/IRAKASGAIA`.

### `leer_asignaturas(hoja) -> dict[str, Asignatura]`

Recorre todas las filas (solo valores) y se queda con las que tienen la forma
`abreviatura | código numérico | nombre`. El filtro clave es
`normalizar_codigo`: si la segunda columna no es un número, la fila no es de
la tabla de asignaturas. Devuelve un diccionario indexado por abreviatura.

### `normalizar_codigo(valor)`

Convierte el código a texto: `26643` (int), `26643.0` (float) o `"26643"`
(texto) → `"26643"`. Cualquier otra cosa → `None`.

---

## 6. Flujo A — Análisis de horas

```
cargar_horario
  └── analizar_horario(ruta)
        ├── leer_horario(ruta)          → list[Clase]
        ├── calcular_horas(clases)      → ResumenHoras
        └── construir_tabla(resumen)    → Tabla
  └── mostrar_resultado(tabla)
```

### 6.1 `analisis.py`

#### Clase `Tabla`

Resultado genérico listo para mostrar: `columnas: list[str]` y
`filas: list[list]`. Las celdas sin valor son `None`; la interfaz decide cómo
representarlas.

#### `analizar_horario(ruta) -> Tabla`

Encadena las tres fases: `construir_tabla(calcular_horas(leer_horario(ruta)))`.
Es la única función que necesita conocer la interfaz para este flujo.

#### `construir_tabla(resumen) -> Tabla`

- Columnas: `Código`, `Asignatura`, una por cada modalidad (su `etiqueta`, en
  el orden de `resumen.modalidades`) y `Total`.
- Filas: una por asignatura, ordenadas por abreviatura. Cada fila contiene
  código, abreviatura, las horas de cada modalidad (`None` si la asignatura no
  tiene esa modalidad) y la suma de todas.

### 6.2 `horas.py`

#### Constantes

- `ORDEN_MODALIDADES = ("T", "GA", "S", "GL", "GO")`: orden de las columnas.
- `POSICION_MODALIDAD`: índice de cada una en ese orden.

#### Clase `ResumenHoras`

- `modalidades: list[Modalidad]`: modalidades presentes, ya ordenadas.
- `horas: dict[Asignatura, dict[str, int]]`: horas por asignatura y clave de
  modalidad.

#### `calcular_horas(clases) -> ResumenHoras`

Cada `Clase` es una sesión (≈ 1 hora). El cálculo:

1. Para cada clase, `separar_tipo` obtiene `(modalidad, grupo)`, p. ej.
   `GL2` → `("GL", "GL2")`.
2. Cuenta sesiones en `sesiones[asignatura][modalidad][grupo]` y apunta qué
   grupos existen de cada modalidad en todo el horario.
3. Las horas de una modalidad son el **máximo** entre sus grupos, no la suma:
   los grupos paralelos (GL1, GL2…) son la misma docencia repetida para
   distintos alumnos.

   ```
   GL1 = 13 sesiones, GL2 = 13 sesiones  →  GL = 13 horas
   ```

4. Construye la lista de `Modalidad` en el orden de `ordenar_modalidades`,
   con la etiqueta que da `etiqueta_modalidad`.

#### `separar_tipo(tipo)`

Separa letras y número: `T` → `("T", "T")`, `GA1` → `("GA", "GA1")`. Si el
texto no encaja con `letras + número opcional`, devuelve el tipo tal cual en
ambas posiciones.

#### `ordenar_modalidades(modalidades)`

Ordena según `ORDEN_MODALIDADES`. Una modalidad desconocida va al final
(posición `len(ORDEN_MODALIDADES)`) y, entre desconocidas, alfabéticamente.

#### `etiqueta_modalidad(modalidad, grupos)`

Si en todo el horario solo aparece un grupo de esa modalidad, la columna
muestra el grupo concreto (`GA1`); si hay varios, la modalidad genérica
(`GA`).

---

## 7. Flujo B — Registros GAUR: `registros.py`

```
generar_registros (interfaz)
  └── generar_registros(rutas, cod_plan, destino)
        ├── leer_horario(ruta)  ×N            → list[Clase] (todas juntas)
        ├── construir_registros(clases, cod_plan)
        │     ├── nombres_asignaturas(clases)
        │     ├── juntar_sesiones(clases)
        │     │     ├── separar_grupo(tipo)
        │     │     ├── horas_franja(franja) → convertir_hora
        │     │     └── hora_gaur(hora)
        │     ├── comprimir_semanas(secuencia)   (por cada hueco)
        │     │     ├── mejor(desde)   (recursiva, memoizada)
        │     │     ├── valores_por_fase(...)
        │     │     └── filas_de_fase(...)
        │     └── formatear_hora(hora)
        └── guardar_registros(registros, destino)
```

### Idea general

GAUR no guarda cada sesión de cada semana. Guarda **registros** que dicen:
"esta actividad se imparte el día D de HH:MM a HH:MM, en las semanas
`SEMANA_INICIO`–`SEMANA_FIN`, una de cada `PERIODO` semanas, empezando por la
`FASE`-ésima". Una semana `s` está incluida si:

```
(s - semana_inicio) % periodo == fase - 1
```

El objetivo es describir todo el horario con **el menor número de filas**.

### Constantes

| Constante | Uso |
|---|---|
| `NOMBRE_HOJA` | Nombre de la hoja de salida: `registrosGAUR`. |
| `COLUMNAS` | Cabeceras de la hoja (hay una columna sin nombre, la `I`). |
| `ANCHOS_COLUMNA` | Ancho de las columnas D, E y H. |
| `NUMERO_DIA` | `Lunes` → 1 … `Viernes` → 5. |
| `DESCANSO_MAXIMO` | 30 min: hueco máximo entre dos sesiones para unirlas en un bloque. |
| `PERIODO_MAXIMO` | 5: GAUR no usa ciclos de más de cinco semanas. |

### Clase `Registro`

Una fila de la hoja de salida: `cod_plan`, `curso`, `cod_asignatura`,
`desc_asig`, `desc_asig_eusk`, `grupo`, `tipo_grupo`, `secuencia`,
`semana_inicio`, `semana_fin`, `fase`, `periodo`, `dia` (1–5), `hora_inicio`
y `hora_fin` (texto `HH:MM:00`).

### `generar_registros(rutas, cod_plan, destino)` — punto de entrada

1. Lee **todos** los horarios seleccionados y junta sus clases en una lista.
   Así un único fichero de salida contiene los grupos de castellano y euskera.
2. `construir_registros` → lista de `Registro`.
3. `guardar_registros` → escribe el `.xlsx`.
4. Devuelve cuántos registros se han escrito (lo muestra la interfaz).

### `construir_registros(clases, cod_plan) -> list[Registro]`

Es el núcleo del flujo:

1. **Nombres**: `nombres_asignaturas` obtiene, por código, el nombre en
   castellano y en euskera.
2. **Bloques**: `juntar_sesiones` une las sesiones consecutivas de la misma
   actividad en bloques con horas de GAUR.
3. **Huecos**: agrupa los bloques en
   `huecos[(grupo, día, inicio, fin)][semana] = {actividades}`.
   Un hueco es "el lunes de 08:30 a 10:30 del grupo 01"; para cada semana se
   anota qué actividades lo ocupan. Una actividad es
   `(código, tipo_grupo, secuencia)`, p. ej. `("26643", "GL", 1)`.
4. **Días con clase**: para cada grupo, qué `(semana, día)` tienen al menos
   una clase, y el curso de cada grupo.
5. **Secuencia por hueco**: para cada hueco se construye una tupla con una
   posición por semana (de la 1 a la última):
   - `frozenset(actividades)` si ese día el grupo tiene clase (vacío si el
     hueco está libre esa semana);
   - `None` si ese día el grupo no tiene **ninguna** clase → se considera
     festivo y valdrá como comodín (no rompe la repetición).
6. **Compresión**: `comprimir_semanas(secuencia)` devuelve las filas
   `(actividad, semana_inicio, semana_fin, fase, periodo)`; cada una se
   convierte en un `Registro` añadiendo nombres, curso, día numérico y horas
   formateadas.
7. **Orden**: por curso, asignatura, grupo, tipo, secuencia, semana de inicio,
   fase, día y hora.

### `es_grupo_euskera(grupo)`

En la UPV/EHU los grupos 31–60 son de euskera.

### `nombres_asignaturas(clases)`

Cada nombre se toma del horario de su idioma: si la clase es de un grupo de
euskera, su nombre va a la columna de euskera; si no, a la de castellano. Si
solo se ha cargado un idioma, la otra columna queda vacía (`None`). Devuelve
`{código: (nombre_castellano, nombre_euskera)}`.

### `horas_franja(franja)` y `convertir_hora(texto)`

`horas_franja` extrae las dos primeras horas de la franja
(`"8:40\n9:30"`) y `convertir_hora` las pasa a `timedelta` para poder
compararlas y restarlas.

### `hora_gaur(hora)`

GAUR trabaja con franjas que empiezan en punto o a y media. Redondea a la
media hora más cercana: `8:40` → `8:30`, `15:55` → `16:00`.

### `formatear_hora(hora)`

`timedelta` → `"HH:MM:00"`.

### `separar_grupo(tipo)`

Como `horas.separar_tipo`, pero devuelve el número como secuencia entera, que
es lo que pide GAUR:

```
T   -> ("T", 0)     teoría sin número: secuencia 0
GA1 -> ("GA", 1)
S   -> ("S", 1)     otra modalidad sin número: secuencia 1
```

### `juntar_sesiones(clases)`

1. Agrupa las franjas por `(grupo, semana, día, actividad)`.
2. Las ordena y recorre: si el hueco entre el fin de una sesión y el inicio de
   la siguiente es como mucho `DESCANSO_MAXIMO`, siguen en el mismo bloque; si
   no, cierra el bloque y abre otro.
3. Pasa a horas de GAUR: el inicio se redondea con `hora_gaur` y el fin es
   **una hora después del inicio redondeado de la última sesión** del bloque.

```
GL1: 15:00-15:50, 15:55-16:45, 17:00-17:50  →  15:00-18:00
```

Devuelve tuplas `(grupo, semana, día, inicio, fin, actividad)`.

### `comprimir_semanas(secuencia)`

Recibe la secuencia semanal de **un** hueco y busca la forma de cubrirla con
el mínimo de filas. Es una **programación dinámica** sobre la posición de
inicio.

Ejemplo (TAP, jueves 10:30):

```
semana:   1 2 3 4 5 | 6 7 8 9 10 11 12 13 14 15
actividad T T T T T | S T S T S  T  S  T  S  T

→ T   1-5   fase 1 periodo 1
  S   6-15  fase 1 periodo 2
  T   6-15  fase 2 periodo 2
```

Tiene tres funciones internas:

#### `mejor(desde)` (memoizada con `lru_cache`)

Devuelve la mejor forma de cubrir las semanas desde `desde` hasta el final,
como `(coste, filas)`.

1. Caso base: si `desde == n`, coste `(0, 0, 0)` y ninguna fila.
2. Si la semana `desde` está vacía o es festiva, se salta sin coste.
3. Si no, prueba todos los **tramos** `[desde, hasta]` que terminan en una
   semana con clase, y para cada uno todos los **periodos** de 1 a
   `min(PERIODO_MAXIMO, longitud)`:
   - Un periodo `p > 1` exige que el tramo tenga al menos `2p - 1` semanas
     (que el patrón se vea casi dos veces); si no, no tendría sentido
     hablar de repetición.
   - `valores_por_fase` comprueba que el tramo realmente se repite con ese
     periodo. Si no, se descarta.
   - `filas_de_fase` genera las filas de cada fase.
   - Se suma el coste del resto: `mejor(hasta + 1)`.
4. Se elige la opción de menor coste. El coste se compara por orden:
   1. número de filas,
   2. suma de periodos (prefiere periodos pequeños),
   3. número de tramos,
   4. en empate, el tramo más largo (`-longitud`).

#### `valores_por_fase(desde, hasta, periodo)`

Reparte las semanas del tramo entre las `periodo` fases
(`(semana - desde) % periodo`). Ignora las festivas (`None`). Si dentro de una
misma fase hay dos semanas con actividades distintas, el tramo no es periódico
con ese periodo y devuelve `None`. Si lo es, devuelve para cada fase
`(actividades, semanas_reales)`, donde `semanas_reales` son las semanas no
festivas de esa fase.

#### `filas_de_fase(actividades, semanas_reales, desde, hasta, fase, periodo)`

Genera una fila por actividad de la fase:

- Normalmente: `(actividad, desde+1, hasta+1, fase+1, periodo)` (las semanas
  se pasan de índice 0 a número de semana 1).
- Si el periodo es mayor que 1 pero la fase solo cae en **una** semana real,
  la escribe como esa semana suelta (periodo 1), que es más legible:

  ```
  semanas 7-9:  GO GL GO  →  GO  7-9  fase 1 periodo 2
                             GL  8-8  fase 1 periodo 1
  ```

Al final, `comprimir_semanas` llama a `mejor(0)` y devuelve sus filas.

### `guardar_registros(registros, destino)`

1. Crea un libro nuevo con una hoja `registrosGAUR`.
2. Escribe las cabeceras de `COLUMNAS` en negrita.
3. Una fila por registro. La columna I (sin nombre) lleva la fórmula
   `=G{fila}&H{fila}`, que concatena tipo de grupo y secuencia (p. ej.
   `GL1`).
4. Ajusta el ancho de las columnas D, E y H y guarda en `destino`.

---

## 8. Resumen de llamadas

```
python -m src.main
└── main()
    └── iniciar_interfaz()
        ├── VentanaPrincipal.__init__()
        │   └── _crear_tabla()
        └── root.mainloop()   ── espera eventos ──┐
                                                  │
   ┌──────────── botón "Seleccionar horario" ─────┤
   │ cargar_horario()                             │
   │ ├── analizar_horario(ruta)                   │
   │ │   ├── leer_horario(ruta)                   │
   │ │   │   └── leer_hoja(hoja)  (por hoja)      │
   │ │   │       ├── leer_asignaturas()           │
   │ │   │       │   └── normalizar_codigo()      │
   │ │   │       ├── obtener_semana()             │
   │ │   │       ├── obtener_curso_grupo()        │
   │ │   │       ├── obtener_columnas_dias()      │
   │ │   │       ├── es_inicio_tabla_asignaturas()│
   │ │   │       ├── es_franja_horaria()          │
   │ │   │       └── leer_franja()                │
   │ │   │           └── interpretar_celda()      │
   │ │   ├── calcular_horas(clases)               │
   │ │   │   ├── separar_tipo()                   │
   │ │   │   ├── ordenar_modalidades()            │
   │ │   │   └── etiqueta_modalidad()             │
   │ │   └── construir_tabla(resumen)             │
   │ └── mostrar_resultado(tabla)                 │
   │                                              │
   ├──────────── Ctrl+C sobre la tabla ───────────┤
   │ copiar_tabla()                               │
   │                                              │
   └──────────── botón "Generar registros GAUR" ──┘
     VentanaPrincipal.generar_registros()
     └── registros.generar_registros(rutas, cod_plan, destino)
         ├── leer_horario(ruta)  (por cada fichero; igual que arriba)
         ├── construir_registros(clases, cod_plan)
         │   ├── nombres_asignaturas()
         │   │   └── es_grupo_euskera()
         │   ├── juntar_sesiones()
         │   │   ├── separar_grupo()
         │   │   ├── horas_franja() → convertir_hora()
         │   │   └── hora_gaur()
         │   ├── comprimir_semanas()  (por cada hueco)
         │   │   └── mejor() → valores_por_fase(), filas_de_fase()
         │   └── formatear_hora()
         └── guardar_registros()
```
