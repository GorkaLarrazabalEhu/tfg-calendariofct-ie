import tkinter as tk

from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
from tkinter import ttk

from src.analisis import Tabla, analizar_horario
from src.registros import generar_registros


TITULO = "Análisis de horario docente"

# Cómo se muestran las celdas sin valor
SIN_VALOR = "0"

ANCHO_MINIMO = 70
ANCHO_MAXIMO = 250


def iniciar_interfaz():
    root = tk.Tk()
    VentanaPrincipal(root)
    root.mainloop()


class VentanaPrincipal:
    """
    Ventana con un botón para elegir el horario
    y una tabla con el resultado del análisis.
    """

    def __init__(self, root):
        root.title(TITULO)
        root.geometry("1000x600")

        ttk.Label(
            root,
            text=TITULO,
            font=("TkDefaultFont", 16, "bold"),
        ).pack(pady=(20, 10))

        self.etiqueta_archivo = ttk.Label(root, text="Ningún archivo seleccionado")
        self.etiqueta_archivo.pack(pady=5)

        self.tabla = self._crear_tabla(root)
        self.tabla.bind("<Control-c>", lambda evento: self.copiar_tabla())

        botones = ttk.Frame(root)
        botones.pack(pady=(0, 20))

        ttk.Button(
            botones,
            text="Seleccionar horario",
            command=self.cargar_horario,
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            botones,
            text="Generar registros GAUR",
            command=self.generar_registros,
        ).pack(side=tk.LEFT, padx=5)

        self.cod_plan = ""

    def _crear_tabla(self, root):
        marco = ttk.Frame(root)
        marco.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tabla = ttk.Treeview(marco, show="headings")

        scroll_vertical = ttk.Scrollbar(marco, orient=tk.VERTICAL, command=tabla.yview)
        scroll_horizontal = ttk.Scrollbar(marco, orient=tk.HORIZONTAL, command=tabla.xview)

        tabla.configure(
            yscrollcommand=scroll_vertical.set,
            xscrollcommand=scroll_horizontal.set,
        )

        tabla.grid(row=0, column=0, sticky="nsew")
        scroll_vertical.grid(row=0, column=1, sticky="ns")
        scroll_horizontal.grid(row=1, column=0, sticky="ew")

        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)

        return tabla

    def cargar_horario(self):
        """
        Selecciona el archivo y solicita su análisis.
        """

        ruta = filedialog.askopenfilename(
            title="Seleccionar horario",
            filetypes=[("Archivos Excel", "*.xlsx")],
        )

        if not ruta:
            return

        try:
            resultado = analizar_horario(ruta)

        except Exception as error:
            messagebox.showerror(
                "Error",
                f"No se ha podido analizar el archivo.\n\n{error}",
            )
            return

        self.etiqueta_archivo.config(text=ruta)
        self.mostrar_resultado(resultado)

    def generar_registros(self):
        """
        Genera la lista de registros (hoja registrosGAUR)
        a partir de uno o varios horarios, p. ej. castellano y euskera.
        """

        rutas = filedialog.askopenfilenames(
            title="Seleccionar horarios (castellano y/o euskera)",
            filetypes=[("Archivos Excel", "*.xlsx")],
        )

        if not rutas:
            return

        cod_plan = simpledialog.askstring(
            "Código de plan",
            "Código del plan (COD_PLAN), p. ej. GELECT30:",
            initialvalue=self.cod_plan,
        )

        if not cod_plan:
            return

        self.cod_plan = cod_plan.strip()

        destino = filedialog.asksaveasfilename(
            title="Guardar registros",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            initialfile="registrosGAUR.xlsx",
        )

        if not destino:
            return

        try:
            total = generar_registros(rutas, self.cod_plan, destino)

        except Exception as error:
            messagebox.showerror(
                "Error",
                f"No se han podido generar los registros.\n\n{error}",
            )
            return

        messagebox.showinfo(
            "Registros generados",
            f"Se han escrito {total} registros en:\n{destino}",
        )

    def mostrar_resultado(self, resultado: Tabla):
        """
        Muestra cualquier resultado tabular.

        No sabe nada de asignaturas ni de cómo se calculan
        las horas: solo conoce columnas y filas.
        """

        self.tabla.delete(*self.tabla.get_children())

        filas = [
            [SIN_VALOR if valor is None else str(valor) for valor in fila]
            for fila in resultado.filas
        ]

        self.tabla["columns"] = resultado.columnas

        for indice, columna in enumerate(resultado.columnas):

            longitud_maxima = max(
                [len(columna), *(len(fila[indice]) for fila in filas)]
            )

            ancho = min(ANCHO_MAXIMO, max(ANCHO_MINIMO, longitud_maxima * 9 + 20))

            self.tabla.heading(columna, text=columna)
            self.tabla.column(columna, width=ancho, anchor=tk.CENTER)

        for fila in filas:
            self.tabla.insert("", tk.END, values=fila)

    def copiar_tabla(self):
        """
        Copia la tabla al portapapeles separada por tabuladores,
        lista para pegar en Excel.

        Si hay filas seleccionadas, copia solo esas.
        Si no hay ninguna seleccionada, copia toda la tabla.
        """

        columnas = self.tabla["columns"]

        lineas = [
            "\t".join(self.tabla.heading(columna)["text"] for columna in columnas)
        ]

        for fila in self.tabla.selection() or self.tabla.get_children():

            valores = self.tabla.item(fila, "values")

            lineas.append(
                "\t".join(
                    "" if valor == SIN_VALOR else str(valor)
                    for valor in valores
                )
            )

        self.tabla.clipboard_clear()
        self.tabla.clipboard_append("\n".join(lineas))
