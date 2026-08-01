# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import os


class AplicacionRIPS:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "Corrección Automática de Facturas RIPS"
        )

        self.root.geometry("850x550")
        self.root.resizable(False, False)

        self.archivo_csv = ""

        self.crear_interfaz()

    # -----------------------------------------------------

    def crear_interfaz(self):

        contenedor = tk.Frame(
            self.root,
            padx=20,
            pady=20
        )

        contenedor.pack(
            fill="both",
            expand=True
        )

        titulo = tk.Label(
            contenedor,
            text="CORRECCIÓN AUTOMÁTICA DE FACTURAS RIPS",
            font=("Segoe UI", 18, "bold")
        )

        titulo.pack(
            pady=(20, 10)
        )

        subtitulo = tk.Label(
            contenedor,
            text=(
                "Seleccione la factura CSV y el periodo "
                "al cual desea mover las fechas."
            ),
            font=("Segoe UI", 10)
        )

        subtitulo.pack(
            pady=(0, 25)
        )

        # -------------------------------------------------
        # CONFIGURACIÓN
        # -------------------------------------------------

        frame_config = tk.LabelFrame(
            contenedor,
            text="Configuración",
            padx=20,
            pady=15
        )

        frame_config.pack(
            fill="x",
            padx=40,
            pady=10
        )

        # Año

        lbl_anio = tk.Label(
            frame_config,
            text="Año:"
        )

        lbl_anio.grid(
            row=0,
            column=0,
            padx=10,
            pady=10
        )

        self.combo_anio = ttk.Combobox(
            frame_config,
            width=15,
            state="readonly"
        )

        self.combo_anio["values"] = (
            2024,
            2025,
            2026,
            2027,
            2028,
            2029,
            2030
        )

        self.combo_anio.set("2026")

        self.combo_anio.grid(
            row=0,
            column=1,
            padx=10
        )

        # Mes

        lbl_mes = tk.Label(
            frame_config,
            text="Mes:"
        )

        lbl_mes.grid(
            row=0,
            column=2,
            padx=10
        )

        self.combo_mes = ttk.Combobox(
            frame_config,
            width=20,
            state="readonly"
        )

        self.combo_mes["values"] = (
            "1 - Enero",
            "2 - Febrero",
            "3 - Marzo",
            "4 - Abril",
            "5 - Mayo",
            "6 - Junio",
            "7 - Julio",
            "8 - Agosto",
            "9 - Septiembre",
            "10 - Octubre",
            "11 - Noviembre",
            "12 - Diciembre"
        )

        self.combo_mes.set(
            "4 - Abril"
        )

        self.combo_mes.grid(
            row=0,
            column=3,
            padx=10
        )

        # -------------------------------------------------
        # ARCHIVO
        # -------------------------------------------------

        frame_archivo = tk.LabelFrame(
            contenedor,
            text="Factura"
        )

        frame_archivo.pack(
            fill="x",
            padx=40,
            pady=20
        )

        self.lbl_archivo = tk.Label(
            frame_archivo,
            text="No se ha seleccionado ningún archivo",
            anchor="w"
        )

        self.lbl_archivo.pack(
            fill="x",
            padx=10,
            pady=10
        )

        btn_archivo = tk.Button(
            frame_archivo,
            text="Seleccionar Factura CSV",
            width=30,
            height=2,
            command=self.seleccionar_archivo
        )

        btn_archivo.pack(
            pady=10
        )

        # -------------------------------------------------
        # BOTÓN EJECUTAR
        # -------------------------------------------------

        btn_ejecutar = tk.Button(
            contenedor,
            text="EJECUTAR CORRECCIÓN",
            font=("Segoe UI", 11, "bold"),
            width=30,
            height=2,
            command=self.ejecutar
        )

        btn_ejecutar.pack(
            pady=20
        )

        # -------------------------------------------------

        self.estado = tk.Label(
            contenedor,
            text="Estado: Esperando archivo",
            font=("Segoe UI", 10)
        )

        self.estado.pack(
            pady=15
        )

    # -----------------------------------------------------

    def seleccionar_archivo(self):

        archivo = filedialog.askopenfilename(
            filetypes=[
                (
                    "Archivos CSV",
                    "*.csv *.CSV"
                )
            ]
        )

        if archivo:

            self.archivo_csv = archivo

            self.lbl_archivo.config(
                text=os.path.basename(archivo)
            )

            self.estado.config(
                text="Archivo cargado correctamente"
            )

    # -----------------------------------------------------

    def ejecutar(self):

        if not self.archivo_csv:

            messagebox.showwarning(
                "Advertencia",
                "Seleccione una factura CSV."
            )

            return

        try:

            anio = self.combo_anio.get()

            mes = (
                self.combo_mes.get()
                .split("-")[0]
                .strip()
            )

            self.estado.config(
                text="Procesando..."
            )

            self.root.update()

            script = r"C:\proyecto\PRUEBA V3.py"

            subprocess.run(
                [
                    "python",
                    script,
                    self.archivo_csv,
                    anio,
                    mes
                ],
                check=True
            )

            self.estado.config(
                text="Proceso finalizado correctamente"
            )

            messagebox.showinfo(
                "Proceso terminado",
                (
                    "Factura corregida exitosamente.\n\n"
                    "Revise:\n"
                    "C:\\proyecto\\factura_corregida"
                )
            )

        except Exception as e:

            self.estado.config(
                text="Error durante la ejecución"
            )

            messagebox.showerror(
                "Error",
                str(e)
            )


# ---------------------------------------------------------
# INICIO DE LA APLICACIÓN
# ---------------------------------------------------------

if __name__ == "__main__":

    root = tk.Tk()

    app = AplicacionRIPS(root)

    root.mainloop()