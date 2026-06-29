import customtkinter as ctk
import os
import pdfplumber
import pandas as pd
import re
from tkinter import filedialog, messagebox

class GUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x300")
        self.title("Writter of bills")
        ctk.set_appearance_mode("dark")

        # Variables para almacenar las rutas
        self.ruta_origen = ""
        self.ruta_destino = ""

        self.buttonSelectFolder = ctk.CTkButton(self, text="Seleccionar Carpeta", command=self.seleccionar_carpeta)
        self.buttonSelectFolder.pack(pady=20)

        self.buttonDestinyExcel = ctk.CTkButton(self, text="Seleccionar Destino Excel", command=self.seleccionar_destino)
        self.buttonDestinyExcel.pack(pady=20)

        self.btn_procesar = ctk.CTkButton(self, text="Procesar Facturas", command=self.procesar_todo, fg_color="green")
        self.btn_procesar.pack(pady=20)

    def seleccionar_carpeta(self):
        self.ruta_origen = filedialog.askdirectory()
        print(f"Carpeta seleccionada: {self.ruta_origen}")

    def seleccionar_destino(self):
        self.ruta_destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        print(f"Destino: {self.ruta_destino}")

    def extraer_datos_factura(self, ruta_pdf):
        datos = {}
        with pdfplumber.open(ruta_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            if "FACTURA:" in texto:
                datos['Numero'] = texto.split("FACTURA: #")[1].split()[0]
            match = re.search(r"Total a pagar:\s*(\d+\.\d{2})", texto)
            if match:
                datos['Total'] = match.group(1)
            datos['Archivo'] = os.path.basename(ruta_pdf)
        return datos

    def procesar_todo(self):
        if not self.ruta_origen or not self.ruta_destino:
            messagebox.showwarning("Error", "Selecciona carpeta y destino primero")
            return

        lista_facturas = []
        for nombre_archivo in os.listdir(self.ruta_origen):
            if nombre_archivo.endswith(".pdf"):
                ruta_completa = os.path.join(self.ruta_origen, nombre_archivo)
                lista_facturas.append(self.extraer_datos_factura(ruta_completa))

        if lista_facturas:
            pd.DataFrame(lista_facturas).to_excel(self.ruta_destino, index=False)
            messagebox.showinfo("Éxito", "Excel generado correctamente")
        else:
            messagebox.showwarning("Aviso", "No se encontraron PDFs")

app = GUI()
app.mainloop() # Nota: es mainloop() en minúsculas