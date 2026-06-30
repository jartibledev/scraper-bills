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
    
    def extraer_columna_importe(self, ruta_pdf):
        with pdfplumber.open(ruta_pdf) as pdf:
            pagina = pdf.pages[0]
            tabla = pagina.extract_table(table_settings={
                "vertical_strategy": "text", 
                "horizontal_strategy": "text"
            }) # Obtienes una lista de filas
            
            if not tabla: return None
            
            # 1. Buscamos el índice de la columna "Importe" en la primera fila
            encabezados = tabla[0]
            try:
                indice_importe = encabezados.index("Importe")
            except ValueError:
                print("No se encontró la columna 'Importe'")
                return None
            
            # 2. Extraemos el valor de esa columna en todas las filas siguientes
            importes = [fila[indice_importe] for fila in tabla[1:] if fila[indice_importe]]
            return importes
        
    def actualizar_excel(self, nueva_lista_datos, ruta_excel):
        # Si la lista está vacía, no hacemos nada
        if not nueva_lista_datos: return

        if os.path.exists(ruta_excel):
            df_existente = pd.read_excel(ruta_excel)
        else:
            df_existente = pd.DataFrame(columns=["Importe"])

        df_nuevos = pd.DataFrame(nueva_lista_datos)
        df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
        df_final.to_excel(ruta_excel, index=False)
        
    def procesar_todo(self):
        if not self.ruta_origen or not self.ruta_destino:
            messagebox.showwarning("Error", "Selecciona carpeta y destino primero")
            return

        
        for nombre_archivo in os.listdir(self.ruta_origen):
            if nombre_archivo.endswith(".pdf"):
                ruta_completa = os.path.join(self.ruta_origen, nombre_archivo)
                # Extraemos la lista de importes
                datos = self.extraer_columna_importe(ruta_completa)
                # Actualizamos el excel inmediatamente
                self.actualizar_excel(datos, self.ruta_destino)
        
        messagebox.showinfo("Éxito", "Proceso terminado")

app = GUI()
app.mainloop() # Nota: es mainloop() en minúsculas