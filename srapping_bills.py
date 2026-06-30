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

        self.buttonSelectFolder = ctk.CTkButton(self, text="Seleccionar Factura", command=self.seleccionar_archivo)
        self.buttonSelectFolder.pack(pady=20)

        self.buttonDestinyExcel = ctk.CTkButton(self, text="Seleccionar Destino Excel", command=self.seleccionar_destino)
        self.buttonDestinyExcel.pack(pady=20)

        self.btn_procesar = ctk.CTkButton(self, text="Procesar Facturas", command=self.procesar_todo, fg_color="green")
        self.btn_procesar.pack(pady=20)

    def seleccionar_archivo(self):
        archivo_pdf = filedialog.askopenfilename(
        title="Selecciona un archivo PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
        if archivo_pdf:
            self.ruta_origen = archivo_pdf
            print(f"PDF seleccionado: {self.ruta_origen}")

    def seleccionar_destino(self):
        self.ruta_destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        print(f"Destino: {self.ruta_destino}")
    
    def extraer_todas_las_tablas(self, ruta_pdf):
        with pdfplumber.open(ruta_pdf) as pdf:
            pagina = pdf.pages[0]
            tablas = pagina.extract_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
            
            datos_totales = []
            
            for tabla in tablas:
                fila_cabecera_idx = -1
                
                # 1. Buscamos en qué fila aparece la palabra "Importe"
                for i, fila in enumerate(tabla):
                    # Convertimos la fila a texto (cuidando valores None)
                    fila_texto = [str(celda).lower() if celda else "" for celda in fila]
                    if any("importe" in celda for celda in fila_texto):
                        fila_cabecera_idx = i
                        encabezados = fila # Esta es nuestra fila de encabezados real
                        break
                
                # 2. Si encontramos la fila, procesamos desde la siguiente
                if fila_cabecera_idx != -1:
                    idx_map = {i: nombre for i, nombre in enumerate(encabezados) if nombre and "importe" in str(nombre).lower()}
                    
                    for fila in tabla[fila_cabecera_idx + 1:]:
                        # Extraemos los datos basándonos en el índice encontrado
                        fila_dict = {f"Importe_{i}": fila[i] for i in idx_map.keys() if fila[i]}
                        if fila_dict:
                            datos_totales.append(fila_dict)
                            
            return datos_totales
        
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

        if not os.path.exists(self.ruta_origen):
            messagebox.showerror("Error", "La ruta seleccionada no existe")
            return
        
        if self.ruta_origen.endswith(".pdf"):
            # Procesar solo ese archivo directamente
            datos = self.extraer_todas_las_tablas(self.ruta_origen)
            self.actualizar_excel(datos, self.ruta_destino)
            messagebox.showinfo("Éxito", "Archivo procesado")
        
        else : 
            for nombre_archivo in os.listdir(self.ruta_origen):
                if nombre_archivo.endswith(".pdf"):
                    ruta_completa = os.path.join(self.ruta_origen, nombre_archivo)
                    # Extraemos la lista de importes
                    datos = self.extraer_todas_las_tablas(ruta_completa)
                    # Actualizamos el excel inmediatamente
                    self.actualizar_excel(datos, self.ruta_destino)
        
        messagebox.showinfo("Éxito", "Proceso terminado")

app = GUI()
app.mainloop() # Nota: es mainloop() en minúsculas