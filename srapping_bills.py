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
        # 1. Definimos los alias dentro o fuera de la función
        alias_columnas = {
            "Importe": ["importe", "total", "amount", "price", "precio", "subtotal"],
            "Cantidad": ["cantidad", "units", "uds", "qty", "cantidad total"],
            "Concepto": ["concepto", "concept", "descripcion", "description"]
        }
        
        with pdfplumber.open(ruta_pdf) as pdf:
            pagina = pdf.pages[0]
            tablas = pagina.extract_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
            
            datos_totales = []
            
            for tabla in tablas:
                idx_map = {}
                fila_cabecera_idx = -1
                
                # Buscamos la fila que contiene los encabezados
                for i, fila in enumerate(tabla):
                    fila_texto = [str(c).lower() if c else "" for c in fila]
                    # Si al menos una celda contiene alguno de nuestros alias
                    if any(any(a in celda for lista in alias_columnas.values() for a in lista) for celda in fila_texto):
                        fila_cabecera_idx = i
                        # Mapeamos qué índice corresponde a qué estándar
                        for col_idx, celda in enumerate(fila_texto):
                            for estandar, alias_list in alias_columnas.items():
                                if any(a in celda for a in alias_list):
                                    idx_map[col_idx] = estandar
                        break
                
                # Si encontramos encabezados, extraemos los datos
                if fila_cabecera_idx != -1:
                    for fila in tabla[fila_cabecera_idx + 1:]:
                        # Usamos los alias para construir el diccionario
                        fila_dict = {idx_map[i]: fila[i] for i in idx_map.keys() if i < len(fila) and fila[i]}
                        if fila_dict:
                            datos_totales.append(fila_dict)
                            
            return datos_totales
        
    def actualizar_excel(self, nueva_lista_datos, ruta_excel):
        if not nueva_lista_datos: return

        try:
            if os.path.exists(ruta_excel):
                df_existente = pd.read_excel(ruta_excel)
            else:
                df_existente = pd.DataFrame(columns=["Concepto", "Cantidad", "Importe"])

            df_nuevos = pd.DataFrame(nueva_lista_datos)
            df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
            df_final.to_excel(ruta_excel, index=False)
            
        except PermissionError:
            messagebox.showerror("Error de Permisos", 
                                 "No puedo escribir en el archivo porque está abierto en Excel.\n"
                                 "Por favor, cierra el archivo y vuelve a intentarlo.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")
        
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