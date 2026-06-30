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
            texto = pagina.extract_text()
            print(f"--- Diagnóstico para {os.path.basename(ruta_pdf)} ---")
            print("Texto detectado en la página:")
            print(texto[:200])
            
            # 'extract_tables' detecta múltiples tablas separadas
            tablas = pagina.extract_tables(table_settings={
                "vertical_strategy": "text", 
                "horizontal_strategy": "text"
            })
            print(f"Número de tablas detectadas: {len(tablas)}")
            
            datos_totales = []
            for i, tabla in enumerate(tablas):
                print(f"Encabezados de la tabla {i}: {tabla[0]}")
            
            for tabla in tablas:
               
                # tabla[0] son los encabezados
                encabezados = tabla[0]
                
                columnas_necesarias = ["Cantidad", "Importe"]

                if encabezados and all(col in encabezados for col in columnas_necesarias):
                    idx_map = {n: i for i, n in enumerate(encabezados) if n in columnas_necesarias}
                    
                    for fila in tabla[1:]:
                        # Usamos 'fila[indice] or ""' para evitar errores si la celda está vacía
                        fila_dict = {nombre: (fila[indice] if fila[indice] is not None else "") for nombre, indice in idx_map.items()}
                        
                        # Solo añadimos si al menos uno de los valores tiene contenido real
                        if any(val != "" for val in fila_dict.values()):
                            datos_totales.append(fila_dict)

            print( f"Datos: {datos_totales}" )                
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