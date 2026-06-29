from CTkColorPicker import *
import customtkinter as ctk
import os
import pdfplumber
import pandas as pd
import re


def get_path ():
    directorio_base = os.path.dirname(os.path.abspath(__file__))
    carpeta_pdf = os.path.join(directorio_base, "facturas")
    archivo_salida = "resumen_facturas.xlsx"

def extraer_datos_factura(ruta_pdf):
    
    """Aquí defines cómo buscar los datos dentro del texto del PDF"""
    datos = {}
    with pdfplumber.open(ruta_pdf) as pdf:
        pagina = pdf.pages[0]
        texto = pagina.extract_text()
        
        # Lógica de extracción simple (ajusta según el formato real de tu PDF)
        # Esto es un ejemplo basado en lo que vimos antes
        lines = texto.split('\n')
        for line in lines:
            if "FACTURA:" in line:
                datos['Numero'] = line.split("#")[1].strip()
            match = re.search(r"Total a pagar:\s*(\d+\.\d{2})", texto)
            if match:
               datos['Total'] = match.group(1)
        datos['Archivo'] = os.path.basename(ruta_pdf)
    return datos

# 2. Bucle principal

def create_excel ():
    lista_facturas = []
    if os.path.exists(carpeta_pdf):
        for nombre_archivo in os.listdir(carpeta_pdf):
            if nombre_archivo.endswith(".pdf"):
                ruta_completa = os.path.join(carpeta_pdf, nombre_archivo)
                print(f"Extrayendo datos de: {nombre_archivo}")
                
                # Extraer y añadir a la lista
                info = extraer_datos_factura(ruta_completa)
                lista_facturas.append(info)

    # 3. Guardar en Excel
        if lista_facturas:
            df = pd.DataFrame(lista_facturas)
            df.to_excel(archivo_salida, index=False)
            print(f"¡Éxito! Datos guardados en {archivo_salida}")
        else:
            print("No se encontraron datos para guardar.")
    else:
        print(f"La carpeta '{carpeta_pdf}' no existe.")

class GUI (ctk.CTk):
    def __init__(self):
        super().__init__()
       
        self.geometry("400x800")
        self.title("Writter of bills")
        ctk.set_appearance_mode("dark")

        self.buttonSelectFolder = ctk.CTkButton(self, text="Select folder of bills", text_color="#D4D4D4", command= self.exportDocument, fg_color="#585858", hover_color="#7A7A7A" )
        self.buttonSelectFolder.pack(padx=20, pady=5)

        self.buttonDestinyExcel = ctk.CTkButton(self, text="Select excel of destination", text_color="#D4D4D4", command= self.exportDocument, fg_color="#585858", hover_color="#7A7A7A" )
        self.buttonDestinyExcel.pack(padx=20, pady=5)      

app = GUI()
app.mainLoop()