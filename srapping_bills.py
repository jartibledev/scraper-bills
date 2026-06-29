import os
import pdfplumber
import pandas as pd

# 1. Configuración de rutas
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
            if "Total a pagar:" in line:
                datos['Total'] = line.split("Total a pagar:")[1].strip()
        
        datos['Archivo'] = os.path.basename(ruta_pdf)
    return datos

# 2. Bucle principal
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