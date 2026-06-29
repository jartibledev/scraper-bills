import pdfplumber
import pandas as pd
import os

def verificar_lectura(ruta_pdf):
    with pdfplumber.open(ruta_pdf) as pdf:
        pagina = pdf.pages[0]
        texto = pagina.extract_text()
        print("--- CONTENIDO DEL PDF ---")
        print(texto)
        print("-------------------------")



def extraer_datos_factura(ruta_pdf):
    datos = {}
    with pdfplumber.open(ruta_pdf) as pdf:
        pagina = pdf.pages[0]  # Asumimos que la info está en la primera página
        texto = pagina.extract_text()
        
        # AQUÍ VA LA LÓGICA DE EXTRACCIÓN
        # Ejemplo: Buscar palabras clave y extraer el valor siguiente
        # (Esto requiere ajustar según el formato específico de tus facturas)
        if "Total:" in texto:
            datos['Total'] = texto.split("Total:")[1].split()[0]
        datos['Archivo'] = os.path.basename(ruta_pdf)
        
    return datos

# 1. Definimos la carpeta y el nombre del archivo
nombre_carpeta = "facturas"
nombre_archivo = "factura_001.pdf"

# 2. Construimos la ruta de forma segura (sin importar el sistema operativo)
directorio_base = os.path.dirname(os.path.abspath(__file__))
ruta_completa = os.path.join(directorio_base, nombre_carpeta, nombre_archivo)

# 3. Verificación
if os.path.exists(ruta_completa):
    print(f"¡Éxito! Archivo encontrado en: {ruta_completa}")
    
    # Aquí iría tu lógica de lectura
    with pdfplumber.open(ruta_completa) as pdf:
        # ... procesar el PDF ...
        pass
else:
    print(f"Error: No se encontró el archivo en: {ruta_completa}")
    print("Por favor, asegúrate de que la carpeta 'facturas' esté al mismo nivel que este script.")

lista_facturas = []

carpeta_pdf = os.path.join(os.path.dirname(os.path.abspath(__file__)), "facturas")

# 2. Listar y procesar archivos en esa carpeta
for nombre_archivo in os.listdir(carpeta_pdf):
    # Verificamos que sea un archivo PDF
    if nombre_archivo.endswith(".pdf"):
        ruta_completa = os.path.join(carpeta_pdf, nombre_archivo)
        print(f"Procesando: {nombre_archivo}")
        
        # Aquí llamas a tu función de extracción
        datos = extraer_datos_factura(ruta_completa)
        lista_facturas.append(datos)

# Convertir a Excel
df = pd.DataFrame(lista_facturas)
df.to_excel("resumen_facturas.xlsx", index=False)
print("¡Archivo Excel generado con éxito!")