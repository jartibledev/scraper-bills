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

directorio_actual = os.path.dirname(os.path.abspath(__file__))
ruta_pdf = os.path.join(directorio_actual, "factura_001.pdf")
print(f"¿Existe el archivo en {ruta_pdf}? {os.path.exists(ruta_pdf)}")
verificar_lectura(ruta_pdf)
# Lista para almacenar los resultados de varias facturas
lista_facturas = []

# Supongamos que tienes una carpeta con tus PDFs
carpeta_pdf = ruta_pdf
for archivo in os.listdir(carpeta_pdf):
    if archivo.endswith(".pdf"):
        ruta = os.path.join(carpeta_pdf, archivo)
        datos = extraer_datos_factura(ruta)
        lista_facturas.append(datos)

# Convertir a Excel
df = pd.DataFrame(lista_facturas)
df.to_excel("resumen_facturas.xlsx", index=False)
print("¡Archivo Excel generado con éxito!")