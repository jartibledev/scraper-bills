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
    
    def limpiar_concepto(self, text):

        match_date_and_code = re.search(r'Ref Reserva:\s(?P<codigo>\d{8,})\s\((?P<fecha_inicio>\d{1,2}/\d{1,2}/\d{2,4})\s-\s(?P<fecha_fin>\d{1,2}/\d{1,2}/\d{2,4})\)')

    def extraer_texto_completo(self, pdf_path):
        
        texto_total = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Extrae el texto plano de toda la página
                texto_total += page.extract_text() + "\n"
        
        return texto_total

    def extraer_todas_las_tablas(self, ruta_pdf):
        alias_columnas = {
            "Importe": ["importe", "total", "amount", "price", "precio", "subtotal"],
            "Concepto": ["concepto", "concept", "descripcion", "description"]
        }
        
        # 1. Extraer todas las filas (igual que antes)
        todas_las_filas = []
        with pdfplumber.open(ruta_pdf) as pdf:
            for pagina in pdf.pages:
                tablas = pagina.extract_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
                for tabla in tablas:
                    idx_map = {}
                    fila_cabecera_idx = -1
                    for i, fila in enumerate(tabla):
                        fila_texto = [str(c).lower() if c else "" for c in fila]
                        if any(any(a in celda for lista in alias_columnas.values() for a in lista) for celda in fila_texto):
                            fila_cabecera_idx = i
                            for col_idx, celda in enumerate(fila_texto):
                                for estandar, alias_list in alias_columnas.items():
                                    if any(a in celda for a in alias_list):
                                        idx_map[col_idx] = estandar
                            break
                    for fila in tabla[fila_cabecera_idx + 1:]:
                        fila_dict = {idx_map[i]: fila[i] for i in idx_map.keys() if i < len(fila) and fila[i]}
                        if fila_dict:
                            todas_las_filas.append(fila_dict)

           # 2. Pre-calcular las reservas globales (para no buscarlas en cada fila)
        texto_total_pdf = self.extraer_texto_completo(ruta_pdf)
        texto_limpio = " ".join(texto_total_pdf.split())
        print (texto_limpio)
        patron = r'Ref\s*Reserva.*?(\d{8,})\s*\((.*?)\s*-\s*(.*?)\)'
        reservas_encontradas = list(re.finditer(patron, texto_total_pdf, re.IGNORECASE | re.DOTALL))
        print (reservas_encontradas)
        array_plano = [m.group(0) for m in reservas_encontradas]
        print (array_plano)

        
        
         # 2. Pre-calcular las reservas globales (para no buscarlas en cada fila)
        texto_total_pdf = self.extraer_texto_completo(ruta_pdf)
        texto_limpio = " ".join(texto_total_pdf.split())
        print (texto_limpio)
        patron = r'Ref\s*Reserva.*?(\d{8,})\s*\((.*?)\s*-\s*(.*?)\)'
        reservas_encontradas = list(re.finditer(patron, texto_total_pdf, re.IGNORECASE | re.DOTALL))
        print (reservas_encontradas)
        array_plano = [m.group(0) for m in reservas_encontradas]
        print (array_plano)

        
        
        datos_finales = []
       # 2. Pre-calcular las reservas globales (para no buscarlas en cada fila)
        texto_total_pdf = self.extraer_texto_completo(ruta_pdf)
        texto_limpio = " ".join(texto_total_pdf.split())
        print (texto_limpio)
        patron = r'Ref\s*Reserva.*?(\d{8,})\s*\((.*?)\s*-\s*(.*?)\)'
        reservas_encontradas = list(re.finditer(patron, texto_total_pdf, re.IGNORECASE | re.DOTALL))
        print (reservas_encontradas)
        array_plano = [m.group(0) for m in reservas_encontradas]
        print (array_plano)

        
        
        datos_finales = []
        
        # 3. Procesar cada fila individualmente
        for fila in todas_las_filas:
            concepto_raw = str(fila.get("Concepto", ""))
            importe_raw = fila.get("Importe", 0)
            texto_low = concepto_raw.lower()
            
            fila_procesada = {
                "Concepto": "",
                "Limpieza": concepto_raw,
                "Importe": importe_raw,
                
            }
            
            # Primero intentamos identificar si es una reserva
            es_reserva = False
            for i, fila in enumerate(todas_las_filas):
                # Si aún quedan reservas en el array, se asigna la siguiente
                if i < len(array_plano):
                    fila["Concepto"] = array_plano[i]
                else:
                    fila["Concepto"] = ""
            
            # Si no fue reserva, comprobamos si es limpieza
            if not es_reserva:
                if "limpieza" in texto_low or "arreglo" in texto_low:
                    fila_procesada["Limpieza"] = "LIMPIEZA FINAL"
            
            datos_finales.append(fila_procesada)
                    
        
                    
                    
                    
        return datos_finales
        
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