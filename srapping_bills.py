import customtkinter as ctk
import flet as ft
import os
import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter  import filedialog, messagebox
import numpy as np

@ft.control
class GUI(ft.Column):
    def __init__(self):
        super().__init__()
        
        self.ruta_origen = ""
        self.ruta_destino = ""
        
        self.contenedor_lista=ft.Column()
        self.width = 600
        self.controls = [
            ft.Column(
                controls=[
               
                ft.Row(
                    controls=[
                        ft.Button("Elige la factura", 
                                icon=ft.Icons.FILE_COPY,
                                on_click=self.seleccionar_archivo),
                        ft.Button("Elige el archivo excel de destino",
                                icon=ft.Icons.LIST, 
                              on_click=self.seleccionar_destino),
                        ft.Button("Procesar Facturas",
                                    style=ft.ButtonStyle(
                                    color=ft.Colors.WHITE,
                                    bgcolor=ft.Colors.GREEN_800,
                                    overlay_color=ft.Colors.GREEN_600,
                               ),
                               icon=ft.Icons.FILE_DOWNLOAD,
                               icon_color=ft.Colors.WHITE,
                               
                               
                               on_click=self.procesar_todo),
                ],
                    alignment=ft.MainAxisAlignment.CENTER, 
                    spacing= 15
            ),
            ft.Column(
                
                controls=[
                    ft.Text("Facturas seleccionadas:", weight="bold"),
                    ft.Container(
                    content=self.contenedor_lista,
                    bgcolor=ft.Colors.BLUE_GREY_600,
                    padding=20,
                    border_radius=10,
                    width=600,
                    height=300
            ),
                    

                ],
            )
            ]) 
            
        ]

    def actualizar_lista_visual(self, archive):
        # Limpiamos y recreamos la lista de textos
        self.contenedor_lista.controls = [
            ft.Text(f"• {archive}", size=12, color=ft.Colors.WHITE) 
        ]
        self.update()
    def seleccionar_archivo(self):
        
        # 1. Creamos una ventana raíz oculta de Tkinter
        root = tk.Tk()
        root.withdraw()
        
        # 2. Hacemos que esta ventana esté siempre encima (esto ayuda con el foco)
        root.attributes("-topmost", True)
        
        # 3. Abrimos el diálogo
        archivo_pdf = filedialog.askopenfilename(
            parent=root, # Le decimos explícitamente que el padre es esta ventana
            title="Selecciona un archivo PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        
        # 4. Destruimos la ventana raíz para limpiar memoria
        root.destroy()
        
        if archivo_pdf:
            self.ruta_origen = archivo_pdf
            self.actualizar_lista_visual(archivo_pdf)
            print(f"PDF seleccionado: {self.ruta_origen}")
            # Si necesitas actualizar la UI de Flet inmediatamente:
            # self.update()

    def seleccionar_destino(self):
        root = tk.Tk()
        root.withdraw()
        
        # 2. Hacemos que esta ventana esté siempre encima (esto ayuda con el foco)
        root.attributes("-topmost", True)
        self.ruta_destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        print(f"Destino: {self.ruta_destino}")
         # 4. Destruimos la ventana raíz para limpiar memoria
        root.destroy()
    
    def limpiar_concepto(self, text):

        match_date_and_code = re.search(r'Ref Reserva:\s(?P<codigo>\d{8,})\s\((?P<fecha_inicio>\d{1,2}/\d{1,2}/\d{2,4})\s-\s(?P<fecha_fin>\d{1,2}/\d{1,2}/\d{2,4})\)')

    def extraer_texto_completo(self, pdf_path):
        
        texto_total = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Extrae el texto plano de toda la página
                texto_total += page.extract_text() + "\n"
        
        return texto_total

    def calcular_diferencia_fechas(self, lista_cadenas):
        resultados = []
        
        for item in lista_cadenas:
            # 1.1 Extraer solo los números de los días (lo que está antes de la primera '/')
            # La regex '\d{2}' busca grupos de dos dígitos. 
            # En '(23/07/2026 - 27/07/2026)', esto encontrará '23' y '27'.
            dias = re.findall(r'(\d{2})/', item)
            
            # 1.2 y 1.3 Convertir a lista de ints
            if len(dias) >= 2:
                dias_int = [int(d) for d in dias]
                
                # 2. Restarlos (27 - 23)
                # Nota: Usamos abs() por si el orden de los días está invertido
                resta = abs(dias_int[1] - dias_int[0])

                nights= resta - 1
                
                # 3. Crear el array final
                resultados.append([nights])
            
        return resultados
    
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

        # 1. Convertir las fechas dentro del paréntesis en dos elementos de una lista. Ej: '(23/07/2026 - 27/072026)'
        #   1.1 Expresión regular que filtre los guiones, los meses y los años. Ej: '23', '27'
        #   1.2 Convertirlos en una lista : un array de dos numeros dentro de una lista Ej: [['23', '27'], ...]
        #   1.3 Convertirlos en int. [[23, 27]...]
        # 2. Restarlos [[23-27]...]
        # 3. Crear un array de numeros  [[5]...]  
        nights = self.calcular_diferencia_fechas(array_plano)
        print (nights)

        # 1. Filtrar las fechas eliminando todo lo que está fuera del paréntesis. Ej: Ref Reserva: 31682853 (27/04/2026 - 04/05/2026) -> '27/04/2026 - 04/05/2026'
        # 2. Convertirlas en una lista de arrays: cada array tendría la fecha de inicio y la fecha de salida. Ej: [['27/04/2026', '04/05/2026'],...]
        # 3. Restar las fechas dentro de cada array: Mediante un bucle for se itera los arrays mientras otro itera los elementos dentro de los arrays. En esta última iteración se hacen las operaciones DataFrame
        #   3.1 Primera iteración sobre los arrays de la lista
        #       3.1.1 Segunda iteración dentro del array : restar las fechas
        #       3.1.2 Devuelve la diferencia
        #   3.2 La diferencia se añade a un nuevo array
        # Pre-calcular precios de limpieza en el orden de todas_las_filas
        precios_limpieza = []
        for fila in todas_las_filas:
            concepto = str(fila.get("Concepto", "")).lower()
            if "limpieza" in concepto or "arreglo" in concepto:
                precios_limpieza.append(fila.get("Importe", 50))
            else:
                precios_limpieza.append(0)
        # Mantenemos solo los elementos que NO son cero
        precios = [x for x in precios_limpieza if x != 0]
        print (precios)
        lista_precios_limpieza = [ float (precio.replace('€', '').replace(',', '.').strip())
                                  for precio in precios
                                  ]
        print (lista_precios_limpieza)
        precios_reserva = []
        for fila in todas_las_filas:
            precio = fila.get("Importe", "")
            if precios_limpieza[0] != precio:
                precios_reserva.append(fila.get("Importe"))
        set_reserva = set(precios)
        reserva = [x for x in precios_reserva if x not in set_reserva] 
        
        lista_precios_reserva = []
        for v in reserva:
            try:
                if isinstance(v, (int, float)):
                    number = float(v)
                elif isinstance(v, str):
                    clean= v.replace('€', '').replace(',','.').strip()
                    number = float(clean)
                    
                else:
                    continue
                if number > 0:
                    lista_precios_reserva.append(number)
            except (ValueError, TypeError):
                continue
            

                                  
        print (lista_precios_reserva)            
        
        

        # Resultado: [50, 100, 75]
        datos_finales = []
        
        
        print(len(precios))
        
        for i, fila in enumerate(todas_las_filas):
            concepto_raw = str(fila.get("Concepto", ""))
            importe_raw = fila.get("Importe") or None
            
            fila_procesada = {
                "Concepto": array_plano[i] if i < len(array_plano) else None,
                "Limpieza": lista_precios_limpieza[i] if i < len(lista_precios_limpieza) else None,
                "Importe": lista_precios_reserva[i] if i < len(lista_precios_reserva)-1 else None,    
            }
            print(fila_procesada["Concepto"])
            print(fila_procesada["Limpieza"])
            print(fila_procesada["Importe"])
            
            
            # Si no fue reserva, comprobamos si es limpieza
            
            datos_finales.append(fila_procesada)
                    
        
                    
                    
                    
        return datos_finales
        
    def actualizar_excel(self, nueva_lista_datos, ruta_excel):
        if not nueva_lista_datos: return

        try:
            # 1. Preparar datos nuevos
            columnas = ['Concepto', 'Noches', 'Limpieza', 'Importe']
            df_nuevos = pd.DataFrame(nueva_lista_datos)
            df_nuevos.replace([0, '0', '', '0,0', '0,00', '0 €', 0.0], np.nan, inplace=True)
            df_nuevos.dropna(subset=['Importe'], inplace=True)
            fechas_extraidas = df_nuevos['Concepto'].str.extract(r'\((.*?)\)')[0]
            separadas = fechas_extraidas.str.split(' - ', expand=True)
            f_inicio = pd.to_datetime(separadas[0], format='%d/%m/%Y', errors='coerce')
            f_fin = pd.to_datetime(separadas[1], format='%d/%m/%Y', errors='coerce')
            df_nuevos['Noches'] = (f_fin - f_inicio).dt.days - 1
            df_nuevos['Noches'] = df_nuevos['Noches'].fillna(0).astype(int)
            # 2. Cargar datos existentes de forma segura
            if os.path.exists(ruta_excel):
                try:
                    df_existente = pd.read_excel(ruta_excel)
                    # Forzamos las columnas si el archivo existe pero está vacío/mal formado
                    if df_existente.empty or 'Importe' not in df_existente.columns:
                        df_existente = pd.DataFrame(columns= columnas)
                    else:
                        # Si tiene datos, limpiamos los "TOTAL" y ceros
                        df_existente.replace([0, '0', '', '0,0', '0,00', '0 €', 0.0], np.nan, inplace=True)
                        df_existente.dropna(subset=['Importe'], inplace=True)
                        df_existente = df_existente[~df_existente['Concepto'].astype(str).str.contains('TOTAL', case=False, na=False)]
                        df_existente = df_existente[~df_existente['Concepto'].astype(str).str.contains('IVA', case=False, na=False)]
                        df_existente = df_existente[~df_existente['Concepto'].astype(str).str.contains('SUBTOTAL DEL IVA', case=False, na=False)]
                        df_existente = df_existente[~df_existente['Concepto'].astype(str).str.contains('TOTAL SIN IMPUESTOS', case=False, na=False)]
                        df_existente = df_existente[~df_existente['Concepto'].astype(str).str.contains('TOTAL CON IMPUESTOS', case=False, na=False)]
                except:
                    df_existente = pd.DataFrame(columns=columnas)
            else:
                df_existente = pd.DataFrame(columns=columnas)

            # 3. Concatenar y asegurar tipos numéricos
            df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
            df_final['Importe'] = pd.to_numeric(df_final['Importe'], errors='coerce').fillna(0)
            df_final = df_final[df_final['Importe'] != 0]
            
            with pd.ExcelWriter(ruta_excel, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Facturas')
                
                # Formato y fórmula
                workbook = writer.book
                worksheet = writer.sheets['Facturas']
                formato_euro = workbook.add_format({'num_format': '#,##0.00 €'})
                worksheet.set_column('C:C', 15, formato_euro)
                worksheet.set_column('D:D', 15, formato_euro)
                
                
                num_filas = len(df_final) + 1
                worksheet.write(f'A{num_filas + 2}', 'TOTAL', workbook.add_format({'bold': True}))
                worksheet.write(f'A{num_filas + 3}', 'IVA', workbook.add_format({'bold': True}))
                worksheet.write(f'A{num_filas + 4}', 'SUBTOTAL DEL IVA', workbook.add_format({'bold': True}))
                worksheet.write(f'A{num_filas + 5}', 'TOTAL SIN IMPUESTOS', workbook.add_format({'bold': True}))
                worksheet.write(f'A{num_filas + 6}', 'TOTAL CON IMPUESTOS', workbook.add_format({'bold': True}))
                worksheet.write_formula(f'B{num_filas + 2}', f'=SUM(B2:B{num_filas})')
                worksheet.write_formula(f'C{num_filas + 2}', f'=SUM(C2:C{num_filas})', formato_euro) 
                worksheet.write_formula(f'D{num_filas + 2}', f'=SUM(D2:D{num_filas})', formato_euro)
                worksheet.write_formula(f'C{num_filas + 3}', f'=SUM(C{num_filas + 2})*0.21', formato_euro)
                worksheet.write_formula(f'D{num_filas + 3}', f'=SUM(D{num_filas + 2})*0.21', formato_euro)
                worksheet.write_formula(f'D{num_filas + 4}', f'=SUM(C{num_filas + 3}, D{num_filas + 3})', formato_euro)
                worksheet.write_formula(f'D{num_filas + 5}', f'=SUM(C{num_filas + 2},D{num_filas + 2})', formato_euro)
                worksheet.write_formula(f'D{num_filas + 6}', f'=SUM(D{num_filas + 4}, D{num_filas + 5})', formato_euro)
                

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

def main(page: ft.Page):
    # Instanciamos nuestra clase y la añadimos a la página
    app = GUI()
    page.add(app)
    page.title = "Procesador de facturas"
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.update()
    

    
ft.run(main)