# Procesador Automático de Facturas

Una herramienta *standalone* diseñada para extraer tablas de archivos PDF y consolidarlas automáticamente en un reporte de Excel con formato profesional.

## 📋 Descripción
Esta aplicación automatiza el flujo de trabajo de facturación. Permite seleccionar una carpeta con múltiples archivos PDF, extraer las tablas contenidas en ellos y generar un archivo Excel consolidado (`.xlsx`). El reporte incluye cálculos automáticos de IVA, totales y un diseño visualmente limpio y profesional.

## 🚀 Características principales
* **Procesamiento por lotes**: Selecciona una carpeta completa o archivos individuales.
* **Consolidación inteligente**: Une todos los datos extraídos en una única hoja de cálculo.
* **Formato profesional**:
    * Aplicación de estilos de celda (fuentes, colores, bordes).
    * Formato de moneda (`€`) automático.
    * Altura de filas y ancho de columnas ajustados dinámicamente al contenido.
* **Cálculos automáticos**: Incluye fórmulas nativas de Excel para Totales, IVA, Subtotales y Totales con impuestos.
* **Alta compatibilidad**: Desarrollado con `openpyxl` para asegurar que el archivo sea perfectamente legible tanto en Microsoft Excel como en LibreOffice Calc.

## 🛠️ Requisitos
* **Python 3.x**
* Librerías necesarias:
  ```bash
  pip install pandas openpyxl