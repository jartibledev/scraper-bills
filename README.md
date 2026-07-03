# Procesador Automático de Facturas
Una herramienta diseñada para automatizar flujos de trabajo extrayendo tablas de facturas en archivos PDF y consolidarlas automáticamente en un reporte de Excel con formato profesional.

## 📋 Descripción
Esta aplicación automatiza el flujo de trabajo de facturación. Permite seleccionar un múltiples facturas en PDF, extraer las tablas contenidas en ellos y generar un archivo Excel consolidado (`.xlsx`). El reporte incluye cálculos automáticos de IVA, las cantidade totales y un diseño visualmente limpio y profesional.

## 🚀 Características principales
* **Procesamiento por lotes**: Selecciona un archivo pdf o varios a la vez.
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
    ```
## ⚙️ Notas técnicas
* **Motor de escritura**: Se utiliza openpyxl como motor principal, lo que permite la manipulación precisa de estilos XML y asegura que los formatos persistan entre distintas suites de ofimática.
* **Lógica de ajuste**: El script calcula automáticamente el ancho de las columnas basándose en el contenido de las cabeceras y los datos, garantizando la legibilidad.

## 🛠️ Tecnologías utilizadas
* **Pandas**: Manipulación de DataFrames y consolidación de datos tabulares.
* **OpenPyXL**: Estilizado avanzado de celdas, fórmulas y estructuración del libro de Excel.
* **Flet/Tkinter**: (Basado en la interfaz de tu proyecto) para la gestión de ventanas y selección de rutas.

## 📄 Licencia

Este proyecto es de uso privado para automatización de tareas.
Desarrollado para optimizar la gestión contable y administrativa.

