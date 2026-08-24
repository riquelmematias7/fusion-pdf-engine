# 🧩 Fusionador y Procesador Avanzado de PDFs Fiscales (PDF-ETL)

Script avanzado en **Python** diseñado para la manipulación, normalización, fusión gráfica y consolidación masiva de documentos PDF institucionales. Implementa técnicas de solapamiento de capas métricas (`overlay`), control de páginas anexas/huérfanas y generación automática de reportes de auditoría en Excel.

## 📋 Descripción del Proyecto
Este proyecto resuelve el desafío operativo de procesar miles de páginas de emisión de tributos y unirlas dinámicamente con plantillas de fondo institucionales (*templates* corporativos). El sistema lee un padrón de cuentas válidas (`cuentas.csv`), escanea los documentos PDF de emisión mediante extracción de texto y expresiones regulares, identifica legajos, solapa geométricamente el contenido sobre los fondos correspondientes respetando la métrica de cajas (`mediabox`), agrupa páginas de detalle asociadas y consolida un PDF maestro junto con una planilla de control y trazabilidad.

## ⚙️ Características Principales
- **Manipulación Gráfica de PDFs (Overlay):** Utiliza `pypdf` para crear páginas en blanco con dimensiones base, calcular centros geométricos exactos (`dx`, `dy`) y fusionar las páginas de contenido traducidas exactamente sobre el fondo institucional.
- **Detección Inteligente de Anexos (Anti-duplicados):** Identifica automáticamente si la página posterior a una emisión corresponde a un "detalle" asociado, agrupándola correctamente al mismo contribuyente y marcándola como utilizada para evitar solapamientos erróneos.
- **Optimización de Memoria:** Carga las bases de diseño institucionales una sola vez en memoria al inicio de la ejecución, agilizando drásticamente el procesamiento de archivos masivos.
- **Trazabilidad y Auditoría Dual:** 
  - Genera automáticamente un reporte `.xlsx` estructurado con el recuento de páginas por legajo y tasas halladas.
  - Implementa un `Logger` dual que transmite el output simultáneamente a la consola y a un archivo de registro fechado (`.log`).

## 🛠️ Stack Tecnológico
- **Lenguaje:** Python 3.x
- **Librerías principales:** 
  - `pypdf` (Manipulación, lectura, solapamiento y escritura de objetos binarios PDF)
  - `pandas` / `openpyxl` (Generación de reportes de control en Excel)
  - `tqdm` (Barra de progreso visual y métricas de rendimiento)
  - `re` / `csv` / `pathlib` (Normalización de texto y gestión de flujos)

## 🚀 Cómo probar el proyecto localmente

1. Clonar el repositorio:
   ```bash
   git clone [https://github.com/riquelmematias7/fusion-pdf-engine.git](https://github.com/riquelmematias7/fusion-pdf-engine.git)

2. Instalar las dependencias requeridas:
   ```bash
   pip install pypdf pandas openpyxl tqdm

3. (Opcional) Configurar directorios locales emision/ y bases/ con PDFs genéricos de prueba si se desea ejecutar el flujo completo de fusión gráfica.

4. Ejecutar el script principal:
   ```bash
   python fusionMerge.py
