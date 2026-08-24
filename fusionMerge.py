import re
import csv
import os
import sys
import datetime
from pypdf import PdfReader, PdfWriter, PageObject
from tqdm import tqdm
from openpyxl import Workbook

# =====================================================
# CONFIGURACIÓN GENERAL
# =====================================================
emision_publicidad = "emision/emision_publicidad.pdf"
emision_ocupacion = "emision/emision_ocupacion.pdf"
emision_antenas = "emision/emision_antenas.pdf"
archivo_cuentas = "cuentas.csv"

base_publicidad = "bases/base_publicidad.pdf"
base_ocupacion = "bases/base_ocupacion.pdf"
base_antenas = "bases/base_antenas.pdf"

archivo_salida = "resultado.pdf"
carpeta_logs = "logs"
os.makedirs(carpeta_logs, exist_ok=True)

# =====================================================
# SISTEMA DE LOGS
# =====================================================
timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_file = os.path.join(carpeta_logs, f"fusionMerge_{timestamp}.log")
sys.stdout.reconfigure(encoding='utf-8')

class Logger:
    def __init__(self, log_path):
        self.terminal = sys.__stdout__
        self.log = open(log_path, "a", encoding="utf-8")
    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()
    def flush(self): pass

sys.stdout = Logger(log_file)

print("===================================================")
print(f"🕒 INICIO DEL PROCESAMIENTO: {datetime.datetime.now():%d/%m/%Y %H:%M:%S}")
print("===================================================")

# =====================================================
# 1. CARGAR CUENTAS VÁLIDAS
# =====================================================
cuentas_validas = set()
with open(archivo_cuentas, newline="", encoding="utf-8-sig") as f:
    reader = csv.reader(f)
    for row in reader:
        for col in row:
            col = col.strip()
            if col.isdigit():
                cuentas_validas.add(col)

print(f"📂 Se cargaron {len(cuentas_validas)} cuentas válidas desde {archivo_cuentas}.")

# =====================================================
# 2. CARGAR BASES UNA SOLA VEZ (optimización de memoria)
# =====================================================
bases_cargadas = {}
for tipo, ruta in {
    "PUBLICIDAD": base_publicidad,
    "OCUPACION": base_ocupacion,
    "ANTENAS": base_antenas
}.items():
    try:
        reader = PdfReader(ruta)
        bases_cargadas[tipo] = [reader.pages[0], reader.pages[1]]
        print(f"🧩 Base {tipo} cargada correctamente ({ruta}).")
    except Exception as e:
        print(f"❌ Error al cargar base {tipo}: {e}")
        sys.exit(1)

# =====================================================
# 3. FUNCIONES AUXILIARES
# =====================================================
def fusionar_con_base(pagina_origen, paginas_base):
    """Fusiona una página con el fondo institucional correspondiente."""
    pagina_base, pagina_extra = paginas_base
    ancho_base = float(pagina_base.mediabox.width)
    alto_base = float(pagina_base.mediabox.height)
    ancho_pag = float(pagina_origen.mediabox.width)
    alto_pag = float(pagina_origen.mediabox.height)

    dx = (ancho_base - ancho_pag) / 2
    dy = (alto_base - alto_pag) / 2

    fondo = PageObject.create_blank_page(width=ancho_base, height=alto_base)
    fondo.merge_page(pagina_base)
    fondo.merge_translated_page(pagina_origen, dx, dy)
    return [fondo, pagina_extra]

def procesar_pdf(ruta_pdf, tipo):
    """Procesa cada PDF de emisión (por tasa)."""
    if not os.path.exists(ruta_pdf):
        print(f"❌ No se encontró el archivo {ruta_pdf}, se omite.")
        return []

    print(f"\n🔎 Procesando archivo {ruta_pdf} ({tipo}) ...")
    reader = PdfReader(ruta_pdf)
    resultados = []

    for i, pagina in enumerate(tqdm(reader.pages, desc=f"{tipo}", unit="pág", colour="green"), start=1):
        if getattr(pagina, "usada", False):
            continue

        texto = pagina.extract_text() or ""
        texto_norm = texto.upper()
        numeros = re.findall(r"\b\d+\b", texto_norm)
        legajo = next((n for n in numeros if n in cuentas_validas), None)

        if not legajo:
            continue

        fusionadas = fusionar_con_base(pagina, bases_cargadas[tipo])
        resultados.append({"legajo": legajo, "tipo": tipo, "paginas": fusionadas})

        # Detectar si la página siguiente corresponde a un detalle del mismo legajo
        if i < len(reader.pages):
            siguiente = reader.pages[i]
            if getattr(siguiente, "usada", False):
                continue

            texto_sig = siguiente.extract_text() or ""
            texto_sig_norm = texto_sig.upper()
            numeros_sig = re.findall(r"\b\d+\b", texto_sig_norm)
            mismo_legajo = legajo in numeros_sig
            tiene_tipo_sig = any(p in texto_sig_norm for p in ["PUBLICIDAD", "OCUPACION", "CAP. XXII"])

            if mismo_legajo and not tiene_tipo_sig:
                fusionadas_detalle = fusionar_con_base(siguiente, bases_cargadas[tipo])
                resultados[-1]["paginas"].extend(fusionadas_detalle)
                setattr(siguiente, "usada", True)
                print(f"   ➕ Página {i+1} agregada como detalle del legajo {legajo}")

    print(f"✅ {len(resultados)} legajos procesados en {tipo}.")
    return resultados

# =====================================================
# 4. PROCESAR LAS TASAS
# =====================================================
publi = procesar_pdf(emision_publicidad, "PUBLICIDAD")
ocupa = procesar_pdf(emision_ocupacion, "OCUPACION")
antenas = procesar_pdf(emision_antenas, "ANTENAS")

# =====================================================
# 5. ORGANIZAR POR LEGAJO
# =====================================================
print("\n📊 Organizando legajos...")

paginas_por_legajo = {}
tasas_por_legajo = {}

for lote in [publi, ocupa, antenas]:
    for item in lote:
        legajo = item["legajo"]
        paginas_por_legajo.setdefault(legajo, []).extend(item["paginas"])
        tasas_por_legajo.setdefault(legajo, set()).add(item["tipo"])

print(f"📌 Se organizaron {len(paginas_por_legajo)} legajos distintos.")

# =====================================================
# 6. GENERAR PDF FINAL
# =====================================================
print("\n📝 Generando PDF final consolidado...")

writer = PdfWriter()
info_legajos = []
for legajo, paginas in tqdm(paginas_por_legajo.items(), desc="Fusionando", unit="legajo", colour="cyan"):
    for p in paginas:
        writer.add_page(p)
    info_legajos.append({
        "legajo": legajo,
        "paginas": len(paginas),
        "tasas": ", ".join(sorted(tasas_por_legajo.get(legajo, set()))),
        "fecha": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    })

with open(archivo_salida, "wb") as f:
    writer.write(f)

print(f"\n🎉 Proceso completado. Archivo final: {archivo_salida}")

# =====================================================
# 7. EXPORTAR EXCEL DE CONTROL
# =====================================================
print("\n📘 Generando Excel de control...")

wb = Workbook()
ws = wb.active
ws.title = "Legajos"
ws.append(["Legajo", "Páginas", "Tasas Encontradas", "Fecha Procesado"])

for item in info_legajos:
    ws.append([item["legajo"], item["paginas"], item["tasas"], item["fecha"]])

excel_path = os.path.join(carpeta_logs, f"fusionMerge_{timestamp}.xlsx")
wb.save(excel_path)

print(f"✅ Excel generado correctamente en: {excel_path}")
print(f"🧾 Log guardado en: {os.path.abspath(log_file)}")
print("===================================================")