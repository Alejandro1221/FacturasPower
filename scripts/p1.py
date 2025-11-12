from pathlib import Path
from pdfminer.high_level import extract_text

# Ruta base: carpeta donde está este script
BASE_DIR = Path(__file__).resolve().parent

# Cambia este nombre por el de la factura que quieras probar
nombre_pdf = "FTAR1288.pdf"

ruta_pdf = BASE_DIR / "Facturas_descargadas" / nombre_pdf

if not ruta_pdf.exists():
    print(f"⚠️ No se encontró el archivo: {ruta_pdf}")
else:
    print(f"📄 Leyendo texto de: {ruta_pdf}\n" + "-" * 60)
    texto = extract_text(str(ruta_pdf))
    print(texto or "⚠️ No se detectó texto en el PDF.")
