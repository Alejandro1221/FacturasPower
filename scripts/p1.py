from pathlib import Path
from pdfminer.high_level import extract_text

BASE_DIR = Path(__file__).resolve().parent

nombre_pdf = "SMP14931.pdf"

ruta_pdf = BASE_DIR / "Facturas_descargadas" / nombre_pdf

if not ruta_pdf.exists():
    print(f"No se encontró el archivo: {ruta_pdf}")
else:
    print(f"Leyendo texto de: {ruta_pdf}\n" + "-" * 60)
    texto = extract_text(str(ruta_pdf))
    print(texto or "No se detectó texto en el PDF.")
