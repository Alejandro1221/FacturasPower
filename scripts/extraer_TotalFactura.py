from __future__ import annotations
import re,  unicodedata
from pathlib import Path
from decimal import Decimal
from typing import List, Optional

try:
    import pdfplumber
    PDFPLUMBER_OK = True
except Exception:
    from pdfminer.high_level import extract_text as pdfminer_extract_text
    PDFPLUMBER_OK = False

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DIR = BASE_DIR / "Facturas_descargadas"

# ====== Utilidades ======
def _strip_accents(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def _norm_amount(s: str) -> Optional[Decimal]:
    s = s.strip().replace(" ", "").replace(chr(160), "")
    if not s: return None
    s = re.sub(r'^[US\$COPcol\$]*', '', s)
    s = s.replace(".", "").replace(",", ".")
    try:
        return Decimal(s)
    except:
        return None

# ====== Números en letras ======
def letras_a_numero(frase: str) -> Decimal:
    frase = _strip_accents(frase.upper())
    frase = re.sub(r"[^A-ZÑ0-9\s/]", " ", frase)
    
    valores = {
        'CERO':0,'UN':1,'UNO':1,'DOS':2,'TRES':3,'CUATRO':4,'CINCO':5,'SEIS':6,'SIETE':7,'OCHO':8,'NUEVE':9,
        'DIEZ':10,'ONCE':11,'DOCE':12,'TRECE':13,'CATORCE':14,'QUINCE':15,'DIECISEIS':16,
        'DIECISIETE':17,'DIECIOCHO':18,'DIECINUEVE':19,'VEINTE':20,'VEINTI':20,'VEINTIUN':21,'VEINTIUNO':21,
        'VEINTIDOS':22,'VEINTITRES':23,'VEINTICUATRO':24,'VEINTICINCO':25,'VEINTISEIS':26,'VEINTISIETE':27,
        'VEINTIOCHO':28,'VEINTINUEVE':29,'TREINTA':30,'CUARENTA':40,'CINCUENTA':50,'SESENTA':60,
        'SETENTA':70,'OCHENTA':80,'NOVENTA':90,'CIEN':100,'CIENTO':100,'DOSCIENTOS':200,'TRESCIENTOS':300,
        'CUATROCIENTOS':400,'QUINIENTOS':500,'SEISCIENTOS':600,'SETECIENTOS':700,'OCHOCIENTOS':800,
        'NOVECIENTOS':900,'MIL':1000,'MILLON':1000000,'MILLONES':1000000
    }
    
    total = grupo = 0
    palabras = frase.split()
    i = 0
    while i < len(palabras):
        p = palabras[i]
        if p in valores:
            if p in ('MIL','MILLON','MILLONES'):
                total += (grupo if grupo else 1) * valores[p]
                grupo = 0
            else:
                grupo += valores[p]
        elif p == "VEINTI" and i+1 < len(palabras):
            sig = palabras[i+1]
            if sig in valores and valores[sig] <= 9:
                grupo += 20 + valores[sig]
                i += 1
        i += 1
    total += grupo
    if m := re.search(r"(\d\d?)/100", frase):
        total += Decimal(m.group(1)) / 100
    return Decimal(total)

# ====== Leer líneas ======
def read_lines(path: Path) -> List[str]:
    if PDFPLUMBER_OK:
        with pdfplumber.open(path) as pdf:
            texto = ""
            for page in pdf.pages:
                texto += page.extract_text() or ""
    else:
        texto = pdfminer_extract_text(str(path))
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    return lineas

# ====== EXTRAER TOTAL ======
def extraer_total(path: Path) -> dict:
    lineas = read_lines(path)
    texto_full = "\n".join(lineas)
    texto_up = texto_full.upper()
    texto_strip = _strip_accents(texto_up.replace("­", "-").replace(chr(173), "-"))  # ARREGLA GUION INVISIBLE
    
    # 1. LETRAS → PRIORIDAD MÁXIMA (AHORA INCLUYE CASOS SIN "SON" NI "PESOS")
    patrones_letras = [
        r"VALOR EN LETRAS.*?([A-ZÑ0-9\s/]{20,})", 
        r"SON[:\.\s]*([A-ZÑ0-9\s/]{15,})",
        r"([A-ZÑ0-9\s/]{20,})PESOS",
        r"([A-ZÑ0-9\s/]{20,})PESO COLOMBIANO",    
        r"([A-ZÑ0-9\s/]{20,})\s*$"                 
    ]
    for pat in patrones_letras:
        if m := re.search(pat, texto_strip, re.DOTALL):
            try:
                frase = re.sub(r"\s+PESOS?.*", "", m.group(1))  # quita "PESOS" al final
                total = letras_a_numero(frase)
                if total >= 1000:
                    return {"total": total, "metodo": "LETRAS (100% SEGURO)", "evidencia": f"LETRAS: {frase.strip()[:80]}"}
            except: pass

    # 2. TOTAL: + LÍNEA INMEDIATA (EL CASO EXACTO QUE FALLÓ)
    for i, linea in enumerate(lineas):
        if "TOTAL:" in linea.upper() and "$" in linea:
            # Busca número en la misma línea
            if nums := re.findall(r"\$\s*[\d.,]+", linea):
                if amt := _norm_amount(nums[-1]):
                    if amt >= 1000:
                        return {"total": amt, "metodo": "TOTAL: EN MISMA LÍNEA", "evidencia": linea.strip()[:100]}

    # 3. VALOR TOTAL DE LA OPERACIÓN (ya estaba)
    keywords_operacion = ["VALOR TOTAL DE LA OPERACIÓN", "TOTAL OPERACIÓN", "VALOR A PAGAR", "TOTAL NETO"]
    if any(k in texto_up for k in keywords_operacion):
        for i, linea in enumerate(lineas):
            if any(k in linea.upper() for k in keywords_operacion):
                for j in range(i, min(i+15, len(lineas))):
                    candidato = lineas[j]
                    if "$" in candidato:
                        if nums := re.findall(r"\$\s*[\d.,]+", candidato):
                            if amt := _norm_amount(nums[-1]):
                                if amt >= 1000:
                                    return {"total": amt, "metodo": "VALOR TOTAL OPERACIÓN", "evidencia": f"{linea[:50]} → {candidato[:50]}"}

    # 4. TOTAL FACTURA + 10 LÍNEAS
    for i, linea in enumerate(lineas):
        if any(x in linea.upper() for x in ["TOTAL FACTURA", "TOTAL A PAGAR"]):
            for j in range(i, min(i+10, len(lineas))):
                candidato = lineas[j]
                if "$" in candidato:
                    if nums := re.findall(r"\$\s*[\d.,]+", candidato):
                        if amt := _norm_amount(nums[-1]):
                            if amt >= 1000:
                                return {"total": amt, "metodo": "TOTAL_FACTURA_VERTICAL", "evidencia": f"{linea[:40]} → {candidato[:40]}"}

    # 5. ÚLTIMO RECURSO
    candidatos = []
    for linea in lineas:
        if "$" in linea:
            for num in re.findall(r"\$\s*[\d.,]+", linea):
                if amt := _norm_amount(num):
                    if amt >= 1000:
                        candidatos.append(amt)
    if candidatos:
        return {"total": max(candidatos), "metodo": "MAX_GLOBAL", "evidencia": "máximo con $"}

    return {"total": Decimal(0), "metodo": "FALLÓ", "evidencia": "nada"}

# ====== MAIN ======
def main():
    import argparse
    parser = argparse.ArgumentParser(description="Extractor")
    parser.add_argument("carpeta", nargs="?", default=str(DEFAULT_DIR))
    args = parser.parse_args()

    carpeta = Path(args.carpeta)
    archivos = list(carpeta.rglob("*.pdf"))
    if not archivos:
        return  

    for pdf in archivos:
        _ = extraer_total(pdf)


if __name__ == "__main__":
    main()