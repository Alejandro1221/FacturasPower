import os
import requests
from dotenv import load_dotenv
from pathlib import Path
import shutil

from  rutas import (
    SP_HOST, SITE_ID, DRIVE_ID, LIB_PARTIAL_NAME,
    SUBCARPETA_SERVER_REL, COLUMNA_FACTURA_INTERNAL
)

# Cargar token desde .env
load_dotenv(dotenv_path=Path(__file__).resolve().parents[1] / ".env")

TOKEN = os.getenv("GRAPH_TOKEN")
if not TOKEN:
    raise RuntimeError("Falta GRAPH_TOKEN en el archivo .env")

HEADERS = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
BASE = "https://graph.microsoft.com/v1.0"

# Funciones auxiliares
def _get(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.json()

def get_site_drives(site_id):
    data = _get(f"{BASE}/sites/{site_id}/drives")
    return data.get("value", [])

def pick_drive_id(drives, partial_name):
    p = partial_name.lower()
    for d in drives:
        if p in d.get("name", "").lower():
            return d["id"], d["name"]
    return None, None

def get_list_id_from_drive(site_id, drive_id):
    data = _get(f"{BASE}/sites/{site_id}/drives/{drive_id}/list")
    return data["id"], data["name"]

def listar_items_por_factura(site_id, list_id, col_internal, factura):
    """Devuelve todos los ítems donde fields/<col_internal> == factura."""
    safe = factura.replace("'", "''")
    filtro = f"fields/{col_internal} eq '{safe}'"
    select = f"fields($select=FileRef,FileDirRef,FileLeafRef,{col_internal})"
    url = (f"{BASE}/sites/{site_id}/lists/{list_id}/items"
           f"?$expand={select}&$filter={filtro}&$top=200")

    items = []
    while url:
        data = _get(url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items

#FUNCIÓN PRINCIPAL
def descargar_archivo(file_ref, nombre_archivo, factura=None):
    #onedrive_base = Path(r"C:\Users\practicante1servicio\OneDrive - kumo2\FN SERVICIOS - FANALCA")
    onedrive_base = Path.home() / "OneDrive - kumo2" / "FN SERVICIOS - FANALCA"
    nombre_destino = f"{factura}.pdf" if factura else nombre_archivo
    destino = Path(__file__).parent / "Facturas_descargadas" / nombre_destino
    destino.parent.mkdir(exist_ok=True, parents=True)

    archivo_local = onedrive_base / nombre_archivo  # el nombre largo real

    if archivo_local.exists():
        try:
            shutil.copy2(archivo_local, destino)
            return destino
        except Exception as e:
            raise RuntimeError(f"Error al copiar el archivo local: {e}")
    else:
        raise FileNotFoundError(f"No se encontró localmente: {archivo_local}")

def buscar(facturas):
    if DRIVE_ID:
        drive_id, drive_name = DRIVE_ID, "(usando DRIVE_ID configurado)"
    else:
        drives = get_site_drives(SITE_ID)
        drive_id, drive_name = pick_drive_id(drives, LIB_PARTIAL_NAME)
        if not drive_id:
            raise RuntimeError(f"No se encontró la biblioteca '{LIB_PARTIAL_NAME}'")

    list_id, list_name = get_list_id_from_drive(SITE_ID, drive_id)

    encontradas = []
    no_encontradas = []

    total = len(facturas)
    print(f"\nIniciando búsqueda en SharePoint ({total} facturas)...")
    print(f"Biblioteca: {drive_name}\nLista: {list_name}\n")

    for i, fac in enumerate(facturas, start=1):
        print(f"[{i}/{total}] Buscando factura: {fac} ...", end=" ", flush=True)

        try:
            items = listar_items_por_factura(SITE_ID, list_id, COLUMNA_FACTURA_INTERNAL, fac)
            hits = []
            for it in items:
                f = it.get("fields", {}) or {}
                dir_ref = f.get("FileDirRef", "")
                if not dir_ref.startswith(SUBCARPETA_SERVER_REL):
                    continue
                hits.append(it)

            if hits:
                encontradas.append(fac)
                print("✔ ENCONTRADA")
                for h in hits:
                    f = h["fields"]
                    file_ref = f.get("FileRef")
                    file_name = f.get("FileLeafRef", f"{fac}.pdf")
                    try:
                        local_path = descargar_archivo(file_ref, file_name, fac)
                        print(f"   ⤷ Copiado como: {local_path.name}")
                    except Exception as e:
                        print(f"Error al descargar {file_name}: {e}")
            else:
                no_encontradas.append(fac)
                print("✖ NO encontrada")

        except Exception as e:
            print(f"Error: {e}")
            no_encontradas.append(fac)

    print("\n=== RESUMEN FINAL ===")
    print(f"Encontradas: {len(encontradas)}")
    print(f"No encontradas: {len(no_encontradas)}")
    print("=====================\n")

    return {
        "encontradas": encontradas,
        "no_encontradas": no_encontradas
    }