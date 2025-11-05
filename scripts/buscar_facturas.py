import os, sys
import requests
from dotenv import load_dotenv,set_key
from pathlib import Path
import shutil

from  rutas import (
    SP_HOST, SITE_ID, DRIVE_ID, LIB_PARTIAL_NAME,
    SUBCARPETA_SERVER_REL, COLUMNA_FACTURA_INTERNAL
)

# Cargar token desde .env
#ENV_PATH = Path(__file__).resolve().parents[1] / ".env"
base_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
ENV_PATH = base_dir / ".env"

def _apply_token(token_str: str):
    global TOKEN, HEADERS
    TOKEN = token_str.strip()
    HEADERS = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}

# lee el token actual (por si ya existe en .env)
load_dotenv(ENV_PATH)
TOKEN = os.getenv("GRAPH_TOKEN") or ""
_apply_token(TOKEN)

BASE = "https://graph.microsoft.com/v1.0"


def set_graph_token(new_token: str, persist: bool = True):
    """
    Actualiza el token en memoria y opcionalmente lo persiste en .env.
    """
    if not new_token or len(new_token.strip()) < 20:
        raise ValueError("Token inválido.")
    if persist:
        ENV_PATH.parent.mkdir(parents=True, exist_ok=True)
        set_key(str(ENV_PATH), "GRAPH_TOKEN", new_token.strip())
    _apply_token(new_token)

# Funciones auxiliares
def _get(url):
    if not TOKEN:
        raise RuntimeError("Falta GRAPH_TOKEN. Configúralo desde la UI (Configurar token).")
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
    onedrive_base = Path.home() / "OneDrive - kumo2" / "FN SERVICIOS - FANALCA"

    base_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
    carpeta_descargas = base_dir / "Facturas_descargadas"
    carpeta_descargas.mkdir(parents=True, exist_ok=True)

    nombre_destino = f"{factura}.pdf" if factura else nombre_archivo
    destino = carpeta_descargas / nombre_destino

    archivo_local = onedrive_base / file_ref.replace("/", "\\")  # <-- FIX RUTA

    if archivo_local.exists():
        try:
            shutil.copy2(archivo_local, destino)
            return destino
        except Exception as e:
            print(f"Error al copiar {archivo_local}: {e}")
            return None
    else:
        print(f"No encontrado localmente: {archivo_local}")
        return None

def buscar(facturas, on_progress=None):
    if on_progress:
        on_progress("Conectando a SharePoint…")

    # =1. Localizar biblioteca y lista 
    if DRIVE_ID:
        drive_id, drive_name = DRIVE_ID, "(usando DRIVE_ID configurado)"
    else:
        drives = get_site_drives(SITE_ID)
        drive_id, drive_name = pick_drive_id(drives, LIB_PARTIAL_NAME)
        if not drive_id:
            raise RuntimeError(f"No se encontró la biblioteca '{LIB_PARTIAL_NAME}'")

    list_id, list_name = get_list_id_from_drive(SITE_ID, drive_id)

    total = len(facturas)
    encontradas, no_encontradas, descargadas = [], [], []

    if on_progress:
        on_progress(f"Biblioteca: {drive_name} | Lista: {list_name}")

    # === 2. BUSCAR Y DESCARGAR CADA FACTURA ===
    for i, fac in enumerate(facturas, start=1):
        if on_progress:
            on_progress(f"Buscando {i}/{total}: {fac}")

        try:
            items = listar_items_por_factura(SITE_ID, list_id, COLUMNA_FACTURA_INTERNAL, fac)
            hits = [
                it for it in items
                if it.get("fields", {}).get("FileDirRef", "").startswith(SUBCARPETA_SERVER_REL)
            ]

            if hits:
                encontradas.append(fac)
                # === DESCARGAR EL PRIMER PDF ===
                try:
                    file_ref = hits[0]["fields"]["FileRef"]
                    nombre_archivo = hits[0]["fields"]["FileLeafRef"]
                    destino = descargar_archivo(file_ref, nombre_archivo, factura=fac)
                    descargadas.append(str(destino))
                    if on_progress:
                        on_progress(f"{fac} → Descargado")
                except Exception as e:
                    if on_progress:
                        on_progress(f"{fac}: Error al descargar ({e})")
            else:
                no_encontradas.append(fac)
                if on_progress:
                    on_progress(f"{fac}: no encontrada")

        except Exception as e:
            no_encontradas.append(fac)
            if on_progress:
                on_progress(f"{fac}: error ({e})")

    # === 3. RESUMEN FINAL ===
    if on_progress:
        on_progress(
            f"Finalizado: {len(encontradas)} encontradas, "
            f"{len(descargadas)} descargadas, {len(no_encontradas)} no encontradas."
        )

    return {
        "encontradas": encontradas,
        "no_encontradas": no_encontradas,
        "descargadas": descargadas
    }


