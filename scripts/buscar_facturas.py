import os
import sys
import requests
from dotenv import load_dotenv, set_key
from pathlib import Path
import shutil

from rutas import (
    SP_HOST, SITE_ID, DRIVE_ID, LIB_PARTIAL_NAME,
    SUBCARPETA_SERVER_REL, COLUMNA_FACTURA_INTERNAL
)

#  Config .env portátil
base_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
ENV_PATH = base_dir / ".env"

def _apply_token(token_str: str):
    global TOKEN, HEADERS
    TOKEN = (token_str or "").strip()
    HEADERS = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}

# Cargar lo que haya en .env (si no hay token, la UI lo pondrá)
load_dotenv(ENV_PATH)
_apply_token(os.getenv("GRAPH_TOKEN"))

BASE = "https://graph.microsoft.com/v1.0"

def set_graph_token(new_token: str, persist: bool = True):
    """
    Actualiza el token en memoria y opcionalmente lo persiste en .env
    (el .env vive junto al .py o al .exe).
    """
    if not new_token or len(new_token.strip()) < 20:
        raise ValueError("Token inválido.")
    if persist:
        ENV_PATH.parent.mkdir(parents=True, exist_ok=True)
        set_key(str(ENV_PATH), "GRAPH_TOKEN", new_token.strip())
    _apply_token(new_token)


#  Graph helpers
def _get(url):
    if not TOKEN:
        raise RuntimeError("Falta GRAPH_TOKEN. Configúralo desde la UI (menú Configuración → Token de Graph…).")
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


#  Solo copia local desde OneDrive
def descargar_archivo(file_ref, nombre_archivo, factura=None):
    """
    Copia el PDF desde OneDrive local hacia <base>/Facturas_descargadas.
    No descarga desde SharePoint.
    La base de OneDrive se puede configurar en .env con ONEDRIVE_BASE=...
    """
    # OneDrive base configurable por .env (portátil entre PCs)
    onedrive_base = Path(
        os.getenv("ONEDRIVE_BASE") or
        (Path.home() / "OneDrive - kumo2" / "FN SERVICIOS - FANALCA")
    )

    # Carpeta de destino junto al .py o .exe
    carpeta_descargas = base_dir / "Facturas_descargadas"
    carpeta_descargas.mkdir(parents=True, exist_ok=True)

    nombre_destino = f"{factura}.pdf" if factura else nombre_archivo
    destino = carpeta_descargas / nombre_destino

    archivo_local = onedrive_base / nombre_archivo

    if archivo_local.exists():
        try:
            shutil.copy2(archivo_local, destino)
            print(f"[LOCAL] Copiado: {archivo_local} → {destino}")
            return destino
        except Exception as e:
            print(f"Error al copiar {archivo_local}: {e}")
            return None
    else:
        print(f"No se encontró localmente: {archivo_local}")
        return None



#  Búsqueda principal
def buscar(facturas, on_progress=None):
    if on_progress:
        on_progress("Conectando a SharePoint…")

    # 1) Drive y List
    if DRIVE_ID:
        drive_id, drive_name = DRIVE_ID, "(usando DRIVE_ID configurado)"
    else:
        drives = get_site_drives(SITE_ID)
        drive_id, drive_name = pick_drive_id(drives, LIB_PARTIAL_NAME)
        if not drive_id:
            raise RuntimeError(f"No se encontró la biblioteca '{LIB_PARTIAL_NAME}'")

    list_id, list_name = get_list_id_from_drive(SITE_ID, drive_id)

    encontradas, no_encontradas, descargadas = [], [], []
    total = len(facturas)

    if on_progress:
        on_progress(f"📁 Biblioteca: {drive_name} | Lista: {list_name}")

    # 2) Buscar cada factura
    for i, fac in enumerate(facturas, start=1):
        if on_progress:
            on_progress(f"Buscando factura {i}/{total}: {fac}")

        try:
            items = listar_items_por_factura(SITE_ID, list_id, COLUMNA_FACTURA_INTERNAL, fac)
            hits = [
                it for it in items
                if (it.get("fields", {}).get("FileDirRef", "").startswith(SUBCARPETA_SERVER_REL))
            ]

            if hits:
                encontradas.append(fac)
                it0 = hits[0]
                # info del archivo
                file_ref = it0["fields"]["FileRef"]
                nombre_archivo = it0["fields"]["FileLeafRef"]

                # solo copia local desde OneDrive
                destino = descargar_archivo(file_ref, nombre_archivo, factura=fac)
                if destino:
                    descargadas.append(str(destino))
                    if on_progress:
                        on_progress(f"✔ {fac}: copiada desde OneDrive")
                else:
                    if on_progress:
                        on_progress(f"⚠️ {fac}: encontrada pero no copiada (no está sincronizada localmente)")

            else:
                no_encontradas.append(fac)
                if on_progress:
                    on_progress(f"✖ {fac}: no encontrada")

        except Exception as e:
            no_encontradas.append(fac)
            if on_progress:
                on_progress(f"⚠️ {fac}: error {e}")

    # 3) Resumen final
    if on_progress:
        on_progress(
            f"Finalizado: {len(encontradas)} encontradas, "
            f"{len(descargadas)} copiadas, {len(no_encontradas)} no encontradas."
        )

    return {
        "encontradas": encontradas,
        "no_encontradas": no_encontradas,
        "descargadas": descargadas
    }
