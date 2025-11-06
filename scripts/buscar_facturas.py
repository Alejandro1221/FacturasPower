# buscar_facturas.py
import os, sys, shutil, requests
from pathlib import Path
from dotenv import load_dotenv, set_key

from rutas import (
    SP_HOST, SITE_ID, DRIVE_ID, LIB_PARTIAL_NAME,
    SUBCARPETA_SERVER_REL, COLUMNA_FACTURA_INTERNAL
)

BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
ENV_PATH = BASE_DIR / ".env"

def _apply_token(token_str: str):
    global TOKEN, HEADERS
    TOKEN = (token_str or "").strip()
    HEADERS = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}

# Cargar variables desde .env (si existe)
load_dotenv(ENV_PATH)
TOKEN = os.getenv("GRAPH_TOKEN") or ""
_apply_token(TOKEN)

# Ruta base de OneDrive configurable por .env (cada usuario la pone)
ONEDRIVE_BASE = os.getenv("ONEDRIVE_BASE")
if not ONEDRIVE_BASE:
    # Fallback razonable, pero se recomienda definir ONEDRIVE_BASE en .env
    ONEDRIVE_BASE = str(Path.home() / "OneDrive - kumo2" / "FN SERVICIOS - FANALCA")
onedrive_base = Path(ONEDRIVE_BASE)

BASE = "https://graph.microsoft.com/v1.0"

# === API Token desde UI ===
def set_graph_token(new_token: str, persist: bool = True):
    """
    Actualiza el token en memoria y opcionalmente lo persiste en .env (junto al .exe).
    """
    if not new_token or len(new_token.strip()) < 20:
        raise ValueError("Token inválido.")
    if persist:
        ENV_PATH.parent.mkdir(parents=True, exist_ok=True)
        set_key(str(ENV_PATH), "GRAPH_TOKEN", new_token.strip())
    _apply_token(new_token)

# === Helpers HTTP Graph ===
def _get(url: str):
    if not TOKEN:
        raise RuntimeError("Falta GRAPH_TOKEN. Configúralo desde la UI (Configuración → Token de Graph…).")
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
            return d["id"], d.get("name", "")
    return None, None

def get_list_id_from_drive(site_id, drive_id):
    data = _get(f"{BASE}/sites/{site_id}/drives/{drive_id}/list")
    return data["id"], data.get("name", "")

def listar_items_por_factura(site_id, list_id, col_internal, factura):
    """
    Devuelve todos los ítems donde fields/<col_internal> == factura.
    """
    safe = str(factura).replace("'", "''")
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

# === Copia local desde OneDrive sincronizado ===
def descargar_archivo(file_ref, nombre_archivo, factura=None):
    """
    Copia desde la carpeta local de OneDrive (ONEDRIVE_BASE) a la carpeta Facturas_descargadas junto al .exe/.py
    """
    # Validación temprana para que sea claro el error en otros PCs
    if not onedrive_base.exists():
        raise RuntimeError(
            f"ONEDRIVE_BASE no existe en este equipo:\n{onedrive_base}\n\n"
            f"Edita el archivo .env (junto al ejecutable) y define ONEDRIVE_BASE con la ruta local correcta."
        )

    # Destino (carpeta estable junto al ejecutable)
    carpeta_descargas = BASE_DIR / "Facturas_descargadas"
    carpeta_descargas.mkdir(parents=True, exist_ok=True)

    nombre_destino = f"{factura}.pdf" if factura else nombre_archivo
    destino = carpeta_descargas / nombre_destino

    # Importante: usar solo FileLeafRef (nombre de archivo) para combinar con ONEDRIVE_BASE
    archivo_local = onedrive_base / nombre_archivo

    if not archivo_local.exists():
        # Mensaje detallado para guiar al usuario
        raise FileNotFoundError(
            f"No se encontró localmente:\n{archivo_local}\n\n"
            f"• Verifica que la ruta ONEDRIVE_BASE del .env sea correcta.\n"
            f"• Asegúrate de que el PDF esté sincronizado (clic derecho → 'Siempre mantener en este dispositivo')."
        )

    shutil.copy2(archivo_local, destino)
    return destino

# === Flujo principal de búsqueda ===
def buscar(facturas, on_progress=None):
    if on_progress:
        on_progress("Conectando a SharePoint…")

    # 1) Biblioteca y lista
    if DRIVE_ID:
        drive_id, drive_name = DRIVE_ID, "(DRIVE_ID fijo)"
    else:
        drives = get_site_drives(SITE_ID)
        drive_id, drive_name = pick_drive_id(drives, LIB_PARTIAL_NAME)
        if not drive_id:
            raise RuntimeError(f"No se encontró la biblioteca '{LIB_PARTIAL_NAME}'")
    list_id, list_name = get_list_id_from_drive(SITE_ID, drive_id)

    if on_progress:
        on_progress(f"📁 Biblioteca: {drive_name} | Lista: {list_name}")

    total = len(facturas)
    encontradas, no_encontradas, descargadas = [], [], []

    # 2) Iterar facturas
    for i, fac in enumerate(facturas, start=1):
        if on_progress:
            on_progress(f"Buscando {i}/{total}: {fac}")

        try:
            items = listar_items_por_factura(SITE_ID, list_id, COLUMNA_FACTURA_INTERNAL, fac)
            hits = [
                it for it in items
                if (it.get("fields", {}).get("FileDirRef", "").startswith(SUBCARPETA_SERVER_REL))
            ]

            if hits:
                encontradas.append(fac)
                # Descargar el primero
                try:
                    file_ref = hits[0]["fields"]["FileRef"]
                    file_name = hits[0]["fields"]["FileLeafRef"]
                    destino = descargar_archivo(file_ref, file_name, factura=fac)
                    descargadas.append(str(destino))
                    if on_progress:
                        on_progress(f"[LOCAL] Copiado → {destino.name}")
                except Exception as e:
                    if on_progress:
                        on_progress(f"{fac}: Error al copiar local ({e})")
            else:
                no_encontradas.append(fac)
                if on_progress:
                    on_progress(f"{fac}: no encontrada")

        except Exception as e:
            no_encontradas.append(fac)
            if on_progress:
                on_progress(f"{fac}: error ({e})")

    # 3) Resumen
    if on_progress:
        on_progress(
            f"Finalizado: {len(encontradas)} encontradas, "
            f"{len(descargadas)} copias locales, {len(no_encontradas)} no encontradas."
        )

    return {
        "encontradas": encontradas,
        "no_encontradas": no_encontradas,
        "descargadas": descargadas,
    }
