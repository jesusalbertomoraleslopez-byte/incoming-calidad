"""
=============================================================================
INCOMING CALIDAD - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia de Alto Rendimiento con Google Cloud Storage (Cloud Run)
- Cliente Singleton reutilizable
- Caché en memoria con TTL (300s) y control de concurrencia thread-safe
- Solo descarga archivos nuevos o modificados
=============================================================================
"""

import os
import time
import threading
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BD_PARAMETROS = os.path.join(BASE_DIR, "BD_Parametros_Materia_Prima.xlsx")
BD_REPORTES = os.path.join(BASE_DIR, "BD_Reportes_Incoming.xlsx")
BD_ATADOS = os.path.join(BASE_DIR, "BD_Atados_Incoming.xlsx")
BD_SALIDAS = os.path.join(BASE_DIR, "BD_Salidas_Incoming.xlsx")
PLANTILLA_PATH = os.path.join(BASE_DIR, "plantilla_incoming_calidad.xlsx")
CARPETAS_DIR = os.path.join(BASE_DIR, "carpetas_electronicas")

GCS_BUCKET = os.environ.get("GCS_BUCKET", "").strip()

_STORAGE_CLIENT = None
_BUCKET_OBJ = None
_SYNC_LOCK = threading.Lock()
_LAST_SYNC_TS = 0.0
_GCS_READY = False
SYNC_TTL = 300.0  # 5 minutos

_GCS_EXCEL_FILES = {
    "data/BD_Parametros_Materia_Prima.xlsx": BD_PARAMETROS,
    "data/BD_Reportes_Incoming.xlsx": BD_REPORTES,
    "data/BD_Atados_Incoming.xlsx": BD_ATADOS,
    "data/BD_Salidas_Incoming.xlsx": BD_SALIDAS,
    "data/plantilla_incoming_calidad.xlsx": PLANTILLA_PATH,
}


def get_storage_client():
    """Retorna una instancia singleton de storage.Client()."""
    global _STORAGE_CLIENT
    if _STORAGE_CLIENT is None:
        from google.cloud import storage
        _STORAGE_CLIENT = storage.Client()
    return _STORAGE_CLIENT


def _gcs_bucket():
    """Retorna una instancia singleton del bucket."""
    global _BUCKET_OBJ
    if not GCS_BUCKET:
        return None
    if _BUCKET_OBJ is None or _BUCKET_OBJ.name != GCS_BUCKET:
        _BUCKET_OBJ = get_storage_client().bucket(GCS_BUCKET)
    return _BUCKET_OBJ


def _gcs_local_path(blob_name):
    if blob_name in _GCS_EXCEL_FILES:
        return _GCS_EXCEL_FILES[blob_name]
    if blob_name.startswith("carpetas_electronicas/"):
        return os.path.join(BASE_DIR, *blob_name.split("/"))
    if blob_name.startswith("data/"):
        filename = blob_name.replace("data/", "")
        return os.path.join(BASE_DIR, filename)
    return None


def sync_from_gcs(force: bool = False) -> bool:
    """Descarga bases de datos y carpetas electrónicas desde GCS al arrancar o si expira el TTL."""
    global _GCS_READY, _LAST_SYNC_TS
    if not GCS_BUCKET:
        return False

    now = time.time()
    if not force and (now - _LAST_SYNC_TS) < SYNC_TTL and _GCS_READY:
        return True

    with _SYNC_LOCK:
        now = time.time()
        if not force and (now - _LAST_SYNC_TS) < SYNC_TTL and _GCS_READY:
            return True

        try:
            bucket = _gcs_bucket()
            if not bucket:
                return False

            count = 0
            for blob in bucket.list_blobs():
                if blob.name.endswith("/"):
                    continue
                local = _gcs_local_path(blob.name)
                if not local:
                    continue
                os.makedirs(os.path.dirname(local), exist_ok=True)
                if not os.path.exists(local) or os.path.getsize(local) != blob.size:
                    blob.download_to_filename(local)
                    count += 1
            _GCS_READY = True
            _LAST_SYNC_TS = time.time()
            print(f"[GCS] Sincronizados exitosamente {count} archivos desde gs://{GCS_BUCKET}")
            return True
        except Exception as e:
            print(f"[GCS] Error al sincronizar desde gs://{GCS_BUCKET}: {e}")
            return False


def push_file_to_gcs(local_path):
    """Sube un archivo especifico (Excel o carpeta electronica) al bucket."""
    if not GCS_BUCKET:
        return False
    try:
        if not os.path.exists(local_path):
            return False
        bucket = _gcs_bucket()
        if not bucket:
            return False
        rel = os.path.relpath(local_path, BASE_DIR).replace(os.sep, "/")
        if rel.startswith("carpetas_electronicas/"):
            blob_name = rel
        else:
            blob_name = f"data/{os.path.basename(local_path)}"
        bucket.blob(blob_name).upload_from_filename(local_path)
        print(f"[GCS] Archivo subido: {blob_name}")
        return True
    except Exception as e:
        print(f"[GCS] Error al subir {local_path} a GCS: {e}")
        return False


def push_folio_dir_to_gcs(folio):
    """Sube todos los archivos generados para un folio al bucket."""
    if not GCS_BUCKET:
        return False
    try:
        folio_dir = os.path.join(CARPETAS_DIR, folio)
        if not os.path.isdir(folio_dir):
            return False
        bucket = _gcs_bucket()
        if not bucket:
            return False
        for root, _, files in os.walk(folio_dir):
            for f in files:
                full_path = os.path.join(root, f)
                rel = os.path.relpath(full_path, BASE_DIR).replace(os.sep, "/")
                bucket.blob(rel).upload_from_filename(full_path)
        print(f"[GCS] Carpeta electronica del folio {folio} sincronizada con GCS.")
        return True
    except Exception as e:
        print(f"[GCS] Error al sincronizar carpeta {folio} con GCS: {e}")
        return False


def delete_folio_from_gcs(folio):
    """Elimina del bucket todos los archivos correspondientes a un folio."""
    if not GCS_BUCKET:
        return False
    try:
        bucket = _gcs_bucket()
        if not bucket:
            return False
        prefix = f"carpetas_electronicas/{folio}/"
        blobs = list(bucket.list_blobs(prefix=prefix))
        for blob in blobs:
            blob.delete()
        print(f"[GCS] Eliminados {len(blobs)} archivos del folio {folio} en GCS.")
        return True
    except Exception as e:
        print(f"[GCS] Error al eliminar folio {folio} de GCS: {e}")
        return False
