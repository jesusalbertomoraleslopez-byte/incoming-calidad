import os
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BD_PARAMETROS = os.path.join(BASE_DIR, "BD_Parametros_Materia_Prima.xlsx")
BD_REPORTES = os.path.join(BASE_DIR, "BD_Reportes_Incoming.xlsx")
BD_ATADOS = os.path.join(BASE_DIR, "BD_Atados_Incoming.xlsx")
BD_SALIDAS = os.path.join(BASE_DIR, "BD_Salidas_Incoming.xlsx")
PLANTILLA_PATH = os.path.join(BASE_DIR, "plantilla_incoming_calidad.xlsx")
CARPETAS_DIR = os.path.join(BASE_DIR, "carpetas_electronicas")

GCS_BUCKET = os.environ.get("GCS_BUCKET", "").strip()
_GCS_READY = False

_GCS_EXCEL_FILES = {
    "data/BD_Parametros_Materia_Prima.xlsx": BD_PARAMETROS,
    "data/BD_Reportes_Incoming.xlsx": BD_REPORTES,
    "data/BD_Atados_Incoming.xlsx": BD_ATADOS,
    "data/BD_Salidas_Incoming.xlsx": BD_SALIDAS,
    "data/plantilla_incoming_calidad.xlsx": PLANTILLA_PATH,
}

def _gcs_bucket():
    from google.cloud import storage
    return storage.Client().bucket(GCS_BUCKET)

def _gcs_local_path(blob_name):
    if blob_name in _GCS_EXCEL_FILES:
        return _GCS_EXCEL_FILES[blob_name]
    if blob_name.startswith("carpetas_electronicas/"):
        return os.path.join(BASE_DIR, *blob_name.split("/"))
    if blob_name.startswith("data/"):
        filename = blob_name.replace("data/", "")
        return os.path.join(BASE_DIR, filename)
    return None

def sync_from_gcs():
    """Descarga bases de datos y carpetas electrónicas desde GCS al arrancar."""
    global _GCS_READY
    if not GCS_BUCKET:
        return False
    try:
        bucket = _gcs_bucket()
        count = 0
        for blob in bucket.list_blobs():
            if blob.name.endswith("/"):
                continue
            local = _gcs_local_path(blob.name)
            if not local:
                continue
            os.makedirs(os.path.dirname(local), exist_ok=True)
            # Solo descargar si no existe o si el tamano difiere
            if not os.path.exists(local) or os.path.getsize(local) != blob.size:
                blob.download_to_filename(local)
                count += 1
        _GCS_READY = True
        print(f"[GCS] Sincronizados exitosamente {count} archivos desde gs://{GCS_BUCKET}")
        return True
    except Exception as e:
        print(f"[GCS] Error al sincronizar desde gs://{GCS_BUCKET}: {e}")
        return False

def push_file_to_gcs(local_path):
    """Sube un archivo especifico (Excel o carpeta electronica) al bucket."""
    if not (GCS_BUCKET and _GCS_READY):
        return False
    try:
        if not os.path.exists(local_path):
            return False
        bucket = _gcs_bucket()
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
    if not (GCS_BUCKET and _GCS_READY):
        return False
    try:
        folio_dir = os.path.join(CARPETAS_DIR, folio)
        if not os.path.isdir(folio_dir):
            return False
        bucket = _gcs_bucket()
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
    if not (GCS_BUCKET and _GCS_READY):
        return False
    try:
        bucket = _gcs_bucket()
        prefix = f"carpetas_electronicas/{folio}/"
        blobs = list(bucket.list_blobs(prefix=prefix))
        for blob in blobs:
            blob.delete()
        print(f"[GCS] Eliminados {len(blobs)} archivos del folio {folio} en GCS.")
        return True
    except Exception as e:
        print(f"[GCS] Error al eliminar folio {folio} de GCS: {e}")
        return False
