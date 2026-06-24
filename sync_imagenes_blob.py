# -*- coding: utf-8 -*-
"""
Sincroniza TODAS las imágenes de Imagenes/Catalogo_Productos/ con Azure Blob Storage.
La estructura local ya incluye la carpeta de marca como primer nivel:
  Local:  Imagenes/Catalogo_Productos/{Marca}/{categoria}/{subcategoria}/{codigo}.png
  Blob:   {Marca}/{categoria}/{subcategoria}/{codigo}.png

Reglas especiales para NutriVita:
  - NutriVita/Medicamentos/Sistema Nervioso/MELATONINA.png  →  NV1670.png
  - NutriVita/Medicamentos/Articular/NV1691.png  →  NutriVita/Fitoterapéuticos/Articular/NV1691.png
    (DYPHLAMIN fue reclasificado de Medicamentos a Fitoterapéuticos)
"""
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
LOCAL_ROOT = os.path.join(BASE_DIR, 'Imagenes', 'Catalogo_Productos')

def _load_env():
    env = {}
    with open(os.path.join(BASE_DIR, '.env'), encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            env[k] = v
    return env

_env = _load_env()
from azure.storage.blob import BlobServiceClient, ContentSettings

CONN_STR  = _env['AZURE_STORAGE_CONNECTION_STRING']
CONTAINER = _env['AZURE_STORAGE_CONTAINER']

# Ruta local relativa (desde LOCAL_ROOT) → blob name
RENAME_RULES = {
    'NutriVita/Medicamentos/Sistema Nervioso/MELATONINA.png':
        'NutriVita/Medicamentos/Sistema Nervioso/NV1670.png',
    'NutriVita/Medicamentos/Sistema Nervioso/MELATONINA_01.png':
        'NutriVita/Medicamentos/Sistema Nervioso/NV1670_01.png',
    'NutriVita/Medicamentos/Articular/NV1691.png':
        'NutriVita/Fitoterapéuticos/Articular/NV1691.png',
    'NutriVita/Medicamentos/Articular/NV1691_01.png':
        'NutriVita/Fitoterapéuticos/Articular/NV1691_01.png',
}


def main():
    client    = BlobServiceClient.from_connection_string(CONN_STR)
    container = client.get_container_client(CONTAINER)

    print('Leyendo blobs existentes en Azure...')
    existing = {b.name: b.size for b in container.list_blobs()}
    print(f'  {len(existing)} blobs en container "{CONTAINER}"')

    archivos = []
    for root, _dirs, files in os.walk(LOCAL_ROOT):
        for fname in files:
            full = os.path.join(root, fname)
            rel  = os.path.relpath(full, LOCAL_ROOT).replace('\\', '/')
            archivos.append((full, rel))

    archivos.sort(key=lambda x: x[1])
    print(f'  {len(archivos)} archivos locales a procesar\n')

    ok = skipped = errores = 0

    for full_path, rel_path in archivos:
        blob_name  = RENAME_RULES.get(rel_path, rel_path)
        local_size = os.path.getsize(full_path)

        if blob_name in existing and existing[blob_name] == local_size:
            print(f'  = {blob_name}  ({local_size:,} B)')
            skipped += 1
            continue

        action = 'NUEVO' if blob_name not in existing else 'ACTUALIZADO'
        try:
            with open(full_path, 'rb') as data:
                container.upload_blob(
                    name=blob_name,
                    data=data,
                    overwrite=True,
                    content_settings=ContentSettings(content_type='image/png'),
                )
            print(f'  ✓ {blob_name}  ({local_size:,} B) [{action}]')
            ok += 1
        except Exception as e:
            print(f'  ✗ {blob_name} — ERROR: {e}')
            errores += 1

    print(f'\n=== Resumen ===')
    print(f'  Subidos/actualizados : {ok}')
    print(f'  Sin cambios (omitidos): {skipped}')
    print(f'  Errores              : {errores}')


if __name__ == '__main__':
    main()
