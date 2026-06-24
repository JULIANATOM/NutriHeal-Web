# -*- coding: utf-8 -*-
"""
Elimina del container de blob los blobs bajo las carpetas raíz antiguas de NutriVita:
  Fitoterapéuticos/
  Medicamentos/
  Suplementos dietarios/

Estas carpetas existen ahora correctamente bajo NutriVita/.
Los blobs raíz son duplicados obsoletos.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
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
from azure.storage.blob import BlobServiceClient

CONN_STR  = _env['AZURE_STORAGE_CONNECTION_STRING']
CONTAINER = _env['AZURE_STORAGE_CONTAINER']

# Prefijos a eliminar (carpetas raíz obsoletas de NutriVita)
PREFIJOS_A_ELIMINAR = [
    'Fitoterapéuticos/',
    'Medicamentos/',
    'Suplementos dietarios/',
]


def main():
    client    = BlobServiceClient.from_connection_string(CONN_STR)
    container = client.get_container_client(CONTAINER)

    # Recopilar blobs a eliminar
    a_eliminar = []
    for prefijo in PREFIJOS_A_ELIMINAR:
        for b in container.list_blobs(name_starts_with=prefijo):
            a_eliminar.append(b.name)

    if not a_eliminar:
        print('No se encontraron blobs con los prefijos obsoletos. Nada que eliminar.')
        return

    print(f'Se eliminarán {len(a_eliminar)} blobs:')
    for nombre in a_eliminar:
        print(f'  {nombre}')

    confirmar = input(f'\n¿Confirmar eliminación de {len(a_eliminar)} blobs? (s/n): ').strip().lower()
    if confirmar != 's':
        print('Cancelado.')
        return

    eliminados = errores = 0
    for nombre in a_eliminar:
        try:
            container.delete_blob(nombre)
            print(f'  ✓ Eliminado: {nombre}')
            eliminados += 1
        except Exception as e:
            print(f'  ✗ Error: {nombre} — {e}')
            errores += 1

    print(f'\n=== Resumen ===')
    print(f'  Eliminados: {eliminados}')
    print(f'  Errores   : {errores}')


if __name__ == '__main__':
    main()
