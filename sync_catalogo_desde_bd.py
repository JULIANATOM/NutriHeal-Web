"""
Regenera catalogo.json a partir de ERP_DB_PROD.nutriheal (Azure SQL).
Ejecutar despues de actualizar precios/productos directamente en la base de datos.

Requiere un archivo .env en la misma carpeta con:
  DB_USER, DB_PASSWORD, DB_SERVER, DB_NAME, DB_PORT
"""
import json
import os
import pyodbc

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def cargar_env():
    env = {}
    with open(os.path.join(BASE_DIR, '.env'), encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            env[k] = v
    return env


def conectar(env):
    pwd_escaped = env['DB_PASSWORD'].replace('}', '}}')
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={env['DB_SERVER']},{env.get('DB_PORT', '1433')};"
        f"DATABASE={env['DB_NAME']};"
        f"UID={{{env['DB_USER']}}};"
        f"PWD={{{pwd_escaped}}};"
        f"Encrypt=yes;TrustServerCertificate=no;Connection Timeout=15;"
    )
    return pyodbc.connect(conn_str)


def exportar():
    env = cargar_env()
    cn = conectar(env)
    cur = cn.cursor()

    cur.execute("SELECT id, nombre FROM nutriheal.categorias ORDER BY id")
    categorias = [{"id": r.id, "nombre": r.nombre, "subcategorias": []} for r in cur.fetchall()]
    cat_by_id = {c["id"]: c for c in categorias}

    cur.execute("SELECT id, categoria_id, nombre FROM nutriheal.subcategorias ORDER BY id")
    for r in cur.fetchall():
        cat_by_id[r.categoria_id]["subcategorias"].append({"id": r.id, "nombre": r.nombre})

    cur.execute("""
        SELECT p.id, p.codigo, p.nombre, p.registro_invima, p.precio, p.observaciones, p.activo,
               s.id AS subcategoria_id, s.nombre AS subcategoria,
               c.id AS categoria_id, c.nombre AS categoria
        FROM nutriheal.productos p
        JOIN nutriheal.subcategorias s ON s.id = p.subcategoria_id
        JOIN nutriheal.categorias c ON c.id = s.categoria_id
        WHERE p.activo = 1
        ORDER BY p.id
    """)
    productos = []
    for r in cur.fetchall():
        productos.append({
            "id": r.id,
            "codigo": r.codigo,
            "nombre": r.nombre,
            "categoria": r.categoria,
            "categoria_id": r.categoria_id,
            "subcategoria": r.subcategoria,
            "subcategoria_id": r.subcategoria_id,
            "registro_invima": r.registro_invima,
            "precio": float(r.precio) if r.precio is not None else 0.0,
            "observaciones": r.observaciones,
            "activo": bool(r.activo),
        })

    catalogo = {"categorias": categorias, "productos": productos}
    out_path = os.path.join(BASE_DIR, 'catalogo.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(catalogo, f, ensure_ascii=False, indent=2)

    cn.close()
    print(f"catalogo.json actualizado: {len(categorias)} categorias, {len(productos)} productos")


if __name__ == '__main__':
    exportar()
