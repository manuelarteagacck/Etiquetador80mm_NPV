import pyodbc
import json
import re

# Cargar configuración
with open("config.json", "r", encoding="utf-8") as f:
    cfg = json.load(f)

conn_str = (
    f"DRIVER={cfg['conexion_odbc']['driver']};"
    f"SERVER={cfg['conexion_odbc']['server']};"
    f"DATABASE={cfg['conexion_odbc']['database']};"
    f"UID={cfg['conexion_odbc']['username']};"
    f"PWD={cfg['conexion_odbc']['password']};"
)

print("Probando conexión...")
try:
    conn = pyodbc.connect(conn_str, timeout=5)
    print("✅ Conexión exitosa a la base de datos NPV.")
except Exception as e:
    print("❌ Error de conexión:", e)
    raise SystemExit()

cursor = conn.cursor()

# ----------- PRUEBA 1: Buscar por ARTÍCULO -----------
articulo = input("Ingresa un ARTICULO (ej. AB1010012): ").strip()
sql_articulo = """
SELECT TOP 1
    a.ARTICULO,
    a.DESCRIPCION,
    pv.PRECIOSVENTA AS PRECIO,
    e.EQUIVALENTE AS UPC
FROM NPV.dbo.NPVFDPreciosVenta AS pv
JOIN NPV.dbo.NPVFDArticulos AS a ON a.ARTICULO = pv.ARTICULO
LEFT JOIN NPV.dbo.NPVFDArticulosEquivalentes AS e ON e.ARTICULO = pv.ARTICULO
WHERE pv.ARTICULO = ?
  AND (pv.FECHAINICIO IS NULL OR pv.FECHAINICIO <= GETDATE())
ORDER BY pv.FECHAINICIO DESC;
"""

row = cursor.execute(sql_articulo, (articulo,)).fetchone()
if row:
    print("\n✅ Resultado por ARTICULO:")
    print(f"  ARTICULO : {row.ARTICULO}")
    print(f"  DESCRIPCION : {row.DESCRIPCION}")
    print(f"  PRECIO : {row.PRECIO}")
    print(f"  UPC : {row.UPC}")
else:
    print("⚠️ No se encontró el artículo en la base de datos.")

# ----------- PRUEBA 2: Buscar por UPC -----------
upc = input("\nIngresa un UPC (12 o 13 dígitos): ").strip()
upc_digits = re.sub(r"[^0-9]", "", upc)

sql_upc = """
SELECT TOP 1
    a.ARTICULO,
    a.DESCRIPCION,
    pv.PRECIOSVENTA AS PRECIO,
    e.EQUIVALENTE AS UPC
FROM NPV.dbo.NPVFDArticulosEquivalentes AS e
JOIN NPV.dbo.NPVFDArticulos AS a ON a.ARTICULO = e.ARTICULO
JOIN NPV.dbo.NPVFDPreciosVenta AS pv ON pv.ARTICULO = a.ARTICULO
WHERE e.EQUIVALENTE = ?
  AND (pv.FECHAINICIO IS NULL OR pv.FECHAINICIO <= GETDATE())
ORDER BY pv.FECHAINICIO DESC;
"""

row = cursor.execute(sql_upc, (upc_digits,)).fetchone()
if row:
    print("\n✅ Resultado por UPC:")
    print(f"  ARTICULO : {row.ARTICULO}")
    print(f"  DESCRIPCION : {row.DESCRIPCION}")
    print(f"  PRECIO : {row.PRECIO}")
    print(f"  UPC : {row.UPC}")
else:
    print("⚠️ No se encontró el UPC en la base de datos.")

conn.close()
print("\nPrueba finalizada ✅")
