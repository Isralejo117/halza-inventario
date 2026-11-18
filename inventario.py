import sqlite3
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

# Carpeta donde viven los datos (la base de datos y los respaldos):
# - Al programar (python interfaz.py): la carpeta de este archivo .py
# - Dentro del .exe: la carpeta donde está el .exe. Ahí NO sirve __file__,
#   porque apunta a una carpeta temporal que Windows borra al cerrar el programa.
if getattr(sys, "frozen", False):  # PyInstaller pone sys.frozen = True en el .exe
    CARPETA_DATOS = Path(sys.executable).parent
else:
    CARPETA_DATOS = Path(__file__).parent

DB_PATH = CARPETA_DATOS / "inventario_halza.db"

# Respaldo automático: una copia por día en esta carpeta; se guardan las últimas 30
CARPETA_RESPALDOS = CARPETA_DATOS / "respaldos"
RESPALDOS_A_GUARDAR = 30

ALMACENES = {
    "1": "Almacén Chico (Medicamentos)",
    "2": "Almacén Grande (Uso Médico)",
}



def conectar():
    """Abre la conexión al archivo .db (lo crea si no existe)."""
    conexion = sqlite3.connect(DB_PATH)
    conexion.row_factory = sqlite3.Row
    return conexion

def crear_tabla():
    """Crea la tabla de barridos solo si todavía no existe."""
    with conectar() as conexion:
        conexion.execute("""
            CREATE TABLE IF NOT EXISTS barridos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                mes TEXT NOT NULL,
                almacen TEXT NOT NULL CHECK (almacen IN ('1', '2')),
                prefijo TEXT NOT NULL,
                codigo TEXT NOT NULL,
                descripcion TEXT NOT NULL,
                existencia INTEGER NOT NULL CHECK (existencia >= 0),
                lote TEXT NOT NULL,
                fecha_caducidad TEXT NOT NULL,
                UNIQUE (mes, almacen, codigo, lote)
            )
        """)


def normalizar_mes(mes):
    """Valida un mes y lo devuelve como AAAA-MM (ej. '2026-1' -> '2026-01')."""
    try:
        return datetime.strptime(mes.strip(), "%Y-%m").strftime("%Y-%m")
    except ValueError:
        raise ValueError("Mes inválido. Usa el formato AAAA-MM (ej. 2026-11).")


def validar_partida(mes, almacen, codigo, descripcion,
                    existencia, lote, fecha_caducidad):
    """Revisa y limpia los datos de una partida.

    almacen: "1" o "2" | mes: AAAA-MM | fecha_caducidad: DD/MM/AAAA
    Devuelve una tupla lista para la base de datos, en este orden:
    (mes, almacen, prefijo, codigo, descripcion, existencia, lote, fecha_iso)
    Lanza ValueError con un mensaje claro si algún dato es inválido.
    """
    # 1. Almacén y prefijo (igual que en el Excel: "1-" o "2-")
    almacen = str(almacen).strip()
    if almacen not in ALMACENES:
        raise ValueError("El almacén debe ser 1 (chico) o 2 (grande).")
    prefijo = f"{almacen}-"

    # 2. Textos obligatorios
    codigo = codigo.strip().upper()
    descripcion = descripcion.strip()
    lote = lote.strip().upper()
    if not codigo or not descripcion or not lote:
        raise ValueError("Código, descripción y lote son obligatorios.")

    # 3. Mes normalizado a AAAA-MM
    mes = normalizar_mes(mes)

    # 4. Existencia: entero no negativo
    try:
        existencia = int(existencia)
    except ValueError:
        raise ValueError("La existencia debe ser un número entero (ej. 30).")
    if existencia < 0:
        raise ValueError("La existencia no puede ser negativa.")

    # 5. Caducidad: el usuario escribe DD/MM/AAAA; se guarda como AAAA-MM-DD
    try:
        fecha_iso = datetime.strptime(fecha_caducidad.strip(), "%d/%m/%Y").date().isoformat()
    except ValueError:
        raise ValueError("Caducidad inválida. Usa el formato dd/mm/aaaa (ej. 03/06/2028).")

    return (mes, almacen, prefijo, codigo, descripcion, existencia, lote, fecha_iso)


def insertar_partida(mes, almacen, codigo, descripcion,
                     existencia, lote, fecha_caducidad):
    """Valida y guarda una partida NUEVA.

    Lanza ValueError si un dato es inválido, e IntegrityError si la
    partida (mes + almacén + código + lote) ya existe.
    """
    datos = validar_partida(mes, almacen, codigo, descripcion,
                            existencia, lote, fecha_caducidad)
    with conectar() as conexion:
        conexion.execute(
            """INSERT INTO barridos
               (mes, almacen, prefijo, codigo, descripcion,
                existencia, lote, fecha_caducidad)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            datos,
        )


def actualizar_partida(id_partida, mes, almacen, codigo, descripcion,
                       existencia, lote, fecha_caducidad):
    """Valida y reemplaza los datos de una partida que ya existe.

    Lanza los mismos errores que insertar_partida (IntegrityError si con
    los cambios queda igual a otra partida del mismo mes y almacén).
    """
    datos = validar_partida(mes, almacen, codigo, descripcion,
                            existencia, lote, fecha_caducidad)
    with conectar() as conexion:
        conexion.execute(
            """UPDATE barridos
               SET mes = ?, almacen = ?, prefijo = ?, codigo = ?, descripcion = ?,
                   existencia = ?, lote = ?, fecha_caducidad = ?
               WHERE id = ?""",
            datos + (id_partida,),
        )


def eliminar_partida(id_partida):
    """Borra una partida por su id."""
    with conectar() as conexion:
        conexion.execute("DELETE FROM barridos WHERE id = ?", (id_partida,))


def fecha_para_mostrar(fecha_iso):
    """Convierte la fecha guardada '2028-06-03' al formato '03/06/2028'."""
    return datetime.strptime(fecha_iso, "%Y-%m-%d").strftime("%d/%m/%Y")


def obtener_partidas(mes=None):
    """Devuelve todas las partidas, o solo las de un mes si se indica."""
    consulta = "SELECT * FROM barridos"
    parametros = ()
    if mes:
        consulta += " WHERE mes = ?"
        parametros = (normalizar_mes(mes),)
    consulta += " ORDER BY mes DESC, almacen, codigo, fecha_caducidad"

    with conectar() as conexion:
        return conexion.execute(consulta, parametros).fetchall()


def obtener_partida(id_partida):
    """Devuelve una sola partida por su id (o None si no existe)."""
    with conectar() as conexion:
        return conexion.execute("SELECT * FROM barridos WHERE id = ?",
                                (id_partida,)).fetchone()


def mes_anterior(mes):
    """Devuelve el mes anterior: '2026-11' -> '2026-10' y '2026-01' -> '2025-12'."""
    primer_dia = datetime.strptime(normalizar_mes(mes), "%Y-%m")  # Día 1 de ese mes
    ultimo_dia_del_anterior = primer_dia - timedelta(days=1)
    return ultimo_dia_del_anterior.strftime("%Y-%m")


def comparar_meses(actual, anterior):
    """Compara lote por lote el barrido de dos meses.

    Devuelve una lista con un diccionario por cada lote que aparece en
    cualquiera de los dos meses. Además de los datos del lote trae:
      existencia_anterior / existencia_actual: número, o None si ese mes
                                               el lote no se registró
      estado: "Nuevo", "Ya no está", "Cambió" o "Sin cambio"
      diferencia: actual - anterior (un lote que no se registró cuenta como 0)
    """
    consulta = """
        WITH actual AS (SELECT * FROM barridos WHERE mes = ?),
             anterior AS (SELECT * FROM barridos WHERE mes = ?)
        SELECT
            COALESCE(actual.almacen, anterior.almacen)                 AS almacen,
            COALESCE(actual.prefijo, anterior.prefijo)                 AS prefijo,
            COALESCE(actual.codigo, anterior.codigo)                   AS codigo,
            COALESCE(actual.descripcion, anterior.descripcion)         AS descripcion,
            COALESCE(actual.lote, anterior.lote)                       AS lote,
            COALESCE(actual.fecha_caducidad, anterior.fecha_caducidad) AS fecha_caducidad,
            anterior.existencia AS existencia_anterior,
            actual.existencia   AS existencia_actual
        FROM actual
        FULL OUTER JOIN anterior
            ON  actual.almacen = anterior.almacen
            AND actual.codigo  = anterior.codigo
            AND actual.lote    = anterior.lote
        ORDER BY almacen, codigo, fecha_caducidad
    """
    with conectar() as conexion:
        filas = conexion.execute(consulta, (normalizar_mes(actual),
                                            normalizar_mes(anterior))).fetchall()

    lotes = []
    for fila in filas:
        lote = dict(fila)  # Una fila de sqlite3 no se puede modificar; un dict sí
        antes = lote["existencia_anterior"]
        ahora = lote["existencia_actual"]

        if antes is None:
            lote["estado"] = "Nuevo"
        elif ahora is None:
            lote["estado"] = "Ya no está"
        elif ahora != antes:
            lote["estado"] = "Cambió"
        else:
            lote["estado"] = "Sin cambio"

        lote["diferencia"] = (ahora or 0) - (antes or 0)
        lotes.append(lote)
    return lotes


def respaldar(ruta_destino):
    """Guarda en ruta_destino una copia completa de la base de datos.

    Usa la función de respaldo de SQLite en lugar de copiar el archivo: así
    la copia sale completa aunque en ese momento se esté guardando algo.
    Lanza ValueError si el destino es la misma base de datos.
    """
    destino = Path(ruta_destino)
    if destino.resolve() == DB_PATH.resolve():
        raise ValueError("El respaldo no puede guardarse encima de la base de datos original.")
    destino.unlink(missing_ok=True)  # Si ya había un archivo con ese nombre, se reemplaza

    origen = conectar()
    copia = sqlite3.connect(destino)
    try:
        origen.backup(copia)
    finally:
        # Se cierran las dos conexiones pase lo que pase, aunque el respaldo falle
        copia.close()
        origen.close()


def respaldo_automatico():
    """Hace el respaldo del día si todavía no existe, y borra los más viejos.

    Los respaldos se llaman auto_AAAA-MM-DD.db y viven en CARPETA_RESPALDOS.
    Solo se conservan los RESPALDOS_A_GUARDAR más recientes. Los demás
    archivos de la carpeta (por ejemplo, respaldos manuales) no se tocan.
    Devuelve la ruta del respaldo automático más reciente, o None si no hay.
    """
    CARPETA_RESPALDOS.mkdir(exist_ok=True)  # Crea la carpeta si no existe

    de_hoy = CARPETA_RESPALDOS / f"auto_{date.today().isoformat()}.db"
    if not de_hoy.exists() and obtener_partidas():  # Si la base está vacía, no hay qué respaldar
        respaldar(de_hoy)

    # Gracias al formato AAAA-MM-DD, ordenar por nombre es ordenar por fecha
    respaldos = sorted(CARPETA_RESPALDOS.glob("auto_*.db"))
    for viejo in respaldos[:-RESPALDOS_A_GUARDAR]:  # Todos menos los últimos 30
        viejo.unlink()

    return respaldos[-1] if respaldos else None


if __name__ == "__main__":
    crear_tabla()
    try:
        insertar_partida("2026-11", "1", "P02836",
                         "Paracetamol 30 mg 10/c", 30, "KH233224", "03/06/2028")
        insertar_partida("2026-11", "1", "P02836",
                         "Paracetamol 30 mg 10/c", 7, "KH999111", "15/01/2027")
    
    except sqlite3.IntegrityError:
        print("Esa partida ya estaba registrada.")

    for fila in obtener_partidas("2026-11"):
        print(ALMACENES[fila["almacen"]], "|",
              fila["prefijo"] + fila["codigo"], fila["lote"],
              fila["existencia"], fecha_para_mostrar(fila["fecha_caducidad"]))
