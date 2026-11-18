import re
import sqlite3
import sys
import unicodedata
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
    # split() corta en palabras y " ".join las vuelve a unir con UN espacio:
    # quita espacios al inicio, al final y repetidos ("PARACETAMOL  30" -> "PARACETAMOL 30")
    descripcion = " ".join(descripcion.split())
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


def obtener_producto(codigo, sin_id=None):
    """Dice con qué nombre y en qué almacén está registrado un código.

    Devuelve {"almacen", "descripcion", "partidas"} o None si el código es nuevo.
    sin_id: id de una partida que NO se cuenta (la que se está editando).
    """
    with conectar() as conexion:
        # "id IS NOT ?" y no "id != ?": si sin_id es None, "!= NULL" no deja pasar
        # ninguna fila (en SQL, comparar con NULL da "desconocido"), en cambio
        # "IS NOT NULL" deja pasar todas.
        filas = conexion.execute(
            """SELECT almacen, descripcion FROM barridos
               WHERE codigo = ? AND id IS NOT ?
               ORDER BY mes DESC, id DESC""",
            (codigo.strip().upper(), sin_id),
        ).fetchall()

    if not filas:
        return None
    # Todas sus partidas deben decir lo mismo; se toma la más reciente
    return {
        "almacen": filas[0]["almacen"],
        "descripcion": filas[0]["descripcion"],
        "partidas": len(filas),
    }


def mismo_producto(producto, almacen, descripcion):
    """¿El almacén y la descripción son los que ya tiene registrados el código?"""
    return (producto["almacen"], producto["descripcion"]) == (almacen, descripcion)


def aviso_otro_producto(codigo, producto):
    """Mensaje para cuando se quiere usar un código que ya es de otro producto."""
    return (f"El código {codigo} ya está registrado como \"{producto['descripcion']}\" "
            f"en el {ALMACENES[producto['almacen']]}. Si es otro producto, revisa el "
            "código. Para corregir su nombre o almacén, edita una de sus partidas.")


def insertar_partida(mes, almacen, codigo, descripcion,
                     existencia, lote, fecha_caducidad):
    """Valida y guarda una partida NUEVA.

    Lanza ValueError si un dato es inválido o si el código ya es de otro
    producto, e IntegrityError si la partida (mes + almacén + código + lote)
    ya existe.
    """
    datos = validar_partida(mes, almacen, codigo, descripcion,
                            existencia, lote, fecha_caducidad)
    mes, almacen, prefijo, codigo, descripcion, existencia, lote, fecha_iso = datos

    # Regla: un código = un producto (un solo nombre y un solo almacén).
    # Puede tener muchos lotes, pero siempre con el mismo nombre y almacén.
    producto = obtener_producto(codigo)
    if producto and not mismo_producto(producto, almacen, descripcion):
        raise ValueError(aviso_otro_producto(codigo, producto))

    with conectar() as conexion:
        conexion.execute(
            """INSERT INTO barridos
               (mes, almacen, prefijo, codigo, descripcion,
                existencia, lote, fecha_caducidad)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            datos,
        )


def actualizar_partida(id_partida, mes, almacen, codigo, descripcion,
                       existencia, lote, fecha_caducidad, confirmar=None):
    """Valida y reemplaza los datos de una partida que ya existe.

    Si se cambió la descripción o el almacén, el producto debe quedar igual en
    TODAS sus partidas (un código = un producto). Antes de corregirlas se llama
    a confirmar(mensaje): una función que recibe la pregunta y devuelve True o
    False (la interfaz le pasa una que muestra una ventana de Sí/No).

    Devuelve True si se guardó, o False si no se confirmó la corrección.
    Lanza ValueError si un dato es inválido o si el código nuevo ya es de otro
    producto, e IntegrityError si queda igual a otra partida del mismo mes y almacén.
    """
    datos = validar_partida(mes, almacen, codigo, descripcion,
                            existencia, lote, fecha_caducidad)
    mes, almacen, prefijo, codigo, descripcion, existencia, lote, fecha_iso = datos

    # Las OTRAS partidas de este código (sin contar la que se está editando)
    producto = obtener_producto(codigo, sin_id=id_partida)
    corregir_todas = producto is not None and not mismo_producto(producto, almacen, descripcion)

    if corregir_todas:
        if codigo != obtener_partida(id_partida)["codigo"]:
            # Se cambió el código por el de OTRO producto que ya existe: no se permite
            raise ValueError(aviso_otro_producto(codigo, producto))

        # Mismo código con otro nombre o almacén: es una corrección del producto
        n = producto["partidas"]
        otras = "otra partida" if n == 1 else f"otras {n} partidas"
        mensaje = (
            f"El código {codigo} tiene {otras} (contando meses pasados) con estos datos:\n\n"
            f"    {producto['descripcion']}\n    {ALMACENES[producto['almacen']]}\n\n"
            "Para que el código siga siendo un solo producto, se cambiarán a:\n\n"
            f"    {descripcion}\n    {ALMACENES[almacen]}\n\n"
            "¿Continuar?"
        )
        if confirmar is None or not confirmar(mensaje):
            return False  # No se guarda nada

    # Los dos UPDATE van en la misma conexión: si el segundo falla, el "with"
    # deshace también el primero (o se guardan los dos, o ninguno)
    with conectar() as conexion:
        if corregir_todas:
            conexion.execute(
                """UPDATE barridos SET almacen = ?, prefijo = ?, descripcion = ?
                   WHERE codigo = ?""",
                (almacen, prefijo, descripcion, codigo),
            )
        conexion.execute(
            """UPDATE barridos
               SET mes = ?, almacen = ?, prefijo = ?, codigo = ?, descripcion = ?,
                   existencia = ?, lote = ?, fecha_caducidad = ?
               WHERE id = ?""",
            datos + (id_partida,),
        )
    return True


def eliminar_partida(id_partida):
    """Borra una partida por su id."""
    with conectar() as conexion:
        conexion.execute("DELETE FROM barridos WHERE id = ?", (id_partida,))


def fecha_para_mostrar(fecha_iso):
    """Convierte la fecha guardada '2028-06-03' al formato '03/06/2028'."""
    return datetime.strptime(fecha_iso, "%Y-%m-%d").strftime("%d/%m/%Y")


def sin_acentos(texto):
    """Prepara un texto para buscar: 'Gasas ESTÉRILES' -> 'gasas esteriles'."""
    # NFD separa cada letra de su acento: "é" -> "e" + "´" (el acento queda como
    # un carácter aparte, de la categoría "Mn"), y luego se quitan los acentos.
    separado = unicodedata.normalize("NFD", texto.lower())
    return "".join(c for c in separado if unicodedata.category(c) != "Mn")


def obtener_partidas(mes=None, almacen=None, texto=""):
    """Devuelve las partidas que cumplen los filtros indicados (todos son opcionales).

    mes: "AAAA-MM" | almacen: "1" o "2"
    texto: palabras a buscar en el código, la descripción o el lote. Sin importar
           mayúsculas, acentos ni el orden: "aguja jeringa 5" encuentra
           "Jeringa 5 ml c/aguja 21G". Cada palabra buscada debe ser el INICIO
           de una palabra del producto: "20" encuentra "20 ml", pero no el lote "JR10200".
    """
    # 1. Filtros de SQL: se juntan solo las condiciones que sí se pidieron
    condiciones = []
    parametros = []
    if mes:
        condiciones.append("mes = ?")
        parametros.append(normalizar_mes(mes))
    if almacen:
        condiciones.append("almacen = ?")
        parametros.append(almacen)

    consulta = "SELECT * FROM barridos"
    if condiciones:
        consulta += " WHERE " + " AND ".join(condiciones)
    consulta += " ORDER BY mes DESC, almacen, codigo, fecha_caducidad"

    with conectar() as conexion:
        partidas = conexion.execute(consulta, parametros).fetchall()

    # 2. Búsqueda de texto en Python (SQL no sabe ignorar acentos)
    buscadas = palabras_de(texto)
    if buscadas:
        partidas = [p for p in partidas if coincide(p, buscadas)]
    return partidas


def palabras_de(texto):
    """Separa un texto en palabras sin acentos ni mayúsculas.

    'Jeringa 5 ml c/aguja' -> ['jeringa', '5', 'ml', 'c', 'aguja']
    '1-P02836'             -> ['1', 'p02836']
    """
    # [a-z0-9]+ = "uno o más caracteres seguidos que sean letra o número".
    # Todo lo demás (espacios, "/", "-", ".") funciona como separador.
    return re.findall(r"[a-z0-9]+", sin_acentos(texto))


def coincide(partida, buscadas):
    """¿Cada palabra buscada es el inicio de alguna palabra de la partida?"""
    palabras = palabras_de(f"{partida['prefijo']}{partida['codigo']} "
                           f"{partida['descripcion']} {partida['lote']}")
    return all(any(palabra.startswith(buscada) for palabra in palabras)
               for buscada in buscadas)


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
