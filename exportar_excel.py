"""Exporta los barridos a un libro de Excel (.xlsx) para que las oficinas
puedan filtrar, ordenar y sumar a su manera."""
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from inventario import (ALMACENES, comparar_meses, mes_anterior, normalizar_mes,
                        obtener_partidas)

NEGRITAS = Font(bold=True)
FONDO_ENCABEZADO = PatternFill("solid", fgColor="D9D9D9")  # Gris claro
FORMATO_FECHA = "DD/MM/YYYY"
FORMATO_CON_SIGNO = "+0;-0;0"  # Formato de Excel: positivos;negativos;cero

# Columnas de cada hoja: (encabezado, ancho en caracteres)
COLUMNAS_PARTIDAS = (
    ("Mes", 10), ("Almacén", 30), ("Código", 12), ("Descripción", 40),
    ("Lote", 14), ("Caducidad", 12), ("Existencia", 11),
)
COLUMNAS_COMPARACION = (
    ("Estado", 12), ("Almacén", 30), ("Código", 12), ("Descripción", 40),
    ("Lote", 14), ("Caducidad", 12), ("Anterior", 10), ("Actual", 10), ("Diferencia", 11),
)

# Mismos colores que la pestaña Comparar (en Excel van sin el "#")
COLORES_ESTADO = {
    "Nuevo": "DFF3E4",
    "Ya no está": "FDE0E0",
    "Cambió": "FFF1C7",
}


def escribir_hoja(hoja, columnas, filas):
    """Escribe el encabezado y las filas en una hoja, con formato cómodo para Excel.

    columnas: tuplas (encabezado, ancho) | filas: listas de valores en el orden de las columnas
    """
    hoja.append([encabezado for encabezado, ancho in columnas])
    for fila in filas:
        hoja.append(fila)

    # Encabezado (fila 1) en negritas con fondo gris
    for celda in hoja[1]:
        celda.font = NEGRITAS
        celda.fill = FONDO_ENCABEZADO

    # Ancho de cada columna: la 1 es "A", la 2 es "B", etc.
    for numero, (encabezado, ancho) in enumerate(columnas, start=1):
        hoja.column_dimensions[get_column_letter(numero)].width = ancho

    # Las fechas se guardan como fechas de verdad y se muestran dd/mm/aaaa
    for fila in hoja.iter_rows(min_row=2):
        for celda in fila:
            if isinstance(celda.value, date):
                celda.number_format = FORMATO_FECHA

    hoja.freeze_panes = "A2"                 # El encabezado se queda fijo al bajar
    hoja.auto_filter.ref = hoja.dimensions   # Flechitas de filtro en cada encabezado


def agregar_total(hoja, letras):
    """Agrega abajo la suma de las columnas indicadas (ej. ["G"]).

    SUBTOTAL(9, ...) suma solo las filas que se ven, así que el total
    se actualiza solo cuando en Excel se filtra por almacén, producto, etc.
    """
    ultima = hoja.max_row
    fila = ultima + 2  # Deja una fila vacía entre los datos y el total
    hoja[f"A{fila}"] = "Total (filas visibles)"
    hoja[f"A{fila}"].font = NEGRITAS
    for letra in letras:
        celda = hoja[f"{letra}{fila}"]
        celda.value = f"=SUBTOTAL(9,{letra}2:{letra}{ultima})"
        celda.font = NEGRITAS
    return fila


def hoja_partidas(libro, partidas):
    """Llena la primera hoja con una fila por partida."""
    hoja = libro.active  # Un libro nuevo ya trae una hoja; se usa esa
    hoja.title = "Partidas"

    filas = []
    for p in partidas:
        filas.append([
            p["mes"],
            ALMACENES[p["almacen"]],
            p["prefijo"] + p["codigo"],
            p["descripcion"],
            p["lote"],
            date.fromisoformat(p["fecha_caducidad"]),  # "2028-06-03" -> fecha real
            p["existencia"],
        ])
    escribir_hoja(hoja, COLUMNAS_PARTIDAS, filas)
    agregar_total(hoja, ["G"])


def hoja_comparacion(libro, mes):
    """Agrega una hoja con la comparación lote por lote contra el mes anterior."""
    hoja = libro.create_sheet("Comparación")

    filas = []
    for lote in comparar_meses(mes, mes_anterior(mes)):
        filas.append([
            lote["estado"],
            ALMACENES[lote["almacen"]],
            lote["prefijo"] + lote["codigo"],
            lote["descripcion"],
            lote["lote"],
            date.fromisoformat(lote["fecha_caducidad"]),
            lote["existencia_anterior"],  # None -> la celda queda vacía
            lote["existencia_actual"],
            lote["diferencia"],
        ])
    escribir_hoja(hoja, COLUMNAS_COMPARACION, filas)
    fila_total = agregar_total(hoja, ["G", "H", "I"])

    # Color de cada fila según su estado (la columna A tiene el estado)
    for fila in hoja.iter_rows(min_row=2, max_row=len(filas) + 1):
        color = COLORES_ESTADO.get(fila[0].value)  # None si es "Sin cambio"
        if color:
            for celda in fila:
                celda.fill = PatternFill("solid", fgColor=color)

    # La diferencia con signo: +7, -11, 0
    for numero_fila in range(2, fila_total + 1):
        hoja[f"I{numero_fila}"].number_format = FORMATO_CON_SIGNO


def generar_excel(mes, ruta):
    """Crea en ruta un .xlsx con las partidas del mes (AAAA-MM), o de todos
    los meses si mes viene vacío.

    Si es un solo mes y el mes anterior tiene datos, agrega la hoja Comparación.
    """
    mes = normalizar_mes(mes) if mes else ""
    libro = Workbook()
    hoja_partidas(libro, obtener_partidas(mes))
    if mes and obtener_partidas(mes_anterior(mes)):
        hoja_comparacion(libro, mes)
    libro.save(ruta)
