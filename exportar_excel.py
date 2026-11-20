"""Exporta los barridos a un libro de Excel (.xlsx) para que las oficinas
puedan filtrar, ordenar y sumar a su manera."""
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from inventario import (ALMACENES, FECHA_NO_CADUCA, NO_CADUCA, comparar_meses,
                        fecha_para_mostrar, limite_de_caducidad, mes_anterior,
                        normalizar_mes, obtener_partidas)
from reporte import COLORES_ESTADO

NEGRITAS = Font(bold=True)
CURSIVAS = Font(italic=True)
FONDO_ENCABEZADO = PatternFill("solid", fgColor="D9D9D9")  # Gris claro
FONDO_POR_CADUCAR = PatternFill("solid", fgColor="F8CACA")  # Mismo rojo que en Consulta
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


def caducidad_para_excel(fecha_iso):
    """'2028-06-03' -> fecha real de Excel; los productos que no caducan -> texto 'No caduca'."""
    if fecha_iso == FECHA_NO_CADUCA:
        return NO_CADUCA
    return date.fromisoformat(fecha_iso)


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
            caducidad_para_excel(p["fecha_caducidad"]),
            p["existencia"],
        ])
    escribir_hoja(hoja, COLUMNAS_PARTIDAS, filas)
    fila_total = agregar_total(hoja, ["G"])

    # Rojo: lotes con menos de 1 año para caducar, igual que en Consulta.
    # La partida 0 de la lista quedó en la fila 2 de Excel (la 1 es el encabezado).
    limite = limite_de_caducidad()
    por_caducar = 0
    for numero_fila, p in enumerate(partidas, start=2):
        if p["fecha_caducidad"] < limite:
            por_caducar += 1
            for celda in hoja[numero_fila]:  # hoja[2] = todas las celdas de la fila 2
                celda.fill = FONDO_POR_CADUCAR

    # Nota debajo del total. El Excel es una "foto" del día en que se exporta:
    # por eso se dice contra qué fecha se calculó el rojo.
    nota = hoja[f"A{fila_total + 1}"]
    lotes = "1 lote" if por_caducar == 1 else f"{por_caducar} lotes"
    nota.value = (f"En rojo: {lotes} con menos de 1 año para caducar "
                  f"(vencen antes del {fecha_para_mostrar(limite)}, "
                  f"según la fecha en que se exportó este archivo).")
    nota.font = CURSIVAS


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
            caducidad_para_excel(lote["fecha_caducidad"]),
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
                # Los colores vienen como "#dff3e4"; Excel los quiere sin el "#"
                celda.fill = PatternFill("solid", fgColor=color.lstrip("#").upper())

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
