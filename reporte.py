"""Reportes PDF imprimibles: el barrido de un mes y la comparación contra el mes anterior.

También tiene los formatos que comparten la pantalla, el Excel y el PDF de la comparación.
"""
from datetime import datetime
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (CondPageBreak, KeepTogether, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

from inventario import (ALMACENES, comparar_meses, fecha_para_mostrar, mes_anterior,
                        normalizar_mes, obtener_partidas)

MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre")

# Estados de la comparación, en el orden del resumen, y el color de fondo de
# su fila. "Sin cambio" no lleva color para que resalte lo que sí cambió.
ESTADOS = ("Nuevo", "Ya no está", "Cambió", "Sin cambio")
COLORES_ESTADO = {
    "Nuevo": "#dff3e4",       # verde claro
    "Ya no está": "#fde0e0",  # rojo claro
    "Cambió": "#fff1c7",      # amarillo claro
}

# Ancho de cada columna: Código, Descripción, Lote, Caducidad, Existencia (suman 18 cm)
ANCHOS = [2.6 * cm, 7.4 * cm, 3 * cm, 2.6 * cm, 2.4 * cm]

# Comparación (hoja horizontal): Estado, Código, Descripción, Lote, Caducidad,
# Anterior, Actual, Diferencia (suman 24.7 cm)
ANCHOS_COMPARACION = [2.3 * cm, 2.3 * cm, 8.5 * cm, 3 * cm, 2.4 * cm, 2 * cm, 2 * cm, 2.2 * cm]

# Formato que comparten todas las tablas del PDF. Cada comando es:
# (qué, (columna, fila) inicial, (columna, fila) final, valores...)
# Las posiciones negativas cuentan desde el final: -1 = última, -2 = penúltima
FORMATO_TABLA = [
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),   # Encabezado gris
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, -1), 9),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -2), 0.5, colors.grey),                  # Cuadrícula sin el subtotal
    ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),             # Fila de subtotal
    ("LINEABOVE", (0, -1), (-1, -1), 1, colors.black),
]

# Estilos de texto: se parte de los estilos de ejemplo de reportlab y se ajustan
_base = getSampleStyleSheet()
ESTILO_TITULO = _base["Title"]
ESTILO_SUBTITULO = ParagraphStyle("Subtitulo", parent=_base["Normal"],
                                  fontSize=11, leading=15, alignment=TA_CENTER)
ESTILO_ALMACEN = _base["Heading2"]
ESTILO_CELDA = ParagraphStyle("Celda", parent=_base["Normal"], fontSize=9, leading=11)
ESTILO_NORMAL = _base["Normal"]


def nombre_mes(mes):
    """Convierte '2026-11' en 'noviembre 2026'."""
    anio, numero = mes.split("-")
    return f"{MESES[int(numero) - 1]} {anio}"


def texto_partidas(cantidad):
    """1 -> '1 partida', 5 -> '5 partidas'."""
    return f"{cantidad} partida" if cantidad == 1 else f"{cantidad} partidas"


def con_signo(numero):
    """Para mostrar diferencias: 5 -> '+5', -3 -> '-3', 0 -> '0'."""
    return f"{numero:+d}" if numero else "0"


def cantidad_o_guion(cantidad):
    """None (el lote no se registró ese mes) -> '—'. Un número se deja igual."""
    return "—" if cantidad is None else cantidad


def resumen_comparacion(lotes, ocultar_sin_cambio=False):
    """Texto con cuántos lotes hay de cada estado y cómo cambió la existencia total.

    Lo usan la pestaña Comparar y el PDF. Se cuentan TODOS los lotes, también
    los ocultos, para que el total cuadre. Devuelve dos líneas.
    """
    partes = []
    for estado in ESTADOS:
        cantidad = sum(1 for lote in lotes if lote["estado"] == estado)
        texto = f"{estado}: {cantidad}"
        if ocultar_sin_cambio and estado == "Sin cambio":
            texto += " (ocultos)"  # Avisa que existen aunque no se vean
        partes.append(texto)
    total_antes = sum(lote["existencia_anterior"] or 0 for lote in lotes)
    total_ahora = sum(lote["existencia_actual"] or 0 for lote in lotes)
    return ("  |  ".join(partes),
            f"Existencia total: {total_antes} → {total_ahora} "
            f"({con_signo(total_ahora - total_antes)})")


def tabla_almacen(partidas):
    """Arma la tabla de un almacén: encabezado, una fila por lote y subtotal."""
    filas = [["Código", "Descripción", "Lote", "Caducidad", "Existencia"]]
    for p in partidas:
        filas.append([
            p["prefijo"] + p["codigo"],
            Paragraph(escape(p["descripcion"]), ESTILO_CELDA),  # Paragraph = texto que se ajusta en varias líneas
            p["lote"],
            fecha_para_mostrar(p["fecha_caducidad"]),
            p["existencia"],
        ])
    subtotal = sum(p["existencia"] for p in partidas)
    filas.append(["", texto_partidas(len(partidas)), "", "Subtotal:", subtotal])

    tabla = Table(filas, colWidths=ANCHOS, repeatRows=1)  # repeatRows: encabezado en cada hoja
    tabla.setStyle(TableStyle(FORMATO_TABLA + [
        ("ROWBACKGROUNDS", (0, 1), (-1, -2),
         [colors.white, colors.HexColor("#f2f2f2")]),                  # Filas alternadas (cebra)
        ("ALIGN", (3, 0), (3, -1), "CENTER"),                          # Columna Caducidad
        ("ALIGN", (4, 0), (4, -1), "RIGHT"),                           # Columna Existencia
    ]))
    return tabla


def tabla_comparacion(lotes):
    """Arma la tabla de un almacén en la comparación: una fila por lote, del color de su estado."""
    filas = [["Estado", "Código", "Descripción", "Lote", "Caducidad",
              "Anterior", "Actual", "Diferencia"]]
    for lote in lotes:
        filas.append([
            lote["estado"],
            lote["prefijo"] + lote["codigo"],
            Paragraph(escape(lote["descripcion"]), ESTILO_CELDA),
            lote["lote"],
            fecha_para_mostrar(lote["fecha_caducidad"]),
            cantidad_o_guion(lote["existencia_anterior"]),
            cantidad_o_guion(lote["existencia_actual"]),
            con_signo(lote["diferencia"]),
        ])
    antes = sum(lote["existencia_anterior"] or 0 for lote in lotes)
    ahora = sum(lote["existencia_actual"] or 0 for lote in lotes)
    lotes_texto = "1 lote" if len(lotes) == 1 else f"{len(lotes)} lotes"
    filas.append(["", "", lotes_texto, "", "Subtotal:", antes, ahora, con_signo(ahora - antes)])

    formato = FORMATO_TABLA + [
        ("ALIGN", (4, 0), (4, -1), "CENTER"),   # Caducidad
        ("ALIGN", (5, 0), (-1, -1), "RIGHT"),   # Anterior, Actual y Diferencia
    ]
    # Fondo de cada lote según su estado. La fila 0 es el encabezado, por eso
    # el primer lote es la fila 1 (enumerate empieza a contar desde 1).
    for numero, lote in enumerate(lotes, start=1):
        color = COLORES_ESTADO.get(lote["estado"])  # None si es "Sin cambio"
        if color:
            formato.append(("BACKGROUND", (0, numero), (-1, numero), colors.HexColor(color)))

    tabla = Table(filas, colWidths=ANCHOS_COMPARACION, repeatRows=1)
    tabla.setStyle(TableStyle(formato))
    return tabla


def lineas_de_firma():
    """Espacio y líneas para firmar; KeepTogether evita que se separen entre dos hojas."""
    firmas = Table([["_" * 32, "_" * 32], ["Elaboró", "Revisó"]], colWidths=[9 * cm, 9 * cm])
    firmas.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    return KeepTogether([Spacer(1, 2 * cm), firmas])


def dibujar_pie(lienzo, documento):
    """reportlab la llama en cada hoja para escribir el pie de página."""
    lienzo.saveState()
    lienzo.setFont("Helvetica", 8)
    lienzo.setFillColor(colors.grey)
    lienzo.drawString(documento.leftMargin, 1 * cm, f"Halza Medic  |  {documento.title}")
    # pagesize[0] es el ancho de la hoja: sirve igual para hoja vertical u horizontal
    lienzo.drawRightString(documento.pagesize[0] - documento.rightMargin, 1 * cm,
                           f"Página {documento.page}")
    lienzo.restoreState()


def generar_reporte_pdf(mes, ruta_pdf):
    """Crea en ruta_pdf el reporte del barrido del mes indicado (AAAA-MM)."""
    mes = normalizar_mes(mes)
    partidas = obtener_partidas(mes)

    documento = SimpleDocTemplate(
        str(ruta_pdf), pagesize=letter,
        title=f"Barrido de inventario {nombre_mes(mes)}", author="Halza Medic",
        leftMargin=1.5 * cm, rightMargin=1.5 * cm, topMargin=1.5 * cm, bottomMargin=2 * cm,
    )

    # "contenido" es la lista de piezas (flowables) que reportlab acomoda hoja por hoja
    contenido = [
        Paragraph("Halza Medic", ESTILO_TITULO),
        Paragraph(f"Reporte de barrido de inventario: <b>{nombre_mes(mes).capitalize()}</b>",
                  ESTILO_SUBTITULO),
        Paragraph(f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}", ESTILO_SUBTITULO),
        Spacer(1, 0.6 * cm),
    ]

    for clave, nombre in ALMACENES.items():
        partidas_almacen = [p for p in partidas if p["almacen"] == clave]
        if not partidas_almacen:
            continue  # Este almacén no tiene partidas en el mes: no se imprime
        contenido.append(CondPageBreak(4 * cm))  # Si quedan menos de 4 cm, hoja nueva
        contenido.append(Paragraph(nombre, ESTILO_ALMACEN))
        contenido.append(tabla_almacen(partidas_almacen))
        contenido.append(Spacer(1, 0.5 * cm))

    total = sum(p["existencia"] for p in partidas)
    contenido.append(Paragraph(
        f"<b>Total general:</b> {texto_partidas(len(partidas))}  |  Existencia total: {total}",
        ESTILO_NORMAL))

    contenido.append(lineas_de_firma())
    documento.build(contenido, onFirstPage=dibujar_pie, onLaterPages=dibujar_pie)


def generar_comparacion_pdf(mes, ruta_pdf, ocultar_sin_cambio=False):
    """Crea en ruta_pdf la comparación lote por lote del mes (AAAA-MM) contra el anterior.

    ocultar_sin_cambio: True = no imprime los lotes "Sin cambio" (el resumen sí los cuenta).
    """
    actual = normalizar_mes(mes)
    anterior = mes_anterior(actual)
    lotes = comparar_meses(actual, anterior)

    documento = SimpleDocTemplate(
        str(ruta_pdf), pagesize=landscape(letter),  # Hoja horizontal: son 8 columnas
        title=f"Comparación {nombre_mes(actual)} contra {nombre_mes(anterior)}",
        author="Halza Medic",
        leftMargin=1.5 * cm, rightMargin=1.5 * cm, topMargin=1.5 * cm, bottomMargin=2 * cm,
    )

    linea_estados, linea_existencia = resumen_comparacion(lotes, ocultar_sin_cambio)
    contenido = [
        Paragraph("Halza Medic", ESTILO_TITULO),
        Paragraph(f"Comparación de barridos: <b>{nombre_mes(actual).capitalize()}</b> "
                  f"contra <b>{nombre_mes(anterior).capitalize()}</b>", ESTILO_SUBTITULO),
        Paragraph(f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}", ESTILO_SUBTITULO),
        Spacer(1, 0.4 * cm),
        Paragraph(f"{linea_estados}<br/>{linea_existencia}", ESTILO_SUBTITULO),
        Spacer(1, 0.4 * cm),
    ]

    if ocultar_sin_cambio:
        lotes = [lote for lote in lotes if lote["estado"] != "Sin cambio"]

    for clave, nombre in ALMACENES.items():
        lotes_almacen = [lote for lote in lotes if lote["almacen"] == clave]
        if not lotes_almacen:
            continue  # Nada que mostrar de este almacén
        contenido.append(CondPageBreak(4 * cm))  # Si quedan menos de 4 cm, hoja nueva
        contenido.append(Paragraph(nombre, ESTILO_ALMACEN))
        contenido.append(tabla_comparacion(lotes_almacen))
        contenido.append(Spacer(1, 0.5 * cm))

    if not lotes:
        contenido.append(Paragraph("Todos los lotes quedaron sin cambio.", ESTILO_SUBTITULO))

    contenido.append(lineas_de_firma())
    documento.build(contenido, onFirstPage=dibujar_pie, onLaterPages=dibujar_pie)
