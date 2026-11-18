"""Reporte PDF imprimible de un barrido mensual, separado por almacén."""
from datetime import datetime
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (CondPageBreak, KeepTogether, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

from inventario import ALMACENES, fecha_para_mostrar, normalizar_mes, obtener_partidas

MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre")

# Ancho de cada columna: Código, Descripción, Lote, Caducidad, Existencia (suman 18 cm)
ANCHOS = [2.6 * cm, 7.4 * cm, 3 * cm, 2.6 * cm, 2.4 * cm]

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
    tabla.setStyle(TableStyle([
        # (comando, (columna, fila) inicial, (columna, fila) final, valores...)
        # Las posiciones negativas cuentan desde el final: -1 = última, -2 = penúltima
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),   # Encabezado gris
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -2), 0.5, colors.grey),                  # Cuadrícula sin el subtotal
        ("ROWBACKGROUNDS", (0, 1), (-1, -2),
         [colors.white, colors.HexColor("#f2f2f2")]),                  # Filas alternadas (cebra)
        ("ALIGN", (3, 0), (3, -1), "CENTER"),                          # Columna Caducidad
        ("ALIGN", (4, 0), (4, -1), "RIGHT"),                           # Columna Existencia
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),             # Fila de subtotal
        ("LINEABOVE", (0, -1), (-1, -1), 1, colors.black),
    ]))
    return tabla


def dibujar_pie(lienzo, documento):
    """reportlab la llama en cada hoja para escribir el pie de página."""
    lienzo.saveState()
    lienzo.setFont("Helvetica", 8)
    lienzo.setFillColor(colors.grey)
    lienzo.drawString(documento.leftMargin, 1 * cm, f"Halza Medic  |  {documento.title}")
    lienzo.drawRightString(letter[0] - documento.rightMargin, 1 * cm,
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

    # Líneas de firma; KeepTogether evita que queden partidas entre dos hojas
    firmas = Table([["_" * 32, "_" * 32], ["Elaboró", "Revisó"]], colWidths=[9 * cm, 9 * cm])
    firmas.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    contenido.append(KeepTogether([Spacer(1, 2 * cm), firmas]))

    documento.build(contenido, onFirstPage=dibujar_pie, onLaterPages=dibujar_pie)
