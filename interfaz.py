import os
import sqlite3
import traceback
from datetime import datetime
from itertools import groupby
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from inventario import (ALMACENES, CARPETA_DATOS, DB_PATH, FECHA_NO_CADUCA, NO_CADUCA,
                        actualizar_partida, comparar_meses,
                        crear_tabla, eliminar_partida, fecha_para_mostrar,
                        insertar_partida, limite_de_caducidad, mes_anterior,
                        normalizar_mes, obtener_partida,
                        obtener_partidas, obtener_producto, respaldar, respaldo_automatico)
from exportar_excel import generar_excel
from formato_fechas import activar_formato_fecha, activar_formato_mes
from reporte import generar_reporte_pdf, nombre_mes

# Diccionario inverso: "Almacén Chico (Medicamentos)" -> "1"
CLAVE_POR_NOMBRE = {nombre: clave for clave, nombre in ALMACENES.items()}

# Colores en pares: (modo claro, modo oscuro)
COLOR_OK = ("#1b7f3b", "#5fd38d")
COLOR_ERROR = ("#c62828", "#ff6b6b")
COLOR_TEXTO = ("gray10", "gray90")
COLOR_AVISO = ("#b45309", "#fbbf24")

# Columnas de la tabla de consulta: (nombre interno, encabezado, ancho, alineación)
COLUMNAS = (
    ("mes", "Mes", 80, "center"),
    ("codigo", "Código", 110, "w"),
    ("descripcion", "Descripción", 300, "w"),
    ("lote", "Lote", 120, "w"),
    ("caducidad", "Caducidad", 100, "center"),
    ("existencia", "Existencia", 90, "e"),
)

# Columnas de la tabla de la pestaña Comparar
COLUMNAS_COMPARAR = (
    ("estado", "Estado", 95, "w"),
    ("codigo", "Código", 95, "w"),
    ("descripcion", "Descripción", 205, "w"),
    ("lote", "Lote", 105, "w"),
    ("caducidad", "Caducidad", 90, "center"),
    ("anterior", "Anterior", 65, "e"),
    ("actual", "Actual", 65, "e"),
    ("diferencia", "Diferencia", 80, "e"),
)

# Estados de la comparación, en el orden del resumen, y el color de fondo de
# su fila. "Sin cambio" no lleva color para que resalte lo que sí cambió.
ESTADOS = ("Nuevo", "Ya no está", "Cambió", "Sin cambio")
COLORES_ESTADO = {
    "Nuevo": "#dff3e4",       # verde claro
    "Ya no está": "#fde0e0",  # rojo claro
    "Cambió": "#fff1c7",      # amarillo claro
}

# Consulta: fondo de las partidas con menos de un año para caducar (o ya caducadas)
COLOR_POR_CADUCAR = "#f8caca"  # rojo claro

# Botones del filtro de almacén en Consulta: texto del botón -> clave del almacén
# (None = sin filtro, se muestran los dos)
FILTRO_ALMACEN = {
    "Todos": None,
    "1 · Chico": "1",
    "2 · Grande": "2",
}


# ------------------------------------------------------------------- Estado

# None = capturando partidas nuevas; un número = id de la partida que se edita
id_en_edicion = None
# (mes, almacén) que había en el formulario antes de editar, para restaurarlos
respaldo_formulario = None
# Último código que ya se buscó, para no buscarlo de nuevo con cada tecla
codigo_revisado = ""


# ---------------------------------------------------------------- Funciones

def crear_etiqueta(fila, texto):
    """Coloca un texto fijo en la columna izquierda del formulario."""
    etiqueta = ctk.CTkLabel(formulario, text=texto)
    etiqueta.grid(row=fila, column=0, sticky="w", padx=(0, 15), pady=6)


def crear_campo(fila, texto, ejemplo=""):
    """Coloca una etiqueta y una caja de texto en la fila indicada y devuelve la caja."""
    crear_etiqueta(fila, texto)
    entrada = ctk.CTkEntry(formulario, width=300, placeholder_text=ejemplo)
    entrada.grid(row=fila, column=1, pady=6)
    return entrada


def crear_tabla_visual(padre, columnas, con_arbol=False):
    """Crea una tabla (ttk.Treeview) con su barra de desplazamiento y la devuelve.

    columnas: tuplas (nombre interno, encabezado, ancho, alineación).
    con_arbol: True = filas que se despliegan (un producto y debajo sus lotes).
    """
    marco = ctk.CTkFrame(padre, fg_color="transparent")
    marco.pack(fill="both", expand=True)

    # "headings" muestra solo las columnas; "tree headings" agrega a la izquierda
    # la columna del árbol (#0), que es donde aparece la flechita para desplegar
    mostrar = "tree headings" if con_arbol else "headings"
    nueva_tabla = ttk.Treeview(marco, columns=[c[0] for c in columnas], show=mostrar,
                               selectmode="browse")  # "browse" = solo una fila a la vez
    if con_arbol:
        nueva_tabla.column("#0", width=30, minwidth=30, stretch=False)  # Solo para la flechita
    for nombre, encabezado, ancho, alineacion in columnas:
        nueva_tabla.heading(nombre, text=encabezado)
        nueva_tabla.column(nombre, width=ancho, anchor=alineacion)
    nueva_tabla.pack(side="left", fill="both", expand=True)

    barra = ctk.CTkScrollbar(marco, command=nueva_tabla.yview)
    barra.pack(side="right", fill="y")
    nueva_tabla.configure(yscrollcommand=barra.set)
    return nueva_tabla


def mostrar_mensaje(texto, color):
    """Escribe un aviso en la parte de abajo de la pestaña Captura."""
    etiqueta_mensaje.configure(text=texto, text_color=color)


def guardar():
    """Guarda lo que hay en el formulario.

    Si no se está editando, crea una partida nueva (INSERT).
    Si se está editando, reemplaza los datos de esa partida (UPDATE); si se
    corrigió el nombre o almacén del producto, antes pregunta si se corrige
    en todas sus partidas.
    """
    datos = dict(
        mes=entrada_mes.get(),
        almacen=CLAVE_POR_NOMBRE[selector_almacen.get()],
        codigo=entrada_codigo.get(),
        descripcion=entrada_descripcion.get(),
        existencia=entrada_existencia.get(),
        lote=entrada_lote.get(),
        fecha_caducidad=entrada_caducidad.get(),
    )
    try:
        if id_en_edicion is None:
            insertar_partida(**datos)
        else:
            # confirmar recibe una FUNCIÓN: actualizar_partida la llama solo si
            # necesita preguntar, y esta muestra la ventana de Sí/No
            guardado = actualizar_partida(
                id_en_edicion, **datos,
                confirmar=lambda mensaje: messagebox.askyesno("Corregir producto", mensaje),
            )
            if not guardado:
                mostrar_mensaje("No se guardó nada: el producto quedó como estaba.",
                                COLOR_AVISO)
                return
    except ValueError as error:
        mostrar_mensaje(str(error), COLOR_ERROR)
        return
    except sqlite3.IntegrityError:
        mostrar_mensaje("Ese lote ya está registrado para este producto, mes y almacén.",
                        COLOR_ERROR)
        return

    resumen = f"{datos['codigo'].strip().upper()} lote {datos['lote'].strip().upper()}"

    if id_en_edicion is None:
        mostrar_mensaje(f"Guardado: {resumen}", COLOR_OK)
        # Se conservan mes, almacén, código y descripción para capturar
        # rápido el siguiente lote del mismo producto.
        for entrada in (entrada_existencia, entrada_lote):
            entrada.delete(0, "end")
        if not casilla_no_caduca.get():  # "No caduca" se queda para el siguiente lote
            entrada_caducidad.delete(0, "end")
        entrada_existencia.focus()
    else:
        nuevo_producto()  # Sale del modo edición y deja el formulario limpio
        mostrar_mensaje(f"Cambios guardados: {resumen}", COLOR_OK)

    cargar_tabla()  # Para que la pestaña Consulta muestre los datos al día


def nuevo_producto():
    """Limpia los datos del producto y pone el cursor en Código.

    Si se estaba editando una partida, también cancela la edición.
    """
    global codigo_revisado

    salir_de_edicion()
    codigo_revisado = ""
    for entrada in (entrada_codigo, entrada_descripcion, entrada_existencia, entrada_lote):
        entrada.delete(0, "end")
    poner_no_caduca(False)  # Desmarca la casilla y deja la caducidad vacía
    etiqueta_mensaje.configure(text="")  # Borra el aviso anterior
    entrada_codigo.focus()


def poner_no_caduca(activa):
    """Marca o desmarca la casilla "No caduca" y ajusta la caja de Caducidad.

    Marcada: la caja dice "No caduca" y se bloquea (no se puede escribir una fecha).
    Desmarcada: la caja queda vacía y lista para escribir una fecha.
    """
    if activa:
        casilla_no_caduca.select()
    else:
        casilla_no_caduca.deselect()

    # Una caja bloqueada ("disabled") no deja cambiar su texto, ni siquiera desde
    # el código; por eso primero se desbloquea, se cambia el texto y al final se bloquea
    entrada_caducidad.configure(state="normal")
    entrada_caducidad.delete(0, "end")
    if activa:
        entrada_caducidad.insert(0, NO_CADUCA)
        entrada_caducidad.configure(state="disabled")


def al_escribir_codigo(evento=None):
    """Si el código escrito ya está registrado, llena solos su descripción y almacén.

    Así cada código conserva siempre el mismo producto (un código = un producto).
    """
    global codigo_revisado

    codigo = entrada_codigo.get().strip().upper()
    if codigo == codigo_revisado:
        return  # El código no cambió (ej. se presionó una flecha o Shift)
    codigo_revisado = codigo

    # Al editar no se cuenta la propia partida: solo importa cómo están las demás
    producto = obtener_producto(codigo, sin_id=id_en_edicion)
    if producto:
        entrada_descripcion.delete(0, "end")
        entrada_descripcion.insert(0, producto["descripcion"])
        selector_almacen.set(ALMACENES[producto["almacen"]])
        # La casilla se cambia solo si hace falta, para no borrar una fecha ya escrita
        if producto["no_caduca"] != bool(casilla_no_caduca.get()):
            poner_no_caduca(producto["no_caduca"])

    if id_en_edicion is not None:
        return  # Al editar no se borra nada ni se cambia el aviso "Editando..."

    if producto:
        mostrar_mensaje(f"{codigo} ya está registrado: se llenaron su descripción y almacén.",
                        COLOR_TEXTO)
    else:
        # Código nuevo: se borra el nombre que quedó del producto anterior,
        # para no guardar por error un producto nuevo con el nombre de otro
        entrada_descripcion.delete(0, "end")
        # Y empieza pidiendo su fecha: así no se queda marcado "No caduca" por
        # olvido cuando después de una silla se captura un medicamento
        if casilla_no_caduca.get():
            poner_no_caduca(False)
        mostrar_mensaje("", COLOR_TEXTO)


def agregar_lote(partida, padre, en_rojo):
    """Agrega una partida (un lote) a la tabla de Consulta.

    padre: "" = va como fila principal (producto de un solo lote);
           el iid de un producto = va debajo de él, al desplegarlo.
    El iid de la fila es el id de la partida: así Editar y Eliminar saben cuál es.
    """
    if padre:
        inicio = ("", "", "")  # Mes, código y descripción ya están en la fila del producto
    else:
        inicio = (partida["mes"], partida["prefijo"] + partida["codigo"], partida["descripcion"])

    tabla.insert(padre, "end", iid=partida["id"],
                 tags=("por_caducar",) if en_rojo else (),  # Esta etiqueta pinta la fila de rojo
                 values=inicio + (partida["lote"],
                                  fecha_para_mostrar(partida["fecha_caducidad"]),
                                  partida["existencia"]))


def cargar_tabla():
    """Vacía la tabla y la vuelve a llenar según los filtros: mes, producto y almacén.

    Cada producto ocupa una fila. Si tiene varios lotes, su fila muestra el total
    y se despliega (flechita a la izquierda) para ver cada lote con su caducidad.
    """
    try:
        partidas = obtener_partidas(
            entrada_filtro_mes.get(),
            almacen=FILTRO_ALMACEN[selector_filtro_almacen.get()],
            texto=entrada_buscar.get(),
        )
    except ValueError as error:
        etiqueta_resumen.configure(text=str(error), text_color=COLOR_ERROR)
        return

    # Recordar qué productos estaban desplegados, para dejarlos igual al recargar
    abiertos = [fila for fila in tabla.get_children() if tabla.item(fila, "open")]
    tabla.delete(*tabla.get_children())  # Borra todas las filas actuales

    # Las fechas AAAA-MM-DD se pueden comparar como texto: "2027-02-28" < "2027-11-19".
    # Los productos que no caducan (9999-12-31) nunca quedan antes del límite.
    limite = limite_de_caducidad()
    por_caducar = 0
    productos = 0

    # groupby junta las filas SEGUIDAS que tienen la misma clave (mes y código).
    # Funciona porque obtener_partidas ya las ordena por mes, almacén y código:
    # todos los lotes de un mismo producto vienen juntos.
    for (mes, codigo), grupo in groupby(partidas, key=lambda p: (p["mes"], p["codigo"])):
        lotes = list(grupo)
        productos += 1
        en_rojo = [lote["fecha_caducidad"] < limite for lote in lotes]  # True/False por lote
        por_caducar += sum(en_rojo)  # sum() cuenta cada True como 1

        if len(lotes) == 1:
            agregar_lote(lotes[0], "", en_rojo[0])  # Un solo lote: fila normal
            continue

        # Varios lotes: primero la fila del producto con el total...
        # Su iid es texto ("2026-11|P00089") para no confundirlo con el id de una partida
        producto = f"{mes}|{codigo}"
        primero = lotes[0]
        tabla.insert("", "end", iid=producto, open=producto in abiertos,
                     tags=("por_caducar",) if any(en_rojo) else (),  # Rojo si algún lote lo está
                     values=(mes, primero["prefijo"] + codigo, primero["descripcion"],
                             f"{len(lotes)} lotes", "",
                             sum(lote["existencia"] for lote in lotes)))
        # ...y debajo de ella, cada lote (zip recorre las dos listas al mismo tiempo)
        for lote, rojo in zip(lotes, en_rojo):
            agregar_lote(lote, producto, rojo)

    total = sum(fila["existencia"] for fila in partidas)
    nombre = "producto" if productos == 1 else "productos"
    etiqueta_resumen.configure(
        text=f"{productos} {nombre} ({len(partidas)} lotes)  |  Existencia: {total}  |  "
             f"En rojo: {por_caducar} vencen en menos de 1 año",
        text_color=COLOR_TEXTO)

def ver_todos():
    """Vacía el filtro de mes y muestra las partidas de todos los meses."""
    entrada_filtro_mes.delete(0, "end")
    cargar_tabla()


def limpiar_busqueda():
    """Quita la búsqueda de producto y el filtro de almacén (el mes se queda igual)."""
    entrada_buscar.delete(0, "end")
    selector_filtro_almacen.set("Todos")
    cargar_tabla()
    entrada_buscar.focus()


def id_seleccionado():
    """Devuelve el id de la partida seleccionada en la tabla, o None si no hay."""
    seleccion = tabla.selection()
    if not seleccion:
        messagebox.showwarning("Sin selección", "Primero selecciona una partida en la tabla.")
        return None
    fila = seleccion[0]
    if tabla.get_children(fila):  # Tiene filas debajo: es un producto, no un lote
        messagebox.showwarning(
            "Producto con varios lotes",
            "Este producto tiene varios lotes. Despliégalo (flechita de la izquierda "
            "o doble clic) y selecciona el lote que quieres editar o eliminar.")
        return None
    return int(fila)


def al_doble_clic(evento):
    """Doble clic: en un producto con varios lotes lo despliega o pliega; en un lote, lo edita."""
    fila = tabla.identify_row(evento.y)  # La fila que está bajo el mouse ("" = ninguna)
    if not fila or tabla.get_children(fila):
        # Sin fila (ej. el encabezado) o un producto con lotes: no se edita nada.
        # El Treeview mismo despliega o pliega el producto con el doble clic.
        return
    editar_seleccion()


def editar_seleccion():
    """Carga la partida seleccionada en el formulario de Captura para modificarla."""
    global id_en_edicion, respaldo_formulario, codigo_revisado

    id_partida = id_seleccionado()
    if id_partida is None:
        return
    partida = obtener_partida(id_partida)

    # Recordar mes y almacén, solo si no se estaba editando ya otra partida
    if id_en_edicion is None:
        respaldo_formulario = (entrada_mes.get(), selector_almacen.get())
    id_en_edicion = id_partida

    # Pasar los datos guardados al formulario
    valores = (
        (entrada_mes, partida["mes"]),
        (entrada_codigo, partida["codigo"]),
        (entrada_descripcion, partida["descripcion"]),
        (entrada_existencia, str(partida["existencia"])),
        (entrada_lote, partida["lote"]),
    )
    for entrada, valor in valores:
        entrada.delete(0, "end")
        entrada.insert(0, valor)
    selector_almacen.set(ALMACENES[partida["almacen"]])

    # Caducidad: o la casilla "No caduca" marcada, o su fecha en la caja
    no_caduca = partida["fecha_caducidad"] == FECHA_NO_CADUCA
    poner_no_caduca(no_caduca)
    if not no_caduca:
        entrada_caducidad.insert(0, fecha_para_mostrar(partida["fecha_caducidad"]))
    codigo_revisado = partida["codigo"]  # Ya está en el formulario: no hay que buscarlo

    # Mostrar claramente que estamos en modo edición
    boton_guardar.configure(text="Guardar cambios")
    boton_nuevo.configure(text="Cancelar edición")
    mostrar_mensaje(f"Editando {partida['prefijo']}{partida['codigo']} "
                    f"lote {partida['lote']}", COLOR_AVISO)
    pestanas.set("Captura")
    entrada_existencia.focus()


def salir_de_edicion():
    """Regresa el formulario al modo normal y restaura el mes y almacén de antes."""
    global id_en_edicion

    if id_en_edicion is None:
        return  # No se estaba editando: no hay nada que restaurar
    id_en_edicion = None

    mes, almacen = respaldo_formulario
    entrada_mes.delete(0, "end")
    entrada_mes.insert(0, mes)
    selector_almacen.set(almacen)

    boton_guardar.configure(text="Guardar partida")
    boton_nuevo.configure(text="Nuevo producto")


def eliminar_seleccion():
    """Pide confirmación y borra la partida seleccionada en la tabla."""
    id_partida = id_seleccionado()
    if id_partida is None:
        return
    partida = obtener_partida(id_partida)

    confirmado = messagebox.askyesno(
        "Eliminar partida",
        "¿Seguro que quieres eliminar esta partida?\n\n"
        f"{partida['prefijo']}{partida['codigo']}  {partida['descripcion']}\n"
        f"Lote {partida['lote']}  |  Existencia {partida['existencia']}  |  Mes {partida['mes']}",
        icon="warning",
    )
    if not confirmado:
        return

    eliminar_partida(id_partida)
    if id_partida == id_en_edicion:
        nuevo_producto()  # La partida que se estaba editando ya no existe
    cargar_tabla()


def guardar_archivo(nombre, extension, tipo, crear_archivo, abrir=True):
    """Pregunta dónde guardar y crea el archivo. Lo usan el PDF, el Excel y el respaldo.

    nombre: nombre sugerido sin extensión (ej. "Barrido_2026-11")
    extension: ".pdf", ".xlsx" o ".db" | tipo: texto para la ventana (ej. "Archivo PDF")
    crear_archivo: función que recibe la ruta elegida y crea ahí el archivo
    abrir: si es True, al terminar se abre el archivo
    Devuelve la ruta del archivo creado, o None si se canceló o hubo un error.
    """
    ruta = filedialog.asksaveasfilename(
        title="Guardar como",
        initialfile=nombre + extension,
        defaultextension=extension,
        filetypes=[(tipo, "*" + extension)],
    )
    if not ruta:
        return None  # El usuario presionó Cancelar

    try:
        crear_archivo(ruta)
    except PermissionError:
        messagebox.showerror("No se pudo guardar",
                             "Ese archivo está abierto en otro programa. Ciérralo e intenta de nuevo.")
        return None
    except ValueError as error:
        messagebox.showerror("No se pudo guardar", str(error))  # Mensaje nuestro, en español
        return None
    except (OSError, sqlite3.Error) as error:
        # Problemas de Windows o de SQLite: USB desconectada, disco lleno, sin permiso...
        messagebox.showerror("No se pudo guardar",
                             "No se pudo escribir en esa ubicación. Revisa que la carpeta "
                             f"o la USB sigan disponibles.\n\nDetalle técnico: {error}")
        return None

    if abrir:
        os.startfile(ruta)  # Abre el archivo con el programa predeterminado de Windows
    return ruta


def generar_pdf():
    """Pide dónde guardar el reporte del mes del filtro, lo genera y lo abre."""
    try:
        mes = normalizar_mes(entrada_filtro_mes.get())
    except ValueError:
        messagebox.showwarning("Falta el mes",
                               "Escribe en el filtro el mes del reporte (ej. 2026-11).")
        return

    if not obtener_partidas(mes):
        messagebox.showinfo("Sin datos", f"No hay partidas registradas en {mes}.")
        return

    guardar_archivo(f"Barrido_{mes}", ".pdf", "Archivo PDF",
                    lambda ruta: generar_reporte_pdf(mes, ruta))


def exportar_excel():
    """Exporta a Excel el mes del filtro, o todos los meses si el filtro está vacío."""
    texto_mes = entrada_filtro_mes.get().strip()
    try:
        mes = normalizar_mes(texto_mes) if texto_mes else ""
    except ValueError as error:
        messagebox.showwarning("Mes inválido", str(error))
        return

    if not obtener_partidas(mes):
        messagebox.showinfo("Sin datos", f"No hay partidas registradas en {mes}." if mes
                            else "No hay partidas registradas todavía.")
        return

    nombre = f"Barrido_{mes}" if mes else "Barridos_todos_los_meses"
    guardar_archivo(nombre, ".xlsx", "Libro de Excel",
                    lambda ruta: generar_excel(mes, ruta))


def respaldar_base():
    """Guarda una copia de toda la base de datos donde el usuario elija."""
    # Fecha y hora en el nombre, ej. respaldo_inventario_2026-11-18_1530.db
    nombre = f"respaldo_inventario_{datetime.now():%Y-%m-%d_%H%M}"
    ruta = guardar_archivo(nombre, ".db", "Base de datos", respaldar, abrir=False)
    if ruta:
        messagebox.showinfo(
            "Respaldo listo",
            f"Se guardó una copia completa de la base de datos en:\n\n{ruta}\n\n"
            "Si la guardaste en esta misma computadora, copia también el respaldo "
            "a una USB o a la nube.")


def error_inesperado(tipo, error, rastreo):
    """Atrapa cualquier error no previsto al usar un botón, una tecla, etc.

    En el .exe no hay consola donde ver los errores: sin esto, el botón
    simplemente "no haría nada". Se avisa en pantalla y el detalle técnico
    se anota en errores.log, en la carpeta de los datos.
    """
    detalle = "".join(traceback.format_exception(tipo, error, rastreo))
    try:
        with open(CARPETA_DATOS / "errores.log", "a", encoding="utf-8") as archivo:
            archivo.write(f"\n===== {datetime.now():%d/%m/%Y %H:%M:%S} =====\n{detalle}")
    except OSError:
        pass  # Si ni el registro se puede escribir, al menos se muestra el aviso
    messagebox.showerror(
        "Error inesperado",
        f"Ocurrió un error no previsto:\n\n{error}\n\n"
        "El detalle quedó anotado en el archivo errores.log, en la carpeta del programa.")


def hacer_respaldo_automatico():
    """Hace el respaldo automático del día y muestra abajo cuándo fue el último.

    Si falla, solo avisa: el programa debe abrir aunque no se pueda respaldar.
    """
    try:
        ultimo = respaldo_automatico()
    except (OSError, sqlite3.Error) as error:
        etiqueta_respaldo.configure(
            text=f"⚠ No se pudo hacer el respaldo automático de hoy: {error}",
            text_color=COLOR_AVISO)
        return

    if ultimo is None:
        texto = "Respaldo automático: todavía no hay partidas que respaldar"
    else:
        # ultimo.stem es el nombre sin ".db": "auto_2026-11-18" -> se quita "auto_"
        fecha = fecha_para_mostrar(ultimo.stem.removeprefix("auto_"))
        texto = f"Último respaldo automático: {fecha}  (carpeta {ultimo.parent.name})"
    etiqueta_respaldo.configure(text=texto, text_color="gray40")


def con_signo(numero):
    """Para mostrar diferencias: 5 -> '+5', -3 -> '-3', 0 -> '0'."""
    return f"{numero:+d}" if numero else "0"


def cantidad_o_guion(cantidad):
    """None (el lote no se registró ese mes) -> '—'. Un número se deja igual."""
    return "—" if cantidad is None else cantidad


def comparar():
    """Compara el mes escrito contra su mes anterior y llena la tabla de Comparar."""
    tabla_comparar.delete(*tabla_comparar.get_children())
    etiqueta_contra.configure(text="")

    try:
        actual = normalizar_mes(entrada_comparar_mes.get())
    except ValueError as error:
        etiqueta_comparacion.configure(text=str(error), text_color=COLOR_ERROR)
        return
    anterior = mes_anterior(actual)
    etiqueta_contra.configure(text=f"{nombre_mes(actual)}  contra  {nombre_mes(anterior)}")

    # Si un mes está vacío, todo saldría "Nuevo" o "Ya no está", y eso confunde
    for mes in (actual, anterior):
        if not obtener_partidas(mes):
            etiqueta_comparacion.configure(
                text=f"No hay partidas registradas en {nombre_mes(mes)}; "
                     "no hay contra qué comparar.",
                text_color=COLOR_AVISO)
            return

    lotes = comparar_meses(actual, anterior)
    ocultar_sin_cambio = casilla_ocultar.get() == 1  # True si la casilla está marcada

    for lote in lotes:
        if ocultar_sin_cambio and lote["estado"] == "Sin cambio":
            continue  # Se salta este lote: no se agrega a la tabla
        tabla_comparar.insert("", "end", tags=(lote["estado"],), values=(
            lote["estado"],
            lote["prefijo"] + lote["codigo"],
            lote["descripcion"],
            lote["lote"],
            fecha_para_mostrar(lote["fecha_caducidad"]),
            cantidad_o_guion(lote["existencia_anterior"]),
            cantidad_o_guion(lote["existencia_actual"]),
            con_signo(lote["diferencia"]),
        ))

    # Resumen: cuántos lotes hay de cada estado y cómo cambió la existencia total.
    # Se cuentan TODOS los lotes, también los ocultos, para que el total cuadre.
    partes = []
    for estado in ESTADOS:
        cantidad = sum(1 for lote in lotes if lote["estado"] == estado)
        texto = f"{estado}: {cantidad}"
        if ocultar_sin_cambio and estado == "Sin cambio":
            texto += " (ocultos)"  # Avisa que existen aunque no se vean en la tabla
        partes.append(texto)
    total_antes = sum(lote["existencia_anterior"] or 0 for lote in lotes)
    total_ahora = sum(lote["existencia_actual"] or 0 for lote in lotes)

    etiqueta_comparacion.configure(
        text="  |  ".join(partes)
             + f"\nExistencia total: {total_antes} → {total_ahora} "
               f"({con_signo(total_ahora - total_antes)})",
        text_color=COLOR_TEXTO)


def al_cambiar_pestana():
    """Al abrir la pestaña Comparar, la recalcula para que muestre los datos al día."""
    if pestanas.get() == "Comparar":
        comparar()


def al_presionar_enter(evento):
    """Enter hace la acción principal de la pestaña que está abierta."""
    pestana = pestanas.get()
    if pestana == "Captura":
        guardar()
    elif pestana == "Consulta":
        cargar_tabla()
    else:
        comparar()


# ------------------------------------------------------------------ Ventana

crear_tabla()  # Asegura que la tabla exista antes de guardar algo

# La tabla de consulta (ttk) no tiene modo oscuro; en modo claro todo combina
ctk.set_appearance_mode("light")

ventana = ctk.CTk()
ventana.report_callback_exception = error_inesperado  # Errores no previstos -> aviso + log
ventana.title("Inventario Halza Medic")
ventana.geometry("900x580")

# Letra de las tablas (aplica a todas las tablas ttk de la ventana)
estilo = ttk.Style()
estilo.configure("Treeview", rowheight=28, font=("Roboto", 11))
estilo.configure("Treeview.Heading", font=("Roboto", 11, "bold"))

# Barra inferior: dónde está la base de datos y el botón de respaldo.
# Se acomoda ANTES que las pestañas para asegurar su lugar abajo.
barra_inferior = ctk.CTkFrame(ventana, fg_color="transparent")
barra_inferior.pack(side="bottom", fill="x", padx=20, pady=(0, 12))

# El botón va primero: pack reparte el espacio en orden, así que el botón
# siempre cabe, y si la ruta es muy larga, lo que se corta es el texto.
boton_respaldo = ctk.CTkButton(barra_inferior, text="Respaldar base de datos", width=170,
                               command=respaldar_base, fg_color="transparent",
                               border_width=2, text_color=COLOR_TEXTO)
boton_respaldo.pack(side="right")

# Dos líneas de texto, una debajo de la otra, a la izquierda del botón
textos_inferiores = ctk.CTkFrame(barra_inferior, fg_color="transparent")
textos_inferiores.pack(side="left", fill="x", expand=True, padx=(0, 15))

etiqueta_ubicacion = ctk.CTkLabel(textos_inferiores, text=f"Base de datos: {DB_PATH}",
                                  text_color="gray40", font=ctk.CTkFont(size=12),
                                  anchor="w", height=18)  # anchor="w": texto a la izquierda
etiqueta_ubicacion.pack(fill="x")

etiqueta_respaldo = ctk.CTkLabel(textos_inferiores, text="", text_color="gray40",
                                 font=ctk.CTkFont(size=12), anchor="w", height=18)
etiqueta_respaldo.pack(fill="x")

# command: se llama cada vez que el usuario cambia de pestaña con un clic
pestanas = ctk.CTkTabview(ventana, command=al_cambiar_pestana)
pestanas.pack(fill="both", expand=True, padx=20, pady=(5, 8))
pestana_captura = pestanas.add("Captura")
pestana_consulta = pestanas.add("Consulta")
pestana_comparar = pestanas.add("Comparar")

# ----------------------------------------------------- Pestaña Captura

# Contenedor invisible que acomoda etiquetas y campos en filas y columnas
formulario = ctk.CTkFrame(pestana_captura, fg_color="transparent")
formulario.pack(padx=30, pady=10)

entrada_mes = crear_campo(0, "Mes:")
entrada_mes.insert(0, datetime.now().strftime("%Y-%m"))  # Mes actual por defecto

crear_etiqueta(1, "Almacén:")
selector_almacen = ctk.CTkOptionMenu(formulario, values=list(ALMACENES.values()),
                                     width=300)
selector_almacen.grid(row=1, column=1, pady=6)

entrada_codigo = crear_campo(2, "Código:", "Ej. P02836")
entrada_descripcion = crear_campo(3, "Descripción:", "Ej. Paracetamol 30 mg 10/c")
entrada_existencia = crear_campo(4, "Existencia:", "Cantidad de este lote")
entrada_lote = crear_campo(5, "Lote:", "Ej. KH233224")
entrada_caducidad = crear_campo(6, "Caducidad:", "dd/mm/aaaa")

# Para lo que no caduca (sillas, tijeras...): bloquea la fecha y guarda "No caduca".
# .get() de una casilla da 1 si está marcada y 0 si no.
casilla_no_caduca = ctk.CTkCheckBox(
    formulario, text="No caduca (sillas, tijeras, batas…)",
    command=lambda: poner_no_caduca(casilla_no_caduca.get() == 1))
casilla_no_caduca.grid(row=7, column=1, sticky="w", pady=(0, 6))

# Contenedor invisible para poner los dos botones lado a lado
botones = ctk.CTkFrame(pestana_captura, fg_color="transparent")
botones.pack(pady=(10, 5))

boton_guardar = ctk.CTkButton(botones, text="Guardar partida", command=guardar)
boton_guardar.pack(side="left", padx=5)

boton_nuevo = ctk.CTkButton(botones, text="Nuevo producto", command=nuevo_producto,
                            fg_color="transparent", border_width=2,
                            text_color=COLOR_TEXTO)
boton_nuevo.pack(side="left", padx=5)

etiqueta_mensaje = ctk.CTkLabel(pestana_captura, text="", wraplength=420)
etiqueta_mensaje.pack(pady=(0, 10))

# ----------------------------------------------------- Pestaña Consulta

# Fila de filtro: [Mes:] [caja] [Buscar]
filtro = ctk.CTkFrame(pestana_consulta, fg_color="transparent")
filtro.pack(fill="x", pady=(5, 10))

ctk.CTkLabel(filtro, text="Mes:").pack(side="left", padx=(0, 10))
entrada_filtro_mes = ctk.CTkEntry(filtro, width=150, placeholder_text="Vacío = todos")
entrada_filtro_mes.pack(side="left")
entrada_filtro_mes.insert(0, datetime.now().strftime("%Y-%m"))

boton_buscar = ctk.CTkButton(filtro, text="Buscar", width=100, command=cargar_tabla)
boton_buscar.pack(side="left", padx=10)

boton_ver_todos = ctk.CTkButton(filtro, text="Ver todos", width=100, command=ver_todos)
boton_ver_todos.pack(side="left", padx=(0, 10))   # 0 a la izquierda, 10 a la derecha

# Botones de acciones, pegados a la orilla derecha. Con side="right" el primero
# que se acomoda queda hasta la orilla, por eso Eliminar va antes que Editar.
boton_eliminar = ctk.CTkButton(filtro, text="Eliminar", width=100, command=eliminar_seleccion,
                               fg_color="#c62828", hover_color="#8e1c1c")
boton_eliminar.pack(side="right")

boton_editar = ctk.CTkButton(filtro, text="Editar", width=100, command=editar_seleccion)
boton_editar.pack(side="right", padx=10)

# Segunda fila: [Producto:] [caja de búsqueda] [Limpiar]       [Almacén:] [Todos|1|2]
busqueda = ctk.CTkFrame(pestana_consulta, fg_color="transparent")
busqueda.pack(fill="x", pady=(0, 10))

ctk.CTkLabel(busqueda, text="Producto:").pack(side="left", padx=(0, 10))
entrada_buscar = ctk.CTkEntry(busqueda, width=300, placeholder_text="Código, nombre o lote")
entrada_buscar.pack(side="left")
# Busca mientras se escribe: cada vez que se suelta una tecla, se recarga la tabla
entrada_buscar.bind("<KeyRelease>", lambda evento: cargar_tabla())

boton_limpiar = ctk.CTkButton(busqueda, text="Limpiar", width=80, command=limpiar_busqueda,
                              fg_color="transparent", border_width=2, text_color=COLOR_TEXTO)
boton_limpiar.pack(side="left", padx=10)

# Botones pegados: solo uno puede estar activo. command recibe el texto del botón
# elegido, pero cargar_tabla() no lo necesita (lo lee con .get()), por eso el lambda.
selector_filtro_almacen = ctk.CTkSegmentedButton(busqueda, values=list(FILTRO_ALMACEN),
                                                 command=lambda valor: cargar_tabla())
selector_filtro_almacen.set("Todos")
selector_filtro_almacen.pack(side="right")
ctk.CTkLabel(busqueda, text="Almacén:").pack(side="right", padx=(0, 10))


tabla = crear_tabla_visual(pestana_consulta, COLUMNAS, con_arbol=True)
tabla.tag_configure("por_caducar", background=COLOR_POR_CADUCAR)

# Atajos: doble clic = desplegar un producto o editar un lote; tecla Supr = Eliminar
tabla.bind("<Double-1>", al_doble_clic)
tabla.bind("<Delete>", lambda evento: eliminar_seleccion())

# Fila inferior: resumen a la izquierda y botones de Excel y PDF a la derecha.
# Los botones se acomodan ANTES que el resumen: pack reparte el espacio en orden,
# así que si el texto no cabe, lo que se corta es el texto y no los botones.
pie = ctk.CTkFrame(pestana_consulta, fg_color="transparent")
pie.pack(fill="x", pady=(10, 0))

boton_pdf = ctk.CTkButton(pie, text="Generar PDF del mes", width=160, command=generar_pdf)
boton_pdf.pack(side="right")

# Verde de Excel, para reconocerlo de un vistazo. Queda a la izquierda del PDF
boton_excel = ctk.CTkButton(pie, text="Exportar a Excel", width=140, command=exportar_excel,
                            fg_color="#1d6f42", hover_color="#155232")
boton_excel.pack(side="right", padx=10)

etiqueta_resumen = ctk.CTkLabel(pie, text="", anchor="w")
etiqueta_resumen.pack(side="left", fill="x", expand=True)

# ----------------------------------------------------- Pestaña Comparar

# Fila superior: [Mes:] [caja] [Comparar]   noviembre 2026 contra octubre 2026
fila_comparar = ctk.CTkFrame(pestana_comparar, fg_color="transparent")
fila_comparar.pack(fill="x", pady=(5, 10))

ctk.CTkLabel(fila_comparar, text="Mes:").pack(side="left", padx=(0, 10))
entrada_comparar_mes = ctk.CTkEntry(fila_comparar, width=150, placeholder_text="AAAA-MM")
entrada_comparar_mes.pack(side="left")
entrada_comparar_mes.insert(0, datetime.now().strftime("%Y-%m"))

boton_comparar = ctk.CTkButton(fila_comparar, text="Comparar", width=100, command=comparar)
boton_comparar.pack(side="left", padx=10)

etiqueta_contra = ctk.CTkLabel(fila_comparar, text="",
                               font=ctk.CTkFont(weight="bold"))
etiqueta_contra.pack(side="left", padx=(10, 0))

# Casilla pegada a la orilla derecha; al marcarla o desmarcarla se vuelve a comparar
casilla_ocultar = ctk.CTkCheckBox(fila_comparar, text="Ocultar sin cambio", command=comparar)
casilla_ocultar.pack(side="right")

tabla_comparar = crear_tabla_visual(pestana_comparar, COLUMNAS_COMPARAR)
# Cada fila lleva como "etiqueta" (tag) su estado; aquí se le da color a cada una
for estado, color in COLORES_ESTADO.items():
    tabla_comparar.tag_configure(estado, background=color)

etiqueta_comparacion = ctk.CTkLabel(pestana_comparar, text="", justify="left")
etiqueta_comparacion.pack(anchor="w", pady=(10, 0))

# ------------------------------------------------------------------ Arranque

ventana.bind("<Return>", al_presionar_enter)

# Al escribir un código ya registrado se llenan solos su descripción y almacén
entrada_codigo.bind("<KeyRelease>", al_escribir_codigo)

# Separadores automáticos y bloqueo de días/meses que no existen (formato_fechas.py)
activar_formato_mes(entrada_mes)
activar_formato_fecha(entrada_caducidad)
activar_formato_mes(entrada_filtro_mes)
activar_formato_mes(entrada_comparar_mes)

hacer_respaldo_automatico()  # Uno por día, al abrir el programa
cargar_tabla()
entrada_codigo.focus()
ventana.mainloop()
