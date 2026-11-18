"""Ayuda a escribir fechas en las cajas de texto mientras el usuario teclea.

- Pone solos los separadores: "03062028" -> "03/06/2028" y "202611" -> "2026-11".
- Entiende las fechas sin ceros: "3/6/2028" -> "03/06/2028".
- No deja escribir un día o mes que no existe (día 45, mes 13, 31 de abril...).
  Si se rechaza una tecla, el borde de la caja parpadea en rojo.

Esto solo ayuda al escribir. La revisión definitiva sigue en inventario.py,
al guardar (por ejemplo, si se corrige un número en medio de la fecha).
"""
from datetime import datetime

import customtkinter as ctk

# Días máximos de cada mes. Febrero va con 29 porque, mientras se escribe el
# día y el mes, todavía no se sabe el año (si es bisiesto o no).
DIAS_POR_MES = (31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

# Símbolos que el usuario puede escribir como separador; la caja los cambia
# por el que corresponde ("/" en fechas, "-" en meses)
SEPARADORES = "/-."

COLOR_RECHAZO = ("#c62828", "#ff6b6b")
BORDE_NORMAL = ctk.ThemeManager.theme["CTkEntry"]["border_color"]


# ------------------------------------------------- ¿Los dígitos son válidos?

def inicio_de_fecha_valido(digitos):
    """¿Estos dígitos pueden ser el principio de una fecha ddmmaaaa que exista?

    Solo revisa lo que ya se escribió: "3" sí (puede ser 30 o 31), "4" no
    (no hay día 40) y "3104" no (abril tiene 30 días).
    """
    if len(digitos) > 8:
        return False
    if len(digitos) >= 1 and digitos[0] not in "0123":
        return False
    if len(digitos) >= 2 and not 1 <= int(digitos[:2]) <= 31:
        return False
    if len(digitos) >= 3 and digitos[2] not in "01":
        return False
    if len(digitos) >= 4:
        dia, mes = int(digitos[:2]), int(digitos[2:4])
        if not 1 <= mes <= 12 or dia > DIAS_POR_MES[mes - 1]:
            return False
    if len(digitos) == 8:
        # Fecha completa: ahora sí se sabe el año, así que se revisa el 29 de febrero
        try:
            datetime.strptime(digitos, "%d%m%Y")
        except ValueError:
            return False
    return True


def inicio_de_mes_valido(digitos):
    """¿Estos dígitos pueden ser el principio de un mes aaaamm que exista?

    El año son 4 dígitos que no empiezan con 0; el mes debe ir de 01 a 12.
    """
    if len(digitos) > 6:
        return False
    if len(digitos) >= 1 and digitos[0] == "0":
        return False
    if len(digitos) >= 5 and digitos[4] not in "01":
        return False
    if len(digitos) == 6 and not 1 <= int(digitos[4:]) <= 12:
        return False
    return True


# ------------------------------------------------ Dar formato al escribir

def marcar_rechazo(entrada):
    """Pone el borde de la caja en rojo un momento para avisar que no se aceptó la tecla."""
    entrada.configure(border_color=COLOR_RECHAZO)
    entrada.after(600, lambda: entrada.configure(border_color=BORDE_NORMAL))


def aplicar_formato(evento, entrada, separador, cortes, es_valido):
    """Limpia lo escrito en la caja y le pone los separadores.

    separador: "/" o "-"
    cortes: después de cuántos dígitos va un separador, ej. (2, 4) para dd/mm/aaaa
    es_valido: función que dice si unos dígitos pueden ser el inicio de una fecha real
    """
    texto = entrada.get()
    if entrada.index("insert") < len(texto):
        return  # Se está corrigiendo en medio de la fecha: no se mueve nada

    inicios = (0,) + cortes  # Dígito donde empieza cada parte: (0, 2, 4) = día, mes, año

    # 1. Revisar carácter por carácter y quedarse solo con lo válido
    digitos = ""
    rechazado = False
    for caracter in texto:
        if caracter.isdigit() and es_valido(digitos + caracter):
            digitos += caracter
        elif (caracter.isdigit() and len(digitos) in inicios
              and es_valido(digitos + "0" + caracter)):
            digitos += "0" + caracter  # Un "4" al empezar el día solo puede ser "04"
        elif caracter in SEPARADORES:
            # "3/" quiere decir "03/": si el día o el mes lleva una sola cifra,
            # se completa con un cero a la izquierda
            if len(digitos) - 1 in inicios and len(digitos) + 1 in cortes:
                con_cero = digitos[:-1] + "0" + digitos[-1]
                if es_valido(con_cero):
                    digitos = con_cero
        else:
            rechazado = True  # Una letra, o un dígito que daría una fecha imposible
            break             # Se descarta esto y todo lo que sigue

    # 2. Volver a armar el texto, con un separador en cada corte
    nuevo = ""
    for posicion, digito in enumerate(digitos):
        if posicion in cortes:
            nuevo += separador
        nuevo += digito
    # Al completar el día o el mes, el separador aparece de una vez,
    # salvo que se esté borrando (si no, nunca se podría borrar la "/")
    if len(digitos) in cortes and evento.keysym != "BackSpace":
        nuevo += separador

    # 3. Solo se reescribe la caja si algo cambió
    if nuevo != texto:
        entrada.delete(0, "end")
        entrada.insert(0, nuevo)
    if rechazado:
        marcar_rechazo(entrada)


def activar_formato_fecha(entrada):
    """La caja dará formato dd/mm/aaaa al texto mientras se escribe."""
    entrada.bind("<KeyRelease>", lambda evento: aplicar_formato(
        evento, entrada, "/", (2, 4), inicio_de_fecha_valido))


def activar_formato_mes(entrada):
    """La caja dará formato aaaa-mm al texto mientras se escribe."""
    entrada.bind("<KeyRelease>", lambda evento: aplicar_formato(
        evento, entrada, "-", (4,), inicio_de_mes_valido))
