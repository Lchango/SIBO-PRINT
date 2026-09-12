"""
Generador de etiquetas ZPL - Modelo "102x102 PAPEL PRODUCTO"
================================================================
Script parametrizable listo para integrar en un flujo de automatizacion
de impresion masiva en impresoras termicas Zebra (o compatibles ZPL).

Por que ZPL y no HTML/CSS:
Una impresora termica Zebra (como la ZD230) interpreta ZPL de forma nativa,
directo al cabezal de impresion, en dots (puntos) a 8 dots/mm (203dpi).
HTML/CSS requeriria un renderizador intermedio (navegador headless ->
imagen -> conversion a ZPL/raster) para poder imprimirse en este tipo de
hardware, lo cual agrega latencia, dependencias y puntos de falla
innecesarios para una etiqueta que ya se construye pixel a pixel.

Este layout es el mismo ya validado visualmente (Labelary) y ajustado
para no producir textos superpuestos, basado en la plantilla
`FORMATO 102x102 PAPEL PRODUCTO.prn` usada por la app SIBO PRINT.

Dimension fisica: 102 x 102 mm -> 816 x 816 dots a 8 dots/mm.
"""

from __future__ import annotations

import re
import socket


# ---------------------------------------------------------------------------
# Plantilla ZPL (idéntica a la usada en producción, con placeholders {})
# ---------------------------------------------------------------------------

_ZPL_TEMPLATE = """CT~~CD,~CC^~CT~
^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR6,6~SD26^JUS^LRN^CI0^XZ
^XA
^MMT
^PW816
^LL816
^LS0
^FO360,0^GFA,12544,12544,00056,:Z64:
eJzt2s+O0zwQAHAHI7wHJHPg7sJzfJI5fA/EkQPCRiBx5BF4FCx9LxKJA9cc80lRzEz+NM7UpfEEBYE62m132/66iT0zcdIV4h73uMc9fmH8w3TvmK7nMcl0iukMuIbhbC9Uy3CuF5rhqtiJyHAqtjIy9s/E1nCci//HWJe7iFHuKqZTgwvFTjOdZTo3OF/sIs/Js9NFbhzO+BW+i9w4LEMUuYUVVX3FdHJxXYnTiysqXcN0bnFFpbuwolJKhrMotZPhLErRX+EKWOq4aVaULokr6/T7XVmnX/KsLnKWlS6p84c4s9sVsaVPFK5hzq5wYaAOdudCKlxQnAuiZrpwjDsXkme6QjYXYPGSd3JFXQnDsqad7wwrXc6uLnWaN+1zQfi9Dgukg41vXY1bg/tt4AZ2RHUwho2dXqxIuqwdPiExKWDg5MrJjGtnJ3HHNb4QJirnurzTOEEGMbxi7Soy7SuH95gZjYBtLXAWt8Th0xdOZNx5XBwWyrAbMAI51+QdPlXFEXviHE0zWwuDf6vGMzXbwzmXi5MbpnQaQ0vTbHEy1rJTMSh0YXJqiwui1bBn0Ud8i0vns07B4691D9sLriHOrNKFOHy+Axfyrs87jY8bmCMX4teWOB1JV0pcNzlbxy9dxrXUwR7Pzjaj64f582en8i6ga8df0X0eXVi55oprJmfq+Ik6SbsSdWFwH2LGhdvuffxv7SralTY6sc35nBMbnPt24fqM82Q8vftOnCMHsSRfkvnzNq7z5WduyLMxX7yhzpKDWCY/Ia+93u7UXA8x5+jFF+L0WH85V19zMLEtFCH2Ca+o0+Tgt+ovstNIwEnq1FU39EGFPVNknb9wzeR6F2Vsse96fJNm7cSFw34N5ZX0aziUDf16cfSyYuKS44MXjrruqkuOR15Y4qr2qpuPf3Whk0jG462HBrx29Bpm4sTQ+LA+4V5TV193rh+vO2bdK+LC7MJSuU5M6xfepdJ1vGG6t0y34bq3xy87/Ph0euipcH6Lm+PxdG+8u73aHdwD3pxmp1t7ezU/OTCn57Pr7e2zBy8endC9Wv6eogeIn2/ny38TF267x8/APZzEKXU3ByZx83ZKuhDJu6oGJ2vxInU3T8c83gz79+Jj6sINF8QDbKV4BvPwZHqoogvJXGRSuFoveDfH4BgflYwr83I3rrAD07E+gKBLyW1hmTs4On+YM7tcMTvc6T2OkdiamS9qj2MULtdJZh0d7TY1wlz8DhcYzu1wnukYjO0s37E+ujfHO0a5851mlfvQKI52nP+fEJJXfntc4LjqD3GC1yb2OBaDBsN0zP/wYTtW2UJDYzrDPEc+2mlWufOdqv9uJ8PBzvNcxWP3uMc97lEQPwDGR0iV:E178
^FO21,232^GFA,04608,04608,00144,:Z64:
eJztzqENACAAxMDfgP23ROKeFXBAcqcqmwBc0pm2GX0glh8/fvz48XP4AwAf2g2bbsE=:1919
^FT27,78^A0N,53,52^FH\\^FD{fecha}^FS
^FT61,220^A0N,95,46^FB750,1,0,L,0^FH\\^FD{codigo}^FS
^FT27,280^A0N,28,27^FH\\^FDCONTENIDO:^FS
^FT27,355^A0N,{contenido_h},{contenido_w}^FB748,2,3,L,0^FH\\^FD{contenido}^FS
^FO27,436^GB748,0,4^FS
^FT27,470^A0N,26,25^FH\\^FDUBICACION:^FS
^FT522,466^A0N,26,25^FH\\^FDBODEGA:^FS
^FT27,518^A0N,46,64^FB480,1,0,L,0^FH\\^FD{ubicacion}^FS
^FT522,511^A0N,36,50^FB220,1,0,L,0^FH\\^FD{bodega}^FS
^FO27,528^GB748,0,4^FS
^FT27,562^A0N,26,25^FH\\^FDRESPONSABLE DE INGRESO:^FS
^FT522,562^A0N,26,25^FH\\^FDCANTIDAD:^FS
^FT21,614^A0N,55,40^FB480,1,0,L,0^FH\\^FD{responsable}^FS
^FT522,608^A0N,38,48^FB220,1,0,L,0^FH\\^FD{cantidad}^FS
^FO21,624^GB748,0,4^FS
^FT112,745^A0N,36,36^FH\\^FDCODIGO:^FS
^FT290,798^BQN,2,6
^FH\\^FDLA,{codigo_qr}^FS
^PQ{copias},0,1,Y^XZ
"""


# ---------------------------------------------------------------------------
# Auto-ajuste de fuente para la columna C (CONTENIDO / descripcion)
# ---------------------------------------------------------------------------

# Espacio real disponible para este campo en la etiqueta (medido sobre la
# plantilla): ancho del bloque de texto y el hueco vertical entre la etiqueta
# "CONTENIDO:" (y=280) y la siguiente linea divisoria (y=436).
_CONTENIDO_FB_WIDTH = 748          # dots (^FB)
_CONTENIDO_MAX_LINES = 2
_CONTENIDO_LINE_GAP = 3            # dots entre lineas (el "3" del ^FB)
_CONTENIDO_Y_START = 355
_CONTENIDO_Y_LIMIT = 436           # antes de este y no debe entrar ningun pixel

# Relacion empirica ancho-de-caracter -> avance real en dots para la fuente A0
# (el avance real de un caracter es menor al parametro de ancho que se le pasa
# a ^A0; este factor sale de comparar los renders de prueba contra el texto
# real impreso).
_CHAR_ADVANCE_RATIO = 0.62

# Tabla de tamanos candidatos, de mas grande a mas chico. `height` siempre dentro
# del limite vertical (2 lineas + espacio <= 81 dots disponibles); `width` se
# reduce junto con la altura para mantener las letras proporcionadas.
_CONTENIDO_SIZE_STEPS = [
    (34, 22),  # tamano por defecto / textos cortos-medianos (hasta ~70 caracteres)
    (30, 19),
    (26, 17),
    (22, 14),
    (18, 12),  # tamano minimo legible; textos mas largos que esto se recortan
]


def _fit_contenido_font(texto: str) -> tuple[int, int]:
    """Calcula (alto, ancho) de fuente para que `texto` quepa en el espacio de
    CONTENIDO sin invadir ni la etiqueta "CONTENIDO:" de arriba, ni la linea
    divisoria de abajo, ni desbordar a los lados.

    Estrategia: probar los tamanos de mas grande a mas chico y quedarse con el
    primero donde el texto quepa en <= 2 lineas dentro del ancho disponible.
    Si ni el tamano minimo alcanza, se usa el minimo de todas formas (el ^FB
    de la plantilla recorta el sobrante de forma limpia, sin superponer texto).
    """
    largo = len(texto or "")
    if largo == 0:
        return _CONTENIDO_SIZE_STEPS[0]

    for height, width in _CONTENIDO_SIZE_STEPS:
        avance = max(1, width * _CHAR_ADVANCE_RATIO)
        chars_por_linea = int(_CONTENIDO_FB_WIDTH / avance)
        capacidad_total = chars_por_linea * _CONTENIDO_MAX_LINES
        # Margen de seguridad del 10% para no rozar el borde del bloque
        if largo <= capacidad_total * 0.90:
            return height, width

    return _CONTENIDO_SIZE_STEPS[-1]


def _zpl_safe(texto: str) -> str:
    """Escapa/limpia texto de datos reales antes de insertarlo en un campo ^FD.

    - Los caracteres ^ y ~ tienen significado especial en ZPL (inician comandos);
      si aparecen dentro de un dato (ej. una descripcion con "~" o "^"), rompen
      el resto de la etiqueta. Se reemplazan por un guion como salvaguarda.
    - Recorta espacios sobrantes.
    """
    if texto is None:
        return ""
    texto = str(texto).strip()
    texto = texto.replace("^", "-").replace("~", "-")
    return texto


def build_zpl_label(
    codigo: str,
    contenido: str,
    ubicacion: str = "",
    bodega: str = "",
    responsable: str = "",
    cantidad: str = "",
    fecha: str | None = None,
    copias: int = 1,
) -> str:
    """Arma el ZPL final de una etiqueta 102x102mm lista para enviar a imprimir.

    Parametros
    ----------
    codigo:      Codigo principal del articulo. Se usa tanto en el texto grande
                 como en el contenido del QR (ej. "10-06-04-013").
    contenido:   Descripcion / presentacion del articulo (ej. "AMARRAS PLASTICAS 35CM (100U)").
                 El tamano de letra se AUTO-AJUSTA segun el largo del texto (ver
                 `_fit_contenido_font`): mientras mas caracteres, mas chica la
                 fuente, para que nunca invada el espacio de "CONTENIDO:" arriba
                 ni la linea divisoria de abajo. Si el texto excede incluso el
                 tamano minimo, se recorta limpiamente (sin superposicion).
    ubicacion:   Codigo de ubicacion/rack (ej. "R1702B01").
    bodega:      Codigo de bodega (ej. "BQEF").
    responsable: Nombre de quien realizo el ingreso (ej. "GGROMERO").
    cantidad:    Cantidad + unidad (ej. "50UND").
    fecha:       Fecha a mostrar (ej. "01/01/2023"). Si se omite, usa la fecha actual.
    copias:      Numero de copias de esta misma etiqueta a imprimir (^PQ).

    Retorna
    -------
    str: cadena ZPL completa, lista para enviar en modo RAW a la impresora.
    """
    import datetime

    if fecha is None:
        fecha = datetime.date.today().strftime("%d/%m/%Y")

    codigo = _zpl_safe(codigo)
    contenido = _zpl_safe(contenido)
    contenido_h, contenido_w = _fit_contenido_font(contenido)

    return _ZPL_TEMPLATE.format(
        fecha=_zpl_safe(fecha),
        codigo=codigo,
        contenido=contenido,
        contenido_h=contenido_h,
        contenido_w=contenido_w,
        ubicacion=_zpl_safe(ubicacion),
        bodega=_zpl_safe(bodega),
        responsable=_zpl_safe(responsable),
        cantidad=_zpl_safe(cantidad),
        codigo_qr=codigo,
        copias=max(1, int(copias)),
    )


# ---------------------------------------------------------------------------
# Envio a impresora
# ---------------------------------------------------------------------------

def send_zpl_over_network(zpl: str, ip: str, port: int = 9100, timeout: float = 5.0) -> None:
    """Envia el ZPL directo a una impresora Zebra en red (puerto RAW 9100).

    Es el metodo recomendado para automatizacion masiva / servidores, ya que
    no depende de que la impresora este instalada como cola de Windows en la
    maquina que ejecuta el script.
    """
    with socket.create_connection((ip, port), timeout=timeout) as sock:
        sock.sendall(zpl.encode("utf-8"))


def send_zpl_windows_printer(zpl: str, printer_name: str) -> None:
    """Envia el ZPL a una impresora instalada en Windows (modo RAW), usando pywin32.

    Requiere: pip install pywin32
    """
    import win32print

    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta ZPL", None, "RAW"))
        try:
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, zpl.encode("utf-8"))
            win32print.EndPagePrinter(hPrinter)
        finally:
            win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)


# ---------------------------------------------------------------------------
# Ejemplo de uso para impresion masiva (batch)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Ejemplo: generar e imprimir un lote de etiquetas a partir de una lista de
    # diccionarios (en un flujo real, esta lista vendria de tu Excel/DB/API).
    registros = [
        {
            "codigo": "10-06-04-013",
            "contenido": "AMARRAS PLASTICAS 35CM (100U)",
            "ubicacion": "R1702B01",
            "bodega": "BQEF",
            "responsable": "GGROMERO",
            "cantidad": "50UND",
        },
        {
            "codigo": "10-01-01-001",
            "contenido": "COBERTOR PARA RJ-45",
            "ubicacion": "R1701A01",
            "bodega": "BQZ5",
            "responsable": "MGUANULEMA",
            "cantidad": "2000UND",
        },
    ]

    for registro in registros:
        zpl = build_zpl_label(**registro)

        # Opcion A: imprimir por red (recomendado para automatizacion / servidor)
        # send_zpl_over_network(zpl, ip="192.168.1.50")

        # Opcion B: imprimir a una cola de Windows
        # send_zpl_windows_printer(zpl, printer_name="ZDesigner ZD230-203dpi ZPL")

        print(f"ZPL generado para {registro['codigo']} ({len(zpl)} caracteres)")
