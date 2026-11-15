"""Conversión de Markdown al XML de OneNote para crear páginas o añadir contenido."""

from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

NS_URI = "http://schemas.microsoft.com/office/onenote/2013/onenote"
NS = f'xmlns:one="{NS_URI}"'

TAMANO_TITULOS = {1: 16, 2: 14, 3: 12, 4: 11, 5: 11, 6: 11}
FUENTE_CODIGO = "Consolas"


# ---------- formato en línea ----------

_INLINE = re.compile(
    r"(?P<code>`[^`]+`)"
    r"|(?P<link>\[(?P<lt>[^\]]+)\]\((?P<lu>[^)\s]+)\))"
    r"|(?P<b>\*\*(?P<bt>.+?)\*\*|__(?P<bt2>.+?)__)"
    r"|(?P<s>~~(?P<st>.+?)~~)"
    r"|(?P<i>(?<![\w*])\*(?P<it>[^*\s][^*]*?)\*(?![\w*])|(?<![\w_])_(?P<it2>[^_\s][^_]*?)_(?![\w_]))"
)


def en_linea(texto: str) -> str:
    """Markdown en línea → HTML que entiende OneNote dentro de <one:T>."""
    out = []
    pos = 0
    for m in _INLINE.finditer(texto):
        out.append(html.escape(texto[pos:m.start()], quote=False))
        if m.group("code"):
            out.append(f"<span style='font-family:{FUENTE_CODIGO}'>{html.escape(m.group('code')[1:-1], quote=False)}</span>")
        elif m.group("link"):
            url = m.group("lu")
            if re.match(r"^(https?|mailto|onenote):", url, re.I):
                out.append(f'<a href="{html.escape(url)}">{en_linea(m.group("lt"))}</a>')
            else:
                out.append(html.escape(m.group(0), quote=False))
        elif m.group("b"):
            out.append(f"<span style='font-weight:bold'>{en_linea(m.group('bt') or m.group('bt2'))}</span>")
        elif m.group("s"):
            out.append(f"<span style='text-decoration:line-through'>{en_linea(m.group('st'))}</span>")
        elif m.group("i"):
            out.append(f"<span style='font-style:italic'>{en_linea(m.group('it') or m.group('it2'))}</span>")
        pos = m.end()
    out.append(html.escape(texto[pos:], quote=False))
    return "".join(out)


def _cdata(contenido: str) -> str:
    return "<![CDATA[" + contenido.replace("]]>", "]]]]><![CDATA[>") + "]]>"


def _attr(v: str) -> str:
    return html.escape(v, quote=True)


# ---------- bloques ----------

@dataclass
class Bloque:
    nivel: int
    tipo: str                 # "p" | "titulo" | "vineta" | "numero" | "tarea" | "codigo" | "tabla"
    texto: str = ""
    nivel_titulo: int = 0
    hecho: bool = False
    filas: list[list[str]] = field(default_factory=list)
    hijos: list["Bloque"] = field(default_factory=list)


_RE_TITULO = re.compile(r"^(#{1,6})\s+(.*)$")
_RE_TAREA = re.compile(r"^[-*+]\s+\[([ xX])\]\s+(.*)$")
_RE_VINETA = re.compile(r"^[-*+]\s+(.*)$")
_RE_NUMERO = re.compile(r"^\d+[.)]\s+(.*)$")
_RE_FILA = re.compile(r"^\|.*\|$")
_RE_SEPARADOR = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$")


def _nivel(linea: str) -> int:
    espacios = len(linea) - len(linea.lstrip(" \t"))
    prefijo = linea[:espacios].replace("\t", "    ")
    return len(prefijo) // 2


def _celdas(linea: str) -> list[str]:
    cuerpo = linea.strip()[1:-1]
    return [c.strip().replace("\\|", "|") for c in re.split(r"(?<!\\)\|", cuerpo)]


def parsear_markdown(md: str) -> list[Bloque]:
    lineas = md.replace("\r\n", "\n").split("\n")
    planos: list[Bloque] = []
    i = 0
    while i < len(lineas):
        linea = lineas[i]
        if not linea.strip():
            i += 1
            continue
        nivel = _nivel(linea)
        s = linea.strip()

        if s.startswith("```"):
            i += 1
            codigo = []
            while i < len(lineas) and not lineas[i].strip().startswith("```"):
                codigo.append(lineas[i])
                i += 1
            i += 1  # cierre
            for c in codigo or [""]:
                planos.append(Bloque(nivel, "codigo", c.rstrip()))
            continue

        if _RE_FILA.match(s):
            filas = []
            while i < len(lineas) and _RE_FILA.match(lineas[i].strip()):
                if not _RE_SEPARADOR.match(lineas[i].strip()):
                    filas.append(_celdas(lineas[i]))
                i += 1
            planos.append(Bloque(nivel, "tabla", filas=filas))
            continue

        if m := _RE_TITULO.match(s):
            planos.append(Bloque(nivel, "titulo", m.group(2), nivel_titulo=len(m.group(1))))
        elif m := _RE_TAREA.match(s):
            planos.append(Bloque(nivel, "tarea", m.group(2), hecho=m.group(1).lower() == "x"))
        elif m := _RE_VINETA.match(s):
            planos.append(Bloque(nivel, "vineta", m.group(1)))
        elif m := _RE_NUMERO.match(s):
            planos.append(Bloque(nivel, "numero", m.group(1)))
        elif s.startswith(">"):
            planos.append(Bloque(nivel, "p", s.lstrip("> ").strip()))
        else:
            planos.append(Bloque(nivel, "p", s))
        i += 1

    # Anidar por sangría: un bloque más sangrado cuelga del anterior
    raiz: list[Bloque] = []
    pila: list[Bloque] = []
    for b in planos:
        while pila and pila[-1].nivel >= b.nivel:
            pila.pop()
        (pila[-1].hijos if pila else raiz).append(b)
        pila.append(b)
    return raiz


# ---------- XML ----------

class Generador:
    def __init__(self, indice_tarea: int = 0) -> None:
        self.indice_tarea = indice_tarea
        self.usa_tareas = False

    def oe(self, b: Bloque) -> str:
        partes = []
        if b.tipo == "tarea":
            self.usa_tareas = True
            partes.append(f'<one:Tag index="{self.indice_tarea}" completed="{"true" if b.hecho else "false"}"/>')
        if b.tipo == "vineta":
            partes.append('<one:List><one:Bullet bullet="2"/></one:List>')
        elif b.tipo == "numero":
            partes.append('<one:List><one:Number numberSequence="0" numberFormat="##."/></one:List>')

        if b.tipo == "titulo":
            tam = TAMANO_TITULOS[b.nivel_titulo]
            contenido = f"<span style='font-weight:bold;font-size:{tam}.0pt'>{en_linea(b.texto)}</span>"
            partes.append(f"<one:T>{_cdata(contenido)}</one:T>")
        elif b.tipo == "codigo":
            contenido = f"<span style='font-family:{FUENTE_CODIGO}'>{html.escape(b.texto, quote=False) or '&nbsp;'}</span>"
            partes.append(f"<one:T>{_cdata(contenido)}</one:T>")
        elif b.tipo == "tabla":
            partes.append(self.tabla(b.filas))
        else:
            partes.append(f"<one:T>{_cdata(en_linea(b.texto))}</one:T>")

        if b.hijos:
            partes.append(self.oechildren(b.hijos))
        return "<one:OE>" + "".join(partes) + "</one:OE>"

    def oechildren(self, bloques: list[Bloque]) -> str:
        return "<one:OEChildren>" + "".join(self.oe(b) for b in bloques) + "</one:OEChildren>"

    def tabla(self, filas: list[list[str]]) -> str:
        if not filas:
            return ""
        ancho = max(len(f) for f in filas)
        cols = "".join(f'<one:Column index="{i}" width="140" isLocked="false"/>' for i in range(ancho))
        xml_filas = []
        for n, fila in enumerate(filas):
            celdas = []
            for c in fila + [""] * (ancho - len(fila)):
                contenido = en_linea(c)
                if n == 0:
                    contenido = f"<span style='font-weight:bold'>{contenido}</span>"
                celdas.append(f"<one:Cell><one:OEChildren><one:OE><one:T>{_cdata(contenido)}</one:T></one:OE></one:OEChildren></one:Cell>")
            xml_filas.append("<one:Row>" + "".join(celdas) + "</one:Row>")
        return f'<one:Table bordersVisible="true"><one:Columns>{cols}</one:Columns>{"".join(xml_filas)}</one:Table>'

    def tagdef(self) -> str:
        return (f'<one:TagDef index="{self.indice_tarea}" type="0" symbol="3" fontColor="automatic" '
                f'highlightColor="none" name="Tareas pendientes"/>')


def xml_nueva_pagina(pagina_id: str, titulo: str, markdown: str) -> str:
    g = Generador()
    cuerpo = g.oechildren(parsear_markdown(markdown)) if markdown.strip() else ""
    tagdefs = g.tagdef() if g.usa_tareas else ""
    outline = f"<one:Outline>{cuerpo}</one:Outline>" if cuerpo else ""
    titulo_xml = f"<one:Title><one:OE><one:T>{_cdata(html.escape(titulo, quote=False))}</one:T></one:OE></one:Title>"
    return f'<?xml version="1.0"?><one:Page {NS} ID="{_attr(pagina_id)}">{tagdefs}{titulo_xml}{outline}</one:Page>'


def xml_agregar(pagina_xml_actual: str, markdown: str) -> str:
    """XML para añadir contenido al final de una página existente.

    Importante: si se envía un Outline existente (con su objectID), OneNote REEMPLAZA todo su
    contenido por el enviado. Por eso se crea un Outline nuevo, sin objectID, colocado debajo
    de todo lo que ya hay en la página: lo existente no se toca.
    """
    raiz = ET.fromstring(pagina_xml_actual)
    ns = "{" + NS_URI + "}"
    pagina_id = raiz.get("ID", "")

    # Reutilizar la etiqueta de "tarea pendiente" de la página, o crear una con índice libre.
    # OneNote exige que el TagDef vaya en el mismo XML que las casillas que lo usan.
    existente: ET.Element | None = None
    usados = []
    for t in raiz.findall(f"{ns}TagDef"):
        usados.append(int(t.get("index", "0")))
        if existente is None and t.get("type") == "0" and t.get("symbol") == "3":
            existente = t
    if existente is not None:
        g = Generador(int(existente.get("index", "0")))
    else:
        g = Generador(max(usados) + 1 if usados else 0)
    cuerpo_oes = "".join(g.oe(b) for b in parsear_markdown(markdown))
    if not cuerpo_oes:
        raise ValueError("No hay contenido que añadir.")
    tagdefs = ""
    if g.usa_tareas:
        if existente is not None:
            attrs = " ".join(f'{k}="{_attr(v)}"' for k, v in existente.attrib.items())
            tagdefs = f"<one:TagDef {attrs}/>"
        else:
            tagdefs = g.tagdef()

    x, y, z = _hueco_debajo(raiz, ns)
    bloque = (f'<one:Outline><one:Position x="{x:.1f}" y="{y:.1f}" z="{z}"/>'
              f"<one:OEChildren>{cuerpo_oes}</one:OEChildren></one:Outline>")
    return f'<?xml version="1.0"?><one:Page {NS} ID="{_attr(pagina_id)}">{tagdefs}{bloque}</one:Page>'


def _hueco_debajo(raiz: ET.Element, ns: str) -> tuple[float, float, int]:
    """Posición libre bajo el contenido de la página (alineada con el bloque más bajo)."""
    x_def, y_def = 36.0, 86.4  # donde OneNote pone el primer bloque bajo el título
    fondo, x, z_max = None, x_def, 0
    for el in raiz:
        if el.tag[len(ns):] not in ("Outline", "Image", "InkDrawing", "InsertedFile", "MediaFile"):
            continue
        pos, tam = el.find(f"{ns}Position"), el.find(f"{ns}Size")
        if pos is None:
            continue
        y = float(pos.get("y", 0)) + (float(tam.get("height", 0)) if tam is not None else 0)
        z_max = max(z_max, int(float(pos.get("z", 0))))
        if fondo is None or y > fondo:
            fondo, x = y, float(pos.get("x", x_def))
    if fondo is None:
        return x_def, y_def, 0
    return x, fondo + 18.0, z_max + 1
