"""Conversión del XML de una página de OneNote a Markdown legible."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from html.parser import HTMLParser

MAX_OCR = 2000  # caracteres de texto reconocido por imagen que se incluyen en el texto


@dataclass
class ImagenRef:
    numero: int
    callback_id: str | None
    ocr: str
    alt: str


@dataclass
class Pagina:
    titulo: str
    markdown: str
    imagenes: list[ImagenRef] = field(default_factory=list)
    texto_plano: str = ""  # para búsquedas (incluye el OCR completo)


# ---------- texto enriquecido de los <one:T> (HTML simplificado) ----------

class _HTMLaMD(HTMLParser):
    """Convierte el HTML de un <one:T> en Markdown: negrita, cursiva, tachado, enlaces y saltos."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.segmentos: list[tuple[str, frozenset[str], str | None]] = []
        self._estilos: list[frozenset[str]] = [frozenset()]
        self._enlace: list[str | None] = [None]

    @staticmethod
    def _de_style(style: str) -> set[str]:
        s = style.replace(" ", "").lower()
        out = set()
        if "font-weight:bold" in s:
            out.add("b")
        if "font-style:italic" in s:
            out.add("i")
        if "line-through" in s:
            out.add("s")
        return out

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        a = dict(attrs)
        if tag == "br":
            self.segmentos.append(("\n", frozenset(), None))
            return
        actual = set(self._estilos[-1])
        if tag in ("b", "strong"):
            actual.add("b")
        elif tag in ("i", "em"):
            actual.add("i")
        elif tag in ("s", "strike", "del"):
            actual.add("s")
        actual |= self._de_style(a.get("style") or "")
        self._estilos.append(frozenset(actual))
        self._enlace.append(a.get("href") if tag == "a" else self._enlace[-1])

    def handle_endtag(self, tag: str) -> None:
        if tag == "br":
            return
        if len(self._estilos) > 1:
            self._estilos.pop()
            self._enlace.pop()

    def handle_data(self, data: str) -> None:
        self.segmentos.append((data, self._estilos[-1], self._enlace[-1]))

    def markdown(self) -> str:
        # Unir segmentos contiguos con el mismo formato
        grupos: list[list] = []
        for texto, est, enl in self.segmentos:
            if grupos and grupos[-1][1] == est and grupos[-1][2] == enl and texto != "\n" and grupos[-1][0] != "\n":
                grupos[-1][0] += texto
            else:
                grupos.append([texto, est, enl])
        out = []
        for texto, est, enl in grupos:
            if not texto.strip() or texto == "\n":
                out.append(texto)
                continue
            # Los marcadores van pegados al texto, con los espacios por fuera
            ini = texto[: len(texto) - len(texto.lstrip())]
            fin = texto[len(texto.rstrip()):]
            nucleo = texto.strip()
            if enl:
                nucleo = f"[{nucleo}]({enl})"
            if "s" in est:
                nucleo = f"~~{nucleo}~~"
            if "i" in est:
                nucleo = f"*{nucleo}*"
            if "b" in est:
                nucleo = f"**{nucleo}**"
            out.append(ini + nucleo + fin)
        return "".join(out)


def html_a_md(html: str, style: str | None = None) -> str:
    p = _HTMLaMD()
    if style:
        p._estilos = [frozenset(p._de_style(style))]
    p.feed(html or "")
    p.close()
    return p.markdown()


def _plano(md: str) -> str:
    """Quita marcas de Markdown (para búsquedas y fragmentos)."""
    md = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", md)
    return re.sub(r"[*~`]+", "", md)


# ---------- recorrido de la página ----------

class _Lector:
    def __init__(self, raiz: ET.Element) -> None:
        self.ns = raiz.tag.split("}")[0] + "}" if raiz.tag.startswith("{") else ""
        self.raiz = raiz
        self.imagenes: list[ImagenRef] = []
        self.ocr_completo: list[str] = []
        self.estilos = {
            q.get("index"): (q.get("name") or "").lower()
            for q in raiz.findall(self._t("QuickStyleDef"))
        }
        self.etiquetas = {
            t.get("index"): (t.get("name") or "", t.get("type"), t.get("symbol"))
            for t in raiz.findall(self._t("TagDef"))
        }

    def _t(self, nombre: str) -> str:
        return self.ns + nombre

    def _hijo(self, el: ET.Element, nombre: str) -> ET.Element | None:
        return el.find(self._t(nombre))

    # --- elementos sueltos ---

    def imagen(self, el: ET.Element) -> str:
        n = len(self.imagenes) + 1
        cb = self._hijo(el, "CallbackID")
        ocr_el = el.find(f"{self._t('OCRData')}/{self._t('OCRText')}")
        ocr = " ".join((ocr_el.text or "").split()) if ocr_el is not None else ""
        alt = el.get("alt") or ""
        self.imagenes.append(ImagenRef(n, cb.get("callbackID") if cb is not None else None, ocr, alt))
        if ocr:
            self.ocr_completo.append(ocr)
        partes = [f"🖼 Imagen {n}"]
        if alt and alt.lower() not in ("imagen", "image"):
            partes.append(f"descripción: {alt}")
        if ocr:
            corto = ocr if len(ocr) <= MAX_OCR else ocr[:MAX_OCR] + "…"
            partes.append(f"texto reconocido: «{corto}»")
        else:
            partes.append("sin texto reconocido")
        return "[" + " — ".join(partes) + "]"

    def archivo(self, el: ET.Element) -> str:
        return f"[📎 Archivo adjunto: {el.get('preferredName') or 'sin nombre'}]"

    def tinta(self, el: ET.Element) -> str:
        palabras = [w.get("recognizedText") for w in el.iter(self._t("InkWord")) if w.get("recognizedText")]
        if palabras:
            return f"[✍ escrito a mano: «{' '.join(palabras)}»]"
        return "[✍ dibujo o escritura a mano]"

    def tabla(self, el: ET.Element, sangria: str) -> str:
        filas = []
        for fila in el.findall(self._t("Row")):
            celdas = []
            for celda in fila.findall(self._t("Cell")):
                hijos = self._hijo(celda, "OEChildren")
                lineas = self.bloque(hijos, 0) if hijos is not None else []
                texto = "<br>".join(l.strip() for l in lineas if l.strip())
                celdas.append(texto.replace("|", "\\|") or " ")
            filas.append(celdas)
        if not filas:
            return ""
        ancho = max(len(f) for f in filas)
        filas = [f + [" "] * (ancho - len(f)) for f in filas]
        out = [f"{sangria}| " + " | ".join(filas[0]) + " |",
               f"{sangria}|" + "---|" * ancho]
        out += [f"{sangria}| " + " | ".join(f) + " |" for f in filas[1:]]
        return "\n".join(out)

    # --- párrafos (OE) ---

    def oe(self, el: ET.Element, nivel: int) -> list[str]:
        sangria = "  " * nivel
        prefijo = ""

        lista = self._hijo(el, "List")
        if lista is not None:
            num = self._hijo(lista, "Number")
            prefijo = (num.get("text") or "1.").strip() + " " if num is not None else "- "

        for tag in el.findall(self._t("Tag")):
            nombre, tipo, simbolo = self.etiquetas.get(tag.get("index"), ("", None, None))
            if tag.get("completed") is not None:
                prefijo += "[x] " if tag.get("completed") == "true" else "[ ] "
            elif nombre:
                prefijo += f"[{nombre}] "

        estilo = self.estilos.get(el.get("quickStyleIndex"), "")
        encabezado = ""
        m = re.fullmatch(r"h([1-6])", estilo)
        if m and not prefijo:
            encabezado = "#" * min(int(m.group(1)) + 1, 6) + " "  # el título de la página es el "#"

        lineas: list[str] = []
        contenido: list[str] = []
        for hijo in el:
            nombre = hijo.tag[len(self.ns):]
            if nombre == "T":
                contenido.append(html_a_md(hijo.text or "", hijo.get("style")))
            elif nombre == "Image":
                contenido.append(self.imagen(hijo))
            elif nombre == "InsertedFile":
                contenido.append(self.archivo(hijo))
            elif nombre in ("InkWord", "InkParagraph", "InkDrawing"):
                contenido.append(self.tinta(hijo))
            elif nombre == "MediaFile":
                contenido.append(f"[🎵 Audio/vídeo: {hijo.get('preferredName') or 'sin nombre'}]")
            elif nombre == "Table":
                if contenido:
                    lineas.append(sangria + prefijo + encabezado + "".join(contenido))
                    contenido, prefijo, encabezado = [], "", ""
                lineas.append(self.tabla(hijo, sangria))

        texto = "".join(contenido)
        if texto or prefijo:
            primera, *resto = (texto or "").split("\n")
            if estilo in ("code",):
                primera = f"`{primera}`" if primera.strip() else primera
            if estilo in ("cite", "quote"):
                primera = f"> {primera}"
            lineas.append(sangria + prefijo + encabezado + primera)
            # Saltos de línea dentro del mismo párrafo: alineados con el texto
            lineas += [sangria + " " * len(prefijo) + r for r in resto]

        hijos = self._hijo(el, "OEChildren")
        if hijos is not None:
            lineas += self.bloque(hijos, nivel + 1)
        return lineas

    def bloque(self, oechildren: ET.Element, nivel: int) -> list[str]:
        lineas: list[str] = []
        for oe in oechildren.findall(self._t("OE")):
            lineas += self.oe(oe, nivel)
        return lineas

    # --- página completa ---

    def leer(self) -> Pagina:
        titulo_el = self.raiz.find(f"{self._t('Title')}//{self._t('T')}")
        titulo = _plano(html_a_md(titulo_el.text or "")).strip() if titulo_el is not None else ""
        titulo = titulo or self.raiz.get("name") or "(sin título)"

        # Los bloques sueltos de la página se leen de arriba abajo y de izquierda a derecha
        bloques = []
        for el in self.raiz:
            nombre = el.tag[len(self.ns):]
            if nombre not in ("Outline", "Image", "InsertedFile", "InkDrawing", "MediaFile"):
                continue
            pos = self._hijo(el, "Position")
            y = float(pos.get("y", 0)) if pos is not None else 0.0
            x = float(pos.get("x", 0)) if pos is not None else 0.0
            bloques.append((round(y / 10), x, el, nombre))
        bloques.sort(key=lambda b: (b[0], b[1]))

        partes: list[str] = []
        for _, _, el, nombre in bloques:
            if nombre == "Outline":
                hijos = self._hijo(el, "OEChildren")
                lineas = self.bloque(hijos, 0) if hijos is not None else []
                texto = "\n".join(lineas).strip("\n")
            elif nombre == "Image":
                texto = self.imagen(el)
            elif nombre == "InsertedFile":
                texto = self.archivo(el)
            elif nombre == "MediaFile":
                texto = f"[🎵 Audio/vídeo: {el.get('preferredName') or 'sin nombre'}]"
            else:
                texto = self.tinta(el)
            if texto.strip():
                partes.append(texto)

        cuerpo = "\n\n".join(partes)
        md = f"# {titulo}\n\n{cuerpo}".rstrip() + "\n"
        plano = _plano(f"{titulo}\n{cuerpo}") + "\n" + "\n".join(self.ocr_completo)
        return Pagina(titulo=titulo, markdown=md, imagenes=self.imagenes, texto_plano=plano)


def leer_pagina(xml: str) -> Pagina:
    return _Lector(ET.fromstring(xml)).leer()
