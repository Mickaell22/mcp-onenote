"""Estructura de cuadernos → secciones → páginas, y resolución de nombres a IDs."""

from __future__ import annotations

import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime, timezone


@dataclass
class Nodo:
    tipo: str            # "cuaderno" | "grupo" | "seccion" | "pagina"
    id: str
    nombre: str
    ruta: str            # "Cuaderno / Sección / Página"
    modificado: datetime | None = None
    nivel_pagina: int = 1  # subpáginas: 2 o 3
    hijos: list["Nodo"] = field(default_factory=list)
    padre: "Nodo | None" = field(default=None, repr=False)
    bloqueada: bool = False  # sección protegida con contraseña


def _ns(raiz: ET.Element) -> str:
    return raiz.tag.split("}")[0] + "}" if raiz.tag.startswith("{") else ""


def _fecha(valor: str | None) -> datetime | None:
    if not valor:
        return None
    try:
        return datetime.fromisoformat(valor.replace("Z", "+00:00"))
    except ValueError:
        return None


def _es_papelera(el: ET.Element) -> bool:
    return el.get("isRecycleBin") == "true" or el.get("isInRecycleBin") == "true"


def parsear(xml: str) -> list[Nodo]:
    """Convierte el XML de GetHierarchy en una lista de cuadernos (sin papelera)."""
    raiz = ET.fromstring(xml)
    ns = _ns(raiz)

    def recorrer(el: ET.Element, padre: Nodo | None) -> Nodo | None:
        etiqueta = el.tag[len(ns):]
        tipo = {"Notebook": "cuaderno", "SectionGroup": "grupo",
                "Section": "seccion", "Page": "pagina"}.get(etiqueta)
        if tipo is None or _es_papelera(el):
            return None
        # OneNote guarda los saltos de línea de los nombres como "^J"
        nombre = " ".join((el.get("nickname") or el.get("name") or "(sin título)").replace("^J", " ").split())
        if tipo == "pagina" and not el.get("name"):
            nombre = "(página sin título)"
        ruta = f"{padre.ruta} / {nombre}" if padre else nombre
        nodo = Nodo(
            tipo=tipo, id=el.get("ID", ""), nombre=nombre, ruta=ruta,
            modificado=_fecha(el.get("lastModifiedTime")),
            nivel_pagina=int(el.get("pageLevel", "1") or 1),
            padre=padre,
            bloqueada=el.get("locked") == "true",
        )
        for hijo in el:
            h = recorrer(hijo, nodo)
            if h:
                nodo.hijos.append(h)
        return nodo

    return [n for el in raiz if (n := recorrer(el, None))]


def todos(cuadernos: list[Nodo], tipo: str | None = None):
    pila = list(reversed(cuadernos))
    while pila:
        n = pila.pop()
        if tipo is None or n.tipo == tipo:
            yield n
        pila.extend(reversed(n.hijos))


def normalizar(texto: str) -> str:
    """Minúsculas y sin tildes, para comparar nombres y buscar texto."""
    t = unicodedata.normalize("NFD", texto.casefold())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


class NoEncontrado(LookupError):
    pass


def resolver(cuadernos: list[Nodo], ref: str, tipo: str) -> Nodo:
    """Encuentra un nodo por ID o por nombre.

    El nombre puede ser solo el título ("Clase 3") o una ruta parcial separada por "/"
    ("Apuntes / Clase 3"). Si hay varias coincidencias se pide precisar.
    """
    ref = ref.strip()
    candidatos = list(todos(cuadernos, tipo))
    if ref.startswith("{"):
        for n in candidatos:
            if n.id == ref:
                return n
        raise NoEncontrado(f"No hay ningún elemento de tipo '{tipo}' con el ID {ref}.")

    partes = [normalizar(p.strip()) for p in ref.split("/") if p.strip()]
    if not partes:
        raise NoEncontrado("Indica un nombre o un ID.")

    def coincide(n: Nodo, exacto: bool) -> bool:
        ruta = [normalizar(p.strip()) for p in n.ruta.split(" / ")]
        # El último trozo debe ser el nombre del nodo; los anteriores, ancestros en orden
        ultimo = partes[-1]
        if exacto and ruta[-1] != ultimo:
            return False
        if not exacto and ultimo not in ruta[-1]:
            return False
        i = 0
        for p in ruta[:-1]:
            if i < len(partes) - 1 and partes[i] in p:
                i += 1
        return i == len(partes) - 1

    for exacto in (True, False):
        hallados = [n for n in candidatos if coincide(n, exacto)]
        if len(hallados) == 1:
            return hallados[0]
        if len(hallados) > 1:
            lista = "\n".join(f"  - {n.ruta}  (id: {n.id})" for n in hallados[:15])
            raise NoEncontrado(
                f"'{ref}' coincide con {len(hallados)} elementos de tipo '{tipo}'. "
                f"Precisa con la ruta (p. ej. 'Cuaderno / Sección / Página') o usa el ID:\n{lista}"
            )
    raise NoEncontrado(f"No se encontró ningún elemento de tipo '{tipo}' llamado '{ref}'.")


def _fmt_fecha(d: datetime | None) -> str:
    if not d:
        return ""
    return d.astimezone().strftime("%Y-%m-%d %H:%M")


def arbol(cuadernos: list[Nodo], con_paginas: bool = True, con_ids: bool = True) -> str:
    lineas: list[str] = []
    iconos = {"cuaderno": "📓", "grupo": "📁", "seccion": "📑", "pagina": "📄"}

    def pintar(n: Nodo, sangria: int) -> None:
        if n.tipo == "pagina" and not con_paginas:
            return
        extra = []
        if n.tipo == "seccion":
            paginas = sum(1 for h in n.hijos if h.tipo == "pagina")
            extra.append(f"{paginas} pág.")
            if n.bloqueada:
                extra.append("🔒 bloqueada")
        if n.tipo == "pagina" and n.modificado:
            extra.append(_fmt_fecha(n.modificado))
        if con_ids:
            extra.append(f"id: {n.id}")
        sub = "  " * (n.nivel_pagina - 1) if n.tipo == "pagina" else ""
        detalle = f"  ({' · '.join(extra)})" if extra else ""
        lineas.append(f"{'  ' * sangria}{sub}{iconos[n.tipo]} {n.nombre}{detalle}")
        for h in n.hijos:
            pintar(h, sangria + 1)

    for c in cuadernos:
        pintar(c, 0)
    return "\n".join(lineas) if lineas else "No hay cuadernos abiertos en OneNote."


def recientes(cuadernos: list[Nodo], dias: int, limite: int) -> list[Nodo]:
    ahora = datetime.now(timezone.utc)
    paginas = [p for p in todos(cuadernos, "pagina")
               if p.modificado and (ahora - p.modificado).total_seconds() <= dias * 86400]
    paginas.sort(key=lambda p: p.modificado, reverse=True)
    return paginas[:limite]
