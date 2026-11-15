"""Búsqueda de texto en las páginas (incluido el OCR de las imágenes).

No depende del indexado de Windows: guarda en memoria el texto de cada página y solo
vuelve a leer las que cambiaron desde la última búsqueda.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime

from . import jerarquia as jq


@dataclass
class _Entrada:
    modificado: datetime | None
    texto: str
    normal: str


@dataclass
class Resultado:
    nodo: jq.Nodo
    puntos: int
    fragmentos: list[str]


class Indice:
    def __init__(self) -> None:
        self._paginas: dict[str, _Entrada] = {}

    def necesita(self, nodo: jq.Nodo) -> bool:
        e = self._paginas.get(nodo.id)
        return e is None or e.modificado != nodo.modificado

    def guardar(self, nodo: jq.Nodo, texto: str) -> None:
        self._paginas[nodo.id] = _Entrada(nodo.modificado, texto, jq.normalizar(texto))

    def olvidar_ausentes(self, ids_vigentes: set[str]) -> None:
        for pid in list(self._paginas):
            if pid not in ids_vigentes:
                del self._paginas[pid]

    def buscar(self, paginas: list[jq.Nodo], consulta: str, max_fragmentos: int = 3) -> list[Resultado]:
        frase = jq.normalizar(consulta.strip().strip('"'))
        terminos = [t for t in re.split(r"\s+", frase) if t]
        if not terminos:
            return []
        resultados = []
        for nodo in paginas:
            e = self._paginas.get(nodo.id)
            if e is None:
                continue
            titulo = jq.normalizar(nodo.nombre)
            # Todas las palabras deben aparecer (en el texto o en el título)
            if not all(t in e.normal or t in titulo for t in terminos):
                continue
            puntos = sum(e.normal.count(t) for t in terminos)
            puntos += 20 * sum(t in titulo for t in terminos)
            if len(terminos) > 1 and frase in e.normal:
                puntos += 10
            resultados.append(Resultado(nodo, puntos, _fragmentos(e, terminos, frase, max_fragmentos)))
        resultados.sort(key=lambda r: (-r.puntos, -(r.nodo.modificado.timestamp() if r.nodo.modificado else 0)))
        return resultados


def _fragmentos(e: _Entrada, terminos: list[str], frase: str, maximo: int, radio: int = 70) -> list[str]:
    # normalizar() conserva la longitud de cada carácter salvo tildes combinadas, que
    # no aparecen en texto ya compuesto; si las longitudes difieren, se usa el texto normalizado.
    base = e.texto if len(e.texto) == len(e.normal) else e.normal
    posiciones: list[int] = []
    for objetivo in ([frase] if len(terminos) > 1 else []) + terminos:
        for m in re.finditer(re.escape(objetivo), e.normal):
            if all(abs(m.start() - p) > radio for p in posiciones):
                posiciones.append(m.start())
            if len(posiciones) >= maximo:
                break
        if len(posiciones) >= maximo:
            break
    out = []
    for p in sorted(posiciones):
        ini, fin = max(0, p - radio), min(len(base), p + radio)
        trozo = " ".join(base[ini:fin].split())
        out.append(("…" if ini > 0 else "") + trozo + ("…" if fin < len(base) else ""))
    return out
