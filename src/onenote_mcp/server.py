"""Servidor MCP de OneNote (escritorio, Windows)."""

from __future__ import annotations

import functools
import time
from typing import Annotated, Any, Awaitable, Callable, TypeVar

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from . import jerarquia as jq
from .com import OneNote, OneNoteError

INSTRUCCIONES = """\
Acceso a los cuadernos de OneNote de escritorio del usuario (Windows).

- Empieza con `listar_cuadernos` o `paginas_recientes` para ubicarte; `buscar` encuentra texto en todas las páginas.
- Las páginas, secciones y cuadernos se pueden indicar por nombre ("Clase 3"), por ruta parcial
  ("Apuntes / Clase 3") o por su ID ("{...}"). Si un nombre es ambiguo, la herramienta lista las opciones.
- OneNote de escritorio solo ve lo que ya sincronizó: si el usuario dice que acaba de apuntar algo en otro
  dispositivo y no aparece, usa `sincronizar` y vuelve a intentarlo.
"""

mcp = MCPServer("onenote", instructions=INSTRUCCIONES)
onenote = OneNote()

SOLO_LECTURA = ToolAnnotations(readOnlyHint=True, openWorldHint=False)

F = TypeVar("F", bound=Callable[..., Awaitable[Any]])


def amable(fn: F) -> F:
    """Convierte los errores esperables en mensajes claros para el modelo (no en un error genérico)."""
    @functools.wraps(fn)
    async def envoltura(*args: Any, **kwargs: Any) -> Any:
        try:
            return await fn(*args, **kwargs)
        except (jq.NoEncontrado, OneNoteError) as e:
            raise ToolError(str(e)) from e
    return envoltura  # type: ignore[return-value]

# ---------- caché de la jerarquía ----------

_CACHE_SEG = 20
_cache: dict[str, object] = {"t": 0.0, "nodos": None}


async def cuadernos(forzar: bool = False) -> list[jq.Nodo]:
    if forzar or _cache["nodos"] is None or time.monotonic() - float(_cache["t"]) > _CACHE_SEG:
        xml = await onenote.llamar(onenote.jerarquia("", "paginas"))
        _cache["nodos"] = jq.parsear(xml)
        _cache["t"] = time.monotonic()
    return _cache["nodos"]  # type: ignore[return-value]


def invalidar_cache() -> None:
    _cache["nodos"] = None


async def resolver(ref: str, tipo: str) -> jq.Nodo:
    try:
        return jq.resolver(await cuadernos(), ref, tipo)
    except jq.NoEncontrado:
        # Quizá se acaba de crear o renombrar: reintentar con la jerarquía fresca
        return jq.resolver(await cuadernos(forzar=True), ref, tipo)


# ---------- herramientas: navegación ----------

@mcp.tool(annotations=SOLO_LECTURA)
@amable
async def listar_cuadernos(
    cuaderno: Annotated[str | None, Field(description="Mostrar solo este cuaderno (nombre o ID). Vacío = todos.")] = None,
    incluir_paginas: Annotated[bool, Field(description="Incluir las páginas de cada sección.")] = True,
) -> str:
    """Muestra el árbol de cuadernos, grupos de secciones, secciones y páginas de OneNote, con sus IDs."""
    nodos = await cuadernos(forzar=True)
    if cuaderno:
        nodos = [jq.resolver(nodos, cuaderno, "cuaderno")]
    return jq.arbol(nodos, con_paginas=incluir_paginas)


@mcp.tool(annotations=SOLO_LECTURA)
@amable
async def paginas_recientes(
    dias: Annotated[int, Field(ge=1, le=3650, description="Páginas modificadas en los últimos N días.")] = 7,
    limite: Annotated[int, Field(ge=1, le=200)] = 20,
) -> str:
    """Lista las páginas modificadas recientemente, de la más nueva a la más antigua."""
    paginas = jq.recientes(await cuadernos(forzar=True), dias, limite)
    if not paginas:
        return f"No hay páginas modificadas en los últimos {dias} días."
    return "\n".join(
        f"- {p.ruta}  ({jq._fmt_fecha(p.modificado)} · id: {p.id})" for p in paginas
    )


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=True))
@amable
async def sincronizar(
    cuaderno: Annotated[str | None, Field(description="Cuaderno a sincronizar (nombre o ID). Vacío = todos.")] = None,
) -> str:
    """Pide a OneNote que sincronice con la nube, para ver apuntes hechos en otros dispositivos.

    La sincronización sigue en segundo plano dentro de OneNote; si algo no aparece aún,
    espera unos segundos y vuelve a consultar.
    """
    nodos = await cuadernos(forzar=True)
    objetivos = [jq.resolver(nodos, cuaderno, "cuaderno")] if cuaderno else nodos
    errores = []
    for c in objetivos:
        try:
            await onenote.llamar(onenote.sincronizar(c.id))
        except OneNoteError as e:
            errores.append(f"{c.nombre}: {e}")
    invalidar_cache()
    nombres = ", ".join(c.nombre for c in objetivos)
    msg = f"Sincronización solicitada para: {nombres}."
    if errores:
        msg += "\nNo se pudo sincronizar:\n" + "\n".join(f"- {e}" for e in errores)
    return msg


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, openWorldHint=False))
@amable
async def abrir_en_onenote(
    elemento: Annotated[str, Field(description="Página o sección (nombre, ruta o ID) que se mostrará en la ventana de OneNote.")],
) -> str:
    """Muestra una página o sección en la ventana de OneNote del usuario."""
    try:
        nodo = await resolver(elemento, "pagina")
    except jq.NoEncontrado:
        nodo = await resolver(elemento, "seccion")
    await onenote.llamar(onenote.navegar(nodo.id))
    return f"Abierto en OneNote: {nodo.ruta}"


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
