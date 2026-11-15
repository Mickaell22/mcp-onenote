"""Servidor MCP de OneNote (escritorio, Windows)."""

from __future__ import annotations

import asyncio
import functools
import time
from typing import Annotated, Any, Awaitable, Callable, TypeVar

from mcp.server.mcpserver import Image, MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from . import imagenes, lectura
from . import jerarquia as jq
from .com import BINARIO_NO_DESCARGADO, OneNote, OneNoteError, es_codigo

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


# ---------- herramientas: lectura ----------

MAX_BYTES_IMAGENES = 4_000_000  # tope total por respuesta, para no saturar el contexto


async def _cargar_pagina(ref: str) -> tuple[jq.Nodo, lectura.Pagina]:
    nodo = await resolver(ref, "pagina")
    xml = await onenote.llamar(onenote.contenido(nodo.id))
    return nodo, lectura.leer_pagina(xml)


async def _descargar_pagina(nodo: jq.Nodo) -> None:
    """En cuadernos de la nube OneNote no baja las imágenes hasta que se abre la página.

    Se abre la página un momento en la ventana de OneNote y luego se vuelve a la que el
    usuario tenía abierta.
    """
    anterior = await onenote.llamar(onenote.pagina_actual())
    await onenote.llamar(onenote.navegar(nodo.id))
    _descargadas.add(nodo.id)
    if anterior and anterior != nodo.id:
        _volver_a[nodo.id] = anterior


_descargadas: set[str] = set()
_volver_a: dict[str, str] = {}


async def _volver(nodo: jq.Nodo) -> None:
    anterior = _volver_a.pop(nodo.id, None)
    if anterior:
        try:
            await onenote.llamar(onenote.navegar(anterior))
        except OneNoteError:
            pass


async def _binario(nodo: jq.Nodo, callback_id: str) -> str:
    try:
        return await onenote.llamar(onenote.binario(nodo.id, callback_id))
    except OneNoteError as e:
        if not es_codigo(e, BINARIO_NO_DESCARGADO):
            raise
    if nodo.id not in _descargadas:
        await _descargar_pagina(nodo)
    # Esperar a que OneNote termine de bajar la imagen (hasta ~15 s)
    ultimo: Exception | None = None
    for _ in range(30):
        await asyncio.sleep(0.5)
        try:
            return await onenote.llamar(onenote.binario(nodo.id, callback_id))
        except OneNoteError as e:
            if not es_codigo(e, BINARIO_NO_DESCARGADO):
                raise
            ultimo = e
    raise OneNoteError(
        "La imagen todavía no se ha descargado de la nube. Comprueba que OneNote tenga conexión "
        "e inténtalo de nuevo en unos segundos."
    ) from ultimo


async def _imagen(nodo: jq.Nodo, img: lectura.ImagenRef) -> Image:
    if not img.callback_id:
        raise OneNoteError(f"La imagen {img.numero} no tiene datos accesibles.")
    b64 = await _binario(nodo, img.callback_id)
    datos, formato = imagenes.preparar(b64)
    return Image(data=datos, format=formato)


@mcp.tool(annotations=SOLO_LECTURA, structured_output=False)
@amable
async def leer_pagina(
    pagina: Annotated[str, Field(description="Página: nombre, ruta parcial ('Sección / Página') o ID.")],
    incluir_imagenes: Annotated[bool, Field(
        description="Adjuntar las imágenes para verlas. Si es False solo se incluye su texto reconocido (OCR), "
                    "que suele bastar para capturas de texto.")] = False,
    max_imagenes: Annotated[int, Field(ge=0, le=30, description="Máximo de imágenes a adjuntar.")] = 6,
) -> list:
    """Lee una página de OneNote y la devuelve en Markdown (texto, listas, tablas, casillas y enlaces).

    Las imágenes aparecen como "[🖼 Imagen N — texto reconocido: …]". Para verlas de verdad usa
    incluir_imagenes=True o `ver_imagen` con su número.
    """
    nodo, pag = await _cargar_pagina(pagina)
    cabecera = f"Ruta: {nodo.ruta}\nModificada: {jq._fmt_fecha(nodo.modificado)}\nID: {nodo.id}\n\n"
    salida: list = [cabecera + pag.markdown]

    if incluir_imagenes and pag.imagenes:
        enviadas = await _adjuntar(nodo, pag.imagenes[:max_imagenes], salida)
        faltan = len(pag.imagenes) - enviadas
        if faltan > 0:
            salida.append(f"({faltan} imagen(es) sin adjuntar; usa `ver_imagen` con su número para verlas.)")
    elif pag.imagenes:
        salida.append(f"(La página tiene {len(pag.imagenes)} imagen(es); usa incluir_imagenes=True o `ver_imagen` para verlas.)")
    return salida


async def _adjuntar(nodo: jq.Nodo, refs: list[lectura.ImagenRef], salida: list) -> int:
    """Añade las imágenes a la salida respetando el tope de tamaño. Devuelve cuántas se adjuntaron."""
    total = enviadas = 0
    try:
        for img in refs:
            try:
                im = await _imagen(nodo, img)
            except Exception as e:  # noqa: BLE001 - una imagen rota no debe tumbar la respuesta
                salida.append(f"(Imagen {img.numero}: no se pudo obtener — {e})")
                continue
            tam = len(im.data or b"")
            if total + tam > MAX_BYTES_IMAGENES:
                salida.append(f"(Límite de tamaño alcanzado en la imagen {img.numero}; pídela por separado con `ver_imagen`.)")
                break
            total += tam
            enviadas += 1
            salida += [f"Imagen {img.numero}:", im]
    finally:
        await _volver(nodo)  # dejar OneNote en la página que tenía el usuario
    return enviadas


@mcp.tool(annotations=SOLO_LECTURA, structured_output=False)
@amable
async def ver_imagen(
    pagina: Annotated[str, Field(description="Página: nombre, ruta parcial o ID.")],
    numeros: Annotated[list[int], Field(description="Números de imagen, tal como aparecen en leer_pagina (Imagen 1, 2…).")],
) -> list:
    """Devuelve una o varias imágenes concretas de una página para verlas."""
    nodo, pag = await _cargar_pagina(pagina)
    por_num = {i.numero: i for i in pag.imagenes}
    salida: list = []
    validas = []
    for n in numeros:
        if n in por_num:
            validas.append(por_num[n])
        else:
            salida.append(f"La página no tiene imagen {n} (tiene {len(pag.imagenes)}).")
    if validas:
        await _adjuntar(nodo, validas, salida)
    return salida


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
