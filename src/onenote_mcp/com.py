"""Conexión con OneNote de escritorio mediante su interfaz COM.

Detalles que resuelve este módulo:

* En muchas instalaciones de Office de 64 bits la "type library" de OneNote solo está
  registrada para Win32, y pywin32 falla con "Biblioteca no registrada". Aquí se carga la
  type library directamente desde el archivo (``ONENOTE.EXE\\3``) con comtypes, que llama
  por vtable y no depende de ese registro.
* COM exige que el objeto se use desde el hilo que lo creó. Todas las llamadas pasan por
  un único hilo dedicado.
* Si cierras OneNote mientras el servidor corre, la siguiente llamada reconecta sola.
"""

from __future__ import annotations

import asyncio
import os
import threading
import winreg
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable, TypeVar

T = TypeVar("T")

# Microsoft OneNote 15.0 Type Library (válida para OneNote 2013, 2016, 2019, 2021 y Microsoft 365)
TYPELIB_GUID = "{0EA692EE-BB50-4E3C-AEF0-356D91732725}"

# Errores COM que indican que OneNote se cerró o se reinició
_DESCONEXION = {
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # RPC_E_DISCONNECTED (llamada remota falló)
    -2147417848,  # RPC_E_DISCONNECTED
    -2147417851,  # RPC_E_SERVERFAULT
}


class OneNoteError(RuntimeError):
    """Error con un mensaje pensado para mostrarse tal cual al usuario."""


def _ruta_typelib() -> str:
    """Busca en el registro la ruta de la type library de OneNote (la versión más alta)."""
    candidatos: list[str] = []
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"TypeLib\{TYPELIB_GUID}") as raiz:
            versiones = []
            i = 0
            while True:
                try:
                    versiones.append(winreg.EnumKey(raiz, i))
                except OSError:
                    break
                i += 1
            for ver in sorted(versiones, key=_clave_version, reverse=True):
                for plataforma in ("win64", "win32"):
                    try:
                        with winreg.OpenKey(raiz, rf"{ver}\0\{plataforma}") as k:
                            candidatos.append(winreg.QueryValue(k, None))
                    except OSError:
                        pass
    except OSError:
        pass

    # Rutas habituales por si el registro no ayuda
    for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        for office in ("Office16", "Office15"):
            candidatos.append(rf"{base}\Microsoft Office\root\{office}\ONENOTE.EXE\3")
            candidatos.append(rf"{base}\Microsoft Office\{office}\ONENOTE.EXE\3")

    for c in candidatos:
        # La ruta puede llevar un índice de recurso al final ("...\\ONENOTE.EXE\\3")
        base, _, ultimo = c.rpartition("\\")
        archivo = base if ultimo.isdigit() else c
        if os.path.exists(archivo):
            return c
    raise OneNoteError(
        "No se encontró OneNote de escritorio. Este servidor necesita la app de OneNote de "
        "Office (OneNote 2013 o posterior / Microsoft 365) instalada en Windows."
    )


def _clave_version(v: str) -> tuple[int, ...]:
    try:
        return tuple(int(p, 16) for p in v.split("."))
    except ValueError:
        return (0,)


class OneNote:
    """Envoltorio de la aplicación COM de OneNote, seguro para usar desde asyncio."""

    def __init__(self) -> None:
        self._ejecutor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="onenote-com",
                                            initializer=self._iniciar_hilo)
        self._mod: Any = None
        self._app: Any = None
        self._hilo: threading.Thread | None = None

    # ----- hilo COM -----

    @staticmethod
    def _iniciar_hilo() -> None:
        import comtypes
        comtypes.CoInitialize()

    def _conectar(self) -> None:
        import comtypes.client as cc
        import comtypes.typeinfo as ti

        if self._mod is None:
            tlib = ti.LoadTypeLibEx(_ruta_typelib())
            self._mod = cc.GetModule(tlib)
        try:
            self._app = cc.CreateObject(self._mod.Application, interface=self._mod.IApplication)
        except OSError as e:
            raise OneNoteError(f"No se pudo abrir OneNote: {e}") from e

    def _ejecutar(self, fn: Callable[[Any, Any], T]) -> T:
        if self._app is None:
            self._conectar()
        try:
            return fn(self._app, self._mod)
        except Exception as e:  # noqa: BLE001
            if _hresult(e) in _DESCONEXION:
                # OneNote se cerró: reconectar una vez y reintentar
                self._app = None
                self._conectar()
                return fn(self._app, self._mod)
            raise OneNoteError(_mensaje_error(e)) from e

    async def llamar(self, fn: Callable[[Any, Any], T]) -> T:
        """Ejecuta fn(app, mod) en el hilo COM y devuelve su resultado."""
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(self._ejecutor, self._ejecutar, fn)

    def llamar_sync(self, fn: Callable[[Any, Any], T]) -> T:
        """Versión síncrona de llamar() (para pruebas y scripts)."""
        return self._ejecutor.submit(self._ejecutar, fn).result()

    # ----- operaciones -----

    def jerarquia(self, inicio: str = "", alcance: str = "paginas") -> Callable[[Any, Any], str]:
        def fn(app: Any, mod: Any) -> str:
            scope = {"cuadernos": mod.hsNotebooks, "secciones": mod.hsSections,
                     "paginas": mod.hsPages, "hijos": mod.hsChildren, "propio": mod.hsSelf}[alcance]
            return app.GetHierarchy(inicio, scope, mod.xs2013)
        return fn

    @staticmethod
    def contenido(pagina_id: str, binario: bool = False) -> Callable[[Any, Any], str]:
        def fn(app: Any, mod: Any) -> str:
            info = mod.piBinaryData if binario else mod.piBasic
            return app.GetPageContent(pagina_id, info, mod.xs2013)
        return fn

    @staticmethod
    def actualizar_pagina(xml: str) -> Callable[[Any, Any], None]:
        def fn(app: Any, mod: Any) -> None:
            app.UpdatePageContent(xml, 0, mod.xs2013, False)
        return fn

    @staticmethod
    def nueva_pagina(seccion_id: str) -> Callable[[Any, Any], str]:
        def fn(app: Any, mod: Any) -> str:
            return app.CreateNewPage(seccion_id, mod.npsDefault)
        return fn

    @staticmethod
    def sincronizar(objeto_id: str) -> Callable[[Any, Any], None]:
        def fn(app: Any, mod: Any) -> None:
            app.SyncHierarchy(objeto_id)
        return fn

    @staticmethod
    def navegar(objeto_id: str) -> Callable[[Any, Any], None]:
        def fn(app: Any, mod: Any) -> None:
            app.NavigateTo(objeto_id, "", False)
        return fn


def _hresult(e: BaseException) -> int | None:
    for attr in ("hresult", "winerror"):
        v = getattr(e, attr, None)
        if isinstance(v, int):
            return v
    if e.args and isinstance(e.args[0], int):
        return e.args[0]
    return None


# Códigos de error de OneNote más comunes, en lenguaje claro
_ERRORES_ONENOTE = {
    0x80042000: "El XML enviado a OneNote está mal formado.",
    0x80042001: "El XML enviado a OneNote no es válido.",
    0x80042004: "La sección no existe.",
    0x80042005: "La página no existe (quizá se borró o el ID ya no es válido).",
    0x8004200B: "La sección es de solo lectura.",
    0x8004200C: "La página es de solo lectura.",
    0x80042010: "La página cambió en OneNote mientras se editaba; vuelve a leerla e inténtalo de nuevo.",
    0x80042014: "El objeto no existe (quizá se borró o el ID ya no es válido).",
    0x80042015: "El cuaderno no existe o no está abierto en OneNote.",
    0x8004201B: "La sección está protegida con contraseña y bloqueada. Desbloquéala en OneNote.",
    0x8004201D: "El contenido aún no se ha sincronizado. Abre OneNote y espera a que termine de sincronizar.",
    0x8004201F: "Error de red al acceder al cuaderno.",
}


def _mensaje_error(e: BaseException) -> str:
    h = _hresult(e)
    if h is not None and (h & 0xFFFFFFFF) in _ERRORES_ONENOTE:
        return _ERRORES_ONENOTE[h & 0xFFFFFFFF]
    return f"Error de OneNote ({h if h is not None else type(e).__name__}): {e}"
