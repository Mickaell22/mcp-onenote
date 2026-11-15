# onenote-mcp

Servidor [MCP](https://modelcontextprotocol.io) para leer, buscar y escribir en **OneNote de escritorio (Windows)** desde Claude Code, Claude Desktop u otro cliente MCP.

Se conecta a la app de OneNote instalada en tu PC mediante su interfaz COM:

- **Sin Azure ni Microsoft Graph**: no hay que registrar ninguna app ni iniciar sesión en otro servicio.
- **Local**: tus apuntes solo viajan desde OneNote hasta el cliente MCP que los pide.
- **Ve tus imágenes**: incluye el texto reconocido (OCR) de cada captura y puede adjuntar las imágenes para que el modelo las vea.
- **Seguro al escribir**: solo crea páginas o añade contenido al final; no hay herramientas para borrar ni sobrescribir.

## Requisitos

- Windows con **OneNote de escritorio** (el de Office: OneNote 2013, 2016, 2019, 2021 o Microsoft 365). La app "OneNote para Windows 10" de la Store no tiene interfaz COM.
- Python 3.10 o superior.
- Los cuadernos que quieras usar deben estar **abiertos en OneNote de escritorio** (Archivo → Abrir). Si están en OneDrive, se sincronizan solos.

## Instalación

```powershell
git clone https://github.com/Mickaell22/onenote-mcp.git
cd onenote-mcp
python -m venv .venv
.venv\Scripts\pip install -e .
```

### Claude Code

```powershell
claude mcp add onenote --scope user -- "C:\ruta\a\onenote-mcp\.venv\Scripts\onenote-mcp.exe"
claude mcp list   # debería mostrar: onenote ... ✔ Connected
```

### Claude Desktop u otros clientes

```json
{
  "mcpServers": {
    "onenote": {
      "command": "C:\\ruta\\a\\onenote-mcp\\.venv\\Scripts\\onenote-mcp.exe"
    }
  }
}
```

## Herramientas

| Herramienta | Qué hace |
|---|---|
| `listar_cuadernos` | Árbol de cuadernos, grupos, secciones y páginas (sin la papelera). |
| `paginas_recientes` | Páginas modificadas en los últimos N días. |
| `buscar` | Busca palabras en todas las páginas, **incluido el OCR de las imágenes**. Sin distinguir mayúsculas ni tildes, con fragmentos de contexto. Se puede limitar a un cuaderno o sección. |
| `leer_pagina` | Página en Markdown: títulos, listas, casillas `[ ]`/`[x]`, tablas, enlaces, adjuntos, y cada imagen con su texto reconocido. Opcionalmente adjunta las imágenes. |
| `ver_imagen` | Devuelve imágenes concretas de una página por su número. |
| `crear_pagina` | Crea una página en una sección a partir de Markdown. |
| `agregar_a_pagina` | Añade Markdown al final de una página sin tocar lo que ya tiene. |
| `sincronizar` | Pide a OneNote que sincronice con la nube (para apuntes hechos en el móvil o en la web). |
| `abrir_en_onenote` | Muestra una página o sección en la ventana de OneNote. |

Las páginas, secciones y cuadernos se pueden indicar por **nombre** (`"Clase 3"`), por **ruta parcial** (`"Redes / Clase 3"`) o por **ID**. Si un nombre es ambiguo, la herramienta lista las opciones.

### Markdown admitido al escribir

`# Títulos` (1–6), `- viñetas`, `1. numeradas`, `- [ ] tareas` / `- [x] hechas`, `**negrita**`, `*cursiva*`, `~~tachado~~`, `` `código` ``, bloques ```` ``` ````, `[enlaces](https://…)`, tablas `| a | b |` y sangría de 2 espacios para anidar.

## Ejemplos de uso

- *"¿Qué apunté esta semana en OneNote?"*
- *"Busca en mis apuntes dónde hablé de subneteo y explícamelo."*
- *"Lee la página «Clase 5» y mira las capturas: ¿qué error sale en la segunda?"*
- *"Crea en la sección Redes una página «Resumen parcial» con los puntos clave de las clases 1 a 5."*
- *"Añade a «Pendientes» una tarea: revisar el informe."*

## Detalles técnicos

- **Type library mal registrada**: en muchas instalaciones de Office de 64 bits la type library de OneNote solo está registrada para Win32, y `pywin32` falla con *"Biblioteca no registrada"*. Este servidor la carga directamente desde `ONENOTE.EXE` con `comtypes`, que llama por vtable.
- **Un solo hilo COM**: todas las llamadas a OneNote pasan por un hilo dedicado; si OneNote se cierra, la siguiente llamada reconecta.
- **Imágenes de cuadernos en la nube**: OneNote no las descarga hasta que se abre la página. Si una imagen aún no está en el equipo, el servidor abre esa página un momento, espera la descarga y **vuelve a la página que tenías abierta**.
- **Añadir contenido**: enviar a OneNote un bloque existente (con su `objectID`) hace que lo **reemplace entero**. Por eso `agregar_a_pagina` crea un bloque nuevo colocado debajo de todo lo que ya hay.
- **Búsqueda**: no depende del indexado de Windows. Mantiene en memoria el texto de cada página (y su OCR) y solo vuelve a leer las que cambiaron.

## Desarrollo

```powershell
.venv\Scripts\pip install -e ".[dev]"
.venv\Scripts\python -m pytest
```

Las pruebas usan XML de ejemplo y no necesitan OneNote.

## Licencia

[MIT](LICENSE)
