"""Pruebas del conversor de páginas a Markdown, con XML de ejemplo (sin OneNote)."""

from onenote_mcp.lectura import html_a_md, leer_pagina

NS = 'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'


def pagina(cuerpo: str, extra: str = "") -> str:
    return f"""<?xml version="1.0"?>
<one:Page {NS} ID="{{P}}" name="Ejemplo">
  <one:TagDef index="0" type="0" symbol="3" name="Tareas pendientes"/>
  <one:TagDef index="1" type="1" symbol="13" name="Importante"/>
  <one:QuickStyleDef index="0" name="PageTitle"/>
  <one:QuickStyleDef index="1" name="p"/>
  <one:QuickStyleDef index="2" name="h1"/>
  {extra}
  <one:Title><one:OE quickStyleIndex="0"><one:T><![CDATA[Clase 3: <span style='font-weight:bold'>Redes</span>]]></one:T></one:OE></one:Title>
  {cuerpo}
</one:Page>"""


def outline(oes: str, y: float = 100) -> str:
    return f"""<one:Outline><one:Position x="36" y="{y}" z="0"/><one:OEChildren>{oes}</one:OEChildren></one:Outline>"""


def test_formato_en_linea():
    md = html_a_md("Texto <span style='font-weight:bold'>en negrita</span> y "
                   "<span style='font-style:italic'>cursiva</span> con "
                   "<a href=\"https://ejemplo.com\">enlace</a>.")
    assert md == "Texto **en negrita** y *cursiva* con [enlace](https://ejemplo.com)."


def test_espacios_fuera_de_los_marcadores():
    assert html_a_md("<span style='font-weight:bold'>hola </span>mundo") == "**hola** mundo"


def test_entidades_y_saltos():
    assert html_a_md("a &amp; b<br>c") == "a & b\nc"


def test_titulo_encabezados_y_listas():
    p = leer_pagina(pagina(outline("""
      <one:OE quickStyleIndex="2"><one:T><![CDATA[Tema]]></one:T></one:OE>
      <one:OE><one:T><![CDATA[Intro]]></one:T>
        <one:OEChildren>
          <one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[uno]]></one:T></one:OE>
          <one:OE><one:List><one:Number numberSequence="0" numberFormat="##." text="1."/></one:List><one:T><![CDATA[dos]]></one:T></one:OE>
        </one:OEChildren>
      </one:OE>""")))
    assert p.titulo == "Clase 3: Redes"
    assert p.markdown.startswith("# Clase 3: **Redes**") or p.markdown.startswith("# Clase 3: Redes")
    assert "## Tema" in p.markdown
    assert "  - uno" in p.markdown
    assert "  1. dos" in p.markdown


def test_casillas_y_etiquetas():
    p = leer_pagina(pagina(outline("""
      <one:OE><one:Tag index="0" completed="true"/><one:T><![CDATA[hecha]]></one:T></one:OE>
      <one:OE><one:Tag index="0" completed="false"/><one:T><![CDATA[pendiente]]></one:T></one:OE>
      <one:OE><one:Tag index="1"/><one:T><![CDATA[ojo]]></one:T></one:OE>""")))
    assert "[x] hecha" in p.markdown
    assert "[ ] pendiente" in p.markdown
    assert "[Importante] ojo" in p.markdown


def test_tabla():
    p = leer_pagina(pagina(outline("""
      <one:OE><one:Table>
        <one:Row>
          <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[A]]></one:T></one:OE></one:OEChildren></one:Cell>
          <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[B|C]]></one:T></one:OE></one:OEChildren></one:Cell>
        </one:Row>
        <one:Row>
          <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[1]]></one:T></one:OE></one:OEChildren></one:Cell>
          <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[2]]></one:T></one:OE>
                                    <one:OE><one:T><![CDATA[3]]></one:T></one:OE></one:OEChildren></one:Cell>
        </one:Row>
      </one:Table></one:OE>""")))
    assert "| A | B\\|C |" in p.markdown
    assert "|---|---|" in p.markdown
    assert "| 1 | 2<br>3 |" in p.markdown


def test_imagenes_con_ocr_y_orden_por_posicion():
    img = """<one:Image alt="captura"><one:Position x="10" y="{y}" z="1"/><one:Size width="100" height="50"/>
             <one:CallbackID callbackID="{{CB{n}}}"/>
             <one:OCRData><one:OCRText><![CDATA[texto  de la
             imagen {n}]]></one:OCRText></one:OCRData></one:Image>"""
    p = leer_pagina(pagina(
        outline("<one:OE><one:T><![CDATA[medio]]></one:T></one:OE>", y=200)
        + img.format(y=400, n=2) + img.format(y=50, n=1)
    ))
    # Se leen de arriba abajo: imagen (y=50), texto (y=200), imagen (y=400)
    i1 = p.markdown.index("Imagen 1")
    medio = p.markdown.index("medio")
    i2 = p.markdown.index("Imagen 2")
    assert i1 < medio < i2
    assert "texto reconocido: «texto de la imagen 1»" in p.markdown
    assert "descripción: captura" in p.markdown
    assert [i.callback_id for i in p.imagenes] == ["{CB1}", "{CB2}"]
    assert "imagen 2" in p.texto_plano


def test_adjuntos_y_tinta():
    p = leer_pagina(pagina(outline("""
      <one:OE><one:InsertedFile pathCache="x" preferredName="tarea.pdf"/></one:OE>
      <one:OE><one:InkParagraph><one:InkWord recognizedText="hola"/></one:InkParagraph></one:OE>""")))
    assert "[📎 Archivo adjunto: tarea.pdf]" in p.markdown
    assert "escrito a mano: «hola»" in p.markdown
