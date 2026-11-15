"""Pruebas de la conversión Markdown → XML de OneNote (sin OneNote)."""

import xml.etree.ElementTree as ET

from onenote_mcp.escritura import en_linea, parsear_markdown, xml_agregar, xml_nueva_pagina
from onenote_mcp.lectura import leer_pagina

NSU = "{http://schemas.microsoft.com/office/onenote/2013/onenote}"
NS = 'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'


def test_en_linea():
    assert en_linea("a **b** *c* ~~d~~ `e<f>` [g](https://x.y)") == (
        "a <span style='font-weight:bold'>b</span> <span style='font-style:italic'>c</span> "
        "<span style='text-decoration:line-through'>d</span> "
        "<span style='font-family:Consolas'>e&lt;f&gt;</span> <a href=\"https://x.y\">g</a>"
    )


def test_enlaces_no_web_se_dejan_como_texto():
    assert en_linea("[x](javascript:alert(1))") == "[x](javascript:alert(1))"


def test_estructura_anidada():
    bloques = parsear_markdown("# T\n- a\n  - b\n    texto\n1. n\n- [x] hecha\n\n| h1 | h2 |\n|---|---|\n| 1 | 2 |")
    tipos = [b.tipo for b in bloques]
    assert tipos == ["titulo", "vineta", "numero", "tarea", "tabla"]
    assert bloques[1].hijos[0].texto == "b"
    assert bloques[1].hijos[0].hijos[0].texto == "texto"
    assert bloques[3].hecho
    assert bloques[4].filas == [["h1", "h2"], ["1", "2"]]


def test_pagina_nueva_es_xml_valido_y_se_lee_igual():
    md = "## Sección\nTexto con **negrita** y ]]> raro\n- uno\n- [ ] tarea\n```\ncodigo()\n```"
    xml = xml_nueva_pagina("{P}", "Título <con> símbolos", md)
    raiz = ET.fromstring(xml)  # debe ser XML bien formado
    assert raiz.get("ID") == "{P}"
    assert raiz.find(f"{NSU}TagDef") is not None  # hay una tarea → hace falta su TagDef

    # Lo que se escribe se puede volver a leer
    leido = leer_pagina(xml.replace("<one:OE>", '<one:OE quickStyleIndex="9">'))
    assert "Título <con> símbolos" in leido.titulo
    assert "**negrita**" in leido.markdown
    assert "]]> raro" in leido.markdown
    assert "- uno" in leido.markdown
    assert "[ ] tarea" in leido.markdown
    assert "`codigo()`" in leido.markdown


PAGINA_EXISTENTE = f"""<?xml version="1.0"?>
<one:Page {NS} ID="{{P}}">
  <one:TagDef index="0" type="1" symbol="13" name="Importante"/>
  <one:TagDef index="1" type="0" symbol="3" fontColor="automatic" highlightColor="none" name="Tareas pendientes"/>
  <one:Title><one:OE><one:T><![CDATA[Existente]]></one:T></one:OE></one:Title>
  <one:Outline objectID="{{O1}}"><one:Position x="36" y="86" z="0"/><one:Size width="300" height="40"/>
    <one:OEChildren><one:OE objectID="{{E1}}"><one:T><![CDATA[contenido previo]]></one:T></one:OE></one:OEChildren>
  </one:Outline>
  <one:Image objectID="{{I1}}"><one:Position x="400" y="200" z="1"/><one:Size width="100" height="150"/></one:Image>
</one:Page>"""


def test_agregar_no_reenvia_bloques_existentes():
    xml = xml_agregar(PAGINA_EXISTENTE, "- [ ] nueva tarea\nmás texto")
    raiz = ET.fromstring(xml)
    outlines = raiz.findall(f"{NSU}Outline")
    # Un solo bloque, NUEVO (sin objectID): si llevara el de uno existente, OneNote lo reemplazaría
    assert len(outlines) == 1
    assert outlines[0].get("objectID") is None
    assert "{O1}" not in xml and "{E1}" not in xml and "{I1}" not in xml
    # Colocado debajo de todo (la imagen termina en y=350), alineado con ella
    pos = outlines[0].find(f"{NSU}Position")
    assert float(pos.get("y")) > 350
    assert int(pos.get("z")) == 2
    # Reutiliza la etiqueta de tareas de la página e incluye su TagDef (OneNote lo exige)
    tag = raiz.find(f".//{NSU}Tag")
    assert tag.get("index") == "1"
    tagdefs = raiz.findall(f"{NSU}TagDef")
    assert [t.get("index") for t in tagdefs] == ["1"]


def test_agregar_sin_tareas_no_manda_tagdef():
    raiz = ET.fromstring(xml_agregar(PAGINA_EXISTENTE, "solo texto"))
    assert raiz.findall(f"{NSU}TagDef") == []
