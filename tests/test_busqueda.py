"""Pruebas del índice de búsqueda y de la resolución de nombres (sin OneNote)."""

from datetime import datetime, timezone

import pytest

from onenote_mcp import jerarquia as jq
from onenote_mcp.busqueda import Indice

NS = 'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'
JERARQUIA = f"""<?xml version="1.0"?>
<one:Notebooks {NS}>
  <one:Notebook name="Universidad" ID="{{N1}}">
    <one:Section name="Redes" ID="{{S1}}">
      <one:Page ID="{{P1}}" name="Clase 1" lastModifiedTime="2026-01-10T10:00:00.000Z" pageLevel="1"/>
      <one:Page ID="{{P2}}" name="Clase 2" lastModifiedTime="2026-01-17T10:00:00.000Z" pageLevel="1"/>
    </one:Section>
    <one:Section name="Cálculo" ID="{{S2}}">
      <one:Page ID="{{P3}}" name="Clase 1" lastModifiedTime="2026-01-11T10:00:00.000Z" pageLevel="1"/>
    </one:Section>
    <one:SectionGroup name="OneNote_RecycleBin" ID="{{R}}" isRecycleBin="true">
      <one:Section name="Páginas eliminadas" ID="{{S3}}" isInRecycleBin="true">
        <one:Page ID="{{P4}}" name="Borrada" pageLevel="1"/>
      </one:Section>
    </one:SectionGroup>
  </one:Notebook>
</one:Notebooks>"""


@pytest.fixture
def nodos():
    return jq.parsear(JERARQUIA)


def test_papelera_excluida(nodos):
    nombres = [p.nombre for p in jq.todos(nodos, "pagina")]
    assert "Borrada" not in nombres
    assert len(nombres) == 3


def test_resolver_por_nombre_ruta_e_id(nodos):
    assert jq.resolver(nodos, "Clase 2", "pagina").id == "{P2}"
    assert jq.resolver(nodos, "calculo / clase 1", "pagina").id == "{P3}"  # sin tildes ni mayúsculas
    assert jq.resolver(nodos, "{P1}", "pagina").ruta == "Universidad / Redes / Clase 1"
    with pytest.raises(jq.NoEncontrado, match="coincide con 2"):
        jq.resolver(nodos, "Clase 1", "pagina")
    with pytest.raises(jq.NoEncontrado):
        jq.resolver(nodos, "No existe", "pagina")


def test_recientes(nodos):
    paginas = jq.recientes(nodos, dias=100000, limite=2)
    assert [p.id for p in paginas] == ["{P2}", "{P3}"]


def test_busqueda_sin_tildes_y_con_ranking(nodos):
    paginas = list(jq.todos(nodos, "pagina"))
    idx = Indice()
    textos = {
        "{P1}": "El modelo OSI tiene siete capas. La capa de red enruta paquetes.",
        "{P2}": "Subredes y máscaras. Ejercicio de subneteo.",
        "{P3}": "Límites y derivadas. La derivada de x² es 2x. Derivada derivada.",
    }
    for p in paginas:
        assert idx.necesita(p)
        idx.guardar(p, textos[p.id])
        assert not idx.necesita(p)

    r = idx.buscar(paginas, "limites")  # sin tilde encuentra "Límites"
    assert [x.nodo.id for x in r] == ["{P3}"]
    assert "Límites" in r[0].fragmentos[0]

    r = idx.buscar(paginas, "capa red")  # todas las palabras deben estar
    assert [x.nodo.id for x in r] == ["{P1}"]

    assert idx.buscar(paginas, "capa derivada") == []

    # Si la página cambia, se vuelve a indexar
    p1 = paginas[0]
    p1.modificado = datetime(2026, 2, 1, tzinfo=timezone.utc)
    assert idx.necesita(p1)
