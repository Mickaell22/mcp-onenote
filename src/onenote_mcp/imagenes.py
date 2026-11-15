"""Preparación de las imágenes de OneNote para enviarlas al modelo."""

from __future__ import annotations

import base64
import io

from PIL import Image

LADO_MAX = 1568          # px; más grande no aporta detalle al modelo y gasta contexto
BYTES_PNG_MAX = 600_000  # por encima, se pasa a JPEG


def preparar(b64: str) -> tuple[bytes, str]:
    """Recibe la imagen en base64 tal como la da OneNote y devuelve (bytes, formato)."""
    datos = base64.b64decode(b64)
    img = Image.open(io.BytesIO(datos))
    img.load()
    formato_original = (img.format or "PNG").upper()

    cambio = False
    if max(img.size) > LADO_MAX:
        img.thumbnail((LADO_MAX, LADO_MAX), Image.Resampling.LANCZOS)
        cambio = True

    # Formatos que el modelo no admite directamente (BMP, TIFF, EMF…) o fotos pesadas → JPEG/PNG
    if formato_original in ("JPEG", "JPG") and not cambio:
        return datos, "jpeg"
    if formato_original in ("PNG", "GIF", "WEBP") and not cambio and len(datos) <= BYTES_PNG_MAX:
        return datos, formato_original.lower()

    salida = io.BytesIO()
    if img.mode in ("RGBA", "LA", "P"):
        # Las capturas con transparencia se ven mejor sobre blanco
        img = img.convert("RGBA")
        fondo = Image.new("RGB", img.size, "white")
        fondo.paste(img, mask=img.split()[-1])
        img = fondo
    elif img.mode != "RGB":
        img = img.convert("RGB")

    img.save(salida, "PNG", optimize=True)
    if salida.tell() <= BYTES_PNG_MAX:
        return salida.getvalue(), "png"
    salida = io.BytesIO()
    img.save(salida, "JPEG", quality=85, optimize=True)
    return salida.getvalue(), "jpeg"
