"""Nombre de archivos PLE, escritura de TXT y empaquetado ZIP."""
from __future__ import annotations

import io
import zipfile

from .config import Config

ENCODING = "cp1252"


def nombre_archivo(cfg: Config, libro: str, con_informacion: bool) -> str:
    """LE + RUC + AAAA + 00 + 00 + LLLLLL + CC + O + I + M + G .txt"""
    indicador = "1" if con_informacion else "0"
    return (
        f"LE{cfg.ruc}{cfg.ejercicio}0000{libro}{cfg.codigo_oportunidad}"
        f"{cfg.indicador_operaciones}{indicador}{cfg.moneda}1.txt"
    )


def contenido_txt(lineas: list[list[str]]) -> bytes:
    """Campos separados por '|', con '|' final y CRLF. Sin líneas -> 0 bytes."""
    if not lineas:
        return b""
    texto = "".join("|".join(campos) + "|\r\n" for campos in lineas)
    return texto.encode(ENCODING)


def armar_zip(archivos: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, datos in archivos.items():
            z.writestr(nombre, datos)
    return buffer.getvalue()
