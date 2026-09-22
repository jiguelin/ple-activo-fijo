"""Funciones auxiliares: números, fechas y limpieza de texto."""
from __future__ import annotations

import datetime as dt
import re
import unicodedata
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

CERO = Decimal("0")
DOS_DEC = Decimal("0.01")
TRES_DEC = Decimal("0.001")


def normalizar(texto) -> str:
    """Mayúsculas, sin tildes y con espacios simples (para comparar títulos)."""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().upper()


def a_decimal(valor) -> Decimal | None:
    """Convierte un valor de Excel en Decimal. Vacío -> None. Texto inválido -> ValueError."""
    if valor is None:
        return None
    if isinstance(valor, bool):
        raise ValueError(f"valor no numérico: {valor!r}")
    if isinstance(valor, (int, Decimal)):
        return Decimal(valor)
    if isinstance(valor, float):
        return Decimal(repr(valor))
    texto = str(valor).strip().replace(" ", "")
    if texto in ("", "-"):
        return None
    # Admite "1,234.56" y "1234,56"
    if "," in texto and "." in texto:
        texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return Decimal(texto)
    except InvalidOperation as exc:
        raise ValueError(f"valor no numérico: {valor!r}") from exc


def fmt_importe(valor: Decimal | None) -> str:
    """Importe con 2 decimales, punto decimal, sin miles. Vacío -> 0.00."""
    if valor is None:
        valor = CERO
    q = valor.quantize(DOS_DEC, rounding=ROUND_HALF_UP)
    if q == 0:
        q = abs(q)  # evita "-0.00"
    return f"{q:.2f}"


def fmt_tc(valor: Decimal | None) -> str:
    """Tipo de cambio con 3 decimales."""
    if valor is None:
        valor = CERO
    return f"{valor.quantize(TRES_DEC, rounding=ROUND_HALF_UP):.3f}"


def redondear(valor: Decimal | None) -> Decimal:
    return (valor or CERO).quantize(DOS_DEC, rounding=ROUND_HALF_UP)


def a_fecha(valor) -> dt.date | None:
    """Convierte fecha de Excel / datetime / texto en date. Vacío -> None."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    if isinstance(valor, (int, float)):
        # número de serie de Excel
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(valor))).date()
    texto = str(valor).strip()
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(texto, formato).date()
        except ValueError:
            continue
    raise ValueError(f"fecha no válida: {valor!r}")


def fmt_fecha(fecha: dt.date) -> str:
    return fecha.strftime("%d/%m/%Y")


def a_texto(valor) -> str:
    """Convierte a texto sin '.0' en números enteros (ej. modelo 2017, serie 202109252)."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if isinstance(valor, (dt.datetime, dt.date)):
        return valor.strftime("%d/%m/%Y")
    return str(valor).strip()


# Caracteres que no deben ir en el TXT
_REEMPLAZOS = {
    "|": "/", "\t": " ", "\r": " ", "\n": " ",
    "“": '"', "”": '"', "‘": "'", "’": "'", "–": "-", "—": "-", "…": "...",
}


def limpiar_texto(valor) -> str:
    """Texto listo para el TXT: mayúsculas, sin '|', sin saltos, solo caracteres de cp1252."""
    texto = a_texto(valor)
    for a, b in _REEMPLAZOS.items():
        texto = texto.replace(a, b)
    salida = []
    for c in texto:
        if not c.isprintable():
            continue
        try:
            c.encode("cp1252")
            salida.append(c)
        except UnicodeEncodeError:
            base = unicodedata.normalize("NFKD", c)
            base = "".join(x for x in base if not unicodedata.combining(x))
            try:
                base.encode("cp1252")
                salida.append(base)
            except UnicodeEncodeError:
                continue
    return re.sub(r"\s+", " ", "".join(salida)).strip().upper()


def ruc_valido(ruc: str) -> bool:
    """Valida longitud, prefijo y dígito verificador del RUC."""
    if not re.fullmatch(r"\d{11}", ruc or ""):
        return False
    if ruc[:2] not in ("10", "15", "16", "17", "20"):
        return False
    factores = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    suma = sum(int(d) * f for d, f in zip(ruc[:10], factores))
    digito = 11 - (suma % 11)
    if digito == 10:
        digito = 0
    elif digito == 11:
        digito = 1
    return digito == int(ruc[10])
