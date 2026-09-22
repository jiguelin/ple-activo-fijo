"""Formato 7.3 - Detalle de la diferencia de cambio (15 campos).

Por defecto se genera VACÍO. Solo se llena si el usuario activa el módulo manual y
marca activos adquiridos en moneda extranjera.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal

from . import config as C
from .config import Config
from .formato_71 import Registro71
from .observaciones import Observaciones
from .utilidades import CERO, a_decimal, fmt_fecha, fmt_importe, fmt_tc

TITULOS_73 = [
    "1 Periodo", "2 CUO", "3 Correlativo asiento", "4 Catálogo", "5 Código activo",
    "6 Fecha adquisición", "7 Valor adq. ME", "8 TC adquisición", "9 Valor adq. MN",
    "10 TC al 31.12", "11 Ajuste dif. cambio", "12 Dep. ejercicio", "13 Dep. retiros/bajas",
    "14 Dep. otros ajustes", "15 Estado operación",
]


@dataclass
class DatoME:
    """Datos que ingresa el usuario para un activo comprado en moneda extranjera."""
    codigo: str
    valor_me: Decimal | float | str | None
    tc_adquisicion: Decimal | float | str | None
    ajuste: Decimal | float | str | None = None
    dep_ejercicio: Decimal | float | str | None = None   # None = tomar la del 7.1


def construir_73(
    registros_71: list[Registro71],
    datos_me: list[DatoME],
    tc_cierre,
    cfg: Config,
) -> tuple[list[list[str]], Observaciones]:
    obs = Observaciones()
    lineas: list[list[str]] = []
    por_codigo = {r.codigo: r for r in registros_71}
    try:
        tc_31 = a_decimal(tc_cierre)
    except ValueError:
        obs.error("Tipo de cambio al 31.12 no válido.", campo="TC al 31.12")
        tc_31 = None
    if datos_me and tc_31 is None:
        obs.advertencia("No se ingresó el tipo de cambio al 31.12 → 0.000.", campo="TC al 31.12")

    for n, dato in enumerate(datos_me, start=1):
        reg = por_codigo.get(dato.codigo)
        if reg is None:
            obs.error("El código no existe en el 7.1.", codigo=dato.codigo, campo="Código")
            continue
        try:
            valor_me = a_decimal(dato.valor_me)
            tc_adq = a_decimal(dato.tc_adquisicion)
            ajuste = a_decimal(dato.ajuste) or CERO
            dep_manual = a_decimal(dato.dep_ejercicio)
        except ValueError as exc:
            obs.error(str(exc), reg.fila, reg.codigo, "7.3")
            continue
        if not valor_me:
            obs.error("Falta el valor de adquisición en moneda extranjera.", reg.fila, reg.codigo, "Valor ME")
        if not tc_adq:
            obs.error("Falta el tipo de cambio de la fecha de adquisición.", reg.fila, reg.codigo, "TC adquisición")
        imp = reg.importes
        valor_mn = imp["saldo_inicial"] + imp["adquisiciones"]
        if valor_me and tc_adq and abs(valor_me * tc_adq - valor_mn) > Decimal("1.00"):
            obs.advertencia(
                f"Valor ME × TC ({fmt_importe(valor_me * tc_adq)}) difiere del costo en soles del 7.1 "
                f"({fmt_importe(valor_mn)}).", reg.fila, reg.codigo, "Valor MN")
        dep = imp["dep_ejercicio"]
        if dep_manual is not None and dep_manual != dep:
            obs.advertencia(
                f"Depreciación del ejercicio ingresada ({fmt_importe(dep_manual)}) distinta a la del 7.1 "
                f"({fmt_importe(dep)}).", reg.fila, reg.codigo, "Dep. ejercicio")
            dep = dep_manual
        lineas.append([
            cfg.periodo,
            reg.campos[1],   # mismo CUO / correlativo del activo en el 7.1
            reg.campos[2],
            C.CATALOGO_OTROS,
            reg.codigo,
            fmt_fecha(reg.fecha_adquisicion) if reg.fecha_adquisicion else "",
            fmt_importe(valor_me),
            fmt_tc(tc_adq),
            fmt_importe(valor_mn),
            fmt_tc(tc_31),
            fmt_importe(ajuste),
            fmt_importe(dep),
            fmt_importe(imp["dep_retiros"]),
            fmt_importe(imp["dep_otros"]),
            C.ESTADO_OPERACION,
        ])
    return lineas, obs
