"""Lectura del Libro Diario PLE (TXT 5.1 o 5.2) para obtener el CUO de cada activo.

Contasis registra la depreciación mensual en el sub diario 09 con un asiento por activo:

    20241200|12.09.1|M1|6841394|...|040100020001 - Depreciacion 12||1079.13|0.00||1|
    20241200|12.09.1|M4|39525  |...|040100020001 - Depreciacion 12||0.00|1079.13||1|

El CUO del activo en el 7.1 es el del asiento del diario (12.09.1) y el correlativo es el
de la línea del asiento (M1 = cuenta 68, M4 = cuenta 39). Para cada activo se toma el
asiento del último mes disponible.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal

from .utilidades import CERO, a_decimal

# Posiciones (0-based) en las estructuras 5.1 y 5.2
I_PERIODO, I_CUO, I_CORRELATIVO, I_CUENTA, I_GLOSA, I_DEBE, I_HABER = 0, 1, 2, 3, 15, 17, 18

RE_NOMBRE = re.compile(r"LE(\d{11})(\d{4})(\d{2})\d{2}(05\d{4})", re.I)


@dataclass
class AsientoActivo:
    mes: int
    cuo: str
    correlativo_68: str = ""     # línea del gasto (cuenta 68)
    correlativo_39: str = ""     # línea de la depreciación acumulada (cuenta 39)
    depreciacion: Decimal = CERO  # haber de la cuenta 39 en el asiento


@dataclass
class IndiceDiario:
    por_codigo: dict[str, list[AsientoActivo]] = field(default_factory=dict)
    meses: set[int] = field(default_factory=set)
    ruc: str = ""
    ejercicio: int | None = None
    archivos: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)

    def ultimo(self, codigo: str) -> AsientoActivo | None:
        asientos = self.por_codigo.get(codigo)
        return max(asientos, key=lambda a: a.mes) if asientos else None

    def depreciacion_anual(self, codigo: str) -> Decimal:
        return sum((a.depreciacion for a in self.por_codigo.get(codigo, [])), CERO)

    @property
    def anio_completo(self) -> bool:
        return self.meses >= set(range(1, 13))


def leer_diarios(archivos: list[tuple[str, bytes]], codigos: set[str], subdiario: str = "09") -> IndiceDiario:
    """Lee uno o varios TXT del diario y arma el índice código de activo -> asientos."""
    ind = IndiceDiario()
    patron_cuo = re.compile(rf"^(\d{{1,2}})\.{re.escape(subdiario)}\.(\S+)$")
    for nombre, contenido in archivos:
        m = RE_NOMBRE.search(nombre or "")
        if m:
            if not m.group(4).startswith(("050100", "050200")):
                ind.avisos.append(f"{nombre}: no es un Libro Diario (5.1 / 5.2); se ignoró.")
                continue
            if ind.ruc and m.group(1) != ind.ruc:
                ind.avisos.append(f"{nombre}: es de otro RUC ({m.group(1)}).")
            ind.ruc = ind.ruc or m.group(1)
            ind.ejercicio = ind.ejercicio or int(m.group(2))
        ind.archivos.append(nombre)
        texto = contenido.decode("cp1252", errors="replace")
        encontrados = 0
        asientos: dict[str, AsientoActivo] = {}
        for linea in texto.splitlines():
            campos = linea.split("|")
            if len(campos) <= I_HABER:
                continue
            cuo = campos[I_CUO].strip()
            mc = patron_cuo.match(cuo)
            if not mc:
                continue
            glosa = campos[I_GLOSA].strip()
            codigo = glosa.split(" ", 1)[0] if glosa else ""
            if codigo not in codigos:
                continue
            mes = int(campos[I_PERIODO][4:6]) if campos[I_PERIODO][4:6].isdigit() else int(mc.group(1))
            clave = f"{codigo}|{cuo}"
            asiento = asientos.get(clave)
            if asiento is None:
                asiento = asientos[clave] = AsientoActivo(mes=mes, cuo=cuo)
                ind.por_codigo.setdefault(codigo, []).append(asiento)
                encontrados += 1
            cuenta = campos[I_CUENTA].strip()
            correlativo = campos[I_CORRELATIVO].strip()
            if cuenta.startswith("68") and not asiento.correlativo_68:
                asiento.correlativo_68 = correlativo
            elif cuenta.startswith("39"):
                asiento.correlativo_39 = asiento.correlativo_39 or correlativo
                try:
                    asiento.depreciacion += (a_decimal(campos[I_HABER]) or CERO) - (a_decimal(campos[I_DEBE]) or CERO)
                except ValueError:
                    pass
            ind.meses.add(mes)
        if not encontrados:
            ind.avisos.append(f"{nombre}: no tiene asientos del sub diario {subdiario} con códigos de activo.")
    return ind
