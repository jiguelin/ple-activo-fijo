"""Controles de totales y de continuidad con el ejercicio anterior."""
from __future__ import annotations

from decimal import Decimal

from .formato_71 import Registro71
from .lector_contasis import ArchivoContasis, NOMBRES_CAMPOS
from .observaciones import Observaciones
from .utilidades import CERO, a_decimal, fmt_importe, redondear

TOLERANCIA = Decimal("0.01")

CAMPOS_TOTALES = [
    "saldo_inicial", "adquisiciones", "mejoras", "retiros", "otros_ajustes",
    "valor_historico", "ajuste_inflacion", "dep_acum_anterior", "dep_ejercicio",
    "dep_retiros", "dep_otros", "dep_acum_historica", "ajuste_inflacion_dep",
]


def totales(registros: list[Registro71]) -> dict[str, Decimal]:
    return {c: sum((r.importes[c] for r in registros), CERO) for c in CAMPOS_TOTALES}


def comparar_totales(archivo: ArchivoContasis, registros: list[Registro71]) -> tuple[list[dict], Observaciones]:
    """Tabla de totales calculados vs. fila 'Totales:' del Excel."""
    obs = Observaciones()
    calc = totales(registros)
    tabla = []
    for campo in CAMPOS_TOTALES:
        excel = None
        if archivo.totales:
            try:
                excel = a_decimal(archivo.totales.get(campo))
            except ValueError:
                excel = None
        excel_r = redondear(excel) if excel is not None else None
        diferencia = (calc[campo] - excel_r) if excel_r is not None else None
        tabla.append({
            "Concepto": NOMBRES_CAMPOS[campo],
            "Calculado": float(calc[campo]),
            "Excel (fila Totales)": float(excel_r) if excel_r is not None else None,
            "Diferencia": float(diferencia) if diferencia is not None else None,
        })
        if diferencia is not None and abs(diferencia) > TOLERANCIA:
            obs.advertencia(
                f"Total de '{NOMBRES_CAMPOS[campo]}' calculado ({fmt_importe(calc[campo])}) ≠ fila Totales "
                f"del Excel ({fmt_importe(excel_r)}).", archivo.fila_totales, campo=NOMBRES_CAMPOS[campo])
    if not archivo.totales:
        obs.advertencia("El Excel no tiene fila 'Totales:'; no se pudo conciliar.")
    tabla.append({"Concepto": "Cantidad de activos", "Calculado": float(len(registros)),
                  "Excel (fila Totales)": None, "Diferencia": None})
    return tabla, obs


def comparar_anio_anterior(actual: list[Registro71], anterior: list[Registro71], ejercicio_anterior) -> Observaciones:
    """Compara saldo inicial y depreciación acumulada con el cierre del año anterior."""
    obs = Observaciones()
    prev = {r.codigo: r for r in anterior}
    act = {r.codigo: r for r in actual}
    for r in actual:
        p = prev.get(r.codigo)
        if p is None:
            if r.importes["saldo_inicial"] != 0:
                obs.advertencia(
                    f"Activo con saldo inicial {fmt_importe(r.importes['saldo_inicial'])} que no estaba en {ejercicio_anterior}.",
                    r.fila, r.codigo, "Saldo inicial")
            continue
        if abs(r.importes["saldo_inicial"] - p.importes["valor_historico"]) > TOLERANCIA:
            obs.advertencia(
                f"Saldo inicial ({fmt_importe(r.importes['saldo_inicial'])}) ≠ valor histórico al cierre de "
                f"{ejercicio_anterior} ({fmt_importe(p.importes['valor_historico'])}).",
                r.fila, r.codigo, "Saldo inicial")
        if abs(r.importes["dep_acum_anterior"] - p.importes["dep_acum_historica"]) > TOLERANCIA:
            obs.advertencia(
                f"Dep. acumulada anterior ({fmt_importe(r.importes['dep_acum_anterior'])}) ≠ dep. acumulada al cierre "
                f"de {ejercicio_anterior} ({fmt_importe(p.importes['dep_acum_historica'])}).",
                r.fila, r.codigo, "Dep. acumulada anterior")
    for p in anterior:
        if p.codigo not in act and p.importes["valor_historico"] != 0:
            obs.advertencia(
                f"Estaba en {ejercicio_anterior} (costo {fmt_importe(p.importes['valor_historico'])}) y ya no aparece, "
                f"sin retiro registrado en este ejercicio.", None, p.codigo, "Retiros/bajas")
    return obs
