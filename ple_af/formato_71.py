"""Construcción del Formato 7.1 (37 campos) a partir del Excel de Contasis."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from decimal import Decimal

from . import config as C
from .config import Config
from .diario import IndiceDiario
from .lector_contasis import ArchivoContasis, CAMPOS_IMPORTE, NOMBRES_CAMPOS
from .observaciones import Observaciones
from .utilidades import (
    CERO, a_decimal, a_fecha, a_texto, fmt_fecha, fmt_importe, limpiar_texto, normalizar,
    redondear,
)

TOLERANCIA = Decimal("0.01")

TITULOS_71 = [
    "1 Periodo", "2 CUO", "3 Correlativo asiento", "4 Catálogo", "5 Código activo",
    "6 Catálogo existencia", "7 Código existencia", "8 Tipo activo", "9 Cuenta contable",
    "10 Estado activo", "11 Descripción", "12 Marca", "13 Modelo", "14 Serie/placa",
    "15 Saldo inicial", "16 Adquisiciones", "17 Mejoras", "18 Retiros/bajas",
    "19 Otros ajustes", "20 Reval. voluntaria", "21 Reval. reorganización",
    "22 Otras revaluaciones", "23 Ajuste inflación", "24 Fecha adquisición",
    "25 Fecha inicio uso", "26 Método", "27 Doc. autorización", "28 % depreciación",
    "29 Dep. acum. anterior", "30 Dep. ejercicio", "31 Dep. retiros/bajas",
    "32 Dep. otros ajustes", "33 Dep. reval. voluntaria", "34 Dep. reval. reorganización",
    "35 Dep. otras reval.", "36 Ajuste inflación dep.", "37 Estado operación",
]


@dataclass
class Registro71:
    fila: int
    codigo: str
    campos: list[str]
    importes: dict = field(default_factory=dict)   # Decimales (para totales y 7.3)
    fecha_adquisicion: dt.date | None = None
    descripcion: str = ""
    descripcion_original: str = ""


def truncar_descripcion(texto: str, largo: int = C.LARGO_DESCRIPCION, conservar_sufijo: bool = False) -> str:
    """Recorta a `largo` caracteres sin partir palabras.

    Con `conservar_sufijo=True` mantiene el correlativo final de Contasis (ej. ' 0001');
    se usa solo cuando dos activos quedarían con la misma descripción recortada.
    """
    if len(texto) <= largo:
        return texto
    if conservar_sufijo:
        m = re.search(r"\s(\d{3,5})$", texto)
        if m:
            sufijo = " " + m.group(1)
            base = truncar_descripcion(texto[: m.start()].rstrip(), largo - len(sufijo))
            return base + sufijo
    corte = texto[:largo]
    if texto[largo] != " " and " " in corte:
        corte = corte[: corte.rfind(" ")]
    return corte.rstrip(" ,;-/")


def _importe(activo, clave, obs, codigo) -> Decimal:
    try:
        valor = a_decimal(activo.get(clave))
    except ValueError as exc:
        obs.error(f"{NOMBRES_CAMPOS[clave]}: {exc}", activo["fila"], codigo, NOMBRES_CAMPOS[clave])
        return CERO
    return valor if valor is not None else CERO


def construir_71(
    archivo: ArchivoContasis, cfg: Config, diario: IndiceDiario | None = None
) -> tuple[list[Registro71], Observaciones]:
    obs = Observaciones()
    registros: list[Registro71] = []
    fin_ejercicio = dt.date(cfg.ejercicio, 12, 31)
    col = archivo.columnas
    codigos_vistos: dict[str, int] = {}
    cuos_vistos: dict[str, str] = {}
    contador_serie = 0

    for fila in archivo.filas_sin_codigo:
        obs.error("Fila con datos pero sin código de activo.", fila, campo="Código", columna=col.get("codigo", ""))

    for n, activo in enumerate(archivo.activos, start=1):
        fila = activo["fila"]
        crudo_codigo = activo["codigo"]
        codigo = limpiar_texto(crudo_codigo)

        # --- Código (campo 5) ---
        if activo.get("codigo_numerico"):
            obs.advertencia("El código vino como número: podría haber perdido ceros a la izquierda.",
                            fila, codigo, "Código", col["codigo"])
        if len(codigo) > C.LARGO_CODIGO:
            obs.error(f"Código de más de {C.LARGO_CODIGO} caracteres.", fila, codigo, "Código", col["codigo"])
        if codigo in codigos_vistos:
            obs.error(f"Código duplicado (también en la fila {codigos_vistos[codigo]}).",
                      fila, codigo, "Código", col["codigo"])
        codigos_vistos.setdefault(codigo, fila)

        # --- Cuenta contable (campo 9) ---
        cuenta = re.sub(r"\D", "", a_texto(activo.get("cuenta")))
        if not cuenta:
            obs.error("Falta la cuenta contable.", fila, codigo, "Cuenta contable", col["cuenta"])
        elif len(cuenta) > C.LARGO_CUENTA:
            obs.error("Cuenta contable de más de 24 dígitos.", fila, codigo, "Cuenta contable", col["cuenta"])

        # --- Descripción (campo 11) ---
        descripcion = limpiar_texto(activo.get("descripcion"))
        descripcion_original = descripcion
        if not descripcion:
            obs.error("Falta la descripción.", fila, codigo, "Descripción", col["descripcion"])
        elif len(descripcion) > C.LARGO_DESCRIPCION:
            corta = truncar_descripcion(descripcion)
            obs.advertencia(f"Descripción de {len(descripcion)} caracteres recortada a 40: '{corta}' "
                            f"(original: '{descripcion}').",
                            fila, codigo, "Descripción", col["descripcion"])
            descripcion = corta

        # --- Marca, modelo, serie (campos 12-14) ---
        def texto_opcional(clave, largo, nombre):
            valor = limpiar_texto(activo.get(clave))
            if not valor:
                return None
            if len(valor) > largo:
                corto = valor[:largo].rstrip()
                obs.advertencia(f"{nombre} de {len(valor)} caracteres recortado a {largo}: '{corto}'.",
                                fila, codigo, nombre, col[clave])
                return corto
            return valor

        marca = texto_opcional("marca", C.LARGO_MARCA, "Marca")
        if marca is None:
            marca = cfg.texto_faltante
            obs.autocompletado(f"Marca vacía → '{marca}'.", fila, codigo, "Marca", col["marca"])
        modelo = texto_opcional("modelo", C.LARGO_MODELO, "Modelo")
        if modelo is None:
            modelo = cfg.texto_faltante
            obs.autocompletado(f"Modelo vacío → '{modelo}'.", fila, codigo, "Modelo", col["modelo"])
        serie = texto_opcional("serie", C.LARGO_SERIE, "Serie")
        if serie is None:
            if cfg.serie_faltante == "-":
                serie = "-"
            else:
                contador_serie += 1
                serie = f"GENERICO-{contador_serie:02d}"
            obs.autocompletado(f"Serie vacía → '{serie}'.", fila, codigo, "Serie/placa", col["serie"])

        # --- Importes ---
        imp = {clave: _importe(activo, clave, obs, codigo) for clave in CAMPOS_IMPORTE}

        # --- Fechas (campos 24 y 25) ---
        fecha_adq = fecha_uso = None
        try:
            fecha_adq = a_fecha(activo.get("fecha_adquisicion"))
        except ValueError as exc:
            obs.error(str(exc), fila, codigo, "Fecha de adquisición", col["fecha_adquisicion"])
        if fecha_adq is None:
            obs.error("Falta la fecha de adquisición.", fila, codigo, "Fecha de adquisición", col["fecha_adquisicion"])
        elif fecha_adq > fin_ejercicio:
            obs.error(f"Fecha de adquisición {fmt_fecha(fecha_adq)} posterior al 31/12/{cfg.ejercicio}.",
                      fila, codigo, "Fecha de adquisición", col["fecha_adquisicion"])
        try:
            fecha_uso = a_fecha(activo.get("fecha_uso"))
        except ValueError as exc:
            obs.error(str(exc), fila, codigo, "Fecha inicio de uso", col["fecha_uso"])
        if fecha_uso is None and fecha_adq is not None:
            fecha_uso = fecha_adq
            obs.autocompletado("Fecha de inicio de uso vacía → igual a la fecha de adquisición.",
                               fila, codigo, "Fecha inicio de uso", col["fecha_uso"])
        if fecha_uso and fecha_uso > fin_ejercicio:
            obs.error(f"Fecha de inicio de uso {fmt_fecha(fecha_uso)} posterior al 31/12/{cfg.ejercicio}.",
                      fila, codigo, "Fecha inicio de uso", col["fecha_uso"])
        if fecha_uso and fecha_adq and fecha_uso < fecha_adq:
            obs.advertencia("La fecha de inicio de uso es anterior a la fecha de adquisición.",
                            fila, codigo, "Fecha inicio de uso", col["fecha_uso"])
        if fecha_adq and fecha_adq < dt.date(2013, 1, 1):
            obs.advertencia(
                "Activo adquirido antes de 2013: revise si su costo incluye diferencias de cambio "
                "capitalizadas; en ese caso podría corresponder informar el 7.3.",
                fila, codigo, "Fecha de adquisición", col["fecha_adquisicion"])

        # --- Método (campo 26) ---
        metodo_txt = normalizar(activo.get("metodo"))
        if not metodo_txt:
            metodo = "1"
            obs.autocompletado("Método vacío → 1 (línea recta).", fila, codigo, "Método", col["metodo"])
        else:
            metodo = C.METODOS.get(metodo_txt, C.METODO_OTROS)
            if metodo == C.METODO_OTROS:
                obs.advertencia(f"Método '{activo.get('metodo')}' no reconocido → 9 (otros).",
                                fila, codigo, "Método", col["metodo"])

        # --- Documento de autorización (campo 27) ---
        doc = limpiar_texto(activo.get("doc_autorizacion"))[: C.LARGO_DOC_AUTORIZACION]
        if not doc:
            doc = f"{n:08d}" if cfg.campo27 == "correlativo" else "-"

        # --- Porcentaje (campo 28) ---
        try:
            pct = a_decimal(activo.get("porcentaje"))
        except ValueError as exc:
            obs.error(str(exc), fila, codigo, "% depreciación", col["porcentaje"])
            pct = CERO
        if pct is None:
            pct = CERO
            obs.advertencia("% de depreciación vacío → 0.00 (correcto solo para terrenos u otros activos no depreciables).",
                            fila, codigo, "% depreciación", col["porcentaje"])
        elif abs(pct) <= 1:
            pct = pct * 100
        if pct < 0 or pct > 100:
            obs.error(f"% de depreciación fuera de rango: {pct}.", fila, codigo, "% depreciación", col["porcentaje"])

        # --- Controles aritméticos ---
        costo_a = imp["saldo_inicial"] + imp["adquisiciones"] + imp["mejoras"] + imp["retiros"] + imp["otros_ajustes"]
        costo_b = imp["saldo_inicial"] + imp["adquisiciones"] + imp["mejoras"] - imp["retiros"] + imp["otros_ajustes"]
        historico = imp["valor_historico"]
        if min(abs(costo_a - historico), abs(costo_b - historico)) > TOLERANCIA:
            obs.advertencia(f"Saldo inicial + movimientos ({fmt_importe(costo_b)}) ≠ valor histórico ({fmt_importe(historico)}).",
                            fila, codigo, "Valor histórico", col["valor_historico"])
        dep_calc = imp["dep_acum_anterior"] + imp["dep_ejercicio"] + imp["dep_retiros"] + imp["dep_otros"]
        dep_calc_b = imp["dep_acum_anterior"] + imp["dep_ejercicio"] - imp["dep_retiros"] + imp["dep_otros"]
        dep_hist = imp["dep_acum_historica"]
        if min(abs(dep_calc - dep_hist), abs(dep_calc_b - dep_hist)) > TOLERANCIA:
            obs.advertencia(f"Dep. anterior + del ejercicio + bajas + ajustes ({fmt_importe(dep_calc)}) ≠ "
                            f"dep. acumulada histórica ({fmt_importe(dep_hist)}).",
                            fila, codigo, "Dep. acumulada histórica", col["dep_acum_historica"])
        if dep_hist > historico + TOLERANCIA:
            obs.advertencia(f"Depreciación acumulada ({fmt_importe(dep_hist)}) mayor que el costo ({fmt_importe(historico)}).",
                            fila, codigo, "Dep. acumulada histórica", col["dep_acum_historica"])
        if imp["dep_retiros"] != 0 and imp["retiros"] == 0:
            obs.advertencia(
                f"Baja incompleta: hay depreciación de bajas ({fmt_importe(imp['dep_retiros'])}) pero no retiro del costo.",
                fila, codigo, "Retiros/bajas", col["retiros"])
        if imp["retiros"] != 0 and imp["dep_retiros"] == 0 and imp["dep_acum_anterior"] + imp["dep_ejercicio"] != 0:
            obs.advertencia(
                f"Baja incompleta: hay retiro del costo ({fmt_importe(imp['retiros'])}) pero no depreciación de bajas.",
                fila, codigo, "Dep. retiros/bajas", col["dep_retiros"])

        # --- CUO y correlativo (campos 2 y 3): asiento del libro diario ---
        cuo, correlativo = f"{cfg.prefijo_cuo}{n}", f"{cfg.prefijo_asiento}{n}"
        if diario is not None:
            asiento = diario.ultimo(codigo)
            if asiento is None:
                obs.advertencia(
                    f"No se encontró asiento del sub diario {cfg.subdiario} para este activo en el diario "
                    f"subido; se usó el CUO propio '{cuo}'.", fila, codigo, "CUO")
            else:
                cuo = asiento.cuo
                preferido = asiento.correlativo_68 if cfg.linea_correlativo == "68" else asiento.correlativo_39
                correlativo = preferido or asiento.correlativo_68 or asiento.correlativo_39 or "M1"
                if diario.anio_completo:
                    dep_diario = redondear(diario.depreciacion_anual(codigo))
                    if abs(dep_diario - redondear(imp["dep_ejercicio"])) > TOLERANCIA:
                        obs.advertencia(
                            f"Depreciación del ejercicio ({fmt_importe(imp['dep_ejercicio'])}) ≠ suma de la "
                            f"depreciación en el diario ({fmt_importe(dep_diario)}).",
                            fila, codigo, "Dep. del ejercicio", col["dep_ejercicio"])
        clave_cuo = f"{cuo}|{correlativo}"
        if clave_cuo in cuos_vistos:
            obs.error(f"CUO {cuo} / {correlativo} repetido (también en el activo {cuos_vistos[clave_cuo]}).",
                      fila, codigo, "CUO")
        cuos_vistos.setdefault(clave_cuo, codigo)

        campos = [
            cfg.periodo,                                   # 1
            cuo,                                           # 2
            correlativo,                                   # 3
            C.CATALOGO_OTROS,                              # 4
            codigo,                                        # 5
            "",                                            # 6
            "",                                            # 7
            C.TIPO_ACTIVO,                                 # 8
            cuenta,                                        # 9
            C.ESTADO_ACTIVO,                               # 10
            descripcion,                                   # 11
            marca,                                         # 12
            modelo,                                        # 13
            serie,                                         # 14
            fmt_importe(imp["saldo_inicial"]),             # 15
            fmt_importe(imp["adquisiciones"]),             # 16
            fmt_importe(imp["mejoras"]),                   # 17
            fmt_importe(imp["retiros"]),                   # 18
            fmt_importe(imp["otros_ajustes"]),             # 19
            "0.00", "0.00", "0.00",                        # 20-22
            fmt_importe(imp["ajuste_inflacion"]),          # 23
            fmt_fecha(fecha_adq) if fecha_adq else "",     # 24
            fmt_fecha(fecha_uso) if fecha_uso else "",     # 25
            metodo,                                        # 26
            doc,                                           # 27
            fmt_importe(pct),                              # 28
            fmt_importe(imp["dep_acum_anterior"]),         # 29
            fmt_importe(imp["dep_ejercicio"]),             # 30
            fmt_importe(imp["dep_retiros"]),               # 31
            fmt_importe(imp["dep_otros"]),                 # 32
            "0.00", "0.00", "0.00",                        # 33-35
            fmt_importe(imp["ajuste_inflacion_dep"]),      # 36
            C.ESTADO_OPERACION,                            # 37
        ]
        assert len(campos) == 37
        registros.append(Registro71(fila, codigo, campos, {k: redondear(v) for k, v in imp.items()},
                                    fecha_adq, descripcion, descripcion_original))

    if len(f"{cfg.prefijo_cuo}{len(archivo.activos)}") > C.LARGO_CUO:
        obs.error("El prefijo del CUO es demasiado largo.")
    if not re.fullmatch(r"[AMC][A-Za-z0-9]*", cfg.prefijo_asiento or ""):
        obs.error("El prefijo del correlativo de asiento debe empezar con A, M o C.")
    # Si dos activos quedan con la misma descripción recortada, se conserva su correlativo final
    por_descripcion: dict[str, list[Registro71]] = {}
    for r in registros:
        por_descripcion.setdefault(r.campos[10], []).append(r)
    for grupo in por_descripcion.values():
        if len(grupo) < 2:
            continue
        for r in grupo:
            if r.descripcion_original != r.campos[10]:
                nueva = truncar_descripcion(r.descripcion_original, conservar_sufijo=True)
                r.campos[10] = r.descripcion = nueva
                obs.advertencia(f"Descripción recortada repetida con otro activo; se conservó el correlativo: '{nueva}'.",
                                r.fila, r.codigo, "Descripción", col["descripcion"])
    return registros, obs

