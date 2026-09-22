"""Pruebas con exportaciones reales de Contasis (tests/fixtures)."""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest

from ple_af import Config, DatoME, generar, leer_diarios, leer_excel_contasis
from ple_af.formato_71 import truncar_descripcion
from ple_af.utilidades import fmt_importe, ruc_valido

FIX = Path(__file__).parent / "fixtures"

ESPERADO = {
    "essentta_2022.xlsx": ("20565449843", 2022, 33),
    "essentta_2024.xlsx": ("20565449843", 2024, 35),
    "contadeus_2023.xlsx": ("20609565625", 2023, 7),
    "dental_2022.xlsx": ("20548446971", 2022, 28),
    "dental_2023.xlsx": ("20548446971", 2023, 28),
    "dental_2024.xlsx": ("20548446971", 2024, 31),
    "dental_2025.xlsx": ("20548446971", 2025, 31),
    "econobrillo_2025.xlsx": ("20601601240", 2025, 5),
}


def leer(nombre):
    return leer_excel_contasis((FIX / nombre).read_bytes(), nombre)


def generar_de(nombre, **kw):
    a = leer(nombre)
    return a, generar(a, Config(ruc=a.ruc, ejercicio=a.ejercicio, **kw))


def lineas_71(res):
    nombre = next(n for n in res.archivos if "070100" in n)
    return res.archivos[nombre].decode("cp1252").split("\r\n")[:-1]


@pytest.mark.parametrize("nombre", ESPERADO)
def test_lectura_y_generacion(nombre):
    ruc, ejercicio, cantidad = ESPERADO[nombre]
    a, res = generar_de(nombre)
    assert (a.ruc, a.ejercicio, len(a.activos)) == (ruc, ejercicio, cantidad)
    assert not res.observaciones.hay_errores, res.observaciones.de_tipo("Error")
    # 3 archivos con los nombres correctos
    assert sorted(res.archivos) == sorted([
        f"LE{ruc}{ejercicio}0000070100001111.txt",
        f"LE{ruc}{ejercicio}0000070300001011.txt",
        f"LE{ruc}{ejercicio}0000070400001011.txt",
    ])
    # 7.3 y 7.4 vacíos (0 bytes)
    for n, datos in res.archivos.items():
        if "070100" not in n:
            assert datos == b""
    # cada línea: 37 campos + '|' final, CRLF
    lineas = lineas_71(res)
    assert len(lineas) == cantidad
    for linea in lineas:
        assert linea.endswith("|")
        assert len(linea.split("|")) == 38
    # los totales calculados coinciden con la fila Totales del Excel
    for fila in res.tabla_totales:
        if fila["Diferencia"] is not None:
            assert abs(fila["Diferencia"]) < 0.01, fila


def test_totales_essentta_2024():
    _, res = generar_de("essentta_2024.xlsx")
    t = {f["Concepto"]: f["Calculado"] for f in res.tabla_totales}
    assert t["Saldo inicial"] == pytest.approx(172507.54)
    assert t["Adquisiciones"] == pytest.approx(4152.54)
    assert t["Dep. del ejercicio"] == pytest.approx(22836.81)


def test_primera_linea_essentta_2024():
    _, res = generar_de("essentta_2024.xlsx")
    campos = lineas_71(res)[0].split("|")
    assert campos[0] == "20240000"
    assert campos[4] == "040100020001"          # conserva ceros
    assert campos[8] == "33411"
    assert campos[10] == "AUTO HONDA 4X4 0001"
    assert campos[13] == "1HGRW2870LL500956/BTD-410"
    assert campos[23:28] == ["21/11/2020", "21/11/2020", "1", "-", "20.00"]
    assert campos[28:30] == ["28057.38", "12949.56"]
    assert campos[36] == "1"


def test_valores_por_defecto_y_numeros_como_texto():
    _, res = generar_de("essentta_2024.xlsx")
    lineas = [l.split("|") for l in lineas_71(res)]
    escritorio = next(l for l in lineas if l[4] == "050100030001")
    assert escritorio[11:14] == ["-", "-", "-"]              # por defecto: '-' como indica SUNAT
    _, res = generar_de("essentta_2024.xlsx", texto_faltante="GENERICO", serie_faltante="GENERICO-nn")
    lineas = [l.split("|") for l in lineas_71(res)]
    escritorio = next(l for l in lineas if l[4] == "050100030001")
    assert escritorio[11:14] == ["GENERICO", "GENERICO", "GENERICO-01"]
    vitrina = next(l for l in lineas if l[4] == "050100080001")
    assert vitrina[12] == "2017"                # modelo numérico sin ".0"
    _, res = generar_de("dental_2022.xlsx")
    assert any("|202109252|" in l for l in lineas_71(res))   # serie numérica


def test_opciones_guion():
    _, res = generar_de("essentta_2024.xlsx", campo27="correlativo")
    escritorio = next(l.split("|") for l in lineas_71(res) if "|050100030001|" in l)
    assert escritorio[26] == "00000002"


def test_descripciones_y_modelo_largos():
    _, res = generar_de("dental_2022.xlsx")
    lineas = [l.split("|") for l in lineas_71(res)]
    assert all(len(l[10]) <= 40 and len(l[11]) <= 20 and len(l[12]) <= 20 and len(l[13]) <= 30 for l in lineas)
    assert "PIEZA DE MANO DENTAL MAS ACCESORIOS 0001" in [l[10] for l in lineas]
    assert truncar_descripcion("ZHIYUN CRANE 4 COMBO, SMALLRING TRIPODE AD-80 0001") == \
        "ZHIYUN CRANE 4 COMBO, SMALLRING TRI 0001"


def test_advertencias_dental_2022():
    _, res = generar_de("dental_2022.xlsx")
    adv = res.observaciones.de_tipo("Advertencia")
    assert any(o.codigo == "060900050001" and "Baja incompleta" in o.mensaje for o in adv)
    assert any("antes de 2013" in o.mensaje for o in adv)


def test_control_anio_anterior():
    anterior = leer("dental_2024.xlsx")
    actual = leer("dental_2025.xlsx")
    res = generar(actual, Config(ruc=actual.ruc, ejercicio=2025), archivo_anterior=anterior)
    continuidad = [o for o in res.observaciones.items if "Dep. acumulada anterior" in o.mensaje]
    assert len(continuidad) == 31


def test_ejercicio_minimo_y_ruc():
    a = leer("essentta_2024.xlsx")
    assert generar(a, Config(ruc=a.ruc, ejercicio=2021)).observaciones.hay_errores
    assert generar(a, Config(ruc="20565449840", ejercicio=2024)).observaciones.hay_errores
    assert ruc_valido("20565449843") and ruc_valido("20548446971")


def test_73_manual():
    a = leer("essentta_2024.xlsx")
    cfg = Config(ruc=a.ruc, ejercicio=2024)
    dron = DatoME(codigo="060900040001", valor_me="1439.83", tc_adquisicion="3.932")
    res = generar(a, cfg, [dron], tc_cierre="3.758")
    nombre = next(n for n in res.archivos if "070300" in n)
    assert nombre.endswith("070300001111.txt")
    linea = res.archivos[nombre].decode("cp1252")
    assert linea == "20240000|AF34|M34|9|060900040001|30/06/2021|1439.83|3.932|5661.41|3.758|0.00|516.12|0.00|0.00|1|\r\n"


def test_formato_igual_al_txt_2021():
    """El TXT de referencia (2021) tiene la misma forma: 37 campos, '|' final, CRLF, ASCII."""
    ref = (FIX / "LE2056544984320210000070100001111.txt").read_bytes()
    assert ref.endswith(b"|\r\n")
    for linea in ref.decode("cp1252").split("\r\n")[:-1]:
        assert len(linea.split("|")) == 38


def test_importes():
    assert fmt_importe(Decimal("-0.001")) == "0.00"
    assert fmt_importe(Decimal("2.005")) == "2.01"
    assert fmt_importe(None) == "0.00"


DIARIO_DIC = "LE2056544984320241200050200002111.txt"
DIARIO_NOV = "LE2056544984320241100050200001111.txt"


def diario(*nombres, codigos):
    return leer_diarios([(n, (FIX / n).read_bytes()) for n in nombres], codigos)


def test_cuo_desde_libro_diario():
    a = leer("essentta_2024.xlsx")
    codigos = {x["codigo"] for x in a.activos}
    ind = diario(DIARIO_NOV, DIARIO_DIC, codigos=codigos)
    assert ind.ruc == "20565449843" and ind.ejercicio == 2024 and ind.meses == {11, 12}
    res = generar(a, Config(ruc=a.ruc, ejercicio=2024), diario=ind)
    assert not res.observaciones.hay_errores
    lineas = [l.split("|") for l in lineas_71(res)]
    por_codigo = {l[4]: l for l in lineas}
    assert por_codigo["040100020001"][1:3] == ["12.09.1", "M1"]     # auto: asiento 1 de diciembre
    assert por_codigo["060100040004"][1:3] == ["12.09.28", "M1"]    # laptop comprada en diciembre
    assert por_codigo["060900050001"][1:3] == ["12.09.35", "M1"]
    assert not any("No se encontró asiento" in o.mensaje for o in res.observaciones.items)
    # opción: correlativo de la línea de la cuenta 39
    res39 = generar(a, Config(ruc=a.ruc, ejercicio=2024, linea_correlativo="39"), diario=ind)
    assert lineas_71(res39)[0].split("|")[1:3] == ["12.09.1", "M4"]


def test_diario_solo_noviembre_usa_ultimo_mes():
    a = leer("essentta_2024.xlsx")
    ind = diario(DIARIO_NOV, codigos={x["codigo"] for x in a.activos})
    res = generar(a, Config(ruc=a.ruc, ejercicio=2024), diario=ind)
    assert lineas_71(res)[0].split("|")[1] == "11.09.1"


def test_diario_de_otra_empresa_bloquea():
    a = leer("dental_2024.xlsx")
    ind = diario(DIARIO_DIC, codigos={x["codigo"] for x in a.activos})
    assert generar(a, Config(ruc=a.ruc, ejercicio=2024), diario=ind).observaciones.hay_errores


def test_sin_diario_advierte():
    _, res = generar_de("essentta_2024.xlsx")
    assert any("No se subió el Libro Diario" in o.mensaje for o in res.observaciones.items)
