"""Generador PLE - Registro de Activos Fijos (Formatos 7.1, 7.3 y 7.4).

Sube el Excel "FORMATO 7.1" exportado desde Contasis y descarga los TXT para el PLE.
"""
from __future__ import annotations

import hmac

import pandas as pd
import streamlit as st

from ple_af import Config, DatoME, generar, leer_diarios, leer_excel_contasis, reporte_excel
from ple_af.config import EJERCICIO_MINIMO
from ple_af.formato_71 import TITULOS_71
from ple_af.formato_73 import TITULOS_73
from ple_af.observaciones import ADVERTENCIA, AUTOCOMPLETADO, ERROR

st.set_page_config(page_title="PLE Activos Fijos", page_icon="📒", layout="wide")


# ---------------------------------------------------------------- contraseña opcional
def _clave_ok() -> bool:
    try:
        clave = st.secrets.get("APP_PASSWORD")
    except Exception:  # sin archivo secrets.toml
        clave = None
    if not clave:
        return True
    if st.session_state.get("autenticado"):
        return True
    st.title("📒 Generador PLE – Activos Fijos")
    ingresada = st.text_input("Contraseña", type="password")
    if ingresada:
        if hmac.compare_digest(ingresada, str(clave)):
            st.session_state["autenticado"] = True
            st.rerun()
        st.error("Contraseña incorrecta.")
    return False


if not _clave_ok():
    st.stop()


@st.cache_data(show_spinner=False)
def _leer(contenido: bytes, nombre: str):
    return leer_excel_contasis(contenido, nombre)


@st.cache_data(show_spinner=False)
def _leer_diario(archivos: tuple, codigos: frozenset, subdiario: str):
    return leer_diarios(list(archivos), set(codigos), subdiario)


# ---------------------------------------------------------------- encabezado
st.title("📒 Generador PLE – Registro de Activos Fijos")
st.caption(
    "Sube el Excel **FORMATO 7.1** exportado desde Contasis y descarga los TXT del PLE: "
    "7.1 con datos, 7.3 y 7.4 vacíos (salvo que actives el 7.3 manual). Ejercicios "
    f"{EJERCICIO_MINIMO} en adelante. Los archivos no se guardan en el servidor."
)

col_a, col_b = st.columns([3, 2])
with col_a:
    subido = st.file_uploader("1. Excel de Contasis (FORMATO 7.1)", type=["xlsx", "xlsm", "xls"])
with col_b:
    anterior_subido = st.file_uploader(
        "Opcional: Excel del ejercicio anterior (para controlar saldos iniciales)",
        type=["xlsx", "xlsm", "xls"],
    )

diarios_subidos = st.file_uploader(
    "2. Libro Diario PLE del mismo ejercicio (TXT 5.1 o 5.2) – puedes subir varios meses; basta diciembre",
    type=["txt"], accept_multiple_files=True,
    help="De aquí se toma el CUO (ej. 12.09.1) y el correlativo (ej. M1) del asiento de depreciación "
         "de cada activo en el sub diario 09. Sin el diario, el CUO se genera como AF1, AF2…",
)

if not subido:
    st.info("Sube el archivo para empezar.")
    st.stop()

try:
    archivo = _leer(subido.getvalue(), subido.name)
except Exception as exc:  # archivo con otra estructura
    st.error(f"No se pudo leer el archivo: {exc}")
    st.stop()

archivo_anterior = None
if anterior_subido:
    try:
        archivo_anterior = _leer(anterior_subido.getvalue(), anterior_subido.name)
    except Exception as exc:
        st.warning(f"No se pudo leer el archivo del año anterior: {exc}")

# ---------------------------------------------------------------- parámetros
with st.sidebar:
    st.header("Datos de la empresa")
    ruc = st.text_input("RUC", value=archivo.ruc, max_chars=11)
    razon = st.text_input("Razón social", value=archivo.razon_social)
    ejercicio = st.number_input("Ejercicio", min_value=2000, max_value=2100,
                                value=int(archivo.ejercicio or 2024), step=1)

    st.header("Opciones del TXT")
    texto_faltante = st.radio("Marca / modelo vacíos", ["GENERICO", "-"], horizontal=True)
    serie_faltante = st.radio("Serie vacía", ["GENERICO-nn", "-"], horizontal=True,
                              help="GENERICO-nn numera GENERICO-01, GENERICO-02…")
    campo27 = st.radio("Campo 27 (doc. autorización) vacío", ["-", "correlativo"], horizontal=True,
                       help="'correlativo' pone 00000001, 00000002…")
    with st.expander("Avanzado"):
        formato_periodo = st.selectbox("Campo 1 (periodo)", ["AAAA0000", "AAAA1200"])
        prefijo_cuo = st.text_input("Prefijo del CUO", value="AF", max_chars=10)
        prefijo_asiento = st.text_input("Prefijo del correlativo de asiento", value="M", max_chars=3,
                                        help="Debe empezar con A, M o C.")
        indicador_op = st.selectbox(
            "Indicador de operaciones", ["1", "0", "2"],
            format_func=lambda x: {"1": "1 - Empresa operativa", "0": "0 - Cierre / baja de RUC",
                                   "2": "2 - Cierre del libro"}[x])
        oportunidad = st.text_input("Código de oportunidad", value="00", max_chars=2)
        subdiario = st.text_input("Sub diario de depreciación (Contasis)", value="09", max_chars=4)
        linea_corr = st.radio(
            "Correlativo (campo 3) = línea del asiento con la cuenta", ["68", "39"], horizontal=True,
            help="68 = gasto por depreciación (M1 en Contasis); 39 = depreciación acumulada (M4).")

cfg = Config(
    ruc=ruc.strip(), ejercicio=int(ejercicio), razon_social=razon,
    formato_periodo=formato_periodo, codigo_oportunidad=oportunidad or "00",
    indicador_operaciones=indicador_op, prefijo_cuo=prefijo_cuo.strip() or "AF",
    prefijo_asiento=prefijo_asiento.strip() or "M", texto_faltante=texto_faltante,
    serie_faltante=serie_faltante, campo27=campo27,
    subdiario=subdiario.strip() or "09", linea_correlativo=linea_corr,
)

diario = None
if diarios_subidos:
    diario = _leer_diario(
        tuple((f.name, f.getvalue()) for f in diarios_subidos),
        frozenset(a["codigo"] for a in archivo.activos), cfg.subdiario,
    )

st.markdown(
    f"**{razon or 'Empresa'}** · RUC {cfg.ruc} · Ejercicio **{cfg.ejercicio}** · "
    f"**{len(archivo.activos)}** activos leídos"
    + (f" · comparando con {archivo_anterior.ejercicio}" if archivo_anterior else "")
)
if diario is not None:
    con_asiento = sum(1 for a in archivo.activos if diario.ultimo(a["codigo"]))
    meses = ", ".join(f"{m:02d}" for m in sorted(diario.meses)) or "ninguno"
    st.markdown(f"Libro Diario: meses {meses} · **{con_asiento} de {len(archivo.activos)}** activos "
                f"con asiento en el sub diario {cfg.subdiario}")

# ---------------------------------------------------------------- 7.3 opcional
st.subheader("Formato 7.3 – Diferencia de cambio")
st.caption(
    "Desde 2013 la diferencia de cambio ya no se suma al costo del activo (se derogó el art. 61 inc. f "
    "de la LIR), así que el 7.3 va **vacío** aunque las compras hayan sido en dólares. "
    "Actívalo solo si en tu caso corresponde (por ejemplo, activos antiguos con diferencia de cambio en su costo)."
)
datos_me: list[DatoME] = []
tc_cierre = None
if st.toggle("Llenar el 7.3 manualmente", value=False):
    base = pd.DataFrame([
        {"Incluir": False, "Código": str(a["codigo"]), "Descripción": str(a.get("descripcion") or ""),
         "Valor ME": None, "TC adquisición": None, "Ajuste dif. cambio": 0.0}
        for a in archivo.activos
    ])
    editado = st.data_editor(
        base, hide_index=True, width="stretch", key="editor73",
        disabled=["Código", "Descripción"],
        column_config={
            "Valor ME": st.column_config.NumberColumn(format="%.2f"),
            "TC adquisición": st.column_config.NumberColumn(format="%.3f"),
            "Ajuste dif. cambio": st.column_config.NumberColumn(format="%.2f"),
        },
    )
    tc_cierre = st.number_input("Tipo de cambio al 31.12", min_value=0.0, step=0.001, format="%.3f")
    for _, f in editado[editado["Incluir"]].iterrows():
        datos_me.append(DatoME(
            codigo=f["Código"],
            valor_me=None if pd.isna(f["Valor ME"]) else f["Valor ME"],
            tc_adquisicion=None if pd.isna(f["TC adquisición"]) else f["TC adquisición"],
            ajuste=None if pd.isna(f["Ajuste dif. cambio"]) else f["Ajuste dif. cambio"],
        ))

# ---------------------------------------------------------------- generación
resultado = generar(archivo, cfg, datos_me, tc_cierre or None, archivo_anterior, diario)
obs = resultado.observaciones

st.subheader("Revisión")
n_err = len(obs.de_tipo(ERROR))
n_adv = len(obs.de_tipo(ADVERTENCIA))
n_auto = len(obs.de_tipo(AUTOCOMPLETADO))
c1, c2, c3 = st.columns(3)
c1.metric("Errores (bloquean)", n_err)
c2.metric("Advertencias", n_adv)
c3.metric("Valores completados", n_auto)

tab_obs, tab_71, tab_73, tab_tot = st.tabs(["Observaciones", "Vista 7.1", "Vista 7.3", "Totales"])
with tab_obs:
    if obs.items:
        tipos = st.multiselect("Mostrar", [ERROR, ADVERTENCIA, AUTOCOMPLETADO],
                               default=[t for t in (ERROR, ADVERTENCIA) if obs.de_tipo(t)] or [AUTOCOMPLETADO])
        df_obs = pd.DataFrame(obs.como_filas())
        st.dataframe(df_obs[df_obs["Tipo"].isin(tipos)], hide_index=True, width="stretch")
    else:
        st.success("Sin observaciones.")
with tab_71:
    if resultado.registros_71:
        st.dataframe(pd.DataFrame([r.campos for r in resultado.registros_71], columns=TITULOS_71),
                     hide_index=True, width="stretch")
with tab_73:
    if resultado.lineas_73:
        st.dataframe(pd.DataFrame(resultado.lineas_73, columns=TITULOS_73), hide_index=True,
                     width="stretch")
    else:
        st.info("El 7.3 se generará vacío (sin información).")
with tab_tot:
    if resultado.tabla_totales:
        st.dataframe(pd.DataFrame(resultado.tabla_totales), hide_index=True, width="stretch",
                     column_config={k: st.column_config.NumberColumn(format="%.2f")
                                    for k in ("Calculado", "Excel (fila Totales)", "Diferencia")})

# ---------------------------------------------------------------- descargas
st.subheader("Descargar")
base_nombre = f"PLE_AF_{cfg.ruc}_{cfg.ejercicio}"
if n_err:
    st.error("Corrige los errores (en Contasis o en los datos de la barra lateral) para descargar los TXT.")
else:
    st.download_button("⬇️ Descargar ZIP con los 3 TXT", data=resultado.zip,
                       file_name=f"{base_nombre}.zip", mime="application/zip", type="primary")
    cols = st.columns(len(resultado.archivos))
    for col, (nombre, datos) in zip(cols, resultado.archivos.items()):
        col.download_button(nombre, data=datos, file_name=nombre, mime="text/plain",
                            help=f"{len(datos):,} bytes" if datos else "Vacío (sin información)")
st.download_button("Descargar reporte de revisión (Excel)", data=reporte_excel(resultado, cfg),
                   file_name=f"{base_nombre}_revision.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.caption(
    "Herramienta interna de apoyo para generar los TXT del PLE a partir de exportaciones de Contasis. "
    "Los archivos se procesan en memoria y no se guardan. Revisa las observaciones y valida los TXT en el "
    "PLE antes de enviarlos; la responsabilidad de la información declarada es del contribuyente."
)
