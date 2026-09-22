"""Parámetros de generación y tablas SUNAT."""
from __future__ import annotations

from dataclasses import dataclass

EJERCICIO_MINIMO = 2022

# Tabla 13 - Catálogo de existencias (se usa 9 = Otros)
CATALOGO_OTROS = "9"
# Tabla 18 - Tipo de activo fijo: 1 = No revaluado o revaluado sin efecto tributario
TIPO_ACTIVO = "1"
# Tabla 19 - Estado del activo fijo: 9 = Resto de activos
ESTADO_ACTIVO = "9"
# Tabla 20 - Método de depreciación
METODOS = {"LINEA RECTA": "1", "UNIDADES PRODUCIDAS": "2"}
METODO_OTROS = "9"
# Estado de la operación: 1 = corresponde al periodo
ESTADO_OPERACION = "1"

LIBRO_71 = "070100"
LIBRO_73 = "070300"
LIBRO_74 = "070400"

# Longitudes máximas SUNAT
LARGO_DESCRIPCION = 40
LARGO_MARCA = 20
LARGO_MODELO = 20
LARGO_SERIE = 30
LARGO_CODIGO = 24
LARGO_CUENTA = 24
LARGO_DOC_AUTORIZACION = 20
LARGO_CUO = 40


@dataclass
class Config:
    ruc: str
    ejercicio: int
    razon_social: str = ""
    formato_periodo: str = "AAAA0000"      # o "AAAA1200"
    codigo_oportunidad: str = "00"
    indicador_operaciones: str = "1"       # 1 operativa, 0 cierre, 2 cierre del libro
    moneda: str = "1"                      # 1 soles
    prefijo_cuo: str = "AF"
    prefijo_asiento: str = "M"             # A, M o C
    texto_faltante: str = "GENERICO"       # marca/modelo faltante: "GENERICO" o "-"
    serie_faltante: str = "GENERICO-nn"    # "GENERICO-nn" o "-"
    campo27: str = "-"                     # "-" o "correlativo"
    subdiario: str = "09"                  # sub diario de Contasis con la depreciación
    linea_correlativo: str = "68"          # "68" (línea del gasto) o "39" (dep. acumulada)

    @property
    def periodo(self) -> str:
        if self.formato_periodo == "AAAA1200":
            return f"{self.ejercicio}1200"
        return f"{self.ejercicio}0000"
