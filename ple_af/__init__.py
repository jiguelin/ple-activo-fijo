"""Generador de TXT PLE - Registro de Activos Fijos (7.1, 7.3, 7.4)."""
from .config import Config
from .formato_73 import DatoME
from .generador import Resultado, generar, reporte_excel
from .diario import leer_diarios
from .lector_contasis import leer_excel_contasis

__all__ = ["Config", "DatoME", "Resultado", "generar", "reporte_excel", "leer_excel_contasis", "leer_diarios"]
