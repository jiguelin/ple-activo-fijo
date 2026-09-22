"""Registro de errores, advertencias y valores completados automáticamente."""
from __future__ import annotations

from dataclasses import dataclass, field

ERROR = "Error"
ADVERTENCIA = "Advertencia"
AUTOCOMPLETADO = "Autocompletado"


@dataclass
class Observacion:
    tipo: str               # Error / Advertencia / Autocompletado
    mensaje: str
    fila: int | None = None
    codigo: str = ""
    campo: str = ""
    columna: str = ""


@dataclass
class Observaciones:
    items: list[Observacion] = field(default_factory=list)

    def agregar(self, tipo, mensaje, fila=None, codigo="", campo="", columna=""):
        self.items.append(Observacion(tipo, mensaje, fila, codigo or "", campo, columna))

    def error(self, *a, **k):
        self.agregar(ERROR, *a, **k)

    def advertencia(self, *a, **k):
        self.agregar(ADVERTENCIA, *a, **k)

    def autocompletado(self, *a, **k):
        self.agregar(AUTOCOMPLETADO, *a, **k)

    def extender(self, otras: "Observaciones"):
        self.items.extend(otras.items)

    def de_tipo(self, tipo):
        return [o for o in self.items if o.tipo == tipo]

    @property
    def hay_errores(self) -> bool:
        return any(o.tipo == ERROR for o in self.items)

    def como_filas(self) -> list[dict]:
        return [
            {"Tipo": o.tipo, "Fila Excel": o.fila, "Columna": o.columna, "Código": o.codigo,
             "Campo": o.campo, "Detalle": o.mensaje}
            for o in self.items
        ]
