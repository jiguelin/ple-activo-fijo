# Generador PLE – Registro de Activos Fijos (Formatos 7.1, 7.3 y 7.4)

Especificación funcional y técnica para que Claude (Claude Code) programe la app.
Repositorio en GitHub + despliegue en Streamlit Community Cloud (streamlit.io).

---

## 1. Objetivo

Subir el Excel "FORMATO 7.1" exportado desde **Contasis** y obtener los archivos TXT del
**Registro de Activos Fijos** listos para validar y enviar en el **PLE de SUNAT**:

| Libro PLE | Código | Contenido en esta app |
|---|---|---|
| 7.1 Detalle de los activos fijos revaluados y no revaluados | `070100` | Con datos (desde el Excel de Contasis) |
| 7.3 Detalle de la diferencia de cambio | `070300` | **Vacío por defecto** (indicador 0), aunque las compras hayan sido en dólares. Llenado opcional solo si el contador lo activa (ver sección 6) |
| 7.4 Activos fijos bajo arrendamiento financiero al 31.12 | `070400` | Siempre vacío (indicador 0) |

**Alcance temporal:** solo ejercicios **2022 en adelante**. La app rechaza ejercicios < 2022.

> Nota: en el PLE **no existe el formato 7.2** (solo aparece en el formato impreso/hojas sueltas,
> porque en el PLE los revaluados y no revaluados van juntos en el 7.1). En el PLE se presentan
> tres archivos: 070100, 070300 y 070400.

Base normativa: R.S. 286-2009/SUNAT y modificatorias (anexos de estructuras R.S. 361-2015,
R.S. 108-2020 y Anexo I de la R.S. 315-2018 para el 7.1). Tablas 13, 18, 19 y 20 del Anexo 3.

---

## 2. Stack y estructura del repositorio

- Python 3.11, `streamlit`, `pandas`, `openpyxl`, `xlrd` (por si llega un `.xls`), `pytest`.
- Sin base de datos. Todo se procesa en memoria; nada se guarda en el servidor.

```
ple-activos-fijos/
├── app.py                  # Interfaz Streamlit
├── ple_af/
│   ├── __init__.py
│   ├── lector_contasis.py  # Lee y normaliza el Excel de Contasis
│   ├── config.py           # Parámetros (Config) y tablas SUNAT 13/18/19/20
│   ├── utilidades.py       # Números (Decimal), fechas, limpieza de texto, RUC
│   ├── observaciones.py    # Errores / advertencias / autocompletados
│   ├── formato_71.py       # Construye filas 7.1 (37 campos)
│   ├── formato_73.py       # Construye filas 7.3 (15 campos, módulo manual)
│   ├── controles.py        # Totales y control con el año anterior
│   ├── archivos.py         # Nombre LE...txt, escritura TXT, ZIP
│   └── generador.py        # Orquesta todo + reporte Excel de revisión
├── tests/
│   ├── fixtures/           # 8 exportaciones reales de Contasis + TXT 2021
│   └── test_generador.py
├── docs/ESPECIFICACIONES.md
├── requirements.txt
├── README.md
└── CLAUDE.md
```

---

## 3. Archivo de entrada (Excel de Contasis "FORMATO 7.1")

Estructura **confirmada con 8 exportaciones reales** de Contasis (4 empresas, ejercicios
2022–2025: Essentta 2022 y 2024, Contadeus International 2023, Dental Salud 2022–2025,
Econobrillo 2025). Todas son idénticas:

- Una sola hoja llamada `FORMATO 7.1`, 28 celdas combinadas en la cabecera.
- Cabecera: `B3` = "EJERCICIO: 2024", `B4` = "R.U.C: 20565449843", `B5` = "RAZÓN SOCIAL: ...",
  `B6` = giro (informativo).
- Títulos de columnas en filas 8 a 11 (celdas combinadas). Datos desde la fila 12.
- Tras el último activo: fila con "Totales: " en la columna G (sumas en H..AA) y luego
  3 filas vacías. La fila de totales **no es un activo**, se usa solo para control.
- Tipos de dato reales: B texto (con ceros iniciales); C entero con formato texto; D/E/G texto;
  F y G a veces **número** (modelo `2017`, serie `202109252`) → convertir a texto sin `.0`;
  importes numéricos con formato `#,##0.0000` (en la práctica, 2 decimales);
  P/Q `datetime`; R texto ("LINEA RECTA" en todos los casos vistos); S siempre vacío;
  T número en formato % (`0.2`).
- Convención de signos de Contasis: la **depreciación de bajas (W) viene en negativo**
  (ej. −4,010.48), de modo que `Y = U + V + W + X`. Se trasladan los signos tal cual al TXT
  (SUNAT admite positivo o negativo en esos campos).
- Casos de muestra reales para pruebas:
  - Descripciones de hasta 55 caracteres (límite 40) en Dental y Contadeus.
  - Modelo de 25 caracteres (`GRETA 1.6 GL 2WD MT SPORT`, límite 20).
  - Activos adquiridos en 2012 (Dental) → dispara la advertencia del 7.3.
  - Baja en Dental 2022 (activo `060900050001`): W = −4,010.48 pero K (retiro del costo) = 0
    y el costo sigue en 5,839.00 → debe salir como **advertencia** de baja incompleta.

| Col | Contenido Contasis | Campo PLE 7.1 |
|---|---|---|
| B | Código relacionado con el activo fijo | 5 |
| C | Cuenta contable | 9 |
| D | Descripción | 11 |
| E | Marca | 12 |
| F | Modelo | 13 |
| G | Número de serie y/o placa | 14 |
| H | Saldo inicial | 15 |
| I | Adquisiciones / adiciones | 16 |
| J | Mejoras | 17 |
| K | Retiros y/o bajas | 18 |
| L | Otros ajustes | 19 |
| M | Valor histórico al 31.12 | (solo control: H+I+J+K+L si K viene negativo, o H+I+J−K+L si viene positivo; aceptar cualquiera de las dos) |
| N | Ajuste por inflación | 23 |
| O | Valor ajustado al 31.12 | (solo control) |
| P | Fecha de adquisición | 24 |
| Q | Fecha de inicio de uso | 25 |
| R | Método aplicado (texto, ej. "LINEA RECTA") | 26 (convertir a código) |
| S | N° documento de autorización | 27 |
| T | Porcentaje de depreciación (ej. 0.2) | 28 (×100 → 20.00) |
| U | Depreciación acumulada al cierre del ejercicio anterior | 29 |
| V | Depreciación del ejercicio | 30 |
| W | Depreciación del ejercicio relacionada con retiros/bajas | 31 |
| X | Depreciación relacionada con otros ajustes | 32 |
| Y | Depreciación acumulada histórica | (solo control: U+V+W+X, con W negativo) |
| Z | Ajuste por inflación de la depreciación | 36 |
| AA | Depreciación acumulada ajustada | (solo control) |

Reglas de lectura:

1. **Detectar la fila de títulos y las columnas por texto** (no por posición fija), buscando
   palabras clave ("CÓDIGO RELACIONADO", "SALDO INICIAL", "FECHA DE ADQUISICIÓN", etc.),
   con respaldo a la posición estándar B..AA si no se encuentran.
2. **Leer el código (col. B) y la cuenta (col. C) como texto** para no perder ceros a la izquierda
   (`040100020001` no debe convertirse en `40100020001`).
3. Leer filas hasta la primera fila vacía en B o hasta la fila "Totales:".
4. Fechas: aceptar fecha de Excel, `datetime`, texto `DD/MM/AAAA` o `AAAA-MM-DD`.
5. Números: aceptar vacío, `0`, texto con comas; vacío = 0.
6. Aceptar `.xlsx`, `.xls` y `.xlsm`.
7. Extraer RUC, razón social y ejercicio de la cabecera; permitir editarlos en pantalla.

---

## 4. Parámetros en pantalla (barra lateral)

| Parámetro | Valor por defecto | Notas |
|---|---|---|
| RUC | Leído de `B4` | 11 dígitos, validar dígito verificador |
| Razón social | Leída de `B5` | Solo informativo |
| Ejercicio (AAAA) | Leído de `B3` | ≥ **2022** |
| Formato del campo 1 (periodo) | `AAAA0000` | Opción `AAAA1200`. Ver nota en sección 5 |
| Código de oportunidad (CC) | `00` | |
| Indicador de operaciones (O) | `1` (empresa operativa) | 0 = cierre/baja, 2 = cierre del libro |
| Moneda (M) | `1` (soles) | |
| Libro Diario PLE (TXT) | — | Uno o varios meses del mismo ejercicio; basta diciembre. Fuente del CUO |
| Sub diario de depreciación | `09` | |
| Línea del correlativo | `68` | Opción `39` |
| Prefijo de CUO (sin diario) | `AF` | Ver campo 2 |
| Prefijo correlativo de asiento | `M` | Debe empezar con A, M o C |
| Texto por defecto marca / modelo | `GENERICO` | La estructura SUNAT indica `-` si no existe; ambas opciones en pantalla |
| Serie por defecto | `GENERICO-01`, `GENERICO-02`… | O `-` |
| Campo 27 por defecto | `-` | Opción: correlativo de 8 dígitos (`00000001`) como en los TXT 2021 |
| Llenar 7.3 manualmente | No | Si "Sí", se habilita la tabla del 7.3 (sección 6) |

---

## 5. Formato 7.1 – construcción de cada línea (37 campos)

| # | Campo | Regla |
|---|---|---|
| 1 | Periodo | `AAAA0000` (ej. `20240000`). Configurable (ver nota abajo) |
| 2 | CUO | **Asiento del Libro Diario** donde se registró la depreciación del activo (sub diario 09 de Contasis), tomado del TXT del diario PLE (5.1/5.2) que sube el usuario: se busca la línea cuya glosa empieza con el código del activo (`040100020001 - Depreciacion 12`) y se usa el CUO del **último mes** disponible (ej. `12.09.1`). Sin diario: `AF{n}` con advertencia |
| 3 | Correlativo del asiento | Correlativo de la línea de ese asiento con la cuenta **68** (por defecto, ej. `M1`) o **39** (opción, ej. `M4`). Sin diario: `M{n}` |
| 4 | Código de catálogo | `9` (Otros – Tabla 13) |
| 5 | Código propio del activo | Col. B (texto, máx. 24). Obligatorio: si falta → error |
| 6 | Catálogo existencia (UNSPSC/GTIN) | Vacío (solo si el comprobante trae UNSPSC/GTIN; no acepta 9) |
| 7 | Código de existencia | Vacío |
| 8 | Tipo de activo fijo (Tabla 18) | `1` = No revaluado o revaluado sin efecto tributario |
| 9 | Cuenta contable | Col. C (solo dígitos, máx. 24) |
| 10 | Estado del activo (Tabla 19) | `9` = Resto de activos (1 desuso, 2 obsoleto) |
| 11 | Descripción | Col. D, máx. **40**. Contasis termina cada descripción con un correlativo (` 0001`); al truncar, **conservar ese sufijo**: `texto[:35] + " 0001"` (ej. "PIEZA DE MANO DENTAL MAS ACCESORIOS MARCA PANA MAX 0001" → "PIEZA DE MANO DENTAL MAS ACCESORIOS 0001"). Advertir |
| 12 | Marca | Col. E, máx. 20; si vacío → valor por defecto |
| 13 | Modelo | Col. F, máx. 20; si vacío → valor por defecto; si es número (2017) → texto "2017" |
| 14 | Serie / placa | Col. G, máx. 30; si vacío → `GENERICO-nn` o `-` |
| 15 | Saldo inicial | Col. H |
| 16 | Adquisiciones / adiciones | Col. I |
| 17 | Mejoras | Col. J |
| 18 | Retiros y/o bajas | Col. K |
| 19 | Otros ajustes | Col. L |
| 20 | Revaluación voluntaria | `0.00` |
| 21 | Revaluación por reorganización | `0.00` |
| 22 | Otras revaluaciones | `0.00` |
| 23 | Ajuste por inflación | Col. N |
| 24 | Fecha de adquisición | Col. P, `DD/MM/AAAA`, ≤ 31/12/AAAA |
| 25 | Fecha de inicio de uso | Col. Q, `DD/MM/AAAA`, ≤ 31/12/AAAA; si vacío → = campo 24 (advertir) |
| 26 | Método (Tabla 20) | "LINEA RECTA" → `1`, "UNIDADES PRODUCIDAS" → `2`, otro → `9`; vacío → `1` |
| 27 | N° documento autorización | Col. S; si vacío → valor por defecto del parámetro |
| 28 | % depreciación | Col. T; si ≤ 1 multiplicar ×100; formato `###.##`; 0 ≤ % ≤ 100 |
| 29 | Depreciación acumulada ejercicio anterior | Col. U |
| 30 | Depreciación del ejercicio | Col. V |
| 31 | Depreciación de retiros/bajas | Col. W |
| 32 | Depreciación de otros ajustes | Col. X |
| 33–35 | Depreciación de revaluaciones | `0.00` |
| 36 | Ajuste por inflación de la depreciación | Col. Z |
| 37 | Estado de la operación | `1` (operación del periodo) |

**Nota sobre el campo 1:** el anexo de la R.S. 108-2020 describe el campo como `AAAAMM00`,
mientras que los TXT anuales usados por el estudio y la macro 2024 usan `AAAA0000` (mes `00`
por ser libro anual, igual que en el nombre del archivo). Por defecto se usa `AAAA0000` y queda
configurable. **Validado:** los 3 TXT de Essentta 2024 (campo 1 `20240000`, campo 27 `-`, CUO `AF{n}`, marca/modelo `GENERICO`) pasaron "Sin errores" en el PLE 5.2.0.7, modo de prueba (22/09/2026).

**Qué se genera y qué bloquea:** los campos que Contasis nunca trae (2, 3, 4, 8, 10, 37) se
generan con valores fijos, porque siempre faltan. Solo bloquean la fila los datos propios del
activo: código (5), cuenta (9), descripción (11), fecha de adquisición (24) y porcentaje (28).

Ejemplo real (TXT 2021 aceptado):

```
20210000|12.09.1|M1|9|040100020001|||1|33411|9|AUTO HONDA 4X4 0001|HONDA|CR-V|BTD-410|107914.01|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|21/11/2020|21/11/2020|1|00000001|20.00|0.00|7553.91|0.00|0.00|0.00|0.00|0.00|0.00|1|
```

---

## 6. Formato 7.3 – Diferencia de cambio (15 campos)

**Regla por defecto (ejercicios 2022+): el 7.3 se genera VACÍO** (0 bytes, indicador de
contenido `0`), aunque los activos se hayan comprado en dólares.

Motivo: el inciso f) del art. 61 de la LIR (que llevaba al costo del activo la diferencia de
cambio de pasivos vinculados a activos fijos) fue derogado por el D. Leg. 1112 desde el
01/01/2013. Desde entonces esa diferencia de cambio va a resultados y no hay "ajuste por
diferencia de cambio del activo" que informar. La moneda de la factura, por sí sola, no obliga
a llenar el 7.3; el costo en soles (al TC de la fecha de compra) va en el 7.1.

**Advertencia automática:** si algún activo tiene fecha de adquisición anterior al 01/01/2013,
mostrar: "Revise si el costo de este activo incluye diferencias de cambio capitalizadas antes
de 2013; en ese caso podría corresponder informar el 7.3".

**Módulo opcional (desactivado por defecto):** interruptor "Llenar 7.3 manualmente". Solo si el
contador lo activa, mostrar un `st.data_editor` con la lista de activos del 7.1
para que marque cuáles fueron en ME y complete:

- Valor de adquisición en ME (ej. 1,439.83 USD)
- Tipo de cambio en la fecha de adquisición (TC venta SBS/SUNAT de esa fecha)
- Un solo campo general: Tipo de cambio al 31.12 del ejercicio

Opcional (fase 2): aceptar esas columnas en una hoja adicional "ME" del Excel subido
(código, valor ME, TC adquisición), para no tipearlo cada año.

| # | Campo | Regla |
|---|---|---|
| 1 | Periodo | `AAAA0000` |
| 2 | CUO | Correlativo propio del 7.3 (máx. 40) |
| 3 | Correlativo asiento | `M{n}` |
| 4 | Catálogo | `9` (no acepta `1`) |
| 5 | Código propio del activo | Mismo código del 7.1 |
| 6 | Fecha de adquisición | Del 7.1 (campo 24) |
| 7 | Valor adquisición en ME | Ingresado por el usuario |
| 8 | TC en fecha de adquisición | Ingresado, `#.###` (3 decimales) |
| 9 | Valor adquisición en MN | Del 7.1 (costo: saldo inicial + adquisiciones). Advertir si difiere de campo 7 × campo 8 en más de S/ 1.00 |
| 10 | TC al 31.12 | Ingresado, `#.###`; si no hay → `0.000` |
| 11 | Ajuste por diferencia de cambio | `0.00` por defecto (desde 2013 la diferencia de cambio no se capitaliza al activo, va a resultados). Editable |
| 12 | Depreciación del ejercicio | Del 7.1 (campo 30). Si el usuario la edita y difiere del 7.1, **advertir** (no corregir) |
| 13 | Depreciación de retiros/bajas | Del 7.1 (campo 31) |
| 14 | Depreciación de otros ajustes | Del 7.1 (campo 32) |
| 15 | Estado de la operación | `1` |

Ejemplo real (TXT 2021):

```
20210000|12.09.33|M1|9|060900040001|30/06/2021|1439.83|3.932|5661.41|3.979|0.00|330.25|0.00|0.00|1|
```

---

## 7. Formato 7.4 – Arrendamiento financiero

Siempre archivo **vacío (0 bytes)** con indicador de contenido `0`.

---

## 8. Nombre de los archivos

`LE` + RUC(11) + AAAA + `00` (mes) + `00` (día) + LLLLLL + CC + O + I + M + G + `.txt`

- LLLLLL: `070100`, `070300`, `070400`
- CC: código de oportunidad (`00`)
- O: indicador de operaciones (`1`)
- I: indicador de contenido: `1` con información, `0` sin información
- M: moneda (`1` soles)
- G: `1` (generado por PLE)

Ejemplos (coinciden con los archivos 2021 adjuntos):

```
LE2056544984320210000070100001111.txt   (7.1 con datos)
LE2056544984320210000070300001111.txt   (7.3 con datos)
LE2056544984320210000070400001011.txt   (7.4 vacío)
```

---

## 9. Formato del TXT

- Separador `|`, **con `|` al final de cada línea**.
- Sin fila de títulos.
- Fin de línea `CRLF` (`\r\n`), incluida la última línea.
- Importes: 2 decimales, punto decimal, sin separador de miles, redondeo "half-up"
  (usar `Decimal`, no `float`). Vacío o cero → `0.00`.
- Tipos de cambio: 3 decimales (`3.932`).
- Fechas: `DD/MM/AAAA`.
- Encoding: ANSI / `cp1252` (sin BOM). Limpiar texto: quitar `|`, saltos de línea,
  tabulaciones, espacios dobles, caracteres no imprimibles; pasar a MAYÚSCULAS;
  reemplazar caracteres fuera de `cp1252` (ej. `ñ`/`Ñ` sí se permiten; comillas tipográficas → `"`).
- Archivo sin información: 0 bytes.

---

## 10. Validaciones

**Errores (bloquean la descarga):**

- RUC inválido o ejercicio < 2022.
- Activo sin código (col. B) o código duplicado.
- Activo sin cuenta contable (col. C) o sin descripción (col. D).
- Fecha de adquisición vacía o posterior al 31/12 del ejercicio.
- % de depreciación fuera de 0–100 (vacío → 0.00 con advertencia, por terrenos).
- Cada error indica la fila y la columna del Excel de origen.
- En 7.3: activo marcado en ME sin valor ME o sin TC.

**Advertencias (permiten descargar, se listan en pantalla y en el reporte):**

- Descripción, marca, modelo o serie truncados por longitud
  (ej. "ZHIYUN CRANE 4 COMBO, SMALLRING TRIPODE AD-80 0001" tiene 50 caracteres → 40).
- Valores por defecto aplicados (marca/modelo/serie/fecha de uso/método/campo 27).
- Costo ≠ col. M, o depreciación U+V+W+X ≠ col. Y (diferencia > 0.01).
- Baja incompleta: W ≠ 0 y K = 0 (se dio de baja la depreciación pero no el costo), o K ≠ 0 y W = 0.
- Fecha de inicio de uso anterior a la fecha de adquisición.
- Depreciación acumulada mayor que el costo.
- Activo con fecha de adquisición anterior a 2013 (posible diferencia de cambio capitalizada → revisar 7.3).
- Totales calculados ≠ fila "Totales:" del Excel.
- Sin Libro Diario subido, o activo sin asiento en el sub diario 09 → CUO propio con advertencia.
- Si se suben los 12 meses del diario: depreciación del ejercicio (col. V) ≠ suma de la cuenta 39 del activo en el diario.
- Diario de otro RUC u otro ejercicio → **error**.

**Control opcional con el ejercicio anterior (recomendado):** permitir subir también el Excel
de Contasis del año anterior y comparar por código de activo:

- Saldo inicial (H) del año ≠ valor histórico (M) del año anterior.
- Depreciación acumulada anterior (U) del año ≠ depreciación acumulada (Y) del año anterior.
- Activos que estaban el año anterior y ya no aparecen, sin baja registrada.

Esto ya encuentra problemas reales en los archivos de muestra: Dental Salud 2024 → 2025 tiene
los 31 activos con U(2025) ≠ Y(2024) (ej. vehículo: 59,285.36 vs 58,287.29), en 2023 → 2024 el
activo `060900060001` cambia de costo (3,696.05 → 2,617.66) sin movimiento, y en 2022 → 2023
desaparecen `060100050001` y `060900050001` sin retiro del costo.

---

## 11. Interfaz (app.py)

1. Título y breve instrucción.
2. `st.file_uploader` para el Excel de Contasis.
3. Barra lateral con los parámetros de la sección 4 (precargados desde la cabecera).
4. Vista previa de los activos leídos (tabla) con resaltado de celdas que recibieron valor por defecto.
5. Sección 7.3: interruptor "¿Activos en moneda extranjera?" → `st.data_editor` + TC al 31.12.
6. Panel de validaciones (errores en rojo, advertencias en amarillo).
7. Totales de control (como la hoja "3_Generacion" de la macro): saldo inicial, adquisiciones,
   mejoras, retiros, otros ajustes, ajuste por inflación, depreciación acumulada anterior,
   depreciación del ejercicio, cantidad de activos.
8. Botones de descarga:
   - **ZIP con los 3 TXT** (principal)
   - Cada TXT por separado
   - Reporte Excel de revisión (hoja "7.1", hoja "7.3", hoja "Advertencias", hoja "Totales")

---

## 12. Pruebas (pytest)

- Fixtures: las 8 exportaciones reales de Contasis. Cantidad de activos esperada:
  Essentta 2022 = 33, Essentta 2024 = 35, Contadeus 2023 = 7, Dental 2022 = 28, 2023 = 28,
  2024 = 31, 2025 = 31, Econobrillo 2025 = 5.
- Totales de cada archivo = fila "Totales:" (ej. Essentta 2024: saldo inicial 172,507.54,
  adquisiciones 4,152.54, depreciación del ejercicio 22,836.81).
- Dental 2022: advertencia de baja incompleta en `060900050001`; advertencias pre-2013.
- Dental 2025 + Dental 2024 como año anterior → 31 advertencias de continuidad.
- Serie numérica `202109252` → `202109252` (sin `.0`); modelo `2017` → `2017`.
- Código `040100020001` conserva el cero inicial.
- Porcentaje `0.2` → `20.00`; "LINEA RECTA" → `1`.
- Marca vacía → `GENERICO`; descripción de 50 caracteres → 40.
- Línea generada tiene exactamente 37 campos + `|` final (7.1) y 15 + `|` (7.3).
- Nombre de archivo 7.4 vacío = `LE{RUC}{AAAA}0000070400001011.txt` y tamaño 0 bytes.
- Formato de TXT igual al TXT 2021 adjunto (misma estructura, CRLF, pipe final).
- 7.3 sin activar el módulo → 0 bytes e indicador `0` en el nombre.
- Ejercicio 2021 → error.
- Prueba final: validar en el PLE los 3 TXT del ejercicio 2024 de Essentta y guardar ese caso
  como prueba de regresión.

---

## 13. Despliegue

1. Crear repo en GitHub (puede ser privado).
2. `requirements.txt`: `streamlit`, `pandas`, `openpyxl`, `xlrd`.
3. En share.streamlit.io → "New app" → elegir repo, rama `main`, archivo `app.py`.
4. (Opcional) Proteger con contraseña usando `st.secrets` si la app queda pública.

---

## 14. Fuera de alcance (por ahora)

- Cálculo de la depreciación (se toma lo que da Contasis).
- Revaluaciones (campos 20–22 y 33–35 en cero).
- Formato 7.4 con datos (leasing).
- Consulta automática del tipo de cambio SUNAT/SBS (posible fase 2).
