# Generador PLE – Registro de Activos Fijos

App web (Streamlit) que convierte el Excel **FORMATO 7.1** exportado desde **Contasis** en los
TXT del **PLE de SUNAT** para el Registro de Activos Fijos:

| Archivo | Contenido |
|---|---|
| `LE{RUC}{AAAA}0000070100001111.txt` | 7.1 con los activos del Excel |
| `LE{RUC}{AAAA}0000070300001011.txt` | 7.3 vacío (salvo que actives el llenado manual) |
| `LE{RUC}{AAAA}0000070400001011.txt` | 7.4 siempre vacío |

Solo ejercicios **2022 en adelante**. La especificación completa está en
[`docs/ESPECIFICACIONES.md`](docs/ESPECIFICACIONES.md).

## Cómo se usa

1. Exporta desde Contasis el **FORMATO 7.1** del ejercicio (Excel).
2. Abre la app y súbelo. RUC, razón social y ejercicio se leen del archivo.
3. Sube también el **Libro Diario PLE** (TXT 5.1 o 5.2) del mismo ejercicio; basta el de diciembre.
   De ahí se toma el CUO de cada activo: el asiento de depreciación del sub diario 09 (ej. `12.09.1` / `M1`).
4. (Opcional) Sube también el Excel del año anterior para comprobar saldos iniciales.
5. Revisa la pestaña **Observaciones**:
   - **Errores**: bloquean la descarga (ej. falta fecha de adquisición, código duplicado).
   - **Advertencias**: se puede descargar, pero conviene revisarlas (textos recortados,
     bajas incompletas, totales que no cuadran, activos anteriores a 2013).
   - **Autocompletados**: valores que la app puso (`-` en marca/modelo/serie, fecha de uso, etc.).
6. Descarga el **ZIP con los 3 TXT** e impórtalos en el PLE.
7. El **reporte de revisión (Excel)** muestra línea por línea lo que se generó.

## Qué completa la app automáticamente

| Dato | Valor |
|---|---|
| CUO / correlativo de asiento | Del Libro Diario (sub diario 09): `12.09.1` / `M1`. Sin diario: `AF1` / `M1`… |
| Catálogo / tipo / estado del activo / estado de operación | `9` / `1` / `9` / `1` |
| Marca o modelo vacíos | `-`, como indica SUNAT (opción: `GENERICO`) |
| Serie vacía | `-`, como indica SUNAT (opción: `GENERICO-01`, `GENERICO-02`…) |
| Fecha de inicio de uso vacía | Igual a la fecha de adquisición |
| Método "LINEA RECTA" | `1` |
| % depreciación `0.2` | `20.00` |
| Documento de autorización vacío | `-` (o correlativo `00000001`) |
| Descripción > 40 caracteres | Se recorta conservando el correlativo final (` 0001`) |

## Desarrollo local (opcional)

```bash
pip install -r requirements-dev.txt
streamlit run app.py
python -m pytest -q
```

## Contraseña (opcional)

Si defines el secreto `APP_PASSWORD` (en Streamlit Cloud → Settings → Secrets), la app pide
contraseña antes de mostrarse:

```toml
APP_PASSWORD = "tu-clave"
```
