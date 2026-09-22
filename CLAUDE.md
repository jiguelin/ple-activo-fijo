# Instrucciones para Claude

App Streamlit que convierte el Excel "FORMATO 7.1" de Contasis en los TXT del PLE
(Registro de Activos Fijos 070100, 070300, 070400).

- La especificación completa está en `docs/ESPECIFICACIONES.md`: respétala al hacer cambios.
- Lógica en `ple_af/`, interfaz en `app.py`. No pongas lógica de negocio en `app.py`.
- Importes siempre con `Decimal` (nunca float) y 2 decimales half-up; TC con 3 decimales.
- TXT: campos separados por `|`, `|` final, CRLF, encoding cp1252; archivo sin datos = 0 bytes
  e indicador de contenido 0 en el nombre.
- Solo ejercicios 2022+. El 7.3 va vacío salvo que el usuario active el módulo manual.
- CUO/correlativo (campos 2 y 3) salen del Libro Diario PLE (sub diario 09, glosa "CODIGO - Depreciacion MM"), ver `ple_af/diario.py`.
- Antes de terminar cualquier cambio: `python -m pytest -q` debe pasar.
- Si cambias una regla, actualiza también `docs/ESPECIFICACIONES.md` y agrega una prueba.
