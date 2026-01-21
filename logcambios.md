# Log de cambios

## v1.0.1 - 2026-01-21
- Restaura la lógica original de transformación para entradas "Taller".
- Añade los nuevos mapeos GFD/ADD en los mismos diccionarios legacy.
- Mantiene el comportamiento previo y amplía los casos soportados.

## v1.1.2 - 2026-01-21
- Añade los nuevos mapeos GFD/ADD directamente a MAPEO_TESAUROS y MAPEO_CERTIFICADOS.
- Unifica la lógica para resolver por "Nombre input" o "Taller" sin eliminar compatibilidad legacy.
- Actualiza la versión visible de la APP.

## v1.1.1 - 2026-01-21
- Reincorpora el mapeo legacy por "Taller" y mantiene los nuevos inputs GFD/ADD.
- Permite generar salidas tanto por "Nombre input" como por el código de taller previo.
- Actualiza la versión visible de la APP.

## v1.1.0 - 2026-01-21
- Actualiza los mapeos para nuevas entradas de GFD y ADD con sus certificados y campos de tesauro.
- Ajusta el procesamiento para leer la columna "Nombre input" y generar filas según el tipo de campo.
- Actualiza la versión visible de la APP.
