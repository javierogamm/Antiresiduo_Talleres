# Log de cambios

## v1.10
- Se interpreta automáticamente la primera fila de la columna C como `ExpedienteCodigo` cuando el encabezado viene vacío en el `.xlsx` de entrada.
- Se actualizó la versión mostrada en la interfaz.

## v1.9
- Se sustituyó la petición manual del programa por un selector desplegable con las opciones `CAAG`, `GFD` y `CAZ` antes de cargar archivos `.xlsx`.
- Se renombró el programa `CAG` a `CAAG` en el mapeo de certificaciones.
- Se ajustó la detección visual de errores para ignorar la columna `ValorCampoAdicional`.
- Se cambió el resaltado de errores para colorear la fila completa cuando cualquier campo relevante falla.
- Se actualizó la versión mostrada en la interfaz.

## v1.8
- Se actualizaron los inputs normalizados de talleres para los programas `CAG`, `CAZ` y `GFD` según la tabla consolidada.
- Se sustituyó el programa `ADD` por `CAZ` para la selección y mapeo de certificados de Analiza.
- Se ajustaron los inputs de `GFD` para aceptar `Kickoff` y `Taller 01` a `Taller 08` sin prefijo del programa.
- Se actualizó la versión mostrada en la interfaz.

## v1.7
- Se añadió la selección obligatoria del programa de certificación al cargar archivos `.xlsx`, con opciones `GFD`, `ADD` y `CAG`.
- Se separaron los mapeos de talleres por programa para aplicar el nombre de tarea correspondiente según la selección realizada.
- Se incorporaron advertencias en pantalla para filas con datos obligatorios faltantes y una visualización con fondo rojo para los campos vacíos.
- Se actualizó la versión mostrada en la interfaz.

## v1.6
- Se incorporó el mapeo completo de entradas de talleres AGG, ADD y GFD al nombre de tarea requerido para la salida CSV.
- Se consolidó el tipo de campo tesauro de la salida como `Texto` para todos los registros generados.
- Se actualizó la versión mostrada en la interfaz.

## v1.5
- Se ajustó el encabezado del CSV UTF-8 para respetar espacios y mayúsculas requeridos.
- Se actualizó la versión mostrada en la interfaz.

## v1.4
- Se ajustó la exportación CSV para usar separador de punto y coma, salto de línea CRLF y escape de valores para visualizar columnas correctamente en Excel.
- Se actualizó la versión mostrada en la interfaz.

## v1.3
- Se modernizó la interfaz con nuevo layout, tipografía y botones estilizados.
- Se reorganizó el contenido en tarjetas para mejorar la lectura del resultado.
- Se actualizó la versión visible en la interfaz.

## v1.2
- Se generó el CSV sin entrecomillar valores en la exportación.
- Se actualizó la versión mostrada en la interfaz.

## v1.1
- Se añadió la descarga de la tabla en formato CSV con codificación UTF-8.
- Se actualizó la interfaz para mostrar la nueva versión y el botón de descarga CSV.
