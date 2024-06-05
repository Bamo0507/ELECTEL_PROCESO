# ExcelEditor

## Descripción
Este proyecto consiste en un procesador de archivos Excel (.xls) que automatiza la tarea de formatear y exportar datos de hojas de cálculo según ciertos criterios predefinidos. Está diseñado para ser utilizado por la empresa como una herramienta interna para el tratamiento de datos específicos contenidos en archivos Excel.

## Funcionalidades
- **Procesamiento de archivos Excel:** El programa busca archivos con extensión .xls en un directorio especificado y los procesa uno por uno.
- **Formateo de hojas de cálculo:** Formatea automáticamente las hojas de cálculo de acuerdo a ciertos criterios establecidos.
- **Exportación de datos:** Exporta los datos procesados de las hojas de cálculo a archivos de texto (.txt) en un directorio de exportación.
- **Detección de problemas:** Identifica archivos y hojas de cálculo que presentan problemas durante el procesamiento, como falta de datos específicos.

## Requisitos
- Java 8 o superior.
- Apache POI, una biblioteca para leer y escribir archivos de Microsoft Office.

## Instrucciones de Uso
1. Coloque los archivos Excel que desee procesar en el directorio `ProcesadorExcel\ArchivosModificar`.
2. Ejecute el programa. Los archivos procesados se exportarán al directorio `ProcesadorExcel\ArchivosModificados`.
3. Revise la salida del programa para identificar cualquier problema detectado durante el procesamiento.

## Configuración
Puede ajustar la configuración del programa modificando las siguientes variables en la clase `ExcelProcessor`:
- `EXPORT_FOLDER`: Ruta del directorio de exportación para los archivos procesados.
- `EXPECTED_ROWS`: Número esperado de filas en las hojas de cálculo.

## Notas
- Este programa asume que los archivos de Excel están en formato .xls.
- Se recomienda ejecutar el programa en un entorno con suficientes recursos de memoria, especialmente al procesar archivos grandes o un gran número de archivos simultáneamente.

## Problemas Conocidos
- No se han identificado problemas conocidos en la versión actual del programa.

## Contribución
Este proyecto está desarrollado como una solución interna para la empresa y actualmente no aceptamos contribuciones externas. Sin embargo, si encuentra algún problema o tiene sugerencias, no dude en comunicarse con el equipo de desarrollo.
