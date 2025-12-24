# 🔓 XLSX Unlocker
XLSX Unlocker es una herramienta ligera de automatización diseñada para eliminar instantáneamente la protección de libros (Workbook) y hojas (Worksheet) en archivos de Microsoft Excel (.xlsx).

A diferencia de otros métodos, este script no intenta "romper" la contraseña por fuerza bruta; en su lugar, modifica la estructura XML interna del archivo para desactivar las restricciones de edición.

🚀 Características
Sin dependencias externas: Solo utiliza librerías estándar de Python (zipfile, re, os).

No destructivo: Crea una copia del archivo (_unlocked.xlsx), manteniendo tu archivo original intacto.

Rápido y eficiente: Procesa archivos de gran tamaño en milisegundos al trabajar directamente con el flujo de datos ZIP.

🛠️ ¿Cómo funciona?
Técnicamente, un archivo .xlsx es un archivo comprimido que contiene múltiples documentos XML siguiendo el estándar OpenXML. El script realiza el siguiente proceso:

Lectura de Contenedor: Abre el archivo como un objeto zip.

Análisis de Estructura:

Accede a xl/workbook.xml para buscar la etiqueta <workbookProtection>.

Escanea la carpeta xl/worksheets/ buscando etiquetas <sheetProtection> en cada hoja.

Inyección de Datos: Utiliza expresiones regulares para eliminar quirúrgicamente estas etiquetas sin corromper el resto del formato.

Re-empaquetado: Genera un nuevo contenedor ZIP con la extensión .xlsx listo para editarse sin restricciones.

📋 Requisitos
Python 3.6 o superior.

No requiere instalar librerías mediante pip.

💻 Uso
Clona este repositorio o descarga el archivo XLSX_Unlocker.py.

Abre una terminal y ejecuta el script:

python XLSX_Unlocker.py

Cuando el programa lo solicite, arrastra y suelta tu archivo Excel bloqueado en la terminal (o escribe la ruta manualmente).

¡Listo! Encontrarás una versión desbloqueada en la misma carpeta que el original.

⚠️ Nota de Seguridad
Este script está diseñado para fines de recuperación de archivos propios o flujos de trabajo administrativos donde se ha perdido la clave de edición. No debe utilizarse para violar la privacidad de documentos protegidos con contraseña de apertura (encriptación), ya que este script solo remueve la protección contra escritura/edición.
