# XLSX_Unlocker
Pyhon scritpt to bypass the worksheet and workbook protection of a excel file.

¿Cómo funciona este script?
Librería zipfile: Los archivos .xlsx son contenedores ZIP. El script abre el archivo original y crea uno nuevo vacío.

Iteración: Copia archivo por archivo (los XML internos) del viejo al nuevo.

Detección (re):

Si el archivo interno es xl/workbook.xml, busca la etiqueta <workbookProtection ...>.

Si el archivo está en xl/worksheets/, busca la etiqueta <sheetProtection ...>.

Eliminación: Si encuentra esas etiquetas, las reemplaza por "nada" (las borra) antes de escribirlas en el nuevo archivo.

Resultado: Genera un archivo _unlocked.xlsx que Excel interpretará como un archivo sin contraseña.

Instrucciones de uso
Asegúrate de tener Python instalado.

Copia el código de arriba y guárdalo en un archivo llamado XLSX_Unlocker.py.

Ejecútalo desde tu terminal:

Bash

python desbloquear.py
El programa te pedirá la ruta del archivo. Puedes arrastrar el archivo Excel a la terminal para que se pegue la ruta automáticamente y presionar Enter.
