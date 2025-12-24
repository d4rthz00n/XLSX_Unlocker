import zipfile
import re
import os
import shutil

def desbloquear_excel(ruta_entrada, ruta_salida):
    """
    Toma un archivo Excel (.xlsx), elimina las protecciones de Workbook y Worksheet,
    y guarda una copia desbloqueada.
    """
    
    # Patrones Regex para encontrar y eliminar las etiquetas de protección
    # Busca <workbookProtection ... /> y <sheetProtection ... />
    patron_workbook = re.compile(rb'<workbookProtection[^>]*>')
    patron_sheet = re.compile(rb'<sheetProtection[^>]*>')

    try:
        print(f"--- Procesando: {os.path.basename(ruta_entrada)} ---")
        
        # Abrimos el archivo original (lectura) y creamos el nuevo (escritura)
        with zipfile.ZipFile(ruta_entrada, 'r') as zin:
            with zipfile.ZipFile(ruta_salida, 'w') as zout:
                
                # Iteramos sobre cada archivo dentro del Excel (que es un zip)
                for item in zin.infolist():
                    contenido = zin.read(item.filename)
                    
                    # 1. Modificar workbook.xml (Protección del libro/estructura)
                    if item.filename == 'xl/workbook.xml':
                        if patron_workbook.search(contenido):
                            print(f"  [+] Protección de libro encontrada y eliminada en: {item.filename}")
                            # Reemplazamos la etiqueta por vacío
                            contenido = patron_workbook.sub(b'', contenido)
                    
                    # 2. Modificar hojas de trabajo (Protección de celdas/hoja)
                    # Usualmente están en xl/worksheets/sheet1.xml, sheet2.xml, etc.
                    elif 'worksheets/' in item.filename and item.filename.endswith('.xml'):
                        if patron_sheet.search(contenido):
                            print(f"  [+] Protección de hoja encontrada y eliminada en: {item.filename}")
                            contenido = patron_sheet.sub(b'', contenido)
                    
                    # Escribimos el contenido (modificado o no) en el nuevo archivo
                    zout.writestr(item, contenido)
        
        print(f"--- ¡Éxito! Archivo guardado en: {ruta_salida} ---\n")
        return True

    except Exception as e:
        print(f"Error procesando el archivo: {e}")
        return False

# --- BLOQUE PRINCIPAL PARA EJECUTAR ---
if __name__ == "__main__":
    # Puedes cambiar esta ruta o pedirla al usuario
    input_path = input("Ingresa la ruta completa del archivo Excel bloqueado (ej: C:\\docs\\archivo.xlsx): ").strip().strip('"')
    
    if os.path.exists(input_path):
        # Crear nombre de salida (ej: archivo_unlocked.xlsx)
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_unlocked{ext}"
        
        desbloquear_excel(input_path, output_path)
    else:
        print("El archivo no existe. Por favor verifica la ruta.")