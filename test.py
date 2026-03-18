import openpyxl
import glob
import os

carpeta = r"C:\Users\josue\Desktop\ALL_TIENDAS"

print("Verificando archivos en:", carpeta)
print("-" * 50)

archivos = glob.glob(carpeta + "\\*.xls*")

if not archivos:
    print("No se encontraron archivos .xlsx en esa carpeta")
    print("Verifica que la ruta sea correcta")
else:
    for ruta in archivos:
        nombre = os.path.basename(ruta)
        if nombre.startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(ruta, read_only=True)
            wb.close()
            print(f"OK  — {nombre}")
        except Exception as e:
            print(f"MAL — {nombre}")
            print(f"      Error: {e}")

print("-" * 50)
print("Listo.")

