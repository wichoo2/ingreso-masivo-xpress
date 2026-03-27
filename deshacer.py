"""
deshacer.py — Sistema de backup y restauración para INGRESO MASIVO.

Flujo:
  1. Antes de tocar cualquier archivo de tienda → guardar copia en carpeta _backups/
  2. Al terminar el proceso → escribir registro JSON con qué archivos se modificaron
  3. Botón Deshacer → leer el registro y restaurar cada backup

El registro se guarda junto al script: _backups/ultimo_proceso.json
Los backups se guardan en:          _backups/<timestamp>/<nombre_archivo>.xlsx
"""

import os
import shutil
import json
import time

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
CARPETA_BAK  = os.path.join(BASE_DIR, "_backups")
REGISTRO_BAK = os.path.join(CARPETA_BAK, "ultimo_proceso.json")

# Solo guardamos el último proceso — sobrescribimos cada vez.
# Si se necesita historial, se puede cambiar a timestamps individuales.

def iniciar_sesion(timestamp: str = None) -> str:
    """
    Crea la carpeta del backup para esta sesión y devuelve su ruta.
    timestamp: cadena única para esta sesión (ej: '20260325_143022')
    """
    if timestamp is None:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
    carpeta_sesion = os.path.join(CARPETA_BAK, timestamp)
    os.makedirs(carpeta_sesion, exist_ok=True)
    return carpeta_sesion, timestamp


def guardar_backup(ruta_original: str, carpeta_sesion: str) -> str:
    """
    Copia el archivo original a la carpeta de backup de la sesión.
    Solo lo copia si no existe ya (evita sobreescribir con versión modificada).
    Devuelve la ruta del backup.
    """
    nombre   = os.path.basename(ruta_original)
    ruta_bak = os.path.join(carpeta_sesion, nombre)
    if not os.path.isfile(ruta_bak):
        shutil.copy2(ruta_original, ruta_bak)
    return ruta_bak


def guardar_registro(timestamp: str, carpeta_sesion: str,
                     archivos_modificados: list,
                     c_listo: int, c_dup: int, c_falta: int):
    """
    Escribe el registro JSON del proceso actual.
    archivos_modificados: lista de rutas originales que se tocaron.
    """
    os.makedirs(CARPETA_BAK, exist_ok=True)
    registro = {
        "timestamp":   timestamp,
        "fecha":       time.strftime("%d/%m/%Y %H:%M:%S"),
        "carpeta_bak": carpeta_sesion,
        "listo":       c_listo,
        "dup":         c_dup,
        "falta":       c_falta,
        "archivos":    [
            {
                "original": ruta,
                "backup":   os.path.join(carpeta_sesion, os.path.basename(ruta)),
            }
            for ruta in archivos_modificados
        ],
    }
    with open(REGISTRO_BAK, "w", encoding="utf-8") as f:
        json.dump(registro, f, ensure_ascii=False, indent=2)


def cargar_registro() -> dict:
    """
    Carga el registro del último proceso. Retorna {} si no existe.
    """
    if not os.path.isfile(REGISTRO_BAK):
        return {}
    try:
        with open(REGISTRO_BAK, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def puede_deshacer() -> tuple:
    """
    Verifica si hay un proceso que se pueda deshacer.
    Retorna (puede: bool, info: str)
    """
    reg = cargar_registro()
    if not reg:
        return False, "No hay ningún proceso para deshacer."
    n = len(reg.get("archivos", []))
    fecha = reg.get("fecha", "?")
    listo = reg.get("listo", 0)
    if n == 0:
        return False, "El último proceso no modificó ningún archivo."
    return True, (
        "Último proceso: {} — {} archivo(s) modificado(s), "
        "{} paquete(s) insertado(s).".format(fecha, n, listo)
    )


def deshacer(callback_log=None) -> tuple:
    """
    Restaura todos los archivos del último proceso desde los backups.
    callback_log: función(msg) para reportar progreso.
    Retorna (exito: bool, mensaje: str)
    """
    def _log(msg):
        if callback_log:
            callback_log(msg)

    reg = cargar_registro()
    if not reg:
        return False, "No hay ningún proceso registrado para deshacer."

    archivos = reg.get("archivos", [])
    if not archivos:
        return False, "El último proceso no modificó ningún archivo."

    errores  = []
    restored = 0

    _log("Deshaciendo proceso del {}...".format(reg.get("fecha", "?")))
    _log("Restaurando {} archivo(s)...".format(len(archivos)))

    for entry in archivos:
        ruta_orig = entry.get("original", "")
        ruta_bak  = entry.get("backup", "")
        nombre    = os.path.basename(ruta_orig)

        if not os.path.isfile(ruta_bak):
            msg = "  [ERROR] Backup no encontrado: {}".format(ruta_bak)
            _log(msg)
            errores.append(nombre)
            continue

        # Verificar que el original no esté abierto en Excel
        directorio = os.path.dirname(ruta_orig)
        lock_path  = os.path.join(directorio, "~$" + nombre)
        if os.path.isfile(lock_path):
            msg = "  [ERROR] '{}' está abierto en Excel — ciérralo y vuelve a intentar.".format(nombre)
            _log(msg)
            errores.append(nombre)
            continue

        try:
            shutil.copy2(ruta_bak, ruta_orig)
            _log("  [OK] Restaurado: {}".format(nombre))
            restored += 1
        except Exception as e:
            msg = "  [ERROR] No se pudo restaurar '{}': {}".format(nombre, e)
            _log(msg)
            errores.append(nombre)

    # Invalidar el registro para no poder deshacer dos veces el mismo proceso
    if restored > 0 and not errores:
        try:
            # Renombrar el registro en vez de borrarlo — mantiene trazabilidad
            ruta_done = REGISTRO_BAK.replace(".json", "_deshecho.json")
            if os.path.isfile(ruta_done):
                os.remove(ruta_done)
            os.rename(REGISTRO_BAK, ruta_done)
        except Exception:
            pass

    if errores:
        return False, (
            "Restaurados {}/{} archivos. Errores en: {}".format(
                restored, len(archivos), ", ".join(errores))
        )
    return True, "Proceso deshecho correctamente. {} archivo(s) restaurado(s).".format(restored)


def limpiar_backups_viejos(dias: int = 7):
    """
    Elimina carpetas de backup con más de N días de antigüedad.
    Llamar periódicamente para no acumular espacio.
    """
    if not os.path.isdir(CARPETA_BAK):
        return
    ahora = time.time()
    for entrada in os.listdir(CARPETA_BAK):
        ruta = os.path.join(CARPETA_BAK, entrada)
        if not os.path.isdir(ruta):
            continue
        try:
            edad = ahora - os.path.getmtime(ruta)
            if edad > dias * 86400:
                shutil.rmtree(ruta)
        except Exception:
            pass
