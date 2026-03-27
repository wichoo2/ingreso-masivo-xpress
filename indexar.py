import os
import sys
import json
import time
import zipfile
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, "cache_cole.json")

sys.path.insert(0, BASE_DIR)
from config_local import CARPETA_TIENDAS, FILA_DATOS_DEST
import logica_local as logica

_PAT_E = re.compile(
    rb'<c r="E(\d+)"[^>]*>'
    rb'(?:<is><t[^>]*>(.*?)</t></is>'
    rb'|<v>(\d+)</v>)',
    re.DOTALL)

_PAT_SS = re.compile(
    rb'<si>(?:<r>)?(?:<rPr>.*?</rPr>)?<t[^>]*>(.*?)</t>(?:</r>)?</si>',
    re.DOTALL)


def _leer_zip(args):
    clave_arch, ruta_arch = args
    nombres = {}
    try:
        with zipfile.ZipFile(ruta_arch, 'r') as z:
            nz = z.namelist()
            strings = []
            for ss in ['xl/sharedStrings.xml', 'xl/SharedStrings.xml']:
                if ss in nz:
                    raw_ss = z.read(ss)
                    strings = [
                        m.group(1).decode('utf-8', errors='ignore').strip()
                        for m in _PAT_SS.finditer(raw_ss)
                    ]
                    break
            hoja_xml = None
            for cand in nz:
                if 'worksheets/sheet' in cand and cand.endswith('.xml'):
                    hoja_xml = cand
                    break
            if not hoja_xml:
                return clave_arch, {}, False
            raw_ws = z.read(hoja_xml)

        for m in _PAT_E.finditer(raw_ws):
            try:
                num_fila = int(m.group(1))
                if num_fila < FILA_DATOS_DEST:
                    continue
            except Exception:
                continue

            inline = m.group(2)
            v_idx  = m.group(3)
            texto  = ""
            if inline:
                texto = inline.decode('utf-8', errors='ignore').strip()
            elif v_idx:
                idx_s = int(v_idx)
                if strings and idx_s < len(strings):
                    texto = strings[idx_s]

            if not texto or texto.startswith("#"):
                continue
            v = logica.normalizar(texto)
            if v and len(v) >= 3 and not v.isdigit() and v not in nombres:
                nombres[v] = clave_arch
            # También indexar versión sin sufijo de tipo
            v_st = logica.normalizar_sin_tipo(texto)
            if v_st and v_st != v and len(v_st) >= 3 and not v_st.isdigit() and v_st not in nombres:
                nombres[v_st] = clave_arch

        return clave_arch, nombres, True
    except Exception:
        return clave_arch, {}, False


def indexar(carpeta, callback_progreso=None):
    """
    Construye el caché de col E con invalidación por mtime.
    Solo re-indexa archivos que cambiaron desde la última indexación.
    """
    # Cargar caché anterior con mtimes
    cache_ant  = {}
    mtimes_ant = {}
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            datos = json.load(f)
            cache_ant  = datos.get("indice", {})
            mtimes_ant = datos.get("mtimes", {})
    except Exception:
        pass

    archivos = {}
    for nombre in os.listdir(carpeta):
        if nombre.lower().endswith('.xlsx') and not nombre.startswith('~$'):
            clave = logica.normalizar(os.path.splitext(nombre)[0])
            archivos[clave] = os.path.join(carpeta, nombre)

    # Detectar archivos que cambiaron (mtime diferente) o son nuevos
    por_procesar = {}
    mtimes_new   = {}
    for clave, ruta in archivos.items():
        try:
            mtime = os.path.getmtime(ruta)
        except Exception:
            mtime = 0
        mtimes_new[clave] = mtime
        if mtimes_ant.get(clave, 0) != mtime:
            por_procesar[clave] = ruta

    # Mantener entradas del caché que siguen válidas (archivos no modificados)
    indice = {}
    for nombre_norm, clave_arch in cache_ant.items():
        if clave_arch in archivos and clave_arch not in por_procesar:
            indice[nombre_norm] = clave_arch

    total    = len(por_procesar)
    lock     = threading.Lock()
    contador = [0]
    n_proc   = 0

    if total > 0:
        with ThreadPoolExecutor(max_workers=12) as executor:
            futuros = {executor.submit(_leer_zip, item): item
                       for item in por_procesar.items()}
            for futuro in as_completed(futuros):
                _, nombres, ok = futuro.result()
                with lock:
                    for v, clave in nombres.items():
                        if v not in indice:
                            indice[v] = clave
                    if ok:
                        n_proc += 1
                    contador[0] += 1
                    if callback_progreso:
                        callback_progreso(contador[0], total)

    # Guardar caché actualizado con mtimes
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "indice":  indice,
            "mtimes":  mtimes_new,
            "updated": time.strftime("%d/%m/%Y %H:%M:%S"),
        }, f, ensure_ascii=False)

    return indice, n_proc, len(indice)


def cargar_cache():
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get("indice", {})
    except Exception:
        return {}


if __name__ == "__main__":
    BARRA = 30
    print("=" * 55)
    print("  INDEXAR COLUMNA E  |  Xpress")
    print("=" * 55)
    print("Carpeta: {}\n".format(CARPETA_TIENDAS))

    if not os.path.isdir(CARPETA_TIENDAS):
        print("ERROR: carpeta no encontrada")
        sys.exit(1)

    def _prog(actual, total):
        pct     = actual / total if total else 1
        relleno = int(BARRA * pct)
        barra   = "#" * relleno + "-" * (BARRA - relleno)
        print("  [{}] {:>3}%  {}/{}".format(barra, int(pct*100), actual, total),
              end="\r", flush=True)

    indice, n_arch, n_nombres = indexar(CARPETA_TIENDAS, _prog)
    print()
    print("\n[OK] {} nombres en {} archivos.".format(n_nombres, n_arch))
    print("Cache: {}\n".format(CACHE_FILE))
