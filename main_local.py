import os
import sys
import time
import glob
import copy
import traceback
from collections import defaultdict
from openpyxl import load_workbook

import logica_local as logica
from config_local import (
    CARPETA_TIENDAS, ARCHIVO_INGRESO,
    HOJA_ORIGEN, HOJA_TIENDAS,
    FILA_INICIO, FILA_DATOS_DEST, FILA_TIENDAS_INI,
    COL_RESULTADO, COL_DATOS_INI, COL_TIENDA, COL_ID,
    COL_TIPOSERV, COL_MUNICIPIO, COL_COMENTARIO
)

SEP            = "=" * 60
MAX_LIBROS     = 25
MAX_REINTENTOS = 5
ESPERA_REINT   = 3

import zipfile as _zipfile_main

# =============================================================================
# LOG PERSISTENTE
# =============================================================================
_LOG_PATH = None

def _init_log():
    global _LOG_PATH
    try:
        base = os.path.dirname(os.path.abspath(__file__))
        _LOG_PATH = os.path.join(base, "ultimo_proceso.log")
        with open(_LOG_PATH, "w", encoding="utf-8") as f:
            f.write("=== INGRESO MASIVO - LOG {} ===\n".format(
                time.strftime("%d/%m/%Y %H:%M:%S")))
    except Exception:
        _LOG_PATH = None

def _log(msg):
    print(msg)
    if _LOG_PATH:
        try:
            with open(_LOG_PATH, "a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except Exception:
            pass

# =============================================================================
# ONEDRIVE SYNC
# =============================================================================
def _forzar_sync_onedrive(ruta):
    try:
        _now = time.time()
        os.utime(ruta, (_now, _now))
    except Exception:
        pass
    try:
        with open(ruta, "ab") as f:
            pass
    except Exception:
        pass

# =============================================================================
# DETECTAR ARCHIVOS BLOQUEADOS
# =============================================================================
def _archivo_bloqueado(ruta):
    directorio = os.path.dirname(ruta)
    nombre     = os.path.basename(ruta)
    lock       = os.path.join(directorio, "~$" + nombre)
    return os.path.isfile(lock)

def _verificar_bloqueados(indice):
    return [os.path.basename(ruta)
            for ruta in indice.values()
            if _archivo_bloqueado(ruta)]

# =============================================================================
# PRESERVAR ZIP (solo tema) — trabaja con BytesIO, sin archivos temporales
# =============================================================================
def _preservar_tema_en_buf(ruta_orig, buf):
    """
    Preserva xl/theme/theme1.xml del archivo original dentro del buffer BytesIO.
    Modifica el buffer en memoria sin tocar el disco.
    """
    import io as _io
    TEMA = "xl/theme/theme1.xml"
    try:
        if not os.path.isfile(ruta_orig):
            return
        with _zipfile_main.ZipFile(ruta_orig, 'r') as z:
            if TEMA not in z.namelist():
                return
            tema_orig = z.read(TEMA)

        buf.seek(0)
        with _zipfile_main.ZipFile(buf, 'r') as z:
            if TEMA not in z.namelist():
                return
            if z.read(TEMA) == tema_orig:
                return  # identicos

        # Reconstruir con tema original
        buf.seek(0)
        datos_zip = {}
        infos_zip = {}
        with _zipfile_main.ZipFile(buf, 'r') as z_in:
            for item in z_in.infolist():
                datos_zip[item.filename] = z_in.read(item.filename)
                infos_zip[item.filename] = item
        datos_zip[TEMA] = tema_orig

        buf_nuevo = _io.BytesIO()
        with _zipfile_main.ZipFile(buf_nuevo, 'w',
                                    compression=_zipfile_main.ZIP_DEFLATED) as z_out:
            for fn, data in datos_zip.items():
                z_out.writestr(infos_zip[fn], data)

        # Reemplazar contenido del buffer original
        contenido = buf_nuevo.getvalue()
        buf.seek(0)
        buf.truncate()
        buf.write(contenido)
        buf.seek(0)
    except Exception:
        pass

# =============================================================================
# CACHE DE LIBROS
# =============================================================================
class CacheLibros:
    def __init__(self):
        self._libros  = {}
        self._orden   = []
        self._resumen = {}

    def tiene(self, clave):      return clave in self._libros
    def obtener(self, clave):    return self._libros.get(clave, {})\


    def count(self):             return len(self._libros)
    def get_resumen(self):       return self._resumen

    def agregar(self, clave, datos):
        self._libros[clave] = datos
        self._orden.append(clave)
        if clave not in self._resumen:
            self._resumen[clave] = {
                "nombre": os.path.basename(datos["path"]),
                "listo": 0, "dup": 0}

    def sumar(self, clave, tipo):
        if clave in self._resumen:
            self._resumen[clave][tipo] = self._resumen[clave].get(tipo, 0) + 1

    def cerrar_todos(self):
        import io as _io
        for clave, datos in list(self._libros.items()):
            nombre    = os.path.basename(datos["path"])
            ruta_orig = datos["path"]

            guardado = False
            for intento in range(1, MAX_REINTENTOS + 1):
                try:
                    # Guardar en buffer de memoria — sin crear ningun archivo
                    buf = _io.BytesIO()
                    datos["wb"].save(buf)
                    datos["wb"].close()
                    contenido = buf.getvalue()

                    if contenido:
                        # Preservar tema de colores del original
                        _preservar_tema_en_buf(ruta_orig, buf)
                        contenido = buf.getvalue()

                        # Escribir directamente sobre el archivo original
                        with open(ruta_orig, 'wb') as f:
                            f.write(contenido)
                        _forzar_sync_onedrive(ruta_orig)
                        _log("[OK] Guardado: {}".format(nombre))
                        guardado = True
                    else:
                        _log("[ERROR] Buffer vacio: {}".format(nombre))
                    break
                except PermissionError:
                    if intento < MAX_REINTENTOS:
                        _log("  [Reintento {}/{}] '{}' bloqueado, esperando {}s...".format(
                            intento, MAX_REINTENTOS, nombre, ESPERA_REINT))
                        time.sleep(ESPERA_REINT)
                    else:
                        _log("[ERROR] No se pudo guardar '{}' despues de {} intentos".format(
                            nombre, MAX_REINTENTOS))
                except Exception as e:
                    _log("[ERROR] Al guardar {}: {}".format(nombre, e))
                    break

            if not guardado:
                _log("[ERROR CRITICO] '{}' NO fue guardado.".format(nombre))

        self._libros.clear()
        self._orden.clear()

# =============================================================================
# INDEXAR CARPETA
# =============================================================================
def indexar_carpeta(carpeta):
    indice = {}
    for ruta in glob.glob(os.path.join(carpeta, "*.xls*")):
        nombre = os.path.basename(ruta)
        if nombre.startswith("~$"):
            continue
        dot     = nombre.rfind(".")
        sin_ext = nombre[:dot] if dot > 0 else nombre
        norm    = logica.normalizar(sin_ext)
        if norm and norm not in indice:
            indice[norm] = ruta
    return indice

# =============================================================================
# CARGAR LIBRO DESTINO
# =============================================================================
def cargar_libro(ruta, clave):
    try:
        wb = load_workbook(ruta, keep_vba=False, data_only=False)
    except Exception as e:
        _log("[ERROR] No se pudo abrir {}: {}".format(os.path.basename(ruta), e))
        return None

    ws_valida = None
    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        if ws.sheet_state != "visible":
            continue
        try:
            if ws.protection.sheet:
                continue
        except Exception:
            pass
        if logica.hoja_valida(ws):
            ws_valida = ws
            break

    if ws_valida is None:
        wb.close()
        return None

    cols_esp       = logica.detectar_cols_especiales(ws_valida)
    ids            = logica.cargar_ids_destino(ws_valida)
    fila_enc_real  = logica.get_fila_encabezado(ws_valida)
    fila_datos_ini = fila_enc_real + 1

    nombres_cole = set()
    for row in ws_valida.iter_rows(min_row=fila_datos_ini, max_row=ws_valida.max_row,
                                    min_col=5, max_col=5, values_only=True):
        raw = row[0]
        if raw is None:
            continue
        v = str(raw).strip()
        if v.startswith("#") or v == "":
            continue
        n = logica.normalizar(v)
        if n and len(n) >= 3 and not n.isdigit():
            nombres_cole.add(n)

    ult_datos  = logica.ultima_fila_con_datos(ws_valida)
    fila_libre = logica.primera_fila_libre(ws_valida, ult_datos + 1)

    return {
        "wb":           wb,
        "ws":           ws_valida,
        "path":         ruta,
        "tipo_q":       cols_esp["tipo_q"],
        "col_q":        cols_esp["col_q"],
        "col_paq":      cols_esp["col_paq"],
        "col_oid":      cols_esp["col_oid"],
        "ids":          ids,
        "nombres_cole": nombres_cole,
        "clave":        clave,
        "fila_libre":   fila_libre,
        "fila_enc":     fila_enc_real,
    }

# =============================================================================
# PRE-CLASIFICAR FILAS POR TIENDA (evita busquedas repetidas)
# =============================================================================
def _clasificar_filas(filas_data, indice, mapa_tiendas, indice_cole,
                      ws_origen, ultima_fila):
    """
    Paso 1: clasifica todas las filas en memoria agrupandolas por tienda.
    Evita llamar normalizar() + busqueda de indice N veces para la misma tienda.
    Retorna: dict clave_archivo -> [lista de (fila_real, fila_vals)]
             lista de filas_falta
    """
    # Cache local de normalizacion — cada nombre de tienda se normaliza UNA sola vez
    _norm_cache = {}

    grupos  = defaultdict(list)   # clave_archivo -> [(fila_real, fila_vals)]
    faltas  = []                  # [(fila_real, fila_vals, motivo)]

    for fila_real in range(FILA_INICIO, ultima_fila + 1):
        fila_vals = filas_data.get(fila_real, ())

        # Leer valores directo del tuple — sin funcion get() redefinida en loop
        def _v(col):
            idx = col - 1
            v = fila_vals[idx] if idx < len(fila_vals) else None
            return str(v).strip() if v is not None else ""

        nombre_tienda = _v(COL_TIENDA)
        id_nuevo      = _v(COL_ID)
        arr_datos     = [fila_vals[COL_DATOS_INI - 1 + j]
                         if (COL_DATOS_INI - 1 + j) < len(fila_vals) else None
                         for j in range(7)]

        # Fila completamente vacia
        if (not nombre_tienda and not id_nuevo and
                all(v is None or str(v).strip() == "" for v in arr_datos)):
            continue

        # Sin tienda
        if not nombre_tienda:
            ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "FALTA"
            faltas.append((fila_real, list(fila_vals), "sin_tienda"))
            continue

        # Normalizar con cache — evita regex por cada fila repetida
        if nombre_tienda not in _norm_cache:
            _norm_cache[nombre_tienda] = logica.normalizar(nombre_tienda)
        norm_tienda = _norm_cache[nombre_tienda]

        # Buscar archivo
        clave_archivo = None
        if norm_tienda in indice:
            clave_archivo = norm_tienda
        elif norm_tienda in mapa_tiendas:
            nd = mapa_tiendas[norm_tienda]
            if nd in indice:
                clave_archivo = nd
        elif norm_tienda in indice_cole:
            clave_archivo = indice_cole[norm_tienda]

        if not clave_archivo:
            ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "FALTA"
            faltas.append((fila_real, list(fila_vals), nombre_tienda))
            continue

        grupos[clave_archivo].append((fila_real, fila_vals))

    return grupos, faltas

# =============================================================================
# PROCESAR UN GRUPO DE FILAS PARA UNA TIENDA
# =============================================================================
def _procesar_grupo(clave_archivo, filas_grupo, cache, indice,
                    indice_cole, ws_origen):
    """
    Procesa todas las filas de UNA tienda de una sola vez.
    El libro se abre una sola vez y se insertan todos sus paquetes juntos.
    """
    c_listo = c_dup = 0

    # Abrir libro si no esta en cache
    if not cache.tiene(clave_archivo):
        ruta        = indice[clave_archivo]
        nombre_arch = os.path.basename(ruta)
        _log("  Abriendo {}...".format(nombre_arch))
        datos = cargar_libro(ruta, clave_archivo)
        if datos is None:
            # Marcar todas las filas del grupo como FALTA
            for fila_real, fila_vals in filas_grupo:
                ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "FALTA"
            _log("  FALTA  hoja invalida: '{}'".format(nombre_arch))
            return 0, len(filas_grupo), [(fr, list(fv), nombre_arch)
                                          for fr, fv in filas_grupo]
        cache.agregar(clave_archivo, datos)
        for ne in datos["nombres_cole"]:
            if ne not in indice_cole:
                indice_cole[ne] = clave_archivo

    datos_lib = cache.obtener(clave_archivo)
    ws_dest   = datos_lib["ws"]
    tipo_q    = datos_lib["tipo_q"]
    col_q     = datos_lib["col_q"]
    col_paq   = datos_lib["col_paq"]
    col_oid   = datos_lib["col_oid"]
    ids_dest  = datos_lib["ids"]
    fila_libre = datos_lib["fila_libre"]

    filas_falta_grupo = []

    for fila_real, fila_vals in filas_grupo:
        def _v(col):
            idx = col - 1
            v = fila_vals[idx] if idx < len(fila_vals) else None
            return str(v).strip() if v is not None else ""

        id_nuevo = _v(COL_ID)
        valor_k  = _v(COL_TIPOSERV)
        valor_m  = _v(COL_MUNICIPIO)
        valor_n  = _v(COL_COMENTARIO)
        arr_datos = [fila_vals[COL_DATOS_INI - 1 + j]
                     if (COL_DATOS_INI - 1 + j) < len(fila_vals) else None
                     for j in range(7)]

        num_paquetes = max(logica.extraer_num_paquetes(valor_n), 1)
        estado, a_insertar = logica.evaluar_duplicado(
            id_nuevo, valor_n, ids_dest, num_paquetes)

        if estado == "DUP":
            ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "DUP"
            cache.sumar(clave_archivo, "dup")
            c_dup += 1
            _log("  DUP    fila {}: ID {}".format(fila_real, id_nuevo))
            continue

        ya_existentes = ids_dest.get(id_nuevo, 0)

        for k in range(a_insertar):
            logica.insertar_paquete(
                ws_dest, fila_libre, arr_datos,
                precio_cero=(ya_existentes > 0 or k > 0),
                valor_k=valor_k, valor_m=valor_m, valor_n=valor_n,
                es_primero=(k == 0),
                tipo_q=tipo_q, col_q=col_q,
                col_paq=col_paq, col_oid=col_oid)
            fila_libre += 1
            fila_libre = logica.primera_fila_libre_rapida(ws_dest, fila_libre)

        if id_nuevo:
            ids_dest[id_nuevo] = ya_existentes + a_insertar

        ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "LISTO"
        cache.sumar(clave_archivo, "listo")
        c_listo += 1
        _log("  LISTO  fila {}".format(fila_real))

    # Actualizar puntero fila libre en cache
    datos_lib["fila_libre"] = fila_libre

    return c_listo, c_dup, filas_falta_grupo

# =============================================================================
# MAIN
# =============================================================================
def main():
    _init_log()
    t_inicio = time.time()

    _log("\n" + SEP)
    _log("  MOVER PAQUETES - MODO LOCAL")
    _log(SEP + "\n")

    if not os.path.isfile(ARCHIVO_INGRESO):
        _log("[ERROR] No se encontro INGRESO_MASIVO:\n  {}".format(ARCHIVO_INGRESO))
        return
    if not os.path.isdir(CARPETA_TIENDAS):
        _log("[ERROR] No se encontro carpeta tiendas:\n  {}".format(CARPETA_TIENDAS))
        return

    _log("Cargando INGRESO_MASIVO...")
    try:
        wb_origen = load_workbook(ARCHIVO_INGRESO, data_only=False)
    except Exception as e:
        _log("[ERROR] {}".format(e))
        return

    if HOJA_ORIGEN not in wb_origen.sheetnames:
        _log("[ERROR] No se encontro hoja '{}'".format(HOJA_ORIGEN))
        return

    ws_origen = wb_origen[HOJA_ORIGEN]

    mapa_tiendas = {}
    if HOJA_TIENDAS in wb_origen.sheetnames:
        mapa_tiendas = logica.cargar_mapa_tiendas(
            wb_origen[HOJA_TIENDAS], FILA_TIENDAS_INI)
        _log("[OK] Hoja TIENDAS: {} variantes".format(len(mapa_tiendas)))
    else:
        _log("[AVISO] Sin hoja TIENDAS")

    _log("\nIndexando carpeta de tiendas...")
    indice = indexar_carpeta(CARPETA_TIENDAS)
    _log("[OK] {} archivos encontrados".format(len(indice)))

    # Advertir bloqueados
    bloqueados = _verificar_bloqueados(indice)
    if bloqueados:
        _log("\n[AVISO] Archivos abiertos en Excel (pueden fallar al guardar):")
        for b in bloqueados:
            _log("  - {}".format(b))
        _log("")

    import indexar as _idx
    indice_cole = _idx.cargar_cache()
    if indice_cole:
        _log("[OK] Cache col E: {} nombres".format(len(indice_cole)))
    else:
        _log("[AVISO] Sin cache col E.")
    _log("")

    ultima_fila = ws_origen.max_row

    # Emitir total para la UI
    print("TOTAL_FILAS:{}".format(ultima_fila - FILA_INICIO + 1), flush=True)

    # ------------------------------------------------------------------
    # PASO 1: Pre-leer TODAS las filas en memoria (una sola pasada)
    # ------------------------------------------------------------------
    t0 = time.time()
    COL_MAX_LEER = max(COL_COMENTARIO, COL_DATOS_INI + 6) + 1
    filas_data = {}
    for row in ws_origen.iter_rows(min_row=FILA_INICIO, max_row=ultima_fila,
                                    min_col=1, max_col=COL_MAX_LEER,
                                    values_only=True):
        fila_num = FILA_INICIO + len(filas_data)
        filas_data[fila_num] = row
    _log("Leidas {} filas en {:.2f}s".format(len(filas_data), time.time() - t0))

    # ------------------------------------------------------------------
    # PASO 2: Clasificar filas por tienda en memoria (sin abrir libros)
    # ------------------------------------------------------------------
    t0 = time.time()
    grupos, faltas_clasificacion = _clasificar_filas(
        filas_data, indice, mapa_tiendas, indice_cole, ws_origen, ultima_fila)

    n_tiendas = len(grupos)
    n_filas   = sum(len(v) for v in grupos.values())
    _log("Clasificadas: {} tiendas, {} filas validas, {} faltas en {:.2f}s".format(
        n_tiendas, n_filas, len(faltas_clasificacion), time.time() - t0))
    _log("")

    # ------------------------------------------------------------------
    # PASO 3: Procesar tienda por tienda (ya agrupadas)
    # ------------------------------------------------------------------
    cache   = CacheLibros()
    c_listo = c_dup = 0
    filas_falta = [(fr, fv) for fr, fv, _ in faltas_clasificacion]

    # Logging de faltas de clasificacion
    for fila_real, fila_vals, motivo in faltas_clasificacion:
        if motivo != "sin_tienda":
            _log("  FALTA  fila {}: '{}'".format(fila_real, motivo))

    tiendas_procesadas = 0
    for clave_archivo, filas_grupo in grupos.items():

        # Control de cache — guardar lote si se llena
        if cache.count() >= MAX_LIBROS and not cache.tiene(clave_archivo):
            _log("\n  [LOTE] Guardando {} libros...".format(MAX_LIBROS))
            cache.cerrar_todos()

        listo, dup, faltas_grupo = _procesar_grupo(
            clave_archivo, filas_grupo, cache, indice, indice_cole, ws_origen)

        c_listo += listo
        c_dup   += dup
        filas_falta.extend([(fr, fv) for fr, fv, _ in faltas_grupo])
        tiendas_procesadas += 1

        if tiendas_procesadas % 10 == 0:
            _log("  [{}/{}] tiendas procesadas...".format(
                tiendas_procesadas, n_tiendas))

    # ------------------------------------------------------------------
    # PASO 4: Guardar todos los libros
    # ------------------------------------------------------------------
    _log("\nGuardando archivos de tiendas...")
    resumen_tiendas = cache.get_resumen()
    cache.cerrar_todos()

    c_falta = len(filas_falta)

    # ------------------------------------------------------------------
    # PASO 5: Actualizar hoja FALTA
    # ------------------------------------------------------------------
    if filas_falta:
        _log("\nActualizando hoja FALTA...")
        HOJA_FALTA = "FALTA"
        if HOJA_FALTA in wb_origen.sheetnames:
            ws_falta = wb_origen[HOJA_FALTA]
            for row in ws_falta.iter_rows(min_row=4, max_row=ws_falta.max_row):
                for cell in row:
                    cell.value = None
        else:
            ws_falta = wb_origen.create_sheet(HOJA_FALTA)

        # Indice id→fila para copiar estilos
        indice_id_fila = {}
        for fr in range(FILA_INICIO, ws_origen.max_row + 1):
            id_val = ws_origen.cell(row=fr, column=6).value
            res    = str(ws_origen.cell(row=fr, column=COL_RESULTADO).value or "").strip().upper()
            if id_val and res == "FALTA":
                clave_id = str(id_val)
                if clave_id not in indice_id_fila:
                    indice_id_fila[clave_id] = fr

        ri = 4
        for fila_real, fila_vals in filas_falta:
            id_buscar        = fila_vals[5] if len(fila_vals) > 5 else None
            fila_origen_real = indice_id_fila.get(str(id_buscar)) if id_buscar else None
            for ci, valor in enumerate(fila_vals, start=1):
                celda_dest = ws_falta.cell(row=ri, column=ci)
                celda_dest.value = valor
                if fila_origen_real:
                    try:
                        celda_orig = ws_origen.cell(row=fila_origen_real, column=ci)
                        if celda_orig.has_style:
                            celda_dest.font      = copy.copy(celda_orig.font)
                            celda_dest.fill      = copy.copy(celda_orig.fill)
                            celda_dest.border    = copy.copy(celda_orig.border)
                            celda_dest.alignment = copy.copy(celda_orig.alignment)
                            celda_dest.number_format = celda_orig.number_format
                    except Exception:
                        pass
            ri += 1
        _log("[OK] {} filas en hoja FALTA".format(len(filas_falta)))

    # ------------------------------------------------------------------
    # PASO 6: Guardar INGRESO_MASIVO
    # ------------------------------------------------------------------
    _log("\nGuardando INGRESO_MASIVO...")
    try:
        import io as _io
        buf = _io.BytesIO()
        wb_origen.save(buf)
        wb_origen.close()
        contenido = buf.getvalue()
        if contenido:
            with open(ARCHIVO_INGRESO, 'wb') as f:
                f.write(contenido)
            _forzar_sync_onedrive(ARCHIVO_INGRESO)
            _log("[OK] INGRESO_MASIVO guardado")
        else:
            _log("[ERROR] No se pudo guardar INGRESO_MASIVO")
    except Exception as e:
        _log("[ERROR] {}".format(e))

    t_total = time.time() - t_inicio

    # Resumen final
    _log("\n" + SEP)
    _log("  PROCESO COMPLETADO")
    _log(SEP)
    _log("  LISTO : {}".format(c_listo))
    _log("  FALTA : {}".format(c_falta))
    _log("  DUP   : {}".format(c_dup))
    _log("  Tiempo: {:.1f}s".format(t_total))

    if resumen_tiendas:
        _log("\n  RESUMEN POR TIENDA:")
        _log("  {:<40} {:>6} {:>5}".format("TIENDA", "LISTO", "DUP"))
        _log("  " + "-" * 54)
        for clave, dat in sorted(resumen_tiendas.items(),
                                  key=lambda x: x[1].get("listo", 0), reverse=True):
            if dat.get("listo", 0) > 0 or dat.get("dup", 0) > 0:
                _log("  {:<40} {:>6} {:>5}".format(
                    dat["nombre"][:40], dat.get("listo", 0), dat.get("dup", 0)))

    _log(SEP + "\n")
    if _LOG_PATH:
        _log("Log guardado en: {}".format(_LOG_PATH))


if __name__ == "__main__":
    main()
