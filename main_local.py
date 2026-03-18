import os
import sys
import time
import glob
import copy
from openpyxl import load_workbook

import logica_local as logica
from config_local import (
    CARPETA_TIENDAS, ARCHIVO_INGRESO,
    HOJA_ORIGEN, HOJA_TIENDAS,
    FILA_INICIO, FILA_DATOS_DEST, FILA_TIENDAS_INI,
    COL_RESULTADO, COL_DATOS_INI, COL_TIENDA, COL_ID,
    COL_TIPOSERV, COL_MUNICIPIO, COL_COMENTARIO
)

SEP = "=" * 60
MAX_LIBROS = 25   # Aumentado de 10 → menos pausas de guardado en lote

import zipfile as _zipfile_main

def _preservar_partes_zip(ruta_orig, ruta_nuevo):
    """
    Preserva solo xl/theme/theme1.xml del original en el nuevo.
    Version rapida: solo lee el tema (pequeno) y lo inyecta.
    Si el tema no cambio o falla, no hace nada.
    """
    TEMA = "xl/theme/theme1.xml"
    try:
        if not os.path.isfile(ruta_orig) or not os.path.isfile(ruta_nuevo):
            return

        # Leer solo el tema del original
        with _zipfile_main.ZipFile(ruta_orig, 'r') as z:
            if TEMA not in z.namelist():
                return
            tema_orig = z.read(TEMA)

        # Verificar si el tema cambio en el nuevo
        with _zipfile_main.ZipFile(ruta_nuevo, 'r') as z:
            if TEMA not in z.namelist():
                return
            tema_nuevo = z.read(TEMA)

        # Si son identicos no hace falta nada
        if tema_orig == tema_nuevo:
            return

        # Solo si cambio: reemplazar el tema
        ruta_patch = ruta_nuevo + ".patch"
        with _zipfile_main.ZipFile(ruta_nuevo, 'r') as z_in:
            with _zipfile_main.ZipFile(ruta_patch, 'w',
                                        compression=_zipfile_main.ZIP_DEFLATED) as z_out:
                for item in z_in.infolist():
                    if item.filename == TEMA:
                        z_out.writestr(item, tema_orig)
                    else:
                        z_out.writestr(item, z_in.read(item.filename))

        if os.path.isfile(ruta_patch) and os.path.getsize(ruta_patch) > 0:
            os.remove(ruta_nuevo)
            os.rename(ruta_patch, ruta_nuevo)
    except Exception:
        try:
            if os.path.isfile(ruta_nuevo + ".patch"):
                os.remove(ruta_nuevo + ".patch")
        except Exception:
            pass



# =============================================================================
# CACHE DE LIBROS
# =============================================================================
class CacheLibros:
    def __init__(self):
        self._libros = {}
        self._orden  = []

    def tiene(self, clave):
        return clave in self._libros

    def obtener(self, clave):
        return self._libros.get(clave, {})

    def agregar(self, clave, datos):
        self._libros[clave] = datos
        self._orden.append(clave)

    def count(self):
        return len(self._libros)

    def cerrar_todos(self):
        for clave, datos in list(self._libros.items()):
            nombre    = os.path.basename(datos["path"])
            ruta_orig = datos["path"]
            ruta_tmp  = ruta_orig + ".tmp"
            try:
                print("  Guardando {}...".format(nombre))
                datos["wb"].save(ruta_tmp)
                datos["wb"].close()
                if os.path.isfile(ruta_tmp) and os.path.getsize(ruta_tmp) > 0:
                    # Preservar partes del ZIP original que openpyxl no maneja:
                    # tema de colores, colores recientes, metadatos
                    _preservar_partes_zip(ruta_orig, ruta_tmp)
                    if os.path.isfile(ruta_orig):
                        os.remove(ruta_orig)
                    os.rename(ruta_tmp, ruta_orig)
                    print("[OK] Guardado: {}".format(nombre))
                else:
                    print("[ERROR] Temporal vacio: {}".format(nombre))
                    if os.path.isfile(ruta_tmp):
                        os.remove(ruta_tmp)
            except Exception as e:
                print("[ERROR] Al guardar {}: {}".format(nombre, e))
                if os.path.isfile(ruta_tmp):
                    try: os.remove(ruta_tmp)
                    except: pass
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
        print("[ERROR] No se pudo abrir {}: {}".format(os.path.basename(ruta), e))
        return None

    ws_valida = None
    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        # Saltar hojas ocultas
        if ws.sheet_state != "visible":
            continue
        # Saltar hojas bloqueadas (con o sin contrasena)
        try:
            if ws.protection.sheet:
                continue
        except Exception:
            pass
        # Usar la primera hoja visible, sin bloqueo y con encabezados correctos
        if logica.hoja_valida(ws):
            ws_valida = ws
            break

    if ws_valida is None:
        wb.close()
        return None

    cols_esp = logica.detectar_cols_especiales(ws_valida)
    ids      = logica.cargar_ids_destino(ws_valida)

    nombres_cole = set()
    for row in ws_valida.iter_rows(min_row=FILA_DATOS_DEST, max_row=ws_valida.max_row,
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

    # Pre-calcular ultima fila UNA sola vez al abrir
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
        "fila_libre":   fila_libre,  # puntero incremental en cache
    }

# =============================================================================
# MAIN
# =============================================================================
def main():
    t_inicio = time.time()

    print("\n" + SEP)
    print("  MOVER PAQUETES - MODO LOCAL")
    print(SEP + "\n")

    if not os.path.isfile(ARCHIVO_INGRESO):
        print("[ERROR] No se encontro INGRESO_MASIVO:\n  {}".format(ARCHIVO_INGRESO))
        return
    if not os.path.isdir(CARPETA_TIENDAS):
        print("[ERROR] No se encontro carpeta tiendas:\n  {}".format(CARPETA_TIENDAS))
        return

    print("Cargando INGRESO_MASIVO...")
    try:
        wb_origen = load_workbook(ARCHIVO_INGRESO, data_only=False)
    except Exception as e:
        print("[ERROR] {}".format(e))
        return

    if HOJA_ORIGEN not in wb_origen.sheetnames:
        print("[ERROR] No se encontro hoja '{}'".format(HOJA_ORIGEN))
        return

    ws_origen = wb_origen[HOJA_ORIGEN]

    mapa_tiendas = {}
    if HOJA_TIENDAS in wb_origen.sheetnames:
        mapa_tiendas = logica.cargar_mapa_tiendas(
            wb_origen[HOJA_TIENDAS], FILA_TIENDAS_INI)
        print("[OK] Hoja TIENDAS: {} variantes".format(len(mapa_tiendas)))
    else:
        print("[AVISO] Sin hoja TIENDAS")

    print("\nIndexando carpeta de tiendas...")
    indice = indexar_carpeta(CARPETA_TIENDAS)
    print("[OK] {} archivos encontrados".format(len(indice)))

    # Cargar cache de col E
    import indexar as _idx
    indice_cole = _idx.cargar_cache()
    if indice_cole:
        print("[OK] Cache col E: {} nombres".format(len(indice_cole)))
    else:
        print("[AVISO] Sin cache col E. Usa 'Indexar tiendas' para generarlo.")
    print()

    cache       = CacheLibros()
    c_listo = c_falta = c_dup = 0
    ultima_fila = ws_origen.max_row
    filas_falta = []

    tiendas_unicas = len(set(
        logica.normalizar(row[0] or "")
        for row in ws_origen.iter_rows(min_row=FILA_INICIO, max_row=ultima_fila,
                                        min_col=COL_TIENDA, max_col=COL_TIENDA,
                                        values_only=True)
        if row[0]
    ))
    print("Procesando filas {} a {} ({} tiendas unicas)...\n".format(
        FILA_INICIO, ultima_fila, tiendas_unicas))
    # Linea especial para que la UI pueda parsear el total exacto
    print("TOTAL_FILAS:{}".format(ultima_fila - FILA_INICIO + 1), flush=True)

    # Pre-leer todas las filas del origen EN MEMORIA de una sola pasada
    # Esto evita cientos de llamadas ws.cell() dentro del loop — mucho mas rapido
    COL_MAX_LEER = max(COL_COMENTARIO, COL_DATOS_INI + 6) + 1
    filas_data = {}
    for row in ws_origen.iter_rows(min_row=FILA_INICIO, max_row=ultima_fila,
                                    min_col=1, max_col=COL_MAX_LEER,
                                    values_only=True):
        fila_num = FILA_INICIO + len(filas_data)
        filas_data[fila_num] = row

    for fila_real in range(FILA_INICIO, ultima_fila + 1):

        fila_vals = filas_data.get(fila_real, ())

        def get(col):
            idx = col - 1
            v   = fila_vals[idx] if idx < len(fila_vals) else None
            return str(v).strip() if v is not None else ""

        ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = ""

        nombre_tienda = get(COL_TIENDA)
        id_nuevo      = get(COL_ID)
        valor_k       = get(COL_TIPOSERV)
        valor_m       = get(COL_MUNICIPIO)
        valor_n       = get(COL_COMENTARIO)

        arr_datos = [fila_vals[COL_DATOS_INI - 1 + j] if (COL_DATOS_INI - 1 + j) < len(fila_vals) else None
                     for j in range(7)]

        # Fila completamente vacia
        if (not nombre_tienda and not id_nuevo and
                all(v is None or str(v).strip() == "" for v in arr_datos)):
            continue

        # Sin tienda
        if not nombre_tienda:
            ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "FALTA"
            filas_falta.append(list(fila_vals[:14]) + [None] * max(0, 14 - len(fila_vals)))
            c_falta += 1
            continue

        norm_tienda   = logica.normalizar(nombre_tienda)
        clave_archivo = None

        # [1] Exacto
        if norm_tienda in indice:
            clave_archivo = norm_tienda
        # [2] Hoja TIENDAS
        elif norm_tienda in mapa_tiendas:
            nd = mapa_tiendas[norm_tienda]
            if nd in indice:
                clave_archivo = nd
        # [3] Cache col E
        elif norm_tienda in indice_cole:
            clave_archivo = indice_cole[norm_tienda]

        if not clave_archivo:
            ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "FALTA"
            filas_falta.append([ws_origen.cell(row=fila_real, column=c).value
                                 for c in range(1, 15)])
            c_falta += 1
            print("  FALTA  fila {}: '{}'".format(fila_real, nombre_tienda))
            continue

        # Limite cache
        if cache.count() >= MAX_LIBROS and not cache.tiene(clave_archivo):
            print("\n  [LOTE] Guardando {} libros...".format(MAX_LIBROS))
            cache.cerrar_todos()

        # Abrir libro
        if not cache.tiene(clave_archivo):
            ruta  = indice[clave_archivo]
            nombre_arch = os.path.basename(ruta)
            print("  Abriendo {}...".format(nombre_arch))
            datos = cargar_libro(ruta, clave_archivo)
            if datos is None:
                ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "FALTA"
                filas_falta.append(list(fila_vals[:14]) + [None] * max(0, 14 - len(fila_vals)))
                c_falta += 1
                print("  FALTA  fila {}: hoja invalida en '{}'".format(
                    fila_real, nombre_tienda))
                continue
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

        num_paquetes = max(logica.extraer_num_paquetes(valor_n), 1)
        estado, a_insertar = logica.evaluar_duplicado(
            id_nuevo, valor_n, ids_dest, num_paquetes)

        if estado == "DUP":
            ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "DUP"
            c_dup += 1
            print("  DUP    fila {}: '{}' ID {}".format(
                fila_real, nombre_tienda, id_nuevo))
            continue

        ya_existentes = ids_dest.get(id_nuevo, 0)

        # Usar fila_libre del cache — ya calculada al abrir, solo avanza
        fila_libre = datos_lib["fila_libre"]

        for k in range(a_insertar):
            logica.insertar_paquete(
                ws_dest, fila_libre, arr_datos,
                precio_cero=(ya_existentes > 0 or k > 0),
                valor_k=valor_k, valor_m=valor_m, valor_n=valor_n,
                es_primero=(k == 0),
                tipo_q=tipo_q, col_q=col_q,
                col_paq=col_paq, col_oid=col_oid)
            fila_libre += 1
            # Version rapida: solo busca en 20 filas desde el punto actual
            fila_libre = logica.primera_fila_libre_rapida(ws_dest, fila_libre)

        # Guardar puntero actualizado en cache
        datos_lib["fila_libre"] = fila_libre

        if id_nuevo:
            ids_dest[id_nuevo] = ya_existentes + a_insertar

        ws_origen.cell(row=fila_real, column=COL_RESULTADO).value = "LISTO"
        c_listo += 1
        print("  LISTO  fila {}: {}".format(fila_real, nombre_tienda))

    # Guardar tiendas
    print("\nGuardando archivos de tiendas...")
    cache.cerrar_todos()

    # Hoja FALTA
    if filas_falta:
        print("\nActualizando hoja FALTA...")
        HOJA_FALTA = "FALTA"
        if HOJA_FALTA in wb_origen.sheetnames:
            ws_falta = wb_origen[HOJA_FALTA]
            for row in ws_falta.iter_rows(min_row=4, max_row=ws_falta.max_row):
                for cell in row:
                    cell.value = None
        else:
            ws_falta = wb_origen.create_sheet(HOJA_FALTA)

        # Construir indice id→fila UNA sola vez (evita O(n²) anterior)
        indice_id_fila = {}
        for fr in range(FILA_INICIO, ws_origen.max_row + 1):
            id_val = ws_origen.cell(row=fr, column=6).value
            res    = str(ws_origen.cell(row=fr, column=COL_RESULTADO).value or "").strip().upper()
            if id_val and res == "FALTA":
                clave_id = str(id_val)
                if clave_id not in indice_id_fila:
                    indice_id_fila[clave_id] = fr

        ri = 4
        for fila_datos in filas_falta:
            id_buscar        = fila_datos[5]
            fila_origen_real = indice_id_fila.get(str(id_buscar)) if id_buscar else None

            for ci, valor in enumerate(fila_datos, start=1):
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
        print("[OK] {} filas en hoja FALTA".format(len(filas_falta)))

    # Guardar INGRESO_MASIVO
    print("\nGuardando INGRESO_MASIVO...")
    ruta_tmp = ARCHIVO_INGRESO + ".tmp"
    try:
        wb_origen.save(ruta_tmp)
        wb_origen.close()
        if os.path.isfile(ruta_tmp) and os.path.getsize(ruta_tmp) > 0:
            if os.path.isfile(ARCHIVO_INGRESO):
                os.remove(ARCHIVO_INGRESO)
            os.rename(ruta_tmp, ARCHIVO_INGRESO)
            print("[OK] INGRESO_MASIVO guardado")
        else:
            print("[ERROR] No se pudo guardar INGRESO_MASIVO")
    except Exception as e:
        print("[ERROR] {}".format(e))
        if os.path.isfile(ruta_tmp):
            try: os.remove(ruta_tmp)
            except: pass

    t_total = time.time() - t_inicio
    print("\n" + SEP)
    print("  PROCESO COMPLETADO")
    print(SEP)
    print("  LISTO : {}".format(c_listo))
    print("  FALTA : {}".format(c_falta))
    print("  DUP   : {}".format(c_dup))
    print("  Tiempo: {:.1f}s".format(t_total))
    print(SEP + "\n")


if __name__ == "__main__":
    main()
