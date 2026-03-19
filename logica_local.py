import re
import unicodedata
import copy
import servicios_variantes as _svc
from datetime import datetime, date
from config_local import (
    FILA_DATOS_DEST, FILA_ENCABEZADO, ENCABEZADOS_VALIDOS,
    SERVICIOS_VALIDOS, SERVICIO_DEFAULT, COLORES_AMARILLO,
    COL_DATOS_INI, COL_ID, COL_DEST_TIPOSERV,
    COL_DEST_PRECIO, COL_DEST_COMENTAR
)

# =============================================================================
# NORMALIZAR
# =============================================================================
def normalizar(txt) -> str:
    if txt is None:
        return ""
    txt = str(txt).lower().strip()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"[^a-z0-9]", "", txt)
    return txt

# =============================================================================
# COMENTARIO PAQ / TYP
# =============================================================================
def tiene_comentario_paq(comentario) -> bool:
    txt = str(comentario).upper().strip()
    return "PAQ" in txt or "TYP" in txt or "T/P" in txt

def extraer_num_paquetes(comentario) -> int:
    """
    Extrae la cantidad de paquetes del comentario.
    Cubre: "2 PAQ", "PAQ2", "18PAQ", "PAQ-5", "PAQ,3",
           "PAQ:3", "5-PAQ", "TYP", "T/P", "2TYP", etc.
    """
    import re as _re
    txt = str(comentario).upper().strip()
    if not txt:
        return 1

    # T/P siempre = 2
    if "T/P" in txt:
        return 2

    # TYP: buscar numero asociado
    if "TYP" in txt:
        m = _re.search(r'(\d+)', txt)
        if m:
            n = int(m.group(1))
            if 2 <= n <= 999:
                return n
        return 2

    # PAQ/PAQS/PAQUETE/PAQUETES: buscar cualquier numero en el texto
    if "PAQ" in txt:
        m = _re.search(r'(\d+)', txt)
        if m:
            n = int(m.group(1))
            if 2 <= n <= 999:
                return n
        return 2

    return 1

# =============================================================================
# TIPO DE SERVICIO
# =============================================================================
def obtener_tipo_servicio(valor_k) -> str:
    """
    Resuelve NRM/PLUS/ECO usando las ~1500 variantes de escritura.
    Si no reconoce nada -> NRM por defecto (nunca FALTA en col tipo servicio).
    """
    resultado = _svc.resolver_servicio(valor_k)
    if resultado in ("VACIO", "FALTA", "NINGUNO"):
        return SERVICIO_DEFAULT  # NRM por defecto
    return resultado  # NRM / PLUS / ECO

def detectar_tipo_col_q(encabezado) -> str:
    if encabezado is None:
        return "NINGUNO"
    # Normalizar: quitar puntos, comas, guiones para comparar
    enc = str(encabezado).upper().strip()
    enc = enc.replace(".", " ").replace(",", " ").replace("-", " ")
    enc = enc.replace("_", " ").replace("/", " ")
    if not enc.strip():
        return "NINGUNO"
    if "TIPO" in enc or "SERV" in enc or "ENVIO" in enc:
        return "SERVICIO"
    if "MUNIC" in enc or "CIUDAD" in enc or "LOCALID" in enc or "DEPTO" in enc:
        return "MUNICIPIO"
    return "NINGUNO"

# =============================================================================
# DETECTAR COLUMNAS ESPECIALES EN HOJA DESTINO
# Lee encabezados fila 5 y devuelve dict con columnas encontradas
# =============================================================================
def _enc_norm_kw(txt) -> str:
    """Normaliza encabezado: mayusculas, sin puntos/comas/guiones."""
    s = str(txt or "").upper().strip()
    for c in ".,_/-": s = s.replace(c, " ")
    return " ".join(s.split())


def _coincide_enc(enc_norm, palabras_exactas):
    """
    True solo si el encabezado normalizado ES EXACTAMENTE
    una de las palabras, o la CONTIENE como palabra completa.
    Evita que "ECO" coincida con "F RECOLECTA" o "TECOMUNICACIONES".
    """
    import re as _re
    for kw in palabras_exactas:
        # Coincidencia exacta total
        if enc_norm == kw:
            return True
        # Coincidencia como palabra completa (no substring de otra palabra)
        patron = r'(?<![A-Z0-9])' + _re.escape(kw) + r'(?![A-Z0-9])'
        if _re.search(patron, enc_norm):
            return True
    return False


def detectar_cols_especiales(ws) -> dict:
    """
    Detecta columnas especiales leyendo encabezados en la fila de encabezado real.
    Usa coincidencia de PALABRA COMPLETA para evitar falsos positivos.

    Libros Metrogalerias   -> tienen "TIPO SERVICIO" o "MUNIC" -> col_q
    Libros Metropolitanos  -> tienen "MUNICIPIO" o "CIUDAD"    -> col_q
    Libros UT SOFTWARE     -> tienen "PAQUETES"                -> col_paq
    Libros con orden ID    -> tienen "ORDEN ID"                -> col_oid
    Libros normales        -> ninguna de las anteriores        -> todo 0
    """
    # Palabras que deben coincidir como PALABRA COMPLETA, no substring
    KW_SERVICIO  = {"TIPO SERVICIO", "TIPO DE SERVICIO",
                    "TIPOSERVICIO", "SCGE", "SERVICIO"}
    KW_MUNICIPIO = {"MUNICIPIO", "MUNIC", "CIUDAD", "LOCALIDAD",
                    "DEPTO", "DEPARTAMENTO", "MUNICIPIO DESTINO",
                    "MUNIC DESTINO"}
    KW_PAQUETES  = {"PAQUETES", "PAQUETE", "CANTIDAD PAQUETES",
                    "CANTIDAD PAQ", "NUM PAQUETES"}
    KW_ORDEN_ID  = {"ORDEN ID", "ORDER ID", "ORDEN ID 2",
                    "ORDER ID 2", "ID 2", "ID2"}

    resultado = {"tipo_q": "NINGUNO", "col_q": 0, "col_paq": 0, "col_oid": 0}

    fila_enc = _buscar_fila_encabezado(ws)

    for col in range(1, 60):
        raw = ws.cell(row=fila_enc, column=col).value
        if raw is None:
            continue
        enc_n = _enc_norm_kw(raw)

        if resultado["col_q"] == 0:
            if _coincide_enc(enc_n, KW_SERVICIO):
                resultado["tipo_q"] = "SERVICIO"
                resultado["col_q"]  = col
                continue
            if _coincide_enc(enc_n, KW_MUNICIPIO):
                resultado["tipo_q"] = "MUNICIPIO"
                resultado["col_q"]  = col
                continue

        if resultado["col_paq"] == 0:
            if _coincide_enc(enc_n, KW_PAQUETES):
                resultado["col_paq"] = col
                continue

        if resultado["col_oid"] == 0:
            if _coincide_enc(enc_n, KW_ORDEN_ID):
                resultado["col_oid"] = col
                continue

    return resultado

# =============================================================================
# VALIDAR HOJA DESTINO
# =============================================================================
def _norm_enc(txt) -> str:
    """Normaliza encabezado: sin puntos, comas, guiones, espacios extra."""
    import re as _re
    s = str(txt or "").upper().strip()
    s = s.replace(".", "").replace(",", "").replace("-", "").replace(
        "_", "").replace("/", "").replace(" ", "")
    return s

# Mapa de encabezados validos normalizados -> columna esperada
_ENCABEZADOS_NORM = None

def _get_enc_norm():
    global _ENCABEZADOS_NORM
    if _ENCABEZADOS_NORM is None:
        _ENCABEZADOS_NORM = {
            col: _norm_enc(enc)
            for col, enc in ENCABEZADOS_VALIDOS.items()
        }
    return _ENCABEZADOS_NORM


def _buscar_fila_encabezado(ws) -> int:
    """
    Busca la fila del encabezado buscando F.RECOLECTA en las primeras 15 filas.
    Devuelve el numero de fila encontrado, o FILA_ENCABEZADO si no encuentra.
    """
    clave = "F.RECOLECTA"
    clave_norm = _norm_enc(clave)
    for fila in range(1, 16):
        for col in range(1, 25):
            val = ws.cell(row=fila, column=col).value
            if val and _norm_enc(str(val)) == clave_norm:
                return fila
    return FILA_ENCABEZADO


def hoja_valida(ws) -> bool:
    """Valida encabezados buscando la fila dinamicamente."""
    fila_enc = _buscar_fila_encabezado(ws)
    for col_1based, esperado in ENCABEZADOS_VALIDOS.items():
        actual = str(ws.cell(row=fila_enc,
                             column=col_1based).value or "").strip().upper()
        if actual != esperado.upper():
            return False
    return True


def get_fila_encabezado(ws) -> int:
    """Devuelve la fila real del encabezado en esta hoja."""
    return _buscar_fila_encabezado(ws)

# =============================================================================
# COLOR AMARILLO
# =============================================================================
def es_color_amarillo(celda) -> bool:
    try:
        fill = celda.fill
        if fill is None:
            return False
        fg = fill.fgColor
        if fg is None:
            return False
        if fg.type == "rgb":
            color_rgb = str(fg.rgb).upper().replace("#", "")
            rrggbb = color_rgb[-6:] if len(color_rgb) >= 6 else color_rgb
            for amarillo in COLORES_AMARILLO:
                if rrggbb == amarillo[-6:].upper():
                    return True
        return False
    except Exception:
        return False

# =============================================================================
# HELPERS DE VALOR
# =============================================================================
def _es_valor_real(v) -> bool:
    """True solo si la celda tiene contenido real.
    Ignora: None, string vacio, y cero numerico (residuo de formulas)."""
    if v is None:
        return False
    if isinstance(v, str) and v.strip() == "":
        return False
    if isinstance(v, (int, float)) and v == 0:
        return False
    return True


# =============================================================================
# ULTIMA FILA CON DATOS REALES
# Usa iter_rows vectorizado en el rango D:J — mas rapido que cell() por celda
# =============================================================================
def ultima_fila_con_datos(ws) -> int:
    col_ini    = COL_DATOS_INI
    col_fin    = COL_DATOS_INI + 6
    fila_datos = _buscar_fila_encabezado(ws) + 1
    ultima     = fila_datos - 1
    for row_idx, row in enumerate(
            ws.iter_rows(min_row=fila_datos, max_row=ws.max_row,
                         min_col=col_ini, max_col=col_fin, values_only=True),
            start=fila_datos):
        if any(_es_valor_real(v) for v in row):
            ultima = row_idx
    return ultima

# =============================================================================
# PRIMERA FILA LIBRE
# =============================================================================
def primera_fila_libre(ws, fila_base: int) -> int:
    fila_datos = _buscar_fila_encabezado(ws) + 1
    fila    = max(fila_base, fila_datos)
    col_ini = COL_DATOS_INI
    col_fin = COL_DATOS_INI + 6
    limite  = fila + 5000
    while fila < limite:
        ocupada = False
        for col in range(col_ini, col_fin + 1):
            celda = ws.cell(row=fila, column=col)
            if _es_valor_real(celda.value):
                ocupada = True
                break
            if es_color_amarillo(celda):
                ocupada = True
                break
        if not ocupada:
            return fila
        fila += 1
    return fila


def primera_fila_libre_rapida(ws, fila_base: int) -> int:
    """
    Ultra-rapida: solo verifica col D para decidir.
    Usar cuando fila_base viene del cache (ya es correcta o muy cerca).
    Fallback a primera_fila_libre si no encuentra en 20 filas.
    """
    fila_datos = _buscar_fila_encabezado(ws) + 1
    fila   = max(fila_base, fila_datos)
    limite = fila + 20
    while fila < limite:
        v = ws.cell(row=fila, column=COL_DATOS_INI).value
        if not _es_valor_real(v):
            if not es_color_amarillo(ws.cell(row=fila, column=COL_DATOS_INI)):
                return fila
        fila += 1
    return primera_fila_libre(ws, fila_base)

# =============================================================================
# ENCONTRAR COLUMNA LIBRE PARA COMENTARIO
# Empieza en COL_DEST_COMENTAR (R), salta si tiene texto o color
# =============================================================================
def encontrar_col_comentario(ws, fila: int) -> int:
    for col in range(COL_DEST_COMENTAR, COL_DEST_COMENTAR + 30):
        celda = ws.cell(row=fila, column=col)
        if celda.value is not None and str(celda.value).strip() != "":
            continue
        try:
            fill = celda.fill
            if fill and fill.fill_type not in (None, "none"):
                fg = fill.fgColor
                if fg and fg.type != "none":
                    rgb = ""
                    if fg.type == "rgb":
                        rgb = str(fg.rgb).upper()
                    elif fg.type == "theme":
                        rgb = "HAS_THEME"
                    if rgb and rgb not in ("00000000", "FFFFFFFF", ""):
                        continue
        except Exception:
            pass
        return col
    return 0

# =============================================================================
# CARGAR IDs EN DESTINO
# Usa iter_rows para acceso vectorizado — mas rapido que while celda a celda
# =============================================================================
def cargar_ids_destino(ws) -> dict:
    ids = {}
    fila_datos = _buscar_fila_encabezado(ws) + 1
    for row in ws.iter_rows(min_row=fila_datos, max_row=ws.max_row,
                             min_col=COL_ID, max_col=COL_ID, values_only=True):
        raw = row[0]
        if raw is not None:
            v = str(raw).strip()
            if v and not v.startswith("#"):
                ids[v] = ids.get(v, 0) + 1
    return ids

# =============================================================================
# CARGAR MAPA DE TIENDAS
# =============================================================================
def cargar_mapa_tiendas(ws_tiendas, fila_ini: int) -> dict:
    mapa = {}
    if ws_tiendas is None:
        return mapa
    for fila in range(fila_ini, ws_tiendas.max_row + 1):
        kA = normalizar(ws_tiendas.cell(row=fila, column=1).value)
        kB = normalizar(ws_tiendas.cell(row=fila, column=2).value)
        if kA and kB and kA not in mapa:
            mapa[kA] = kB
    return mapa

# =============================================================================
# CONVERTIR VALOR AL TIPO CORRECTO
# - Fechas  -> "12-mar"
# - Numeros -> int o float
# - Texto   -> texto limpio
# =============================================================================
def _convertir_valor(valor, es_fecha=False):
    if valor is None:
        return None

    if isinstance(valor, (datetime, date)) or es_fecha:
        try:
            if isinstance(valor, str):
                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
                    try:
                        valor = datetime.strptime(valor.strip(), fmt)
                        break
                    except ValueError:
                        continue

            # CAMBIO CLAVE: devolver datetime real, no string
            if isinstance(valor, (datetime, date)):
                return valor

            return str(valor)
        except Exception:
            return str(valor)

    if isinstance(valor, (int, float)):
        return valor

    s = str(valor).strip()
    if s == "":
        return None

    s_limpio = s.replace("$", "").replace(",", "").strip()
    try:
        return int(s_limpio) if "." not in s_limpio else float(s_limpio)
    except ValueError:
        return s
# =============================================================================
# INSERTAR PAQUETE
# Solo toca: D:J, col tipo_q/municipio, col paquetes, col orden_id, comentario R
# =============================================================================
def insertar_paquete(ws, fila_dest, arr_datos, precio_cero,
                     valor_k, valor_m, valor_n,
                     es_primero, tipo_q, col_q=0,
                     col_paq=0, col_oid=0,
                     forzar_sin_comentario=False):

    COL_D = COL_DATOS_INI        # 4 = D
    COL_J = COL_DATOS_INI + 6   # 10 = J

    # D:J — no tocar celdas con formula
    for j in range(7):
        col   = COL_D + j
        valor = arr_datos[j] if j < len(arr_datos) else None
        celda = ws.cell(row=fila_dest, column=col)

        v_actual = celda.value
        if v_actual is not None and str(v_actual).strip().startswith("="):
            continue

        if col == COL_J and precio_cero:
            celda.value = 0
            continue

        es_fecha = (col == COL_DATOS_INI)
        celda.value = _convertir_valor(valor, es_fecha=es_fecha)
        if es_fecha and isinstance(celda.value, (datetime, date)):
            celda.number_format = "D-MMM"

    # Tipo servicio / municipio
    # PROTECCION: col_q debe estar fuera del rango D:J (cols 4-10)
    # Si col_q apunta a una columna de datos, ignorar para no corromper
    if tipo_q in ("SERVICIO", "MUNICIPIO") and col_q > COL_J:
        celda_q = ws.cell(row=fila_dest, column=col_q)
        if tipo_q == "SERVICIO":
            celda_q.value = obtener_tipo_servicio(valor_k)
        else:
            celda_q.value = str(valor_m).strip() if valor_m else ""

    # Paquetes — solo si esta fuera del rango D:J
    if col_paq > COL_J:
        num_paq = extraer_num_paquetes(valor_n)
        ws.cell(row=fila_dest, column=col_paq).value = max(num_paq, 1)

    # Orden ID — solo si esta fuera del rango D:J
    if col_oid > COL_J:
        val_oid = str(valor_k).strip() if valor_k else ""
        ws.cell(row=fila_dest, column=col_oid).value = val_oid if val_oid else "EN BLANCO"

    # Comentario — en R o siguiente columna libre (sin texto ni color)
    if es_primero and tiene_comentario_paq(valor_n) and not forzar_sin_comentario:
        col_com = encontrar_col_comentario(ws, fila_dest)
        if col_com > 0:
            ws.cell(row=fila_dest, column=col_com).value = valor_n

# =============================================================================
# EVALUAR DUPLICADO
# =============================================================================
def evaluar_duplicado(id_nuevo, valor_n, ids_existentes, num_paquetes):
    """
    Reglas exactas:

    1. Sin comentario PAQ/TYP:
       - Siempre DUP — no importa si el ID existe o no.
         Un paquete sin comentario de multiples solo se ingresa
         la primera vez que aparece. Si ya existe o no tiene
         comentario -> DUP.
       EXCEPCION: si el ID no existe aun -> OK, insertar 1.

    2. Con comentario PAQ/TYP:
       - ID no existe -> OK, insertar los N paquetes juntos
       - ID existe y ya tiene todos -> DUP
       - ID existe y tiene algunos -> DUP (ya fue procesado antes,
         no fragmentar — deben ir juntos de una sola vez)
       - Mismo ID con comentario aparece dos veces en INGRESO:
         el primero se procesa, el segundo es DUP
    """
    tiene_paq = tiene_comentario_paq(valor_n)

    # Sin comentario: solo insertar si el ID es completamente nuevo
    if not tiene_paq:
        if not id_nuevo or id_nuevo in ids_existentes:
            return "DUP", 0
        return "OK", 1

    # Con comentario PAQ/TYP: ID debe ser completamente nuevo
    # Si ya existe (parcial o total) -> DUP para no fragmentar
    if not id_nuevo or id_nuevo in ids_existentes:
        return "DUP", 0

    # ID nuevo con comentario -> insertar todos juntos de una vez
    return "OK", num_paquetes
