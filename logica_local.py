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
# NORMALIZAR ID — convierte float a entero para comparaciones exactas
# =============================================================================
def normalizar_id(raw) -> str:
    """
    Convierte cualquier representación de ID a string entero limpio.
    3936869   → '3936869'
    3936869.0 → '3936869'
    '3936869.0' → '3936869'
    None / '' → ''
    """
    if raw is None:
        return ""
    if isinstance(raw, float):
        return str(int(raw)) if raw == int(raw) else str(raw)
    s = str(raw).strip()
    if not s:
        return ""
    if "." in s:
        try:
            f = float(s)
            if f == int(f):
                return str(int(f))
        except (ValueError, OverflowError):
            pass
    return s

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

_RE_TIPO_SUFIJO = re.compile(
    r"[\s\-_.,;()\[\]/]*(scge|pls|plus|lockers?|metrogalerias?|"
    r"metropolitano|diamante|bonny|ss|sv)"
    r"[\s\-_.,;()\[\]/]*$",
    re.IGNORECASE,
)

def normalizar_sin_tipo(txt) -> str:
    s = str(txt or "").strip()
    for _ in range(4):
        nuevo = _RE_TIPO_SUFIJO.sub("", s).strip()
        if not nuevo or nuevo == s:
            break
        s = nuevo
    resultado = normalizar(s)
    return resultado if len(resultado) >= 4 else normalizar(txt)

# =============================================================================
# COMENTARIO PAQ / TYP
# =============================================================================
def tiene_comentario_paq(comentario) -> bool:
    txt = str(comentario).upper().strip()
    return "PAQ" in txt or "TYP" in txt or "T/P" in txt

def extraer_num_paquetes(comentario) -> int:
    import re as _re
    txt = str(comentario).upper().strip()
    if not txt:
        return 1
    if "T/P" in txt:
        return 2
    if "TYP" in txt:
        m = _re.search(r'(\d+)', txt)
        if m:
            n = int(m.group(1))
            if 2 <= n <= 999:
                return n
        return 2
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
    resultado = _svc.resolver_servicio(valor_k)
    if resultado in ("VACIO", "FALTA", "NINGUNO"):
        return SERVICIO_DEFAULT
    return resultado

def detectar_tipo_col_q(encabezado) -> str:
    if encabezado is None:
        return "NINGUNO"
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
# DETECTAR COLUMNAS ESPECIALES
# =============================================================================
def _enc_norm_kw(txt) -> str:
    s = str(txt or "").upper().strip()
    for c in ".,_/-": s = s.replace(c, " ")
    return " ".join(s.split())

def _coincide_enc(enc_norm, palabras_exactas):
    import re as _re
    for kw in palabras_exactas:
        if enc_norm == kw:
            return True
        patron = r'(?<![A-Z0-9])' + _re.escape(kw) + r'(?![A-Z0-9])'
        if _re.search(patron, enc_norm):
            return True
    return False

def detectar_cols_especiales(ws) -> dict:
    KW_SERVICIO  = {"TIPO SERVICIO", "TIPO DE SERVICIO", "TIPOSERVICIO", "SCGE", "SERVICIO"}
    KW_MUNICIPIO = {"MUNICIPIO", "MUNIC", "CIUDAD", "LOCALIDAD",
                    "DEPTO", "DEPARTAMENTO", "MUNICIPIO DESTINO", "MUNIC DESTINO"}
    KW_PAQUETES  = {"PAQUETES", "PAQUETE", "CANTIDAD PAQUETES", "CANTIDAD PAQ", "NUM PAQUETES"}
    KW_ORDEN_ID  = {"ORDEN ID", "ORDER ID", "ORDEN ID 2", "ORDER ID 2", "ID 2", "ID2"}

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
    import re as _re
    s = str(txt or "").upper().strip()
    s = s.replace(".", "").replace(",", "").replace("-", "").replace(
        "_", "").replace("/", "").replace(" ", "")
    return s

_ENCABEZADOS_NORM = None

def _get_enc_norm():
    global _ENCABEZADOS_NORM
    if _ENCABEZADOS_NORM is None:
        _ENCABEZADOS_NORM = {col: _norm_enc(enc) for col, enc in ENCABEZADOS_VALIDOS.items()}
    return _ENCABEZADOS_NORM

_CACHE_FILA_ENC = {}

def _buscar_fila_encabezado(ws) -> int:
    ws_id = id(ws)
    if ws_id in _CACHE_FILA_ENC:
        return _CACHE_FILA_ENC[ws_id]
    clave      = "F.RECOLECTA"
    clave_norm = _norm_enc(clave)
    fila_enc   = FILA_ENCABEZADO
    for fila in range(1, 16):
        for col in range(1, 25):
            val = ws.cell(row=fila, column=col).value
            if val and _norm_enc(str(val)) == clave_norm:
                fila_enc = fila
                break
        else:
            continue
        break
    _CACHE_FILA_ENC[ws_id] = fila_enc
    return fila_enc

def hoja_valida(ws) -> bool:
    """
    Valida que la hoja tenga los encabezados esperados.
    FIX 1: Los encabezados se validan relativos a donde está F.RECOLECTA,
    no en columnas absolutas. Si una tienda agregó columnas al inicio,
    la validación sigue funcionando.
    FIX 2: Se eliminó el check de ws.protection.sheet — openpyxl puede
    escribir en hojas con protección sin contraseña, y el check descartaba
    hojas válidas que solo tienen protección de encabezados.
    """
    fila_enc = _buscar_fila_encabezado(ws)

    # Encontrar la columna donde está F.RECOLECTA en esta hoja
    col_recolecta = None
    clave_norm = _norm_enc("F.RECOLECTA")
    for col in range(1, 25):
        val = ws.cell(row=fila_enc, column=col).value
        if val and _norm_enc(str(val)) == clave_norm:
            col_recolecta = col
            break

    if col_recolecta is None:
        return False  # No encontró F.RECOLECTA

    # Validar encabezados relativos a la posición de F.RECOLECTA
    # ENCABEZADOS_VALIDOS tiene col 4 = F.RECOLECTA como referencia
    # offset = col_recolecta - 4
    offset = col_recolecta - 4
    for col_base, esperado in ENCABEZADOS_VALIDOS.items():
        col_real = col_base + offset
        actual = str(ws.cell(row=fila_enc, column=col_real).value or "").strip().upper()
        esperado_norm = esperado.strip().upper().replace(" ", "").replace(".", "")
        actual_norm   = actual.replace(" ", "").replace(".", "")
        if actual_norm != esperado_norm:
            return False
    return True

def get_fila_encabezado(ws) -> int:
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

def _fila_tiene_id(ws, fila: int) -> bool:
    """
    Si col F (ID) tiene cualquier valor no vacío, la fila está ocupada.
    Protección extra contra sobreescritura de filas con precio $0.
    """
    v = ws.cell(row=fila, column=COL_ID).value
    return v is not None and str(v).strip() not in ("", "0")

# =============================================================================
# ULTIMA FILA CON DATOS REALES
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
# PRIMERA FILA LIBRE — respeta la línea amarilla como barrera
# MEJORA: inserta SIEMPRE antes de la última línea amarilla (no después).
# Si no hay amarilla, comportamiento original (primera fila vacía al final).
# =============================================================================
def _buscar_ultima_amarilla(ws) -> int:
    """
    Devuelve la fila de la ÚLTIMA línea amarilla del archivo.
    0 si no existe ninguna.
    Solo busca en cols D:J para velocidad.
    """
    col_ini    = COL_DATOS_INI
    col_fin    = COL_DATOS_INI + 6
    fila_datos = _buscar_fila_encabezado(ws) + 1
    ultima_am  = 0
    for fila in range(fila_datos, ws.max_row + 1):
        for col in range(col_ini, col_fin + 1):
            if es_color_amarillo(ws.cell(row=fila, column=col)):
                ultima_am = fila
                break
    return ultima_am

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
def encontrar_col_comentario(ws, fila: int) -> int:
    """
    Busca la primera columna libre desde COL_DEST_COMENTAR hacia la derecha.
    FIX: Ahora también salta celdas con:
      - Fórmulas (valor empieza con '=')
      - Estilos definidos (bordes, número de formato) que indican celda de plantilla
      - Colores de fill (como antes)
    Evita sobreescribir columnas de fórmulas de liquidación.
    """
    for col in range(COL_DEST_COMENTAR, COL_DEST_COMENTAR + 30):
        celda = ws.cell(row=fila, column=col)

        # 1. Tiene texto o fórmula — ocupada
        v = celda.value
        if v is not None:
            sv = str(v).strip()
            if sv != "":
                continue  # texto visible
            if sv.startswith("="):
                continue  # fórmula vacía

        # 2. Tiene número de formato personalizado — celda de plantilla
        try:
            nf = celda.number_format
            if nf and nf not in (None, "General", "@", ""):
                continue
        except Exception:
            pass

        # 3. Tiene borde definido — celda de plantilla
        try:
            b = celda.border
            if b and any([
                b.left and b.left.style,
                b.right and b.right.style,
                b.top and b.top.style,
                b.bottom and b.bottom.style,
            ]):
                continue
        except Exception:
            pass

        # 4. Tiene color de fill distinto de blanco/transparente
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
# CARGAR IDs EN DESTINO — MEJORA: normalizar float a entero
# =============================================================================
def cargar_ids_destino(ws) -> dict:
    """
    Carga todos los IDs existentes en el archivo de tienda.
    MEJORA: normaliza IDs con normalizar_id() para que
    '3936869' y '3936869.0' sean la misma clave.
    """
    ids = {}
    fila_datos = _buscar_fila_encabezado(ws) + 1
    for row in ws.iter_rows(min_row=fila_datos, max_row=ws.max_row,
                             min_col=COL_ID, max_col=COL_ID, values_only=True):
        raw = row[0]
        if raw is not None:
            v = normalizar_id(raw)
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
# CONVERTIR ID PARA ESCRITURA — siempre entero si es número
# =============================================================================
def _convertir_id_escritura(raw):
    """
    Convierte el ID al formato correcto para escribir en la tienda.
    Si es numérico → escribe como int (no como float).
    Evita que quede '3936869.0' en col F.
    """
    if raw is None:
        return None
    if isinstance(raw, float):
        return int(raw) if raw == int(raw) else raw
    s = str(raw).strip()
    if "." in s:
        try:
            f = float(s)
            if f == int(f):
                return int(f)
        except (ValueError, OverflowError):
            pass
    try:
        return int(s)
    except (ValueError, TypeError):
        return s if s else None

# =============================================================================
# INSERTAR PAQUETE — MEJORA: ID siempre como entero en col F
# =============================================================================
def insertar_paquete(ws, fila_dest, arr_datos, precio_cero,
                     valor_k, valor_m, valor_n,
                     es_primero, tipo_q, col_q=0,
                     col_paq=0, col_oid=0,
                     forzar_sin_comentario=False):

    COL_D  = COL_DATOS_INI        # 4 = D
    COL_J  = COL_DATOS_INI + 6   # 10 = J
    COL_F  = COL_DATOS_INI + 2   # 6  = F (ID)

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

        if col == COL_F:
            # MEJORA: ID siempre como entero limpio — nunca '3936869.0'
            celda.value = _convertir_id_escritura(valor)
            continue

        es_fecha = (col == COL_DATOS_INI)
        celda.value = _convertir_valor(valor, es_fecha=es_fecha)
        if es_fecha and isinstance(celda.value, (datetime, date)):
            celda.number_format = "D-MMM"

    if tipo_q in ("SERVICIO", "MUNICIPIO") and col_q > COL_J:
        celda_q = ws.cell(row=fila_dest, column=col_q)
        if tipo_q == "SERVICIO":
            celda_q.value = obtener_tipo_servicio(valor_k)
        else:
            celda_q.value = str(valor_m).strip() if valor_m else ""

    if col_paq > COL_J:
        num_paq = extraer_num_paquetes(valor_n)
        ws.cell(row=fila_dest, column=col_paq).value = max(num_paq, 1)

    if col_oid > COL_J:
        val_oid = str(valor_k).strip() if valor_k else ""
        ws.cell(row=fila_dest, column=col_oid).value = val_oid if val_oid else "EN BLANCO"

    if es_primero and tiene_comentario_paq(valor_n) and not forzar_sin_comentario:
        col_com = encontrar_col_comentario(ws, fila_dest)
        if col_com > 0:
            ws.cell(row=fila_dest, column=col_com).value = valor_n

# =============================================================================
# EVALUAR DUPLICADO — MEJORA: normalizar ID antes de comparar
# =============================================================================
def evaluar_duplicado(id_nuevo, valor_n, ids_existentes, num_paquetes):
    """
    MEJORA: id_nuevo se normaliza con normalizar_id() antes de comparar,
    para que '3936869' y '3936869.0' sean tratados como el mismo ID.

    Motivos de DUP diferenciados:
      DUP_EXISTE  → ID ya está en la tienda
      DUP_SESION  → ID apareció dos veces en el INGRESO en la misma sesión
      DUP_PAQ     → PAQ pero ID ya existe (no fragmentar)
    """
    id_norm   = normalizar_id(id_nuevo)
    tiene_paq = tiene_comentario_paq(valor_n)

    if not tiene_paq:
        if not id_norm or id_norm in ids_existentes:
            motivo = "DUP_EXISTE" if id_norm in ids_existentes else "DUP_VACIO"
            return motivo, 0
        return "OK", 1

    if not id_norm or id_norm in ids_existentes:
        motivo = "DUP_PAQ" if id_norm in ids_existentes else "DUP_VACIO"
        return motivo, 0

    return "OK", num_paquetes
