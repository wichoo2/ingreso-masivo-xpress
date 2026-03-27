# =============================================================================
# CONFIG_LOCAL.PY - Configuracion para modo local
# Solo edita las rutas de abajo con las tuyas.
# =============================================================================

import os

# --- Rutas ---
CARPETA_TIENDAS = r"C:/Users/USUARIO/OneDrive/Clientes_Xpress_2.0"
ARCHIVO_INGRESO = r"C:/Users/USUARIO/Downloads/INGRESO_MAS_v5 (3)/INGRESO_MASIVO.xlsx"

# --- Hojas ---
HOJA_ORIGEN      = "INGRESO_MASIVO"
HOJA_TIENDAS     = "TIENDAS"
FILA_INICIO      = 4
FILA_ENCABEZADO  = 5
FILA_DATOS_DEST  = 6
FILA_TIENDAS_INI = 2

# --- Columnas en INGRESO_MASIVO (1=A, 2=B ...) ---
COL_RESULTADO  = 2   # B
COL_DATOS_INI  = 4   # D
COL_TIENDA     = 5   # E
COL_ID         = 6   # F
COL_TIPOSERV   = 11  # K  tipo servicio NRM/PLUS/ECO
COL_ORDEN_ID2  = 11  # K  ORDEN ID 2
COL_MUNICIPIO  = 13  # M
COL_COMENTARIO = 14  # N

# --- Columnas en archivos DESTINO ---
COL_DEST_PRECIO   = 10  # J
COL_DEST_TIPOSERV = 17  # Q
COL_DEST_COMENTAR = 18  # R

# --- Encabezados que validan una hoja destino (fila 5) ---
ENCABEZADOS_VALIDOS = {
    4:  "F.RECOLECTA",
    5:  "TIENDA",
    6:  "ID",
    7:  "NOMBRE",
    8:  "ZONA",
    9:  "TELEFONO",
    10: "PRECIO",
}

# --- Servicios validos ---
SERVICIOS_VALIDOS = {"NRM", "PLUS", "ECO"}
SERVICIO_DEFAULT  = "NRM"

# --- Color amarillo divisor (separador de secciones) ---
# Solo estos colores bloquean el pegado — verde, azul y otros se ignoran
COLORES_AMARILLO = {
    "FFFF00",
    "FFFF33",
    "FFCC00",
}
