# =============================================================================
# SERVICIOS_VARIANTES.PY
# Mapeo de variantes escritas mal (o bien) hacia el servicio correcto.
# =============================================================================
# HOW TO ADD MORE:
#   Agrega entradas al diccionario VARIANTES_EXTRA al final del archivo.
#   Formato: "variante_en_minusculas": "SERVICIO_CORRECTO"
#   Ejemplo: "nrmal": "NRM"
#
# Los 3 servicios validos son: NRM, PLUS, ECO
# Si un valor no esta en ninguna variante -> se marca FALTA
# =============================================================================

import re
import unicodedata


def _n(txt):
    """Normaliza: minusculas, sin acentos, solo letras y numeros."""
    if txt is None:
        return ""
    txt = str(txt).lower().strip()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"[^a-z0-9]", "", txt)
    return txt


# =============================================================================
# VARIANTES GENERADAS AUTOMATICAMENTE
# =============================================================================
_MAPA = {}


def _registrar(variantes_raw, servicio):
    for v in variantes_raw:
        k = _n(v)
        if k and k not in _MAPA:
            _MAPA[k] = servicio


def _generar(base):
    """Genera variantes tipograficas de una palabra base."""
    variantes = set()
    b = base.upper()

    # Exacta y con espacios
    variantes.add(b)
    for i in range(1, 4):
        variantes.add(" " * i + b)
        variantes.add(b + " " * i)

    # Letras duplicadas / triplicadas
    for idx in range(len(b)):
        variantes.add(b[:idx] + b[idx] + b[idx:])
        variantes.add(b[:idx] + b[idx] * 2 + b[idx:])

    # Letras faltantes
    for idx in range(len(b)):
        variantes.add(b[:idx] + b[idx + 1:])

    # Letras adyacentes en teclado
    teclado = {
        'Q': 'WA',  'W': 'QEA', 'E': 'WRS', 'R': 'ET',  'T': 'RYG',
        'Y': 'TUH', 'U': 'YIJ', 'I': 'UOK', 'O': 'IPL', 'P': 'O',
        'A': 'QWS', 'S': 'AWDE','D': 'SEF', 'F': 'DGR', 'G': 'FHT',
        'H': 'GJY', 'J': 'HKU', 'K': 'JLI', 'L': 'KO',
        'Z': 'AX',  'X': 'ZCS', 'C': 'XVD', 'V': 'CBF', 'B': 'VNG',
        'N': 'BM',  'M': 'N',
    }
    for idx, letra in enumerate(b):
        for alt in teclado.get(letra, ''):
            variantes.add(b[:idx] + alt + b[idx + 1:])

    # Con separadores
    for sep in ['.', '-', '/', '_', ' ']:
        for idx in range(1, len(b)):
            variantes.add(b[:idx] + sep + b[idx:])

    # Mayusculas/minusculas/mixto
    copia = list(variantes)
    for v in copia:
        variantes.add(v.lower())
        variantes.add(v.capitalize())

    # Con prefijos comunes
    copia2 = list(variantes)
    for v in copia2:
        for pre in ['TIP', 'TIPO', 'SERV', 'S/']:
            variantes.add(pre + v)
            variantes.add(pre + ' ' + v)

    return variantes


_registrar(_generar("NRM"),  "NRM")
_registrar(_generar("PLUS"), "PLUS")
_registrar(_generar("ECO"),  "ECO")

# Alias comunes manuales
_registrar([
    "normal", "nrml", "norm", "nrm1", "nrm0",
    "nrmai", "nrmai", "nrma",
    "standard", "estandar", "est",
], "NRM")

_registrar([
    "plus1", "plus2", "pl us", "plu s", "p.l.u.s",
    "pls", "pluss", "pluus", "plusss",
    "expres", "express", "rapido", "prioritario",
], "PLUS")

_registrar([
    "economico", "econ", "ec0", "ec o", "e.c.o",
    "ecoo", "eeco", "ecco", "economy",
    "barato", "lento", "ordinario",
], "ECO")


# =============================================================================
# VARIANTES EXTRA — AGREGA AQUI TUS PROPIAS VARIANTES
# Formato: "texto_que_llega": "SERVICIO_CORRECTO"
# =============================================================================
VARIANTES_EXTRA = {
    # Ejemplos (puedes borrarlos o agregar los tuyos):
    # "nr": "NRM",
    # "pl": "PLUS",
    # "ec": "ECO",
}

# Registrar variantes extra
for v_raw, srv in VARIANTES_EXTRA.items():
    k = _n(v_raw)
    if k:
        _MAPA[k] = srv


# =============================================================================
# FUNCION PUBLICA
# =============================================================================
def resolver_servicio(valor_k) -> str:
    """
    Recibe el valor de la columna K (tipo servicio) de INGRESO_MASIVO.
    Devuelve: "NRM", "PLUS", "ECO", "VACIO" o "FALTA"
      - VACIO  -> columna K estaba en blanco -> se usara NRM por defecto
      - FALTA  -> valor no reconocido -> marcar como FALTA
    """
    if valor_k is None:
        return "VACIO"
    s = str(valor_k).strip()
    if s == "":
        return "VACIO"

    # Buscar en el mapa normalizado
    k = _n(s)
    if not k:
        return "VACIO"

    resultado = _MAPA.get(k)
    if resultado:
        return resultado

    # Ultimo intento: coincidencia parcial si contiene la palabra clave
    if "nrm" in k or "norm" in k:
        return "NRM"
    if "plus" in k or "plu" in k:
        return "PLUS"
    if "eco" in k:
        return "ECO"

    return "FALTA"


if __name__ == "__main__":
    # Test rapido
    pruebas = [
        ("NRM", "NRM"), ("PLUS", "PLUS"), ("ECO", "ECO"),
        ("nrm", "NRM"), ("plus", "PLUS"), ("eco", "ECO"),
        ("NRMM", "NRM"), ("PLUSS", "PLUS"), ("ECOO", "ECO"),
        ("NR M", "NRM"), ("PL US", "PLUS"),
        ("", "VACIO"), (None, "VACIO"),
        ("XPTO", "FALTA"), ("123", "FALTA"),
    ]
    ok = err = 0
    for entrada, esperado in pruebas:
        res = resolver_servicio(entrada)
        estado = "OK" if res == esperado else "FAIL"
        if estado == "FAIL":
            print(f"  {estado}  '{entrada}' -> '{res}' (esperado '{esperado}')")
            err += 1
        else:
            ok += 1
    print(f"\n{ok} OK  |  {err} FAIL  |  {len(_MAPA)} variantes en mapa")
