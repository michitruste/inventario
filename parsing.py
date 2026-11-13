"""
Limpieza y lectura de datos: descripciones, números de serie, costos y la
tabla que se copia/pega desde la web.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

import config

SERIAL_OK = re.compile(r"^[A-Z]{5}$")
TOKEN_CODIGO = re.compile(r"^[A-Z0-9]+$")


# ---------------------------------------------------------------- texto
def sin_acentos(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def norm_texto(s) -> str:
    """'  Adaptador   Ethernet ' -> 'adaptador ethernet'"""
    if s is None:
        return ""
    s = sin_acentos(str(s)).lower().strip()
    s = re.sub(r"[^a-z0-9|\- ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def norm_clave(s) -> str:
    """Claves de catálogo: MAYÚSCULAS, sin espacios ni acentos."""
    if s is None:
        return ""
    return re.sub(r"\s+", "", sin_acentos(str(s)).upper().strip())


def variantes_0_1(s: str) -> set[str]:
    """Variantes típicas de captura: 0<->O, 1<->I. Para tolerar errores."""
    tabla = [("0", "O"), ("O", "0"), ("1", "I"), ("I", "1")]
    out = {s}
    for a, b in tabla:
        out |= {x.replace(a, b) for x in list(out)}
    return out


# ---------------------------------------------------------------- descripción
@dataclass
class Descripcion:
    original: str
    categoria: str = ""          # 'HVAC'
    prefijo: str = ""            # 'ABCD'
    codigo: str = ""             # 'K7TR'
    nombre: str = ""             # 'adaptador ethernet' (cuando no hay formato de código)
    claves: list[tuple[str, str]] = field(default_factory=list)  # [(tipo, clave)] de más a menos específica


def parse_descripcion(texto) -> Descripcion:
    """
    'HVAC | ABCD-K7TR-001'  -> categoria=HVAC, prefijo=ABCD, codigo=K7TR
    'ABCD-K7TR'             -> prefijo=ABCD, codigo=K7TR
    'Adaptador ethernet'    -> nombre='adaptador ethernet'
    'RED | Adaptador ethernet' -> categoria=RED, nombre='adaptador ethernet'
    """
    original = "" if texto is None else str(texto).strip()
    d = Descripcion(original=original)
    if not original:
        return d

    partes = [p.strip() for p in original.split("|") if p.strip()]
    resto = partes[-1] if partes else ""
    if len(partes) > 1:
        d.categoria = norm_clave(partes[0])

    tokens = [norm_clave(t) for t in resto.split("-")]
    es_codigo = (
        len(tokens) >= 2
        and " " not in resto.strip()
        and all(t and TOKEN_CODIGO.match(t) for t in tokens)
    )
    if es_codigo:
        d.prefijo, d.codigo = tokens[0], tokens[1]
    else:
        d.nombre = norm_texto(resto)

    # Claves para buscar en el catálogo, de más específica a menos
    if d.prefijo and d.codigo:
        d.claves.append(("prefijo_codigo", f"{d.prefijo}-{d.codigo}"))
    if d.codigo:
        d.claves.append(("codigo", d.codigo))
    if d.nombre:
        d.claves.append(("nombre", d.nombre))
    if d.categoria:
        d.claves.append(("categoria", d.categoria))
    return d


# ---------------------------------------------------------------- serial
def validar_serial(s) -> tuple[str, str, str]:
    """
    Devuelve (serial_limpio, aviso, sugerencia).
    aviso: '' si está bien, 'posible 0/1 por O/I', 'formato inválido'
    """
    limpio = norm_clave(s)
    if SERIAL_OK.match(limpio):
        return limpio, "", ""
    sug = limpio.replace("0", "O").replace("1", "I")
    if SERIAL_OK.match(sug):
        return limpio, "posible 0/1 por O/I", sug
    return limpio, "formato inválido", ""


# ---------------------------------------------------------------- números
def parse_numero(v) -> float | None:
    """'$1,234.50' -> 1234.5 ; '1.234,50' -> 1234.5 ; '' -> None"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s or s in {"-", ".", ","}:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):     # 1.234,50
            s = s.replace(".", "").replace(",", ".")
        else:                               # 1,234.50
            s = s.replace(",", "")
    elif "," in s:
        ent, _, dec = s.rpartition(",")
        s = s.replace(",", ".") if len(dec) != 3 else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


# ---------------------------------------------------------------- encabezados
def mapear_encabezados(encabezados, mapa: dict[str, list[str]]) -> dict[str, int]:
    """Encuentra qué índice de columna corresponde a cada campo interno."""
    norm = [norm_texto(h) for h in encabezados]
    res = {}
    for campo, opciones in mapa.items():
        ops = [norm_texto(o) for o in opciones]
        for i, h in enumerate(norm):
            if h in ops:
                res[campo] = i
                break
        else:
            for i, h in enumerate(norm):               # coincidencia parcial
                if h and any(o in h for o in ops) and i not in res.values():
                    res[campo] = i
                    break
    return res


# ---------------------------------------------------------------- tabla pegada de la web
def parse_tabla_pegada(texto: str) -> list[dict]:
    """
    Recibe el texto que queda al copiar la tabla de la web (Ctrl+C) y pegarlo.
    Quita filas vacías, filas de encabezado repetidas y basura.
    """
    lineas = [l for l in texto.replace("\r", "").split("\n")]
    filas = [[c.strip() for c in l.split("\t")] for l in lineas]
    filas = [f for f in filas if any(c for c in f)]           # sin filas vacías
    if not filas:
        return []

    idx = None
    for f in filas[:5]:
        m = mapear_encabezados(f, config.WEB_COLUMNAS)
        if "serial" in m and "descripcion" in m:
            idx = m
            break
    if idx is None:
        idx = {c: i for i, c in enumerate(config.WEB_ORDEN_COLUMNAS)}

    out = []
    for f in filas:
        get = lambda campo: f[idx[campo]] if campo in idx and idx[campo] < len(f) else ""
        serial_txt = get("serial")
        # descartar encabezados repetidos y filas sin serial válido-ish
        if norm_texto(serial_txt) in [norm_texto(x) for x in config.WEB_COLUMNAS["serial"]]:
            continue
        serial, aviso, sug = validar_serial(serial_txt)
        if not serial or len(serial) < 3:
            continue
        out.append({
            "serial": serial,
            "serial_aviso": aviso,
            "serial_sugerido": sug,
            "descripcion": get("descripcion"),
            "costo": parse_numero(get("costo")),
            "web_height": parse_numero(get("height")),
            "web_length": parse_numero(get("length")),
            "web_breadth": parse_numero(get("breadth")),
            "web_weight": parse_numero(get("weight")),
        })
    return out
