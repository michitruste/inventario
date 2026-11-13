"""
Base de datos SQLite: un solo archivo (datos/inventario.db).
Se puede abrir también con 'DB Browser for SQLite' si quieres verla por dentro.
"""
from __future__ import annotations

import getpass
import sqlite3
from contextlib import contextmanager
from datetime import datetime

import pandas as pd

import config
from parsing import parse_descripcion

ESQUEMA = """
CREATE TABLE IF NOT EXISTS catalogo (
    id          INTEGER PRIMARY KEY,
    etiqueta    TEXT NOT NULL,           -- nombre legible: 'K7TR', 'HVAC', 'Adaptador ethernet'
    descripcion TEXT,                    -- descripción de ejemplo de donde salió
    height REAL, length REAL, breadth REAL, weight REAL,
    imagen      TEXT,                    -- ruta relativa dentro de datos/fotos
    notas       TEXT,
    actualizado TEXT
);
-- Varias claves pueden apuntar al mismo registro del catálogo
CREATE TABLE IF NOT EXISTS claves (
    tipo        TEXT NOT NULL,           -- prefijo_codigo | codigo | nombre | categoria | aprendido
    clave       TEXT NOT NULL,
    catalogo_id INTEGER NOT NULL REFERENCES catalogo(id) ON DELETE CASCADE,
    origen      TEXT,                    -- excel | manual | aprendido
    PRIMARY KEY (tipo, clave)
);
CREATE TABLE IF NOT EXISTS hojas (
    numero      TEXT PRIMARY KEY,
    fecha_carga TEXT,
    notas       TEXT
);
CREATE TABLE IF NOT EXISTS items (
    id              INTEGER PRIMARY KEY,
    hoja            TEXT NOT NULL REFERENCES hojas(numero),
    serial          TEXT NOT NULL,
    serial_aviso    TEXT,
    serial_sugerido TEXT,
    descripcion     TEXT,
    costo           REAL,
    categoria TEXT, prefijo TEXT, codigo TEXT, nombre TEXT,
    web_height REAL, web_length REAL, web_breadth REAL, web_weight REAL,
    status          TEXT NOT NULL DEFAULT 'missing',
    catalogo_id     INTEGER REFERENCES catalogo(id),
    match_tipo      TEXT,
    match_score     REAL,
    confirmado      INTEGER DEFAULT 0,
    notas           TEXT,
    creado          TEXT,
    actualizado     TEXT,
    UNIQUE (hoja, serial)
);
CREATE TABLE IF NOT EXISTS historial (
    id      INTEGER PRIMARY KEY,
    item_id INTEGER,
    campo   TEXT,
    antes   TEXT,
    despues TEXT,
    fecha   TEXT,
    usuario TEXT
);
CREATE INDEX IF NOT EXISTS ix_items_status ON items(status);
CREATE INDEX IF NOT EXISTS ix_items_codigo ON items(codigo);
"""


def ahora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def usuario() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return "?"


@contextmanager
def conectar(ruta=None):
    ruta = ruta or config.DB_PATH
    ruta.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(ruta, timeout=30)   # timeout: tolera 2-3 personas a la vez
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    try:
        yield con
        con.commit()
    finally:
        con.close()


def init():
    with conectar() as con:
        con.executescript(ESQUEMA)


# ---------------------------------------------------------------- catálogo
def crear_catalogo(con, etiqueta, descripcion=None, height=None, length=None,
                   breadth=None, weight=None, imagen=None, notas=None) -> int:
    cur = con.execute(
        "INSERT INTO catalogo (etiqueta, descripcion, height, length, breadth, weight, imagen, notas, actualizado)"
        " VALUES (?,?,?,?,?,?,?,?,?)",
        (etiqueta, descripcion, height, length, breadth, weight, imagen, notas, ahora()),
    )
    return cur.lastrowid


def actualizar_catalogo(con, cat_id: int, **campos):
    permitidos = {"etiqueta", "descripcion", "height", "length", "breadth", "weight", "imagen", "notas"}
    campos = {k: v for k, v in campos.items() if k in permitidos}
    if not campos:
        return
    sets = ", ".join(f"{k}=?" for k in campos) + ", actualizado=?"
    con.execute(f"UPDATE catalogo SET {sets} WHERE id=?", (*campos.values(), ahora(), cat_id))


def agregar_clave(con, tipo: str, clave: str, cat_id: int, origen="manual", reemplazar=False):
    if not clave:
        return
    verbo = "INSERT OR REPLACE" if reemplazar else "INSERT OR IGNORE"
    con.execute(f"{verbo} INTO claves (tipo, clave, catalogo_id, origen) VALUES (?,?,?,?)",
                (tipo, clave, cat_id, origen))


def catalogo_df(con) -> pd.DataFrame:
    df = pd.read_sql_query("""
        SELECT c.*, GROUP_CONCAT(k.tipo || ':' || k.clave, '  ') AS claves
        FROM catalogo c LEFT JOIN claves k ON k.catalogo_id = c.id
        GROUP BY c.id ORDER BY c.etiqueta""", con)
    return df


def catalogo_completo(row) -> bool:
    """Un registro sirve para subir si tiene las 4 medidas y foto."""
    return all(row[k] not in (None, "") for k in ("height", "length", "breadth", "weight", "imagen"))


# ---------------------------------------------------------------- hojas e items
def cargar_hoja(con, numero: str, filas: list[dict]) -> dict:
    """Inserta los ítems de una hoja. Si un ítem ya existía, actualiza descripción/costo."""
    numero = str(numero).strip()
    con.execute("INSERT OR IGNORE INTO hojas (numero, fecha_carga) VALUES (?,?)", (numero, ahora()))
    nuevos = actualizados = 0
    for f in filas:
        d = parse_descripcion(f.get("descripcion"))
        web_med = [f.get(k) for k in ("web_height", "web_length", "web_breadth", "web_weight")]
        status = "completed" if (config.WEB_MEDIDAS_IMPLICA_COMPLETED and all(web_med)) else "missing"
        existe = con.execute("SELECT id, descripcion, costo FROM items WHERE hoja=? AND serial=?",
                             (numero, f["serial"])).fetchone()
        if existe:
            if (existe["descripcion"], existe["costo"]) != (f.get("descripcion"), f.get("costo")):
                con.execute("""UPDATE items SET descripcion=?, costo=?, categoria=?, prefijo=?, codigo=?,
                               nombre=?, actualizado=? WHERE id=?""",
                            (f.get("descripcion"), f.get("costo"), d.categoria, d.prefijo, d.codigo,
                             d.nombre, ahora(), existe["id"]))
                actualizados += 1
            continue
        con.execute("""INSERT INTO items (hoja, serial, serial_aviso, serial_sugerido, descripcion, costo,
                         categoria, prefijo, codigo, nombre, web_height, web_length, web_breadth, web_weight,
                         status, creado, actualizado)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (numero, f["serial"], f.get("serial_aviso"), f.get("serial_sugerido"),
                     f.get("descripcion"), f.get("costo"), d.categoria, d.prefijo, d.codigo, d.nombre,
                     *web_med, status, ahora(), ahora()))
        nuevos += 1
    return {"nuevos": nuevos, "actualizados": actualizados}


def registrar_cambio(con, item_id, campo, antes, despues, quien=None):
    if str(antes) == str(despues):
        return
    con.execute("INSERT INTO historial (item_id, campo, antes, despues, fecha, usuario) VALUES (?,?,?,?,?,?)",
                (item_id, campo, None if antes is None else str(antes),
                 None if despues is None else str(despues), ahora(), quien or usuario()))


def set_status(con, item_id: int, status: str, quien=None):
    if status not in config.STATUS:
        raise ValueError(f"Status no válido: {status}")
    antes = con.execute("SELECT status FROM items WHERE id=?", (item_id,)).fetchone()["status"]
    if antes != status:
        con.execute("UPDATE items SET status=?, actualizado=? WHERE id=?", (status, ahora(), item_id))
        registrar_cambio(con, item_id, "status", antes, status, quien)


def set_notas(con, item_id: int, notas: str):
    con.execute("UPDATE items SET notas=?, actualizado=? WHERE id=?", (notas, ahora(), item_id))


def items_df(con, where="1=1", params=()) -> pd.DataFrame:
    return pd.read_sql_query(f"""
        SELECT i.id, i.hoja, i.serial, i.serial_aviso, i.descripcion, i.costo, i.status,
               i.match_score, i.match_tipo, i.confirmado, i.catalogo_id,
               c.etiqueta AS match, c.height, c.length, c.breadth, c.weight, c.imagen,
               i.codigo, i.prefijo, i.categoria, i.nombre, i.notas, i.actualizado
        FROM items i LEFT JOIN catalogo c ON c.id = i.catalogo_id
        WHERE {where}
        ORDER BY i.hoja, i.id""", con, params=params)


def historial_df(con) -> pd.DataFrame:
    return pd.read_sql_query("SELECT * FROM historial ORDER BY fecha", con)
