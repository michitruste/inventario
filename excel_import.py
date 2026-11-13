"""
Migración desde tus Excel actuales:

  * importar_catalogo(): Excel de medidas/fotos/descripciones -> tabla catálogo
      - fotos flotantes sobre la hoja (Insertar > Imagen)
      - fotos "en la celda" (Excel 365: Colocar en celda / =IMAGEN)
      - nombre o ruta de archivo escrito en la celda (busca en carpetas de la PC)
  * importar_seguimiento(): Excel de hojas con serie, descripción, costo y status
"""
from __future__ import annotations

import hashlib
import posixpath
import re
import zipfile
from collections import defaultdict
from pathlib import Path

import openpyxl
from lxml import etree

import config
import db
from parsing import mapear_encabezados, norm_clave, norm_texto, parse_descripcion, parse_numero

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "rv": "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata",
    "rvr": "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel",
    "xlrd": "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata",
}
EXT_IMG = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff", ".emf", ".wmf"}


# ================================================================ imágenes EN la celda
def _leer_xml(z: zipfile.ZipFile, ruta: str):
    try:
        return etree.fromstring(z.read(ruta))
    except KeyError:
        return None


def _rels(z, ruta_rels: str) -> dict[str, str]:
    x = _leer_xml(z, ruta_rels)
    if x is None:
        return {}
    base = posixpath.dirname(posixpath.dirname(ruta_rels))  # carpeta del archivo dueño
    out = {}
    for r in x.findall("rel:Relationship", NS):
        t = r.get("Target")
        out[r.get("Id")] = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(base, t))
    return out


def _celda_a_rc(ref: str) -> tuple[int, int]:
    letras, num = re.match(r"([A-Z]+)(\d+)", ref).groups()
    col = 0
    for ch in letras:
        col = col * 26 + ord(ch) - 64
    return int(num), col


def imagenes_en_celda(xlsx: Path) -> dict[str, dict[tuple[int, int], tuple[bytes, str]]]:
    """
    Devuelve {hoja: {(fila, columna): (bytes, extensión)}} para las imágenes
    que Excel 365 guarda DENTRO de la celda (no las flotantes).
    Cadena interna: celda vm -> metadata.xml -> rdrichvalue.xml -> richValueRel.xml -> xl/media/...
    """
    out: dict = defaultdict(dict)
    with zipfile.ZipFile(xlsx) as z:
        nombres = set(z.namelist())
        rv_xml = _leer_xml(z, "xl/richData/rdrichvalue.xml")
        if rv_xml is None:
            return {}

        # posición de la clave _rvRel:LocalImageIdentifier en cada estructura
        pos_img = {}
        st = _leer_xml(z, "xl/richData/rdrichvaluestructure.xml")
        if st is not None:
            for i, s in enumerate(st.findall("rv:s", NS)):
                ks = [k.get("n") for k in s.findall("rv:k", NS)]
                if "_rvRel:LocalImageIdentifier" in ks:
                    pos_img[str(i)] = ks.index("_rvRel:LocalImageIdentifier")

        # richvalue index -> índice de relación
        rv_a_rel = []
        for rv in rv_xml.findall("rv:rv", NS):
            vals = [v.text for v in rv.findall("rv:v", NS)]
            p = pos_img.get(rv.get("s"), 0)
            rv_a_rel.append(int(vals[p]) if p < len(vals) and vals[p] and vals[p].isdigit() else None)

        # índice de relación -> archivo de imagen
        rel_ids = []
        for ruta in ("xl/richData/richValueRel.xml",):
            x = _leer_xml(z, ruta)
            if x is not None:
                rel_ids = [e.get(f"{{{NS['r']}}}id") for e in x]
        rel_map = _rels(z, "xl/richData/_rels/richValueRel.xml.rels")
        rel_a_archivo = [rel_map.get(r) for r in rel_ids]

        # vm (1-based) -> índice richvalue, usando metadata.xml
        vm_a_rv = {}
        meta = _leer_xml(z, "xl/metadata.xml")
        if meta is not None:
            fut = []
            for fm in meta.findall("m:futureMetadata", NS):
                if fm.get("name") == "XLRICHVALUE":
                    fut = [bk.find(".//xlrd:rvb", NS) for bk in fm.findall("m:bk", NS)]
            vm = meta.find("m:valueMetadata", NS)
            if vm is not None:
                for i, bk in enumerate(vm.findall("m:bk", NS), start=1):
                    rc = bk.find("m:rc", NS)
                    v = int(rc.get("v")) if rc is not None else i - 1
                    if fut and v < len(fut) and fut[v] is not None:
                        vm_a_rv[i] = int(fut[v].get("i"))
                    else:
                        vm_a_rv[i] = v

        # hojas: nombre -> xml
        wb = _leer_xml(z, "xl/workbook.xml")
        wb_rels = _rels(z, "xl/_rels/workbook.xml.rels")
        for sh in wb.find("m:sheets", NS):
            nombre = sh.get("name")
            ruta = wb_rels.get(sh.get(f"{{{NS['r']}}}id"))
            if not ruta or ruta not in nombres:
                continue
            x = _leer_xml(z, ruta)
            for c in x.iter(f"{{{NS['m']}}}c"):
                vmv = c.get("vm")
                if not vmv:
                    continue
                rv_i = vm_a_rv.get(int(vmv), int(vmv) - 1)
                if rv_i is None or rv_i >= len(rv_a_rel):
                    continue
                rel_i = rv_a_rel[rv_i]
                if rel_i is None or rel_i >= len(rel_a_archivo):
                    continue
                archivo = rel_a_archivo[rel_i]
                if archivo and archivo in nombres:
                    out[nombre][_celda_a_rc(c.get("r"))] = (z.read(archivo), Path(archivo).suffix.lower())
    return dict(out)


# ================================================================ imágenes flotantes
def imagenes_flotantes(ws) -> dict[int, list[tuple[bytes, str]]]:
    """{fila: [(bytes, ext)]} según la celda donde está la esquina superior izquierda."""
    out = defaultdict(list)
    for img in getattr(ws, "_images", []):
        ancla = getattr(img.anchor, "_from", None)
        if ancla is None:
            continue
        try:
            datos = img._data()
        except Exception:
            continue
        out[ancla.row + 1].append((datos, "." + (img.format or "png").lower()))
    return out


# ================================================================ guardar fotos
def _nombre_seguro(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", s)[:40] or "img"


def guardar_foto(datos: bytes, etiqueta: str, ext: str) -> str:
    """Guarda en datos/fotos con nombre <etiqueta>_<hash>.ext. Si ya existe idéntica, la reutiliza."""
    config.FOTOS_DIR.mkdir(parents=True, exist_ok=True)
    h = hashlib.sha1(datos).hexdigest()[:10]
    existente = next(config.FOTOS_DIR.glob(f"*_{h}.*"), None)
    if existente:
        return existente.name
    nombre = f"{_nombre_seguro(etiqueta)}_{h}{ext if ext.startswith('.') else '.' + ext}"
    (config.FOTOS_DIR / nombre).write_bytes(datos)
    return nombre


def _indice_fotos_pc(carpetas) -> dict[str, Path]:
    """{nombre_normalizado_sin_extensión: ruta} de las fotos sueltas en la PC."""
    idx = {}
    for carpeta in carpetas:
        carpeta = Path(carpeta)
        if not carpeta.exists():
            continue
        for p in carpeta.rglob("*"):
            if p.suffix.lower() in EXT_IMG:
                idx.setdefault(norm_clave(p.stem), p)
                idx.setdefault(p.name.lower(), p)
    return idx


# ================================================================ catálogo
def claves_para_catalogo(texto: str, clave_explicita: str | None = None) -> list[tuple[str, str]]:
    """Qué claves registrar para una fila del Excel de medidas."""
    base = clave_explicita or texto
    d = parse_descripcion(base)
    claves = []
    if d.codigo:
        if d.prefijo:
            claves.append(("prefijo_codigo", f"{d.prefijo}-{d.codigo}"))
        claves.append(("codigo", d.codigo))
    elif d.nombre:
        claves.append(("nombre", d.nombre))
        if " " not in d.nombre:                     # 'HVAC', 'K7TR' sueltos
            claves.append(("categoria", norm_clave(d.nombre)))
            claves.append(("codigo", norm_clave(d.nombre)))
    elif d.categoria:
        claves.append(("categoria", d.categoria))
    return claves


def _buscar_encabezado(ws, mapa, max_filas=15):
    for fila in ws.iter_rows(min_row=1, max_row=max_filas):
        valores = [c.value for c in fila]
        m = mapear_encabezados([v if v is not None else "" for v in valores], mapa)
        if "descripcion" in m or "clave" in m:
            return fila[0].row, {k: v + 1 for k, v in m.items()}   # columnas 1-based
    return None, {}


def importar_catalogo(xlsx: Path, hojas: list[str] | None = None, log=print) -> dict:
    xlsx = Path(xlsx)
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    en_celda = imagenes_en_celda(xlsx)
    fotos_pc = _indice_fotos_pc([xlsx.parent, *config.CARPETAS_FOTOS_EXTRA])
    stats = defaultdict(int)

    with db.conectar() as con:
        for ws in wb.worksheets:
            if hojas and ws.title not in hojas:
                continue
            fila_enc, col = _buscar_encabezado(ws, config.CATALOGO_COLUMNAS)
            if not fila_enc:
                log(f"[{ws.title}] sin encabezados reconocibles, se omite")
                continue
            flotantes = imagenes_flotantes(ws)
            celda_imgs = en_celda.get(ws.title, {})
            celda_por_fila = defaultdict(list)
            for (r, c), v in celda_imgs.items():
                celda_por_fila[r].append((c, v))

            for r in range(fila_enc + 1, ws.max_row + 1):
                val = lambda campo: ws.cell(r, col[campo]).value if campo in col else None
                desc = val("descripcion")
                clave = val("clave")
                if not desc and not clave:
                    continue
                claves = claves_para_catalogo(str(desc or ""), str(clave) if clave else None)
                if not claves:
                    continue
                etiqueta = claves[0][1] if claves[0][0] != "nombre" else str(desc or clave).strip()

                # ---- foto
                foto = None
                imgs = celda_por_fila.get(r)
                if imgs:
                    imgs.sort(key=lambda x: abs(x[0] - col.get("imagen", x[0])))
                    datos, ext = imgs[0][1]
                    foto = guardar_foto(datos, etiqueta, ext); stats["fotos_en_celda"] += 1
                elif flotantes.get(r):
                    datos, ext = flotantes[r][0]
                    foto = guardar_foto(datos, etiqueta, ext); stats["fotos_flotantes"] += 1
                else:
                    txt = val("imagen")
                    candidatos = []
                    if txt:
                        p = Path(str(txt).strip().strip('"'))
                        candidatos += [p, xlsx.parent / p.name]
                        candidatos.append(fotos_pc.get(p.name.lower()))
                        candidatos.append(fotos_pc.get(norm_clave(p.stem)))
                    for _, k in claves:
                        candidatos.append(fotos_pc.get(norm_clave(k)))
                    p = next((c for c in candidatos if c and Path(c).exists()), None)
                    if p:
                        foto = guardar_foto(Path(p).read_bytes(), etiqueta, Path(p).suffix.lower())
                        stats["fotos_de_carpeta"] += 1

                medidas = {k: parse_numero(val(k)) for k in ("height", "length", "breadth", "weight")}

                # ---- crear o completar registro existente con la misma clave principal
                tipo0, clave0 = claves[0]
                ya = con.execute("SELECT catalogo_id FROM claves WHERE tipo=? AND clave=?",
                                 (tipo0, clave0)).fetchone()
                if ya:
                    cat_id = ya["catalogo_id"]
                    act = con.execute("SELECT * FROM catalogo WHERE id=?", (cat_id,)).fetchone()
                    cambios = {k: v for k, v in medidas.items() if v is not None and act[k] is None}
                    if foto and not act["imagen"]:
                        cambios["imagen"] = foto
                    if cambios:
                        db.actualizar_catalogo(con, cat_id, **cambios)
                    stats["duplicados_fusionados"] += 1
                else:
                    cat_id = db.crear_catalogo(con, etiqueta, str(desc or clave), imagen=foto, **medidas)
                    stats["registros"] += 1
                for tipo, k in claves:
                    db.agregar_clave(con, tipo, k, cat_id, "excel")
            log(f"[{ws.title}] leída")
    stats = dict(stats)
    log(f"Catálogo importado: {stats}")
    return stats


# ================================================================ seguimiento (hojas + status)
def _norm_status(s) -> str | None:
    t = norm_texto(s)
    if not t:
        return None
    for st in config.STATUS:
        if t == norm_texto(st) or t.startswith(norm_texto(st)[:6]):
            return st
    if "regres" in t or "devuel" in t:
        return "regresado al cliente"
    if "upload" in t or "subir" in t:
        return "to upload"
    if "correc" in t:
        return "correction"
    return None


def importar_seguimiento(xlsx: Path, log=print) -> dict:
    """Trae tu Excel de seguimiento actual (para no perder los status que ya tienes)."""
    from parsing import validar_serial
    import matching

    wb = openpyxl.load_workbook(Path(xlsx), data_only=True, read_only=True)
    por_hoja: dict[str, list[dict]] = defaultdict(list)
    for ws in wb.worksheets:
        filas = list(ws.iter_rows(values_only=True))
        enc_i, m = None, {}
        for i, f in enumerate(filas[:15]):
            m = mapear_encabezados([v if v is not None else "" for v in f], config.SEGUIMIENTO_COLUMNAS)
            if "serial" in m and "descripcion" in m:
                enc_i = i
                break
        if enc_i is None:
            log(f"[{ws.title}] sin encabezados de seguimiento, se omite")
            continue
        hoja_actual = None
        for f in filas[enc_i + 1:]:
            g = lambda k: f[m[k]] if k in m and m[k] < len(f) else None
            if g("hoja") not in (None, ""):
                hoja_actual = str(g("hoja")).strip().removesuffix(".0")
            hoja = hoja_actual or ws.title          # si no hay columna hoja, usa el nombre de la pestaña
            serial, aviso, sug = validar_serial(g("serial"))
            if not serial:
                continue
            por_hoja[hoja].append({"serial": serial, "serial_aviso": aviso, "serial_sugerido": sug,
                                   "descripcion": g("descripcion"), "costo": parse_numero(g("costo")),
                                   "_status": _norm_status(g("status"))})
    wb.close()   # read_only deja el archivo abierto; cerrarlo permite borrar la copia temporal
    stats ={"hojas": 0, "items": 0, "con_status": 0}
    with db.conectar() as con:
        for hoja, filas in por_hoja.items():
            db.cargar_hoja(con, hoja, filas)
            stats["hojas"] += 1
            for f in filas:
                stats["items"] += 1
                if f["_status"]:
                    iid = con.execute("SELECT id FROM items WHERE hoja=? AND serial=?",
                                      (hoja, f["serial"])).fetchone()["id"]
                    db.set_status(con, iid, f["_status"], quien="importación")
                    stats["con_status"] += 1
        matching.aplicar(con)
    log(f"Seguimiento importado: {stats}")
    return stats
