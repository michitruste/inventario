"""
Exporta la base de datos a Excel, con el mismo estilo de tu hoja de seguimiento:
lista desplegable de status, colores por status, foto en miniatura y medidas.
"""
from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import config
import db
from matching import confianza

COLOR_STATUS = {"completed": "C6EFCE", "to upload": "FFEB9C", "missing": "FFC7CE",
                "correction": "F4B183", "regresado al cliente": "D9D9D9"}


def _miniatura(ruta: Path, alto=60):
    try:
        from PIL import Image
        im = Image.open(ruta)
        im.thumbnail((alto * 2, alto))
        b = io.BytesIO(); im.convert("RGB").save(b, "PNG"); b.seek(0)
        x = XLImage(b); x.width, x.height = im.size
        return x
    except Exception:
        return None


def exportar(destino: Path | None = None, con_fotos: bool = True) -> Path:
    destino = Path(destino or config.BASE / f"inventario_{datetime.now():%Y%m%d_%H%M}.xlsx")
    with db.conectar() as con:
        items = db.items_df(con)
        cat = db.catalogo_df(con)
    items = items.astype(object).where(items.notna(), None)   # NaN -> celda vacía
    cat = cat.astype(object).where(cat.notna(), None)

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Items"
    cols = ["hoja", "serial", "descripcion", "costo", "status", "match", "confianza", "height", "length",
            "breadth", "weight", "foto", "serial_aviso", "notas", "id"]
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True)
    col_foto = cols.index("foto") + 1
    for _, r in items.iterrows():
        ws.append([r.hoja, r.serial, r.descripcion, r.costo, r.status, r.match, confianza(r.match_score),
                   r.height, r.length, r.breadth, r.weight, r.imagen, r.serial_aviso, r.notas, r.id])
        fila = ws.max_row
        if con_fotos and r.imagen:
            img = _miniatura(config.FOTOS_DIR / r.imagen)
            if img:
                ws.row_dimensions[fila].height = 48
                ws.add_image(img, f"{get_column_letter(col_foto)}{fila}")
    n = ws.max_row
    dv = DataValidation(type="list", formula1='"' + ",".join(config.STATUS) + '"', allow_blank=False)
    ws.add_data_validation(dv); dv.add(f"E2:E{max(n, 2)}")
    for st, color in COLOR_STATUS.items():
        ws.conditional_formatting.add(f"E2:E{max(n, 2)}",
                                      CellIsRule(operator="equal", formula=[f'"{st}"'],
                                                 fill=PatternFill("solid", fgColor=color)))
    for col, ancho in zip("ABCDEFGHIJKLMNO", [8, 9, 30, 11, 18, 22, 10, 8, 8, 8, 8, 18, 18, 20, 6]):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions

    wc = wb.create_sheet("Catalogo")
    ccols = ["id", "etiqueta", "descripcion", "height", "length", "breadth", "weight", "imagen", "claves"]
    wc.append(ccols)
    for c in wc[1]:
        c.font = Font(bold=True)
    for _, r in cat.iterrows():
        wc.append([r[c] for c in ccols])

    wr = wb.create_sheet("Resumen")
    wr.append(["hoja", "items", "completed", "% completed", "costo total", "costo completed"])
    for c in wr[1]:
        c.font = Font(bold=True)
    for hoja, g in items.groupby("hoja"):
        comp = g[g.status == "completed"]
        costo = lambda d: float(sum(x for x in d.costo if x is not None))
        wr.append([hoja, len(g), len(comp), round(len(comp) / len(g), 3), costo(g), costo(comp)])
    wb.save(destino)
    return destino


def leer_status_desde_excel(ruta: Path, quien=None) -> int:
    """
    Si alguien prefiere seguir trabajando en el Excel exportado: cambia los status ahí
    y con esto se pasan de vuelta a la base de datos (usa la columna 'id').
    """
    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    ws = wb["Items"]
    filas = list(ws.iter_rows(values_only=True))
    wb.close()   # read_only deja el archivo abierto; cerrarlo permite borrar la copia temporal
    enc = list(filas[0]); i_id, i_st, i_no = enc.index("id"), enc.index("status"), enc.index("notas")
    n = 0
    with db.conectar() as con:
        for f in filas[1:]:
            if f[i_id] is None or f[i_st] not in config.STATUS:
                continue
            actual = con.execute("SELECT status, notas FROM items WHERE id=?", (f[i_id],)).fetchone()
            if actual and actual["status"] != f[i_st]:
                db.set_status(con, f[i_id], f[i_st], quien); n += 1
            if actual and f[i_no] and f[i_no] != actual["notas"]:
                db.set_notas(con, f[i_id], f[i_no])
    return n
