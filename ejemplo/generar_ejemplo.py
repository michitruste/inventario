"""
Genera datos de ejemplo INVENTADOS con la misma forma que los reales:
  ejemplo/catalogo_medidas.xlsx   -> medidas + fotos (flotantes, en celda, y por nombre de archivo)
  ejemplo/fotos_pc/               -> fotos sueltas "guardadas en la PC"
  ejemplo/seguimiento.xlsx        -> Excel de seguimiento con hojas y status
  ejemplo/hoja_1003.txt           -> lo que queda al copiar/pegar una hoja nueva desde la web
  ejemplo/web_simulada/           -> página falsa para probar Playwright
Uso:  python ejemplo/generar_ejemplo.py
"""
import io
import json
import random
import re
import shutil
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.datavalidation import DataValidation
from PIL import Image, ImageDraw

AQUI = Path(__file__).resolve().parent
random.seed(7)


def foto(texto, color):
    img = Image.new("RGB", (160, 120), color)
    d = ImageDraw.Draw(img)
    d.rectangle([20, 25, 140, 95], outline="white", width=3)
    d.text((30, 50), texto, fill="white")
    b = io.BytesIO(); img.save(b, "PNG")
    return b.getvalue()


# ------------------------------------------------------------------ catálogo
# (descripción, height, length, breadth, weight, tipo_foto)
CATALOGO = [
    ("HVAC | ABCD-K7TR-001", 30, 45, 20, 3.2, "flotante"),
    ("HVAC | ABCD-M2PX-004", 12, 20, 10, 0.8, "en_celda"),
    ("SRV | WXYZ-Q9ZO-010", 8, 44, 70, 12.5, "archivo"),
    ("ABCD-R4TS", 15, 15, 15, 1.1, "flotante"),        # mismo código, distinto prefijo:
    ("EFGH-R4TS", 25, 25, 25, 2.0, "en_celda"),        # se distinguen por prefijo
    ("HVAC", 40, 60, 30, 5.0, "archivo"),              # genérico para cualquier HVAC
    ("Adaptador ethernet", 2, 8, 3, 0.1, "en_celda"),
    ("Fuente de poder 500W", 9, 15, 14, 1.9, "flotante"),
    ("NET | JKLM-T8VW-002", 4, 30, 20, 1.4, None),     # registrado pero sin foto -> missing
    ("PWR | NOPQ-B3CY-007", None, None, None, None, "archivo"),  # foto pero sin medidas
]
COLORES = ["#c0392b", "#2980b9", "#27ae60", "#8e44ad", "#d35400", "#16a085", "#2c3e50", "#7f8c8d",
           "#b9770e", "#1f618d"]

fotos_pc = AQUI / "fotos_pc"
shutil.rmtree(fotos_pc, ignore_errors=True); fotos_pc.mkdir()

wb = openpyxl.Workbook()
ws = wb.active; ws.title = "Medidas"
ws.append(["Registro de medidas y fotos"])          # título arriba, como en muchos Excel reales
ws.append([])
ws.append(["#", "Descripción", "Height", "Lenght", "Breadth", "Weight", "Imagen", "Proveedor", "Comentarios"])
en_celda = []   # (celda, bytes)
for i, (desc, h, l, b, w, tf) in enumerate(CATALOGO, start=1):
    r = ws.max_row + 1
    ws.append([i, desc, h, l, b, w, None, random.choice(["Cliente A", "Cliente B"]), ""])
    ws.row_dimensions[r].height = 70
    datos = foto(desc.split("|")[-1].strip()[:18], COLORES[i - 1])
    if tf == "flotante":
        img = XLImage(io.BytesIO(datos)); img.width, img.height = 100, 75
        ws.add_image(img, f"G{r}")
    elif tf == "en_celda":
        ws.cell(r, 7, 999000 + len(en_celda))   # marcador, se convierte abajo
        en_celda.append((f"G{r}", datos))
    elif tf == "archivo":
        nombre = re.sub(r"\W+", "_", desc.split("|")[-1].strip()) + ".png"
        (fotos_pc / nombre).write_bytes(datos)
        ws.cell(r, 7, nombre)
# fila duplicada (pasa en la vida real): mismo código, completa la foto que faltaba
ws.append([99, "NET | JKLM-T8VW-002", None, None, None, None, "JKLM_T8VW_002.png", "", "foto tomada después"])
(fotos_pc / "JKLM_T8VW_002.png").write_bytes(foto("T8VW", "#5d6d7e"))
ws.column_dimensions["B"].width = 28; ws.column_dimensions["G"].width = 16
ruta_cat = AQUI / "catalogo_medidas.xlsx"
wb.save(ruta_cat)


def meter_imagenes_en_celda(ruta: Path, imagenes):
    """Convierte los marcadores en imágenes 'Colocar en celda' de Excel 365 (richData)."""
    tmp = ruta.with_suffix(".tmp")
    with zipfile.ZipFile(ruta) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                s = data.decode()
                for k, (celda, _) in enumerate(imagenes):
                    s = re.sub(rf'<c r="{celda}"([^>]*)><v>{999000 + k}</v></c>',
                               f'<c r="{celda}" t="e" vm="{k + 1}"><v>#VALUE!</v></c>', s)
                data = s.encode()
            elif item.filename == "xl/_rels/workbook.xml.rels":
                s = data.decode().replace("</Relationships>",
                    '<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>'
                    '<Relationship Id="rIdRv1" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>'
                    '<Relationship Id="rIdRv2" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>'
                    '<Relationship Id="rIdRv3" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>'
                    "</Relationships>")
                data = s.encode()
            elif item.filename == "[Content_Types].xml":
                s = data.decode()
                if 'Extension="png"' not in s:
                    s = s.replace("<Override", '<Default Extension="png" ContentType="image/png"/><Override', 1)
                s = s.replace("</Types>",
                    '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>'
                    '<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>'
                    '<Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>'
                    '<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>'
                    "</Types>")
                data = s.encode()
            zout.writestr(item, data)
        n = len(imagenes)
        for k, (_, b) in enumerate(imagenes):
            zout.writestr(f"xl/media/encelda{k + 1}.png", b)
        zout.writestr("xl/richData/rdrichvaluestructure.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">'
            '<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s></rvStructures>')
        zout.writestr("xl/richData/rdrichvalue.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="{n}">'
            + "".join(f'<rv s="0"><v>{k}</v><v>5</v></rv>' for k in range(n)) + "</rvData>")
        zout.writestr("xl/richData/richValueRel.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            + "".join(f'<rel r:id="rId{k + 1}"/>' for k in range(n)) + "</richValueRels>")
        zout.writestr("xl/richData/_rels/richValueRel.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/encelda{k + 1}.png"/>' for k in range(n))
            + "</Relationships>")
        zout.writestr("xl/metadata.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">'
            '<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>'
            f'<futureMetadata name="XLRICHVALUE" count="{n}">'
            + "".join(f'<bk><extLst><ext uri="{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}"><xlrd:rvb i="{k}"/></ext></extLst></bk>' for k in range(n))
            + f'</futureMetadata><valueMetadata count="{n}">'
            + "".join(f'<bk><rc t="1" v="{k}"/></bk>' for k in range(n))
            + "</valueMetadata></metadata>")
    tmp.replace(ruta)


meter_imagenes_en_celda(ruta_cat, en_celda)

# ------------------------------------------------------------------ hojas (items)
def serial():
    return "".join(random.choice("ABCDEFGHJKLMNPQRSTUVWXYZ") for _ in range(5))

DESCS = ["HVAC | ABCD-K7TR-001", "HVAC | ABCD-K7TR-002", "HVAC | ABCD-M2PX-004", "HVAC | ZZZZ-H5GG-003",
         "SRV | WXYZ-Q9ZO-010", "SRV | WXYZ-Q9ZO-011", "ABCD-R4TS-001", "EFGH-R4TS-002",
         "Adaptador Ethernet", "adaptador  ethernet ", "Fuente de poder 500 W", "NET | JKLM-T8VW-002",
         "PWR | NOPQ-B3CY-007", "SRV | WXYZ-Q9Z0-012",   # cero en vez de O: error de captura
         "Cable HDMI 2m", "SRV | UVWX-P6LK-001"]

seg = openpyxl.Workbook(); s = seg.active; s.title = "Seguimiento"
s.append(["Hoja", "No. de serie", "Descripción", "Costo", "Status"])
estados = ["completed", "to upload", "missing", "regresado al cliente", "correction"]
for hoja in ("1001", "1002"):
    for _ in range(random.randint(6, 9)):
        d = random.choice(DESCS)
        s.append([int(hoja), serial(), d, round(random.uniform(40, 2500), 2),
                  random.choices(estados, [4, 2, 3, 1, 1])[0]])
s.append([None, None, None, None, None])            # filas vacías que se cuelan
dv = DataValidation(type="list", formula1='"' + ",".join(estados) + '"'); s.add_data_validation(dv)
dv.add(f"E2:E{s.max_row}")
seg.save(AQUI / "seguimiento.xlsx")

# ------------------------------------------------------------------ hoja nueva copiada de la web
filas = ["Serial\tDescription\tCost\tHeight\tLength\tBreadth\tWeight\tImage\tUpload", ""]
web_items = []
for k in range(10):
    d = random.choice(DESCS)
    sr = serial() if k != 3 else "AB0DE"            # serie con un 0 (subida mal)
    c = round(random.uniform(40, 2500), 2)
    web_items.append({"serial": sr, "descripcion": d, "costo": c})
    filas.append(f"{sr}\t{d}\t${c:,.2f}\t\t\t\t\t\tUpload image")
    if k % 4 == 1:
        filas.append("\t\t\t\t\t\t\t\t")            # filas basura
filas.append("")
(AQUI / "hoja_1003.txt").write_text("\n".join(filas), encoding="utf-8")

# ------------------------------------------------------------------ web simulada para Playwright
web = AQUI / "web_simulada"; web.mkdir(exist_ok=True)
(web / "items.json").write_text(json.dumps(web_items))
filas_html = "".join(
    f'<tr data-serial="{it["serial"]}"><td class="serial">{it["serial"]}</td><td class="desc">{it["descripcion"]}</td>'
    f'<td class="cost">${it["costo"]:,.2f}</td>'
    f'<td><input name="height"></td><td><input name="length"></td><td><input name="breadth"></td>'
    f'<td><input name="weight"></td><td><input type="file" name="image"></td>'
    f'<td><button class="save">Save</button><span class="ok"></span></td></tr>'
    for it in web_items)
(web / "hoja_1003.html").write_text(f"""<!doctype html><html><body>
<h1>Hoja 1003</h1>
<table id="items"><thead><tr><th>Serial</th><th>Description</th><th>Cost</th><th>Height</th><th>Length</th>
<th>Breadth</th><th>Weight</th><th>Image</th><th></th></tr></thead><tbody>{filas_html}
<tr><td></td><td></td><td></td></tr></tbody></table>
<script>
document.querySelectorAll('button.save').forEach(b => b.onclick = () => {{
  const tr = b.closest('tr');
  const vals = [...tr.querySelectorAll('input:not([type=file])')].map(i => i.value);
  const f = tr.querySelector('input[type=file]').files.length;
  tr.querySelector('.ok').textContent = (vals.every(v => v) && f) ? 'Saved' : 'Error';
}});
</script></body></html>""", encoding="utf-8")

print("Ejemplos generados en", AQUI)
