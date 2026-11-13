"""
App para ver, editar y hacer match del inventario.
Ejecutar:   streamlit run app.py
Para que otra persona de la red la abra:  streamlit run app.py --server.address 0.0.0.0
(y entra desde su navegador a http://IP_DE_TU_PC:8501)
"""
import base64
import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import config
import db
import excel_import
import exportar
import matching
import reportes
from parsing import parse_descripcion, parse_tabla_pegada

st.set_page_config(page_title="Inventario", page_icon="📦", layout="wide")
db.init()


# ------------------------------------------------------------------ utilidades
@st.cache_data(show_spinner=False)
def miniatura(nombre: str, mtime: float, lado: int = 96) -> str | None:
    """Data URI de la foto para mostrarla dentro de las tablas."""
    from PIL import Image
    try:
        im = Image.open(config.FOTOS_DIR / nombre)
        im.thumbnail((lado, lado))
        b = io.BytesIO(); im.convert("RGB").save(b, "JPEG", quality=80)
        return "data:image/jpeg;base64," + base64.b64encode(b.getvalue()).decode()
    except Exception:
        return None


def thumb(nombre):
    if not isinstance(nombre, str) or not nombre:
        return None
    p = config.FOTOS_DIR / nombre
    return miniatura(nombre, p.stat().st_mtime) if p.exists() else None


def quien() -> str:
    return st.session_state.get("usuario") or db.usuario()


def guardar_upload(archivo) -> Path:
    tmp = Path(tempfile.mkdtemp()) / archivo.name
    tmp.write_bytes(archivo.getvalue())
    return tmp


def txt(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)


def num(v):
    return None if v is None or (isinstance(v, float) and pd.isna(v)) or v == 0 else float(v)


ICONO = {"alta": "🟢 alta", "confirmado": "✅ confirmado", "media": "🟡 media", "baja": "🔴 baja",
         "sin match": "⚪ sin match"}

# ------------------------------------------------------------------ barra lateral
st.sidebar.title("📦 Inventario")
st.sidebar.text_input("Tu nombre (para el historial)", key="usuario", value=db.usuario())
pagina = st.sidebar.radio("Ir a", ["📥 Cargar hoja", "🔎 Revisar ítems", "📚 Catálogo", "📊 Dashboard",
                                    "⚙️ Importar / Exportar"])


# ================================================================== CARGAR HOJA
def pagina_cargar():
    st.header("📥 Cargar hoja nueva")
    st.caption("En la web: selecciona la tabla de la hoja, Ctrl+C, y pégala aquí. "
               "Las filas vacías y los encabezados se limpian solos.")
    c1, c2 = st.columns([1, 4])
    numero = c1.text_input("Número de hoja")
    texto = c2.text_area("Tabla pegada", height=200)
    if not texto:
        return
    filas = parse_tabla_pegada(texto)
    if not filas:
        st.warning("No encontré filas con número de serie.")
        return
    df = pd.DataFrame(filas)
    st.write(f"**{len(df)} ítems** detectados")
    st.dataframe(df[["serial", "serial_aviso", "descripcion", "costo"]], hide_index=True, width="stretch")
    malos = df[df.serial_aviso != ""]
    if not malos.empty:
        st.warning("Series con formato raro (se guardan tal cual, solo es aviso): " +
                   ", ".join(f"{r.serial} → ¿{r.serial_sugerido}?" if r.serial_sugerido else r.serial
                             for r in malos.itertuples()))
    if st.button("Guardar hoja y buscar matches", type="primary", disabled=not numero):
        with db.conectar() as con:
            r = db.cargar_hoja(con, numero, filas)
            m = matching.aplicar(con)
        st.success(f"Hoja {numero}: {r['nuevos']} nuevos, {r['actualizados']} actualizados. "
                   f"Confianza → alta {m['alta']}, media {m['media']}, baja {m['baja']}, "
                   f"sin match {m['sin match']}. Pasaron a 'to upload': {m['a_to_upload']}.")


# ================================================================== REVISAR
def panel_match(item):
    """Candidatos para un ítem + asignar / crear registro nuevo."""
    st.markdown(f"**{item.serial}** · {item.descripcion} · hoja {item.hoja} · status `{item.status}`")
    with db.conectar() as con:
        cands = matching.Indice(con).candidatos(item.descripcion, top=5)
        cat = db.catalogo_df(con).set_index("id")
    if cands:
        cols = st.columns(len(cands))
        for col, c in zip(cols, cands):
            r = cat.loc[c["catalogo_id"]]
            with col:
                if isinstance(r.imagen, str) and (config.FOTOS_DIR / r.imagen).exists():
                    st.image(str(config.FOTOS_DIR / r.imagen), width=110)
                else:
                    st.caption("(sin foto)")
                st.markdown(f"**{r.etiqueta}**  \n{c['score']:.0f}% · {c['tipo']}")
                st.caption(f"H{r.height} L{r.length} B{r.breadth} · {r.weight} kg")
                if st.button("Usar este", key=f"usar_{item.id}_{c['catalogo_id']}"):
                    with db.conectar() as con:
                        matching.confirmar(con, int(item.id), int(c["catalogo_id"]), quien=quien())
                    st.rerun()
    else:
        st.info("Sin candidatos en el catálogo.")

    with st.expander("Elegir otro registro del catálogo"):
        opciones = {f"{r.etiqueta}  (#{i})": i for i, r in cat.iterrows()}
        elegido = st.selectbox("Registro", list(opciones), index=None, key=f"otro_{item.id}")
        if elegido and st.button("Asignar", key=f"asig_{item.id}"):
            with db.conectar() as con:
                matching.confirmar(con, int(item.id), int(opciones[elegido]), quien=quien())
            st.rerun()

    with st.expander("➕ Registrar medidas y foto nuevas para este código"):
        d = parse_descripcion(item.descripcion)
        sugerida = d.codigo or d.nombre or item.descripcion
        with st.form(f"nuevo_{item.id}"):
            etiqueta = st.text_input("Clave (código, prefijo-código o nombre)", value=sugerida)
            c = st.columns(4)
            h = c[0].number_input("Height", min_value=0.0, step=0.1)
            l = c[1].number_input("Length", min_value=0.0, step=0.1)
            b = c[2].number_input("Breadth", min_value=0.0, step=0.1)
            w = c[3].number_input("Weight", min_value=0.0, step=0.01)
            foto = st.file_uploader("Foto", type=["png", "jpg", "jpeg", "webp", "bmp"])
            if st.form_submit_button("Guardar y asignar", type="primary"):
                with db.conectar() as con:
                    img = excel_import.guardar_foto(foto.getvalue(), etiqueta, Path(foto.name).suffix.lower()) if foto else None
                    cid = db.crear_catalogo(con, etiqueta, item.descripcion, num(h), num(l), num(b), num(w), img)
                    for tipo, k in excel_import.claves_para_catalogo(etiqueta):
                        db.agregar_clave(con, tipo, k, cid, "manual")
                    matching.confirmar(con, int(item.id), cid, quien=quien())
                st.success("Guardado. Todos los ítems con esa clave se actualizaron.")
                st.rerun()


def pagina_revisar():
    st.header("🔎 Revisar ítems")
    with db.conectar() as con:
        df = reportes.items_con_extras(con)
    if df.empty:
        st.info("Todavía no hay ítems. Carga una hoja o importa tu Excel de seguimiento.")
        return
    c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
    hojas = c1.multiselect("Hoja", sorted(df.hoja.unique()))
    status = c2.multiselect("Status", config.STATUS)
    conf = c3.multiselect("Confianza", list(ICONO))
    buscar = c4.text_input("Buscar (serie, descripción, código)")
    f = df
    if hojas:
        f = f[f.hoja.isin(hojas)]
    if status:
        f = f[f.status.isin(status)]
    if conf:
        f = f[f.confianza.isin(conf)]
    if buscar:
        b = buscar.lower()
        f = f[f.serial.str.lower().str.contains(b) | f.descripcion.fillna("").str.lower().str.contains(b, regex=False)
              | f.match.fillna("").str.lower().str.contains(b, regex=False)]

    vista = f.assign(foto=f.imagen.map(thumb), confianza=f.confianza.map(ICONO))[
        ["id", "foto", "hoja", "serial", "descripcion", "costo", "status", "match", "confianza",
         "height", "length", "breadth", "weight", "serial_aviso", "notas"]]
    st.caption(f"{len(vista)} ítems · edita **status** y **notas** directo en la tabla y pulsa Guardar")
    editado = st.data_editor(
        vista, hide_index=True, width="stretch", height=420, key="tabla_items",
        disabled=[c for c in vista.columns if c not in ("status", "notas")],
        column_config={
            "id": st.column_config.NumberColumn("id", width="small"),
            "foto": st.column_config.ImageColumn("foto", width="small"),
            "status": st.column_config.SelectboxColumn("status", options=config.STATUS, required=True),
            "costo": st.column_config.NumberColumn("costo", format="$%.2f"),
            "descripcion": st.column_config.TextColumn("descripción", width="medium"),
        })
    if st.button("💾 Guardar cambios", type="primary"):
        cambios = 0
        orig = vista.set_index("id")
        with db.conectar() as con:
            for r in editado.itertuples():
                o = orig.loc[r.id]
                if r.status != o.status:
                    db.set_status(con, int(r.id), r.status, quien()); cambios += 1
                if txt(r.notas) != txt(o.notas):
                    db.set_notas(con, int(r.id), txt(r.notas)); cambios += 1
        st.success(f"{cambios} cambios guardados.")
        st.rerun()

    st.divider()
    st.subheader("Revisar match")
    pendientes = f[f.status.isin(["missing", "to upload"])].sort_values("match_score", na_position="first")
    if pendientes.empty:
        st.caption("Nada pendiente con estos filtros.")
        return
    etiquetas = {f"{r.serial} · {r.descripcion} · {ICONO.get(r.confianza, r.confianza)}": r.id
                 for r in pendientes.itertuples()}
    sel = st.selectbox("Ítem (los de menor confianza primero)", list(etiquetas))
    panel_match(df.set_index("id", drop=False).loc[etiquetas[sel]])


# ================================================================== CATÁLOGO
def pagina_catalogo():
    st.header("📚 Catálogo de medidas y fotos")
    with db.conectar() as con:
        cat = db.catalogo_df(con)
        usos = pd.read_sql_query("SELECT catalogo_id, COUNT(*) n FROM items GROUP BY catalogo_id", con)
    cat = cat.merge(usos, left_on="id", right_on="catalogo_id", how="left").fillna({"n": 0})
    vista = cat.assign(foto=cat.imagen.map(thumb))[
        ["id", "foto", "etiqueta", "height", "length", "breadth", "weight", "n", "claves", "notas"]]
    editado = st.data_editor(
        vista, hide_index=True, width="stretch", height=420, key="tabla_cat",
        disabled=["id", "foto", "n", "claves"],
        column_config={"foto": st.column_config.ImageColumn("foto", width="small"),
                       "n": st.column_config.NumberColumn("ítems que lo usan", format="%d")})
    if st.button("💾 Guardar cambios del catálogo", type="primary"):
        orig = vista.set_index("id")
        with db.conectar() as con:
            for r in editado.itertuples():
                o = orig.loc[r.id]
                camb = {k: (txt(getattr(r, k)) if k in ("etiqueta", "notas") else num(getattr(r, k)))
                        for k in ("etiqueta", "height", "length", "breadth", "weight", "notas")
                        if txt(getattr(r, k)) != txt(o[k])}
                if camb:
                    db.actualizar_catalogo(con, int(r.id), **camb)
            m = matching.aplicar(con)
        st.success(f"Guardado. Pasaron a 'to upload': {m['a_to_upload']}")
        st.rerun()

    c1, c2 = st.columns(2)
    with c1.form("cambiar_foto"):
        st.subheader("Cambiar / agregar foto")
        opciones = {f"{r.etiqueta} (#{r.id})": r.id for r in cat.itertuples()}
        reg = st.selectbox("Registro", list(opciones))
        foto = st.file_uploader("Foto nueva", type=["png", "jpg", "jpeg", "webp", "bmp"])
        marcar = st.checkbox("Marcar como 'correction' los ítems completed que usan este registro")
        if st.form_submit_button("Guardar foto") and foto and reg:
            cid = opciones[reg]
            with db.conectar() as con:
                nombre = excel_import.guardar_foto(foto.getvalue(), reg.split(" (#")[0], Path(foto.name).suffix.lower())
                db.actualizar_catalogo(con, int(cid), imagen=nombre)
                if marcar:
                    for r in con.execute("SELECT id FROM items WHERE catalogo_id=? AND status='completed'", (cid,)).fetchall():
                        db.set_status(con, r["id"], "correction", quien())
                matching.aplicar(con)
            st.success("Foto actualizada.")
            st.rerun()
    with c2.form("nuevo_registro"):
        st.subheader("Registro nuevo")
        clave = st.text_input("Clave: código (K7TR), prefijo-código (ABCD-K7TR), categoría (HVAC) o nombre")
        c = st.columns(4)
        h = c[0].number_input("Height", min_value=0.0, step=0.1)
        l = c[1].number_input("Length", min_value=0.0, step=0.1)
        b = c[2].number_input("Breadth", min_value=0.0, step=0.1)
        w = c[3].number_input("Weight", min_value=0.0, step=0.01)
        foto = st.file_uploader("Foto", type=["png", "jpg", "jpeg", "webp", "bmp"], key="foto_nueva")
        if st.form_submit_button("Crear") and clave:
            with db.conectar() as con:
                img = excel_import.guardar_foto(foto.getvalue(), clave, Path(foto.name).suffix.lower()) if foto else None
                cid = db.crear_catalogo(con, clave, clave, num(h), num(l), num(b), num(w), img)
                for tipo, k in excel_import.claves_para_catalogo(clave):
                    db.agregar_clave(con, tipo, k, cid, "manual")
                m = matching.aplicar(con)
            st.success(f"Creado. Pasaron a 'to upload': {m['a_to_upload']}")
            st.rerun()


# ================================================================== DASHBOARD
def pagina_dashboard():
    st.header("📊 Avance")
    with db.conectar() as con:
        df = reportes.items_con_extras(con)
        tiempo = reportes.avance_en_el_tiempo(con, df)
    if df.empty:
        st.info("Sin datos todavía.")
        return
    k = reportes.kpis(df)
    c = st.columns(5)
    c[0].metric("Ítems", k["items"])
    c[1].metric("Completed", f"{k['pct_completed']:.0%}", f"{k['completed']} ítems", delta_color="off")
    c[2].metric("Costo completed", f"${k['costo_completed']:,.0f}", f"{k['pct_costo']:.0%} de ${k['costo_total']:,.0f}",
                delta_color="off")
    c[3].metric("To upload", k["to_upload"])
    c[4].metric("Missing", k["missing"], f"{k['correction']} correction", delta_color="off")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Ítems por status")
        st.bar_chart(df.status.value_counts().reindex(config.STATUS).fillna(0), horizontal=True)
    with c2:
        st.subheader("Costo por status")
        st.bar_chart(df.groupby("status").costo.sum().reindex(config.STATUS).fillna(0), horizontal=True)

    st.subheader("Avance por hoja")
    st.dataframe(reportes.por_hoja(df), hide_index=True, width="stretch", column_config={
        "avance": st.column_config.ProgressColumn("avance", format="percent", min_value=0, max_value=1),
        "costo": st.column_config.NumberColumn(format="$%.0f"),
        "costo completed": st.column_config.NumberColumn(format="$%.0f")})

    if not tiempo.empty:
        st.subheader("Completed acumulado en el tiempo")
        st.line_chart(tiempo.set_index("fecha")[["completed"]])

    st.subheader("🎯 Qué registrar primero")
    st.caption("Ítems en 'missing' agrupados por código: registrar uno desbloquea todos los de su fila.")
    st.dataframe(reportes.prioridades(df), hide_index=True, width="stretch",
                 column_config={"costo": st.column_config.NumberColumn(format="$%.0f")})

    avisos = df[df.serial_aviso.fillna("") != ""]
    if not avisos.empty:
        st.subheader("⚠️ Series con formato raro")
        st.dataframe(avisos[["hoja", "serial", "serial_aviso", "descripcion"]], hide_index=True)


# ================================================================== IMPORTAR / EXPORTAR
def pagina_io():
    st.header("⚙️ Importar / Exportar")
    log = []
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Excel de medidas y fotos")
        f = st.file_uploader("Catálogo (.xlsx)", type=["xlsx", "xlsm"], key="up_cat")
        if f and st.button("Importar catálogo"):
            with st.spinner("Leyendo Excel y extrayendo fotos…"):
                excel_import.importar_catalogo(guardar_upload(f), log=log.append)
                with db.conectar() as con:
                    log.append(f"Match: {matching.aplicar(con)}")
            st.success("\n\n".join(log))
    with c2:
        st.subheader("Excel de seguimiento (hojas + status)")
        f = st.file_uploader("Seguimiento (.xlsx)", type=["xlsx", "xlsm"], key="up_seg")
        if f and st.button("Importar seguimiento"):
            excel_import.importar_seguimiento(guardar_upload(f), log=log.append)
            st.success("\n\n".join(log))

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Exportar a Excel")
        fotos = st.checkbox("Incluir miniaturas de las fotos", value=True)
        if st.button("Generar Excel"):
            ruta = exportar.exportar(Path(tempfile.mkdtemp()) / "inventario.xlsx", con_fotos=fotos)
            st.download_button("⬇️ Descargar", ruta.read_bytes(), file_name=ruta.name)
    with c2:
        st.subheader("Traer status editados en Excel")
        st.caption("Si alguien trabajó sobre el Excel exportado, súbelo aquí y se pasan sus status.")
        f = st.file_uploader("Excel exportado y editado", type=["xlsx"], key="up_back")
        if f and st.button("Actualizar status"):
            n = exportar.leer_status_desde_excel(guardar_upload(f), quien())
            st.success(f"{n} status actualizados.")

    st.divider()
    st.subheader("🤖 Web (Playwright)")
    st.code("python web_bot.py login\n"
            "python web_bot.py extraer 1003        # trae la hoja de la web sin copiar/pegar\n"
            "python web_bot.py subir 1003          # llena medidas y foto; tú das clic en guardar\n"
            "python web_bot.py subir 1003 --auto   # cuando ya confíes en él", language="bash")


{"📥 Cargar hoja": pagina_cargar, "🔎 Revisar ítems": pagina_revisar, "📚 Catálogo": pagina_catalogo,
 "📊 Dashboard": pagina_dashboard, "⚙️ Importar / Exportar": pagina_io}[pagina]()
