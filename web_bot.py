"""
Automatización de la página web con Playwright.

  python web_bot.py login                 abre el navegador, inicias sesión TÚ, y se guarda la sesión
  python web_bot.py extraer 1003          lee la tabla de la hoja 1003 y la carga a la base (sin copiar/pegar)
  python web_bot.py subir 1003            llena medidas + foto de los 'to upload'; tú revisas y das Guardar
  python web_bot.py subir 1003 --auto     da Guardar solo y marca 'completed' si la web confirma
  opciones: --config otro.json  --limite 5  --headless

Primera vez:  pip install playwright  y luego  playwright install chromium
La contraseña nunca se guarda en el código: solo la sesión del navegador (datos/sesion.json).
Ese archivo da acceso a tu cuenta: no lo compartas ni lo subas a ningún lado.
"""
from __future__ import annotations

import argparse
import json
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

import config
import db
import matching
from parsing import parse_numero, validar_serial

SESION = config.BASE / "datos" / "sesion.json"
CFG_LOCAL = config.BASE / "web_config.json"               # tus URLs reales: NO va al repositorio
CFG_PLANTILLA = config.BASE / "web_config.plantilla.json"


def cargar_cfg(ruta) -> dict:
    ruta = Path(ruta)
    if ruta == CFG_LOCAL and not ruta.exists():
        ruta.write_text(CFG_PLANTILLA.read_text(encoding="utf-8"), encoding="utf-8")
        raise SystemExit(f"Se creó {ruta.name} a partir de la plantilla. "
                         "Edítalo con las URLs y selectores de tu página y vuelve a correr el comando.")
    texto = ruta.read_text(encoding="utf-8")
    texto = texto.replace("{carpeta}", config.BASE.as_uri())   # solo para la web simulada de ejemplo
    return json.loads(texto)


def nuevo_contexto(p, headless: bool):
    nav = p.chromium.launch(headless=headless, slow_mo=0 if headless else 100)
    ctx = nav.new_context(storage_state=str(SESION) if SESION.exists() else None)
    return nav, ctx


def fmt(v) -> str:
    return ("%g" % v) if v is not None else ""


# ------------------------------------------------------------------ login
def login(cfg):
    with sync_playwright() as p:
        nav, ctx = nuevo_contexto(p, headless=False)
        pg = ctx.new_page()
        pg.goto(cfg["url_login"])
        input("Inicia sesión en la ventana del navegador y luego presiona Enter aquí… ")
        SESION.parent.mkdir(parents=True, exist_ok=True)
        ctx.storage_state(path=str(SESION))
        nav.close()
    print("Sesión guardada en datos/sesion.json (no la compartas: da acceso a tu cuenta).")


# ------------------------------------------------------------------ extraer
def leer_tabla(pg, cfg) -> list[dict]:
    t = cfg["tabla"]
    pg.wait_for_selector(t["filas"], timeout=cfg.get("timeout_seg", 20) * 1000)
    filas = []
    for tr in pg.locator(t["filas"]).all():
        def celda(campo):
            sel = t.get(campo)
            if not sel:
                return ""
            loc = tr.locator(sel)
            if loc.count() == 0:
                return ""
            el = loc.first
            # si la celda tiene un input (medidas editables), leer su valor
            if el.evaluate("e => e.tagName") in ("INPUT", "TEXTAREA"):
                return el.input_value()
            inp = el.locator("input:not([type=file])")
            return inp.first.input_value() if inp.count() else el.inner_text()
        serial, aviso, sug = validar_serial(celda("serial"))
        if len(serial) < 3:
            continue                         # filas vacías / basura
        filas.append({"serial": serial, "serial_aviso": aviso, "serial_sugerido": sug,
                      "descripcion": celda("descripcion").strip(), "costo": parse_numero(celda("costo")),
                      **{f"web_{k}": parse_numero(celda(k)) for k in ("height", "length", "breadth", "weight")}})
    return filas


def extraer(cfg, hoja: str, headless=True):
    with sync_playwright() as p:
        nav, ctx = nuevo_contexto(p, headless)
        pg = ctx.new_page()
        pg.goto(cfg["url_hoja"].format(hoja=hoja))
        filas = leer_tabla(pg, cfg)
        nav.close()
    with db.conectar() as con:
        r = db.cargar_hoja(con, hoja, filas)
        # verificación: si la web ya tiene medidas, ese ítem está completo
        ya = 0
        for f in filas:
            if all(f[f"web_{k}"] for k in ("height", "length", "breadth", "weight")):
                it = con.execute("SELECT id, status FROM items WHERE hoja=? AND serial=?", (hoja, f["serial"])).fetchone()
                if it and it["status"] in ("missing", "to upload"):
                    db.set_status(con, it["id"], "completed", quien="web_bot"); ya += 1
        m = matching.aplicar(con)
    print(f"Hoja {hoja}: {len(filas)} filas leídas, {r['nuevos']} nuevas, {ya} ya estaban completas en la web.")
    print(f"Match: {m}")
    for f in filas:
        if f["serial_aviso"]:
            print(f"  ! {f['serial']}: {f['serial_aviso']} {f['serial_sugerido']}")


# ------------------------------------------------------------------ subir
def pendientes(hoja: str):
    with db.conectar() as con:
        return [dict(r) for r in con.execute("""
            SELECT i.id, i.serial, i.descripcion, c.height, c.length, c.breadth, c.weight, c.imagen
            FROM items i JOIN catalogo c ON c.id = i.catalogo_id
            WHERE i.hoja=? AND i.status='to upload' ORDER BY i.id""", (hoja,))]


def subir(cfg, hoja: str, auto=False, limite=None, headless=False):
    items = pendientes(hoja)[:limite]
    if not items:
        print("No hay ítems 'to upload' en esa hoja."); return
    s = cfg["subir"]
    url = cfg["url_hoja"].format(hoja=hoja)
    to = cfg.get("timeout_seg", 20) * 1000
    ok = fallos = 0
    with sync_playwright() as p:
        nav, ctx = nuevo_contexto(p, headless=headless and auto)
        pg = ctx.new_page()
        pg.set_default_timeout(to)
        pg.goto(url)
        for it in items:
            print(f"→ {it['serial']}  {it['descripcion']}")
            try:
                fila = pg.locator(cfg["tabla"]["filas"]).filter(has_text=it["serial"]).first
                fila.wait_for()
                zona = fila
                if s.get("modo") == "pagina":
                    fila.locator(s["abrir"]).first.click()
                    pg.wait_for_load_state()
                    zona = pg
                for campo in ("height", "length", "breadth", "weight"):
                    zona.locator(s[campo]).first.fill(fmt(it[campo]))
                zona.locator(s["imagen"]).first.set_input_files(str(config.FOTOS_DIR / it["imagen"]))

                if auto:
                    zona.locator(s["guardar"]).first.click()
                    zona.locator(s["exito"]).first.wait_for(timeout=to)
                    confirmado = True
                else:
                    r = input("   Revisa en el navegador y da Guardar. [Enter]=quedó bien  s=saltar  q=salir: ").strip().lower()
                    if r == "q":
                        break
                    confirmado = r != "s"
                if confirmado:
                    with db.conectar() as con:
                        db.set_status(con, it["id"], "completed", quien="web_bot")
                    ok += 1
                    print("   ✔ completed")
            except PWTimeout:
                fallos += 1
                print("   ✘ no se confirmó en la web (se deja en 'to upload')")
                pg.screenshot(path=str(config.BASE / "datos" / f"error_{it['serial']}.png"))
            if s.get("modo") == "pagina" and s.get("volver_a_la_hoja", True):
                pg.goto(url)
            time.sleep(cfg.get("pausa_entre_items_seg", 1))
        nav.close()
    print(f"Listo: {ok} completed, {fallos} con error.")


def main(argv=None):
    a = argparse.ArgumentParser()
    a.add_argument("accion", choices=["login", "extraer", "subir"])
    a.add_argument("hoja", nargs="?")
    a.add_argument("--config", default=str(CFG_LOCAL))
    a.add_argument("--auto", action="store_true")
    a.add_argument("--limite", type=int)
    a.add_argument("--headless", action="store_true")
    args = a.parse_args(argv)
    cfg = cargar_cfg(args.config)
    db.init()
    if args.accion == "login":
        login(cfg)
    elif not args.hoja:
        a.error("falta el número de hoja")
    elif args.accion == "extraer":
        extraer(cfg, args.hoja)
    else:
        subir(cfg, args.hoja, auto=args.auto, limite=args.limite, headless=args.headless)


if __name__ == "__main__":
    main()
