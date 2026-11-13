"""
Comandos de consola (todo esto también se puede hacer desde la app):

  python cli.py init
  python cli.py importar-catalogo  ruta\\al\\excel_medidas.xlsx
  python cli.py importar-seguimiento ruta\\al\\excel_seguimiento.xlsx
  python cli.py cargar-hoja 1003 hoja_1003.txt      (texto copiado de la web)
  python cli.py cargar-hoja 1003 --portapapeles      (copias la tabla en la web y corres esto)
  python cli.py match
  python cli.py resumen
  python cli.py exportar [destino.xlsx]
  python cli.py leer-excel inventario_xxx.xlsx        (trae de vuelta status editados en Excel)
"""
import argparse
import sys
from pathlib import Path

import db
import excel_import
import exportar
import matching
from parsing import parse_tabla_pegada


def leer_portapapeles() -> str:
    import tkinter
    r = tkinter.Tk(); r.withdraw()
    try:
        return r.clipboard_get()
    finally:
        r.destroy()


def resumen():
    import pandas as pd
    import reportes
    pd.set_option("display.width", 200)
    with db.conectar() as con:
        df = reportes.items_con_extras(con)
    if df.empty:
        print("Sin ítems todavía."); return
    k = reportes.kpis(df)
    print(df.groupby("status").agg(items=("id", "count"), costo=("costo", "sum")).to_string())
    print(f"\nTotal: {k['items']} ítems | completed {k['pct_completed']:.0%} "
          f"| costo completed ${k['costo_completed']:,.2f} de ${k['costo_total']:,.2f} ({k['pct_costo']:.0%})")
    print("\nPor hoja:\n" + reportes.por_hoja(df).to_string(index=False))
    pr = reportes.prioridades(df, top=10)
    if not pr.empty:
        print("\nQué registrar/confirmar primero:\n" + pr.to_string(index=False))


def main(argv=None):
    p = argparse.ArgumentParser(description="Inventario")
    sp = p.add_subparsers(dest="cmd", required=True)
    sp.add_parser("init")
    a = sp.add_parser("importar-catalogo"); a.add_argument("xlsx"); a.add_argument("--hojas", nargs="*")
    a = sp.add_parser("importar-seguimiento"); a.add_argument("xlsx")
    a = sp.add_parser("cargar-hoja"); a.add_argument("numero"); a.add_argument("archivo", nargs="?")
    a.add_argument("--portapapeles", action="store_true")
    sp.add_parser("match"); sp.add_parser("resumen")
    a = sp.add_parser("exportar"); a.add_argument("destino", nargs="?")
    a = sp.add_parser("leer-excel"); a.add_argument("xlsx")
    args = p.parse_args(argv)

    db.init()
    if args.cmd == "importar-catalogo":
        excel_import.importar_catalogo(Path(args.xlsx), args.hojas)
        with db.conectar() as con:
            print("Match:", matching.aplicar(con))
    elif args.cmd == "importar-seguimiento":
        excel_import.importar_seguimiento(Path(args.xlsx))
    elif args.cmd == "cargar-hoja":
        texto = leer_portapapeles() if args.portapapeles else Path(args.archivo).read_text(encoding="utf-8")
        filas = parse_tabla_pegada(texto)
        avisos = [f for f in filas if f["serial_aviso"]]
        with db.conectar() as con:
            r = db.cargar_hoja(con, args.numero, filas)
            m = matching.aplicar(con)
        print(f"Hoja {args.numero}: {r['nuevos']} nuevos, {r['actualizados']} actualizados | match {m}")
        for f in avisos:
            print(f"  ! serie {f['serial']}: {f['serial_aviso']} {f['serial_sugerido']}")
    elif args.cmd == "match":
        with db.conectar() as con:
            print(matching.aplicar(con))
    elif args.cmd == "resumen":
        resumen()
    elif args.cmd == "exportar":
        print("Exportado:", exportar.exportar(args.destino))
    elif args.cmd == "leer-excel":
        print("Status actualizados:", exportar.leer_status_desde_excel(Path(args.xlsx)))
    elif args.cmd == "init":
        print("Base de datos lista.")


if __name__ == "__main__":
    sys.exit(main())
