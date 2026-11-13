"""
Motor de matching: encuentra qué registro del catálogo (medidas + foto)
le corresponde a cada ítem, con un puntaje de confianza 0-100.

Niveles (de más a menos seguro):
  100  aprendido      -> una persona ya confirmó esa misma descripción
  100  prefijo_codigo -> 'ABCD-K7TR' exacto
   97  codigo         -> 'K7TR' exacto
   95  nombre         -> 'adaptador ethernet' exacto
   88  codigo ~0/1    -> 'K7TR' vs 'K7TR' con 0/O o 1/I intercambiados
   80  categoria      -> solo 'HVAC' (genérico, conviene revisar)
 <=85  parecido       -> similitud de texto (errores de dedo, nombres parecidos)
"""
from __future__ import annotations

import difflib

import config
import db
from parsing import norm_texto, parse_descripcion, variantes_0_1

try:
    from rapidfuzz import fuzz

    def _sim(a, b):
        return fuzz.ratio(a, b)

    def _sim_texto(a, b):
        return max(fuzz.token_sort_ratio(a, b), fuzz.partial_ratio(a, b) * 0.9)
except ImportError:  # funciona igual, un poco más lento
    def _sim(a, b):
        return difflib.SequenceMatcher(None, a, b).ratio() * 100

    def _sim_texto(a, b):
        a2, b2 = " ".join(sorted(a.split())), " ".join(sorted(b.split()))
        return max(_sim(a, b), _sim(a2, b2))

PUNTAJE_EXACTO = {"aprendido": 100, "prefijo_codigo": 100, "codigo": 97, "nombre": 95, "categoria": 80}
PUNTAJE_VARIANTE = 88
TOPE_PARECIDO = 85


def confianza(score) -> str:
    if score is None:
        return "sin match"
    if score >= config.UMBRAL_ALTA:
        return "alta"
    if score >= config.UMBRAL_MEDIA:
        return "media"
    return "baja"


class Indice:
    """Carga todas las claves una vez para no consultar la BD por cada ítem."""

    def __init__(self, con):
        self.exactas: dict[tuple[str, str], int] = {}
        for r in con.execute("SELECT tipo, clave, catalogo_id FROM claves"):
            self.exactas[(r["tipo"], r["clave"])] = r["catalogo_id"]
        self.codigos = [(k, v) for (t, k), v in self.exactas.items() if t in ("codigo", "prefijo_codigo")]
        self.nombres = [(k, v) for (t, k), v in self.exactas.items() if t == "nombre"]

    def candidatos(self, descripcion: str, top: int = 5) -> list[dict]:
        d = parse_descripcion(descripcion)
        mejores: dict[int, dict] = {}

        def poner(cat_id, tipo, score, clave):
            if cat_id is None or score < config.UMBRAL_MINIMO:
                return
            if cat_id not in mejores or score > mejores[cat_id]["score"]:
                mejores[cat_id] = {"catalogo_id": cat_id, "tipo": tipo, "score": round(score, 1), "clave": clave}

        # 1. ya confirmado antes por una persona
        clave_ap = norm_texto(d.original)
        poner(self.exactas.get(("aprendido", clave_ap)), "aprendido", 100, clave_ap)

        # 2. claves exactas
        for tipo, clave in d.claves:
            poner(self.exactas.get((tipo, clave)), tipo, PUNTAJE_EXACTO[tipo], clave)
            # el código puede estar registrado solo como 'nombre' o al revés
            if tipo == "codigo":
                poner(self.exactas.get(("nombre", norm_texto(clave))), "codigo", 93, clave)

        # 3. errores de captura 0/O 1/I en el código
        for tipo, clave in d.claves:
            if tipo in ("codigo", "prefijo_codigo"):
                for v in variantes_0_1(clave) - {clave}:
                    poner(self.exactas.get((tipo, v)), f"{tipo} ~0/1", PUNTAJE_VARIANTE, v)

        # 4. parecido (difuso)
        if d.codigo:
            buscado = f"{d.prefijo}-{d.codigo}" if d.prefijo else d.codigo
            for clave, cat_id in self.codigos:
                ref = d.codigo if "-" not in clave else buscado
                s = _sim(ref, clave)
                poner(cat_id, "parecido", min(s * 0.85, TOPE_PARECIDO), clave)
        if d.nombre:
            for clave, cat_id in self.nombres:
                s = _sim_texto(d.nombre, clave)
                poner(cat_id, "parecido", min(s * 0.9, TOPE_PARECIDO), clave)

        return sorted(mejores.values(), key=lambda x: -x["score"])[:top]


def aplicar(con, ids: list[int] | None = None) -> dict:
    """
    Recalcula el match de los ítems NO confirmados por una persona y ajusta el status
    (solo entre 'missing' y 'to upload'; nunca toca completed / correction / regresado).
    """
    idx = Indice(con)
    completos = {r["id"]: db.catalogo_completo(r) for r in con.execute("SELECT * FROM catalogo")}
    q = "SELECT id, descripcion, status, catalogo_id, confirmado FROM items"
    if ids:
        q += f" WHERE id IN ({','.join('?' * len(ids))})"
    stats = {"alta": 0, "media": 0, "baja": 0, "sin match": 0, "a_to_upload": 0}

    for it in con.execute(q, ids or ()).fetchall():
        if it["confirmado"]:
            cat_id, score = it["catalogo_id"], 100
        else:
            cands = idx.candidatos(it["descripcion"], top=1)
            if cands:
                c = cands[0]
                cat_id, score = c["catalogo_id"], c["score"]
                con.execute("UPDATE items SET catalogo_id=?, match_tipo=?, match_score=? WHERE id=?",
                            (cat_id, c["tipo"], score, it["id"]))
            else:
                cat_id, score = None, None
                con.execute("UPDATE items SET catalogo_id=NULL, match_tipo=NULL, match_score=NULL WHERE id=?",
                            (it["id"],))
        stats[confianza(score)] += 1

        if it["status"] in config.STATUS_MANUALES:
            continue
        listo = cat_id is not None and score >= config.AUTO_ASIGNAR_DESDE and completos.get(cat_id, False)
        nuevo = "to upload" if listo else "missing"
        if nuevo != it["status"]:
            db.set_status(con, it["id"], nuevo, quien="matching")
            stats["a_to_upload"] += nuevo == "to upload"
    return stats


def confirmar(con, item_id: int, cat_id: int, aprender: bool = True, quien=None):
    """Una persona dice 'este ítem es este registro del catálogo'. El sistema lo recuerda."""
    it = con.execute("SELECT * FROM items WHERE id=?", (item_id,)).fetchone()
    db.registrar_cambio(con, item_id, "catalogo_id", it["catalogo_id"], cat_id, quien)
    con.execute("UPDATE items SET catalogo_id=?, match_tipo='manual', match_score=100, confirmado=1, "
                "actualizado=? WHERE id=?", (cat_id, db.ahora(), item_id))
    if aprender:
        d = parse_descripcion(it["descripcion"])
        db.agregar_clave(con, "aprendido", norm_texto(d.original), cat_id, "aprendido", reemplazar=True)
        # generaliza: todos los ítems con este mismo código/nombre usarán este registro
        if d.codigo:
            ya = con.execute("SELECT catalogo_id FROM claves WHERE tipo='codigo' AND clave=?",
                             (d.codigo,)).fetchone()
            if ya is None:
                db.agregar_clave(con, "codigo", d.codigo, cat_id, "aprendido")
            elif ya["catalogo_id"] != cat_id and d.prefijo:
                db.agregar_clave(con, "prefijo_codigo", f"{d.prefijo}-{d.codigo}", cat_id, "aprendido",
                                 reemplazar=True)
        elif d.nombre:
            db.agregar_clave(con, "nombre", d.nombre, cat_id, "aprendido")
    aplicar(con)   # los demás ítems parecidos se actualizan solos


def quitar_confirmacion(con, item_id: int):
    con.execute("UPDATE items SET confirmado=0 WHERE id=?", (item_id,))
    aplicar(con, [item_id])
