"""Cálculos para el dashboard (también usados por la consola)."""
from __future__ import annotations

import pandas as pd

import config
import db
from matching import confianza


def items_con_extras(con) -> pd.DataFrame:
    df = db.items_df(con)
    df["confianza"] = [("confirmado" if c else confianza(s)) for s, c in zip(df.match_score, df.confirmado)]
    df["clave"] = df.codigo.fillna("").where(df.codigo.fillna("") != "", df.nombre.fillna(""))
    df["clave"] = df.clave.where(df.clave != "", df.descripcion)
    return df


def kpis(df: pd.DataFrame) -> dict:
    comp = df[df.status == "completed"]
    activos = df[df.status != "regresado al cliente"]
    return {
        "items": len(df),
        "completed": len(comp),
        "pct_completed": len(comp) / len(activos) if len(activos) else 0,
        "costo_total": float(activos.costo.sum()),
        "costo_completed": float(comp.costo.sum()),
        "pct_costo": float(comp.costo.sum() / activos.costo.sum()) if activos.costo.sum() else 0,
        "to_upload": int((df.status == "to upload").sum()),
        "missing": int((df.status == "missing").sum()),
        "correction": int((df.status == "correction").sum()),
    }


def por_hoja(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("hoja")
    out = pd.DataFrame({
        "items": g.size(),
        "completed": g.apply(lambda x: (x.status == "completed").sum(), include_groups=False),
        "to upload": g.apply(lambda x: (x.status == "to upload").sum(), include_groups=False),
        "missing": g.apply(lambda x: (x.status == "missing").sum(), include_groups=False),
        "costo": g.costo.sum(),
        "costo completed": g.apply(lambda x: x[x.status == "completed"].costo.sum(), include_groups=False),
    })
    activos = g.apply(lambda x: (x.status != "regresado al cliente").sum(), include_groups=False)
    out["avance"] = (out.completed / activos.replace(0, pd.NA)).fillna(0).astype(float)
    return out.reset_index()


def motivo_missing(row) -> str:
    if pd.isna(row.catalogo_id):
        return "sin registro en catálogo"
    if not row.confirmado and (row.match_score or 0) < config.AUTO_ASIGNAR_DESDE:
        return "match por confirmar"
    falta = [k for k in ("height", "length", "breadth", "weight") if pd.isna(row[k])]
    if pd.isna(row.imagen) or not row.imagen:
        falta.append("foto")
    return "falta " + ", ".join(falta) if falta else "?"


def prioridades(df: pd.DataFrame, top=20) -> pd.DataFrame:
    """Qué registrar/confirmar primero: lo que desbloquea más ítems y costo."""
    m = df[df.status == "missing"].copy()
    if m.empty:
        return pd.DataFrame(columns=["clave", "motivo", "items", "costo", "ejemplo"])
    m["motivo"] = m.apply(motivo_missing, axis=1)
    return (m.groupby(["clave", "motivo"])
             .agg(items=("id", "count"), costo=("costo", "sum"), ejemplo=("descripcion", "first"))
             .reset_index().sort_values(["items", "costo"], ascending=False).head(top))


def avance_en_el_tiempo(con, df: pd.DataFrame) -> pd.DataFrame:
    """Acumulado de ítems y costo marcados completed por día (según el historial)."""
    h = db.historial_df(con)
    h = h[(h.campo == "status") & (h.despues == "completed")]
    if h.empty:
        return pd.DataFrame(columns=["fecha", "completed", "costo"])
    h = h.sort_values("fecha").drop_duplicates("item_id", keep="last")
    h = h[["item_id", "fecha"]].merge(df[["id", "costo", "status"]].rename(columns={"id": "item_id"}), on="item_id")
    h = h[h.status == "completed"]
    h["fecha"] = pd.to_datetime(h.fecha).dt.date
    d = h.groupby("fecha").agg(completed=("item_id", "count"), costo=("costo", "sum")).cumsum()
    return d.reset_index()
