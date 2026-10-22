"""
Motor de verificación y asignación de solicitudes de traspaso.

Cada solicitud (origen -> destino, estadístico, cantidad) se cubre en este orden:
  1. Origen solicitado, sin bajar de su exhibición.
  2. Locales alternativos con sobrante (físico > exhibición), de mayor a menor sobrante.
     Nunca quedan bajo exhibición.
  3. Origen solicitado, aunque quede bajo exhibición (hasta agotar su físico).
  4. Locales extremos: otros locales aunque queden bajo exhibición, de mayor a menor
     stock físico (hasta agotar su físico).
  Lo que no se cubre queda como pendiente.

El stock se descuenta en tiempo real: lo que un local cede en una fila ya no está
disponible para las filas siguientes. Las primeras filas tienen prioridad.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

import pandas as pd

# ---------------------------------------------------------------- utilidades


def normalizar(texto) -> str:
    """Mayúsculas, sin tildes ni espacios extra (para comparar encabezados)."""
    s = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def codigo(valor):
    """Extrae el código numérico de '189-GUASMO...', '051404000_MEDIAS...', 189.0, etc."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int,)) and not isinstance(valor, bool):
        return int(valor)
    if isinstance(valor, float):
        return int(valor)
    m = re.match(r"\s*0*(\d+)", str(valor))
    return int(m.group(1)) if m else None


def _num(v) -> float:
    try:
        x = float(v)
        return 0.0 if pd.isna(x) else x
    except (TypeError, ValueError):
        return 0.0


def _fmt(n) -> int | float:
    return int(n) if float(n).is_integer() else round(n, 2)


# ------------------------------------------------------------- lectura base BI

COLUMNAS_BASE = {
    "local": ["# LOCAL", "LOCAL", "NRO LOCAL", "COD LOCAL"],
    "estadistico": ["ESTADISTICO", "CODIGO ESTADISTICO"],
    "descripcion": ["DESCRIPCION"],
    "supervisor": ["SUPERV.", "SUPERVISOR", "SUPERV"],
    "designacion": ["DESIGNACION", "NOMBRE LOCAL"],
    "fisico": ["FISICO", "STOCK FISICO"],
    "exhibicion": ["UNI EXHIBICION", "UNI. EXHIBICION", "EXHIBICION", "UND EXHIBICION"],
    "consumo": ["CONSUMOS ACU", "CONSUMO ACU", "CONSUMOS"],
}
OBLIGATORIAS_BASE = ["local", "estadistico", "fisico", "exhibicion"]


def _buscar_encabezado(df_raw: pd.DataFrame, claves: list[str], max_filas: int = 15) -> int | None:
    """Devuelve el índice de la fila que contiene todos los textos clave."""
    for i in range(min(max_filas, len(df_raw))):
        fila = " | ".join(normalizar(v) for v in df_raw.iloc[i].tolist() if pd.notna(v))
        if all(k in fila for k in claves):
            return i
    return None


def _mapear(columnas, aliases: dict[str, list[str]]) -> dict[str, str]:
    norm = {normalizar(c): c for c in columnas}
    out = {}
    for clave, opciones in aliases.items():
        for op in opciones:
            if op in norm:
                out[clave] = norm[op]
                break
    return out


def leer_base(archivo, hoja: str | int | None = None) -> pd.DataFrame:
    """Lee la base del BI (xlsx/xls/xlsb/csv) y devuelve columnas estándar."""
    nombre = getattr(archivo, "name", str(archivo)).lower()
    if nombre.endswith(".csv"):
        raw = pd.read_csv(archivo, header=None, dtype=object, sep=None, engine="python")
    else:
        engine = "pyxlsb" if nombre.endswith(".xlsb") else None
        raw = pd.read_excel(archivo, sheet_name=hoja if hoja is not None else 0, header=None, dtype=object, engine=engine)
    h = _buscar_encabezado(raw, ["ESTADISTICO", "FISICO"])
    if h is None:
        raise ValueError("No encontré los encabezados de la base (se esperan columnas como '# Local', 'ESTADISTICO', 'Físico', 'Uni Exhibición').")
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c) for c in raw.iloc[h].tolist()]
    m = _mapear(df.columns, COLUMNAS_BASE)
    faltan = [c for c in OBLIGATORIAS_BASE if c not in m]
    if faltan:
        raise ValueError(f"A la base le faltan columnas: {', '.join(faltan)}")
    out = pd.DataFrame({k: df[v].values for k, v in m.items()})
    for c in ["descripcion", "supervisor", "designacion"]:
        if c not in out:
            out[c] = ""
    if "consumo" not in out:
        out["consumo"] = 0
    out["local"] = out["local"].map(codigo)
    out["estadistico"] = out["estadistico"].map(codigo)
    out = out.dropna(subset=["local", "estadistico"])
    out["local"] = out["local"].astype(int)
    out["estadistico"] = out["estadistico"].astype(int)
    for c in ["fisico", "exhibicion", "consumo"]:
        out[c] = out[c].map(_num)
    for c in ["descripcion", "supervisor", "designacion"]:
        out[c] = out[c].fillna("").astype(str).str.strip()
    # si hay duplicados local+estadístico, se suman físico/consumo y se toma la mayor exhibición
    out = (out.groupby(["local", "estadistico"], as_index=False)
              .agg(fisico=("fisico", "sum"), exhibicion=("exhibicion", "max"), consumo=("consumo", "sum"),
                   descripcion=("descripcion", "first"), supervisor=("supervisor", "first"),
                   designacion=("designacion", "first")))
    return out


# ---------------------------------------------------------- lectura solicitudes

COLUMNAS_SOL = {
    "origen": ["LOC. ORIGEN", "LOCAL ORIGEN", "LOCAL DE ORIGEN", "ORIGEN", "LOC ORIGEN"],
    "destino": ["LOC. DESTINO", "LOCAL DESTINO", "LOCAL DE DESTINO", "DESTINO", "LOC DESTINO", "SUCURSAL DESTINO"],
    "estadistico": ["ESTADISTICO", "ESTADÍSTICO", "CODIGO", "ARTICULO"],
    "cantidad": ["CANTIDAD", "CANTIDAD SOLICITADA", "CANT", "UNIDADES"],
}


def leer_solicitudes(archivo, hoja: str | int | None = None) -> pd.DataFrame:
    nombre = getattr(archivo, "name", str(archivo)).lower()
    if nombre.endswith(".csv"):
        raw = pd.read_csv(archivo, header=None, dtype=object, sep=None, engine="python")
    else:
        raw = pd.read_excel(archivo, sheet_name=hoja if hoja is not None else 0, header=None, dtype=object)
    h = _buscar_encabezado(raw, ["ORIGEN", "ESTADISTICO", "CANT"], max_filas=20)
    if h is None:
        raise ValueError("No encontré los encabezados de las solicitudes (se esperan 'Loc. Origen', 'Loc. Destino', 'Estadístico', 'Cantidad').")
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c) for c in raw.iloc[h].tolist()]
    return normalizar_solicitudes(df)


def normalizar_solicitudes(df: pd.DataFrame) -> pd.DataFrame:
    m = _mapear(df.columns, COLUMNAS_SOL)
    faltan = [c for c in ["origen", "estadistico", "cantidad"] if c not in m]
    if faltan:
        raise ValueError(f"A las solicitudes les faltan columnas: {', '.join(faltan)}")
    out = pd.DataFrame({
        "origen_txt": df[m["origen"]].values,
        "destino_txt": df[m["destino"]].values if "destino" in m else None,
        "estadistico_txt": df[m["estadistico"]].values,
        "cantidad": df[m["cantidad"]].map(_num).values,
    })
    out["origen"] = out["origen_txt"].map(codigo)
    out["destino"] = out["destino_txt"].map(codigo)
    out["estadistico"] = out["estadistico_txt"].map(codigo)
    out = out[out["origen"].notna() & out["estadistico"].notna() & (out["cantidad"] > 0)].copy()
    out["origen"] = out["origen"].astype(int)
    out["estadistico"] = out["estadistico"].astype(int)
    out["destino"] = out["destino"].map(lambda x: int(x) if x is not None and not pd.isna(x) else None)
    out = out.reset_index(drop=True)
    out.insert(0, "fila", range(1, len(out) + 1))
    return out[["fila", "origen", "destino", "estadistico", "cantidad"]]


# ------------------------------------------------------------------ cálculo


@dataclass
class Opciones:
    excluir_origenes: bool = True      # un local que es origen del ítem en la solicitud no se usa como alternativo/extremo
    mismo_supervisor: bool = False     # alternativos/extremos solo del mismo supervisor del origen
    usar_extremos: bool = True         # paso 4
    origen_bajo_exhibicion: bool = True  # paso 3


@dataclass
class Resultado:
    detalle: pd.DataFrame
    movimientos: pd.DataFrame
    traspasos: pd.DataFrame
    reporte: pd.DataFrame
    pendientes: pd.DataFrame
    resumen: dict = field(default_factory=dict)


TIPO_ORIGEN = "ORIGEN SOLICITADO"
TIPO_ORIGEN_BAJO = "ORIGEN BAJO EXHIBICIÓN"
TIPO_ALT = "ALTERNATIVO (con sobrante)"
TIPO_EXT = "EXTREMO (bajo exhibición)"


def calcular(base: pd.DataFrame, sol: pd.DataFrame, op: Opciones | None = None) -> Resultado:
    op = op or Opciones()
    inv = {(r.local, r.estadistico): r for r in base.itertuples(index=False)}
    fis = {k: max(0.0, v.fisico) for k, v in inv.items()}          # stock disponible que se va descontando
    fis_ini = {k: v.fisico for k, v in inv.items()}
    exh = {k: v.exhibicion for k, v in inv.items()}
    locales_item: dict[int, list[int]] = {}
    for (l, g) in inv:
        locales_item.setdefault(g, []).append(l)
    desc_item = base.drop_duplicates("estadistico").set_index("estadistico")["descripcion"].to_dict()
    sup_local = base[base["supervisor"] != ""].drop_duplicates("local").set_index("local")["supervisor"].to_dict()
    nom_local = base[base["designacion"] != ""].drop_duplicates("local").set_index("local")["designacion"].to_dict()
    origenes_item: dict[int, set[int]] = {}
    for r in sol.itertuples(index=False):
        origenes_item.setdefault(r.estadistico, set()).add(r.origen)

    # repetidos
    k3 = sol.groupby(["origen", "destino", "estadistico"], dropna=False)["fila"].transform("count")
    k2 = sol.groupby(["origen", "estadistico"])["fila"].transform("count")
    repetido = ["REPETIDO" if a > 1 else ("MISMO ÍTEM Y ORIGEN" if b > 1 else "") for a, b in zip(k3, k2)]

    movs = []
    filas = []
    acumulado: dict[tuple, float] = {}

    for idx, r in enumerate(sol.itertuples(index=False)):
        o, d, g, q = r.origen, r.destino, r.estadistico, float(r.cantidad)
        ko = (o, g)
        en_base = ko in inv
        exh_o = exh.get(ko, 0.0)
        fis_o_ini = fis_ini.get(ko, 0.0)
        cons_o = inv[ko].consumo if en_base else 0.0
        acumulado[ko] = acumulado.get(ko, 0.0) + q
        # validación inicial (como si se enviara todo lo solicitado al origen)
        if not en_base:
            val = "SIN DATO EN BASE"
        elif fis_o_ini <= 0:
            val = "SIN STOCK EN ORIGEN"
        elif acumulado[ko] > fis_o_ini:
            val = "STOCK INSUFICIENTE"
        elif fis_o_ini - acumulado[ko] < exh_o:
            val = "DEJA BAJO EXHIBICIÓN"
        else:
            val = "OK"

        falta = q

        def tomar(local, cant, tipo, paso):
            nonlocal falta
            if cant <= 0:
                return 0.0
            k = (local, g)
            fis[k] = fis.get(k, 0.0) - cant
            falta -= cant
            movs.append({"fila": r.fila, "local_envia": local, "destino": d, "estadistico": g,
                         "cantidad": cant, "tipo": tipo, "paso": paso})
            return cant

        def candidatos():
            sup_o = sup_local.get(o)
            excl = origenes_item.get(g, set()) if op.excluir_origenes else {o}
            for l in locales_item.get(g, []):
                if l == o or l == d or l in excl:
                    continue
                if op.mismo_supervisor and sup_o and sup_local.get(l) != sup_o:
                    continue
                yield l

        # paso 1: origen sin bajar de exhibición
        p1 = tomar(o, min(falta, max(0.0, fis.get(ko, 0.0) - exh_o)), TIPO_ORIGEN, 1) if en_base else 0.0

        # paso 2: alternativos con sobrante (dinámico)
        alts = []
        if falta > 0:
            cands = sorted(((fis[(l, g)] - exh[(l, g)], inv[(l, g)].consumo, l) for l in candidatos()),
                           key=lambda t: (-t[0], t[1], t[2]))
            for sob, _, l in cands:
                if falta <= 0 or sob <= 0:
                    break
                c = tomar(l, min(falta, sob), TIPO_ALT, 2)
                alts.append((l, c))

        # paso 3: origen bajo exhibición
        p3 = 0.0
        if falta > 0 and en_base and op.origen_bajo_exhibicion:
            p3 = tomar(o, min(falta, max(0.0, fis.get(ko, 0.0))), TIPO_ORIGEN_BAJO, 3)

        # paso 4: extremos (otros locales bajo exhibición)
        exts = []
        if falta > 0 and op.usar_extremos:
            cands = sorted(((fis[(l, g)], inv[(l, g)].consumo, l) for l in candidatos()),
                           key=lambda t: (-t[0], t[1], t[2]))
            for fl, _, l in cands:
                if falta <= 0 or fl <= 0:
                    break
                c = tomar(l, min(falta, fl), TIPO_EXT, 4)
                exts.append((l, c))

        txt = lambda lst: ", ".join(f"{l} ({_fmt(c)})" for l, c in lst)
        filas.append({
            "Fila": r.fila, "Loc. Origen": o, "Loc. Destino": d, "Estadístico": g,
            "Descripción": desc_item.get(g, "NO EXISTE EN BASE"),
            "Supervisor Origen": sup_local.get(o, "-"),
            "Cantidad Solicitada": q,
            "Exhibición Origen": exh_o, "Físico Origen": fis_o_ini, "Consumo Acum. Origen": cons_o,
            "Validación": val, "Rotación Origen": "CON ROTACIÓN" if cons_o > 0 else "SIN ROTACIÓN",
            "Repetido": repetido[idx],
            "1. Origen sin bajar exhibición": p1,
            "2. Alternativos": txt(alts), "Und. Alternativos": sum(c for _, c in alts),
            "3. Origen bajo exhibición": p3,
            "4. Extremos": txt(exts), "Und. Extremos": sum(c for _, c in exts),
            "Total a Enviar Origen": p1 + p3,
            "Stock Origen tras esta fila": fis.get(ko, 0.0) if en_base else None,
            "Pendiente sin Cubrir": max(0.0, falta),
        })

    detalle = pd.DataFrame(filas)
    movimientos = pd.DataFrame(movs, columns=["fila", "local_envia", "destino", "estadistico", "cantidad", "tipo", "paso"])

    # ------------------------------------------------ traspasos consolidados
    if len(movimientos):
        tr = (movimientos.groupby(["local_envia", "destino", "estadistico"], dropna=False, sort=False)
              .agg(cantidad=("cantidad", "sum"),
                   tipo=("tipo", lambda s: " + ".join(dict.fromkeys(s))),
                   filas=("fila", lambda s: ", ".join(str(x) for x in dict.fromkeys(s))))
              .reset_index())
        envia_item = movimientos.groupby(["local_envia", "estadistico"])["cantidad"].sum().to_dict()
        rows = []
        for t in tr.itertuples(index=False):
            k = (t.local_envia, t.estadistico)
            fi, ex = fis_ini.get(k, 0.0), exh.get(k, 0.0)
            total = envia_item[k]
            final = fi - total
            causadas = max(0.0, ex - final) - max(0.0, ex - fi)
            rows.append({
                "Loc. Origen": t.local_envia, "Loc. Destino": t.destino, "Estadístico": t.estadistico,
                "Cantidad a Traspasar": t.cantidad,
                "Descripción": desc_item.get(t.estadistico, ""), "Tipo de Origen": t.tipo,
                "Supervisor Origen": sup_local.get(t.local_envia, "-"),
                "Exhibición Origen": ex, "Físico Origen": fi, "Total que Envía (ítem)": total,
                "Stock Final Origen": final, "¿Queda Bajo Exhibición?": "SÍ" if final < ex else "NO",
                "Und. Bajo Exhibición Causadas": causadas, "Filas Solicitud": t.filas,
            })
        traspasos = pd.DataFrame(rows)
    else:
        traspasos = pd.DataFrame(columns=["Loc. Origen", "Loc. Destino", "Estadístico", "Cantidad a Traspasar"])

    # ------------------------------------------------ reporte bajo exhibición (por local + ítem)
    rep = []
    if len(movimientos):
        for (l, g), grp in movimientos.groupby(["local_envia", "estadistico"]):
            fi, ex = fis_ini.get((l, g), 0.0), exh.get((l, g), 0.0)
            total = grp["cantidad"].sum()
            final = fi - total
            if final >= ex:
                continue
            por_dest = grp.groupby("destino", dropna=False)["cantidad"].sum()
            tipos = list(dict.fromkeys(grp["tipo"]))
            if TIPO_EXT in tipos:
                motivo = "SE TOMA DE OTRO LOCAL BAJO EXHIBICIÓN (el origen solicitado no tenía stock suficiente)"
            elif fi < ex and (TIPO_ORIGEN_BAJO not in tipos):
                motivo = "YA ESTABA BAJO EXHIBICIÓN"
            else:
                motivo = "ORIGEN SOLICITADO QUEDA BAJO EXHIBICIÓN"
            rep.append({
                "Local que Envía": l, "Nombre Local": nom_local.get(l, ""), "Supervisor": sup_local.get(l, "-"),
                "Estadístico": g, "Descripción": desc_item.get(g, ""),
                "Exhibición": ex, "Físico Inicial": fi, "Total Enviado": total, "Stock Final": final,
                "Faltante Exhibición": max(0.0, ex - final),
                "Und. Bajo Exhibición Causadas": max(0.0, ex - final) - max(0.0, ex - fi),
                "Destinos": ", ".join(f"{dd} ({_fmt(c)})" for dd, c in por_dest.items()),
                "Motivo": motivo,
            })
    reporte = pd.DataFrame(rep)

    pendientes = detalle[detalle["Pendiente sin Cubrir"] > 0][
        ["Fila", "Loc. Origen", "Loc. Destino", "Estadístico", "Descripción", "Cantidad Solicitada", "Pendiente sin Cubrir"]
    ].reset_index(drop=True) if len(detalle) else pd.DataFrame()

    def suma(paso):
        return float(movimientos.loc[movimientos["paso"] == paso, "cantidad"].sum()) if len(movimientos) else 0.0

    resumen = {
        "solicitudes": int(len(sol)),
        "und_solicitadas": float(sol["cantidad"].sum()) if len(sol) else 0.0,
        "paso1": suma(1), "paso2": suma(2), "paso3": suma(3), "paso4": suma(4),
        "pendiente": float(detalle["Pendiente sin Cubrir"].sum()) if len(detalle) else 0.0,
        "traspasos": int(len(traspasos)),
        "locales_bajo_exh": int(len(reporte)),
        "repetidos": int(sum(1 for x in repetido if x == "REPETIDO")),
        "validacion": detalle["Validación"].value_counts().to_dict() if len(detalle) else {},
    }
    return Resultado(detalle, movimientos, traspasos, reporte, pendientes, resumen)
