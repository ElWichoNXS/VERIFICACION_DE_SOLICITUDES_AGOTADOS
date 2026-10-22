"""
app.py  -  Verificación de solicitudes de traspaso (Streamlit)
Ejecutar:  streamlit run app.py
"""
import io
from datetime import datetime

import altair as alt
import pandas as pd
import streamlit as st

import motor
from exportar import generar_excel
from portal import menu_apps

st.set_page_config(page_title="Verificación de Traspasos", page_icon="📦", layout="wide")
menu_apps("Verificación de traspasos")

# ----------------------------------------------------------------------
# Estilo (mismo lenguaje visual que la app Min/Max): rojo y blanco como base;
# texto y fondos heredan del tema claro/oscuro.
# ----------------------------------------------------------------------
ROJO, ROJO_OSC = "#E30613", "#A30410"
C_OK, C_ALT, C_BAJO, C_EXT, C_PEND = "#1FA391", "#4B86E6", "#D39A2E", "#E5414B", "#8C8C8C"
BORDE = "color-mix(in srgb, currentColor 16%, transparent)"
TARJETA = "color-mix(in srgb, currentColor 4%, transparent)"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, .stApp, h1, h2, h3, h4, p, label, li, button, input, textarea {{ font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif; }}
.block-container {{ padding-top: 1rem; max-width: 1480px; }}
header[data-testid="stHeader"] {{ background: transparent; }}

.hero {{ background: linear-gradient(115deg, {ROJO} 0%, {ROJO_OSC} 100%); border-radius: 16px;
    padding: 22px 30px; display: flex; justify-content: space-between; align-items: center; gap: 24px;
    box-shadow: 0 8px 24px rgba(227,6,19,.25); flex-wrap: wrap; }}
.hero h1 {{ color:#fff !important; margin:0; font-size:1.75rem; font-weight:800; letter-spacing:-.4px; padding:0; }}
.hero p {{ color:#ffe3e5 !important; margin:4px 0 0 0; font-size:.95rem; }}
.pasos {{ display:flex; gap:10px; flex-wrap:wrap; }}
.paso {{ display:flex; align-items:center; gap:8px; background:rgba(255,255,255,.14); color:#fff;
    padding:7px 14px 7px 8px; border-radius:30px; font-size:.82rem; font-weight:600; }}
.paso b {{ width:22px; height:22px; border-radius:50%; background:rgba(255,255,255,.28); display:flex;
    align-items:center; justify-content:center; font-size:.78rem; }}
.paso.ok {{ background:#fff; color:{ROJO}; }}
.paso.ok b {{ background:{ROJO}; color:#fff; }}

.kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:14px; margin:6px 0 4px 0; }}
.kpi {{ background:{TARJETA}; border:1px solid {BORDE}; border-radius:14px; padding:16px 18px; position:relative; overflow:hidden; }}
.kpi::before {{ content:""; position:absolute; left:0; top:0; bottom:0; width:5px; background:var(--ac, {ROJO}); }}
.kpi .l {{ font-size:.8rem; font-weight:600; opacity:.7; }}
.kpi .v {{ font-size:1.9rem; font-weight:800; letter-spacing:-.6px; line-height:1.15; margin-top:2px; }}
.kpi .s {{ font-size:.78rem; opacity:.62; margin-top:2px; }}

.seccion {{ margin: 22px 0 8px 0; font-size:1.12rem; font-weight:800; letter-spacing:-.2px; }}
.seccion small {{ display:block; font-weight:400; opacity:.62; font-size:.84rem; letter-spacing:0; margin-top:1px; }}
.nota {{ border-left:4px solid {ROJO}; background:{TARJETA}; padding:10px 14px; border-radius:0 10px 10px 0; font-size:.88rem; }}
.nota.ok {{ border-left-color:{C_OK}; }}
.nota.warn {{ border-left-color:{C_BAJO}; }}

/* Embudo de cobertura */
.flujo {{ display:flex; gap:0; border-radius:12px; overflow:hidden; border:1px solid {BORDE}; margin:6px 0 2px 0; }}
.flujo div {{ padding:10px 12px; color:#fff; font-size:.8rem; font-weight:700; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.leyenda {{ display:flex; gap:16px; flex-wrap:wrap; font-size:.8rem; opacity:.85; margin-top:6px; }}
.leyenda span::before {{ content:""; display:inline-block; width:10px; height:10px; border-radius:3px; margin-right:6px; background:var(--c); }}

section[data-testid="stSidebar"] {{ border-right:1px solid {BORDE}; }}
section[data-testid="stSidebar"] h2 {{ color:{ROJO} !important; font-size:1.15rem; font-weight:800; }}
div[data-testid="stFileUploader"] section {{ background:rgba(227,6,19,.06); border:2px dashed {ROJO}; border-radius:14px; min-height:92px; align-items:center; }}
div[data-testid="stFileUploader"] label p {{ white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
button[data-testid="stBaseButton-primary"], div[data-testid="stDownloadButton"] button {{
    background:{ROJO}; color:#fff; border:none; border-radius:10px; font-weight:700; box-shadow:0 3px 10px rgba(227,6,19,.3); }}
button[data-testid="stBaseButton-primary"]:hover, div[data-testid="stDownloadButton"] button:hover {{ background:{ROJO_OSC}; color:#fff; border:none; }}
button[data-testid="stBaseButton-primary"] p, div[data-testid="stDownloadButton"] button p {{ color:#fff; }}
button[data-testid="stBaseButton-secondary"] {{ border:1.5px solid {ROJO}; border-radius:10px; font-weight:600; }}
button[data-baseweb="tab"] {{ font-weight:700; }}
button[data-baseweb="tab"][aria-selected="true"] {{ color:{ROJO}; }}
div[data-baseweb="tab-highlight"] {{ background-color:{ROJO}; height:3px; }}
div[data-testid="stDataFrame"] {{ border:1px solid {BORDE}; border-radius:12px; overflow:hidden; }}
div[data-testid="stExpander"] {{ border:1px solid {BORDE}; border-radius:12px; background:{TARJETA}; }}
div[data-testid="stExpander"] summary p {{ font-weight:700; }}
div[data-testid="stVerticalBlockBorderWrapper"] {{ border-radius:14px; }}
hr {{ border-color:{BORDE}; }}
</style>
""", unsafe_allow_html=True)


# ----------------------------------------------------------------------
# Utilidades de presentación
# ----------------------------------------------------------------------
def seccion(titulo, sub=""):
    st.markdown(f'<div class="seccion">{titulo}<small>{sub}</small></div>', unsafe_allow_html=True)


def kpis(items):
    """items: (etiqueta, valor, subtítulo, color)"""
    html = "".join(
        f'<div class="kpi" style="--ac:{c}"><div class="l">{l}</div><div class="v">{v}</div><div class="s">{s}</div></div>'
        for l, v, s, c in items)
    st.markdown(f'<div class="kpis">{html}</div>', unsafe_allow_html=True)


def nota(texto, tipo=""):
    st.markdown(f'<div class="nota {tipo}">{texto}</div>', unsafe_allow_html=True)


def n(x):
    x = float(x or 0)
    return f"{int(round(x)):,}".replace(",", ".") if abs(x - round(x)) < 1e-9 else f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def alto_tabla(filas, maximo=460):
    """Alto justo para la tabla (sin filas vacías), con tope."""
    return min(maximo, 35 * (filas + 1) + 3)


def pct(a, b):
    return f"{a / b:.0%}" if b else "0%"


def grafico(titulo, sub, chart):
    with st.container(border=True):
        st.markdown(f"**{titulo}**" + (f"  \n<span style='opacity:.62;font-size:.82rem'>{sub}</span>" if sub else ""),
                    unsafe_allow_html=True)
        # fit-x: el ancho se ajusta al contenedor; el alto se respeta (la leyenda no aplasta las barras)
        chart = chart.properties(autosize=alt.AutoSizeParams(type="fit-x", contains="padding"))
        st.altair_chart(chart, width="stretch")


def barras_h(df, cat, val, color=None, escala=None, alto=None, titulo_val="Unidades", orden="-x"):
    """Barras horizontales; el valor va en la etiqueta del eje (legible en claro y oscuro) y en el tooltip."""
    df = df.copy()
    df["_et"] = df[cat].astype(str) + "  ·  " + df[val].map(n)
    if isinstance(orden, list):
        orden = [df.loc[df[cat] == o, "_et"].iloc[0] for o in orden if (df[cat] == o).any()]
    alto = alto or max(120, 34 * len(df) + 20)
    enc_color = (alt.Color(f"{color}:N", scale=escala, legend=alt.Legend(title=None, orient="bottom"))
                 if color else alt.value(C_EXT))
    return alt.Chart(df).mark_bar(cornerRadiusEnd=4, size=20).encode(
        y=alt.Y("_et:N", sort=orden, title=None, axis=alt.Axis(labelLimit=330)),
        x=alt.X(f"{val}:Q", title=None, axis=alt.Axis(format="~s", grid=True)),
        color=enc_color,
        tooltip=[alt.Tooltip(f"{cat}:N", title="Categoría"), alt.Tooltip(f"{val}:Q", title=titulo_val, format=",")]
    ).properties(height=alto)


PASOS = ["1. Origen sin bajar exhibición", "2. Alternativos con sobrante", "3. Origen bajo exhibición",
         "4. Otros locales bajo exhibición", "Pendiente sin cubrir"]
ESC_PASO = alt.Scale(domain=PASOS, range=[C_OK, C_ALT, C_BAJO, C_EXT, C_PEND])
TIPO_A_PASO = {motor.TIPO_ORIGEN: PASOS[0], motor.TIPO_ALT: PASOS[1], motor.TIPO_ORIGEN_BAJO: PASOS[2], motor.TIPO_EXT: PASOS[3]}
VALIDACIONES = ["OK", "DEJA BAJO EXHIBICIÓN", "STOCK INSUFICIENTE", "SIN STOCK EN ORIGEN", "SIN DATO EN BASE"]
ESC_VAL = alt.Scale(domain=VALIDACIONES, range=[C_OK, C_BAJO, C_EXT, ROJO_OSC, C_PEND])


# ----------------------------------------------------------------------
# Lectura (en caché)
# ----------------------------------------------------------------------
class _Archivo(io.BytesIO):
    """BytesIO con nombre, para que el motor detecte la extensión."""
    def __init__(self, data: bytes, name: str):
        super().__init__(data)
        self.name = name


@st.cache_data(show_spinner=False)
def hojas_de(data: bytes, nombre: str) -> list[str]:
    if nombre.lower().endswith(".csv"):
        return []
    engine = "pyxlsb" if nombre.lower().endswith(".xlsb") else None
    return pd.ExcelFile(io.BytesIO(data), engine=engine).sheet_names


@st.cache_data(show_spinner="Leyendo base del BI…")
def cargar_base(data: bytes, nombre: str, hoja):
    return motor.leer_base(_Archivo(data, nombre), hoja)


@st.cache_data(show_spinner="Leyendo solicitudes…")
def cargar_solicitudes(data: bytes, nombre: str, hoja):
    return motor.leer_solicitudes(_Archivo(data, nombre), hoja)


@st.cache_data(show_spinner="Calculando traspasos…", max_entries=5)
def calcular(base, sol, opciones: tuple):
    res = motor.calcular(base, sol, motor.Opciones(*opciones))
    return res, generar_excel(res)


def elegir_hoja(contenedor, data, nombre, preferidas, key):
    hojas = hojas_de(data, nombre)
    if len(hojas) <= 1:
        return hojas[0] if hojas else None
    idx = next((i for i, h in enumerate(hojas) if any(p in motor.normalizar(h) for p in preferidas)), 0)
    return contenedor.selectbox("Hoja", hojas, index=idx, key=key)


# ----------------------------------------------------------------------
# Barra lateral
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("Parámetros")
    st.caption("El signo ❓ explica cada opción.")
    with st.expander("Reglas de búsqueda", expanded=True):
        excluir = st.checkbox("No usar como alternativo a un local que ya es origen del mismo ítem", value=True,
                              help="Evita sugerir un local que ya está cediendo ese estadístico en otra fila de la solicitud.")
        mismo_sup = st.checkbox("Alternativos solo del mismo supervisor del origen", value=False,
                                help="Requiere la columna SUPERV. en la base.")
    with st.expander("Cuando no alcanza el sobrante", expanded=True):
        origen_bajo = st.checkbox("Paso 3 · Completar con el origen aunque quede bajo exhibición", value=True,
                                  help="Lo que no cubren los alternativos se saca del local de origen, hasta agotar su físico.")
        extremos = st.checkbox("Paso 4 · Completar con otros locales aunque queden bajo exhibición", value=True,
                               help="Solo se usa cuando el origen ya quedó en 0. Empieza por el local con más stock físico.")
    with st.expander("📘 Cómo decide el programa"):
        st.markdown(
            "**Orden de cobertura de cada solicitud**\n\n"
            "1. **Origen** sin bajar de su exhibición.\n"
            "2. **Alternativos** con sobrante (físico > exhibición), de mayor a menor sobrante, sin límite de locales.\n"
            "3. **Origen** aunque quede bajo exhibición.\n"
            "4. **Otros locales** aunque queden bajo exhibición, de mayor a menor stock.\n\n"
            "El stock se descuenta fila por fila: lo que un local cede ya no está disponible para las siguientes. "
            "**Las primeras filas tienen prioridad.**\n\n"
            "Nunca se usa el local destino como alternativo.")
opciones = (excluir, mismo_sup, extremos, origen_bajo)

# ----------------------------------------------------------------------
# Encabezado + pasos
# ----------------------------------------------------------------------
hay_base = st.session_state.get("base") is not None
hay_sol = st.session_state.get("sol") is not None or bool(st.session_state.get("_hay_pegado"))
paso_cls = lambda ok: "paso ok" if ok else "paso"
st.markdown(f"""
<div class="hero">
  <div>
    <h1>Verificación de Solicitudes de Traspaso</h1>
    <p>Valida que los traspasos no dejen locales sin exhibición y propone desde dónde sacar la mercadería.</p>
  </div>
  <div class="pasos">
    <div class="{paso_cls(hay_base)}"><b>1</b>Cargar base del BI</div>
    <div class="{paso_cls(hay_sol)}"><b>2</b>Cargar solicitudes</div>
    <div class="{paso_cls(hay_base and hay_sol)}"><b>3</b>Revisar y descargar</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------
# Carga
# ----------------------------------------------------------------------
seccion("Carga de archivos", "La base del BI debe incluir todos los locales que pueden ceder mercadería")
c1, c2 = st.columns(2)
base = sol = None
with c1:
    f_base = st.file_uploader("1) Base de inventario del BI", type=["xlsx", "xlsm", "xls", "xlsb", "csv"], key="base",
                              help="Columnas necesarias: '# Local', 'ESTADISTICO', 'Físico', 'Uni Exhibición'. "
                                   "Opcionales: DESCRIPCION, SUPERV., DESIGNACION, CONSUMOS ACU.")
    if f_base:
        data = f_base.getvalue()
        hoja = elegir_hoja(st, data, f_base.name, ["INVENTARIO", "BASE"], "hoja_base")
        try:
            base = cargar_base(data, f_base.name, hoja)
            st.caption(f"✅ {n(len(base))} registros · {base['local'].nunique()} locales · "
                       f"{base['estadistico'].nunique()} estadísticos")
        except Exception as e:
            st.error(str(e))
with c2:
    t_arch, t_pegar = st.tabs(["2) Subir solicitudes", "Pegar / digitar"])
    with t_arch:
        f_sol = st.file_uploader("Solicitudes de traspaso", type=["xlsx", "xlsm", "xls", "csv"], key="sol",
                                 help="Columnas: 'Loc. Origen', 'Loc. Destino', 'Estadístico', 'Cantidad'. "
                                      "Sirve el reporte de Solicitud de Traspasos.")
        if f_sol:
            data = f_sol.getvalue()
            hoja = elegir_hoja(st, data, f_sol.name, ["TRASLADOS", "SOLICITUD", "TRASPASO"], "hoja_sol")
            try:
                sol = cargar_solicitudes(data, f_sol.name, hoja)
            except Exception as e:
                st.error(str(e))
    with t_pegar:
        st.caption("Copia las 4 columnas desde Excel y pégalas en la primera celda (Ctrl+V).")
        vacio = pd.DataFrame({"Loc. Origen": [""] * 8, "Loc. Destino": [""] * 8,
                              "Estadístico": [""] * 8, "Cantidad": [float("nan")] * 8})
        editado = st.data_editor(vacio, num_rows="dynamic", width="stretch", key="editor", height=250,
                                 column_config={"Cantidad": st.column_config.NumberColumn(min_value=0, step=1)})
        pegado = editado.replace("", pd.NA).dropna(how="all")
        st.session_state["_hay_pegado"] = len(pegado) > 0
        if sol is None and len(pegado):
            try:
                sol = motor.normalizar_solicitudes(pegado)
            except Exception as e:
                st.error(str(e))
    if sol is not None:
        st.caption(f"✅ {len(sol)} solicitudes · {n(sol['cantidad'].sum())} unidades")

if base is None or sol is None or not len(sol):
    nota("Sube la base del BI y las solicitudes (o pégalas) para ver el análisis y descargar el Excel. "
         "Las opciones de la izquierda ya traen sus valores recomendados.")
    st.stop()

# ----------------------------------------------------------------------
# Cálculo
# ----------------------------------------------------------------------
res, excel = calcular(base, sol, opciones)
r = res.resumen
mov = res.movimientos.copy()
tot = r["und_solicitadas"]
sin_afectar = r["paso1"] + r["paso2"]
bajo = r["paso3"] + r["paso4"]

st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
d1, d2 = st.columns([3, 1.3])
if r["pendiente"] > 0:
    d1.markdown(f'<div class="nota">⚠️ <b>{n(r["pendiente"])}</b> unidades no se pueden cubrir con ningún local de la base. '
                f'Revisa la pestaña <b>Pendientes</b>.</div>', unsafe_allow_html=True)
elif bajo > 0:
    d1.markdown(f'<div class="nota warn">Se cubre todo lo solicitado, pero <b>{n(bajo)}</b> unidades dejan locales bajo su exhibición '
                f'({r["locales_bajo_exh"]} local/ítem). Revisa el <b>Reporte bajo exhibición</b>.</div>', unsafe_allow_html=True)
else:
    d1.markdown('<div class="nota ok">✅ Se cubre todo lo solicitado sin dejar ningún local bajo su exhibición.</div>',
                unsafe_allow_html=True)
d2.download_button("⬇️ Descargar Excel de resultados", data=excel,
                   file_name=f"VERIFICACION_TRASPASOS_{datetime.now():%Y%m%d_%H%M}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

t_res, t_tg, t_rep, t_det, t_pen = st.tabs([
    "Resumen", f"Traspasos a grabar ({len(res.traspasos)})", f"Reporte bajo exhibición ({len(res.reporte)})",
    f"Detalle por solicitud ({len(res.detalle)})", f"Pendientes ({len(res.pendientes)})"])

# ============================ RESUMEN ============================
with t_res:
    kpis([
        ("Unidades solicitadas", n(tot), f"{r['solicitudes']} solicitudes", ROJO),
        ("Sin afectar exhibición", n(sin_afectar), f"{pct(sin_afectar, tot)} de lo solicitado", C_OK),
        ("Bajo exhibición", n(bajo), f"{pct(bajo, tot)} · origen {n(r['paso3'])} · otros {n(r['paso4'])}", C_BAJO),
        ("Pendiente sin cubrir", n(r["pendiente"]), f"{pct(r['pendiente'], tot)} de lo solicitado", C_PEND),
        ("Traspasos a grabar", n(r["traspasos"]), f"{mov['local_envia'].nunique() if len(mov) else 0} locales envían", C_ALT),
        ("Locales/ítems afectados", n(r["locales_bajo_exh"]), "Quedan bajo su exhibición", C_EXT),
    ])

    # Embudo horizontal de cobertura
    partes = [(PASOS[0], r["paso1"], C_OK), (PASOS[1], r["paso2"], C_ALT), (PASOS[2], r["paso3"], C_BAJO),
              (PASOS[3], r["paso4"], C_EXT), (PASOS[4], r["pendiente"], C_PEND)]
    barras = "".join(f'<div style="flex:{v};background:{c}" title="{p}: {n(v)}">{pct(v, tot)}</div>' for p, v, c in partes if v > 0)
    ley = "".join(f'<span style="--c:{c}">{p}: <b>{n(v)}</b></span>' for p, v, c in partes)
    seccion("¿De dónde sale lo solicitado?", "Reparto de las unidades según el paso en que se cubrieron")
    st.markdown(f'<div class="flujo">{barras}</div><div class="leyenda">{ley}</div>', unsafe_allow_html=True)

    g1, g2 = st.columns([1.6, 1])
    with g1:
        desc = base.drop_duplicates("estadistico").set_index("estadistico")["descripcion"].to_dict()
        dprod = mov.assign(paso=mov["tipo"].map(TIPO_A_PASO))[["estadistico", "paso", "cantidad"]] if len(mov) else pd.DataFrame(columns=["estadistico", "paso", "cantidad"])
        pend = res.detalle.loc[res.detalle["Pendiente sin Cubrir"] > 0, ["Estadístico", "Pendiente sin Cubrir"]]
        pend = pend.rename(columns={"Estadístico": "estadistico", "Pendiente sin Cubrir": "cantidad"}).assign(paso=PASOS[4])
        dprod = pd.concat([dprod, pend], ignore_index=True)
        dprod["prod"] = dprod["estadistico"].map(lambda g: f"{g} · {str(desc.get(g, ''))[:34].title()}")
        dprod = dprod.groupby(["prod", "paso"], as_index=False)["cantidad"].sum()
        # ordenar por unidades que afectan exhibición (pasos 3, 4 y pendiente) y luego por total
        crit = dprod.assign(mal=dprod["cantidad"].where(dprod["paso"].isin(PASOS[2:]), 0)).groupby("prod")[["mal", "cantidad"]].sum()
        top_p = crit.sort_values(["mal", "cantidad"], ascending=False).head(10).index.tolist()
        dprod = dprod[dprod["prod"].isin(top_p)]
        ch_p = alt.Chart(dprod).mark_bar(cornerRadiusEnd=3).encode(
            y=alt.Y("prod:N", sort=top_p, title=None, axis=alt.Axis(labelLimit=300)),
            x=alt.X("cantidad:Q", title=None, axis=alt.Axis(format="~s")),
            color=alt.Color("paso:N", scale=ESC_PASO, legend=None),
            order=alt.Order("paso:N"),
            tooltip=[alt.Tooltip("prod:N", title="Estadístico"), alt.Tooltip("paso:N", title="Paso"),
                     alt.Tooltip("cantidad:Q", title="Unidades", format=",")]
        ).properties(height=max(190, 30 * len(top_p) + 50))
        grafico("Estadísticos que más afectan la exhibición", "Top 10: primero los que tuvieron que sacarse bajo exhibición o quedaron pendientes. Colores como en la barra de arriba.", ch_p)
    with g2:
        vc = res.detalle["Validación"].value_counts().reindex(VALIDACIONES).dropna().rename_axis("Validación").reset_index(name="Solicitudes")
        vc["Porcentaje"] = vc["Solicitudes"] / max(len(res.detalle), 1)
        ch = alt.Chart(vc).mark_arc(innerRadius=58, stroke="rgba(0,0,0,0)").encode(
            theta="Solicitudes:Q", color=alt.Color("Validación:N", scale=ESC_VAL, legend=alt.Legend(title=None, orient="bottom", columns=2)),
            tooltip=["Validación:N", "Solicitudes:Q", alt.Tooltip("Porcentaje:Q", format=".0%")]).properties(height=210)
        grafico("Estado inicial de las solicitudes", "Si se enviara todo desde el origen solicitado", ch)

    if len(mov):
        mov["paso"] = mov["tipo"].map(TIPO_A_PASO)
        mov["Local"] = mov["local_envia"].astype(str)
        nombres = base.drop_duplicates("local").set_index("local")["designacion"].to_dict()
        mov["loc"] = mov["local_envia"].map(lambda l: f"{l} · {str(nombres.get(l, '')).title()}".strip(" ·"))
        mov["dest"] = mov["destino"].map(lambda d: f"{int(d)} · {str(nombres.get(int(d), '')).title()}".strip(" ·")
                                            if pd.notna(d) else "Sin destino")

        h1, h2 = st.columns(2)
        with h1:
            top = mov.groupby("loc")["cantidad"].sum().sort_values(ascending=False).head(12).index.tolist()
            dl = mov[mov["loc"].isin(top)].groupby(["loc", "paso"], as_index=False)["cantidad"].sum()
            ch_l = alt.Chart(dl).mark_bar(cornerRadiusEnd=3).encode(
                y=alt.Y("loc:N", sort=top, title=None, axis=alt.Axis(labelLimit=240)),
                x=alt.X("cantidad:Q", title=None, axis=alt.Axis(format="~s")),
                color=alt.Color("paso:N", scale=ESC_PASO, legend=None),
                order=alt.Order("paso:N"),
                tooltip=[alt.Tooltip("loc:N", title="Local"), alt.Tooltip("paso:N", title="Paso"),
                         alt.Tooltip("cantidad:Q", title="Unidades", format=",")]
            ).properties(height=max(180, 30 * len(top) + 50))
            grafico("Locales que envían mercadería", "Top 12 por unidades; el color indica el paso en que ceden.", ch_l)
        with h2:
            dd = mov.groupby(["dest", "paso"], as_index=False)["cantidad"].sum()
            top_d = dd.groupby("dest")["cantidad"].sum().sort_values(ascending=False).head(12).index.tolist()
            dd = dd[dd["dest"].isin(top_d)]
            ch_d = alt.Chart(dd).mark_bar(cornerRadiusEnd=3).encode(
                y=alt.Y("dest:N", sort=top_d, title=None, axis=alt.Axis(labelLimit=240)),
                x=alt.X("cantidad:Q", title=None, axis=alt.Axis(format="~s")),
                color=alt.Color("paso:N", scale=ESC_PASO, legend=None),
                order=alt.Order("paso:N"),
                tooltip=[alt.Tooltip("dest:N", title="Destino"), alt.Tooltip("paso:N", title="Paso"),
                         alt.Tooltip("cantidad:Q", title="Unidades", format=",")]
            ).properties(height=max(180, 30 * len(top_d) + 50))
            grafico("Locales que reciben mercadería", "Top 12 destinos por unidades; el color indica de qué paso vienen.", ch_d)

        # Matriz origen x destino
        mx = mov.groupby(["Local", "destino"], as_index=False, dropna=False)["cantidad"].sum()
        mx["Destino"] = mx["destino"].map(lambda d: str(int(d)) if pd.notna(d) else "—")
        if mx["Local"].nunique() > 1 or mx["Destino"].nunique() > 1:
            base_m = alt.Chart(mx).encode(
                x=alt.X("Destino:N", title="Local destino", axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Local:N", title="Local que envía"))
            heat = base_m.mark_rect(cornerRadius=4).encode(
                color=alt.Color("cantidad:Q", scale=alt.Scale(scheme="reds"), legend=alt.Legend(title="Unidades")),
                tooltip=[alt.Tooltip("Local:N", title="Envía"), alt.Tooltip("Destino:N", title="Recibe"),
                         alt.Tooltip("cantidad:Q", title="Unidades", format=",")])
            txt = base_m.mark_text(fontSize=11, fontWeight="bold").encode(
                text=alt.Text("cantidad:Q", format=",.0f"),
                color=alt.condition(alt.datum.cantidad > mx["cantidad"].max() * 0.5, alt.value("white"), alt.value("black")))
            alto = max(160, 34 * mx["Local"].nunique() + 50)
            grafico("Matriz de traspasos", "Unidades que envía cada local (filas) a cada destino (columnas).",
                    (heat + txt).properties(height=alto))

    if len(res.reporte):
        rp = res.reporte.copy()
        rp["li"] = rp["Local que Envía"].astype(str) + " · " + rp["Descripción"].str.slice(0, 38)
        rp = rp.sort_values("Faltante Exhibición", ascending=False).head(15)
        largo = rp.melt(id_vars=["li"], value_vars=["Stock Final", "Faltante Exhibición"],
                        var_name="Concepto", value_name="Unidades")
        ch_r = alt.Chart(largo).mark_bar(cornerRadiusEnd=3).encode(
            y=alt.Y("li:N", sort=rp["li"].tolist(), title=None, axis=alt.Axis(labelLimit=360)),
            x=alt.X("Unidades:Q", title="Unidades de exhibición", stack="zero"),
            color=alt.Color("Concepto:N", scale=alt.Scale(domain=["Stock Final", "Faltante Exhibición"], range=[C_OK, C_EXT]),
                            legend=alt.Legend(title=None, orient="bottom")),
            tooltip=[alt.Tooltip("li:N", title="Local · Ítem"), "Concepto:N", alt.Tooltip("Unidades:Q", format=",")]
        ).properties(height=max(160, 28 * len(rp) + 40))
        grafico("Exhibición que queda descubierta", "Top 15 local/ítem: en verde lo que queda en percha, en rojo lo que falta para completar la exhibición.", ch_r)

# ============================ TRASPASOS A GRABAR ============================
with t_tg:
    nota("Lista final consolidada por <b>origen + destino + estadístico</b>. Las 4 primeras columnas están en el orden para grabar "
         "en el sistema; el Excel descargable las trae igual.")
    tg = res.traspasos.copy()
    if len(tg):
        f1, f2, f3 = st.columns([1.2, 1.2, 1])
        tipos = sorted({t for x in tg["Tipo de Origen"] for t in x.split(" + ")})
        sel_t = f1.multiselect("Tipo de origen", tipos, placeholder="Todos")
        sel_l = f2.multiselect("Local que envía", sorted(tg["Loc. Origen"].unique()), placeholder="Todos")
        solo_bajo = f3.toggle("Solo los que quedan bajo exhibición")
        if sel_t:
            tg = tg[tg["Tipo de Origen"].map(lambda x: any(t in x for t in sel_t))]
        if sel_l:
            tg = tg[tg["Loc. Origen"].isin(sel_l)]
        if solo_bajo:
            tg = tg[tg["¿Queda Bajo Exhibición?"] == "SÍ"]
        tg["Cobertura Exhibición Final"] = (tg["Stock Final Origen"] / tg["Exhibición Origen"].where(tg["Exhibición Origen"] > 0)).clip(0, 1).fillna(1)
        orden = ["Loc. Origen", "Loc. Destino", "Estadístico", "Cantidad a Traspasar", "Tipo de Origen", "¿Queda Bajo Exhibición?",
                 "Cobertura Exhibición Final", "Exhibición Origen", "Físico Origen", "Total que Envía (ítem)", "Stock Final Origen",
                 "Und. Bajo Exhibición Causadas", "Descripción", "Supervisor Origen", "Filas Solicitud"]
        tg = tg[[c for c in orden if c in tg.columns]]
        st.dataframe(tg, hide_index=True, width="stretch", height=alto_tabla(len(tg)), column_config={
            "Cantidad a Traspasar": st.column_config.NumberColumn(format="%d"),
            "Cobertura Exhibición Final": st.column_config.ProgressColumn("Cobertura exhib. final", format="percent", min_value=0, max_value=1),
            "Descripción": st.column_config.TextColumn(width="medium"),
            "Tipo de Origen": st.column_config.TextColumn(width="medium"),
        })
        st.caption(f"{len(tg)} traspasos · {n(tg['Cantidad a Traspasar'].sum())} unidades")
    else:
        st.info("No hay traspasos para grabar.")

# ============================ REPORTE ============================
with t_rep:
    if len(res.reporte):
        nota("Cada fila es un <b>local + estadístico</b> que, después de todos los traspasos, queda por debajo de su exhibición. "
             "<b>Und. bajo exhibición causadas</b> descuenta lo que ya faltaba antes del traspaso.", "warn")
        rep = res.reporte.copy()
        kpis([("Locales/ítems afectados", n(len(rep)), f"{rep['Local que Envía'].nunique()} locales distintos", C_EXT),
              ("Und. de exhibición descubiertas", n(rep["Faltante Exhibición"].sum()), "Total faltante tras los traspasos", C_BAJO),
              ("Causadas por los traspasos", n(rep["Und. Bajo Exhibición Causadas"].sum()), "Sin contar lo que ya faltaba", ROJO)])
        rep["Cobertura Exhibición"] = (rep["Stock Final"] / rep["Exhibición"].where(rep["Exhibición"] > 0)).clip(0, 1).fillna(1)
        orden = ["Local que Envía", "Nombre Local", "Estadístico", "Descripción", "Exhibición", "Físico Inicial", "Total Enviado",
                 "Stock Final", "Cobertura Exhibición", "Faltante Exhibición", "Und. Bajo Exhibición Causadas", "Destinos", "Motivo", "Supervisor"]
        rep = rep[[c for c in orden if c in rep.columns]]
        st.dataframe(rep, hide_index=True, width="stretch", height=alto_tabla(len(rep)), column_config={
            "Cobertura Exhibición": st.column_config.ProgressColumn("Cobertura exhibición", format="percent", min_value=0, max_value=1),
            "Descripción": st.column_config.TextColumn(width="medium"),
            "Motivo": st.column_config.TextColumn(width="large"),
        })
    else:
        nota("✅ Ningún local queda bajo su exhibición con estos traspasos.", "ok")

# ============================ DETALLE ============================
with t_det:
    nota("Cómo se cubrió cada solicitud, paso por paso. Entre paréntesis, las unidades que aporta cada local.")
    det = res.detalle.copy()
    f1, f2 = st.columns(2)
    sel_v = f1.multiselect("Validación inicial", [v for v in VALIDACIONES if v in set(det["Validación"])], placeholder="Todas")
    solo_rep = f2.toggle("Solo filas repetidas")
    if sel_v:
        det = det[det["Validación"].isin(sel_v)]
    if solo_rep:
        det = det[det["Repetido"] != ""]
    st.dataframe(det, hide_index=True, width="stretch", height=alto_tabla(len(det)),
                 column_config={"Descripción": st.column_config.TextColumn(width="large")})

# ============================ PENDIENTES ============================
with t_pen:
    if len(res.pendientes):
        nota(f"<b>{n(r['pendiente'])}</b> unidades no tienen stock en ningún local de la base. "
             "Si la base solo trae algunos locales, descarga una con más locales para buscar alternativas.")
        st.dataframe(res.pendientes, hide_index=True, width="stretch", height=alto_tabla(len(res.pendientes)),
                     column_config={"Descripción": st.column_config.TextColumn(width="large")})
    else:
        nota("✅ Todo lo solicitado quedó cubierto.", "ok")
