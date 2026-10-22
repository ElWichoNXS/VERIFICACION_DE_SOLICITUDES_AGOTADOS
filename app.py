"""Verificación de solicitudes de traspaso — app Streamlit."""
import io
from datetime import datetime

import pandas as pd
import streamlit as st

import motor
from exportar import generar_excel

st.set_page_config(page_title="Verificación de Traspasos", page_icon="📦", layout="wide")


# ------------------------------------------------------------------ helpers
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


def elegir_hoja(data, nombre, preferidas, key):
    hojas = hojas_de(data, nombre)
    if not hojas:
        return None
    idx = 0
    for i, h in enumerate(hojas):
        if any(p in motor.normalizar(h) for p in preferidas):
            idx = i
            break
    if len(hojas) == 1:
        return hojas[0]
    return st.selectbox("Hoja", hojas, index=idx, key=key)


def n(x):
    return f"{int(x):,}".replace(",", ".") if float(x).is_integer() else f"{x:,.2f}"


# ------------------------------------------------------------------ barra lateral
with st.sidebar:
    st.header("⚙️ Opciones")
    excluir = st.checkbox("No usar como alternativo a un local que ya es origen del mismo ítem", value=True,
                          help="Evita sugerir un local que ya está cediendo ese ítem en otra fila de la solicitud.")
    mismo_sup = st.checkbox("Alternativos solo del mismo supervisor del origen", value=False)
    origen_bajo = st.checkbox("Paso 3: completar con el origen aunque quede bajo exhibición", value=True)
    extremos = st.checkbox("Paso 4: completar con otros locales aunque queden bajo exhibición", value=True,
                           help="Solo se usa cuando el origen ya quedó en 0 y no hay más sobrantes.")
    st.divider()
    st.markdown(
        "**Orden de cobertura**\n\n"
        "1. Origen sin bajar de exhibición\n"
        "2. Alternativos con sobrante (sin límite de locales)\n"
        "3. Origen bajo exhibición\n"
        "4. Otros locales bajo exhibición\n\n"
        "Las primeras filas tienen prioridad sobre el stock."
    )

st.title("📦 Verificación de Solicitudes de Traspaso")
st.caption("Valida que los traspasos no dejen locales sin exhibición y propone desde dónde sacar la mercadería.")

col1, col2 = st.columns(2)

# ------------------------------------------------------------------ 1. base
with col1:
    st.subheader("1. Base de inventario (BI)")
    f_base = st.file_uploader("Archivo con '# Local', 'ESTADISTICO', 'Físico', 'Uni Exhibición'",
                              type=["xlsx", "xlsm", "xls", "xlsb", "csv"], key="base")
    base = None
    if f_base:
        data = f_base.getvalue()
        hoja = elegir_hoja(data, f_base.name, ["INVENTARIO", "BASE"], "hoja_base")
        try:
            base = cargar_base(data, f_base.name, hoja)
            st.success(f"{n(len(base))} registros · {base['local'].nunique()} locales · "
                       f"{base['estadistico'].nunique()} estadísticos")
        except Exception as e:
            st.error(str(e))

# ------------------------------------------------------------------ 2. solicitudes
with col2:
    st.subheader("2. Solicitudes de traspaso")
    t_arch, t_pegar = st.tabs(["Subir archivo", "Pegar / digitar"])
    sol = None
    with t_arch:
        f_sol = st.file_uploader("Archivo con 'Loc. Origen', 'Loc. Destino', 'Estadístico', 'Cantidad'",
                                 type=["xlsx", "xlsm", "xls", "csv"], key="sol")
        if f_sol:
            data = f_sol.getvalue()
            hoja = elegir_hoja(data, f_sol.name, ["TRASLADOS", "SOLICITUD", "TRASPASO"], "hoja_sol")
            try:
                sol = cargar_solicitudes(data, f_sol.name, hoja)
            except Exception as e:
                st.error(str(e))
    with t_pegar:
        st.caption("Copia las 4 columnas desde Excel y pégalas en la primera celda (Ctrl+V). "
                   "Acepta '189' o '189-GUASMO…' y '51404000' o '051404000_MEDIAS…'.")
        vacio = pd.DataFrame({"Loc. Origen": [""] * 10, "Loc. Destino": [""] * 10,
                              "Estadístico": [""] * 10, "Cantidad": [float("nan")] * 10})
        editado = st.data_editor(vacio, num_rows="dynamic", width="stretch", key="editor",
                                 column_config={"Cantidad": st.column_config.NumberColumn(min_value=0, step=1)})
        if f_sol is None:
            pegado = editado.replace("", pd.NA).dropna(how="all")
            if len(pegado):
                try:
                    sol = motor.normalizar_solicitudes(pegado)
                except Exception as e:
                    st.error(str(e))
    if sol is not None:
        st.success(f"{len(sol)} solicitudes · {n(sol['cantidad'].sum())} unidades")

# ------------------------------------------------------------------ 3. calcular
st.divider()
listo = base is not None and sol is not None and len(sol) > 0
if st.button("▶️ Calcular traspasos", type="primary", disabled=not listo, width="stretch"):
    op = motor.Opciones(excluir_origenes=excluir, mismo_supervisor=mismo_sup,
                        usar_extremos=extremos, origen_bajo_exhibicion=origen_bajo)
    with st.spinner("Calculando…"):
        st.session_state["res"] = motor.calcular(base, sol, op)
if not listo:
    st.info("Carga la base y las solicitudes para calcular.")

res = st.session_state.get("res")
if res is not None and listo:
    r = res.resumen
    st.subheader("Resultado")
    m = st.columns(6)
    m[0].metric("Unidades solicitadas", n(r["und_solicitadas"]))
    m[1].metric("1. Origen OK", n(r["paso1"]))
    m[2].metric("2. Alternativos", n(r["paso2"]))
    m[3].metric("3. Origen bajo exh.", n(r["paso3"]))
    m[4].metric("4. Extremos", n(r["paso4"]))
    m[5].metric("Pendiente", n(r["pendiente"]))
    if r["repetidos"]:
        st.warning(f"Hay {r['repetidos']} filas REPETIDAS (mismo origen, destino y estadístico). Revisa el detalle.")
    if r["pendiente"] > 0:
        st.error(f"{n(r['pendiente'])} unidades no se pueden cubrir con ningún local de la base.")

    st.download_button("⬇️ Descargar Excel de resultados", data=generar_excel(res),
                       file_name=f"VERIFICACION_TRASPASOS_{datetime.now():%Y%m%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", width="stretch")

    tabs = st.tabs([f"Traspasos a grabar ({len(res.traspasos)})",
                    f"Reporte bajo exhibición ({len(res.reporte)})",
                    f"Detalle por solicitud ({len(res.detalle)})",
                    f"Pendientes ({len(res.pendientes)})"])
    with tabs[0]:
        st.dataframe(res.traspasos, width="stretch", hide_index=True)
    with tabs[1]:
        st.dataframe(res.reporte, width="stretch", hide_index=True)
    with tabs[2]:
        st.dataframe(res.detalle, width="stretch", hide_index=True)
    with tabs[3]:
        st.dataframe(res.pendientes, width="stretch", hide_index=True)
