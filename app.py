"""
Tablero de eficiencia de produccion - planta transaccional de alimentos.

Aplicacion de Streamlit sobre la base de datos de 4 lineas, 3 turnos y 12 meses.
Los datos son simulados; el aviso completo esta en la hoja Leame del xlsx.

Esta aplicacion no contiene el modelo de frontera eficiente ni ninguna rutina de
optimizacion. Entrega la matriz de retornos por periodo y ahi termina.

Uso local:
    streamlit run app.py
"""

from pathlib import Path
import sys

import pandas as pd
import streamlit as st

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ))

from src.gestion_bd import BaseDatosPlanta  # noqa: E402

st.set_page_config(page_title="Eficiencia de produccion",
                   page_icon=None, layout="wide")


# ------------------------------------------------------------------ carga
@st.cache_data(show_spinner="Cargando la base de datos")
def cargar_tablas() -> dict[str, pd.DataFrame]:
    bd = BaseDatosPlanta.cargar(RAIZ / "data" / "csv",
                                RAIZ / "data" / "BD_Eficiencia_Planta.xlsx")
    return {k: getattr(bd, k) for k in BaseDatosPlanta.HOJAS}


def construir(tablas: dict[str, pd.DataFrame]) -> BaseDatosPlanta:
    return BaseDatosPlanta(ruta=None, tablas={k: v.copy() for k, v in tablas.items()})


def kpis(d: pd.DataFrame) -> dict[str, float]:
    """Agregacion ponderada. Nunca promedio simple de porcentajes."""
    if d.empty:
        return {k: 0.0 for k in ["Disponibilidad", "Desempeno", "Calidad", "OEE",
                                 "Rendimiento", "YTP", "MC_hora", "Cajas",
                                 "Margen", "Horas_Prog"]}
    disp = d.Horas_Operativas.sum() / d.Horas_Programadas.sum()
    des = d.Cajas_Producidas.sum() / d.Cajas_Teoricas.sum()
    cal = d.Cajas_Buenas.sum() / d.Cajas_Producidas.sum()
    return {
        "Disponibilidad": disp, "Desempeno": des, "Calidad": cal,
        "OEE": disp * des * cal,
        "Rendimiento": d.Kg_MP_Perfection.sum() / d.Kg_MP_Real.sum(),
        "YTP": d.Act_Abs_Loss_USD.sum() / d.Perfection_Cost_USD.sum(),
        "MC_hora": d.Margen_Contribucion_USD.sum() / d.Horas_Programadas.sum(),
        "Cajas": d.Cajas_Buenas.sum(),
        "Margen": d.Margen_Contribucion_USD.sum(),
        "Horas_Prog": d.Horas_Programadas.sum(),
    }


tablas = cargar_tablas()
bd_total = construir(tablas)
prod = bd_total.prod

# ---------------------------------------------------------------- filtros
st.sidebar.header("Filtros")
f_min, f_max = prod.Fecha.min().date(), prod.Fecha.max().date()
rango = st.sidebar.date_input("Rango de fechas", (f_min, f_max),
                              min_value=f_min, max_value=f_max)
if isinstance(rango, tuple) and len(rango) == 2:
    desde, hasta = rango
else:
    desde, hasta = f_min, f_max

lineas_sel = st.sidebar.multiselect("Lineas", sorted(prod.ID_Linea.unique()),
                                    default=sorted(prod.ID_Linea.unique()))
turnos_sel = st.sidebar.multiselect("Turnos", sorted(prod.ID_Turno.unique()),
                                    default=sorted(prod.ID_Turno.unique()))
skus_sel = st.sidebar.multiselect("SKU", sorted(prod.ID_SKU.unique()),
                                  default=sorted(prod.ID_SKU.unique()))
frecuencia = st.sidebar.selectbox("Frecuencia de la matriz de retornos",
                                  ["D", "W", "ME"], index=0,
                                  format_func=lambda x: {"D": "Diaria",
                                                         "W": "Semanal",
                                                         "ME": "Mensual"}[x])
definicion_activo = st.sidebar.selectbox(
    "Definicion de activo", ["linea_turno", "linea"], index=0,
    format_func=lambda x: {"linea_turno": "Linea x turno (12)",
                           "linea": "Linea (4)"}[x])

mask = (prod.Fecha.dt.date.between(desde, hasta)
        & prod.ID_Linea.isin(lineas_sel)
        & prod.ID_Turno.isin(turnos_sel)
        & prod.ID_SKU.isin(skus_sel))
d = prod[mask]

if d.empty:
    st.warning("Los filtros no dejan ningun registro. Amplia el rango o agrega lineas.")
    st.stop()

sub = {k: v.copy() for k, v in tablas.items()}
sub["prod"] = d.copy()
sub["paros"] = bd_total.paros[bd_total.paros.ID_Registro.isin(d.ID_Registro)].copy()
bd = BaseDatosPlanta(ruta=None, tablas=sub)

# ------------------------------------------------------------- encabezado
st.title("Eficiencia de produccion")
st.caption(f"{len(d):,} registros de turno entre {desde} y {hasta}. "
           "Datos simulados: no provienen de ningun sistema real.")

k = kpis(d)
c = st.columns(6)
c[0].metric("OEE", f"{k['OEE']:.1%}")
c[1].metric("Disponibilidad", f"{k['Disponibilidad']:.1%}")
c[2].metric("Desempeno", f"{k['Desempeno']:.1%}")
c[3].metric("Calidad", f"{k['Calidad']:.2%}")
c[4].metric("Rendimiento diario", f"{k['Rendimiento']:.2%}")
c[5].metric("Yield loss to perfection", f"{k['YTP']:.2%}")

c = st.columns(3)
c[0].metric("Cajas buenas", f"{k['Cajas']:,.0f}")
c[1].metric("Margen de contribucion", f"${k['Margen']:,.0f}")
c[2].metric("Margen por hora programada", f"${k['MC_hora']:,.0f}")

t1, t2, t3, t4, t5, t6 = st.tabs(
    ["Resumen", "Tendencia", "Paros", "Yield", "Matriz de retornos", "Validacion"])

# ----------------------------------------------------------------- resumen
with t1:
    corte = st.radio("Agrupar por", ["linea", "turno", "linea_turno", "sku", "mes"],
                     horizontal=True, key="corte_resumen")
    cols = ["Disponibilidad", "Desempeno", "Calidad", "OEE", "Rendimiento_pct",
            "YTP_pct", "Cajas_Buenas", "Margen_Contribucion_USD",
            "MC_por_Hora_USD", "MC_por_Hora_sd"]
    r = bd.resumen(corte)[cols]
    st.dataframe(
        r.style.format({"Disponibilidad": "{:.1%}", "Desempeno": "{:.1%}",
                        "Calidad": "{:.2%}", "OEE": "{:.1%}",
                        "Rendimiento_pct": "{:.2%}", "YTP_pct": "{:.2%}",
                        "Cajas_Buenas": "{:,.0f}",
                        "Margen_Contribucion_USD": "${:,.0f}",
                        "MC_por_Hora_USD": "${:,.0f}",
                        "MC_por_Hora_sd": "${:,.0f}"}),
        use_container_width=True)
    if corte in ("linea", "turno", "linea_turno"):
        st.bar_chart(r["OEE"])

# --------------------------------------------------------------- tendencia
with t2:
    grano = st.radio("Grano", ["Semanal", "Mensual"], horizontal=True, key="grano_tend")
    periodo = "W" if grano == "Semanal" else "M"
    tmp = d.copy()
    tmp["Periodo"] = tmp.Fecha.dt.to_period(periodo).dt.to_timestamp()
    agr = (tmp.groupby(["Periodo", "ID_Linea"])
           .apply(lambda g: pd.Series(kpis(g)), include_groups=False)
           .reset_index())
    st.write("OEE por linea")
    st.line_chart(agr.pivot(index="Periodo", columns="ID_Linea", values="OEE"))
    st.write("Margen de contribucion por hora programada (USD/h)")
    st.line_chart(agr.pivot(index="Periodo", columns="ID_Linea", values="MC_hora"))
    st.write("Yield loss to perfection")
    st.line_chart(agr.pivot(index="Periodo", columns="ID_Linea", values="YTP"))

# ------------------------------------------------------------------ paros
with t3:
    top = st.slider("Motivos a mostrar", 3, 16, 10)
    par = bd.pareto_paros(top)
    st.dataframe(par.style.format({"Horas": "{:,.1f}",
                                   "Cajas_No_Producidas": "{:,.0f}",
                                   "Pct_Horas": "{:.1%}",
                                   "Pct_Acumulado": "{:.1%}"}),
                 use_container_width=True)
    st.bar_chart(par.reset_index().set_index("Motivo")["Horas"])
    st.caption("Solo paros no programados. El detalle concilia contra la columna "
               "Paro_No_Programado_h de la tabla de hechos.")

# ------------------------------------------------------------------ yield
with t4:
    st.write("Marco de yield: perfection cost, MUV, act abs loss y YTP")
    y = bd.resumen("linea")[["Perfection_Cost_USD", "Act_Abs_Loss_USD",
                             "Rendimiento_pct", "YTP_pct"]]
    y["Brecha_pp"] = (1 - y.Rendimiento_pct) * 100
    st.dataframe(y.style.format({"Perfection_Cost_USD": "${:,.0f}",
                                 "Act_Abs_Loss_USD": "${:,.0f}",
                                 "Rendimiento_pct": "{:.2%}",
                                 "YTP_pct": "{:.2%}",
                                 "Brecha_pp": "{:.2f}"}),
                 use_container_width=True)
    st.caption("Act abs loss = costo de scrap estandar + MUV. "
               "YTP = act abs loss / perfection cost. "
               "El rendimiento diario (kg producto / kg materia prima) es otra "
               "medida y por eso se reporta en columna aparte.")
    st.write("MUV acumulado por linea (USD)")
    muv = d.groupby([d.Fecha.dt.to_period("M").dt.to_timestamp(), "ID_Linea"]) \
        .MUV_USD.sum().unstack().cumsum()
    st.line_chart(muv)

# ------------------------------------------------------- matriz de retornos
with t5:
    m = bd.matriz_retornos(activo=definicion_activo, frecuencia=frecuencia)
    st.write(f"Margen de contribucion por hora programada. "
             f"{m.shape[0]} periodos x {m.shape[1]} activos.")
    st.dataframe(m.style.format("${:,.0f}"), use_container_width=True, height=360)
    desc = pd.DataFrame({"Media_USD_h": m.mean(), "Desviacion_USD_h": m.std(),
                         "Minimo": m.min(), "Maximo": m.max(),
                         "Periodos": m.count()})
    st.write("Estadistica descriptiva por activo")
    st.dataframe(desc.style.format({"Media_USD_h": "${:,.0f}",
                                    "Desviacion_USD_h": "${:,.0f}",
                                    "Minimo": "${:,.0f}", "Maximo": "${:,.0f}",
                                    "Periodos": "{:,.0f}"}),
                 use_container_width=True)
    st.download_button("Descargar matriz de retornos (CSV)",
                       m.to_csv(float_format="%.6f").encode("utf-8"),
                       file_name=f"matriz_retornos_{definicion_activo}_{frecuencia}.csv",
                       mime="text/csv")
    st.download_button("Descargar detalle filtrado (CSV)",
                       d.to_csv(index=False).encode("utf-8"),
                       file_name="fact_produccion_filtrado.csv", mime="text/csv")
    st.info("Este tablero se detiene aqui a proposito: entrega la matriz de "
            "retornos y no incluye el modelo de optimizacion.")

# -------------------------------------------------------------- validacion
with t6:
    st.write("Pruebas de integridad sobre la base completa, sin filtros")
    v = bd_total.validar()
    fallas = int((v.Resultado == "FALLA").sum())
    (st.success if fallas == 0 else st.error)(
        f"{len(v) - fallas} de {len(v)} pruebas en OK")
    st.dataframe(v, use_container_width=True, height=520)
