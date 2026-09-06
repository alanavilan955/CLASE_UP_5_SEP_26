"""
Tablero de eficiencia de produccion - planta transaccional de alimentos.

Version de archivo unico: no importa nada del repositorio, asi que funciona
aunque la carpeta src no se haya subido a GitHub. Localiza los datos sola,
buscando primero un paquete de CSV y luego el libro de Excel.

Los datos son simulados. Esta aplicacion no contiene el modelo de frontera
eficiente ni ninguna rutina de optimizacion: entrega la matriz de retornos por
periodo y ahi termina.

Uso local:
    streamlit run app.py
"""

from pathlib import Path

import numpy as np  # noqa: F401  (usado por pandas en algunas agregaciones)
import pandas as pd
import streamlit as st

RAIZ = Path(__file__).resolve().parent
ARCHIVO = "data/BD_Eficiencia_Planta.xlsx"
CARPETA_CSV = "data/csv"

# las reproduce para verificar que el archivo y el codigo dan el mismo numero.
COLUMNAS_FORMULA = [
    "Horas_Programadas", "Horas_Operativas", "Cajas_Teoricas", "Cajas_Buenas",
    "Disponibilidad", "Desempeno", "Calidad", "OEE", "Kg_MP_Perfection",
    "Kg_MP_Estandar", "Rendimiento_pct", "Perfection_Cost_USD",
    "Costo_Scrap_Estandar_USD", "Standard_Prime_Cost_USD", "MUV_USD",
    "Act_Abs_Loss_USD", "YTP_pct", "Costo_MP_Real_USD", "Costo_MOD_USD",
    "Costo_Envase_USD", "Ventas_USD", "Margen_Contribucion_USD",
    "MC_por_Hora_Programada_USD", "Costo_Total_por_Caja_USD",
]

# Posicion de cada formula en la hoja, para dar de alta registros nuevos.
FORMULAS_EXCEL = {
    16: "=N{r}-O{r}", 18: "=P{r}-Q{r}", 20: "=S{r}*R{r}", 23: "=U{r}-V{r}",
    24: "=IFERROR(R{r}/P{r},0)", 25: "=IFERROR(U{r}/T{r},0)",
    26: "=IFERROR(W{r}/U{r},0)", 27: "=X{r}*Y{r}*Z{r}",
    30: "=W{r}*AB{r}", 31: "=AD{r}*(1+AC{r})", 33: "=IFERROR(AD{r}/AF{r},0)",
    35: "=AD{r}*AH{r}", 36: "=AI{r}*AC{r}", 37: "=AI{r}+AJ{r}",
    38: "=(AF{r}-AE{r})*AH{r}", 39: "=AJ{r}+AL{r}", 40: "=IFERROR(AM{r}/AI{r},0)",
    41: "=AF{r}*AH{r}", 44: "=AP{r}*AQ{r}", 47: "=W{r}*AT{r}", 49: "=W{r}*AV{r}",
    50: "=AW{r}-AO{r}-AR{r}-AS{r}-AU{r}", 51: "=IFERROR(AX{r}/P{r},0)",
    52: "=IFERROR((AO{r}+AR{r}+AS{r}+AU{r})/W{r},0)",
}


class BaseDatosPlanta:
    """Acceso y control de la base de datos de eficiencia de produccion."""

    HOJAS = {"lineas": "Dim_Lineas", "turnos": "Dim_Turnos", "sku": "Dim_SKU",
             "motivos": "Dim_Motivos_Paro", "calendario": "Dim_Calendario",
             "prod": "Fact_Produccion", "paros": "Fact_Paros"}

    def __init__(self, ruta: str | Path = ARCHIVO,
                 tablas: dict[str, pd.DataFrame] | None = None):
        if tablas is None:
            self.ruta = Path(ruta)
            if not self.ruta.exists():
                raise FileNotFoundError(f"No existe {self.ruta}")
            libro = pd.read_excel(self.ruta, sheet_name=list(self.HOJAS.values()))
            tablas = {k: libro[v] for k, v in self.HOJAS.items()}
        else:
            self.ruta = Path(ruta) if ruta else None
        for k in self.HOJAS:
            setattr(self, k, tablas[k])
        self.prod["Fecha"] = pd.to_datetime(self.prod["Fecha"])
        self.paros["Fecha"] = pd.to_datetime(self.paros["Fecha"])

    @classmethod
    def desde_csv(cls, carpeta: str | Path = CARPETA_CSV) -> "BaseDatosPlanta":
        """Carga desde el paquete de CSV, unas diez veces mas rapido que el xlsx.

        Genera esa carpeta con: python scripts/preparar_datos.py
        """
        carpeta = Path(carpeta)
        tablas = {k: pd.read_csv(carpeta / f"{v.lower()}.csv") for k, v in cls.HOJAS.items()}
        return cls(ruta=None, tablas=tablas)

    @classmethod
    def cargar(cls, carpeta_csv: str | Path = CARPETA_CSV,
               xlsx: str | Path = ARCHIVO) -> "BaseDatosPlanta":
        """Usa los CSV si existen; si no, cae al archivo de Excel."""
        if Path(carpeta_csv).exists():
            return cls.desde_csv(carpeta_csv)
        return cls(xlsx)

    # ------------------------------------------------------------- validacion
    def validar(self, tolerancia: float = 1e-6) -> pd.DataFrame:
        """Devuelve una tabla de pruebas con su resultado."""
        d = self.prod
        pruebas = []

        def prueba(nombre, ok, detalle=""):
            pruebas.append({"Prueba": nombre,
                            "Resultado": "OK" if ok else "FALLA",
                            "Detalle": detalle})

        prueba("Clave unica por fecha-linea-turno",
               not d.duplicated(["Fecha", "ID_Linea", "ID_Turno"]).any(),
               f"{int(d.duplicated(['Fecha','ID_Linea','ID_Turno']).sum())} duplicados")
        prueba("ID_Registro unico", d["ID_Registro"].is_unique)
        prueba("Sin nulos en campos clave",
               not d[["Fecha", "ID_Linea", "ID_Turno", "ID_SKU",
                      "Cajas_Producidas"]].isna().any().any())
        prueba("ID_Linea existe en Dim_Lineas",
               set(d.ID_Linea) <= set(self.lineas.ID_Linea))
        prueba("ID_SKU existe en Dim_SKU", set(d.ID_SKU) <= set(self.sku.ID_SKU))
        prueba("ID_Registro de paros existe en Fact_Produccion",
               set(self.paros.ID_Registro) <= set(d.ID_Registro))
        prueba("Codigo_Motivo existe en Dim_Motivos_Paro",
               set(self.paros.Codigo_Motivo) <= set(self.motivos.Codigo_Motivo))

        prueba("Horas operativas no negativas", (d.Horas_Operativas >= 0).all(),
               f"minimo {d.Horas_Operativas.min():.3f} h")
        prueba("Paros no exceden el tiempo calendario",
               ((d.Paro_Programado_h + d.Paro_No_Programado_h)
                <= d.Horas_Calendario + tolerancia).all())
        prueba("Cajas rechazadas <= cajas producidas",
               (d.Cajas_Rechazadas <= d.Cajas_Producidas).all())
        prueba("OEE entre 0 y 1", ((d.OEE > 0) & (d.OEE <= 1)).all(),
               f"rango {d.OEE.min():.3f} a {d.OEE.max():.3f}")
        prueba("Rendimiento entre 0.85 y 1",
               ((d.Rendimiento_pct > 0.85) & (d.Rendimiento_pct < 1)).all(),
               f"rango {d.Rendimiento_pct.min():.4f} a {d.Rendimiento_pct.max():.4f}")

        np_detalle = self.paros.loc[self.paros.Tipo == "No programado", "Duracion_h"].sum()
        prueba("Detalle de paros concilia con la tabla de hechos",
               abs(np_detalle - d.Paro_No_Programado_h.sum()) < 0.01,
               f"detalle {np_detalle:,.2f} h contra hechos "
               f"{d.Paro_No_Programado_h.sum():,.2f} h")

        rec = self.kpis_recalculados()
        for col in COLUMNAS_FORMULA:
            dif = (rec[col] - d[col]).abs().max()
            prueba(f"Formula de Excel reproducible: {col}", dif < 0.01,
                   f"diferencia maxima {dif:.6f}")
        return pd.DataFrame(pruebas)

    # ---------------------------------------------------------------- calculo
    def kpis_recalculados(self) -> pd.DataFrame:
        """Reconstruye en pandas cada columna calculada de la hoja."""
        d = self.prod
        r = pd.DataFrame(index=d.index)
        r["Horas_Programadas"] = d.Horas_Calendario - d.Paro_Programado_h
        r["Horas_Operativas"] = r.Horas_Programadas - d.Paro_No_Programado_h
        r["Cajas_Teoricas"] = d.Tasa_Estandar_cajas_h * r.Horas_Operativas
        r["Cajas_Buenas"] = d.Cajas_Producidas - d.Cajas_Rechazadas
        r["Disponibilidad"] = r.Horas_Operativas / r.Horas_Programadas
        r["Desempeno"] = d.Cajas_Producidas / r.Cajas_Teoricas
        r["Calidad"] = r.Cajas_Buenas / d.Cajas_Producidas
        r["OEE"] = r.Disponibilidad * r.Desempeno * r.Calidad
        r["Kg_MP_Perfection"] = r.Cajas_Buenas * d.Kg_Producto_por_Caja
        r["Kg_MP_Estandar"] = r.Kg_MP_Perfection * (1 + d.Factor_Scrap_Estandar)
        r["Rendimiento_pct"] = r.Kg_MP_Perfection / d.Kg_MP_Real
        r["Perfection_Cost_USD"] = r.Kg_MP_Perfection * d.Precio_MP_USD_kg
        r["Costo_Scrap_Estandar_USD"] = r.Perfection_Cost_USD * d.Factor_Scrap_Estandar
        r["Standard_Prime_Cost_USD"] = r.Perfection_Cost_USD + r.Costo_Scrap_Estandar_USD
        r["MUV_USD"] = (d.Kg_MP_Real - r.Kg_MP_Estandar) * d.Precio_MP_USD_kg
        r["Act_Abs_Loss_USD"] = r.Costo_Scrap_Estandar_USD + r.MUV_USD
        r["YTP_pct"] = r.Act_Abs_Loss_USD / r.Perfection_Cost_USD
        r["Costo_MP_Real_USD"] = d.Kg_MP_Real * d.Precio_MP_USD_kg
        r["Costo_MOD_USD"] = d.Horas_MOD * d.Tarifa_MOD_USD_h
        r["Costo_Envase_USD"] = r.Cajas_Buenas * d.Costo_Envase_USD_caja
        r["Ventas_USD"] = r.Cajas_Buenas * d.Precio_Venta_USD_caja
        r["Margen_Contribucion_USD"] = (r.Ventas_USD - r.Costo_MP_Real_USD
                                        - r.Costo_MOD_USD - d.Costo_Energia_USD
                                        - r.Costo_Envase_USD)
        r["MC_por_Hora_Programada_USD"] = r.Margen_Contribucion_USD / r.Horas_Programadas
        r["Costo_Total_por_Caja_USD"] = ((r.Costo_MP_Real_USD + r.Costo_MOD_USD
                                          + d.Costo_Energia_USD + r.Costo_Envase_USD)
                                         / r.Cajas_Buenas)
        return r

    # ----------------------------------------------------------- agregaciones
    def resumen(self, por: str = "linea") -> pd.DataFrame:
        """Agrega la tabla de hechos. por: linea, turno, sku, linea_turno, mes."""
        claves = {"linea": ["ID_Linea"], "turno": ["ID_Turno"], "sku": ["ID_SKU"],
                  "linea_turno": ["ID_Linea", "ID_Turno"], "mes": ["Anio", "Mes"]}
        if por not in claves:
            raise ValueError(f"por debe ser uno de {list(claves)}")
        d = self.prod
        g = d.groupby(claves[por], sort=False).agg(
            Horas_Programadas=("Horas_Programadas", "sum"),
            Horas_Operativas=("Horas_Operativas", "sum"),
            Cajas_Teoricas=("Cajas_Teoricas", "sum"),
            Cajas_Producidas=("Cajas_Producidas", "sum"),
            Cajas_Buenas=("Cajas_Buenas", "sum"),
            Perfection_Cost_USD=("Perfection_Cost_USD", "sum"),
            Act_Abs_Loss_USD=("Act_Abs_Loss_USD", "sum"),
            Kg_MP_Real=("Kg_MP_Real", "sum"),
            Kg_MP_Perfection=("Kg_MP_Perfection", "sum"),
            Ventas_USD=("Ventas_USD", "sum"),
            Margen_Contribucion_USD=("Margen_Contribucion_USD", "sum"),
            MC_por_Hora_sd=("MC_por_Hora_Programada_USD", "std"),
        )
        g["Disponibilidad"] = g.Horas_Operativas / g.Horas_Programadas
        g["Desempeno"] = g.Cajas_Producidas / g.Cajas_Teoricas
        g["Calidad"] = g.Cajas_Buenas / g.Cajas_Producidas
        g["OEE"] = g.Disponibilidad * g.Desempeno * g.Calidad
        g["Rendimiento_pct"] = g.Kg_MP_Perfection / g.Kg_MP_Real
        g["YTP_pct"] = g.Act_Abs_Loss_USD / g.Perfection_Cost_USD
        g["MC_por_Hora_USD"] = g.Margen_Contribucion_USD / g.Horas_Programadas
        return g

    def pareto_paros(self, top: int = 10, tipo: str = "No programado") -> pd.DataFrame:
        """Ordena los motivos de paro por horas perdidas y cajas no producidas."""
        p = self.paros[self.paros.Tipo == tipo]
        g = p.groupby(["Codigo_Motivo", "Motivo", "Categoria_Perdida"]).agg(
            Eventos=("ID_Paro", "count"),
            Horas=("Duracion_h", "sum"),
            Cajas_No_Producidas=("Cajas_No_Producidas", "sum"),
        ).sort_values("Horas", ascending=False)
        g["Pct_Horas"] = g.Horas / g.Horas.sum()
        g["Pct_Acumulado"] = g.Pct_Horas.cumsum()
        return g.head(top)

    # ------------------------------------------------- insumo de optimizacion
    def matriz_retornos(self, metrica: str = "MC_por_Hora_Programada_USD",
                        activo: str = "linea_turno",
                        frecuencia: str = "D") -> pd.DataFrame:
        """Matriz ancha periodo x activo con la metrica elegida.

        activo: linea o linea_turno.
        frecuencia: D diaria, W semanal, ME mensual. Con W o ME la metrica se
        agrega como margen total dividido entre horas programadas totales, no
        como promedio simple de tasas.
        metrica: cualquier columna numerica de Fact_Produccion. Con
        MC_por_Hora_Programada_USD la agregacion es ponderada por horas.

        Esta es la ultima etapa cubierta por este archivo. El modelo de
        optimizacion se construye aparte, sobre el resultado de esta funcion.
        """
        d = self.prod.copy()
        if activo == "linea":
            d["Activo"] = d.ID_Linea
        elif activo == "linea_turno":
            d["Activo"] = d.ID_Linea + "_" + d.ID_Turno
        else:
            raise ValueError("activo debe ser linea o linea_turno")

        if frecuencia == "D":
            d["Periodo"] = d.Fecha
        else:
            d["Periodo"] = d.Fecha.dt.to_period(
                {"W": "W", "ME": "M"}[frecuencia]).dt.to_timestamp(how="end").dt.normalize()

        if metrica == "MC_por_Hora_Programada_USD":
            num = d.pivot_table(index="Periodo", columns="Activo",
                                values="Margen_Contribucion_USD", aggfunc="sum")
            den = d.pivot_table(index="Periodo", columns="Activo",
                                values="Horas_Programadas", aggfunc="sum")
            m = num / den
        else:
            m = d.pivot_table(index="Periodo", columns="Activo",
                              values=metrica, aggfunc="mean")
        return m.sort_index()

    def exportar_matriz(self, destino: str | Path, **kwargs) -> Path:
        destino = Path(destino)
        m = self.matriz_retornos(**kwargs)
        if destino.suffix == ".parquet":
            m.to_parquet(destino)
        else:
            m.to_csv(destino, float_format="%.6f")
        return destino

    # ------------------------------------------------------------------ alta
    def agregar_registros(self, nuevos: pd.DataFrame) -> int:
        """Agrega filas a Fact_Produccion escribiendo las formulas de la hoja.

        nuevos debe traer las columnas de dato (las de color azul). Las columnas
        calculadas se ignoran si vienen: se escriben como formula de Excel.
        Tras usar esta funcion hay que recalcular el archivo en Excel o con
        LibreOffice para que las celdas nuevas tengan valor en cache.
        """
        from openpyxl import load_workbook

        wb = load_workbook(self.ruta)
        ws = wb["Fact_Produccion"]
        encabezados = [c.value for c in ws[1]]
        faltantes = [h for i, h in enumerate(encabezados, start=1)
                     if i not in FORMULAS_EXCEL and h not in nuevos.columns]
        if faltantes:
            raise ValueError(f"Faltan columnas de dato: {faltantes}")

        fila = ws.max_row + 1
        for _, reg in nuevos.iterrows():
            for j, h in enumerate(encabezados, start=1):
                if j in FORMULAS_EXCEL:
                    ws.cell(row=fila, column=j, value=FORMULAS_EXCEL[j].format(r=fila))
                else:
                    ws.cell(row=fila, column=j, value=reg[h])
            fila += 1
        ws.auto_filter.ref = f"A1:AZ{fila - 1}"
        wb.save(self.ruta)
        return len(nuevos)



# ------------------------------------------------------- ubicacion de datos
def localizar_datos(raiz: Path) -> tuple[str | None, Path | None]:
    """Encuentra los datos sin importar como quedo la estructura del repositorio."""
    for carpeta in [raiz / "data" / "csv", raiz / "csv", raiz / "data", raiz]:
        if (carpeta / "fact_produccion.csv").exists():
            return "csv", carpeta
    for patron in ["data/*.xlsx", "*.xlsx", "**/*.xlsx"]:
        for x in sorted(raiz.glob(patron)):
            if not x.name.startswith("~$"):
                return "xlsx", x
    return None, None


def arbol(raiz: Path, limite: int = 60) -> str:
    rutas = [p.relative_to(raiz).as_posix() for p in sorted(raiz.rglob("*"))
             if "__pycache__" not in p.as_posix() and ".git/" not in p.as_posix()]
    extra = f"\n... y {len(rutas) - limite} rutas mas" if len(rutas) > limite else ""
    return "\n".join(rutas[:limite]) + extra


st.set_page_config(page_title="Eficiencia de produccion",
                   page_icon=None, layout="wide")


# ------------------------------------------------------------------ carga
@st.cache_data(show_spinner="Cargando la base de datos")
def cargar_tablas() -> dict[str, pd.DataFrame]:
    tipo, ruta = localizar_datos(RAIZ)
    if tipo is None:
        st.error("No encontre los datos en el repositorio. Falta data/csv o el "
                 "libro BD_Eficiencia_Planta.xlsx. Esto es lo que hay:")
        st.code(arbol(RAIZ))
        st.stop()
    bd = BaseDatosPlanta.desde_csv(ruta) if tipo == "csv" else BaseDatosPlanta(ruta)
    tablas = {k: getattr(bd, k) for k in BaseDatosPlanta.HOJAS}
    tablas["_origen"] = pd.DataFrame({"origen": [str(ruta.relative_to(RAIZ))],
                                      "tipo": [tipo]})
    return tablas


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
origen = tablas.pop("_origen").iloc[0]
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
           f"Fuente: {origen.origen} ({origen.tipo}). "
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
