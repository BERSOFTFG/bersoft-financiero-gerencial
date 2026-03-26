import streamlit as st
import pandas as pd
import numpy as np
import re

st.set_page_config(page_title="Control Financiero Inteligente", layout="wide")

st.title("Control Financiero Inteligente")
st.caption("Sube tu estado financiero en Excel y obtén análisis, alertas y recomendaciones gerenciales.")

TIPOS_MAPA = {
    "4": "Ingreso",
    "5": "Gasto",
    "6": "Costo",
}

COLUMNAS_EXPORT = [
    "CodCuentaNivel1",
    "CodCuentaNivel2",
    "CodCuentaNivel3",
    "Cuenta Tercero",
]


def formato_moneda(valor):
    try:
        return f"${float(valor):,.0f}".replace(",", ".")
    except Exception:
        return "$0"


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def extraer_codigo(texto: str) -> str:
    match = re.match(r"^(\d+)", str(texto).strip())
    return match.group(1) if match else ""


def limpiar_nombre_cuenta(texto: str) -> str:
    texto = str(texto).strip()
    return re.sub(r"^\d+\s*", "", texto).strip()


def detectar_columna_valor(df: pd.DataFrame) -> str | None:
    candidatos = []
    for col in df.columns:
        nombre = str(col).lower()
        if "enero" in nombre or "febrero" in nombre or "marzo" in nombre or "abril" in nombre or "mayo" in nombre or "junio" in nombre or "julio" in nombre or "agosto" in nombre or "sept" in nombre or "oct" in nombre or "nov" in nombre or "dic" in nombre:
            candidatos.append(col)
    if candidatos:
        return candidatos[0]

    for col in df.columns:
        if "valtot" in str(col).lower() or "total" in str(col).lower():
            return col
    return None


def detectar_hoja_exportable(xls: pd.ExcelFile) -> str:
    for hoja in xls.sheet_names:
        prueba = pd.read_excel(xls, sheet_name=hoja, nrows=5)
        prueba = normalizar_columnas(prueba)
        if all(c in prueba.columns for c in COLUMNAS_EXPORT):
            return hoja
    return xls.sheet_names[0]


def transformar_excel_contable(df: pd.DataFrame, nombre_col_valor: str) -> pd.DataFrame:
    df = normalizar_columnas(df)
    base = df.copy()

    for col in COLUMNAS_EXPORT + [nombre_col_valor]:
        if col not in base.columns:
            raise ValueError(f"No se encontró la columna requerida: {col}")

    base = base[COLUMNAS_EXPORT + [nombre_col_valor]].copy()
    base[nombre_col_valor] = pd.to_numeric(base[nombre_col_valor], errors="coerce")
    base = base.dropna(subset=[nombre_col_valor])

    base["CodigoNivel1"] = base["CodCuentaNivel1"].astype(str).apply(extraer_codigo)
    base = base[base["CodigoNivel1"].isin(TIPOS_MAPA.keys())].copy()
    base["Tipo"] = base["CodigoNivel1"].map(TIPOS_MAPA)

    base["Categoria"] = base["CodCuentaNivel2"].astype(str).apply(limpiar_nombre_cuenta)
    base["Subcategoria"] = base["CodCuentaNivel3"].astype(str).apply(limpiar_nombre_cuenta)
    base["Cuenta"] = base["Cuenta Tercero"].astype(str).apply(limpiar_nombre_cuenta)
    base["CodigoCuenta"] = base["Cuenta Tercero"].astype(str).apply(extraer_codigo)
    base["Valor"] = base[nombre_col_valor].astype(float)

    base["Periodo"] = str(nombre_col_valor)
    base["Fecha"] = pd.to_datetime("today").normalize()

    base = base[[
        "Fecha",
        "Periodo",
        "Tipo",
        "Categoria",
        "Subcategoria",
        "CodigoCuenta",
        "Cuenta",
        "Valor",
    ]].copy()

    return base


def calcular_metricas(df: pd.DataFrame) -> dict:
    ingresos = df.loc[df["Tipo"] == "Ingreso", "Valor"].sum()
    costos = df.loc[df["Tipo"] == "Costo", "Valor"].sum()
    gastos = df.loc[df["Tipo"] == "Gasto", "Valor"].sum()

    utilidad_bruta = ingresos - costos
    utilidad_operativa = ingresos - costos - gastos
    utilidad_neta = utilidad_operativa

    margen_bruto = (utilidad_bruta / ingresos * 100) if ingresos else 0
    margen_operativo = (utilidad_operativa / ingresos * 100) if ingresos else 0
    margen_neto = (utilidad_neta / ingresos * 100) if ingresos else 0

    return {
        "ingresos": ingresos,
        "costos": costos,
        "gastos": gastos,
        "utilidad_bruta": utilidad_bruta,
        "utilidad_operativa": utilidad_operativa,
        "utilidad_neta": utilidad_neta,
        "margen_bruto": margen_bruto,
        "margen_operativo": margen_operativo,
        "margen_neto": margen_neto,
    }


def calcular_punto_equilibrio(metricas: dict):
    ingresos = metricas["ingresos"]
    costos = metricas["costos"]
    gastos = metricas["gastos"]

    margen_contribucion = ((ingresos - costos) / ingresos) if ingresos else 0
    if margen_contribucion <= 0:
        return None, margen_contribucion

    punto_equilibrio = gastos / margen_contribucion
    return punto_equilibrio, margen_contribucion


def clasificar_estado(metricas: dict, punto_equilibrio):
    utilidad = metricas["utilidad_neta"]
    ingresos = metricas["ingresos"]
    margen_neto = metricas["margen_neto"]

    if utilidad < 0:
        return "🔴 Mal", "La empresa presenta pérdida en el periodo analizado."
    if punto_equilibrio is not None and ingresos < punto_equilibrio:
        return "🟡 Riesgo", "La empresa no está en pérdida, pero aún no tiene suficiente holgura frente al punto de equilibrio."
    if margen_neto >= 10:
        return "🟢 Bien", "La empresa muestra utilidad y un comportamiento financiero saludable."
    return "🟡 Riesgo", "La empresa tiene utilidad, pero el margen es todavía ajustado."


def generar_alertas(df: pd.DataFrame, metricas: dict, punto_equilibrio):
    alertas = []

    ingresos_negativos = df[(df["Tipo"] == "Ingreso") & (df["Valor"] < 0)]
    if not ingresos_negativos.empty:
        for _, fila in ingresos_negativos.head(5).iterrows():
            alertas.append(
                f"Se detectó un ingreso negativo en '{fila['Cuenta']}' por {formato_moneda(fila['Valor'])}. Conviene revisar si es ajuste, nota crédito o error de registro."
            )

    if metricas["utilidad_neta"] < 0:
        alertas.append("La utilidad neta del periodo es negativa.")

    if metricas["ingresos"] > 0:
        porcentaje_gastos = metricas["gastos"] / metricas["ingresos"] * 100
        if porcentaje_gastos > 40:
            alertas.append(f"Los gastos representan {porcentaje_gastos:.1f}% de los ingresos. Es un nivel alto.")

        sueldos = df[df["Cuenta"].str.contains("SUELDOS", case=False, na=False)]["Valor"].sum()
        if sueldos / metricas["ingresos"] > 0.30:
            alertas.append(f"El rubro de sueldos equivale a {sueldos / metricas['ingresos'] * 100:.1f}% de los ingresos.")

    if punto_equilibrio is not None and metricas["ingresos"] < punto_equilibrio:
        faltante = punto_equilibrio - metricas["ingresos"]
        alertas.append(f"La facturación actual no alcanza el punto de equilibrio. Faltan aproximadamente {formato_moneda(faltante)}.")

    return alertas


def generar_recomendaciones(df: pd.DataFrame, metricas: dict, punto_equilibrio):
    recomendaciones = []

    top_gastos = (
        df[df["Tipo"] == "Gasto"]
        .groupby("Cuenta", as_index=False)["Valor"]
        .sum()
        .sort_values("Valor", ascending=False)
    )

    if not top_gastos.empty:
        primero = top_gastos.iloc[0]
        recomendaciones.append(
            f"Revisar el rubro '{primero['Cuenta']}', porque es el gasto más alto del periodo con {formato_moneda(primero['Valor'])}."
        )

    if metricas["margen_neto"] < 10:
        recomendaciones.append("El margen neto es bajo. Conviene revisar gastos administrativos y fortalecer la facturación recurrente.")

    if metricas["utilidad_neta"] < 0:
        recomendaciones.append("Priorizar control de egresos no esenciales y definir una meta mínima de facturación semanal y mensual.")

    if punto_equilibrio is not None:
        recomendaciones.append(f"La empresa debería facturar al menos {formato_moneda(punto_equilibrio)} para cubrir costos y gastos del periodo.")

    ingresos = metricas["ingresos"]
    if ingresos > 0:
        top_ingresos = (
            df[df["Tipo"] == "Ingreso"]
            .groupby("Cuenta", as_index=False)["Valor"]
            .sum()
            .sort_values("Valor", ascending=False)
        )
        if not top_ingresos.empty:
            recomendaciones.append(
                f"El principal ingreso del periodo es '{top_ingresos.iloc[0]['Cuenta']}'. Conviene proteger esa línea y revisar por qué otras líneas aportan menos."
            )

    if not recomendaciones:
        recomendaciones.append("La estructura financiera luce estable. Mantén seguimiento mensual y metas de gasto por rubro.")

    return recomendaciones


st.sidebar.header("Carga de archivo")
archivo = st.sidebar.file_uploader("Sube tu Excel", type=["xlsx", "xls"])

if archivo is None:
    st.info("Sube tu archivo de Excel para iniciar el análisis.")
    st.markdown(
        """
        ### Esta app está pensada para estados financieros exportados desde contabilidad
        Puede leer archivos con estructura similar a:
        - CodCuentaNivel1
        - CodCuentaNivel2
        - CodCuentaNivel3
        - Cuenta Tercero
        - una columna de periodo, por ejemplo: Enero 2026

        ### Qué te entrega
        - Ingresos, costos, gastos y utilidad
        - Estado financiero general
        - Punto de equilibrio
        - Rubros donde más estás gastando
        - Alertas de inconsistencias
        - Recomendaciones gerenciales
        """
    )
else:
    try:
        xls = pd.ExcelFile(archivo)
        hoja = detectar_hoja_exportable(xls)
        df_raw = pd.read_excel(xls, sheet_name=hoja)
        df_raw = normalizar_columnas(df_raw)

        columna_valor = detectar_columna_valor(df_raw)
        if columna_valor is None:
            st.error("No pude detectar la columna del periodo o del valor a analizar.")
            st.stop()

        df = transformar_excel_contable(df_raw, columna_valor)
        metricas = calcular_metricas(df)
        punto_equilibrio, margen_contribucion = calcular_punto_equilibrio(metricas)
        estado, mensaje_estado = clasificar_estado(metricas, punto_equilibrio)
        alertas = generar_alertas(df, metricas, punto_equilibrio)
        recomendaciones = generar_recomendaciones(df, metricas, punto_equilibrio)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Ingresos", formato_moneda(metricas["ingresos"]))
        col2.metric("Costos", formato_moneda(metricas["costos"]))
        col3.metric("Gastos", formato_moneda(metricas["gastos"]))
        col4.metric("Utilidad neta", formato_moneda(metricas["utilidad_neta"]))

        st.markdown("---")
        col5, col6, col7, col8 = st.columns(4)
        col5.metric("Margen bruto", f"{metricas['margen_bruto']:.1f}%")
        col6.metric("Margen operativo", f"{metricas['margen_operativo']:.1f}%")
        col7.metric("Margen neto", f"{metricas['margen_neto']:.1f}%")
        col8.metric(
            "Punto de equilibrio",
            formato_moneda(punto_equilibrio) if punto_equilibrio is not None else "No calculable",
        )

        st.subheader("Estado general")
        st.markdown(f"### {estado}")
        st.write(mensaje_estado)

        st.subheader("Distribución general")
        resumen_tipo = df.groupby("Tipo", as_index=False)["Valor"].sum()
        st.dataframe(resumen_tipo, use_container_width=True)
        st.bar_chart(resumen_tipo.set_index("Tipo")["Valor"])

        st.subheader("Rubros con mayor gasto")
        top_gastos = (
            df[df["Tipo"] == "Gasto"][["Cuenta", "Subcategoria", "Categoria", "Valor"]]
            .sort_values("Valor", ascending=False)
            .head(15)
        )
        st.dataframe(top_gastos, use_container_width=True)

        st.subheader("Rubros con mayor ingreso")
        top_ingresos = (
            df[df["Tipo"] == "Ingreso"][["Cuenta", "Subcategoria", "Categoria", "Valor"]]
            .sort_values("Valor", ascending=False)
            .head(15)
        )
        st.dataframe(top_ingresos, use_container_width=True)

        st.subheader("Alertas automáticas")
        if alertas:
            for alerta in alertas:
                st.warning(alerta)
        else:
            st.success("No se detectaron alertas críticas en el archivo analizado.")

        st.subheader("Recomendaciones")
        for rec in recomendaciones:
            st.write(f"- {rec}")

        st.subheader("Base transformada para análisis")
        st.dataframe(df, use_container_width=True)

        with st.expander("Ver reglas utilizadas por la app"):
            st.markdown(
                """
                **Reglas actuales**
                - 4 = Ingreso
                - 5 = Gasto
                - 6 = Costo
                - Punto de equilibrio = Gastos / margen de contribución
                - Se detectan ingresos negativos como alerta
                - Se resaltan los rubros de mayor peso

                **Siguiente mejora recomendada**
                - Separar gastos fijos y variables
                - Comparar varios meses
                - Crear meta de facturación mensual y semanal
                - Agregar tablero ejecutivo por áreas
                """
            )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
