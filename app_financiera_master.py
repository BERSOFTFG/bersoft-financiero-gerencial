
import re
import math
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Control Financiero Gerencial", layout="wide")

TIPOS_MAPA = {
    "4": "Ingreso",
    "5": "Gasto",
    "6": "Costo",
}

COLUMNAS_BASE = [
    "CodCuentaNivel1",
    "CodCuentaNivel2",
    "CodCuentaNivel3",
    "Cuenta Tercero",
]

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "sept", "octubre", "noviembre", "diciembre",
    "total"
]


def formato_moneda(valor):
    try:
        return f"${float(valor):,.0f}".replace(",", ".")
    except Exception:
        return "$0"


def formato_pct(valor):
    try:
        return f"{float(valor):.1f}%"
    except Exception:
        return "0.0%"


def safe_div(a, b):
    return a / b if b not in (0, None) else 0


def normalizar_columnas(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def extraer_codigo(texto):
    m = re.match(r"^(\d+)", str(texto).strip())
    return m.group(1) if m else ""


def limpiar_nombre(texto):
    return re.sub(r"^\d+\s*", "", str(texto).strip()).strip()


def detectar_hoja_exportable(xls):
    for hoja in xls.sheet_names:
        prueba = pd.read_excel(xls, sheet_name=hoja, nrows=5)
        prueba = normalizar_columnas(prueba)
        if all(c in prueba.columns for c in COLUMNAS_BASE):
            return hoja
    return xls.sheet_names[0]


def detectar_columnas_periodo(df):
    cols = []
    for col in df.columns:
        nombre = str(col).lower()
        if any(m in nombre for m in MESES) and not nombre.startswith("pasypat") and "clas" not in nombre:
            cols.append(col)
    return cols


def transformar_periodo(df, periodo):
    base = normalizar_columnas(df).copy()
    requeridas = COLUMNAS_BASE + [periodo]
    for col in requeridas:
        if col not in base.columns:
            raise ValueError(f"No se encontró la columna requerida: {col}")

    base = base[requeridas].copy()
    base[periodo] = pd.to_numeric(base[periodo], errors="coerce")
    base = base.dropna(subset=[periodo])

    base["CodigoNivel1"] = base["CodCuentaNivel1"].astype(str).apply(extraer_codigo)
    base = base[base["CodigoNivel1"].isin(TIPOS_MAPA.keys())].copy()
    base["Tipo"] = base["CodigoNivel1"].map(TIPOS_MAPA)

    base["Categoria"] = base["CodCuentaNivel2"].astype(str).apply(limpiar_nombre)
    base["Subcategoria"] = base["CodCuentaNivel3"].astype(str).apply(limpiar_nombre)
    base["Cuenta"] = base["Cuenta Tercero"].astype(str).apply(limpiar_nombre)
    base["CodigoCuenta"] = base["Cuenta Tercero"].astype(str).apply(extraer_codigo)
    base["Valor"] = base[periodo].astype(float)
    base["Periodo"] = str(periodo)

    base = base[["Periodo", "Tipo", "Categoria", "Subcategoria", "CodigoCuenta", "Cuenta", "Valor"]].copy()
    return base


def construir_base_todos_periodos(df_raw, columnas_periodo):
    frames = []
    for periodo in columnas_periodo:
        try:
            frames.append(transformar_periodo(df_raw, periodo))
        except Exception:
            pass
    if not frames:
        return pd.DataFrame(columns=["Periodo", "Tipo", "Categoria", "Subcategoria", "CodigoCuenta", "Cuenta", "Valor"])
    return pd.concat(frames, ignore_index=True)


def calcular_metricas(df):
    ingresos = df.loc[df["Tipo"] == "Ingreso", "Valor"].sum()
    costos = df.loc[df["Tipo"] == "Costo", "Valor"].sum()
    gastos = df.loc[df["Tipo"] == "Gasto", "Valor"].sum()
    utilidad_bruta = ingresos - costos
    utilidad_operativa = ingresos - costos - gastos
    utilidad_neta = utilidad_operativa

    margen_bruto = safe_div(utilidad_bruta * 100, ingresos)
    margen_operativo = safe_div(utilidad_operativa * 100, ingresos)
    margen_neto = safe_div(utilidad_neta * 100, ingresos)

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


def calcular_punto_equilibrio(metricas):
    ingresos = metricas["ingresos"]
    costos = metricas["costos"]
    gastos = metricas["gastos"]
    margen_contribucion = safe_div(ingresos - costos, ingresos)
    if margen_contribucion <= 0:
        return None, margen_contribucion
    punto_equilibrio = gastos / margen_contribucion
    return punto_equilibrio, margen_contribucion


def tabla_por_tipo(df, tipo):
    return (
        df[df["Tipo"] == tipo]
        .groupby(["Cuenta", "Subcategoria", "Categoria"], as_index=False)["Valor"]
        .sum()
        .sort_values("Valor", ascending=False)
    )


def resumen_periodos(base_all):
    piv = (
        base_all.groupby(["Periodo", "Tipo"], as_index=False)["Valor"]
        .sum()
        .pivot(index="Periodo", columns="Tipo", values="Valor")
        .fillna(0)
        .reset_index()
    )
    for col in ["Ingreso", "Costo", "Gasto"]:
        if col not in piv.columns:
            piv[col] = 0.0
    piv["Utilidad neta"] = piv["Ingreso"] - piv["Costo"] - piv["Gasto"]
    piv["Margen neto %"] = piv.apply(lambda r: safe_div(r["Utilidad neta"] * 100, r["Ingreso"]), axis=1)
    return piv


def analizar_reduccion_gastos(df_actual, base_all):
    gastos = tabla_por_tipo(df_actual, "Gasto").copy()
    if gastos.empty:
        return gastos

    total_gastos = gastos["Valor"].sum()
    gastos["% del gasto"] = gastos["Valor"].apply(lambda x: safe_div(x * 100, total_gastos))

    hist = (
        base_all[base_all["Tipo"] == "Gasto"]
        .groupby(["Periodo", "Cuenta"], as_index=False)["Valor"]
        .sum()
        .sort_values(["Cuenta", "Periodo"])
    )

    crecimiento = {}
    if not hist.empty:
        for cuenta, sub in hist.groupby("Cuenta"):
            if len(sub) >= 2:
                ult = sub.iloc[-1]["Valor"]
                ant = sub.iloc[-2]["Valor"]
                crecimiento[cuenta] = safe_div((ult - ant) * 100, ant) if ant != 0 else None
            else:
                crecimiento[cuenta] = None

    gastos["Crecimiento vs periodo anterior %"] = gastos["Cuenta"].map(crecimiento)

    def clasificar_prioridad(row):
        pct = row["% del gasto"]
        crec = row["Crecimiento vs periodo anterior %"]
        sub = str(row["Subcategoria"]).upper()
        cuenta = str(row["Cuenta"]).upper()

        no_esencial = any(x in sub or x in cuenta for x in ["VIAJE", "PASAJES", "DIVERSOS", "RESTAURANTE", "HONORARIOS"])
        if pct >= 10 or (crec is not None and crec >= 20) or no_esencial:
            return "Alta"
        if pct >= 5 or (crec is not None and crec >= 10):
            return "Media"
        return "Baja"

    def motivo(row):
        razones = []
        if row["% del gasto"] >= 10:
            razones.append("alto peso")
        if pd.notnull(row["Crecimiento vs periodo anterior %"]) and row["Crecimiento vs periodo anterior %"] >= 20:
            razones.append("crecimiento fuerte")
        texto = f"{row['Subcategoria']} {row['Cuenta']}".upper()
        if any(x in texto for x in ["VIAJE", "PASAJES", "DIVERSOS", "RESTAURANTE", "HONORARIOS"]):
            razones.append("posible gasto ajustable")
        return ", ".join(razones) if razones else "seguimiento"

    gastos["Prioridad de reducción"] = gastos.apply(clasificar_prioridad, axis=1)
    gastos["Motivo"] = gastos.apply(motivo, axis=1)

    cols = ["Cuenta", "Subcategoria", "Categoria", "Valor", "% del gasto", "Crecimiento vs periodo anterior %", "Prioridad de reducción", "Motivo"]
    return gastos[cols].sort_values(["Prioridad de reducción", "Valor"], ascending=[True, False])


def generar_alertas(df_actual, metricas, punto_equilibrio, base_all):
    alertas = []
    ingresos = metricas["ingresos"]
    costos = metricas["costos"]
    gastos = metricas["gastos"]
    utilidad = metricas["utilidad_neta"]

    ingresos_negativos = df_actual[(df_actual["Tipo"] == "Ingreso") & (df_actual["Valor"] < 0)]
    for _, fila in ingresos_negativos.head(5).iterrows():
        alertas.append(f"Ingreso negativo detectado en '{fila['Cuenta']}' por {formato_moneda(fila['Valor'])}. Revisa si es ajuste, nota crédito o error contable.")

    if utilidad < 0:
        alertas.append("La utilidad neta del periodo es negativa.")

    if ingresos > 0:
        porc_gastos = safe_div(gastos * 100, ingresos)
        porc_costos = safe_div(costos * 100, ingresos)
        if porc_gastos > 100:
            alertas.append(f"Los gastos equivalen al {porc_gastos:.1f}% de los ingresos. La estructura de gasto está absorbiendo el negocio.")
        elif porc_gastos > 40:
            alertas.append(f"Los gastos representan {porc_gastos:.1f}% de los ingresos. Es un nivel alto.")
        if porc_costos > 60:
            alertas.append(f"Los costos representan {porc_costos:.1f}% de los ingresos. Revisa rentabilidad operativa.")

        personal = df_actual[df_actual["Cuenta"].str.contains("SUELDOS|SALARIOS|NOMINA|NÓMINA|CESANTIAS|CESANTÍAS|PRIMA|APORTES", case=False, na=False)]["Valor"].sum()
        if personal > 0 and personal / ingresos > 0.30:
            alertas.append(f"Los gastos de personal pesan aproximadamente {personal / ingresos * 100:.1f}% sobre los ingresos.")

        top_ing = tabla_por_tipo(df_actual, "Ingreso")
        if not top_ing.empty and safe_div(top_ing.iloc[0]["Valor"], ingresos) > 0.60:
            alertas.append("Existe alta dependencia de un solo rubro de ingreso.")

    if punto_equilibrio is not None and ingresos < punto_equilibrio:
        alertas.append(f"La facturación actual no alcanza el punto de equilibrio. Faltan aproximadamente {formato_moneda(punto_equilibrio - ingresos)}.")

    resum = resumen_periodos(base_all)
    if len(resum) >= 2:
        ult = resum.iloc[-1]
        ant = resum.iloc[-2]
        if ant["Ingreso"] != 0:
            crec_ing = safe_div((ult["Ingreso"] - ant["Ingreso"]) * 100, ant["Ingreso"])
            if crec_ing < -10:
                alertas.append(f"Los ingresos cayeron {abs(crec_ing):.1f}% frente al periodo anterior.")
        if ant["Gasto"] != 0:
            crec_gas = safe_div((ult["Gasto"] - ant["Gasto"]) * 100, ant["Gasto"])
            if crec_gas > 15:
                alertas.append(f"Los gastos crecieron {crec_gas:.1f}% frente al periodo anterior.")
    return alertas


def generar_recomendaciones(df_actual, metricas, punto_equilibrio, analisis_gastos):
    recomendaciones = []

    top_gastos = tabla_por_tipo(df_actual, "Gasto")
    top_ingresos = tabla_por_tipo(df_actual, "Ingreso")

    if not top_gastos.empty:
        primero = top_gastos.iloc[0]
        recomendaciones.append(f"Revisar el rubro '{primero['Cuenta']}', porque es el gasto más alto del periodo con {formato_moneda(primero['Valor'])}.")

    altas = analisis_gastos[analisis_gastos["Prioridad de reducción"] == "Alta"].head(3)
    for _, row in altas.iterrows():
        recomendaciones.append(f"Prioridad alta para revisar '{row['Cuenta']}' por {formato_moneda(row['Valor'])}: {row['Motivo']}.")

    if metricas["margen_neto"] < 10:
        recomendaciones.append("El margen neto es bajo. Conviene reducir gastos administrativos y fortalecer la facturación recurrente.")
    if metricas["utilidad_neta"] < 0:
        recomendaciones.append("Definir una meta de ventas semanal y mensual y congelar egresos no esenciales hasta recuperar utilidad.")
    if punto_equilibrio is not None:
        recomendaciones.append(f"La empresa debería facturar al menos {formato_moneda(punto_equilibrio)} para cubrir costos y gastos del periodo.")

    if not top_ingresos.empty:
        recomendaciones.append(f"Proteger y potenciar la línea '{top_ingresos.iloc[0]['Cuenta']}', porque hoy es el principal ingreso con {formato_moneda(top_ingresos.iloc[0]['Valor'])}.")
    return recomendaciones


def calcular_metas_proximo_mes(metricas, utilidad_objetivo, crecimiento_ingresos_pct):
    ingresos = metricas["ingresos"]
    costos = metricas["costos"]
    gastos = metricas["gastos"]

    margen_contribucion = safe_div((ingresos - costos), ingresos)
    if margen_contribucion <= 0:
        return {
            "para_no_perder": None,
            "para_utilidad_objetivo": None,
            "meta_crecimiento": None,
        }

    para_no_perder = gastos / margen_contribucion
    para_utilidad_objetivo = (gastos + utilidad_objetivo) / margen_contribucion
    meta_crecimiento = ingresos * (1 + crecimiento_ingresos_pct / 100)

    return {
        "para_no_perder": para_no_perder,
        "para_utilidad_objetivo": para_utilidad_objetivo,
        "meta_crecimiento": meta_crecimiento,
    }


def clasificar_estado(metricas, punto_equilibrio):
    utilidad = metricas["utilidad_neta"]
    ingresos = metricas["ingresos"]
    margen_neto = metricas["margen_neto"]
    if utilidad < 0:
        return "🔴 Crítico", "La empresa presenta pérdida en el periodo analizado."
    if punto_equilibrio is not None and ingresos < punto_equilibrio:
        return "🟡 En riesgo", "La empresa todavía no tiene suficiente holgura frente al punto de equilibrio."
    if margen_neto >= 10:
        return "🟢 Saludable", "La empresa muestra utilidad y una operación más estable."
    return "🟡 Ajustado", "La empresa tiene utilidad, pero el margen sigue siendo apretado."


st.title("Control Financiero Gerencial")
st.caption("Versión única para reemplazar las anteriores: análisis, crecimiento, reducción de gastos y metas de facturación.")

archivo = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])

if archivo is None:
    st.info("Sube tu archivo para iniciar el análisis.")
    st.markdown(
        """
        ### Esta versión incluye
        - análisis por mes y acumulado
        - comparativo entre periodos
        - gráficas de crecimiento
        - dónde deberías reducir gastos
        - metas de facturación para el próximo mes
        - alertas y recomendaciones gerenciales
        """
    )
    st.stop()

try:
    xls = pd.ExcelFile(archivo)
    hoja = detectar_hoja_exportable(xls)
    df_raw = pd.read_excel(xls, sheet_name=hoja)
    df_raw = normalizar_columnas(df_raw)

    columnas_periodo = detectar_columnas_periodo(df_raw)
    if not columnas_periodo:
        st.error("No encontré columnas de periodo para analizar.")
        st.stop()

    base_all = construir_base_todos_periodos(df_raw, columnas_periodo)
    resumen_all = resumen_periodos(base_all)

    st.sidebar.header("Configuración")
    periodo_actual = st.sidebar.selectbox("Periodo actual", columnas_periodo, index=len(columnas_periodo)-1)
    comparar_opciones = ["Sin comparación"] + [c for c in columnas_periodo if c != periodo_actual]
    periodo_comp = st.sidebar.selectbox("Comparar contra", comparar_opciones)
    utilidad_objetivo = st.sidebar.number_input("Utilidad objetivo próximo mes", min_value=0.0, value=10000000.0, step=1000000.0)
    crecimiento_ventas = st.sidebar.number_input("Meta de crecimiento de ventas %", min_value=0.0, value=10.0, step=1.0)

    df_actual = transformar_periodo(df_raw, periodo_actual)
    metricas = calcular_metricas(df_actual)
    punto_equilibrio, _ = calcular_punto_equilibrio(metricas)
    estado, mensaje_estado = clasificar_estado(metricas, punto_equilibrio)
    analisis_gastos = analizar_reduccion_gastos(df_actual, base_all)
    alertas = generar_alertas(df_actual, metricas, punto_equilibrio, base_all)
    recomendaciones = generar_recomendaciones(df_actual, metricas, punto_equilibrio, analisis_gastos)
    metas = calcular_metas_proximo_mes(metricas, utilidad_objetivo, crecimiento_ventas)

    st.subheader(f"Periodo analizado: {periodo_actual}")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ingresos", formato_moneda(metricas["ingresos"]))
    c2.metric("Costos", formato_moneda(metricas["costos"]))
    c3.metric("Gastos", formato_moneda(metricas["gastos"]))
    c4.metric("Utilidad neta", formato_moneda(metricas["utilidad_neta"]))

    st.markdown("---")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Margen bruto", formato_pct(metricas["margen_bruto"]))
    c6.metric("Margen operativo", formato_pct(metricas["margen_operativo"]))
    c7.metric("Margen neto", formato_pct(metricas["margen_neto"]))
    c8.metric("Punto de equilibrio", formato_moneda(punto_equilibrio) if punto_equilibrio is not None else "No calculable")

    st.subheader("Estado general")
    st.markdown(f"### {estado}")
    st.write(mensaje_estado)

    st.subheader("Metas de facturación para el próximo mes")
    m1, m2, m3 = st.columns(3)
    m1.metric("Para no perder", formato_moneda(metas["para_no_perder"]) if metas["para_no_perder"] is not None else "No calculable")
    m2.metric("Para lograr utilidad objetivo", formato_moneda(metas["para_utilidad_objetivo"]) if metas["para_utilidad_objetivo"] is not None else "No calculable")
    m3.metric("Meta por crecimiento comercial", formato_moneda(metas["meta_crecimiento"]) if metas["meta_crecimiento"] is not None else "No calculable")

    if periodo_comp != "Sin comparación":
        df_comp = transformar_periodo(df_raw, periodo_comp)
        metricas_comp = calcular_metricas(df_comp)

        def delta(a, b):
            return a - b

        st.subheader(f"Comparativo: {periodo_actual} vs {periodo_comp}")
        cc1, cc2, cc3, cc4 = st.columns(4)
        cc1.metric("Δ Ingresos", formato_moneda(metricas["ingresos"]), formato_moneda(delta(metricas["ingresos"], metricas_comp["ingresos"])))
        cc2.metric("Δ Costos", formato_moneda(metricas["costos"]), formato_moneda(delta(metricas["costos"], metricas_comp["costos"])))
        cc3.metric("Δ Gastos", formato_moneda(metricas["gastos"]), formato_moneda(delta(metricas["gastos"], metricas_comp["gastos"])))
        cc4.metric("Δ Utilidad", formato_moneda(metricas["utilidad_neta"]), formato_moneda(delta(metricas["utilidad_neta"], metricas_comp["utilidad_neta"])))

    st.subheader("Gráficas de crecimiento")
    chart_df = resumen_all.copy().set_index("Periodo")[["Ingreso", "Costo", "Gasto", "Utilidad neta"]]
    st.line_chart(chart_df)
    st.bar_chart(chart_df[["Ingreso", "Gasto", "Utilidad neta"]])

    st.subheader("Indicadores de tendencia")
    tendencia = resumen_all.copy()
    tendencia["Crecimiento ingresos %"] = tendencia["Ingreso"].pct_change() * 100
    tendencia["Crecimiento gastos %"] = tendencia["Gasto"].pct_change() * 100
    tendencia["Crecimiento utilidad %"] = tendencia["Utilidad neta"].pct_change() * 100
    mostrar_tend = tendencia[["Periodo", "Ingreso", "Gasto", "Utilidad neta", "Margen neto %", "Crecimiento ingresos %", "Crecimiento gastos %", "Crecimiento utilidad %"]].copy()
    for col in ["Ingreso", "Gasto", "Utilidad neta"]:
        mostrar_tend[col] = mostrar_tend[col].apply(formato_moneda)
    for col in ["Margen neto %", "Crecimiento ingresos %", "Crecimiento gastos %", "Crecimiento utilidad %"]:
        mostrar_tend[col] = mostrar_tend[col].apply(lambda x: formato_pct(x) if pd.notnull(x) else "N/A")
    st.dataframe(mostrar_tend, use_container_width=True)

    st.subheader("Dónde deberías reducir gastos")
    mostrar_red = analisis_gastos.copy()
    if not mostrar_red.empty:
        mostrar_red["Valor"] = mostrar_red["Valor"].apply(formato_moneda)
        mostrar_red["% del gasto"] = mostrar_red["% del gasto"].apply(formato_pct)
        mostrar_red["Crecimiento vs periodo anterior %"] = mostrar_red["Crecimiento vs periodo anterior %"].apply(lambda x: formato_pct(x) if pd.notnull(x) else "N/A")
        st.dataframe(mostrar_red.head(20), use_container_width=True)
    else:
        st.info("No hay gastos suficientes para analizar reducción.")

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Rubros con mayor gasto")
        top_gastos = tabla_por_tipo(df_actual, "Gasto").head(15).copy()
        if not top_gastos.empty:
            top_gastos["Valor"] = top_gastos["Valor"].apply(formato_moneda)
        st.dataframe(top_gastos, use_container_width=True)

    with col_b:
        st.subheader("Rubros con mayor ingreso")
        top_ing = tabla_por_tipo(df_actual, "Ingreso").head(15).copy()
        if not top_ing.empty:
            top_ing["Valor"] = top_ing["Valor"].apply(formato_moneda)
        st.dataframe(top_ing, use_container_width=True)

    st.subheader("Alertas automáticas")
    if alertas:
        for alerta in alertas:
            st.warning(alerta)
    else:
        st.success("No se detectaron alertas críticas en el periodo seleccionado.")

    st.subheader("Recomendaciones gerenciales")
    for i, rec in enumerate(recomendaciones, start=1):
        st.write(f"{i}. {rec}")

    with st.expander("Ver base transformada"):
        st.dataframe(df_actual, use_container_width=True)

except Exception as e:
    st.error(f"Ocurrió un error al procesar el archivo: {e}")
