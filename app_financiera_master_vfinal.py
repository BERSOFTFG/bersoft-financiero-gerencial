import io
import re
from datetime import datetime
 
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    PageBreak,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER
 
 
st.set_page_config(page_title="Control Financiero Inteligente PRO", layout="wide")
 
MESES_ORDEN = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12
}
 
 
def formato_moneda(valor):
    try:
        return f"${valor:,.0f}"
    except Exception:
        return "$0"
 
 
def formato_pct(valor):
    try:
        return f"{valor:.1f}%"
    except Exception:
        return "0.0%"
 
 
def semaforo_html(nivel, texto):
    color = {
        "verde": "#D1FAE5",
        "amarillo": "#FEF3C7",
        "rojo": "#FEE2E2"
    }.get(nivel, "#E5E7EB")
    borde = {
        "verde": "#10B981",
        "amarillo": "#F59E0B",
        "rojo": "#EF4444"
    }.get(nivel, "#9CA3AF")
    return f"""
    <div style="padding:10px 12px;border-left:6px solid {borde};background:{color};
    border-radius:8px;margin-bottom:8px;">{texto}</div>
    """
 
 
def detectar_hoja_bersoft(archivo):
    hojas = pd.ExcelFile(archivo).sheet_names
    if "ExportarAExcel" in hojas:
        return "ExportarAExcel"
    if "Tabla" in hojas:
        return "Tabla"
    return hojas[0]
 
 
def detectar_columnas_mes(df):
    columnas_mes = []
    for c in df.columns:
        c_txt = str(c).strip().lower()
        patron = re.search(
            r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\s+(\d{4})",
            c_txt
        )
        if patron:
            mes = patron.group(1)
            anio = int(patron.group(2))
            columnas_mes.append((c, mes, anio, MESES_ORDEN[mes]))
    return columnas_mes
 
 
def leer_excel_bersoft(archivo, nombre_origen="Archivo"):
    hoja = detectar_hoja_bersoft(archivo)
    df = pd.read_excel(archivo, sheet_name=hoja)
    df.columns = [str(c).strip() for c in df.columns]
 
    columnas_requeridas = ["CodCuentaNivel1", "CodCuentaNivel2", "CodCuentaNivel3", "Cuenta Tercero"]
    for col in columnas_requeridas:
        if col not in df.columns:
            df[col] = ""
 
    columnas_mes = detectar_columnas_mes(df)
    if not columnas_mes:
        raise ValueError("No se encontraron columnas con formato tipo 'Enero 2026'.")
 
    filas = []
    for col_original, mes_nombre, anio, mes_num in columnas_mes:
        temp = df[["CodCuentaNivel1", "CodCuentaNivel2", "CodCuentaNivel3", "Cuenta Tercero", col_original]].copy()
        temp = temp.rename(columns={col_original: "Valor"})
        temp["Valor"] = pd.to_numeric(temp["Valor"], errors="coerce").fillna(0)
        temp["MesNombre"] = mes_nombre.capitalize()
        temp["MesNum"] = mes_num
        temp["Año"] = anio
        temp["Periodo"] = f"{mes_nombre.capitalize()} {anio}"
        temp["Origen"] = nombre_origen
        filas.append(temp)
 
    data = pd.concat(filas, ignore_index=True)
    data["CodCuentaNivel1"] = data["CodCuentaNivel1"].astype(str).str.strip()
    data["CodCuentaNivel2"] = data["CodCuentaNivel2"].astype(str).str.strip()
    data["CodCuentaNivel3"] = data["CodCuentaNivel3"].astype(str).str.strip()
    data["Cuenta Tercero"] = data["Cuenta Tercero"].astype(str).str.strip()
 
    data["Tipo"] = data["CodCuentaNivel1"].str[0].map({
        "4": "Ingresos",
        "5": "Gastos",
        "6": "Costos"
    }).fillna("Otros")
 
    data["CategoriaGasto"] = data["CodCuentaNivel2"].replace("", "Sin categoría")
    data["ProductoServicio"] = data["Cuenta Tercero"].replace("", "Sin detalle")
    return data
 
 
def resumen_periodo(data, anio, meses):
    base = data[(data["Año"] == anio) & (data["MesNum"].isin(meses))].copy()
 
    ingresos = base.loc[base["Tipo"] == "Ingresos", "Valor"].sum()
    gastos = base.loc[base["Tipo"] == "Gastos", "Valor"].sum()
    costos = base.loc[base["Tipo"] == "Costos", "Valor"].sum()
 
    utilidad_bruta = ingresos - costos
    utilidad_operativa = ingresos - costos - gastos
    impuestos_estimados = utilidad_operativa * 0.35 if utilidad_operativa > 0 else 0
    utilidad_neta = utilidad_operativa - impuestos_estimados
 
    margen_valor = utilidad_neta
    margen_pct = (utilidad_neta / ingresos * 100) if ingresos else 0
 
    gastos_fijos = gastos
    margen_contribucion = ((ingresos - costos) / ingresos) if ingresos else 0
    punto_equilibrio = (gastos_fijos / margen_contribucion) if margen_contribucion > 0 else 0
    falta_facturar = max(0, punto_equilibrio - ingresos)
 
    return {
        "ingresos": float(ingresos),
        "gastos": float(gastos),
        "costos": float(costos),
        "utilidad_bruta": float(utilidad_bruta),
        "utilidad_operativa": float(utilidad_operativa),
        "utilidad_neta": float(utilidad_neta),
        "impuestos_estimados": float(impuestos_estimados),
        "margen_valor": float(margen_valor),
        "margen_pct": float(margen_pct),
        "punto_equilibrio": float(punto_equilibrio),
        "falta_facturar": float(falta_facturar),
    }
 
 
def comparar_valores(valor_base, valor_comp):
    dif = valor_base - valor_comp
    pct = (dif / valor_comp * 100) if valor_comp else 0
    return dif, pct
 
 
def tabla_comparativa(resumen_base, resumen_comp):
    filas = []
    for etiqueta, clave in [
        ("Ingresos", "ingresos"),
        ("Costos", "costos"),
        ("Gastos", "gastos"),
        ("Utilidad bruta", "utilidad_bruta"),
        ("Utilidad operativa", "utilidad_operativa"),
        ("Utilidad neta", "utilidad_neta"),
    ]:
        dif, pct = comparar_valores(resumen_base[clave], resumen_comp[clave])
        filas.append({
            "Rubro": etiqueta,
            "Periodo base": resumen_base[clave],
            "Periodo comparativo": resumen_comp[clave],
            "Variación $": dif,
            "Variación %": pct,
        })
    return pd.DataFrame(filas)
 
 
def ranking_ingresos(data, anio, meses):
    base = data[(data["Año"] == anio) & (data["MesNum"].isin(meses)) & (data["Tipo"] == "Ingresos")]
    rank = (
        base.groupby("ProductoServicio", as_index=False)["Valor"]
        .sum()
        .sort_values("Valor", ascending=False)
    )
    return rank.head(10)
 
 
def gastos_categoria(data, anio, meses):
    base = data[(data["Año"] == anio) & (data["MesNum"].isin(meses)) & (data["Tipo"] == "Gastos")]
    gasto_cat = (
        base.groupby("CategoriaGasto", as_index=False)["Valor"]
        .sum()
        .sort_values("Valor", ascending=False)
    )
    return gasto_cat
 
 
def serie_mensual(data):
    tabla = (
        data[data["Tipo"].isin(["Ingresos", "Costos", "Gastos"])]
        .groupby(["Año", "MesNum", "MesNombre", "Tipo"], as_index=False)["Valor"]
        .sum()
    )
    piv = tabla.pivot_table(index=["Año", "MesNum", "MesNombre"], columns="Tipo", values="Valor", fill_value=0).reset_index()
    for col in ["Ingresos", "Costos", "Gastos"]:
        if col not in piv.columns:
            piv[col] = 0
    piv["Utilidad"] = piv["Ingresos"] - piv["Costos"] - piv["Gastos"]
    piv = piv.sort_values(["Año", "MesNum"])
    piv["Etiqueta"] = piv["MesNombre"] + " " + piv["Año"].astype(str)
    return piv
 
 
def proyeccion_cierre_anual(data, anio):
    piv = serie_mensual(data)
    año_data = piv[piv["Año"] == anio].copy()
    meses_cargados = len(año_data)
    if meses_cargados == 0:
        return {"normal": 0, "optimista": 0, "pesimista": 0}
    promedio_utilidad = año_data["Utilidad"].mean()
    utilidad_actual = año_data["Utilidad"].sum()
    faltan = max(0, 12 - meses_cargados)
 
    normal = utilidad_actual + promedio_utilidad * faltan
    optimista = utilidad_actual + promedio_utilidad * 1.15 * faltan
    pesimista = utilidad_actual + promedio_utilidad * 0.85 * faltan
    return {"normal": normal, "optimista": optimista, "pesimista": pesimista}
 
 
def generar_alertas(resumen, gasto_cat, comparativo_df=None):
    alertas = []
    if resumen["utilidad_neta"] < 0:
        alertas.append(("rojo", "La utilidad neta es negativa. Hay pérdidas en el periodo analizado."))
    elif resumen["margen_pct"] < 5:
        alertas.append(("amarillo", "El margen de utilidad es bajo. Conviene revisar costos y gastos."))
    else:
        alertas.append(("verde", "La utilidad y el margen muestran una condición saludable del negocio."))
 
    if resumen["gastos"] > resumen["ingresos"] * 0.30:
        alertas.append(("amarillo", "Los gastos superan el 30% de los ingresos. Se recomienda control administrativo."))
 
    if not gasto_cat.empty:
        gasto_max = gasto_cat.iloc[0]
        alertas.append(("amarillo", f"La categoría con mayor gasto es {gasto_max['CategoriaGasto']} por {formato_moneda(gasto_max['Valor'])}."))
 
    if comparativo_df is not None and not comparativo_df.empty:
        fila_ing = comparativo_df[comparativo_df["Rubro"] == "Ingresos"].iloc[0]
        fila_gas = comparativo_df[comparativo_df["Rubro"] == "Gastos"].iloc[0]
        if fila_ing["Variación %"] < -10:
            alertas.append(("rojo", "Los ingresos cayeron más de 10% frente al periodo comparado."))
        if fila_gas["Variación %"] > 15:
            alertas.append(("rojo", "Los gastos subieron más de 15% frente al periodo comparado."))
    return alertas
 
 
def conclusiones_automaticas(resumen, ranking, gasto_cat):
    top_ingreso = ranking.iloc[0]["ProductoServicio"] if not ranking.empty else "Sin información"
    top_gasto = gasto_cat.iloc[0]["CategoriaGasto"] if not gasto_cat.empty else "Sin información"
 
    texto = []
    texto.append(f"Los ingresos del periodo fueron {formato_moneda(resumen['ingresos'])} y la utilidad neta estimada fue {formato_moneda(resumen['utilidad_neta'])}.")
    texto.append(f"El margen neto se ubicó en {formato_pct(resumen['margen_pct'])}.")
    texto.append(f"El producto o servicio con mayor aporte fue {top_ingreso}.")
    texto.append(f"La categoría de gasto más representativa fue {top_gasto}.")
    if resumen["utilidad_neta"] < 0:
        texto.append("Se recomienda revisar gastos de rápida contención y fortalecer los servicios con mejor aporte.")
    elif resumen["margen_pct"] < 5:
        texto.append("La empresa genera utilidad, pero el margen es estrecho; conviene fortalecer ingresos y contener gastos fijos.")
    else:
        texto.append("El resultado del periodo es favorable y permite proyectar un cierre de año positivo si se mantiene la tendencia.")
    return " ".join(texto)
 
 
def crear_grafico_ingresos_vs_gastos(piv):
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(piv["Etiqueta"], piv["Ingresos"], marker="o", label="Ingresos")
    ax.plot(piv["Etiqueta"], piv["Gastos"], marker="o", label="Gastos")
    ax.plot(piv["Etiqueta"], piv["Costos"], marker="o", label="Costos")
    ax.legend()
    ax.set_title("Ingresos vs gastos por mes")
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer
 
 
def crear_grafico_utilidad(piv):
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.bar(piv["Etiqueta"], piv["Utilidad"])
    ax.set_title("Utilidad por mes")
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer
 
 
def crear_grafico_participacion(rank):
    top = rank.head(8).copy()
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.pie(top["Valor"], labels=top["ProductoServicio"], autopct="%1.1f%%")
    ax.set_title("Participación por producto o servicio")
    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer
 
 
def exportar_excel_resultados(df_comp, ranking, gasto_cat):
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        df_comp.to_excel(writer, sheet_name="Comparativo", index=False)
        ranking.to_excel(writer, sheet_name="Ranking Ingresos", index=False)
        gasto_cat.to_excel(writer, sheet_name="Gastos Categoria", index=False)
    salida.seek(0)
    return salida
 
 
def generar_pdf(periodo_base_txt, periodo_comp_txt, resumen_base, resumen_comp, df_comp, ranking, gasto_cat, alertas, conclusion, img1, img2, img3):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    estilos = getSampleStyleSheet()
    estilos.add(ParagraphStyle(name="TituloBersoft", parent=estilos["Heading1"], fontSize=22, leading=26, alignment=TA_CENTER, textColor=colors.HexColor("#1F3A5F")))
    estilos.add(ParagraphStyle(name="Sub", parent=estilos["Normal"], fontSize=10, leading=13, alignment=TA_CENTER, textColor=colors.HexColor("#5B6470")))
    estilos.add(ParagraphStyle(name="Caja", parent=estilos["Normal"], fontSize=10, leading=14, alignment=TA_LEFT))
 
    story = []
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("BERSOFT SAS", estilos["TituloBersoft"]))
    story.append(Paragraph("Informe Financiero Gerencial", estilos["TituloBersoft"]))
    story.append(Paragraph(f"Periodo base: {periodo_base_txt} | Periodo comparativo: {periodo_comp_txt}", estilos["Sub"]))
    story.append(Paragraph(f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilos["Sub"]))
    story.append(Spacer(1, 0.5*cm))
 
    resumen_data = [
        ["Indicador", "Periodo base", "Periodo comparativo"],
        ["Ingresos", formato_moneda(resumen_base["ingresos"]), formato_moneda(resumen_comp["ingresos"])],
        ["Costos", formato_moneda(resumen_base["costos"]), formato_moneda(resumen_comp["costos"])],
        ["Gastos", formato_moneda(resumen_base["gastos"]), formato_moneda(resumen_comp["gastos"])],
        ["Utilidad neta", formato_moneda(resumen_base["utilidad_neta"]), formato_moneda(resumen_comp["utilidad_neta"])],
        ["Margen neto", formato_pct(resumen_base["margen_pct"]), formato_pct(resumen_comp["margen_pct"])],
    ]
    tabla_res = Table(resumen_data, colWidths=[6*cm, 5*cm, 5*cm])
    tabla_res.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F3A5F")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(tabla_res)
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("<b>Resumen ejecutivo</b>", estilos["Heading2"]))
    story.append(Paragraph(conclusion, estilos["Caja"]))
    story.append(Spacer(1, 0.3*cm))
 
    story.append(Paragraph("<b>Alertas gerenciales</b>", estilos["Heading2"]))
    for nivel, txt in alertas:
        color = {"verde": "#10B981", "amarillo": "#F59E0B", "rojo": "#EF4444"}.get(nivel, "#9CA3AF")
        story.append(Paragraph(f'<font color="{color}">●</font> {txt}', estilos["Caja"]))
    story.append(Spacer(1, 0.4*cm))
 
    comp_table = [["Rubro", "Base", "Comparativo", "Variación $", "Variación %"]]
    for _, row in df_comp.iterrows():
        comp_table.append([
            row["Rubro"],
            formato_moneda(row["Periodo base"]),
            formato_moneda(row["Periodo comparativo"]),
            formato_moneda(row["Variación $"]),
            formato_pct(row["Variación %"]),
        ])
    t2 = Table(comp_table, colWidths=[4.5*cm, 3.5*cm, 3.5*cm, 3.5*cm, 3*cm])
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F3A5F")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
        ("TOPPADDING", (0,0), (-1,-1), 7),
    ]))
    story.append(t2)
    story.append(PageBreak())
 
    story.append(Paragraph("<b>Gráficos de gestión</b>", estilos["Heading2"]))
    story.append(Spacer(1, 0.2*cm))
    story.append(RLImage(img1, width=17*cm, height=7*cm))
    story.append(Spacer(1, 0.3*cm))
    story.append(RLImage(img2, width=17*cm, height=7*cm))
    story.append(PageBreak())
    story.append(Paragraph("<b>Participación por producto o servicio</b>", estilos["Heading2"]))
    story.append(Spacer(1, 0.2*cm))
    story.append(RLImage(img3, width=14*cm, height=9*cm))
 
    if not ranking.empty:
        story.append(Spacer(1, 0.4*cm))
        story.append(Paragraph("<b>Ranking de ingresos</b>", estilos["Heading2"]))
        rank_table = [["Producto/Servicio", "Ingreso"]]
        for _, row in ranking.head(10).iterrows():
            rank_table.append([str(row["ProductoServicio"])[:55], formato_moneda(row["Valor"])])
        t3 = Table(rank_table, colWidths=[12*cm, 5*cm])
        t3.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F3A5F")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
            ("BOTTOMPADDING", (0,0), (-1,-1), 7),
            ("TOPPADDING", (0,0), (-1,-1), 7),
        ]))
        story.append(t3)
 
    doc.build(story)
    buffer.seek(0)
    return buffer
 
 
st.title("Control Financiero Inteligente PRO")
st.caption("Versión gerencial para Bersoft | Lectura automática del Excel sin organizarlo manualmente")
 
col_up1, col_up2 = st.columns(2)
with col_up1:
    archivo_base = st.file_uploader("Archivo base", type=["xlsx"], key="base")
with col_up2:
    archivo_comp = st.file_uploader("Archivo comparativo", type=["xlsx"], key="comp")
 
if archivo_base is not None:
    try:
        data_base = leer_excel_bersoft(archivo_base, "Base")
    except Exception as e:
        st.error(f"Error en archivo base: {e}")
        st.stop()
 
    data_total = data_base.copy()
    if archivo_comp is not None:
        try:
            data_comp_arch = leer_excel_bersoft(archivo_comp, "Comparativo")
            data_total = pd.concat([data_base, data_comp_arch], ignore_index=True)
        except Exception as e:
            st.error(f"Error en archivo comparativo: {e}")
            st.stop()
 
    años_disponibles = sorted(data_total["Año"].dropna().unique().tolist())
    if not años_disponibles:
        st.warning("No se detectaron años en los archivos.")
        st.stop()
 
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        año_base = st.selectbox("Año base", años_disponibles, index=len(años_disponibles)-1)
    with c2:
        año_comp = st.selectbox("Año comparativo", años_disponibles, index=max(0, len(años_disponibles)-2) if len(años_disponibles) > 1 else 0)
    with c3:
        modo = st.selectbox("Modo de análisis", ["Mes a mes", "Acumulado", "Periodo personalizado"])
    with c4:
        mes_max = st.selectbox("Mes final", list(range(1, 13)), index=0, format_func=lambda x: list(MESES_ORDEN.keys())[x-1].capitalize())
 
    if modo == "Mes a mes":
        meses_base = [mes_max]
        meses_comp = [mes_max]
        periodo_base_txt = f"{list(MESES_ORDEN.keys())[mes_max-1].capitalize()} {año_base}"
        periodo_comp_txt = f"{list(MESES_ORDEN.keys())[mes_max-1].capitalize()} {año_comp}"
    elif modo == "Acumulado":
        meses_base = list(range(1, mes_max + 1))
        meses_comp = list(range(1, mes_max + 1))
        periodo_base_txt = f"Acumulado enero a {list(MESES_ORDEN.keys())[mes_max-1].capitalize()} {año_base}"
        periodo_comp_txt = f"Acumulado enero a {list(MESES_ORDEN.keys())[mes_max-1].capitalize()} {año_comp}"
    else:
        mes_inicio = st.selectbox("Mes inicial", list(range(1, 13)), index=0, format_func=lambda x: list(MESES_ORDEN.keys())[x-1].capitalize())
        if mes_inicio > mes_max:
            st.warning("El mes inicial no puede ser mayor al mes final.")
            st.stop()
        meses_base = list(range(mes_inicio, mes_max + 1))
        meses_comp = list(range(mes_inicio, mes_max + 1))
        periodo_base_txt = f"Periodo {list(MESES_ORDEN.keys())[mes_inicio-1].capitalize()} a {list(MESES_ORDEN.keys())[mes_max-1].capitalize()} {año_base}"
        periodo_comp_txt = f"Periodo {list(MESES_ORDEN.keys())[mes_inicio-1].capitalize()} a {list(MESES_ORDEN.keys())[mes_max-1].capitalize()} {año_comp}"
 
    resumen_base = resumen_periodo(data_total, año_base, meses_base)
    resumen_comp = resumen_periodo(data_total, año_comp, meses_comp)
 
    df_comp = tabla_comparativa(resumen_base, resumen_comp)
    ranking = ranking_ingresos(data_total, año_base, meses_base)
    gasto_cat = gastos_categoria(data_total, año_base, meses_base)
    piv = serie_mensual(data_total)
    piv_año = piv[piv["Año"] == año_base]
 
    alertas = generar_alertas(resumen_base, gasto_cat, df_comp)
    conclusion = conclusiones_automaticas(resumen_base, ranking, gasto_cat)
    proy = proyeccion_cierre_anual(data_total, año_base)
 
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Resumen gerencial", "Comparativos", "Ingresos", "Gastos", "Utilidad y proyección", "Reportes"
    ])
 
    with tab1:
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Ingresos", formato_moneda(resumen_base["ingresos"]))
        k2.metric("Costos", formato_moneda(resumen_base["costos"]))
        k3.metric("Gastos", formato_moneda(resumen_base["gastos"]))
        k4.metric("Utilidad neta", formato_moneda(resumen_base["utilidad_neta"]))
 
        k5, k6, k7, k8 = st.columns(4)
        k5.metric("Margen neto", formato_pct(resumen_base["margen_pct"]))
        k6.metric("Punto de equilibrio", formato_moneda(resumen_base["punto_equilibrio"]))
        k7.metric("Falta por facturar", formato_moneda(resumen_base["falta_facturar"]))
        k8.metric("Proyección cierre año", formato_moneda(proy["normal"]))
 
        st.subheader("Semáforos y alertas")
        for nivel, txt in alertas:
            st.markdown(semaforo_html(nivel, txt), unsafe_allow_html=True)
 
        st.subheader("Conclusión automática")
        st.info(conclusion)
 
        st.subheader("Gráficos principales")
        img1 = crear_grafico_ingresos_vs_gastos(piv_año)
        st.image(img1)
        img2 = crear_grafico_utilidad(piv_año)
        st.image(img2)
 
    with tab2:
        st.subheader("Comparativo del periodo")
        st.dataframe(df_comp, use_container_width=True)
 
        st.subheader("Comparativo resumido en texto")
        for _, fila in df_comp.iterrows():
            signo = "subió" if fila["Variación $"] > 0 else "bajó"
            st.write(f"- {fila['Rubro']}: {signo} {formato_moneda(abs(fila['Variación $']))} ({formato_pct(abs(fila['Variación %']))}).")
 
    with tab3:
        st.subheader("Ranking de productos o servicios más rentables")
        st.dataframe(ranking, use_container_width=True)
        img3 = crear_grafico_participacion(ranking if not ranking.empty else pd.DataFrame({"ProductoServicio":["Sin datos"],"Valor":[1]}))
        st.image(img3)
 
        st.subheader("Ingresos mensuales")
        st.dataframe(piv_año[["Etiqueta", "Ingresos"]], use_container_width=True)
 
        st.metric("Ingresos acumulados del año", formato_moneda(piv_año["Ingresos"].sum()))
 
    with tab4:
        st.subheader("Gastos por categoría")
        st.dataframe(gasto_cat, use_container_width=True)
        if not gasto_cat.empty:
            mayor_gasto = gasto_cat.iloc[0]
            st.warning(f"La categoría con mayor gasto es {mayor_gasto['CategoriaGasto']} por {formato_moneda(mayor_gasto['Valor'])}.")
 
    with tab5:
        u1, u2, u3 = st.columns(3)
        u1.metric("Utilidad bruta", formato_moneda(resumen_base["utilidad_bruta"]))
        u2.metric("Utilidad operativa", formato_moneda(resumen_base["utilidad_operativa"]))
        u3.metric("Utilidad neta", formato_moneda(resumen_base["utilidad_neta"]))
 
        p1, p2, p3 = st.columns(3)
        p1.metric("Escenario pesimista", formato_moneda(proy["pesimista"]))
        p2.metric("Escenario normal", formato_moneda(proy["normal"]))
        p3.metric("Escenario optimista", formato_moneda(proy["optimista"]))
 
        st.line_chart(piv_año.set_index("Etiqueta")[["Utilidad"]])
 
    with tab6:
        st.subheader("Descargas gerenciales")
        xlsx_buffer = exportar_excel_resultados(df_comp, ranking, gasto_cat)
 
        img1 = crear_grafico_ingresos_vs_gastos(piv_año)
        img2 = crear_grafico_utilidad(piv_año)
        img3 = crear_grafico_participacion(ranking if not ranking.empty else pd.DataFrame({"ProductoServicio":["Sin datos"],"Valor":[1]}))
 
        pdf_buffer = generar_pdf(
            periodo_base_txt, periodo_comp_txt,
            resumen_base, resumen_comp, df_comp, ranking, gasto_cat,
            alertas, conclusion, img1, img2, img3
        )
 
        st.download_button(
            "Descargar informe PDF",
            data=pdf_buffer,
            file_name="informe_financiero_gerencial_bersoft.pdf",
            mime="application/pdf"
        )
 
        st.download_button(
            "Descargar resultados en Excel",
            data=xlsx_buffer,
            file_name="resultados_financieros_bersoft.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Sube al menos un archivo Excel para iniciar el análisis.")
