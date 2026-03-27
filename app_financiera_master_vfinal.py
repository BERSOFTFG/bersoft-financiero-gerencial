import io
import re
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image as RLImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

st.set_page_config(page_title="Control Financiero Inteligente PRO", layout="wide")

MESES_ORDEN = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}
MESES_NOMBRE = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def formato_moneda(valor):
    try:
        return f"${float(valor):,.0f}"
    except Exception:
        return "$0"



def formato_pct(valor):
    try:
        return f"{float(valor):.1f}%"
    except Exception:
        return "0.0%"



def semaforo_html(nivel, texto):
    color = {
        "verde": "#D1FAE5",
        "amarillo": "#FEF3C7",
        "rojo": "#FEE2E2",
    }.get(nivel, "#E5E7EB")
    borde = {
        "verde": "#10B981",
        "amarillo": "#F59E0B",
        "rojo": "#EF4444",
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
            c_txt,
        )
        if patron:
            mes = patron.group(1)
            anio = int(patron.group(2))
            columnas_mes.append((str(c).strip(), mes, anio, MESES_ORDEN[mes]))
    return columnas_mes



def leer_excel_bersoft(archivo, nombre_origen="Archivo"):
    hoja = detectar_hoja_bersoft(archivo)
    df = pd.read_excel(archivo, sheet_name=hoja)
    df.columns = [str(c).strip() for c in df.columns]

    for col in ["CodCuentaNivel1", "CodCuentaNivel2", "CodCuentaNivel3", "Cuenta Tercero"]:
        if col not in df.columns:
            df[col] = ""

    columnas_mes = detectar_columnas_mes(df)
    if not columnas_mes:
        raise ValueError("No se encontraron columnas con formato tipo 'Enero 2026'.")

    filas = []
    for col_original, mes_nombre, anio, mes_num in columnas_mes:
        temp = df[
            ["CodCuentaNivel1", "CodCuentaNivel2", "CodCuentaNivel3", "Cuenta Tercero", col_original]
        ].copy()
        temp = temp.rename(columns={col_original: "Valor"})
        temp["Valor"] = pd.to_numeric(temp["Valor"], errors="coerce").fillna(0)
        temp["MesNombre"] = mes_nombre.capitalize()
        temp["MesNum"] = mes_num
        temp["Año"] = anio
        temp["Periodo"] = f"{mes_nombre.capitalize()} {anio}"
        temp["Origen"] = nombre_origen
        filas.append(temp)

    data = pd.concat(filas, ignore_index=True)
    for col in ["CodCuentaNivel1", "CodCuentaNivel2", "CodCuentaNivel3", "Cuenta Tercero"]:
        data[col] = data[col].astype(str).str.strip()

    data["Tipo"] = data["CodCuentaNivel1"].str[0].map(
        {"4": "Ingresos", "5": "Gastos", "6": "Costos"}
    ).fillna("Otros")
    data["CategoriaGasto"] = data["CodCuentaNivel2"].replace("", "Sin categoría")
    data["ProductoServicio"] = data["Cuenta Tercero"].replace("", "Sin detalle")
    return data



def obtener_anios_disponibles(data):
    return sorted(data["Año"].dropna().astype(int).unique().tolist())



def obtener_periodo_texto(modo, mes_inicio, mes_final, anio):
    mes_ini_txt = MESES_NOMBRE[mes_inicio]
    mes_fin_txt = MESES_NOMBRE[mes_final]
    if modo == "Mes a mes":
        return f"{mes_fin_txt} {anio}"
    if modo == "Acumulado":
        return f"Acumulado Enero a {mes_fin_txt} {anio}"
    return f"Periodo {mes_ini_txt} a {mes_fin_txt} {anio}"



def construir_meses(modo, mes_inicio, mes_final):
    if modo == "Mes a mes":
        return [mes_final]
    return list(range(mes_inicio, mes_final + 1))



def resumen_periodo(data, anio, meses, utilidad_objetivo=0.0):
    base = data[(data["Año"] == anio) & (data["MesNum"].isin(meses))].copy()

    ingresos = float(base.loc[base["Tipo"] == "Ingresos", "Valor"].sum())
    gastos = float(base.loc[base["Tipo"] == "Gastos", "Valor"].sum())
    costos = float(base.loc[base["Tipo"] == "Costos", "Valor"].sum())

    utilidad_bruta = ingresos - costos
    utilidad_operativa = ingresos - costos - gastos
    impuestos_estimados = utilidad_operativa * 0.35 if utilidad_operativa > 0 else 0
    utilidad_neta = utilidad_operativa - impuestos_estimados

    margen_neto_pct = (utilidad_neta / ingresos * 100) if ingresos else 0.0
    margen_contribucion_ratio = ((ingresos - costos) / ingresos) if ingresos else 0.0
    punto_equilibrio = (gastos / margen_contribucion_ratio) if margen_contribucion_ratio > 0 else 0.0

    facturacion_total_objetivo = (
        (gastos + utilidad_objetivo) / margen_contribucion_ratio
        if margen_contribucion_ratio > 0
        else 0.0
    )
    facturacion_adicional_objetivo = max(0.0, facturacion_total_objetivo - ingresos)

    return {
        "ingresos": ingresos,
        "gastos": gastos,
        "costos": costos,
        "utilidad_bruta": utilidad_bruta,
        "utilidad_operativa": utilidad_operativa,
        "utilidad_neta": utilidad_neta,
        "impuestos_estimados": impuestos_estimados,
        "margen_neto_pct": margen_neto_pct,
        "margen_valor": utilidad_neta,
        "margen_contribucion_ratio": margen_contribucion_ratio,
        "punto_equilibrio": punto_equilibrio,
        "facturacion_total_objetivo": facturacion_total_objetivo,
        "facturacion_adicional_objetivo": facturacion_adicional_objetivo,
    }



def comparar_valores(valor_base, valor_comp):
    diferencia = valor_base - valor_comp
    porcentaje = ((diferencia / abs(valor_comp)) * 100) if valor_comp else 0.0
    return diferencia, porcentaje



def tabla_comparativa(resumen_base, resumen_comp, anio_base, anio_comp):
    filas = []
    col_base = f"Periodo base ({anio_base})"
    col_comp = f"Periodo comparativo ({anio_comp})"
    for etiqueta, clave in [
        ("Ingresos", "ingresos"),
        ("Costos", "costos"),
        ("Gastos", "gastos"),
        ("Utilidad bruta", "utilidad_bruta"),
        ("Utilidad operativa", "utilidad_operativa"),
        ("Utilidad neta", "utilidad_neta"),
    ]:
        dif, pct = comparar_valores(resumen_base[clave], resumen_comp[clave])
        filas.append(
            {
                "Rubro": etiqueta,
                col_base: resumen_base[clave],
                col_comp: resumen_comp[clave],
                "Variación $": dif,
                "Variación %": pct,
            }
        )
    return pd.DataFrame(filas)



def ranking_ingresos(data, anio, meses):
    base = data[
        (data["Año"] == anio) & (data["MesNum"].isin(meses)) & (data["Tipo"] == "Ingresos")
    ].copy()
    rank = (
        base.groupby("ProductoServicio", as_index=False)["Valor"]
        .sum()
        .sort_values("Valor", ascending=False)
    )
    return rank



def gastos_categoria(data, anio, meses):
    base = data[
        (data["Año"] == anio) & (data["MesNum"].isin(meses)) & (data["Tipo"] == "Gastos")
    ].copy()
    gasto_cat = (
        base.groupby("CategoriaGasto", as_index=False)["Valor"]
        .sum()
        .sort_values("Valor", ascending=False)
    )
    return gasto_cat



def serie_mensual(data, anio=None):
    tabla = (
        data[data["Tipo"].isin(["Ingresos", "Costos", "Gastos"])]
        .groupby(["Año", "MesNum", "MesNombre", "Tipo"], as_index=False)["Valor"]
        .sum()
    )
    piv = (
        tabla.pivot_table(
            index=["Año", "MesNum", "MesNombre"], columns="Tipo", values="Valor", fill_value=0
        )
        .reset_index()
        .sort_values(["Año", "MesNum"])
    )
    for col in ["Ingresos", "Costos", "Gastos"]:
        if col not in piv.columns:
            piv[col] = 0.0
    piv["Utilidad"] = piv["Ingresos"] - piv["Costos"] - piv["Gastos"]
    piv["Etiqueta"] = piv["MesNombre"] + " " + piv["Año"].astype(str)
    if anio is not None:
        piv = piv[piv["Año"] == anio].copy()
    return piv



def proyeccion_cierre_anual(data, anio):
    piv = serie_mensual(data, anio)
    meses_cargados = len(piv)
    if meses_cargados == 0:
        return {"normal": 0.0, "optimista": 0.0, "pesimista": 0.0}
    promedio_utilidad = piv["Utilidad"].mean()
    utilidad_actual = piv["Utilidad"].sum()
    faltan = max(0, 12 - meses_cargados)
    return {
        "normal": utilidad_actual + promedio_utilidad * faltan,
        "optimista": utilidad_actual + promedio_utilidad * 1.15 * faltan,
        "pesimista": utilidad_actual + promedio_utilidad * 0.85 * faltan,
    }



def generar_alertas(resumen, gasto_cat, comparativo_df=None):
    alertas = []
    if resumen["utilidad_neta"] < 0:
        alertas.append(
            ("rojo", "La utilidad neta es negativa. El periodo presenta pérdida y requiere revisión inmediata.")
        )
    elif resumen["margen_neto_pct"] < 5:
        alertas.append(
            ("amarillo", "La utilidad es positiva, pero el margen neto es bajo. Conviene revisar costos y gastos.")
        )
    else:
        alertas.append(("verde", "La utilidad y el margen muestran una condición saludable del negocio."))

    if resumen["gastos"] > resumen["ingresos"] * 0.30:
        alertas.append(
            ("amarillo", "Los gastos superan el 30% de los ingresos. Se recomienda control administrativo.")
        )

    if not gasto_cat.empty:
        gasto_max = gasto_cat.iloc[0]
        alertas.append(
            (
                "amarillo",
                f"La categoría con mayor gasto es {gasto_max['CategoriaGasto']} por {formato_moneda(gasto_max['Valor'])}.",
            )
        )

    if comparativo_df is not None and not comparativo_df.empty:
        fila_ing = comparativo_df[comparativo_df["Rubro"] == "Ingresos"].iloc[0]
        fila_gas = comparativo_df[comparativo_df["Rubro"] == "Gastos"].iloc[0]
        if fila_ing["Variación %"] < -10:
            alertas.append(("rojo", "Los ingresos cayeron más de 10% frente al periodo comparado."))
        if fila_gas["Variación %"] > 15:
            alertas.append(("rojo", "Los gastos subieron más de 15% frente al periodo comparado."))
    return alertas



def conclusiones_automaticas(resumen, ranking, gasto_cat, periodo_txt):
    top_ingreso = ranking.iloc[0]["ProductoServicio"] if not ranking.empty else "Sin información"
    top_gasto = gasto_cat.iloc[0]["CategoriaGasto"] if not gasto_cat.empty else "Sin información"

    texto = []
    texto.append(
        f"En {periodo_txt}, los ingresos fueron {formato_moneda(resumen['ingresos'])} y la utilidad neta estimada fue {formato_moneda(resumen['utilidad_neta'])}."
    )
    texto.append(f"El margen neto se ubicó en {formato_pct(resumen['margen_neto_pct'])}.")
    texto.append(f"El producto o servicio con mayor aporte fue {top_ingreso}.")
    texto.append(f"La categoría de gasto más representativa fue {top_gasto}.")
    if resumen["utilidad_neta"] < 0:
        texto.append(
            "La empresa presenta pérdida neta en el periodo analizado. Los gastos y costos están superando los ingresos."
        )
    elif resumen["margen_neto_pct"] < 5:
        texto.append(
            "La empresa genera utilidad, pero el margen es estrecho; conviene fortalecer ingresos y contener gastos."
        )
    else:
        texto.append(
            "El resultado del periodo es favorable y permite proyectar un cierre de año positivo si se mantiene la tendencia."
        )
    return " ".join(texto)



def crear_grafico_ingresos_vs_gastos(piv, anio):
    fig, ax = plt.subplots(figsize=(11, 4.5))
    ax.plot(piv["Etiqueta"], piv["Ingresos"], marker="o", label="Ingresos")
    ax.plot(piv["Etiqueta"], piv["Gastos"], marker="o", label="Gastos")
    ax.plot(piv["Etiqueta"], piv["Costos"], marker="o", label="Costos")
    ax.set_title(f"Ingresos, gastos y costos por mes - {anio}")
    ax.tick_params(axis="x", rotation=45)
    ax.legend()
    for _, row in piv.iterrows():
        ax.annotate(
            formato_moneda(row["Gastos"]),
            (row["Etiqueta"], row["Gastos"]),
            textcoords="offset points",
            xytext=(0, 8),
            ha="center",
            fontsize=7,
        )
    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer



def crear_grafico_utilidad(piv, anio):
    fig, ax = plt.subplots(figsize=(11, 4.5))
    colores = ["green" if v >= 0 else "red" for v in piv["Utilidad"]]
    barras = ax.bar(piv["Etiqueta"], piv["Utilidad"], color=colores)
    ax.set_title(f"Utilidad por mes - {anio}")
    ax.tick_params(axis="x", rotation=45)
    for barra, valor in zip(barras, piv["Utilidad"]):
        ax.text(
            barra.get_x() + barra.get_width() / 2,
            valor,
            formato_moneda(valor),
            ha="center",
            va="bottom" if valor >= 0 else "top",
            fontsize=7,
        )
    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer



def crear_grafico_participacion(rank, periodo_txt):
    fig, ax = plt.subplots(figsize=(10, 6))

    if rank is None or rank.empty:
        ax.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        ax.axis("off")
    else:
        top = rank.copy()
        top["Valor"] = pd.to_numeric(top["Valor"], errors="coerce").fillna(0)
        top = top[top["Valor"] > 0].sort_values("Valor", ascending=False)

        if top.empty:
            ax.text(0.5, 0.5, "No hay valores positivos para graficar", ha="center", va="center")
            ax.axis("off")
        else:
            top = top.head(6).copy()
            resto = rank.copy()
            resto["Valor"] = pd.to_numeric(resto["Valor"], errors="coerce").fillna(0)
            resto = resto[resto["Valor"] > 0].sort_values("Valor", ascending=False).iloc[6:]
            if not resto.empty:
                otros_valor = resto["Valor"].sum()
                if otros_valor > 0:
                    top = pd.concat(
                        [top, pd.DataFrame([{"ProductoServicio": "Otros", "Valor": otros_valor}])],
                        ignore_index=True,
                    )

            wedges, _, _ = ax.pie(
                top["Valor"],
                labels=None,
                autopct="%1.1f%%",
                startangle=90,
                pctdistance=0.7,
            )
            ax.legend(
                wedges,
                top["ProductoServicio"].astype(str),
                title="Producto / Servicio",
                loc="center left",
                bbox_to_anchor=(1.0, 0.5),
                fontsize=9,
            )
            ax.set_title(f"Participación por producto o servicio - {periodo_txt}")

    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer



def exportar_excel_resultados(df_comp, ranking, gasto_cat, piv_año):
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        df_comp.to_excel(writer, sheet_name="Comparativo", index=False)
        ranking.to_excel(writer, sheet_name="Ranking Ingresos", index=False)
        gasto_cat.to_excel(writer, sheet_name="Gastos Categoria", index=False)
        piv_año.to_excel(writer, sheet_name="Series Mensuales", index=False)
    salida.seek(0)
    return salida



def generar_pdf(
    periodo_base_txt,
    periodo_comp_txt,
    resumen_base,
    resumen_comp,
    df_comp,
    ranking,
    gasto_cat,
    alertas,
    conclusion,
    img1,
    img2,
    img3,
    anio_base,
    anio_comp,
):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )
    estilos = getSampleStyleSheet()
    estilos.add(
        ParagraphStyle(
            name="TituloBersoft",
            parent=estilos["Heading1"],
            fontSize=22,
            leading=26,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#1F3A5F"),
        )
    )
    estilos.add(
        ParagraphStyle(
            name="Sub",
            parent=estilos["Normal"],
            fontSize=10,
            leading=13,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#5B6470"),
        )
    )
    estilos.add(
        ParagraphStyle(name="Caja", parent=estilos["Normal"], fontSize=10, leading=14, alignment=TA_LEFT)
    )

    story = []
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("BERSOFT SAS", estilos["TituloBersoft"]))
    story.append(Paragraph("Informe Financiero Gerencial", estilos["TituloBersoft"]))
    story.append(
        Paragraph(
            f"Periodo base: {periodo_base_txt} | Periodo comparativo: {periodo_comp_txt}",
            estilos["Sub"],
        )
    )
    story.append(Paragraph(f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilos["Sub"]))
    story.append(Spacer(1, 0.5 * cm))

    resumen_data = [
        ["Indicador", f"Periodo base ({anio_base})", f"Periodo comparativo ({anio_comp})"],
        ["Ingresos", formato_moneda(resumen_base["ingresos"]), formato_moneda(resumen_comp["ingresos"])],
        ["Costos", formato_moneda(resumen_base["costos"]), formato_moneda(resumen_comp["costos"])],
        ["Gastos", formato_moneda(resumen_base["gastos"]), formato_moneda(resumen_comp["gastos"])],
        ["Utilidad neta", formato_moneda(resumen_base["utilidad_neta"]), formato_moneda(resumen_comp["utilidad_neta"])],
        ["Margen neto", formato_pct(resumen_base["margen_neto_pct"]), formato_pct(resumen_comp["margen_neto_pct"])],
    ]
    tabla_res = Table(resumen_data, colWidths=[6 * cm, 5 * cm, 5 * cm])
    tabla_res.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(tabla_res)
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("<b>Resumen ejecutivo</b>", estilos["Heading2"]))
    story.append(Paragraph(conclusion, estilos["Caja"]))
    story.append(Spacer(1, 0.3 * cm))

    story.append(Paragraph("<b>Alertas gerenciales</b>", estilos["Heading2"]))
    for nivel, txt in alertas:
        color = {"verde": "#10B981", "amarillo": "#F59E0B", "rojo": "#EF4444"}.get(nivel, "#9CA3AF")
        story.append(Paragraph(f'<font color="{color}">●</font> {txt}', estilos["Caja"]))
    story.append(Spacer(1, 0.4 * cm))

    col_base = f"Periodo base ({anio_base})"
    col_comp = f"Periodo comparativo ({anio_comp})"
    comp_table = [["Rubro", col_base, col_comp, "Variación $", "Variación %"]]
    for _, row in df_comp.iterrows():
        comp_table.append(
            [
                row["Rubro"],
                formato_moneda(row[col_base]),
                formato_moneda(row[col_comp]),
                formato_moneda(row["Variación $"]),
                formato_pct(row["Variación %"]),
            ]
        )
    t2 = Table(comp_table, colWidths=[4.5 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm, 3 * cm])
    t2.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(t2)
    story.append(PageBreak())

    story.append(Paragraph("<b>Gráficos de gestión</b>", estilos["Heading2"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(RLImage(img1, width=17 * cm, height=7 * cm))
    story.append(Spacer(1, 0.3 * cm))
    story.append(RLImage(img2, width=17 * cm, height=7 * cm))
    story.append(PageBreak())
    story.append(Paragraph("<b>Participación por producto o servicio</b>", estilos["Heading2"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(RLImage(img3, width=14 * cm, height=9 * cm))

    if not ranking.empty:
        story.append(Spacer(1, 0.4 * cm))
        story.append(Paragraph("<b>Ranking de ingresos</b>", estilos["Heading2"]))
        rank_table = [["Producto/Servicio", "Ingreso"]]
        for _, row in ranking.head(10).iterrows():
            rank_table.append([str(row["ProductoServicio"])[:55], formato_moneda(row["Valor"])])
        t3 = Table(rank_table, colWidths=[12 * cm, 5 * cm])
        t3.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                    ("TOPPADDING", (0, 0), (-1, -1), 7),
                ]
            )
        )
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

    data_comp_arch = None
    if archivo_comp is not None:
        try:
            data_comp_arch = leer_excel_bersoft(archivo_comp, "Comparativo")
        except Exception as e:
            st.error(f"Error en archivo comparativo: {e}")
            st.stop()

    anios_base = obtener_anios_disponibles(data_base)
    anios_comp = obtener_anios_disponibles(data_comp_arch) if data_comp_arch is not None else anios_base

    if not anios_base:
        st.warning("No se detectaron años en el archivo base.")
        st.stop()

    anio_base_default = max(anios_base)
    if anios_comp:
        posibles_comp = [a for a in anios_comp if a != anio_base_default]
        anio_comp_default = max(posibles_comp) if posibles_comp else anios_comp[0]
    else:
        anio_comp_default = anio_base_default - 1

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        anio_base = st.selectbox(
            "Año base",
            anios_base,
            index=anios_base.index(anio_base_default),
        )
    with c2:
        anio_comp = st.selectbox(
            "Año comparativo",
            anios_comp,
            index=anios_comp.index(anio_comp_default) if anio_comp_default in anios_comp else 0,
        )
    with c3:
        modo = st.selectbox("Modo de análisis", ["Mes a mes", "Acumulado", "Periodo personalizado"], index=1)
    with c4:
        mes_max = st.selectbox(
            "Mes final",
            list(range(1, 13)),
            index=4,
            format_func=lambda x: MESES_NOMBRE[x],
        )

    mes_inicio = 1
    if modo == "Periodo personalizado":
        mes_inicio = st.selectbox(
            "Mes inicial",
            list(range(1, 13)),
            index=0,
            format_func=lambda x: MESES_NOMBRE[x],
        )
        if mes_inicio > mes_max:
            st.warning("El mes inicial no puede ser mayor al mes final.")
            st.stop()

    utilidad_objetivo = st.number_input(
        "Utilidad objetivo deseada",
        min_value=0.0,
        value=50000000.0,
        step=1000000.0,
    )

    meses_base = construir_meses(modo, mes_inicio, mes_max)
    meses_comp = construir_meses(modo, mes_inicio, mes_max)

    periodo_base_txt = obtener_periodo_texto(modo, mes_inicio, mes_max, anio_base)
    periodo_comp_txt = obtener_periodo_texto(modo, mes_inicio, mes_max, anio_comp)

    resumen_base = resumen_periodo(data_base, anio_base, meses_base, utilidad_objetivo)
    fuente_comp = data_comp_arch if data_comp_arch is not None else data_base
    resumen_comp = resumen_periodo(fuente_comp, anio_comp, meses_comp, utilidad_objetivo)

    df_comp = tabla_comparativa(resumen_base, resumen_comp, anio_base, anio_comp)
    ranking = ranking_ingresos(data_base, anio_base, meses_base)
    gasto_cat = gastos_categoria(data_base, anio_base, meses_base)
    piv_año = serie_mensual(data_base, anio_base)

    alertas = generar_alertas(resumen_base, gasto_cat, df_comp)
    conclusion = conclusiones_automaticas(resumen_base, ranking, gasto_cat, periodo_base_txt)
    proy = proyeccion_cierre_anual(data_base, anio_base)

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        ["Resumen gerencial", "Comparativos", "Ingresos", "Gastos", "Utilidad y proyección", "Reportes"]
    )

    with tab1:
        k1, k2, k3, k4 = st.columns(4)
        k1.metric(f"Ingresos - {periodo_base_txt}", formato_moneda(resumen_base["ingresos"]))
        k2.metric(f"Costos - {periodo_base_txt}", formato_moneda(resumen_base["costos"]))
        k3.metric(f"Gastos - {periodo_base_txt}", formato_moneda(resumen_base["gastos"]))
        k4.metric(f"Utilidad neta - {periodo_base_txt}", formato_moneda(resumen_base["utilidad_neta"]))

        k5, k6, k7, k8 = st.columns(4)
        k5.metric("Margen neto", formato_pct(resumen_base["margen_neto_pct"]))
        k6.metric("Punto de equilibrio estimado", formato_moneda(resumen_base["punto_equilibrio"]))
        k7.metric("Facturación requerida para utilidad objetivo", formato_moneda(resumen_base["facturacion_total_objetivo"]))
        k8.metric("Facturación adicional requerida", formato_moneda(resumen_base["facturacion_adicional_objetivo"]))

        st.subheader("Semáforos y alertas")
        for nivel, txt in alertas:
            st.markdown(semaforo_html(nivel, txt), unsafe_allow_html=True)

        st.subheader(f"Conclusión automática - {periodo_base_txt}")
        st.info(conclusion)

        st.subheader(f"Gráficos principales - {anio_base}")
        img1 = crear_grafico_ingresos_vs_gastos(piv_año, anio_base)
        st.image(img1)
        img2 = crear_grafico_utilidad(piv_año, anio_base)
        st.image(img2)

    with tab2:
        st.subheader(f"Comparativo del periodo: {periodo_base_txt} vs {periodo_comp_txt}")
        st.dataframe(df_comp, use_container_width=True)

        st.subheader(f"Comparativo resumido en texto: {anio_base} vs {anio_comp}")
        for _, fila in df_comp.iterrows():
            signo = "subió" if fila["Variación $"] > 0 else "bajó"
            st.write(
                f"- {fila['Rubro']}: {signo} {formato_moneda(abs(fila['Variación $']))} ({formato_pct(abs(fila['Variación %']))})."
            )

    with tab3:
        st.subheader(f"Ranking de productos o servicios con mayor ingreso - {periodo_base_txt}")
        st.dataframe(ranking, use_container_width=True)
        img3 = crear_grafico_participacion(ranking, periodo_base_txt)
        st.image(img3)

        st.subheader(f"Ingresos mensuales - Año {anio_base}")
        st.dataframe(piv_año[["Etiqueta", "Ingresos"]], use_container_width=True)
        st.metric(f"Ingresos acumulados del año {anio_base}", formato_moneda(piv_año["Ingresos"].sum()))

    with tab4:
        st.subheader(f"Gastos por categoría - {periodo_base_txt}")
        st.dataframe(gasto_cat, use_container_width=True)
        if not gasto_cat.empty:
            mayor_gasto = gasto_cat.iloc[0]
            st.warning(
                f"La categoría con mayor gasto en {periodo_base_txt} es {mayor_gasto['CategoriaGasto']} por {formato_moneda(mayor_gasto['Valor'])}."
            )

    with tab5:
        u1, u2, u3 = st.columns(3)
        u1.metric(f"Utilidad bruta - {periodo_base_txt}", formato_moneda(resumen_base["utilidad_bruta"]))
        u2.metric(f"Utilidad operativa - {periodo_base_txt}", formato_moneda(resumen_base["utilidad_operativa"]))
        u3.metric(f"Utilidad neta - {periodo_base_txt}", formato_moneda(resumen_base["utilidad_neta"]))

        p1, p2, p3 = st.columns(3)
        p1.metric(f"Escenario pesimista - cierre {anio_base}", formato_moneda(proy["pesimista"]))
        p2.metric(f"Escenario normal - cierre {anio_base}", formato_moneda(proy["normal"]))
        p3.metric(f"Escenario optimista - cierre {anio_base}", formato_moneda(proy["optimista"]))

        if resumen_base["utilidad_neta"] < 0:
            st.error(
                f"En {periodo_base_txt}, la empresa presenta pérdida neta. Los gastos y costos están superando los ingresos."
            )
        elif resumen_base["utilidad_neta"] > 0:
            st.success(
                f"En {periodo_base_txt}, la empresa presenta utilidad neta positiva y el negocio es rentable en este periodo."
            )
        else:
            st.info(f"En {periodo_base_txt}, la empresa está en punto de equilibrio.")

        st.image(crear_grafico_utilidad(piv_año, anio_base))

    with tab6:
        st.subheader("Descargas gerenciales")
        xlsx_buffer = exportar_excel_resultados(df_comp, ranking, gasto_cat, piv_año)

        img1 = crear_grafico_ingresos_vs_gastos(piv_año, anio_base)
        img2 = crear_grafico_utilidad(piv_año, anio_base)
        img3 = crear_grafico_participacion(ranking, periodo_base_txt)

        pdf_buffer = generar_pdf(
            periodo_base_txt,
            periodo_comp_txt,
            resumen_base,
            resumen_comp,
            df_comp,
            ranking,
            gasto_cat,
            alertas,
            conclusion,
            img1,
            img2,
            img3,
            anio_base,
            anio_comp,
        )

        st.download_button(
            "Descargar informe PDF",
            data=pdf_buffer,
            file_name="informe_financiero_gerencial_bersoft.pdf",
            mime="application/pdf",
        )

        st.download_button(
            "Descargar resultados en Excel",
            data=xlsx_buffer,
            file_name="resultados_financieros_bersoft.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Sube al menos un archivo Excel para iniciar el análisis.")
