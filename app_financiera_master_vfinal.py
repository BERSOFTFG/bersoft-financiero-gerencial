
import streamlit as st
import pandas as pd
import io
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Control Financiero PRO", layout="wide")

st.title("Control Financiero Inteligente PRO")

archivo_1 = st.file_uploader("Archivo base", type=["xlsx"])
archivo_2 = st.file_uploader("Archivo comparativo", type=["xlsx"])

def calcular_totales(df):
    ingresos = df[df.iloc[:,0].astype(str).str.contains("INGRES", case=False, na=False)]
    gastos = df[df.iloc[:,0].astype(str).str.contains("GASTO", case=False, na=False)]
    costos = df[df.iloc[:,0].astype(str).str.contains("COSTO", case=False, na=False)]

    total_ingresos = ingresos.select_dtypes(include='number').sum().sum()
    total_gastos = gastos.select_dtypes(include='number').sum().sum()
    total_costos = costos.select_dtypes(include='number').sum().sum()

    utilidad = total_ingresos - total_gastos - total_costos
    return total_ingresos, total_costos, total_gastos, utilidad

def generar_pdf(ing, cost, gast, util):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer)
    styles = getSampleStyleSheet()

    story = []
    story.append(Paragraph("Informe Financiero", styles["Title"]))
    story.append(Spacer(1,12))
    story.append(Paragraph(f"Ingresos: {ing}", styles["Normal"]))
    story.append(Paragraph(f"Costos: {cost}", styles["Normal"]))
    story.append(Paragraph(f"Gastos: {gast}", styles["Normal"]))
    story.append(Paragraph(f"Utilidad: {util}", styles["Normal"]))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

if archivo_1:
    df1 = pd.read_excel(archivo_1)
    ing1, cost1, gast1, util1 = calcular_totales(df1)

    st.metric("Ingresos", ing1)
    st.metric("Costos", cost1)
    st.metric("Gastos", gast1)
    st.metric("Utilidad", util1)

    pdf = generar_pdf(ing1, cost1, gast1, util1)

    st.download_button("Descargar PDF", pdf, "reporte.pdf", "application/pdf")
