import streamlit as st
import pandas as pd

st.set_page_config(page_title="Control Financiero Inteligente V2", layout="wide")

st.title("Control Financiero Inteligente V2")
st.caption("Análisis financiero con comparativos, alertas y recomendaciones")

archivo = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if archivo:
    df = pd.read_excel(archivo)

    st.subheader("Vista previa")
    st.dataframe(df.head())

    # Simulación básica (puedes mejorar luego)
    total = df.select_dtypes(include='number').sum().sum()

    st.subheader("Resumen general")
    st.write(f"Total detectado: ${total:,.0f}")

    if total < 0:
        st.error("Estado crítico: pérdidas detectadas")
    else:
        st.success("Estado positivo")

    st.subheader("Recomendaciones")
    st.write("- Revisar gastos altos")
    st.write("- Aumentar ingresos principales")
    st.write("- Controlar costos fijos")
