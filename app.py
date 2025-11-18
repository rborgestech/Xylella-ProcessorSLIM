# -*- coding: utf-8 -*-
import streamlit as st
from processor import process_pre_to_dgav

st.set_page_config(page_title="Xylella – Conversor para DGAV", page_icon="🧪", layout="centered")

st.title("Xylella – Conversor para ficheiro DGAV")
st.write(
    "Carregue o ficheiro Excel **“AVALIAÇÃO PRÉ-REGISTO – Amostras Xylella”** "
    "e obtenha como resultado o ficheiro **DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA.xlsx** preenchido."
)

uploaded_file = st.file_uploader("Ficheiro XLSX de entrada (pré-registo)", type=["xlsx"])

if uploaded_file is not None:
    st.success("Ficheiro de pré-registo carregado com sucesso.")

    if st.button("Gerar ficheiro DGAV"):
        try:
            output_bytes, log_msg = process_pre_to_dgav(uploaded_file)
            st.info(log_msg)
            st.download_button(
                label="📥 Download do ficheiro DGAV preenchido",
                data=output_bytes,
                file_name="DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA_preenchido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o ficheiro: {e}")
else:
    st.warning("Por favor, selecione o ficheiro de pré-registo em formato XLSX para continuar.")
