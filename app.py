# -*- coding: utf-8 -*-
import streamlit as st
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime

from processor import process_pre_to_dgav

st.set_page_config(page_title="Xylella – Conversor para DGAV", page_icon="🧪", layout="centered")

st.title("Xylella – Conversor para ficheiro DGAV")

st.write(
    "Carregue um ou vários ficheiros Excel **“AVALIAÇÃO PRÉ-REGISTO – Amostras Xylella”** "
    "e obtenha como resultado o(s) ficheiro(s) **DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA.xlsx** preenchido(s)."
)

uploaded_files = st.file_uploader(
    "Ficheiro(s) XLSX de entrada (pré-registo)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} ficheiro(s) carregado(s).")

    if st.button("Processar ficheiro(s)"):
        outputs = []
        logs = []

        for f in uploaded_files:
            # garantir que o ponteiro do ficheiro está no início
            f.seek(0)
            try:
                output_bytes, log_msg = process_pre_to_dgav(f)
                logs.append(f"✅ {f.name}: {log_msg}")
                outputs.append((f.name, output_bytes))
            except Exception as e:
                logs.append(f"❌ {f.name}: erro ao processar -> {e}")

        # mostrar logs
        for line in logs:
            st.write(line)

        if not outputs:
            st.error("Nenhum ficheiro foi processado com sucesso.")
        elif len(outputs) == 1:
            # apenas 1 ficheiro -> download direto
            original_name, data = outputs[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = f"DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA_{timestamp}.xlsx"

            st.download_button(
                label="📥 Download do ficheiro DGAV",
                data=data,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            # vários ficheiros -> criar ZIP em memória
            zip_buffer = BytesIO()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            with ZipFile(zip_buffer, "w") as zip_file:
                for original_name, data in outputs:
                    base = original_name.rsplit(".", 1)[0]
                    out_name = f"{base}_DGAV_{timestamp}.xlsx"
                    zip_file.writestr(out_name, data)

            zip_buffer.seek(0)

            st.download_button(
                label="📦 Download ZIP com ficheiros DGAV",
                data=zip_buffer,
                file_name=f"DGAV_FILES_{timestamp}.zip",
                mime="application/zip",
            )
else:
    st.info("Selecione pelo menos um ficheiro de pré-registo em formato XLSX.")
