# -*- coding: utf-8 -*-
import io
import zipfile
from datetime import datetime
from typing import List, Tuple, Dict

import streamlit as st
from openpyxl import load_workbook

from processor import process_pre_to_dgav, REQUIRED_DGAV_COLS, _norm


# ───────────────────────────────────────────────
# Análise do ficheiro DGAV gerado (para avisos/erros)
# ───────────────────────────────────────────────
def analyse_output_xlsx(xlsx_bytes: bytes) -> Tuple[int, List[str], List[str]]:
    """
    Analisa o ficheiro DGAV gerado:
      - conta nº de amostras (linhas com CODIGO_AMOSTRA)
      - devolve (warnings, hard_errors)
        * hard_errors: coluna obrigatória ausente OU alguma célula vazia
    """
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb["Default"]

    # Mapear cabeçalhos normalizados -> índice
    header_indices: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v:
            header_indices[_norm(v)] = col

    warnings: List[str] = []
    hard_errors: List[str] = []

    # 1) Determinar última linha de dados com base em CODIGO_AMOSTRA
    codigo_idx = header_indices.get(_norm("CODIGO_AMOSTRA"))
    last_row = 1
    sample_count = 0
    if codigo_idx:
        for row in range(2, ws.max_row + 1):
            v = ws.cell(row=row, column=codigo_idx).value
            if v not in (None, ""):
                sample_count += 1
                last_row = row
    else:
        hard_errors.append("Coluna obrigatória ausente no output: CODIGO_AMOSTRA")
        last_row = ws.max_row

    # 2) Verificar colunas obrigatórias (modo 2: qualquer célula vazia = erro)
    for col_name in REQUIRED_DGAV_COLS:
        col_idx = header_indices.get(_norm(col_name))
        if col_idx is None:
            hard_errors.append(f"Coluna obrigatória ausente no output: {col_name}")
            continue

        any_empty = False
        for row in range(2, last_row + 1):
            v = ws.cell(row=row, column=col_idx).value
            if v in (None, ""):
                any_empty = True
                break

        if any_empty:
            hard_errors.append(f"Coluna obrigatória com células vazias: {col_name}")

    return sample_count, warnings, hard_errors


# ───────────────────────────────────────────────
# ZIP em memória com outputs + summary
# ───────────────────────────────────────────────
def build_zip_with_summary(
    outputs: List[Tuple[str, bytes, int, List[str], List[str]]],
    summary_lines: List[str],
    timestamp: str,
) -> bytes:
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for original_name, data, sample_count, warns, errs in outputs:
            base = original_name.rsplit(".", 1)[0]
            out_name = f"{base}_DGAV_{timestamp}.xlsx"
            z.writestr(out_name, data)
        z.writestr("summary.txt", "\n".join(summary_lines))
    mem.seek(0)
    return mem.getvalue()


# ───────────────────────────────────────────────
# Configuração base e CSS
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella → DGAV", page_icon="🧪", layout="centered")

st.title("🧪 Xylella – Conversor Pré-registo → DGAV")
st.caption(
    "Carrega ficheiros **“AVALIAÇÃO PRÉ-REGISTO – Amostras Xylella”** "
    "e gera automaticamente o ficheiro **DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA.xlsx**."
)

# CSS inspirado no app dos PDFs
st.markdown(
    """
<style>
.stButton > button[kind="primary"]{
  background:#CA4300!important;border:1px solid #CA4300!important;color:#fff!important;
  font-weight:600!important;border-radius:6px!important;transition:background-color .2s ease-in-out!important;
}
.stButton > button[kind="primary"]:hover{background:#A13700!important;border-color:#A13700!important;}
[data-testid="stFileUploader"]>div:first-child{
  border:2px dashed #CA4300!important;border-radius:10px!important;padding:1rem!important
}

/* Caixas de estado */
.file-box{border-radius:8px;padding:.6rem 1rem;margin-bottom:.5rem;opacity:0;
          animation:fadeIn .4s ease forwards}
@keyframes fadeIn{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:translateY(0)}}
.file-box.processing{background:#E8F1FB;border-left:4px solid #2B6CB0}
.file-box.success{background:#e6f9ee;border-left:4px solid #1a7f37}
.file-box.warning{background:#fff8e5;border-left:4px solid #e6a100}
.file-box.error{background:#fdeaea;border-left:4px solid #cc0000}
.file-title{font-size:.9rem;font-weight:600;color:#1A365D}
.file-sub{font-size:.8rem;color:#2A4365}

/* Pontinhos animados */
.dots::after{content:'...';display:inline-block;animation:dots 1.5s steps(4,end) infinite}
@keyframes dots{
  0%,20%{color:rgba(42,67,101,0);text-shadow:.25em 0 0 rgba(42,67,101,0),.5em 0 0 rgba(42,67,101,0)}
  40%{color:#2A4365;text-shadow:.25em 0 0 rgba(42,67,101,0),.5em 0 0 rgba(42,67,101,0)}
  60%{text-shadow:.25em 0 0 #2A4365,.5em 0 0 rgba(42,67,101,0)}
  80%,100%{text-shadow:.25em 0 0 #2A4365,.5em 0 0 #2A4365}
}

/* Botão branco "novo processamento" */
.clean-btn{
  background:#fff!important;border:1px solid #ccc!important;color:#333!important;font-weight:600!important;
  border-radius:8px!important;padding:.5rem 1.2rem!important;transition:all .2s ease
}
.clean-btn:hover{border-color:#999!important;color:#000!important}
</style>
""",
    unsafe_allow_html=True,
)

# ───────────────────────────────────────────────
# Estado da app
# ───────────────────────────────────────────────
if "stage" not in st.session_state:
    st.session_state.stage = "idle"
if "uploads" not in st.session_state:
    st.session_state.uploads = None


def reset_app():
    st.session_state.stage = "idle"
    st.session_state.uploads = None


# ───────────────────────────────────────────────
# Interface principal (2 fases: idle / processing)
# ───────────────────────────────────────────────
if st.session_state.stage == "idle":
    uploads = st.file_uploader(
        "📂 Carrega um ou vários ficheiros de pré-registo (XLSX)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="file_uploader",
    )

    if uploads:
        if st.button("📄 Processar ficheiros de Input", type="primary"):
            st.session_state.uploads = uploads
            st.session_state.stage = "processing"
            st.rerun()
    else:
        st.info("💡 Carrega pelo menos um ficheiro de pré-registo para ativar o processamento.")

elif st.session_state.stage == "processing":
    uploads = st.session_state.uploads
    total = len(uploads)

    st.info("⏳ A processar ficheiros... aguarde até o processo terminar.")
    progress = st.progress(0.0)

    outputs: List[Tuple[str, bytes, int, List[str], List[str]]] = []
    summary_lines: List[str] = []
    total_samples = 0
    warning_files = 0
    error_files = 0

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    for i, up in enumerate(uploads, start=1):
        placeholder = st.empty()
        placeholder.markdown(
            f"""
            <div class='file-box processing'>
              <div class='file-title'>📄 {up.name}</div>
              <div class='file-sub'>Ficheiro {i} de {total} — a processar<span class="dots"></span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        try:
            data_in = io.BytesIO(up.getbuffer())
            output_bytes, log_msg = process_pre_to_dgav(data_in)

            sample_count, col_warnings, hard_errors = analyse_output_xlsx(output_bytes)
            total_samples += sample_count

            if hard_errors:
                status_class = "error"
                error_files += 1
                bullets_err = "<br>".join(f"• {e}" for e in hard_errors)
                extra_html = f"<div class='file-sub'>❌ Erros:<br>{bullets_err}</div>"
            elif col_warnings:
                status_class = "warning"
                warning_files += 1
                bullets_warn = "<br>".join(f"• {w}" for w in col_warnings)
                extra_html = f"<div class='file-sub'>⚠️ Avisos:<br>{bullets_warn}</div>"
            else:
                status_class = "success"
                extra_html = ""

            outputs.append((up.name, output_bytes, sample_count, col_warnings, hard_errors))

            html = (
                f"<div class='file-box {status_class}'>"
                f"<div class='file-title'>📄 {up.name}</div>"
                f"<div class='file-sub'><b>{sample_count}</b> amostra(s) processadas.</div>"
                f"{extra_html}</div>"
            )
            placeholder.markdown(html, unsafe_allow_html=True)

            # resumo textual
            summary_line = f"{up.name}: {sample_count} amostra(s). {log_msg}"
            if hard_errors:
                summary_line += " ❌ " + " | ".join(hard_errors)
            elif col_warnings:
                summary_line += " ⚠ " + " | ".join(col_warnings)
            summary_lines.append(summary_line)

        except Exception as e:
            error_files += 1
            msg = f"{up.name}: erro ao processar ({e})"
            summary_lines.append(msg)
            html = (
                f"<div class='file-box error'>"
                f"<div class='file-title'>📄 {up.name}</div>"
                f"<div class='file-sub'>❌ Erro ao processar: {e}</div>"
                f"</div>"
            )
            placeholder.markdown(html, unsafe_allow_html=True)

        progress.progress(i / total)

    # ───────────────────────────────────────
    # Resumo final + botões de download/reset
    # ───────────────────────────────────────
    st.markdown(
        f"""
        <div style='text-align:center;margin-top:1.5rem;'>
          <h3>🏁 Processamento concluído!</h3>
          <p>Foram processados <b>{len(outputs)}</b> ficheiro(s) válido(s),
          com um total de <b>{total_samples}</b> amostras.<br>
          Ficheiros com avisos: <b>{warning_files}</b>.<br>
          Ficheiros com erro: <b>{error_files}</b>.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # acrescentar linhas finais ao summary
    summary_lines.append("")
    summary_lines.append(f"Total de ficheiros válidos: {len(outputs)}")
    summary_lines.append(f"Total de amostras: {total_samples}")
    summary_lines.append(f"Ficheiros com avisos: {warning_files}")
    summary_lines.append(f"Ficheiros com erro: {error_files}")
    summary_lines.append(f"Executado em: {datetime.now():%d/%m/%Y às %H:%M:%S}")

    if len(outputs) == 1:
        original_name, data, sample_count, warns, errs = outputs[0]
        base = original_name.rsplit(".", 1)[0]
        out_name = f"{base}_DGAV_{timestamp}.xlsx"

        st.download_button(
            label="⬇️ Descarregar ficheiro DGAV",
            data=data,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            label="📝 Descarregar resumo (summary.txt)",
            data="\n".join(summary_lines),
            file_name=f"summary_{timestamp}.txt",
            mime="text/plain",
        )
    elif len(outputs) > 1:
        zip_bytes = build_zip_with_summary(outputs, summary_lines, timestamp)
        zip_name = f"xylella_dgav_{timestamp}.zip"

        st.download_button(
            label="📦 Descarregar resultados (ZIP)",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
        )

    # botão de novo processamento
    st.markdown("<br>", unsafe_allow_html=True)
    st.button(
        "🔁 Novo processamento",
        key="btn_reset",
        use_container_width=True,
        on_click=reset_app,
    )
