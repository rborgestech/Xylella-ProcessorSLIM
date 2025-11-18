# -*- coding: utf-8 -*-
from io import BytesIO
from pathlib import Path
from typing import Tuple, Dict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# Caminho do template DGAV – será lido APENAS como bytes (nunca modificado)
DGAV_TEMPLATE_PATH = Path(__file__).parent / "DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA.xlsx"


# Mapeamento entre colunas do pré-registo e colunas DGAV
INPUT_TO_DGAV_COLMAP = {
    "DATA_RECEPCAO": "Data recepção amostras",
    "DATA_COLHEITA": "Data colheita",
    "CODIGO_AMOSTRA": "Código_amostra (Código original / Referência amostra)",
    "HOSPEDEIRO": "Espécie indicada / Hospedeiro",
    "TIPO_AMOSTRA": "Tipo amostra Simples / Composta",
    "ID_ZONA": "Id Zona (Classificação de zona de origem)",
    "COD_INT_LAB": "Código interno Lab",
    "DATA_REQUERIDO": "Data requerida",
    "RESPONSAVEL_AMOSTRAGEM": "Responsável Amostragem (Zona colheita)",
    "RESP_COLHEITA": "Responsável colheita (Técnico responsável)",
    "PREP_COMMENTS": "Prep_Comments (Observações cliente)",
    "PROCEDURE": "Procedure",
}

# Colunas obrigatórias no ficheiro DGAV
REQUIRED_DGAV_COLS = [
    "DATA_RECEPCAO",
    "DATA_COLHEITA",
    "CODIGO_AMOSTRA",
    "HOSPEDEIRO",
    "TIPO_AMOSTRA",
    "ID_ZONA",
    "PROCEDURE",
]


def _find_header_row(df_raw: pd.DataFrame, target: str) -> int:
    """Encontra a linha de cabeçalho no pré-registo."""
    target = str(target).strip()
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.strip().eq(target).any():
            return idx

    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains("Código_amostra", na=False).any():
            return idx

    raise ValueError("Não foi possível encontrar a linha de cabeçalho no ficheiro de pré-registo.")


def _load_pre_registo_df(uploaded_file) -> pd.DataFrame:
    """
    Lê o pré-registo usando openpyxl para obter valores calculados
    e converte para DataFrame.
    """
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active

    rows = list(ws.values)
    df_raw = pd.DataFrame(rows)

    header_row = _find_header_row(df_raw, "Código_amostra (Código original / Referência amostra)")
    headers = df_raw.iloc[header_row].tolist()

    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all")
    return df


def _build_header_index(ws) -> Dict[str, int]:
    """Constrói dicionário nome_coluna → índice_coluna."""
    header_indices = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val:
            header_indices[str(val)] = col
    return header_indices


def _mark_required_empty_columns(ws, header_indices: Dict[str, int], start_row: int = 2):
    """Pinta a vermelho colunas obrigatórias sem qualquer registo."""
    red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

    for col_name in REQUIRED_DGAV_COLS:
        col_idx = header_indices.get(col_name)
        if col_idx is None:
            continue

        has_value = False
        for row in range(start_row, ws.max_row + 1):
            if ws.cell(row=row, column=col_idx).value not in (None, ""):
                has_value = True
                break

        if not has_value:
            ws.cell(row=1, column=col_idx).fill = red_fill


def process_pre_to_dgav(uploaded_file) -> Tuple[bytes, str]:
    """
    Converte pré-registo → DGAV, devolvendo bytes do Excel.
    TOTALMENTE EM MEMÓRIA, ZERO ESCRITA EM DISCO.
    """
    if not DGAV_TEMPLATE_PATH.exists():
        raise FileNotFoundError("Template DGAV não encontrado.")

    # --------- LER PRÉ-REGISTO ----------
    df_in = _load_pre_registo_df(uploaded_file)
    df_in = df_in.reset_index(drop=True)

    # --------- CARREGAR TEMPLATE COMO BYTES (SEGURO) ----------
    # Lê conteúdo original sem nunca modificar o ficheiro físico
    template_bytes = DGAV_TEMPLATE_PATH.read_bytes()

    # Cria cópia em memória – workbook LIMPO por processamento
    template_stream = BytesIO(template_bytes)

    # workBook isolado completamente, nunca guardado no disco
    wb = load_workbook(template_stream)
    ws = wb["Default"]

    # --------- PREPARAR ESCRITA ----------
    max_col = ws.max_column
    base_values = {col: ws.cell(row=2, column=col).value for col in range(1, max_col + 1)}
    header_indices = _build_header_index(ws)

    # Remover qualquer conteúdo anterior → mas só desta cópia em memória
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # --------- ESCREVER REGISTOS ----------
    start_row = 2
    for i, (_, row_in) in enumerate(df_in.iterrows(), start=0):
        excel_row = start_row + i

        # copiar linha 2 (valores estáticos)
        for col in range(1, max_col + 1):
            ws.cell(row=excel_row, column=col).value = base_values.get(col)

        # copiar valores variáveis
        for dgav_col, input_col in INPUT_TO_DGAV_COLMAP.items():
            col_idx = header_indices.get(dgav_col)
            if col_idx is None:
                continue

            value = row_in.get(input_col, None)

            # remover NaN
            if isinstance(value, float) and pd.isna(value):
                value = None

            # remover horas se datetime
            if hasattr(value, "date"):
                value = value.date()

            ws.cell(row=excel_row, column=col_idx).value = value

    # --------- VALIDAÇÃO ----------
    _mark_required_empty_columns(ws, header_indices, start_row=start_row)

    # --------- EXPORTAÇÃO SEGURA EM MEMÓRIA ----------
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue(), f"Foram processadas {len(df_in)} amostras."
