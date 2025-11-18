# -*- coding: utf-8 -*-
from io import BytesIO
from pathlib import Path
from typing import Tuple, Dict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# Caminho do template DGAV (NUNCA é alterado em disco)
DGAV_TEMPLATE_PATH = Path(__file__).parent / "DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA.xlsx"


# Mapeamento entre colunas do pré-registo e colunas DGAV
INPUT_TO_DGAV_COLMAP = {
    "DATA_RECEPCAO": "Data recepção amostras",
    "DATA_COLHEITA": "Data colheita",
    "CODIGO_AMOSTRA": "Código_amostra (Código original / Referência amostra)",
    "HOSPEDEIRO": "Espécie indicada / Hospedeiro",
    "TIPO_AMOSTRA": "Tipo amostra Simples / Composta",
    "ID_ZONA": "Id_Zona (Classificação de zona de origem)",
    "COD_INT_LAB": "Código interno Lab",
    "DATA_REQUERIDO": "Data requerido",
    "RESPONSAVEL_AMOSTRAGEM": "Responsável Amostragem (Zona colheita)",
    "RESP_COLHEITA": "Responsável colheita (Técnico responsável)",
    "PREP_COMMENTS": "Prep_Comments (Observações cliente)",
    "PROCEDURE": "Procedure",
}

# Colunas DGAV obrigatórias
REQUIRED_DGAV_COLS = [
    "DATA_RECEPCAO",
    "DATA_COLHEITA",
    "CODIGO_AMOSTRA",
    "HOSPEDEIRO",
    "TIPO_AMOSTRA",
    "ID_ZONA",
    "PROCEDURE",
    "COD_INT_LAB",
    "DATA_REQUERIDO"
]


# ───────────────────────────────────────────────
# Leitura do pré-registo (com fórmulas calculadas)
# ───────────────────────────────────────────────
def _find_header_row(df_raw: pd.DataFrame, target: str) -> int:
    """Encontra a linha de cabeçalho no ficheiro de pré-registo."""
    target = target.strip()

    for idx, row in df_raw.iterrows():
        if row.astype(str).str.strip().eq(target).any():
            return idx

    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains("Código_amostra", na=False).any():
            return idx

    raise ValueError("Não foi possível identificar a linha de cabeçalho.")


def _load_pre_registo_df(uploaded_file) -> pd.DataFrame:
    """Lê os valores do pré-registo já calculados usando openpyxl."""
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active

    rows = list(ws.values)
    df_raw = pd.DataFrame(rows)

    header_row = _find_header_row(df_raw, "Código_amostra (Código original / Referência amostra)")
    headers = df_raw.iloc[header_row].tolist()

    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")

    return df


# ───────────────────────────────────────────────
# Utilitários
# ───────────────────────────────────────────────
def _build_header_index(ws) -> Dict[str, int]:
    """Mapeia o nome das colunas → índice em Excel."""
    header_indices = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v:
            header_indices[str(v)] = col
    return header_indices


def _mark_required_empty_columns(ws, header_indices: Dict[str, int], start_row: int, last_row: int):
    """
    Pinta de vermelho colunas obrigatórias que não tenham QUALQUER valor
    entre start_row e last_row.
    """
    red = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

    for col_name in REQUIRED_DGAV_COLS:
        col_idx = header_indices.get(col_name)
        if col_idx is None:
            continue

        has_value = any(
            ws.cell(row=r, column=col_idx).value not in (None, "")
            for r in range(start_row, last_row + 1)
        )

        if not has_value:
            ws.cell(row=1, column=col_idx).fill = red


# ───────────────────────────────────────────────
# PROCESSAMENTO PRINCIPAL
# ───────────────────────────────────────────────
def process_pre_to_dgav(uploaded_file) -> Tuple[bytes, str]:
    """
    Converte ficheiro de PRÉ-REGISTO → DGAV.
    100% em memória, sem alterar o template no disco.
    """
    if not DGAV_TEMPLATE_PATH.exists():
        raise FileNotFoundError("Template DGAV não encontrado.")

    # 1) Carregar pré-registo
    df_in = _load_pre_registo_df(uploaded_file)
    df_in = df_in.reset_index(drop=True)
    n_samples = len(df_in)

    # 2) Carregar template DGAV sempre fresco (bytes → BytesIO)
    template_bytes = DGAV_TEMPLATE_PATH.read_bytes()
    template_stream = BytesIO(template_bytes)
    wb = load_workbook(template_stream)
    ws = wb["Default"]

    header_indices = _build_header_index(ws)
    max_col = ws.max_column

    # 3) Determinar colunas variáveis vs constantes
    variable_cols = set(INPUT_TO_DGAV_COLMAP.keys())
    all_cols = list(header_indices.keys())
    constant_cols = [name for name in all_cols if name not in variable_cols]

    # Guardar valores da linha 2 do template original para TODAS as colunas
    base_values = {}
    for col_name, col_idx in header_indices.items():
        base_values[col_name] = ws.cell(row=2, column=col_idx).value

    # 4) APAGAR TODAS AS LINHAS A PARTIR DA LINHA 2
    # (fica apenas a linha de cabeçalho original)
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # 5) Escrever uma linha por amostra, a partir da linha 2
    start_row = 2
    last_row = start_row + n_samples - 1 if n_samples > 0 else 1

    for i, (_, row_in) in enumerate(df_in.iterrows()):
        excel_row = start_row + i

        # 5.1 Preencher TODAS as colunas com os valores da antiga linha 2 (base_values)
        for col_name, col_idx in header_indices.items():
            ws.cell(row=excel_row, column=col_idx).value = base_values.get(col_name)

        # 5.2 Substituir colunas variáveis com dados do pré-registo
        for dgav_col, input_col in INPUT_TO_DGAV_COLMAP.items():
            col_idx = header_indices.get(dgav_col)
            if col_idx is None:
                continue

            value = row_in.get(input_col)

            # Converte NaN em None
            if isinstance(value, float) and pd.isna(value):
                value = None

            # Remove hora se for datetime
            if hasattr(value, "date"):
                value = value.date()

            ws.cell(row=excel_row, column=col_idx).value = value

    # 6) Validar colunas obrigatórias apenas no intervalo de amostras
    if n_samples > 0:
        _mark_required_empty_columns(ws, header_indices, start_row=start_row, last_row=last_row)

    # 7) Exportar para bytes (sem tocar no disco)
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue(), f"Foram processadas {n_samples} amostras."
