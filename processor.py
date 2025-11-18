# -*- coding: utf-8 -*-
from io import BytesIO
from pathlib import Path
from typing import Tuple, Dict

import pandas as pd
from openpyxl import load_workbook


# Caminho para o template DGAV de saída
DGAV_TEMPLATE_PATH = Path(__file__).parent / "DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA.xlsx"


# Mapeamento entre colunas do ficheiro de pré-registo e colunas DGAV
# (keys = coluna DGAV, values = coluna no ficheiro de pré-registo)
INPUT_TO_DGAV_COLMAP = {
    "DATA_RECEPCAO": "Data recepção amostras",
    "DATA_COLHEITA": "Data colheita",
    "CODIGO_AMOSTRA": "Código_amostra (Código original / Referência amostra)",
    "HOSPEDEIRO": "Espécie indicada / Hospedeiro",
    "TIPO_AMOSTRA": "Tipo amostra Simples / Composta",
    "ID_ZONA": "Id_Zona (Classificação de zona de origem)",
    "COD_INT_LAB": "Código interno Lab",
    "DATA_REQUERIDO": "Data requerido",
    # ------------------------
    "RESPONSAVEL_AMOSTRAGEM": "Responsável Amostragem (Zona colheita)",
    "RESP_COLHEITA": "Responsável colheita (Técnico responsável)",
    "PREP_COMMENTS": "Prep_Comments (Observações cliente)",
    "PROCEDURE": "Procedure",
}


def _find_header_row(df_raw: pd.DataFrame, target: str) -> int:
    """Tenta localizar a linha de cabeçalho procurando uma célula igual ao texto `target`."""
    target = str(target).strip()
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.strip().eq(target).any():
            return idx
    # fallback: procura substring
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains("Código_amostra", na=False).any():
            return idx
    raise ValueError("Não foi possível encontrar a linha de cabeçalho no ficheiro de pré-registo.")




def _load_pre_registo_df(uploaded_file) -> pd.DataFrame:
    """
    Lê o Excel de pré-registo usando openpyxl data_only=True
    para obter os valores calculados das fórmulas.
    Depois converte para DataFrame mantendo o cabeçalho correto.
    """
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active

    # Ler todas as linhas como listas
    rows = list(ws.values)

    df_raw = pd.DataFrame(rows)

    # Encontrar a linha com o cabeçalho
    header_row = _find_header_row(df_raw, "Código_amostra (Código original / Referência amostra)")
    headers = df_raw.iloc[header_row].tolist()

    # DataFrame final
    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = headers

    # Remover linhas totalmente vazias
    df = df.dropna(how="all")

    return df


def _build_header_index(ws) -> Dict[str, int]:
    """Cria um dicionário {nome_coluna: índice_coluna} a partir da linha 1 da folha DGAV."""
    header_row = 1
    header_indices: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val:
            header_indices[str(val)] = col
    return header_indices


def process_pre_to_dgav(uploaded_file) -> Tuple[bytes, str]:
    """
    Recebe o Excel de pré-registo e devolve o ficheiro DGAV preenchido (em bytes).

    Returns
    -------
    (xlsx_bytes, log_msg)
    """
    if not DGAV_TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template DGAV não encontrado em: {DGAV_TEMPLATE_PATH}")

    # DataFrame com os registos de amostras do ficheiro de entrada
    df_in = _load_pre_registo_df(uploaded_file)
    df_in = df_in.reset_index(drop=True)

    # Carrega template DGAV
    wb = load_workbook(DGAV_TEMPLATE_PATH)
    ws = wb["Default"]

    # Guardar valores base da linha 2 (cliente, projecto, etc.) para replicar em todas as linhas
    max_col = ws.max_column
    base_values = {col: ws.cell(row=2, column=col).value for col in range(1, max_col + 1)}

    # Índice de cabeçalhos DGAV
    header_indices = _build_header_index(ws)

    # Limpa linhas de dados existentes (a partir da linha 2)
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # Escreve cada amostra
    start_row = 2
    for i, (_, row_in) in enumerate(df_in.iterrows(), start=0):
        excel_row = start_row + i

        # Preenche com valores base
        for col in range(1, max_col + 1):
            ws.cell(row=excel_row, column=col).value = base_values.get(col)

        # Preenche campos específicos vindos do ficheiro de pré-registo
        for dgav_col, input_col in INPUT_TO_DGAV_COLMAP.items():
            col_idx = header_indices.get(dgav_col)
            if col_idx is None:
                continue

            value = row_in.get(input_col, None)

            # Converte NaN para None
            if isinstance(value, float) and pd.isna(value):
                value = None

            # Se for datetime, remove a hora
            if hasattr(value, "date"):
                value = value.date()

            ws.cell(row=excel_row, column=col_idx).value = value

    # Exporta para bytes em memória
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    log_msg = f"Foram processadas {len(df_in)} amostras para o ficheiro DGAV."
    return output.getvalue(), log_msg
