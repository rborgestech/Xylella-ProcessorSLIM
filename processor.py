# -*- coding: utf-8 -*-
from io import BytesIO
from pathlib import Path
from typing import Tuple, Dict
from copy import copy

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.cell_range import MultiCellRange

import unicodedata


DGAV_TEMPLATE_PATH = Path(__file__).parent / "DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA.xlsx"


INPUT_TO_DGAV_COLMAP = {
    "DATA_RECEPCAO": "Data recepção amostras",
    "DATA_COLHEITA": "Data colheita",
    "CODIGO_AMOSTRA": "Código_amostra (Código original / Referência amostra)",
    "HOSPEDEIRO": "Espécie indicada / Hospedeiro",
    "TIPO_AMOSTRA": "Tipo amostra Simples / Composta",
    "ID_ZONA": "Id Zona (Classificação de zona de origem)",
    "COD_INT_LAB": "Código interno Lab",
    "DATA_REQUERIDO": "Data requerido",
    "RESPONSAVEL_AMOSTRAGEM": "Responsável Amostragem (Zona colheita)",
    "RESP_COLHEITA": "Responsável colheita (Técnico responsável)",
    "PREP_COMMENTS": "Prep_Comments (Observações cliente)",
    "PROCEDURE": "Procedure",
}

REQUIRED_DGAV_COLS = [
    "DATA_RECEPCAO",
    "DATA_COLHEITA",
    "CODIGO_AMOSTRA",
    "HOSPEDEIRO",
    "TIPO_AMOSTRA",
    "ID_ZONA",
    "PROCEDURE",
    "COD_INT_LAB",
    "DATA_REQUERIDO",
]


# ---------------------------------------------------------
# Normalização tolerante de nomes
# ---------------------------------------------------------
def _norm(text: str | None) -> str:
    if text is None:
        return ""
    s = str(text)
    s = s.replace("\u00A0", " ")
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.replace("_", " ").replace("-", " ")
    s = s.lower()
    return " ".join(s.split()).strip()


# ---------------------------------------------------------
# Encontrar cabeçalho no pré-registo
# ---------------------------------------------------------
def _find_header_row(df_raw: pd.DataFrame, target: str) -> int:
    target_norm = _norm(target)

    for idx, row in df_raw.iterrows():
        if row.astype(str).apply(_norm).eq(target_norm).any():
            return idx

    for idx, row in df_raw.iterrows():
        if row.astype(str).apply(_norm).str.contains("codigo amostra", na=False).any():
            return idx

    raise ValueError("Não foi possível identificar a linha de cabeçalho no pré-registo.")


# ---------------------------------------------------------
# Leitura do pré-registo com fórmulas já calculadas
# ---------------------------------------------------------
def _load_pre_registo_df(uploaded_file) -> pd.DataFrame:
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active

    rows = list(ws.values)
    df_raw = pd.DataFrame(rows)

    header_row = _find_header_row(
        df_raw, "Código_amostra (Código original / Referência amostra)"
    )
    headers = df_raw.iloc[header_row].tolist()

    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")

    return df


# ---------------------------------------------------------
# Filtrar apenas linhas com CODIGO_AMOSTRA preenchido
# ---------------------------------------------------------
def _filter_sample_rows(df: pd.DataFrame) -> pd.DataFrame:
    target_norm = _norm(INPUT_TO_DGAV_COLMAP["CODIGO_AMOSTRA"])
    col = None

    for c in df.columns:
        if _norm(c) == target_norm:
            col = c
            break

    if col is None:
        return df

    mask = df[col].notna() & (df[col].astype(str).str.strip() != "")
    return df[mask].copy()


# ---------------------------------------------------------
# Mapear colunas do pré-registo para DGAV
# ---------------------------------------------------------
def _map_input_columns(df: pd.DataFrame) -> Dict[str, str]:
    norm_to_real = { _norm(col): col for col in df.columns }

    mapped = {}
    for dgav_col, input_label in INPUT_TO_DGAV_COLMAP.items():
        key_norm = _norm(input_label)
        mapped[dgav_col] = norm_to_real.get(key_norm)

    return mapped


# ---------------------------------------------------------
# Construir mapa de cabeçalhos no template
# ---------------------------------------------------------
def _build_header_index(ws) -> Dict[str, int]:
    return {
        _norm(ws.cell(row=1, column=c).value): c
        for c in range(1, ws.max_column + 1)
        if ws.cell(row=1, column=c).value
    }


# ---------------------------------------------------------
# Marcar colunas obrigatórias com vermelho se faltar valor
# ---------------------------------------------------------
def _mark_required_empty_columns(ws, header_indices, start_row, last_row):
    red = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

    for col_name in REQUIRED_DGAV_COLS:
        col_idx = header_indices.get(_norm(col_name))
        if col_idx is None:
            continue

        any_empty = any(
            ws.cell(row=r, column=col_idx).value in (None, "")
            for r in range(start_row, last_row + 1)
        )

        if any_empty:
            ws.cell(row=1, column=col_idx).fill = red


# ---------------------------------------------------------
# 🔥 Função principal FINAL
# ---------------------------------------------------------
def process_pre_to_dgav(uploaded_file) -> Tuple[bytes, str]:
    if not DGAV_TEMPLATE_PATH.exists():
        raise FileNotFoundError("Template DGAV não encontrado.")

    # 1) Carregar pré-registo
    df_in = _load_pre_registo_df(uploaded_file)

    # Filtrar só amostras reais
    df_in = _filter_sample_rows(df_in)
    df_in = df_in.reset_index(drop=True)

    n_samples = len(df_in)
    input_colmap = _map_input_columns(df_in)

    # 2) Carregar template
    template_bytes = DGAV_TEMPLATE_PATH.read_bytes()
    wb = load_workbook(BytesIO(template_bytes))
    ws = wb["Default"]

    header_indices = _build_header_index(ws)

    # Guardar linha default
    base_values = {
        norm_name: ws.cell(row=2, column=col_idx).value
        for norm_name, col_idx in header_indices.items()
    }

    # Guardar listas de validação
    original_validations = list(ws.data_validations.dataValidation)

    # 3) Limpar linhas abaixo da linha 1
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # 4) Escrever linhas
    start_row = 2
    last_row = start_row + n_samples - 1

    for i, (_, row_in) in enumerate(df_in.iterrows()):
        excel_row = start_row + i

        # Copiar defaults
        for norm_name, col_idx in header_indices.items():
            ws.cell(row=excel_row, column=col_idx).value = base_values[norm_name]

        # Substituir pelas colunas do pré-registo
        for dgav_col, input_label in INPUT_TO_DGAV_COLMAP.items():
            col_idx = header_indices.get(_norm(dgav_col))
            if col_idx is None:
                continue

            df_col = input_colmap.get(dgav_col)
            if not df_col:
                continue

            value = row_in.get(df_col)
            if isinstance(value, float) and pd.isna(value):
                value = None
            if hasattr(value, "date"):
                value = value.date()

            ws.cell(row=excel_row, column=col_idx).value = value

    # 5) Marcar erros
    if n_samples > 0:
        _mark_required_empty_columns(ws, header_indices, start_row, last_row)

    # 6) Cortar linhas extra
    if ws.max_row > last_row:
        ws.delete_rows(last_row + 1, ws.max_row - last_row)

    # 7) Reaplicar validações de dados apenas ao intervalo real
    ws.data_validations.dataValidation = []

    for dv in original_validations:
        new_dv = copy(dv)

        new_ranges = []
        for r in dv.ranges.ranges:
            col_letter = ''.join(filter(str.isalpha, str(r)))
            new_ranges.append(f"{col_letter}{start_row}:{col_letter}{last_row}")

        new_dv.ranges = MultiCellRange(new_ranges)
        ws.data_validations.append(new_dv)

    # 8) Exportar
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue(), f"Foram processadas {n_samples} amostras."
