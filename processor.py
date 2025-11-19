# -*- coding: utf-8 -*-
from io import BytesIO
from pathlib import Path
from typing import Tuple, Dict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import unicodedata


# Caminho do template DGAV (NUNCA é alterado em disco)
DGAV_TEMPLATE_PATH = Path(__file__).parent / "DGAV_SAMPLE_REGISTRATION_FILE_XYLELLA.xlsx"


# Mapeamento entre colunas do pré-registo e colunas DGAV
# keys -> nome da coluna DGAV (na folha "Default")
# values -> nome (humano) da coluna no pré-registo
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

# Colunas DGAV obrigatórias – modo 2: erro se QUALQUER célula estiver vazia
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


# ───────────────────────────────────────────────
# Normalização "inteligente" de nomes de colunas
# ───────────────────────────────────────────────
def _norm(text: str | None) -> str:
    """
    Normaliza nomes de colunas para comparação tolerante:
      - trata None como string vazia
      - converte para str
      - remove acentos
      - passa para minúsculas
      - converte NBSP para espaço normal
      - substitui '_' e '-' por espaço
      - comprime espaços múltiplos num só
      - faz strip
    """
    if text is None:
        return ""
    s = str(text)

    # NBSP -> espaço normal
    s = s.replace("\u00A0", " ")

    # remover acentos
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")

    # underscores e hífens como espaço
    s = s.replace("_", " ").replace("-", " ")

    # para minúsculas
    s = s.lower()

    # comprimir espaços múltiplos
    s = " ".join(s.split())

    return s.strip()


# ───────────────────────────────────────────────
# Leitura do pré-registo (com fórmulas calculadas)
# ───────────────────────────────────────────────
def _find_header_row(df_raw: pd.DataFrame, target: str) -> int:
    """Encontra a linha de cabeçalho no ficheiro de pré-registo."""
    target_norm = _norm(target)

    # 1ª passagem: match exato (normalizado)
    for idx, row in df_raw.iterrows():
        if row.astype(str).apply(_norm).eq(target_norm).any():
            return idx

    # 2ª passagem: procura por substring "codigo_amostra"
    for idx, row in df_raw.iterrows():
        if row.astype(str).apply(_norm).str.contains("codigo amostra", na=False).any():
            return idx

    raise ValueError("Não foi possível identificar a linha de cabeçalho no pré-registo.")


def _load_pre_registo_df(uploaded_file) -> pd.DataFrame:
    """
    Lê os valores do pré-registo já calculados (data_only=True) e
    devolve DataFrame com cabeçalhos originais.
    """
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


def _map_input_columns(df: pd.DataFrame) -> Dict[str, str]:
    """
    Cria um mapa:
        { dgav_col_name -> nome_coluna_df }
    usando matching tolerante entre os nomes esperados em INPUT_TO_DGAV_COLMAP
    e os cabeçalhos do DataFrame.
    """
    # mapa normalizado -> nome real
    norm_to_real: Dict[str, str] = {}
    for col in df.columns:
        norm_to_real[_norm(col)] = col

    mapped: Dict[str, str] = {}

    for dgav_col, input_label in INPUT_TO_DGAV_COLMAP.items():
        key_norm = _norm(input_label)
        real = norm_to_real.get(key_norm)

        if real is None:
            # não encontrou – pode ser template diferente; deixamos sem mapear
            # e o valor ficará None
            # (se for obrigatório, depois a validação acusa)
            mapped[dgav_col] = None
        else:
            mapped[dgav_col] = real

    return mapped


# ───────────────────────────────────────────────
# Utilitários para o template DGAV
# ───────────────────────────────────────────────
def _build_header_index(ws) -> Dict[str, int]:
    """Mapeia nome de coluna DGAV normalizado -> índice de coluna."""
    header_indices: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v:
            header_indices[_norm(v)] = col
    return header_indices


def _mark_required_empty_columns(ws, header_indices: Dict[str, int], start_row: int, last_row: int):
    """
    Pinta de vermelho colunas obrigatórias em que QUALQUER célula esteja vazia
    entre start_row e last_row (modo 2).
    """
    red = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

    for col_name in REQUIRED_DGAV_COLS:
        col_idx = header_indices.get(_norm(col_name))
        if col_idx is None:
            # coluna obrigatória nem sequer existe na folha
            continue

        any_empty = False
        for r in range(start_row, last_row + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v in (None, ""):
                any_empty = True
                break

        if any_empty:
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

    # Criar mapa tolerante entre colunas do DF e labels esperados
    input_colmap = _map_input_columns(df_in)

    # 2) Carregar template DGAV sempre fresco (bytes → BytesIO)
    template_bytes = DGAV_TEMPLATE_PATH.read_bytes()
    template_stream = BytesIO(template_bytes)
    wb = load_workbook(template_stream)
    ws = wb["Default"]

    header_indices = _build_header_index(ws)

    # Guardar valores da linha 2 original (defaults do template)
    base_values: Dict[str, object] = {}
    for norm_name, col_idx in header_indices.items():
        base_values[norm_name] = ws.cell(row=2, column=col_idx).value

    # 3) APAGAR TODAS AS LINHAS A PARTIR DA LINHA 2
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # 4) Escrever uma linha por amostra
    start_row = 2
    last_row = start_row + n_samples - 1 if n_samples > 0 else 1

    for i, (_, row_in) in enumerate(df_in.iterrows()):
        excel_row = start_row + i

        # 4.1 Preencher linha com defaults do template (linha 2)
        for norm_name, col_idx in header_indices.items():
            ws.cell(row=excel_row, column=col_idx).value = base_values.get(norm_name)

        # 4.2 Substituir colunas do pré-registo (mantendo defaults se não existir no DF)
        for dgav_col, input_label in INPUT_TO_DGAV_COLMAP.items():
            col_idx = header_indices.get(_norm(dgav_col))
            if col_idx is None:
                continue

            df_col_name = input_colmap.get(dgav_col)

            # Se a coluna não existir no pré-registo → manter default
            if not df_col_name:
                continue

            value = row_in.get(df_col_name)

            # Converte NaN em None
            if isinstance(value, float) and pd.isna(value):
                value = None

            # Remove hora se for datetime
            if hasattr(value, "date"):
                value = value.date()

            ws.cell(row=excel_row, column=col_idx).value = value

    # 5) Validar colunas obrigatórias (modo 2)
    if n_samples > 0:
        _mark_required_empty_columns(ws, header_indices, start_row=start_row, last_row=last_row)

    # 6) Cortar linhas extra (garantir que só existem as linhas necessárias)
    if ws.max_row > last_row:
        ws.delete_rows(last_row + 1, ws.max_row - last_row)
    
    # 6B) Encolher também a tabela do Excel
    # (senão o Excel "estica" as fórmulas/defaults até ao tamanho antigo da tabela)
    for table in ws._tables:
        start, end = table.ref.split(':')
        start_col = ''.join(filter(str.isalpha, start))
        start_row = int(''.join(filter(str.isdigit, start)))
        end_col = ''.join(filter(str.isalpha, end))
    
        table.ref = f"{start_col}{start_row}:{end_col}{last_row}"

    # 7) Exportar para bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue(), f"Foram processadas {n_samples} amostras."
