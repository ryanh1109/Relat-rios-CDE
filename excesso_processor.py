import re
import io
import os
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from data_processor import load_excel_smart_header, find_column

def process_excesso_data(file_cad, file_pos=None, include_details=False):
    """
    Processes product registration file and 'Posição Estoque' file for 'Excesso em Estoque' report.
    Columns generated:
      - CÓDIGO
      - DESCRIÇÃO
      - GRUPO
      - (Optional) PP, LE, ES
      - ESTOQUE MÁXIMO (PP + LE + ES)
      - (If file_pos) SALDO ATUAL (soma de SALDO EM ESTOQUE, desconsiderando armazém 51)
      - (If file_pos) EXCESSO EM ESTOQUE (fórmula: SALDO ATUAL - ESTOQUE MÁXIMO)
      - (If file_pos) VALOR EM ESTOQUE (soma de VALOR EM ESTOQUE, desconsiderando armazém 51)
      - (If file_pos) VALOR EM EXCESSO (fórmula: (VALOR EM ESTOQUE / SALDO ATUAL) * EXCESSO EM ESTOQUE)
    Returns (df_result, excel_bytes, metrics).
    """
    # 1. PROCESSAR PLANILHA 1: CADASTRO DE PRODUTOS
    keywords_cad = ['CODIGO', 'CÓDIGO', 'DESCRICAO', 'DESCRIÇÃO', 'GRUPO', 'PP', 'LE', 'ES']
    df_cad_raw = load_excel_smart_header(file_cad, keywords_cad)

    col_cod = find_column(df_cad_raw, ['CODIGO', 'CÓDIGO', 'COD PRODUTO', 'COD. PRODUTO', 'PRODUTO', 'ITEM'])
    col_desc = find_column(df_cad_raw, ['DESCRICAO', 'DESCRIÇÃO', 'DESCRICAO DO PRODUTO'])
    col_grupo = find_column(df_cad_raw, ['GRUPO', 'DESCRICAO GRUPO', 'DESCRIÇÃO GRUPO'])
    col_pp = find_column(df_cad_raw, ['PP', 'PP ATUAL', 'PONTO DE PEDIDO'])
    col_le = find_column(df_cad_raw, ['LE', 'LE ATUAL', 'LOTE ECONOMICO', 'LOTE ECONÔMICO'])
    col_es = find_column(df_cad_raw, ['ES', 'ES ATUAL', 'ESTOQUE DE SEGURANCA', 'ESTOQUE DE SEGURANÇA'])

    if not col_cod:
        cols_found = [str(c) for c in df_cad_raw.columns if not str(c).startswith("Unnamed")]
        raise ValueError(f"Planilha de Cadastro não possui coluna de Código identificável. Colunas detectadas: {cols_found}")

    df = df_cad_raw.copy()
    
    # Padroniza código e remove duplicidades
    df['COD_CLEAN'] = df[col_cod].astype(str).str.strip().str.upper()
    df = df[df['COD_CLEAN'] != ""]
    df = df[~df['COD_CLEAN'].isin(["NAN", "NONE", "NAT"])]
    df = df.drop_duplicates(subset=['COD_CLEAN'])

    desc_series = df[col_desc].fillna("").astype(str).str.strip() if col_desc else ""
    grupo_series = df[col_grupo].fillna("").astype(str).str.strip() if col_grupo else ""

    pp_series = pd.to_numeric(df[col_pp], errors='coerce').fillna(0) if col_pp else pd.Series(0, index=df.index)
    le_series = pd.to_numeric(df[col_le], errors='coerce').fillna(0) if col_le else pd.Series(0, index=df.index)
    es_series = pd.to_numeric(df[col_es], errors='coerce').fillna(0) if col_es else pd.Series(0, index=df.index)

    estoque_maximo = pp_series + le_series + es_series

    # 2. PROCESSAR PLANILHA 2: POSIÇÃO DE ESTOQUE (SE FORNECIDA)
    saldo_mapped = None
    valor_mapped = None

    if file_pos is not None:
        keywords_pos = ['PRODUTO', 'PROD', 'CODIGO', 'CÓDIGO', 'DESCRICAO', 'DESCRIÇÃO', 'ARMZ', 'ARMAZEM', 'LOCAL', 'SALDO EM ESTOQUE', 'SALDO', 'VALOR EM ESTOQUE', 'VALOR']
        df_pos_raw = load_excel_smart_header(file_pos, keywords_pos)

        col_cod_pos = find_column(df_pos_raw, ['CODIGO', 'CÓDIGO', 'COD PRODUTO', 'COD. PRODUTO', 'PRODUTO', 'ITEM', 'CD_PRODUTO'])
        col_saldo_pos = find_column(df_pos_raw, ['SALDO EM ESTOQUE', 'SALDO ATUAL', 'SALDO', 'QUANTIDADE', 'QTDE ATUAL', 'QTDE', 'SLD_ATUAL'])
        col_valor_pos = find_column(df_pos_raw, ['VALOR EM ESTOQUE', 'VALOR ATUAL', 'VALOR', 'CUSTO TOTAL', 'VALOR TOTAL'])
        col_arm_pos = find_column(df_pos_raw, ['ARMZ', 'ARMAZEM', 'ARMAZÉM', 'LOCAL', 'ALMOXARIFADO', 'ARM'])

        if not col_cod_pos:
            cols_found = [str(c) for c in df_pos_raw.columns if not str(c).startswith("Unnamed")]
            raise ValueError(f"Planilha Posição de Estoque não possui coluna de Código identificável. Colunas detectadas: {cols_found}")

        df_pos = df_pos_raw.copy()

        # Filtrar fora o armazém 51
        if col_arm_pos:
            arm_series = df_pos[col_arm_pos].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_pos = df_pos[~arm_series.isin(['51', '051'])]

        df_pos['COD_CLEAN'] = df_pos[col_cod_pos].astype(str).str.strip().str.upper()

        if col_saldo_pos:
            df_pos['SALDO_CLEAN'] = pd.to_numeric(df_pos[col_saldo_pos], errors='coerce').fillna(0)
            saldo_grouped = df_pos.groupby('COD_CLEAN')['SALDO_CLEAN'].sum()
            saldo_mapped = df['COD_CLEAN'].map(saldo_grouped).fillna(0)
        else:
            saldo_mapped = pd.Series(0.0, index=df.index)

        if col_valor_pos:
            df_pos['VALOR_CLEAN'] = pd.to_numeric(df_pos[col_valor_pos], errors='coerce').fillna(0)
            valor_grouped = df_pos.groupby('COD_CLEAN')['VALOR_CLEAN'].sum()
            valor_mapped = df['COD_CLEAN'].map(valor_grouped).fillna(0)
        else:
            valor_mapped = pd.Series(0.0, index=df.index)

    # 3. CONSTRUIR DATAFRAME FINAL
    data_dict = {
        'CÓDIGO': df['COD_CLEAN'],
        'DESCRIÇÃO': desc_series,
        'GRUPO': grupo_series,
    }

    if include_details:
        data_dict['PP'] = pp_series
        data_dict['LE'] = le_series
        data_dict['ES'] = es_series

    data_dict['ESTOQUE MÁXIMO'] = estoque_maximo

    if saldo_mapped is not None:
        data_dict['SALDO ATUAL'] = saldo_mapped
        # Excesso em Estoque = Saldo Atual - Estoque Máximo
        data_dict['EXCESSO EM ESTOQUE'] = saldo_mapped - estoque_maximo
        if valor_mapped is not None:
            data_dict['VALOR EM ESTOQUE'] = valor_mapped
            # Valor em Excesso = (Valor em Estoque / Saldo Atual) * Excesso em Estoque
            custo_unit = np.where(saldo_mapped > 0, valor_mapped / saldo_mapped, 0.0)
            data_dict['VALOR EM EXCESSO'] = custo_unit * data_dict['EXCESSO EM ESTOQUE']

    df_result = pd.DataFrame(data_dict)

    # 4. GERAR ARQUIVO EXCEL ESTILIZADO COM FÓRMULAS
    excel_bytes = generate_excesso_excel(df_result, include_details=include_details, has_saldo=(saldo_mapped is not None))

    # 5. RESUMO DE MÉTRICAS EXECUTIVAS
    metrics = {
        'total_products': int(len(df_result)),
        'total_with_max': int((estoque_maximo > 0).sum()),
        'total_zero_max': int((estoque_maximo == 0).sum()),
        'sum_max': float(estoque_maximo.sum()),
        'avg_max': float(estoque_maximo.mean()) if len(df_result) > 0 else 0.0,
        'total_groups': int(pd.Series(grupo_series).nunique()) if col_grupo else 0,
        'has_posicao': (saldo_mapped is not None)
    }

    if saldo_mapped is not None:
        excesso_series = df_result['EXCESSO EM ESTOQUE']
        excesso_pos = excesso_series.clip(lower=0)
        custo_unit_est = np.where(saldo_mapped > 0, valor_mapped / saldo_mapped, 0)
        valor_excesso_est = excesso_pos * custo_unit_est

        metrics.update({
            'total_saldo': float(saldo_mapped.sum()),
            'total_valor': float(valor_mapped.sum()) if valor_mapped is not None else 0.0,
            'total_itens_com_excesso': int((excesso_series > 0).sum()),
            'total_qtd_excesso': float(excesso_pos.sum()),
            'total_valor_excesso': float(valor_excesso_est.sum())
        })

    return df_result, excel_bytes, metrics

def generate_excesso_excel(df, include_details=False, has_saldo=False):
    """
    Generates a corporate OpenPyXL spreadsheet for 'Excesso em Estoque' with dynamic formulas.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Excesso em Estoque"
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = None  # Sem painéis congelados conforme padrão corporativo

    # Cores Corporativas (Tema Índigo / Roxo Escuro)
    COLOR_HEADER = "312E81"      # Deep Indigo
    COLOR_ZEBRA_LIGHT = "F8FAFC" # Slate clarinho
    COLOR_BORDER = "CBD5E1"      # Cinza claro

    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=10)
    font_excess = Font(name="Segoe UI", size=10, bold=True, color="991B1B")  # Vermelho para excesso
    
    fill_header = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type="solid")
    fill_zebra = PatternFill(start_color=COLOR_ZEBRA_LIGHT, end_color=COLOR_ZEBRA_LIGHT, fill_type="solid")
    fill_excess = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    thin_border_side = Side(border_style="thin", color=COLOR_BORDER)
    border_data = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    headers = list(df.columns)
    ws.row_dimensions[1].height = 28

    # Identificar letras de colunas dinamicamente
    col_max_let = get_column_letter(headers.index('ESTOQUE MÁXIMO') + 1) if 'ESTOQUE MÁXIMO' in headers else None
    col_saldo_let = get_column_letter(headers.index('SALDO ATUAL') + 1) if 'SALDO ATUAL' in headers else None
    col_excesso_let = get_column_letter(headers.index('EXCESSO EM ESTOQUE') + 1) if 'EXCESSO EM ESTOQUE' in headers else None
    col_valor_let = get_column_letter(headers.index('VALOR EM ESTOQUE') + 1) if 'VALOR EM ESTOQUE' in headers else None

    # Escrever Cabeçalho
    for col_idx, h_name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_data

    # Escrever Dados
    for r_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        ws.row_dimensions[r_idx].height = 20
        is_even = (r_idx % 2 == 0)

        for c_idx, col_name in enumerate(headers, start=1):
            val = row_data[c_idx - 1]
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = font_data
            cell.border = border_data
            if is_even:
                cell.fill = fill_zebra

            if col_name == 'CÓDIGO':
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = align_center
            elif col_name == 'DESCRIÇÃO':
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = align_left
            elif col_name == 'GRUPO':
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = align_center
            elif col_name in ['PP', 'LE', 'ES']:
                try:
                    cell.value = float(val) if pd.notna(val) else 0.0
                    cell.number_format = '#,##0'
                except (ValueError, TypeError):
                    cell.value = val
                cell.alignment = align_right
            elif col_name == 'ESTOQUE MÁXIMO':
                if include_details:
                    # Se tiver detalhes de PP (col D=4), LE (col E=5), ES (col F=6):
                    # Fórmula nativa: =D{r}+E{r}+F{r}
                    cell.value = f"=D{r_idx}+E{r_idx}+F{r_idx}"
                else:
                    try:
                        cell.value = float(val) if pd.notna(val) else 0.0
                    except (ValueError, TypeError):
                        cell.value = val
                cell.number_format = '#,##0'
                cell.alignment = align_right
            elif col_name == 'SALDO ATUAL':
                try:
                    cell.value = float(val) if pd.notna(val) else 0.0
                    cell.number_format = '#,##0'
                except (ValueError, TypeError):
                    cell.value = val
                cell.alignment = align_right
            elif col_name == 'EXCESSO EM ESTOQUE':
                # Fórmula nativa: =[SALDO ATUAL] - [ESTOQUE MÁXIMO]
                if col_saldo_let and col_max_let:
                    cell.value = f"={col_saldo_let}{r_idx}-{col_max_let}{r_idx}"
                else:
                    try:
                        cell.value = float(val) if pd.notna(val) else 0.0
                    except (ValueError, TypeError):
                        cell.value = val
                cell.number_format = '#,##0'
                cell.alignment = align_right

                # Destaque visual caso haja excesso positivo
                try:
                    val_num = float(val) if pd.notna(val) else 0.0
                    if val_num > 0:
                        cell.font = font_excess
                        cell.fill = fill_excess
                except (ValueError, TypeError):
                    pass

            elif col_name == 'VALOR EM ESTOQUE':
                try:
                    cell.value = float(val) if pd.notna(val) else 0.0
                    cell.number_format = 'R$ #,##0.00'
                except (ValueError, TypeError):
                    cell.value = val
                cell.alignment = align_right

            elif col_name == 'VALOR EM EXCESSO':
                # Fórmula nativa: =IFERROR(([VALOR EM ESTOQUE] / [SALDO ATUAL]) * [EXCESSO EM ESTOQUE], 0)
                if col_valor_let and col_saldo_let and col_excesso_let:
                    cell.value = f"=IFERROR(({col_valor_let}{r_idx}/{col_saldo_let}{r_idx})*{col_excesso_let}{r_idx}, 0)"
                else:
                    try:
                        cell.value = float(val) if pd.notna(val) else 0.0
                    except (ValueError, TypeError):
                        cell.value = val
                cell.number_format = 'R$ #,##0.00'
                cell.alignment = align_right

                # Destaque visual caso haja excesso financeiro positivo
                try:
                    val_num = float(val) if pd.notna(val) else 0.0
                    if val_num > 0:
                        cell.font = font_excess
                        cell.fill = fill_excess
                except (ValueError, TypeError):
                    pass

    # Ajustar largura das colunas automaticamente
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if 'R$' in str(cell.number_format) and isinstance(cell.value, (int, float)):
                val_str = f"R$ {cell.value:,.2f}"
            elif cell.number_format == '#,##0' and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value:,.0f}"
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
