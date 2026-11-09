import re
import io
import os
import pandas as pd
import numpy as np
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def normalize_str(val):
    """Normalize string for fuzzy matching (lowercase, stripped, no accents)."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip().lower()
    accents = {
        'á': 'a', 'à': 'a', 'ã': 'a', 'â': 'a', 'ä': 'a',
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
        'ó': 'o', 'ò': 'o', 'õ': 'o', 'ô': 'o', 'ö': 'o',
        'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
        'ç': 'c'
    }
    for k, v in accents.items():
        s = s.replace(k, v)
    return s

def parse_bloqueado(val):
    """Normalize block status string to SIM or NÃO."""
    if pd.isna(val) or val is None:
        return "NÃO"
    n = normalize_str(val)
    if n in ["sim", "s", "true", "1", "bloqueado"]:
        return "SIM"
    elif n in ["nao", "n", "false", "0", "desbloqueado"]:
        return "NÃO"
    return str(val).strip().upper()

def parse_date_safe(val):
    """Safely parse date from Excel serial number, ISO string or Brazilian date string."""
    if pd.isna(val) or val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'nat']:
        return None
    try:
        try:
            num = float(s)
            return pd.to_datetime('1899-12-30') + pd.Timedelta(days=num)
        except (ValueError, TypeError):
            pass
        
        if re.match(r'^\d{4}-\d{2}-\d{2}', s):
            return pd.to_datetime(s, format='%Y-%m-%d', errors='coerce')
        elif re.match(r'^\d{2}/\d{2}/\d{4}', s):
            return pd.to_datetime(s, format='%d/%m/%Y', errors='coerce')
        else:
            return pd.to_datetime(s, dayfirst=True, errors='coerce')
    except Exception:
        return None

def get_bytes_stream(file_obj):
    """Creates a fresh BytesIO copy from file_obj (filepath, Streamlit UploadedFile, or BytesIO)."""
    if isinstance(file_obj, str):
        with open(file_obj, "rb") as f:
            return io.BytesIO(f.read())
    elif hasattr(file_obj, 'getvalue'):
        return io.BytesIO(file_obj.getvalue())
    elif hasattr(file_obj, 'read'):
        file_obj.seek(0)
        return io.BytesIO(file_obj.read())
    return file_obj

def load_excel_smart_header(file_obj, target_keywords):
    """
    Scans the first 25 rows of an Excel file to find the actual header row index.
    Ignores single-cell title banners and prioritizes rows with multiple matching column names.
    """
    stream_preview = get_bytes_stream(file_obj)
    try:
        df_raw = pd.read_excel(stream_preview, header=None, nrows=25)
    except Exception:
        stream_fallback = get_bytes_stream(file_obj)
        return pd.read_excel(stream_fallback)

    norm_target = [normalize_str(k) for k in target_keywords]
    best_row_idx = 0
    max_score = 0

    for row_idx in range(len(df_raw)):
        row_vals = [normalize_str(val) for val in df_raw.iloc[row_idx].values if pd.notna(val) and str(val).strip() != ""]
        
        if len(row_vals) < 2:
            continue

        score = 0
        for cell_val in row_vals:
            for kw in norm_target:
                if cell_val == kw:
                    score += 3
                elif cell_val.startswith(kw) or kw.startswith(cell_val):
                    score += 2
                elif kw in cell_val and len(cell_val) < 35:
                    score += 1
        
        if score > max_score:
            max_score = score
            best_row_idx = row_idx

    stream_final = get_bytes_stream(file_obj)
    if max_score > 0:
        df = pd.read_excel(stream_final, header=best_row_idx)
    else:
        df = pd.read_excel(stream_final)

    cleaned_cols = []
    for i, c in enumerate(df.columns):
        if pd.notna(c) and str(c).strip() != "" and not str(c).startswith("Unnamed:"):
            cleaned_cols.append(str(c).strip())
        else:
            cleaned_cols.append(f"Unnamed_{i}")
    df.columns = cleaned_cols
    return df

def find_column(df, possible_names, default=None):
    """Find the best matching column name from a list of possibilities."""
    normalized_cols = {normalize_str(col): col for col in df.columns}
    
    for name in possible_names:
        norm_name = normalize_str(name)
        if norm_name in normalized_cols:
            return normalized_cols[norm_name]
    
    for name in possible_names:
        norm_name = normalize_str(name)
        for col_norm, original_col in normalized_cols.items():
            if col_norm.startswith(norm_name) or norm_name.startswith(col_norm):
                return original_col

    for name in possible_names:
        norm_name = normalize_str(name)
        for col_norm, original_col in normalized_cols.items():
            if norm_name in col_norm or col_norm in norm_name:
                return original_col

    return default

def parse_month_year_to_date(row, col_mes, col_ano):
    """Convert month and year columns into a standard datetime object (1st of month)."""
    val_mes = row[col_mes]
    val_ano = row[col_ano]
    
    if pd.isna(val_mes) or pd.isna(val_ano):
        return pd.NaT
        
    try:
        ano = int(float(str(val_ano).strip()))
    except (ValueError, TypeError):
        return pd.NaT

    month_map = {
        'jan': 1, 'janeiro': 1, '01': 1, '1': 1,
        'fev': 2, 'fevereiro': 2, '02': 2, '2': 2,
        'mar': 3, 'marco': 3, '03': 3, '3': 3,
        'abr': 4, 'abril': 4, '04': 4, '4': 4,
        'mai': 5, 'maio': 5, '05': 5, '5': 5,
        'jun': 6, 'junho': 6, '06': 6, '6': 6,
        'jul': 7, 'julho': 7, '07': 7, '7': 7,
        'ago': 8, 'agosto': 8, '08': 8, '8': 8,
        'set': 9, 'setembro': 9, '09': 9, '9': 9,
        'out': 10, 'outubro': 10, '10': 10,
        'nov': 11, 'novembro': 11, '11': 11,
        'dez': 12, 'dezembro': 12, '12': 12
    }
    
    norm_mes = normalize_str(val_mes)
    mes = month_map.get(norm_mes)
    
    if mes is None:
        try:
            mes = int(float(str(val_mes).strip()))
            if not (1 <= mes <= 12):
                return pd.NaT
        except (ValueError, TypeError):
            return pd.NaT
            
    try:
        return pd.Timestamp(year=ano, month=mes, day=1)
    except Exception:
        return pd.NaT

def calculate_period_months(start_date, end_date):
    """Calculate the number of months in a date range."""
    d1 = pd.to_datetime(start_date)
    d2 = pd.to_datetime(end_date)
    
    if d2 < d1:
        return 1.0
        
    days_diff = (d2 - d1).days + 1
    
    if d1.day == 1 and (d2 + pd.Timedelta(days=1)).day == 1:
        num_months = (d2.year - d1.year) * 12 + (d2.month - d1.month) + 1
        return float(max(1, num_months))
    
    approx_months = round(days_diff / 30.4375, 2)
    return float(max(1.0, approx_months))

def calculate_effective_months(dt_cad, p_start, p_end, default_period_months):
    """
    Calcula a quantidade de meses efetiva para o rateio de médias de um produto.
    Se a data de cadastro estiver dentro do período (dt_cad > p_start e dt_cad <= p_end),
    o fator de divisão é a quantidade de meses entre dt_cad e p_end.
    Caso contrário, retorna default_period_months.
    """
    dt = parse_date_safe(dt_cad)
    if dt is None or pd.isna(dt):
        return float(default_period_months)

    d_start = pd.to_datetime(p_start)
    d_end = pd.to_datetime(p_end)

    # Se cadastrado após o fim do período ou antes/no primeiro dia do período:
    if dt <= d_start or dt > d_end:
        return float(default_period_months)

    # Cadastrado dentro do período analisado:
    days_diff = (d_end - dt).days + 1
    
    # Se cadastrado no primeiro dia de algum mês até o último dia de um mês:
    if dt.day == 1 and (d_end + pd.Timedelta(days=1)).day == 1:
        months = (d_end.year - dt.year) * 12 + (d_end.month - dt.month) + 1
        return float(max(0.5, months))

    # Proporcional em dias (30.4375 dias por mês em média):
    approx = round(days_diff / 30.4375, 2)
    
    # Se muito próximo de meio mês (ex: 47 dias / 30.4375 = 1.54 -> arredonda para 1.5):
    nearest_half = round(approx * 2) / 2
    if abs(approx - nearest_half) <= 0.08:
        approx = nearest_half

    return float(max(0.1, approx))

def process_data(file_cad, file_orc, file_op, periods, file_vendas=None):
    """
    Core data processor function that ingests excel files and list of periods,
    and returns a tuple of (consolidated_df, excel_bytes, metrics_summary).
    - file_cad: Planilha 1 (Cadastro de Produtos)
    - file_orc: Planilha 2 (Orçamentos)
    - file_op: Planilha 3 (Consumo em Fabricação OP)
    - file_vendas: Planilha 4 (Planilha de Vendas - faturamento/devoluções Protheus D2_COD, D2_QUANT, F2_EMISSAO)
    """
    keywords_cad = ['CODIGO', 'CÓDIGO', 'BLOQUEIO', 'BLQ', 'DESCRICAO', 'DESCRIÇÃO', 'FABRICANTE', 'PP', 'LE', 'ES', 'CODIGO_ALTERNATIVO', 'CUSTO_STAND', 'PRAZO_DE_ENTREGA', 'GRUPO', 'DESCRCAO_GRUPO']
    keywords_orc = ['COD PRODUTOS', 'CODIGO', 'CÓDIGO', 'PRODUTO', 'MES', 'MÊS', 'ANO', 'QTDE', 'STATUS']
    keywords_op = ['PRODUTO', 'PROD', 'CODIGO', 'CÓDIGO', 'TIPO RE/DE', 'DT EMISSAO', 'EMISSAO', 'QUANTIDADE', 'RE/DE', 'CONSUMO']
    keywords_vendas = ['D2_COD', 'D2_QUANT', 'F2_EMISSAO', 'CODIGO', 'CÓDIGO', 'QUANTIDADE', 'EMISSAO']

    df_cad_raw = load_excel_smart_header(file_cad, keywords_cad)
    df_orc_raw = load_excel_smart_header(file_orc, keywords_orc)
    df_op_raw = load_excel_smart_header(file_op, keywords_op)
    df_vendas_raw = load_excel_smart_header(file_vendas, keywords_vendas) if file_vendas is not None else None

    # 2. PROCESS PLANILHA 1 (CADASTRO) - INCLUI PRODUTOS ATIVOS E BLOQUEADOS
    col_blq = find_column(df_cad_raw, ['BLOQUEIO', 'BLQ DE TELA', 'BLQ. DE TELA', 'BLOQUEADO', 'STATUS BLQ'])
    col_cod_cad = find_column(df_cad_raw, ['CODIGO', 'CÓDIGO', 'COD PRODUTO', 'COD. PRODUTO', 'PRODUTO', 'ITEM'])
    col_cod_alt = find_column(df_cad_raw, ['CODIGO ALTERNATIVO', 'CODIGO_ALTERNATIVO', 'CÓDIGO ALTERNATIVO'])
    col_desc = find_column(df_cad_raw, ['DESCRICAO', 'DESCRIÇÃO', 'DESCRICAO DO PRODUTO'])
    col_grupo = find_column(df_cad_raw, ['GRUPO'])
    col_desc_grupo = find_column(df_cad_raw, ['DESCRICAO GRUPO', 'DESCRIÇÃO GRUPO', 'DESCRCAO_GRUPO', 'DESCRCAO GRUPO'])
    col_dt_cad = find_column(df_cad_raw, ['DT CADASTRO', 'DATA CADASTRO', 'DT. CADASTRO', 'CADASTRO'])
    col_fab = find_column(df_cad_raw, ['FABRICANTE', 'MARCA'])
    col_custo_stand = find_column(df_cad_raw, ['CUSTO STAND', 'CUSTO_STAND'])
    col_prazo_ent = find_column(df_cad_raw, ['PRAZO DE ENTREGA', 'PRAZO_DE_ENTREGA', 'PRAZO PARA ENTREGA'])
    col_pp = find_column(df_cad_raw, ['PP', 'PP ATUAL', 'PONTO DE PEDIDO'])
    col_le = find_column(df_cad_raw, ['LE', 'LE ATUAL', 'LOTE ECONOMICO', 'LOTE ECONÔMICO'])
    col_es = find_column(df_cad_raw, ['ES', 'ES ATUAL', 'ESTOQUE DE SEGURANCA', 'ESTOQUE DE SEGURANÇA'])

    if not col_cod_cad:
        cols_found = [str(c) for c in df_cad_raw.columns if not str(c).startswith("Unnamed")]
        raise ValueError(f"Planilha 1 (Cadastro) não possui coluna de Código identificável. Colunas detectadas: {cols_found}")

    df_cad = df_cad_raw.copy()
    
    # Padroniza chave de junção do produto
    df_cad['COD_JOIN'] = df_cad[col_cod_cad].astype(str).str.strip().str.upper()
    df_cad = df_cad.drop_duplicates(subset=['COD_JOIN'])

    # Padroniza coluna BLOQUEADO (SIM / NÃO) - sem filtrar fora os bloqueados!
    if col_blq:
        df_cad['BLOQUEADO_CLEAN'] = df_cad[col_blq].apply(parse_bloqueado)
    else:
        df_cad['BLOQUEADO_CLEAN'] = "NÃO"

    # Padroniza CÓDIGO ALTERNATIVO (se vazio, assume o próprio código)
    if col_cod_alt:
        df_cad['COD_ALT_CLEAN'] = df_cad.apply(
            lambda r: str(r[col_cod_alt]).strip().upper()
            if pd.notna(r[col_cod_alt]) and str(r[col_cod_alt]).strip() not in ["", "nan", "NAN", "None"]
            else str(r[col_cod_cad]).strip().upper(),
            axis=1
        )
    else:
        df_cad['COD_ALT_CLEAN'] = df_cad['COD_JOIN']

    base_df = pd.DataFrame({
        'COD_JOIN': df_cad['COD_JOIN'],
        'CÓDIGO': df_cad[col_cod_cad].astype(str).str.strip(),
        'CÓDIGO ALTERNATIVO': df_cad['COD_ALT_CLEAN'],
        'DESCRIÇÃO': df_cad[col_desc].astype(str).str.strip() if col_desc else "",
        'BLOQUEADO': df_cad['BLOQUEADO_CLEAN'],
        'GRUPO': df_cad[col_grupo].astype(str).str.strip() if col_grupo else "",
        'DESCRIÇÃO GRUPO': df_cad[col_desc_grupo].astype(str).str.strip() if col_desc_grupo else "",
        'DATA CADASTRO': df_cad[col_dt_cad] if col_dt_cad else "",
        'FABRICANTE': df_cad[col_fab].astype(str).str.strip() if col_fab else "",
        'CUSTO STAND': df_cad[col_custo_stand] if col_custo_stand else 0,
        'PRAZO PARA ENTREGA': df_cad[col_prazo_ent] if col_prazo_ent else 0,
        'PP': df_cad[col_pp] if col_pp else 0,
        'LE': df_cad[col_le] if col_le else 0,
        'ES': df_cad[col_es] if col_es else 0,
    })

    # 3. PROCESS PLANILHA 2 (ORÇAMENTOS E VENDAS)
    col_cod_orc = find_column(df_orc_raw, ['COD PRODUTOS', 'COD PRODUTO', 'COD. PRODUTO', 'CODIGO', 'CÓDIGO', 'PRODUTO', 'ITEM'])
    col_mes_orc = find_column(df_orc_raw, ['MES', 'MÊS'])
    col_ano_orc = find_column(df_orc_raw, ['ANO'])
    col_dt_orc = find_column(df_orc_raw, ['DATA', 'DT ORCAMENTO', 'DT EMISSAO'])
    col_qtd_orc = find_column(df_orc_raw, ['QTDE DO ITEM', 'QTDE ITEM', 'QUANTIDADE', 'QTDE'])
    col_status_orc = find_column(df_orc_raw, ['COD. STATUS', 'COD STATUS', 'STATUS'])

    if not col_cod_orc:
        cols_found = [str(c) for c in df_orc_raw.columns if not str(c).startswith("Unnamed")]
        raise ValueError(f"Planilha 2 (Orçamentos/Vendas) não possui coluna de Código identificável. Colunas detectadas: {cols_found}")

    df_orc = df_orc_raw.copy()
    df_orc['COD_JOIN'] = df_orc[col_cod_orc].astype(str).str.strip().str.upper()

    if col_mes_orc and col_ano_orc:
        df_orc['DT_PROCESSED'] = df_orc.apply(lambda r: parse_month_year_to_date(r, col_mes_orc, col_ano_orc), axis=1)
    elif col_dt_orc:
        df_orc['DT_PROCESSED'] = pd.to_datetime(df_orc[col_dt_orc], errors='coerce')
    else:
        df_orc['DT_PROCESSED'] = pd.NaT

    if col_qtd_orc:
        df_orc['QTDE_CLEAN'] = pd.to_numeric(df_orc[col_qtd_orc], errors='coerce').fillna(0)
    else:
        df_orc['QTDE_CLEAN'] = 0

    if col_status_orc:
        df_orc['STATUS_CLEAN'] = df_orc[col_status_orc].astype(str).str.strip().str.upper()
    else:
        df_orc['STATUS_CLEAN'] = ""

    # 4. PROCESS PLANILHA 3 (CONSUMO OP)
    possible_op_prod_cols = [
        'PRODUTO', 'PROD', 'COD PRODUTO', 'COD. PRODUTO', 'CÓD. PRODUTO', 
        'CODIGO PRODUTO', 'CÓDIGO PRODUTO', 'CODIGO DO PRODUTO', 'CÓDIGO DO PRODUTO',
        'COD PROD', 'COD. PROD.', 'CÓD. PROD.', 'CD_PRODUTO', 'CD_PROD', 'COD_PROD',
        'PRODUTO/SERVIÇO', 'PRODUTO / SERVIÇO', 'CODIGO', 'CÓDIGO', 'ITEM', 'CD_ITEM', 'COMPONENTE'
    ]
    col_cod_op = find_column(df_op_raw, possible_op_prod_cols)
    
    col_tipo_op = find_column(df_op_raw, ['TIPO RE/DE', 'TIPO RE / DE', 'RE/DE', 'TIPO', 'RE / DE'])
    col_dt_op = find_column(df_op_raw, ['DT EMISSAO', 'DT EMISSÃO', 'DATA EMISSAO', 'DATA EMISSÃO', 'EMISSAO', 'EMISSÃO', 'DATA'])
    col_qtd_op = find_column(df_op_raw, ['QUANTIDADE', 'QTDE', 'QUANT', 'QTD'])

    if not col_cod_op:
        cols_found = [str(c) for c in df_op_raw.columns if not str(c).startswith("Unnamed")]
        raise ValueError(f"Planilha 3 (OP/Fabricação) não possui coluna de Produto identificável. Colunas detectadas: {cols_found}")

    df_op = df_op_raw.copy()
    df_op['COD_JOIN'] = df_op[col_cod_op].astype(str).str.strip().str.upper()

    # Filtro de Exclusão da OP: Desconsiderar Tipo RE/DE em ["PRO", "PR0", "DE0", "DEO"]
    if col_tipo_op:
        excluded_tipos = {"pro", "pr0", "de0", "deo"}
        df_op = df_op[~df_op[col_tipo_op].apply(lambda x: normalize_str(x).replace(" ", "") in excluded_tipos)]

    if col_dt_op:
        df_op['DT_PROCESSED'] = pd.to_datetime(df_op[col_dt_op], errors='coerce')
    else:
        df_op['DT_PROCESSED'] = pd.NaT

    if col_qtd_op:
        df_op['QTDE_CLEAN'] = pd.to_numeric(df_op[col_qtd_op], errors='coerce').fillna(0)
    else:
        df_op['QTDE_CLEAN'] = 0

    # 4.1 PROCESS PLANILHA 4 (PLANILHA DE VENDAS) - SE FORNECIDA
    df_vendas = None
    if df_vendas_raw is not None:
        col_cod_vendas = find_column(df_vendas_raw, ['D2_COD', 'CODIGO', 'CÓDIGO', 'COD PRODUTO', 'PRODUTO', 'ITEM'])
        col_qtd_vendas = find_column(df_vendas_raw, ['D2_QUANT', 'QUANTIDADE', 'QTDE', 'QTD VENDIDA'])
        col_dt_vendas = find_column(df_vendas_raw, ['F2_EMISSAO', 'EMISSAO', 'EMISSÃO', 'DATA EMISSAO', 'DATA'])

        if not col_cod_vendas:
            cols_found = [str(c) for c in df_vendas_raw.columns if not str(c).startswith("Unnamed")]
            raise ValueError(f"Planilha 4 (Planilha de Vendas) não possui coluna de Código (D2_COD) identificável. Colunas detectadas: {cols_found}")

        df_vendas = df_vendas_raw.copy()
        df_vendas['COD_JOIN'] = df_vendas[col_cod_vendas].astype(str).str.strip().str.upper()

        if col_dt_vendas:
            df_vendas['DT_PROCESSED'] = pd.to_datetime(df_vendas[col_dt_vendas], errors='coerce')
        else:
            df_vendas['DT_PROCESSED'] = pd.NaT

        if col_qtd_vendas:
            df_vendas['QTDE_CLEAN'] = pd.to_numeric(df_vendas[col_qtd_vendas], errors='coerce').fillna(0)
        else:
            df_vendas['QTDE_CLEAN'] = 0

    # 5. CALCULATE PERIOD METRICS FOR MAIN DF
    result_df = base_df.copy()
    period_summaries = []

    for idx, p in enumerate(periods):
        p_label = p.get('label') or f"PERÍODO {idx+1}"
        p_start = pd.to_datetime(p['start'])
        p_end = pd.to_datetime(p['end'])
        
        num_months = calculate_period_months(p_start, p_end)
        
        # 1. ORÇAMENTOS: vêm da Planilha 2 (Orçamentos)
        mask_orc = (df_orc['DT_PROCESSED'] >= p_start) & (df_orc['DT_PROCESSED'] <= p_end)
        df_orc_p = df_orc[mask_orc]
        orc_grouped = df_orc_p.groupby('COD_JOIN')['QTDE_CLEAN'].sum()

        # 2. VENDAS: se fornecida a Planilha 4 (Planilha de Vendas), vêm dela considerando devoluções (D2_QUANT negativos)
        if df_vendas is not None:
            mask_vendas = (df_vendas['DT_PROCESSED'] >= p_start) & (df_vendas['DT_PROCESSED'] <= p_end)
            df_vendas_p = df_vendas[mask_vendas]
            vend_grouped = df_vendas_p.groupby('COD_JOIN')['QTDE_CLEAN'].sum()
        else:
            # Fallback histórico para cálculo via Planilha 2 com STATUS == 'B'
            mask_vend = mask_orc & (df_orc['STATUS_CLEAN'] == 'B')
            df_vend_p = df_orc[mask_vend]
            vend_grouped = df_vend_p.groupby('COD_JOIN')['QTDE_CLEAN'].sum()

        # 3. CONSUMO OP: vem da Planilha 3 (Detalhamento OP)
        mask_op = (df_op['DT_PROCESSED'] >= p_start) & (df_op['DT_PROCESSED'] <= p_end)
        df_op_p = df_op[mask_op]
        op_grouped = df_op_p.groupby('COD_JOIN')['QTDE_CLEAN'].sum()

        # Map metrics to base products
        q_orc = result_df['COD_JOIN'].map(orc_grouped).fillna(0)
        q_vend = result_df['COD_JOIN'].map(vend_grouped).fillna(0)
        
        # Fator de meses efetivo por produto considerando DATA CADASTRO
        eff_months_cad = result_df['DATA CADASTRO'].apply(
            lambda dt: calculate_effective_months(dt, p_start, p_end, num_months)
        )
        
        q_vend_mes = q_vend / eff_months_cad
        efetividade = np.where(q_orc > 0, q_vend / q_orc, 0.0)
        
        q_cons = result_df['COD_JOIN'].map(op_grouped).fillna(0)
        q_cons_mes = q_cons / eff_months_cad

        prefix = f"[{p_label}] "
        result_df[prefix + 'QUANT. ORÇADA'] = q_orc
        result_df[prefix + 'QUANT. VENDIDA'] = q_vend
        result_df[prefix + 'QUANT. VENDIDA MÊS'] = q_vend_mes
        result_df[prefix + 'EFETIVIDADE (%)'] = efetividade
        result_df[prefix + 'QUANT. CONSUMIDA'] = q_cons
        result_df[prefix + 'QUANT. CONSUMIDA MÊS'] = q_cons_mes

        period_summaries.append({
            'label': p_label,
            'range': f"{p_start.strftime('%d/%m/%Y')} A {p_end.strftime('%d/%m/%Y')}",
            'months': num_months,
            'total_orcada': q_orc.sum(),
            'total_vendida': q_vend.sum(),
            'total_consumida': q_cons.sum(),
            'avg_efetividade': (q_vend.sum() / q_orc.sum() * 100) if q_orc.sum() > 0 else 0.0
        })

    # 5.1 BUILD ALTERNATIVE CODE AGGREGATION DATAFRAME (ABA 2)
    alt_groups = []
    last_p = periods[-1]
    last_p_label = last_p.get('label') or f"PERÍODO {len(periods)}"
    last_prefix = f"[{last_p_label}] "

    # Dicionário mapeando o CÓDIGO limpo para sua respectiva linha no cadastro.
    # Garante que descrição, grupo, descrição do grupo, data de cadastro, fabricante, prazo, PP, LE e ES
    # sejam rigorosamente iguais aos dados do código que está sendo analisado (CÓDIGO == CÓDIGO ALTERNATIVO)
    cad_by_code = {}
    for _, r in result_df.iterrows():
        cod_key = str(r['CÓDIGO']).strip().upper()
        if cod_key and cod_key not in cad_by_code:
            cad_by_code[cod_key] = r

    for cod_alt, grp in result_df.groupby('CÓDIGO ALTERNATIVO', sort=True):
        clean_alt = str(cod_alt).strip().upper()

        # O código que está sendo analisado é prioritariamente o produto cujo CÓDIGO seja o próprio CÓDIGO ALTERNATIVO
        if clean_alt in cad_by_code:
            rep_row = cad_by_code[clean_alt]
        else:
            rep_row = grp.iloc[0]

        dt_cad_alt = rep_row['DATA CADASTRO']

        row_dict = {
            'CÓDIGO ALTERNATIVO': cod_alt,
            'DESCRIÇÃO': rep_row['DESCRIÇÃO'],
            'GRUPO': rep_row['GRUPO'],
            'DESCRIÇÃO GRUPO': rep_row['DESCRIÇÃO GRUPO'],
            'DATA CADASTRO': dt_cad_alt,
            'QTD. CÓDIGOS': len(grp),
            'FABRICANTE': rep_row['FABRICANTE'],
            'PRAZO PARA ENTREGA': rep_row['PRAZO PARA ENTREGA'],
            'PP': rep_row['PP'],
            'LE': rep_row['LE'],
            'ES': rep_row['ES'],
        }

        for idx, p in enumerate(periods):
            p_label = p.get('label') or f"PERÍODO {idx+1}"
            prefix = f"[{p_label}] "
            p_start = pd.to_datetime(p['start'])
            p_end = pd.to_datetime(p['end'])
            num_months = calculate_period_months(p_start, p_end)
            
            # Fator de meses efetivo para o código alternativo
            eff_months_alt = calculate_effective_months(dt_cad_alt, p_start, p_end, num_months)

            q_orc = grp[prefix + 'QUANT. ORÇADA'].sum()
            q_vend = grp[prefix + 'QUANT. VENDIDA'].sum()
            q_vend_mes = q_vend / eff_months_alt
            efetividade = (q_vend / q_orc) if q_orc > 0 else 0.0
            q_cons = grp[prefix + 'QUANT. CONSUMIDA'].sum()
            q_cons_mes = q_cons / eff_months_alt

            row_dict[prefix + 'QUANT. ORÇADA'] = q_orc
            row_dict[prefix + 'QUANT. VENDIDA'] = q_vend
            row_dict[prefix + 'QUANT. VENDIDA MÊS'] = q_vend_mes
            row_dict[prefix + 'EFETIVIDADE (%)'] = efetividade
            row_dict[prefix + 'QUANT. CONSUMIDA'] = q_cons
            row_dict[prefix + 'QUANT. CONSUMIDA MÊS'] = q_cons_mes

        # Cálculo dos parâmetros sugeridos (com base no último período)
        v_mes_last = row_dict[last_prefix + 'QUANT. VENDIDA MÊS']
        c_mes_last = row_dict[last_prefix + 'QUANT. CONSUMIDA MÊS']
        es_sug = round(v_mes_last + c_mes_last)

        try:
            prazo_num = float(rep_row['PRAZO PARA ENTREGA']) if pd.notna(rep_row['PRAZO PARA ENTREGA']) and str(rep_row['PRAZO PARA ENTREGA']).strip() != '' else 0.0
        except (ValueError, TypeError):
            prazo_num = 0.0

        le_sug = round(es_sug * prazo_num)
        pp_sug = round(le_sug + es_sug)

        try:
            pp_atual = float(rep_row['PP']) if pd.notna(rep_row['PP']) else 0.0
        except (ValueError, TypeError):
            pp_atual = 0.0

        if pp_atual == 0:
            var_pp = 1.0 if pp_sug > 0 else 0.0
            analisar = "ANALISAR" if pp_sug > 0 else "OK"
        else:
            var_pp = (pp_sug - pp_atual) / pp_atual
            if pp_sug >= pp_atual * 1.2 or pp_sug <= pp_atual * 0.8:
                analisar = "ANALISAR"
            else:
                analisar = "OK"

        row_dict['ES SUGERIDO'] = es_sug
        row_dict['LE SUGERIDO'] = le_sug
        row_dict['PP SUGERIDO'] = pp_sug
        row_dict['VARIAÇÃO PP (%)'] = var_pp
        row_dict['ANALISAR PP?'] = analisar

        alt_groups.append(row_dict)

    df_alt_result = pd.DataFrame(alt_groups)

    result_df = result_df.drop(columns=['COD_JOIN'])

    # 6. GENERATE STYLED OPENPYXL EXCEL WITH 2 SHEETS AND NATIVE FORMULAS
    excel_bytes = generate_styled_excel(result_df, df_alt_result, periods)

    total_analisar = int((df_alt_result['ANALISAR PP?'] == 'ANALISAR').sum()) if 'ANALISAR PP?' in df_alt_result.columns else 0

    metrics_summary = {
        'total_products': len(result_df),
        'total_alt_products': len(df_alt_result),
        'total_analisar_pp': total_analisar,
        'df_alt': df_alt_result,
        'periods': period_summaries
    }

    return result_df, excel_bytes, metrics_summary

def generate_styled_excel(df, df_alt, periods):
    """Creates a corporate Excel workbook with 2 sheets and dynamic formulas using OpenPyXL."""
    wb = openpyxl.Workbook()
    
    COLOR_PRIMARY_HEADER = "1F4E78"    # Navy Dark Blue
    COLOR_PERIOD_HEADER = "1B365D"     # Deep Corporate Navy
    COLOR_SUBHEADER = "2C4D75"         # Slate Blue
    COLOR_ZEBRA_LIGHT = "F8FAFC"       # Very light slate/gray
    COLOR_BORDER = "CBD5E1"            # Light gray border
    COLOR_SUG_HEADER = "0E7490"        # Dark Teal / Cyan
    COLOR_SUG_SUB = "155E75"           # Deep Teal Subheader

    font_header_main = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_header_sub = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=10)
    font_blocked = Font(name="Segoe UI", size=10, bold=True, color="991B1B")

    fill_main_cad = PatternFill(start_color=COLOR_PRIMARY_HEADER, end_color=COLOR_PRIMARY_HEADER, fill_type="solid")
    fill_period = PatternFill(start_color=COLOR_PERIOD_HEADER, end_color=COLOR_PERIOD_HEADER, fill_type="solid")
    fill_sub = PatternFill(start_color=COLOR_SUBHEADER, end_color=COLOR_SUBHEADER, fill_type="solid")
    fill_zebra = PatternFill(start_color=COLOR_ZEBRA_LIGHT, end_color=COLOR_ZEBRA_LIGHT, fill_type="solid")
    fill_sug_main = PatternFill(start_color=COLOR_SUG_HEADER, end_color=COLOR_SUG_HEADER, fill_type="solid")
    fill_sug_sub = PatternFill(start_color=COLOR_SUG_SUB, end_color=COLOR_SUG_SUB, fill_type="solid")

    thin_border_side = Side(border_style="thin", color=COLOR_BORDER)
    border_data = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    period_subheaders = ['QUANT. ORÇADA', 'QUANT. VENDIDA', 'QUANT. VENDIDA MÊS', 'EFETIVIDADE (%)', 'QUANT. CONSUMIDA', 'QUANT. CONSUMIDA MÊS']
    sub_cols_per_period = 6

    # =========================================================================
    # ABA 1: ANÁLISE CONSOLIDADA (DETALHADA POR PRODUTO)
    # =========================================================================
    ws = wb.active
    ws.title = "Análise Consolidada"
    ws.views.sheetView[0].showGridLines = True

    base_headers_s1 = [
        'CÓDIGO', 'CÓDIGO ALTERNATIVO', 'DESCRIÇÃO', 'BLOQUEADO', 'GRUPO', 
        'DESCRIÇÃO GRUPO', 'DATA CADASTRO', 'FABRICANTE', 'CUSTO STAND', 
        'PRAZO PARA ENTREGA', 'PP', 'LE', 'ES'
    ]
    base_col_count_s1 = len(base_headers_s1) # 13 colunas

    # ROW 1 (Aba 1): Top Merged Headers
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=base_col_count_s1)
    cell_cad = ws.cell(row=1, column=1, value="DADOS CADASTRAIS DO PRODUTO")
    cell_cad.font = font_header_main
    cell_cad.fill = fill_main_cad
    cell_cad.alignment = align_center

    for c in range(1, base_col_count_s1 + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = fill_main_cad
        cell.border = border_data

    curr_col = base_col_count_s1 + 1
    for idx, p in enumerate(periods):
        p_label = p.get('label') or f"PERÍODO {idx+1}"
        p_start = pd.to_datetime(p['start']).strftime('%d/%m/%Y')
        p_end = pd.to_datetime(p['end']).strftime('%d/%m/%Y')
        header_text = f"{p_label.upper()} ({p_start} A {p_end})"

        end_c = curr_col + sub_cols_per_period - 1
        ws.merge_cells(start_row=1, start_column=curr_col, end_row=1, end_column=end_c)
        cell_p = ws.cell(row=1, column=curr_col, value=header_text)
        cell_p.font = font_header_main
        cell_p.fill = fill_period
        cell_p.alignment = align_center

        for c in range(curr_col, end_c + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = fill_period
            cell.border = border_data

        curr_col = end_c + 1

    # ROW 2 (Aba 1): Column Subheaders
    all_subheaders_s1 = list(base_headers_s1)
    for _ in periods:
        all_subheaders_s1.extend(period_subheaders)

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 25

    for c_idx, h_text in enumerate(all_subheaders_s1, start=1):
        cell = ws.cell(row=2, column=c_idx, value=h_text)
        cell.font = font_header_sub
        cell.fill = fill_sub
        cell.alignment = align_center
        cell.border = border_data

    # ROW 3+ (Aba 1): Data Rows com Fórmulas Dinâmicas
    for r_idx, row_data in enumerate(df.itertuples(index=False), start=3):
        ws.row_dimensions[r_idx].height = 20
        is_even = (r_idx % 2 == 0)

        # Preenche colunas cadastrais
        for c_idx in range(1, base_col_count_s1 + 1):
            val = row_data[c_idx - 1]
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = border_data
            if is_even:
                cell.fill = fill_zebra
            cell.font = font_data

            col_name = base_headers_s1[c_idx - 1]
            if col_name in ['DESCRIÇÃO', 'DESCRIÇÃO GRUPO']:
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = align_left
            elif col_name == 'BLOQUEADO':
                cell.value = str(val) if pd.notna(val) else "NÃO"
                cell.alignment = align_center
                if str(val).strip().upper() == "SIM":
                    cell.font = font_blocked
            elif col_name == 'DATA CADASTRO':
                dt_parsed = parse_date_safe(val)
                if dt_parsed is not None and pd.notna(dt_parsed):
                    cell.value = dt_parsed.to_pydatetime()
                    cell.number_format = 'dd/mm/yyyy'
                else:
                    cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = align_center
            elif col_name in ['PP', 'LE', 'ES', 'PRAZO PARA ENTREGA']:
                try:
                    cell.value = int(val) if pd.notna(val) else 0
                    cell.number_format = '#,##0'
                except (ValueError, TypeError):
                    cell.value = val
                cell.alignment = align_right
            elif col_name == 'CUSTO STAND':
                try:
                    cell.value = float(val) if pd.notna(val) else 0.0
                    cell.number_format = '#,##0.00'
                except (ValueError, TypeError):
                    cell.value = val
                cell.alignment = align_right
            else:
                cell.value = str(val) if pd.notna(val) else ""
                cell.alignment = align_center

        # Preenche colunas de períodos com fórmulas
        dt_cad_prod = row_data[base_headers_s1.index('DATA CADASTRO')]

        for p_idx, p in enumerate(periods):
            p_start = pd.to_datetime(p['start'])
            p_end = pd.to_datetime(p['end'])
            num_months = calculate_period_months(p_start, p_end)
            
            # Fator de meses efetivo para este produto
            eff_m = calculate_effective_months(dt_cad_prod, p_start, p_end, num_months)
            eff_m_val = int(eff_m) if float(eff_m).is_integer() else (f"{eff_m:.1f}" if round(eff_m, 1) == round(eff_m, 2) else f"{eff_m:.2f}")

            col_p_start = base_col_count_s1 + 1 + p_idx * sub_cols_per_period
            val_offset = base_col_count_s1 + p_idx * sub_cols_per_period

            col_orc_idx = col_p_start
            col_vend_idx = col_p_start + 1
            col_vend_mes_idx = col_p_start + 2
            col_efet_idx = col_p_start + 3
            col_cons_idx = col_p_start + 4
            col_cons_mes_idx = col_p_start + 5

            col_orc_let = get_column_letter(col_orc_idx)
            col_vend_let = get_column_letter(col_vend_idx)
            col_cons_let = get_column_letter(col_cons_idx)

            raw_orc = row_data[val_offset]
            raw_vend = row_data[val_offset + 1]
            raw_cons = row_data[val_offset + 4]

            # 1. QUANT. ORÇADA (Valor)
            c_orc = ws.cell(row=r_idx, column=col_orc_idx)
            c_orc.value = float(raw_orc) if pd.notna(raw_orc) else 0.0
            c_orc.number_format = '#,##0'
            c_orc.alignment = align_right
            c_orc.border = border_data
            c_orc.font = font_data
            if is_even: c_orc.fill = fill_zebra

            # 2. QUANT. VENDIDA (Valor)
            c_vend = ws.cell(row=r_idx, column=col_vend_idx)
            c_vend.value = float(raw_vend) if pd.notna(raw_vend) else 0.0
            c_vend.number_format = '#,##0'
            c_vend.alignment = align_right
            c_vend.border = border_data
            c_vend.font = font_data
            if is_even: c_vend.fill = fill_zebra

            # 3. QUANT. VENDIDA MÊS (Fórmula: =VENDIDA / MESES_EFETIVOS)
            c_vmes = ws.cell(row=r_idx, column=col_vend_mes_idx)
            c_vmes.value = f"={col_vend_let}{r_idx}/{eff_m_val}"
            c_vmes.number_format = '#,##0.00'
            c_vmes.alignment = align_right
            c_vmes.border = border_data
            c_vmes.font = font_data
            if is_even: c_vmes.fill = fill_zebra

            # 4. EFETIVIDADE (%) (Fórmula: =SEERRO(VENDIDA / ORCADA; 0))
            c_efet = ws.cell(row=r_idx, column=col_efet_idx)
            c_efet.value = f"=IFERROR({col_vend_let}{r_idx}/{col_orc_let}{r_idx}, 0)"
            c_efet.number_format = '0.0%'
            c_efet.alignment = align_right
            c_efet.border = border_data
            c_efet.font = font_data
            if is_even: c_efet.fill = fill_zebra

            # 5. QUANT. CONSUMIDA (Valor)
            c_cons = ws.cell(row=r_idx, column=col_cons_idx)
            c_cons.value = float(raw_cons) if pd.notna(raw_cons) else 0.0
            c_cons.number_format = '#,##0'
            c_cons.alignment = align_right
            c_cons.border = border_data
            c_cons.font = font_data
            if is_even: c_cons.fill = fill_zebra

            # 6. QUANT. CONSUMIDA MÊS (Fórmula: =CONSUMIDA / MESES_EFETIVOS)
            c_cmes = ws.cell(row=r_idx, column=col_cons_mes_idx)
            c_cmes.value = f"={col_cons_let}{r_idx}/{eff_m_val}"
            c_cmes.number_format = '#,##0.00'
            c_cmes.alignment = align_right
            c_cmes.border = border_data
            c_cmes.font = font_data
            if is_even: c_cmes.fill = fill_zebra

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row == 1:
                continue
            val_str = str(cell.value or '')
            if cell.number_format == '0.0%' and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value * 100:.1f}%"
            elif cell.number_format in ['#,##0', '#,##0.00'] and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value:,.2f}"
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    # =========================================================================
    # ABA 2: CONSOLIDADO POR CÓDIGO ALTERNATIVO (TABELA DINÂMICA COM SOMASE)
    # =========================================================================
    ws2 = wb.create_sheet(title="Consolidado Cód. Alternativo")
    ws2.views.sheetView[0].showGridLines = True

    base_headers_s2 = [
        'CÓDIGO ALTERNATIVO', 'DESCRIÇÃO', 'GRUPO', 'DESCRIÇÃO GRUPO', 
        'DATA CADASTRO', 'QTD. CÓDIGOS', 'FABRICANTE', 'PRAZO PARA ENTREGA', 'PP', 'LE', 'ES'
    ]
    base_col_count_s2 = len(base_headers_s2) # 11 colunas

    suggested_headers = [
        'ES SUGERIDO', 'LE SUGERIDO', 'PP SUGERIDO', 'VARIAÇÃO PP (%)', 'ANALISAR PP?'
    ]
    sug_col_count = len(suggested_headers) # 5 colunas

    # ROW 1 (Aba 2): Top Merged Headers
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=base_col_count_s2)
    cell_cad2 = ws2.cell(row=1, column=1, value="DADOS DO CÓDIGO ALTERNATIVO")
    cell_cad2.font = font_header_main
    cell_cad2.fill = fill_main_cad
    cell_cad2.alignment = align_center

    for c in range(1, base_col_count_s2 + 1):
        cell = ws2.cell(row=1, column=c)
        cell.fill = fill_main_cad
        cell.border = border_data

    curr_col2 = base_col_count_s2 + 1
    for idx, p in enumerate(periods):
        p_label = p.get('label') or f"PERÍODO {idx+1}"
        p_start = pd.to_datetime(p['start']).strftime('%d/%m/%Y')
        p_end = pd.to_datetime(p['end']).strftime('%d/%m/%Y')
        header_text = f"{p_label.upper()} ({p_start} A {p_end})"

        end_c = curr_col2 + sub_cols_per_period - 1
        ws2.merge_cells(start_row=1, start_column=curr_col2, end_row=1, end_column=end_c)
        cell_p = ws2.cell(row=1, column=curr_col2, value=header_text)
        cell_p.font = font_header_main
        cell_p.fill = fill_period
        cell_p.alignment = align_center

        for c in range(curr_col2, end_c + 1):
            cell = ws2.cell(row=1, column=c)
            cell.fill = fill_period
            cell.border = border_data

        curr_col2 = end_c + 1

    # Cabeçalho da Seção de Sugestão de Parâmetros
    end_sug = curr_col2 + sug_col_count - 1
    ws2.merge_cells(start_row=1, start_column=curr_col2, end_row=1, end_column=end_sug)
    cell_sug = ws2.cell(row=1, column=curr_col2, value="PARÂMETROS SUGERIDOS (ÚLTIMO PERÍODO)")
    cell_sug.font = font_header_main
    cell_sug.fill = fill_sug_main
    cell_sug.alignment = align_center

    for c in range(curr_col2, end_sug + 1):
        cell = ws2.cell(row=1, column=c)
        cell.fill = fill_sug_main
        cell.border = border_data

    # ROW 2 (Aba 2): Column Subheaders
    all_subheaders_s2 = list(base_headers_s2)
    for _ in periods:
        all_subheaders_s2.extend(period_subheaders)
    all_subheaders_s2.extend(suggested_headers)

    ws2.row_dimensions[1].height = 28
    ws2.row_dimensions[2].height = 25

    for c_idx, h_text in enumerate(all_subheaders_s2, start=1):
        cell = ws2.cell(row=2, column=c_idx, value=h_text)
        cell.font = font_header_sub
        if c_idx >= curr_col2:
            cell.fill = fill_sug_sub
        else:
            cell.fill = fill_sub
        cell.alignment = align_center
        cell.border = border_data

    # ROW 3+ (Aba 2): Data Rows com Fórmulas SOMASE referenciando Aba 1
    last_row_s1 = len(df) + 2 # Última linha de dados na Aba 1
    last_p_idx = len(periods) - 1
    col_vend_mes_last_idx = base_col_count_s2 + 1 + last_p_idx * sub_cols_per_period + 2
    col_cons_mes_last_idx = base_col_count_s2 + 1 + last_p_idx * sub_cols_per_period + 5
    let_vmes_last = get_column_letter(col_vend_mes_last_idx)
    let_cmes_last = get_column_letter(col_cons_mes_last_idx)

    # Letras das colunas de Prazo e PP Atual na Aba 2
    let_prazo = get_column_letter(8)    # PRAZO PARA ENTREGA (Col H = 8)
    let_pp_atual = get_column_letter(9) # PP (Col I = 9)

    c_es_sug_idx = curr_col2
    c_le_sug_idx = curr_col2 + 1
    c_pp_sug_idx = curr_col2 + 2
    c_var_idx = curr_col2 + 3
    c_status_idx = curr_col2 + 4

    let_es_sug = get_column_letter(c_es_sug_idx)
    let_le_sug = get_column_letter(c_le_sug_idx)
    let_pp_sug = get_column_letter(c_pp_sug_idx)

    for r_idx, row_data in enumerate(df_alt.itertuples(index=False), start=3):
        ws2.row_dimensions[r_idx].height = 20
        is_even = (r_idx % 2 == 0)

        # 1. Dados cadastrais do Alternativo
        for c_idx in range(1, base_col_count_s2 + 1):
            val = row_data[c_idx - 1]
            cell = ws2.cell(row=r_idx, column=c_idx)
            cell.border = border_data
            if is_even:
                cell.fill = fill_zebra
            cell.font = font_data

            col_name = base_headers_s2[c_idx - 1]
            if col_name in ['DESCRIÇÃO', 'DESCRIÇÃO GRUPO']:
                cell.value = str(val).strip() if pd.notna(val) else ""
                cell.alignment = align_left
            elif col_name == 'DATA CADASTRO':
                dt_parsed = parse_date_safe(val)
                if dt_parsed is not None and pd.notna(dt_parsed):
                    cell.value = dt_parsed.to_pydatetime()
                    cell.number_format = 'dd/mm/yyyy'
                else:
                    cell.value = str(val).strip() if pd.notna(val) else ""
                cell.alignment = align_center
            elif col_name in ['QTD. CÓDIGOS', 'PRAZO PARA ENTREGA', 'PP', 'LE', 'ES']:
                try:
                    cell.value = int(val) if pd.notna(val) else 0
                    cell.number_format = '#,##0'
                except (ValueError, TypeError):
                    cell.value = val
                cell.alignment = align_right
            else:
                cell.value = str(val).strip() if pd.notna(val) else ""
                cell.alignment = align_center

        # 2. Colunas de Períodos com SOMASE e Fórmulas
        dt_cad_alt = row_data[base_headers_s2.index('DATA CADASTRO')]

        for p_idx, p in enumerate(periods):
            p_start = pd.to_datetime(p['start'])
            p_end = pd.to_datetime(p['end'])
            num_months = calculate_period_months(p_start, p_end)
            
            # Fator de meses efetivo para o alternativo
            eff_m = calculate_effective_months(dt_cad_alt, p_start, p_end, num_months)
            eff_m_val = int(eff_m) if float(eff_m).is_integer() else (f"{eff_m:.1f}" if round(eff_m, 1) == round(eff_m, 2) else f"{eff_m:.2f}")

            col_s2_start = base_col_count_s2 + 1 + p_idx * sub_cols_per_period
            col_s1_start = base_col_count_s1 + 1 + p_idx * sub_cols_per_period

            # Índices e letras na Aba 1 (Análise Consolidada)
            col_orc_s1_let = get_column_letter(col_s1_start)
            col_vend_s1_let = get_column_letter(col_s1_start + 1)
            col_cons_s1_let = get_column_letter(col_s1_start + 4)

            # Índices e letras na Aba 2 (Consolidado Cód. Alternativo)
            col_orc_s2_idx = col_s2_start
            col_vend_s2_idx = col_s2_start + 1
            col_vend_mes_s2_idx = col_s2_start + 2
            col_efet_s2_idx = col_s2_start + 3
            col_cons_s2_idx = col_s2_start + 4
            col_cons_mes_s2_idx = col_s2_start + 5

            col_orc_s2_let = get_column_letter(col_orc_s2_idx)
            col_vend_s2_let = get_column_letter(col_vend_s2_idx)
            col_cons_s2_let = get_column_letter(col_cons_s2_idx)

            # 1. QUANT. ORÇADA (Fórmula: =SOMASE('Análise Consolidada'!$B$3:$B$FIM; A{LINHA}; 'Análise Consolidada'!$COL_ORÇADA$3:$COL_ORÇADA$FIM))
            c_orc = ws2.cell(row=r_idx, column=col_orc_s2_idx)
            c_orc.value = f"=SUMIF('Análise Consolidada'!$B$3:$B${last_row_s1}, A{r_idx}, 'Análise Consolidada'!${col_orc_s1_let}$3:${col_orc_s1_let}${last_row_s1})"
            c_orc.number_format = '#,##0'
            c_orc.alignment = align_right
            c_orc.border = border_data
            c_orc.font = font_data
            if is_even: c_orc.fill = fill_zebra

            # 2. QUANT. VENDIDA (Fórmula: =SOMASE(...))
            c_vend = ws2.cell(row=r_idx, column=col_vend_s2_idx)
            c_vend.value = f"=SUMIF('Análise Consolidada'!$B$3:$B${last_row_s1}, A{r_idx}, 'Análise Consolidada'!${col_vend_s1_let}$3:${col_vend_s1_let}${last_row_s1})"
            c_vend.number_format = '#,##0'
            c_vend.alignment = align_right
            c_vend.border = border_data
            c_vend.font = font_data
            if is_even: c_vend.fill = fill_zebra

            # 3. QUANT. VENDIDA MÊS (Fórmula: =VENDIDA / MESES_EFETIVOS)
            c_vmes = ws2.cell(row=r_idx, column=col_vend_mes_s2_idx)
            c_vmes.value = f"={col_vend_s2_let}{r_idx}/{eff_m_val}"
            c_vmes.number_format = '#,##0.00'
            c_vmes.alignment = align_right
            c_vmes.border = border_data
            c_vmes.font = font_data
            if is_even: c_vmes.fill = fill_zebra

            # 4. EFETIVIDADE (%) (Fórmula: =SEERRO(VENDIDA / ORCADA; 0))
            c_efet = ws2.cell(row=r_idx, column=col_efet_s2_idx)
            c_efet.value = f"=IFERROR({col_vend_s2_let}{r_idx}/{col_orc_s2_let}{r_idx}, 0)"
            c_efet.number_format = '0.0%'
            c_efet.alignment = align_right
            c_efet.border = border_data
            c_efet.font = font_data
            if is_even: c_efet.fill = fill_zebra

            # 5. QUANT. CONSUMIDA (Fórmula: =SOMASE(...))
            c_cons = ws2.cell(row=r_idx, column=col_cons_s2_idx)
            c_cons.value = f"=SUMIF('Análise Consolidada'!$B$3:$B${last_row_s1}, A{r_idx}, 'Análise Consolidada'!${col_cons_s1_let}$3:${col_cons_s1_let}${last_row_s1})"
            c_cons.number_format = '#,##0'
            c_cons.alignment = align_right
            c_cons.border = border_data
            c_cons.font = font_data
            if is_even: c_cons.fill = fill_zebra

            # 6. QUANT. CONSUMIDA MÊS (Fórmula: =CONSUMIDA / MESES_EFETIVOS)
            c_cmes = ws2.cell(row=r_idx, column=col_cons_mes_s2_idx)
            c_cmes.value = f"={col_cons_s2_let}{r_idx}/{eff_m_val}"
            c_cmes.number_format = '#,##0.00'
            c_cmes.alignment = align_right
            c_cmes.border = border_data
            c_cmes.font = font_data
            if is_even: c_cmes.fill = fill_zebra

        # 3. Colunas de Parâmetros Sugeridos
        # 3.1 ES SUGERIDO: =ARRED(VENDIDA_MES_ULTIMO + CONSUMIDA_MES_ULTIMO; 0)
        c_es = ws2.cell(row=r_idx, column=c_es_sug_idx)
        c_es.value = f"=ROUND({let_vmes_last}{r_idx}+{let_cmes_last}{r_idx}, 0)"
        c_es.number_format = '#,##0'
        c_es.alignment = align_right
        c_es.border = border_data
        c_es.font = font_data
        if is_even: c_es.fill = fill_zebra

        # 3.2 LE SUGERIDO: =ARRED(ES_SUGERIDO * PRAZO_ENTREGA; 0)
        c_le = ws2.cell(row=r_idx, column=c_le_sug_idx)
        c_le.value = f"=ROUND({let_es_sug}{r_idx}*{let_prazo}{r_idx}, 0)"
        c_le.number_format = '#,##0'
        c_le.alignment = align_right
        c_le.border = border_data
        c_le.font = font_data
        if is_even: c_le.fill = fill_zebra

        # 3.3 PP SUGERIDO: =ARRED(LE_SUGERIDO + ES_SUGERIDO; 0)
        c_pp = ws2.cell(row=r_idx, column=c_pp_sug_idx)
        c_pp.value = f"=ROUND({let_le_sug}{r_idx}+{let_es_sug}{r_idx}, 0)"
        c_pp.number_format = '#,##0'
        c_pp.alignment = align_right
        c_pp.border = border_data
        c_pp.font = font_data
        if is_even: c_pp.fill = fill_zebra

        # 3.4 VARIAÇÃO PP (%): =(PP_SUGERIDO - PP_ATUAL) / PP_ATUAL
        c_var = ws2.cell(row=r_idx, column=c_var_idx)
        c_var.value = f"=IF({let_pp_atual}{r_idx}=0, IF({let_pp_sug}{r_idx}>0, 1, 0), ({let_pp_sug}{r_idx}-{let_pp_atual}{r_idx})/{let_pp_atual}{r_idx})"
        c_var.number_format = '+0.0%;-0.0%;0.0%'
        c_var.alignment = align_right
        c_var.border = border_data
        c_var.font = font_data
        if is_even: c_var.fill = fill_zebra

        # 3.5 ANALISAR PP?: se variação >= 20% ou <= -20% -> ANALISAR, senão OK
        c_status = ws2.cell(row=r_idx, column=c_status_idx)
        c_status.value = f"=IF({let_pp_atual}{r_idx}=0, IF({let_pp_sug}{r_idx}>0, \"ANALISAR\", \"OK\"), IF(OR({let_pp_sug}{r_idx}>={let_pp_atual}{r_idx}*1.2, {let_pp_sug}{r_idx}<={let_pp_atual}{r_idx}*0.8), \"ANALISAR\", \"OK\"))"
        c_status.alignment = align_center
        c_status.border = border_data
        
        # Estilo de alerta visual se marcado como ANALISAR
        val_status_precalc = str(row_data[-1]) if len(row_data) >= 1 else ""
        if val_status_precalc == "ANALISAR":
            c_status.font = font_blocked
        else:
            c_status.font = font_data
        if is_even: c_status.fill = fill_zebra

    for col in ws2.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row == 1:
                continue
            val_str = str(cell.value or '')
            if cell.number_format in ['0.0%', '+0.0%;-0.0%;0.0%'] and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value * 100:.1f}%"
            elif cell.number_format in ['#,##0', '#,##0.00'] and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value:,.2f}"
            max_len = max(max_len, len(val_str))
        ws2.column_dimensions[col_letter].width = max(max_len + 4, 13)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
