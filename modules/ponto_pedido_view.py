import streamlit as st
import pandas as pd
from datetime import date, datetime
import io
import time
import os
import importlib
import storage_manager
import data_processor

def render_ponto_pedido_view():
    """
    Renders the complete Ponto de Pedido (Consolidador de Vendas & Fabricação) environment.
    """
    # 1. HEADER SECTION
    st.markdown("""
    <div class="header-container">
        <div class="header-title">
            📦 Relatório: Ponto de Pedido (Consolidador de Vendas & Fabricação)
        </div>
        <div class="header-subtitle">
            Automação de cruzamento de cadastro de produtos, itens orçados/vendidos e consumo de OPs com sugestão de PP, LE e ES por código alternativo.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 2. EXTRACTION INSTRUCTION GUIDE EXPANDER
    with st.expander("📖 Instruções de Extração e Preparação dos Arquivos (Clique para expandir)", expanded=False):
        st.markdown("##### 📋 Passo a Passo para Extração das Planilhas")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown("**🗂️ 1. Cadastro de Produtos**")
            st.markdown("""
1. Abrir o **SQL**
2. Atualizar a consulta
3. Copiar todas as colunas e linhas **sem formatação e fórmulas** para outra planilha
4. Salvar o arquivo no computador
""")

        with col2:
            st.markdown("**📋 2. Orçamentos**")
            st.markdown("""
1. Abrir o **SQL**
2. Atualizar a consulta
3. Selecionar o **período desejado**
4. Copiar todas as colunas e linhas para outra planilha
5. Salvar o arquivo no computador
""")

        with col3:
            st.markdown("**🏭 3. Consumo Fabricação (OP)**")
            st.markdown("""
1. Gerar o relatório **"Detalhamento da Ordem de Produção"** no Protheus
2. Selecionar a data desejada e **todas as filiais**
3. Após gerar, **excluir a primeira aba** (parâmetros)
4. Unificar na primeira aba restante e excluir as demais
5. Salvar o arquivo no computador
""")

        with col4:
            st.markdown("**🛒 4. Planilha de Vendas**")
            st.markdown("""
1. Gerar a extração de faturamento/vendas no **Protheus** (SD2/SF2)
2. Certificar-se de conter as colunas:
   - `D2_COD`: código do produto
   - `D2_QUANT`: quantidade vendida / devoluções (valores negativos)
   - `F2_EMISSAO`: data de emissão
3. Salvar o arquivo no computador
""")

        st.info("💡 **Dica de Organização:** É recomendado salvar todos os arquivos na mesma pasta do seu computador para facilitar a localização.")

    # 3. INITIALIZE SESSION STATE FOR DYNAMIC PERIODS
    if "periods" not in st.session_state:
        st.session_state.periods = [
            {
                "id": 1,
                "label": "1º Semestre 2024",
                "start": date(2024, 1, 1),
                "end": date(2024, 6, 30)
            },
            {
                "id": 2,
                "label": "2º Semestre 2024",
                "start": date(2024, 7, 1),
                "end": date(2024, 12, 31)
            }
        ]

    # 4. SIDEBAR - FILE UPLOAD WITH PERSISTENT STORAGE (ESPECÍFICO DE PONTO DE PEDIDO)
    with st.sidebar:
        st.markdown("### 📦 Arquivos: Ponto de Pedido")
        st.caption("Planilhas salvas no servidor local:")

        def render_file_section(title, file_key, help_text):
            st.markdown(f"#### {title}")
            info = storage_manager.get_file_info(file_key)
            uploaded = None
            
            if info:
                st.markdown(f"""
                <div class="saved-card">
                    <div class="saved-card-title">✅ Arquivo Salvo</div>
                    <div class="saved-card-sub">📄 <b>{info['original_name']}</b> ({info['size']})</div>
                    <div class="saved-card-sub">🕒 Atualizado: {info['last_updated']}</div>
                </div>
                """, unsafe_allow_html=True)
                
                c_up, c_del = st.columns([4, 1])
                with c_up:
                    uploaded = st.file_uploader(
                        f"🔄 Substituir {title}",
                        type=["xlsx", "xls"],
                        key=f"up_{file_key}",
                        help=help_text,
                        label_visibility="collapsed"
                    )
                with c_del:
                    if st.button("🗑️", key=f"del_{file_key}", help=f"Excluir {title} salva"):
                        storage_manager.delete_file(file_key)
                        st.rerun()
            else:
                uploaded = st.file_uploader(
                    f"Upload {title}",
                    type=["xlsx", "xls"],
                    key=f"up_{file_key}",
                    help=help_text
                )
                
            effective_file = storage_manager.get_effective_file(uploaded, file_key)
            return effective_file

        eff_cad = render_file_section(
            "1. Cadastro de Produtos",
            "planilha_1",
            "SQL: Copiar colunas/linhas sem formatação e fórmulas."
        )
        st.markdown("---")

        eff_orc = render_file_section(
            "2. Orçamentos",
            "planilha_2",
            "SQL: Atualizar período desejado e copiar colunas/linhas."
        )
        st.markdown("---")

        eff_op = render_file_section(
            "3. Consumo em Fabricação (OP)",
            "planilha_3",
            "Protheus: Detalhamento da OP, excluir 1ª aba de parâmetros e unificar demais abas."
        )
        st.markdown("---")

        eff_vendas = render_file_section(
            "4. Planilha de Vendas",
            "planilha_4",
            "Protheus: Faturamento e Devoluções com D2_COD, D2_QUANT e F2_EMISSAO."
        )

        st.markdown("---")
        if st.button("🧹 Limpar Planilhas Salvas deste Módulo", use_container_width=True):
            storage_manager.delete_file("planilha_1")
            storage_manager.delete_file("planilha_2")
            storage_manager.delete_file("planilha_3")
            storage_manager.delete_file("planilha_4")
            st.rerun()

    # 5. MAIN CONTENT - PERIOD CONFIGURATION
    col_main, col_summary = st.columns([7, 5])

    with col_main:
        st.subheader("📅 Configuração dos Períodos de Análise")
        st.caption("Adicione ou edite os intervalos de datas para comparar no relatório final:")

        col_add, col_space = st.columns([2, 5])
        with col_add:
            if st.button("➕ Adicionar Período", use_container_width=True):
                new_id = len(st.session_state.periods) + 1
                st.session_state.periods.append({
                    "id": new_id,
                    "label": f"Período {new_id}",
                    "start": date(2024, 1, 1),
                    "end": date(2024, 12, 31)
                })

        periods_to_remove = []

        for idx, p in enumerate(st.session_state.periods):
            with st.container():
                st.markdown(f"**📌 Configuração do {idx + 1}º Período**")
                p_col1, p_col2, p_col3, p_col4 = st.columns([3, 3, 3, 1])
                
                with p_col1:
                    p["label"] = st.text_input(
                        f"Nome/Rótulo",
                        value=p["label"],
                        key=f"label_{idx}"
                    )
                with p_col2:
                    p["start"] = st.date_input(
                        f"Data Inicial",
                        value=p["start"],
                        format="DD/MM/YYYY",
                        key=f"start_{idx}"
                    )
                with p_col3:
                    p["end"] = st.date_input(
                        f"Data Final",
                        value=p["end"],
                        format="DD/MM/YYYY",
                        key=f"end_{idx}"
                    )
                with p_col4:
                    st.write("")
                    st.write("")
                    if len(st.session_state.periods) > 1:
                        if st.button("🗑️", key=f"del_p_{idx}", help="Remover período"):
                            periods_to_remove.append(idx)

        if periods_to_remove:
            for idx in sorted(periods_to_remove, reverse=True):
                st.session_state.periods.pop(idx)
            st.rerun()

    with col_summary:
        st.subheader("📋 Resumo dos Parâmetros")
        st.markdown("""
        <div style="background-color: #FFFFFF; padding: 18px; border-radius: 10px; border: 1px solid #E2E8F0;">
            <p style="margin-bottom: 8px;"><strong>Regras de Negócio Aplicadas:</strong></p>
            <ul style="font-size: 13px; color: #475569; padding-left: 20px; margin-bottom: 0;">
                <li>Rateio Data Cadastro: Divisor proporcional pelo período em que o produto existiu (ex: 15/05 a 30/06 = /1.5).</li>
                <li>Filtro Cadastro: Inclui todos os produtos (ativos e bloqueados), indicados na coluna <b>BLOQUEADO</b>.</li>
                <li>Aba 2 (Alternativo): Consolidação por <b>CÓDIGO ALTERNATIVO</b> com fórmulas <b>SOMASE</b>.</li>
                <li>Sugestão de Parâmetros (Aba 2): Cálculo dinâmico de <b>ES, LE e PP Sugeridos</b> com base no último período.</li>
                <li>Alerta de Ponto de Pedido: Sinaliza <b>ANALISAR</b> se houver variação de <b>&ge; 20%</b> do PP sugerido vs atual.</li>
                <li>Filtro Vendas / OP: <b>COD. STATUS = "B"</b> e desconsidera <b>Tipo RE/DE = "PRO"/"PR0" ou "DE0"</b>.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("##### Períodos Selecionados:")
        for p in st.session_state.periods:
            st.markdown(f"- <span class='period-badge'>{p['label']}</span>: **{p['start'].strftime('%d/%m/%Y')}** a **{p['end'].strftime('%d/%m/%Y')}**", unsafe_allow_html=True)

    st.markdown("---")

    # 6. PROCESSING & DOWNLOAD SECTION
    ready_to_process = eff_cad is not None and eff_orc is not None and eff_op is not None and eff_vendas is not None

    col_proc, col_msg = st.columns([4, 6])

    with col_proc:
        process_btn = st.button(
            "🚀 Processar e Gerar Planilha Consolidada",
            type="primary",
            disabled=not ready_to_process,
            use_container_width=True
        )

    with col_msg:
        if not ready_to_process:
            missing = []
            if eff_cad is None: missing.append("1. Cadastro")
            if eff_orc is None: missing.append("2. Orçamentos")
            if eff_op is None: missing.append("3. Consumo OP")
            if eff_vendas is None: missing.append("4. Planilha de Vendas")
            st.warning(f"⚠️ Faltando planilhas no servidor/upload: **{', '.join(missing)}**.")
        else:
            st.success("✅ Todas as 4 planilhas estão prontas no servidor local! Clique no botão para iniciar.")

    # 7. EXECUTION LOGIC
    if process_btn and ready_to_process:
        progress_bar = st.progress(0, text="Iniciando leitura e validação dos arquivos...")
        
        try:
            importlib.reload(data_processor)
            importlib.reload(storage_manager)

            time.sleep(0.15)
            progress_bar.progress(20, text="Filtrando Cadastro de Produtos ativos e bloqueados...")
            time.sleep(0.15)
            progress_bar.progress(40, text="Processando Orçamentos por período...")
            time.sleep(0.15)
            progress_bar.progress(60, text="Processando Vendas e Devoluções da Planilha de Vendas...")
            time.sleep(0.15)
            progress_bar.progress(80, text="Calculando Consumo de OPs e gerando abas Excel com fórmulas...")

            df_result, excel_bytes, metrics = data_processor.process_data(
                file_cad=eff_cad,
                file_orc=eff_orc,
                file_op=eff_op,
                periods=st.session_state.periods,
                file_vendas=eff_vendas
            )

            progress_bar.progress(100, text="Concluído com sucesso!")
            time.sleep(0.3)
            progress_bar.empty()

            st.toast("🎉 Planilha consolidada com sucesso!", icon="✅")

            # 8. RESULTS DASHBOARD & PREVIEW
            st.subheader("📊 Resumo Executivo dos Resultados")

            m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
            m_col1.metric("Produtos Processados", f"{metrics['total_products']:,}")
            m_col2.metric("Códigos Alternativos", f"{metrics.get('total_alt_products', 0):,}")
            m_col3.metric("🚨 Analisar Ponto Pedido", f"{metrics.get('total_analisar_pp', 0):,}", help="Produtos com variação >= 20% no PP sugerido em relação ao atual")
            
            if metrics['periods']:
                p1 = metrics['periods'][0]
                m_col4.metric(f"Total Vendido ({p1['label']})", f"{p1['total_vendida']:,.0f}")
                m_col5.metric(f"Efetividade Média ({p1['label']})", f"{p1['avg_efetividade']:.1f}%")

            st.markdown("#### Detalhamento por Período:")
            tabs = st.tabs([p['label'] for p in metrics['periods']])
            for idx, tab in enumerate(tabs):
                p_data = metrics['periods'][idx]
                with tab:
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Total Orçado", f"{p_data['total_orcada']:,.0f}")
                    c2.metric("Total Vendido", f"{p_data['total_vendida']:,.0f}")
                    c3.metric("Efetividade Global", f"{p_data['avg_efetividade']:.1f}%")
                    c4.metric("Consumo Total OP", f"{p_data['total_consumida']:,.0f}")

            with st.expander("🔍 Visualizar Prévia das Tabelas Consolidadas", expanded=False):
                tab_p1, tab_p2 = st.tabs(["📋 Aba 1: Detalhado por Código de Produto", "🔄 Aba 2: Consolidado por Código Alternativo"])
                with tab_p1:
                    st.caption("Primeiras 10 linhas da Aba 1 (Análise Consolidada):")
                    st.dataframe(df_result.head(10), use_container_width=True)
                with tab_p2:
                    if 'df_alt' in metrics and not metrics['df_alt'].empty:
                        st.caption("Primeiras 10 linhas da Aba 2 (Consolidado por Código Alternativo):")
                        st.dataframe(metrics['df_alt'].head(10), use_container_width=True)

            # 9. DOWNLOAD BUTTON
            st.markdown("---")
            st.markdown("### 📥 Download do Relatório Formatado")
            
            now_str = datetime.now().strftime("%Y%m%d_%H%M")
            file_name = f"Analise_Vendas_Fabricacao_Consolidada_{now_str}.xlsx"

            st.download_button(
                label="💾 Baixar Planilha Consolidada (.xlsx)",
                data=excel_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            progress_bar.empty()
            st.error(f"❌ Ocorreu um erro durante o processamento: {str(e)}")
            st.exception(e)
