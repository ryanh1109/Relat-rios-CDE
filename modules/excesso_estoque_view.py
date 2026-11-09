import streamlit as st
import pandas as pd
from datetime import datetime
import time
import os
import io
import storage_manager
from excesso_processor import process_excesso_data

def render_excesso_estoque_view():
    """
    Renders the Excesso em Estoque environment with standardized sidebar file management,
    matching the layout and workflow of the 'Ponto de Pedido' module.
    """
    # Invalida cache de sessão antigo se as novas colunas não estiverem presentes
    if "ee_df_result" in st.session_state and st.session_state.ee_df_result is not None:
        if "EXCESSO EM ESTOQUE" not in st.session_state.ee_df_result.columns or "VALOR EM ESTOQUE" not in st.session_state.ee_df_result.columns or "VALOR EM EXCESSO" not in st.session_state.ee_df_result.columns:
            del st.session_state["ee_df_result"]
            if "ee_excel_bytes" in st.session_state:
                del st.session_state["ee_excel_bytes"]
            if "ee_metrics" in st.session_state:
                del st.session_state["ee_metrics"]

    # 1. HEADER SECTION
    st.markdown("""
    <div class="header-container" style="background: linear-gradient(135deg, #1E1B4B 0%, #312E81 50%, #4338CA 100%);">
        <div class="header-title">
            📊 Relatório: Excesso em Estoque
        </div>
        <div class="header-subtitle" style="color: #C7D2FE;">
            Diagnóstico de materiais excedentes, cálculo de Estoque Máximo (PP + LE + ES) e cruzamento com a Posição de Estoque do Protheus.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 2. INSTRUÇÕES DE EXTRAÇÃO E PREPARAÇÃO DAS PLANILHAS
    with st.expander("📖 Instruções de Extração e Preparação das Planilhas (Clique para expandir)", expanded=False):
        st.markdown("##### 📋 Passo a Passo para Extração das Planilhas")
        col_cad_inst, col_pos_inst = st.columns(2)

        with col_cad_inst:
            st.markdown("**🗂️ 1. Cadastro de Produtos (SQL)**")
            st.markdown("""
1. Abrir o **SQL**
2. Atualizar a consulta do **Cadastro de Produtos**
3. Copiar todas as colunas e linhas **sem formatação e fórmulas** para outra planilha
4. Salvar o arquivo no seu computador (com qualquer nome)

💡 *Dica:* Se você já enviou a **Planilha 1** no módulo **Ponto de Pedido**, ela já é herdada automaticamente aqui!
""")

        with col_pos_inst:
            st.markdown("**📦 2. Posição de Estoque (Protheus)**")
            st.markdown("""
1. Acessar o sistema **Protheus**
2. Abrir o relatório de **Posição de Estoque**
3. Preencher os parâmetros de extração:
   - **Aglutina por?** `Armazem`
   - **Da filial?** `0201`
   - **Até filial?** `0203`
   - **Do armazém?** `01`
   - **Até armazém?** `62`
4. Gerar e exportar o relatório para arquivo **Excel (.xlsx)**
5. Salvar o arquivo em qualquer pasta do computador (com qualquer nome)
""")

    # 3. SIDEBAR - FILE UPLOAD PADRONIZADO (IGUAL AO PONTO DE PEDIDO)
    up_exclusive_cad = None
    up_posicao = None

    with st.sidebar:
        st.markdown("### 📁 Gerenciador de Planilhas (Excesso)")
        st.caption("Planilhas salvas no servidor local para este módulo:")

        # --- SEÇÃO 1: CADASTRO DE PRODUTOS ---
        st.markdown("#### 1. Cadastro de Produtos")
        info_ee_cad = storage_manager.get_file_info("ee_cadastro")
        info_shared_cad = storage_manager.get_file_info("planilha_1")

        if info_ee_cad:
            st.markdown(f"""
            <div class="saved-card" style="background-color: #EEF2FF; border-color: #C7D2FE;">
                <div class="saved-card-title" style="color: #312E81;">🏷️ Arquivo Salvo (Exclusivo)</div>
                <div class="saved-card-sub">📄 <b>{info_ee_cad['original_name']}</b> ({info_ee_cad['size']})</div>
                <div class="saved-card-sub">🕒 Atualizado: {info_ee_cad['last_updated']}</div>
            </div>
            """, unsafe_allow_html=True)

            c_up_cad, c_del_cad = st.columns([4, 1])
            with c_up_cad:
                up_exclusive_cad = st.file_uploader(
                    "🔄 Substituir 1. Cadastro de Produtos",
                    type=["xlsx", "xls"],
                    key="side_up_ee_cad_sub",
                    help="Substituir planilha exclusiva de cadastro",
                    label_visibility="collapsed"
                )
            with c_del_cad:
                if st.button("🗑️", key="side_del_ee_cad", help="Excluir cadastro exclusivo e voltar ao compartilhado"):
                    storage_manager.delete_file("ee_cadastro")
                    if "ee_excel_bytes" in st.session_state:
                        del st.session_state["ee_excel_bytes"]
                    st.rerun()

            if info_shared_cad:
                if st.button("↩️ Usar do Ponto de Pedido", key="btn_revert_shared", use_container_width=True, help="Volta a usar o cadastro compartilhado"):
                    storage_manager.delete_file("ee_cadastro")
                    if "ee_excel_bytes" in st.session_state:
                        del st.session_state["ee_excel_bytes"]
                    st.rerun()

        elif info_shared_cad:
            st.markdown(f"""
            <div class="saved-card">
                <div class="saved-card-title">🔗 Compartilhado (Ponto de Pedido)</div>
                <div class="saved-card-sub">📄 <b>{info_shared_cad['original_name']}</b> ({info_shared_cad['size']})</div>
                <div class="saved-card-sub">🕒 Atualizado: {info_shared_cad['last_updated']}</div>
            </div>
            """, unsafe_allow_html=True)

            up_exclusive_cad = st.file_uploader(
                "🔄 Anexar Cadastro Exclusivo",
                type=["xlsx", "xls"],
                key="side_up_ee_cad_new",
                help="Caso envie um arquivo aqui, ele substituirá o cadastro do Ponto de Pedido exclusivamente para este relatório."
            )
        else:
            up_exclusive_cad = st.file_uploader(
                "Upload 1. Cadastro de Produtos",
                type=["xlsx", "xls"],
                key="side_up_ee_cad_first",
                help="SQL: Copiar colunas/linhas sem formatação e fórmulas."
            )

        st.markdown("---")

        # --- SEÇÃO 2: POSIÇÃO DE ESTOQUE ---
        st.markdown("#### 2. Posição Estoque")
        info_pos = storage_manager.get_file_info("ee_posicao_estoque")

        if info_pos:
            st.markdown(f"""
            <div class="saved-card">
                <div class="saved-card-title">✅ Arquivo Salvo</div>
                <div class="saved-card-sub">📄 <b>{info_pos['original_name']}</b> ({info_pos['size']})</div>
                <div class="saved-card-sub">🕒 Atualizado: {info_pos['last_updated']}</div>
            </div>
            """, unsafe_allow_html=True)

            c_up_pos, c_del_pos = st.columns([4, 1])
            with c_up_pos:
                up_posicao = st.file_uploader(
                    "🔄 Substituir 2. Posição Estoque",
                    type=["xlsx", "xls"],
                    key="side_up_ee_pos_sub",
                    help="Enviar nova versão da Posição de Estoque",
                    label_visibility="collapsed"
                )
            with c_del_pos:
                if st.button("🗑️", key="side_del_ee_pos", help="Excluir planilha Posição de Estoque salva"):
                    storage_manager.delete_file("ee_posicao_estoque")
                    if "ee_excel_bytes" in st.session_state:
                        del st.session_state["ee_excel_bytes"]
                    st.rerun()
        else:
            up_posicao = st.file_uploader(
                "Upload 2. Posição Estoque",
                type=["xlsx", "xls"],
                key="side_up_ee_pos_first",
                help="Protheus: Aglutina por? Armazem | Filiais: 0201 a 0203 | Armazéns: 01 a 62"
            )

        st.markdown("---")
        if st.button("🧹 Limpar Planilhas Salvas deste Módulo", use_container_width=True):
            storage_manager.delete_file("ee_cadastro")
            storage_manager.delete_file("ee_posicao_estoque")
            if "ee_excel_bytes" in st.session_state:
                del st.session_state["ee_excel_bytes"]
            st.rerun()

    # 4. RESOLUÇÃO DOS ARQUIVOS NO SERVIDOR
    effective_cad, source_type_cad, source_info_cad = storage_manager.get_effective_cadastro_excesso(up_exclusive_cad)
    effective_pos = storage_manager.get_effective_file(up_posicao, "ee_posicao_estoque")
    info_pos_effective = storage_manager.get_file_info("ee_posicao_estoque")

    # 5. CONTEÚDO PRINCIPAL - CONFIGURAÇÃO E RESUMO DOS PARÂMETROS
    col_main, col_summary = st.columns([7, 5])

    with col_main:
        st.subheader("⚙️ Configurações do Relatório")
        st.caption("Confirme as fontes de dados e opções de visualização:")

        # Status do Cadastro
        if source_type_cad == "exclusive":
            st.success(f"🏷️ **Cadastro de Produtos (Exclusivo):** `{source_info_cad['original_name']}` ({source_info_cad['size']})")
        elif source_type_cad == "shared":
            st.info(f"🔗 **Cadastro de Produtos (Compartilhado do Ponto de Pedido):** `{source_info_cad['original_name']}` ({source_info_cad['size']})")
        else:
            st.error("❌ **Cadastro de Produtos:** Pendente de upload na barra lateral.")

        # Status da Posição Estoque
        if effective_pos and info_pos_effective:
            st.success(f"📦 **Posição de Estoque (Protheus):** `{info_pos_effective['original_name']}` ({info_pos_effective['size']})")
        else:
            st.warning("⚠️ **Posição de Estoque:** Não enviada. Faça o upload na barra lateral para cruzar o saldo real com o Estoque Máximo.")

        st.markdown("---")
        include_details = st.checkbox(
            "📋 Incluir colunas base de cálculo (PP, LE, ES) no relatório final",
            value=False,
            help="Se marcado, adiciona as colunas de PP, LE e ES com fórmulas no Excel. Se desmarcado, traz exatamente as colunas principais."
        )

    with col_summary:
        st.subheader("📋 Resumo dos Parâmetros")
        st.markdown("""
        <div style="background-color: #FFFFFF; padding: 18px; border-radius: 10px; border: 1px solid #E2E8F0;">
            <p style="margin-bottom: 8px;"><strong>Regras de Negócio Aplicadas:</strong></p>
            <ul style="font-size: 13px; color: #475569; padding-left: 20px; margin-bottom: 0; line-height: 1.6;">
                <li><b>Estoque Máximo:</b> <code>PP + LE + ES</code> (Ponto de Pedido + Lote Econômico + Estoque de Segurança).</li>
                <li><b>Saldo Atual:</b> Soma consolidada de <code>SALDO EM ESTOQUE</code> por produto (<b>desconsiderando armazém 51</b>).</li>
                <li><b>Excesso em Estoque:</b> Fórmula nativa no Excel <code>=SALDO ATUAL - ESTOQUE MÁXIMO</code>.</li>
                <li><b>Valor em Estoque:</b> Soma consolidada de <code>VALOR EM ESTOQUE</code> por produto (<b>desconsiderando armazém 51</b>).</li>
                <li><b>Valor em Excesso:</b> Fórmula nativa no Excel <code>=(VALOR EM ESTOQUE / SALDO ATUAL) * EXCESSO EM ESTOQUE</code>.</li>
                <li><b>Filtros Protheus (Posição):</b>
                    <br>• Aglutina por? <b>Armazem</b>
                    <br>• Filiais: de <b>0201</b> até <b>0203</b>
                    <br>• Armazéns: de <b>01</b> até <b>62</b>
                </li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # 6. SEÇÃO DE PROCESSAMENTO (PADRONIZADA COM PONTO DE PEDIDO)
    has_cad = effective_cad is not None
    ready_to_process = has_cad

    col_proc, col_msg = st.columns([4, 6])

    with col_proc:
        btn_process = st.button(
            "🚀 Processar e Gerar Planilha de Excesso",
            type="primary",
            disabled=not ready_to_process,
            use_container_width=True
        )

    with col_msg:
        if not has_cad:
            st.warning("⚠️ Faltando planilha de **1. Cadastro de Produtos** na barra lateral.")
        elif not effective_pos:
            st.info("💡 Planilha de Cadastro pronta! Você pode processar agora ou anexar a **2. Posição Estoque** na barra lateral para cruzar os dados.")
        else:
            st.success("✅ Ambas as planilhas estão prontas no servidor local! Clique no botão para iniciar.")

    # 7. EXECUÇÃO DO PROCESSAMENTO
    if btn_process and ready_to_process:
        progress_bar = st.progress(0, text="Iniciando processamento das planilhas...")
        
        try:
            time.sleep(0.15)
            progress_bar.progress(30, text="Lendo Cadastro de Produtos e calculando Estoque Máximo (PP + LE + ES)...")
            
            if effective_pos:
                time.sleep(0.15)
                progress_bar.progress(60, text="Cruzando com dados da Posição de Estoque (Filial 0201 / Armazém)...")

            time.sleep(0.15)
            progress_bar.progress(85, text="Formatando e estilizando planilha Excel...")

            df_result, excel_bytes, metrics = process_excesso_data(
                file_cad=effective_cad,
                file_pos=effective_pos,
                include_details=include_details
            )

            progress_bar.progress(100, text="Concluído com sucesso!")
            time.sleep(0.2)
            progress_bar.empty()

            # Salva no session state para retenção
            st.session_state.ee_df_result = df_result
            st.session_state.ee_excel_bytes = excel_bytes
            st.session_state.ee_metrics = metrics
            st.toast("🎉 Planilha de Excesso em Estoque gerada com sucesso!", icon="✅")

        except Exception as e:
            progress_bar.empty()
            st.error(f"❌ Ocorreu um erro durante o processamento: {str(e)}")
            st.exception(e)

    # 8. EXIBIÇÃO DOS RESULTADOS (PERSISTIDO EM SESSION STATE)
    if "ee_df_result" in st.session_state and st.session_state.ee_df_result is not None:
        df_result = st.session_state.ee_df_result
        excel_bytes = st.session_state.ee_excel_bytes
        metrics = st.session_state.ee_metrics

        st.markdown("---")
        st.subheader("📊 Resumo Executivo dos Resultados")

        # DASHBOARD EXECUTIVO
        if metrics.get('has_posicao'):
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Produtos Analisados", f"{metrics['total_products']:,}")
            m2.metric("🚨 Itens com Excesso", f"{metrics.get('total_itens_com_excesso', 0):,}", help="Itens onde o Saldo Atual supera o Estoque Máximo")
            m3.metric("Qtd Total em Excesso", f"{metrics.get('total_qtd_excesso', 0):,.0f}")
            m4.metric("Saldo Total em Estoque", f"{metrics.get('total_saldo', 0):,.0f}", help="Soma do Saldo Atual (excluindo armazém 51)")
            m5.metric("Valor Total Estoque", f"R$ {metrics.get('total_valor', 0):,.2f}", help="Soma do Valor em Estoque (excluindo armazém 51)")

            val_exc = metrics.get('total_valor_excesso', 0.0)
            if val_exc > 0:
                st.markdown(f"""
                <div style="background-color: #FEF2F2; border: 1px solid #FECACA; border-radius: 8px; padding: 12px 18px; margin-top: 10px; margin-bottom: 14px;">
                    <span style="color: #991B1B; font-weight: 700; font-size: 14px;">⚠️ Valor Financeiro em Excesso Estimado:</span>
                    <span style="color: #B91C1C; font-weight: 800; font-size: 16px; margin-left: 8px;">R$ {val_exc:,.2f}</span>
                    <span style="color: #7F1D1D; font-size: 12px; margin-left: 10px;">(Calculado com base no custo unitário proporcional dos itens com saldo acima do teto)</span>
                </div>
                """, unsafe_allow_html=True)
        else:
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total de Produtos", f"{metrics['total_products']:,}")
            m2.metric("Com Estoque Máximo > 0", f"{metrics['total_with_max']:,}")
            m3.metric("Sem Parâmetros (= 0)", f"{metrics['total_zero_max']:,}")
            m4.metric("Estoque Máximo Médio", f"{metrics['avg_max']:,.1f}")
            m5.metric("Grupos Encontrados", f"{metrics['total_groups']:,}")

        # PRÉVIA DA TABELA
        with st.expander("🔍 Visualizar Prévia da Tabela Consolidada", expanded=True):
            st.caption(f"Exibindo as 10 primeiras linhas da tabela ({len(df_result):,} produtos no total):")
            st.dataframe(df_result.head(10), use_container_width=True)

        # BOTÃO DE DOWNLOAD
        st.markdown("---")
        st.markdown("### 📥 Download do Relatório Formatado")
        
        now_str = datetime.now().strftime("%Y%m%d_%H%M")
        file_name = f"Relatorio_Excesso_Estoque_{now_str}.xlsx"

        st.download_button(
            label="💾 Baixar Planilha de Excesso em Estoque (.xlsx)",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
