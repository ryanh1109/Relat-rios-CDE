import streamlit as st

def render_home_view():
    """
    Renders the Hub / Selection Portal where the user chooses between
    'Ponto de Pedido' and 'Excesso em Estoque'.
    """
    # Header do Portal
    st.markdown("""
    <div style="background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #1F4E78 100%); padding: 32px 36px; border-radius: 16px; color: white; margin-bottom: 28px; box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);">
        <div style="font-size: 30px; font-weight: 800; letter-spacing: -0.5px; display: flex; align-items: center; gap: 14px;">
            🏢 Central de Inteligência de Estoque & Vendas
        </div>
        <div style="font-size: 15px; color: #94A3B8; margin-top: 8px; font-weight: 400; max-width: 900px; line-height: 1.5;">
            Plataforma corporativa para consolidação analítica, gestão de suprimentos e inteligência de estoque.
            Escolha o ambiente do relatório que deseja acessar ou gerar abaixo:
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 🎯 Selecione o Tipo de Relatório:")
    st.caption("Cada relatório opera em um ambiente independente com regras, planilhas e parâmetros próprios.")

    col1, col2 = st.columns(2, gap="large")

    # CARD 1: PONTO DE PEDIDO
    with col1:
        st.markdown("""
        <div style="background-color: #FFFFFF; border: 2px solid #E2E8F0; border-radius: 14px; padding: 24px; min-height: 380px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); display: flex; flex-direction: column; justify-content: space-between;">
            <div>
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                    <span style="font-size: 26px;">📦</span>
                    <span style="background-color: #DCFCE7; color: #166534; padding: 4px 10px; border-radius: 20px; font-size: 12px; font-weight: 700;">✅ ATIVO & HOMOLOGADO</span>
                </div>
                <h3 style="margin: 0 0 10px 0; color: #0F172A; font-size: 20px; font-weight: 700;">Relatório: Ponto de Pedido</h3>
                <p style="color: #475569; font-size: 13.5px; line-height: 1.5; margin-bottom: 16px;">
                    Consolidação completa de vendas, orçamentos e fabricação com comparativo por períodos e sugestão de parâmetros de reposição.
                </p>
                <div style="background-color: #F8FAFC; border-radius: 8px; padding: 12px; border: 1px solid #E2E8F0; margin-bottom: 18px;">
                    <p style="font-size: 12px; font-weight: 700; color: #1E293B; margin-bottom: 6px;">Recursos inclusos:</p>
                    <ul style="font-size: 12px; color: #475569; padding-left: 18px; margin: 0; line-height: 1.6;">
                        <li>Cruzamento Cadastro + Vendas (SQL) + Consumo OPs (Protheus)</li>
                        <li>Rateio proporcional por <b>Data de Cadastro</b> do produto</li>
                        <li>Aba 2 com agrupamento por <b>Código Alternativo</b> e fórmulas SOMASE</li>
                        <li>Cálculo de <b>ES, LE e PP Sugeridos</b> com gatilho de 20% (ANALISAR)</li>
                        <li>Geração de arquivo Excel formatado com fórmulas dinâmicas</li>
                    </ul>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚀 Acessar Módulo: Ponto de Pedido", key="btn_go_pp", type="primary", use_container_width=True):
            st.session_state.current_page = "ponto_pedido"
            st.rerun()

    # CARD 2: EXCESSO EM ESTOQUE
    with col2:
        st.markdown("""
        <div style="background-color: #FFFFFF; border: 2px solid #E2E8F0; border-radius: 14px; padding: 24px; min-height: 380px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); display: flex; flex-direction: column; justify-content: space-between;">
            <div>
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                    <span style="font-size: 26px;">📊</span>
                    <span style="background-color: #FEF3C7; color: #92400E; padding: 4px 10px; border-radius: 20px; font-size: 12px; font-weight: 700;">🆕 NOVO PROJETO (EM INÍCIO)</span>
                </div>
                <h3 style="margin: 0 0 10px 0; color: #0F172A; font-size: 20px; font-weight: 700;">Relatório: Excesso em Estoque</h3>
                <p style="color: #475569; font-size: 13.5px; line-height: 1.5; margin-bottom: 16px;">
                    Ambiente dedicado para diagnóstico de excedentes de estoque, cobertura de meses, giro de materiais e otimização de capital.
                </p>
                <div style="background-color: #F8FAFC; border-radius: 8px; padding: 12px; border: 1px solid #E2E8F0; margin-bottom: 18px;">
                    <p style="font-size: 12px; font-weight: 700; color: #1E293B; margin-bottom: 6px;">Etapa de Planejamento:</p>
                    <ul style="font-size: 12px; color: #475569; padding-left: 18px; margin: 0; line-height: 1.6;">
                        <li>Ambiente 100% isolado para evitar conflito de dados</li>
                        <li>Espaço preparado para definição das planilhas de entrada</li>
                        <li>Configuração de parâmetros de giro, cobertura e corte de excesso</li>
                        <li>Estruturação das regras de negócio e layout dos resultados</li>
                    </ul>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("⚙️ Iniciar Módulo: Excesso em Estoque", key="btn_go_ee", use_container_width=True):
            st.session_state.current_page = "excesso_estoque"
            st.rerun()

    st.markdown("---")

    # Rodapé informativo
    st.info("""
    💡 **Organização em Ambientes Separados:**
    - Você pode navegar livremente entre os dois ambientes a qualquer momento pelo menu lateral.
    - As planilhas salvas no ambiente **Ponto de Pedido** permanecem preservadas e independentes do novo projeto **Excesso em Estoque**.
    """)
