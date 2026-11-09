import streamlit as st
import os
import importlib
import modules.home_view
import modules.ponto_pedido_view
import modules.excesso_estoque_view
import data_processor
import excesso_processor
import storage_manager

# Force live reload of all modules and views so changes are always instantly active
importlib.reload(data_processor)
importlib.reload(excesso_processor)
importlib.reload(storage_manager)
importlib.reload(modules.home_view)
importlib.reload(modules.ponto_pedido_view)
importlib.reload(modules.excesso_estoque_view)

# 1. PAGE CONFIGURATION & GLOBAL STYLING
st.set_page_config(
    page_title="Portal de Inteligência de Estoque & Vendas",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium UI
st.markdown("""
<style>
    /* Main Background & Fonts */
    .main {
        background-color: #F8FAFC;
    }
    
    /* Header Banner Styling */
    .header-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #1F4E78 100%);
        padding: 24px 32px;
        border-radius: 14px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
    }
    .header-title {
        font-size: 26px;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .header-subtitle {
        font-size: 14px;
        color: #94A3B8;
        margin-top: 6px;
        font-weight: 400;
    }
    
    /* Saved File Card */
    .saved-card {
        background-color: #F0FDF4;
        border: 1px solid #BBF7D0;
        border-radius: 8px;
        padding: 10px 14px;
        margin-bottom: 8px;
    }
    .saved-card-title {
        color: #166534;
        font-weight: 700;
        font-size: 13px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .saved-card-sub {
        color: #374151;
        font-size: 11px;
        margin-top: 2px;
    }

    /* Metric Cards */
    div[data-testid="stMetricValue"] {
        font-size: 24px !important;
        font-weight: 700 !important;
        color: #0F172A !important;
    }
    
    /* Buttons Styling */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        height: 44px;
        transition: all 0.2s ease-in-out;
    }
    
    .stDownloadButton>button {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        height: 50px !important;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3) !important;
    }
    .stDownloadButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(16, 185, 129, 0.4) !important;
    }

    /* Badges */
    .period-badge {
        background-color: #E0F2FE;
        color: #0369A1;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# 2. SESSION STATE FOR PAGE ROUTING
PAGE_MAP = {
    "home": "🏠 Início (Hub de Relatórios)",
    "ponto_pedido": "📦 Ponto de Pedido",
    "excesso_estoque": "📊 Excesso em Estoque"
}
REV_PAGE_MAP = {v: k for k, v in PAGE_MAP.items()}

# Sync with query params if available
if "current_page" not in st.session_state:
    query_page = st.query_params.get("page", None)
    if query_page in PAGE_MAP:
        st.session_state.current_page = query_page
    else:
        st.session_state.current_page = "home"

# 3. SIDEBAR NAVIGATION
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=180)
    
    st.markdown("### 🧭 Ambiente de Trabalho")
    
    current_label = PAGE_MAP.get(st.session_state.current_page, PAGE_MAP["home"])
    options_list = list(PAGE_MAP.values())
    selected_index = options_list.index(current_label) if current_label in options_list else 0

    selected_label = st.radio(
        "Selecione o Relatório:",
        options=options_list,
        index=selected_index,
        key="nav_radio",
        label_visibility="collapsed"
    )

    new_page = REV_PAGE_MAP[selected_label]
    if new_page != st.session_state.current_page:
        st.session_state.current_page = new_page
        st.query_params["page"] = new_page
        st.rerun()

    if st.session_state.current_page != "home":
        if st.button("⬅️ Trocar de Relatório / Início", use_container_width=True):
            st.session_state.current_page = "home"
            st.query_params["page"] = "home"
            st.rerun()

    st.markdown("---")

# 4. VIEW ROUTER
if st.session_state.current_page == "home":
    modules.home_view.render_home_view()
elif st.session_state.current_page == "ponto_pedido":
    modules.ponto_pedido_view.render_ponto_pedido_view()
elif st.session_state.current_page == "excesso_estoque":
    modules.excesso_estoque_view.render_excesso_estoque_view()
else:
    modules.home_view.render_home_view()
