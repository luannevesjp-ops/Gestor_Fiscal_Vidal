# ============================================================================
# GESTOR FISCAL - LUATECH
# Sistema de Gestão Fiscal com Streamlit
# ============================================================================

import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from io import BytesIO
import requests
import time
import os
import base64

# ============================================================================
# CONFIGURAÇÕES INICIAIS
# ============================================================================

st.set_page_config(page_title="LuaTech - Gestão Fiscal", layout="wide")

if 'main_container' not in st.session_state:
    st.session_state.main_container = st.empty()

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1bp7qtkKvsMHMvHjGznT6OwyX_YSQWMa3jVvylOJWSxM/export?format=xlsx"
SHEET_EMPRESAS = "GERAL"

# ============================================================================
# CSS E ESTILOS
# ============================================================================

st.markdown("""
<meta name="google" content="notranslate">
<meta name="googlebot" content="notranslate">
<script>
window.addEventListener('error', function(e) {
    if (e.message && e.message.includes('removeChild')) {
        e.preventDefault();
        console.warn('Erro removeChild suprimido:', e.message);
    }
});
</script>
""", unsafe_allow_html=True)

st.markdown("""
<style>
.header-class .ag-header-cell-label {
    color: white !important;
    font-weight: bold !important;
    background-color: #1d3f77 !important;
}
.sidebar-lt {
    background-color: #1d3f77;
    padding: 0;
    margin: 0;
}
.sidebar-lt img {
    width: 100%;
    display: block;
    border-radius: 0;
}
/* Botões do menu principal */
.menu-btn button {
    width: 100%;
    height: 110px;
    font-size: 22px !important;
    font-weight: bold !important;
    border-radius: 12px !important;
    border: 2px solid #1d3f77 !important;
    background-color: #1d3f77 !important;
    color: white !important;
    cursor: pointer;
    transition: background-color 0.2s;
}
.menu-btn button:hover {
    background-color: #163066 !important;
}
</style>
""", unsafe_allow_html=True)

grid_container = st.empty()

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

@st.cache_data(ttl=600)
def le_planilha_google(url: str, aba: str):
    try:
        resp = requests.get(url)
        resp.raise_for_status()
        df = pd.read_excel(BytesIO(resp.content), sheet_name=aba, engine='openpyxl')
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        return None


def exibe_aggrid(df, height=400, grid_key="grid", selection_mode='none'):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True, sortable=True, editable=False, resizable=True)
    
    if selection_mode != 'none':
        gb.configure_selection(selection_mode=selection_mode, use_checkbox=True)
    
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            gb.configure_column(col, filter="agNumberColumnFilter")
        else:
            gb.configure_column(col, filter="agTextColumnFilter")
    
    gb.configure_grid_options(
        domLayout="normal", floatingFilter=True, headerHeight=40, rowHeight=30,
        enableBrowserTooltips=True, enableCellTextSelection=True, suppressMenuHide=True,
        localeText={
            'filterOoo': 'Filtrar...', 'contains': 'Contém', 'notContains': 'Não contém',
            'equals': 'Igual', 'notEqual': 'Diferente', 'blank': 'Em branco',
            'notBlank': 'Não em branco', 'noRowsToShow': 'Nenhum registro para mostrar',
        }
    )
    
    grid_options = gb.build()
    update_on = ['selectionChanged'] if selection_mode != 'none' else []
    
    return AgGrid(df, gridOptions=grid_options, height=height, key=grid_key,
                  fit_columns_on_grid_load=True, enable_enterprise_modules=False,
                  update_on=update_on, allow_unsafe_jscode=True, reload_data=False)


def exibe_aggrid_com_oculta(df, height=400, grid_key="grid", selection_mode='none', colunas_ocultas=None):
    if colunas_ocultas is None:
        colunas_ocultas = []
    
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True, sortable=True, editable=False, resizable=True)
    
    if selection_mode != 'none':
        gb.configure_selection(selection_mode=selection_mode, use_checkbox=True)
    
    for col in df.columns:
        if col in colunas_ocultas:
            gb.configure_column(col, hide=True)
        elif pd.api.types.is_numeric_dtype(df[col]):
            gb.configure_column(col, filter="agNumberColumnFilter")
        else:
            gb.configure_column(col, filter="agTextColumnFilter")
    
    gb.configure_grid_options(
        domLayout="normal", floatingFilter=True, headerHeight=40, rowHeight=30,
        enableBrowserTooltips=True, enableCellTextSelection=True, suppressMenuHide=True,
        localeText={'filterOoo': 'Filtrar...', 'noRowsToShow': 'Nenhum registro'}
    )
    
    grid_options = gb.build()
    update_on = ['selectionChanged'] if selection_mode != 'none' else []
    
    return AgGrid(df, gridOptions=grid_options, height=height, key=grid_key,
                  fit_columns_on_grid_load=True, enable_enterprise_modules=False,
                  update_on=update_on, allow_unsafe_jscode=True, reload_data=False)

# ============================================================================
# AUTENTICAÇÃO
# ============================================================================

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

# Controle da área principal do menu
if "menu_area" not in st.session_state:
    st.session_state["menu_area"] = None

def tela_login():
    st.markdown("<h1 style='text-align:center; color:#0f4fa3;'>Gestão Fiscal</h1>", unsafe_allow_html=True)
    senha = st.text_input("Senha", type="password", max_chars=20)
    if st.button("Entrar"):
        if senha == "VIDAL":
            st.session_state["autenticado"] = True
        else:
            st.error("Senha incorreta.")

if not st.session_state["autenticado"]:
    tela_login()
    st.stop()

# ============================================================================
# MENU PRINCIPAL (FISCAL / PARALEGAL / CONTÁBIL)
# ============================================================================

def tela_menu_principal():
    """Tela de seleção da área após o login"""
    st.sidebar.markdown("""
    <div class="sidebar-lt" >
        <img src="data:image/png;base64,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" style="width:100%; display:block;">
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align:center; color:#1d3f77;'>Selecione a Área</h1>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown('<div class="menu-btn">', unsafe_allow_html=True)
        if st.button("FISCAL", use_container_width=True, key="btn_fiscal"):
            st.session_state["menu_area"] = "FISCAL"
            st.session_state["pagina_atual"] = "EMPRESAS"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="menu-btn">', unsafe_allow_html=True)
        if st.button("DEPARTAMENTO PARALEGAL", use_container_width=True, key="btn_paralegal"):
            st.session_state["menu_area"] = "PARALEGAL"
            st.session_state["pagina_atual"] = "EMPRESAS"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col3:
        st.markdown('<div class="menu-btn">', unsafe_allow_html=True)
        if st.button("CONTÁBIL", use_container_width=True, key="btn_contabil"):
            st.session_state["menu_area"] = "CONTÁBIL"
            st.session_state["pagina_atual"] = "EMPRESAS"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

if st.session_state["menu_area"] is None:
    tela_menu_principal()
    st.stop()

# ============================================================================
# SIDEBAR
# ============================================================================

st.sidebar.markdown("""
<div class="sidebar-lt" >
    <img src="data:image/png;base64,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" style="width:100%; display:block;">
</div>
""", unsafe_allow_html=True)

# Exibe a área atual e botão de voltar
label_area = {"FISCAL": "FISCAL", "PARALEGAL": "DEPARTAMENTO PARALEGAL", "CONTÁBIL": "CONTÁBIL"}
st.sidebar.markdown(f"<p style='text-align:center; color:#1d3f77; font-weight:bold; margin-top:10px;'>{label_area.get(st.session_state['menu_area'], st.session_state['menu_area'])}</p>", unsafe_allow_html=True)

if st.sidebar.button("← DEPARTAMENTOS", use_container_width=True):
    st.session_state["menu_area"] = None
    st.session_state["pagina_atual"] = None
    st.rerun()

st.sidebar.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

# Define as páginas disponíveis por área
if st.session_state["menu_area"] == "FISCAL":
    paginas_disponiveis = ["EMPRESAS", "SIMPLES NACIONAL", "REINF", "DCTF WEB",
                           "DMS", "SERVIÇOS TOMADOS", "SEFAZ"]

elif st.session_state["menu_area"] == "PARALEGAL":
    paginas_disponiveis = ["Dashboard", "EMPRESAS", "CND Municipal"]

elif st.session_state["menu_area"] == "CONTÁBIL":
    paginas_disponiveis = ["EMPRESAS"]

else:
    paginas_disponiveis = ["EMPRESAS"]

pagina = st.sidebar.radio("", paginas_disponiveis,
                          index=0, label_visibility="collapsed")

if "pagina_atual" not in st.session_state:
    st.session_state["pagina_atual"] = pagina

if st.session_state["pagina_atual"] != pagina:
    st.session_state["pagina_atual"] = pagina
    st.rerun()

# ============================================================================
# PÁGINAS
# ============================================================================

def pagina_empresas():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    if "Situação" in df.columns:
        df_empresas = df[df["Situação"].astype(str).str.upper() == "ATIVA"]
    else:
        st.error("Coluna 'Situação' não encontrada.")
        return
    
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado", "Matriz / Filial", "Situação"]
    df_empresas = df_empresas[[c for c in colunas if c in df_empresas.columns]]
    total_empresas = df_empresas.shape[0]
    
    st.subheader("Empresas - Apenas ATIVAS")
    st.markdown(f"<p style='text-align:right; font-size:20px;'><b>Total:</b> {total_empresas} | <b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
    
    with st.container():
        exibe_aggrid(df_empresas, height=400, grid_key="grid_empresas")
    
    output = BytesIO()
    df_empresas.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="empresas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_simples():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    df_ativas = df[(df["Situação"].astype(str).str.upper() == "ATIVA") & 
                   (df["Regime"].astype(str).str.upper() == "SIMPLES NACIONAL")] if "Situação" in df.columns else pd.DataFrame()
    
    if df_ativas.empty:
        st.warning("Nenhuma empresa SIMPLES NACIONAL ATIVA encontrada.")
        return
    
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado", "SIMPLES GERADO", "Situação"]
    df_simples = df_ativas[[c for c in colunas if c in df_ativas.columns]]
    df_simples["SIMPLES GERADO"] = df_simples["SIMPLES GERADO"].apply(
        lambda x: "Filial" if str(x).upper() == "FILIAL" else ("Concluída" if pd.notna(x) else "Não"))
    
    concluidas = df_simples[df_simples["SIMPLES GERADO"].isin(["Concluída", "Filial"])].shape[0]
    nao_concluidas = df_simples[df_simples["SIMPLES GERADO"] == "Não"].shape[0]
    
    st.markdown(f"<h2>SIMPLES NACIONAL</h2><p style='text-align:right; font-size:20px;'>"
                f"<b>Concluídas:</b> {concluidas} | <b>Não concluídas:</b> {nao_concluidas} | "
                f"<b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
    
    time.sleep(1)
    exibe_aggrid(df_simples, height=400, grid_key="grid_simples")
    
    output = BytesIO()
    df_simples.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="simples_nacional.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_reinf():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None or df.empty:
        st.warning("Nenhum dado encontrado.")
        return
    
    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    df_reinf = df[df["Situação"].astype(str).str.upper() == "ATIVA"] if "Situação" in df.columns else pd.DataFrame()
    if df_reinf.empty:
        st.warning("Nenhuma empresa ATIVA para REINF.")
        return
    
    if "TRANSMISSÃO" in df_reinf.columns:
        df_reinf["TRANSMISSÃO"] = df_reinf["TRANSMISSÃO"].astype(str).str.upper().replace({
            "OK": "Transmitida", "FILIAL": "FILIAL", "NAN": "Não", "": "Não"})
    else:
        df_reinf["TRANSMISSÃO"] = "Não"
    
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "TRANSMISSÃO", "Situação"]
    df_reinf = df_reinf[[c for c in colunas if c in df_reinf.columns]]
    
    total_filial = df_reinf[df_reinf["TRANSMISSÃO"] == "FILIAL"].shape[0]
    total_transmitida = df_reinf[df_reinf["TRANSMISSÃO"] == "Transmitida"].shape[0]
    total_nao = df_reinf[df_reinf["TRANSMISSÃO"] == "Não"].shape[0]
    
    st.markdown(f"<h2>REINF</h2><p style='text-align:right; font-size:20px;'>"
                f"<b>Filial:</b> {total_filial} | <b>Transmitida:</b> {total_transmitida} | "
                f"<b>Não transmitida:</b> {total_nao} | <b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
    
    time.sleep(1)
    exibe_aggrid(df_reinf, height=400, grid_key="grid_reinf")
    
    output = BytesIO()
    df_reinf.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="reinf.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_dctf_web():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None or df.empty:
        st.warning("Nenhum dado encontrado.")
        return
    
    df = df.fillna("")
    df = df[df["Situação"].astype(str).str.upper() == "ATIVA"]
    if df.empty:
        st.warning("Nenhuma empresa ATIVA encontrada.")
        return
    
    competencia = ""
    if "PERÍODO DE COMPETÊNCIA" in df.columns:
        try:
            competencia_dt = pd.to_datetime(df["PERÍODO DE COMPETÊNCIA"].iloc[0], errors="coerce")
            if not pd.isna(competencia_dt):
                competencia = competencia_dt.strftime("%m/%Y")
        except:
            pass
    
    if "PERÍODO" in df.columns:
        df["PERÍODO"] = pd.to_datetime(df["PERÍODO"], errors="coerce").dt.strftime("%m-%Y").fillna("")
    
    df_dctf = df[["Código", "Razão Social", "CNPJ", "Regime", "PERÍODO", "ORIGEM", "TIPO", 
                  "SITUAÇÃO DCTF", "MATRIZ / FILIAL", "Situação"]].copy()
    
    concluidas = df_dctf[df_dctf["SITUAÇÃO DCTF"].astype(str).str.upper() == "ATIVA"].shape[0]
    sem_procuracao = df_dctf[df_dctf["SITUAÇÃO DCTF"].astype(str).str.upper() == "SEM PROCURAÇÃO"].shape[0]
    filiais = df_dctf[df_dctf["MATRIZ / FILIAL"].astype(str).str.upper() == "FILIAL"].shape[0]
    nao_concluidas_total = df_dctf[~df_dctf["SITUAÇÃO DCTF"].astype(str).str.upper().isin(["ATIVA", "SEM PROCURAÇÃO"])].shape[0]
    nao_concluidas = max(0, nao_concluidas_total - filiais)
    
    st.markdown(f"<h2>DCTF WEB</h2><p style='text-align:right; font-size:20px;'>"
                f"<b>Concluídas:</b> {concluidas} | <b>Sem Procuração:</b> {sem_procuracao} | "
                f"<b>Filiais:</b> {filiais} | <b>Não concluídas:</b> {nao_concluidas} | "
                f"<b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
    
    gb = GridOptionsBuilder.from_dataframe(df_dctf)
    gb.configure_default_column(resizable=True, filter=True, sortable=True)
    gb.configure_grid_options(domLayout="normal")
    
    AgGrid(df_dctf, gridOptions=gb.build(), update_mode=GridUpdateMode.NO_UPDATE,
           fit_columns_on_grid_load=True, height=600)
    
    output = BytesIO()
    df_dctf.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="dctf_web.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_dms():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    df_dms = df[df["Situação"].astype(str).str.upper() == "ATIVA"] if "Situação" in df.columns else pd.DataFrame()
    if df_dms.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para DMS.")
        return
    
    for col in ["FATURAMENTO SERVIÇOS", "BASE DE CÁLCULO ISS"]:
        if col in df_dms.columns:
            df_dms[col] = df_dms[col].fillna(0).astype(float).map(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    if "DMS" in df_dms.columns:
        df_dms["DMS"] = df_dms["DMS"].fillna("")
    
    if "GUIA ISS DMS" not in df_dms.columns:
        df_dms["GUIA ISS DMS"] = "Não"
    else:
        df_dms["GUIA ISS DMS"] = df_dms["GUIA ISS DMS"].astype(str).str.upper().replace({
            "OK": "Guia salva", "NAN": "Não", "": "Não"})
    
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado",
               "FATURAMENTO SERVIÇOS", "BASE DE CÁLCULO ISS", "XML DMS", "DMS", "GUIA ISS DMS", "Situação"]
    df_dms = df_dms[[c for c in colunas if c in df_dms.columns]]
    
    concluidas = df_dms[df_dms["DMS"].astype(str).str.upper() == "DMS SALVA"].shape[0]
    sem_acesso = df_dms[df_dms["DMS"].astype(str).str.upper() == "SEM ACESSO"].shape[0]
    nao_concluidas = df_dms[~df_dms["DMS"].astype(str).str.upper().isin(["DMS SALVA", "SEM ACESSO"])].shape[0]
    
    st.markdown(f"<h2>DMS</h2><p style='text-align:right; font-size:20px;'>"
                f"<b>Concluídas:</b> {concluidas} | <b>Sem acesso:</b> {sem_acesso} | "
                f"<b>Não concluídas:</b> {nao_concluidas} | <b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
    
    time.sleep(1)
    exibe_aggrid(df_dms, height=400, grid_key="grid_dms")
    
    output = BytesIO()
    df_dms.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="dms.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_rest():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    df_rest = df[df["Situação"].astype(str).str.upper() == "ATIVA"] if "Situação" in df.columns else pd.DataFrame()
    if df_rest.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para SERVIÇOS TOMADOS.")
        return
    
    for col in ["REST", "GUIA ISS REST"]:
        if col in df_rest.columns:
            df_rest[col] = df_rest[col].fillna("").astype(str)
    
    if "REST" in df_rest.columns:
        df_rest["REST"] = df_rest["REST"].replace({
            "REST SALVA": "Concluído", "SEM ACESSO": "Sem acesso", "": "Não concluído"})
    
    colunas = ["Código", "Razão Social", "CNPJ", "REST", "XML REST", "GUIA ISS REST", "Situação"]
    df_rest = df_rest[[c for c in colunas if c in df_rest.columns]]
    
    concluidas = df_rest[df_rest["REST"] == "Concluído"].shape[0]
    sem_acesso = df_rest[df_rest["REST"] == "Sem acesso"].shape[0]
    nao_concluidas = df_rest[df_rest["REST"] == "Não concluído"].shape[0]
    
    st.markdown(f"<h2>SERVIÇOS TOMADOS</h2><p style='text-align:right; font-size:20px;'>"
                f"<b>Concluídas:</b> {concluidas} | <b>Sem acesso:</b> {sem_acesso} | "
                f"<b>Não concluídas:</b> {nao_concluidas} | <b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
    
    time.sleep(1)
    exibe_aggrid(df_rest, height=400, grid_key="grid_rest")
    
    output = BytesIO()
    df_rest.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="servicos_tomados.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_sefaz():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    df_sefaz = df[df["Situação"].astype(str).str.upper() == "ATIVA"] if "Situação" in df.columns else pd.DataFrame()
    if df_sefaz.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para SEFAZ.")
        return
    
    colunas = ["Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual",
               "XML ENTRADA", "XML SAÍDA", "IMPORTAÇÃO", "TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO", "Situação"]
    df_sefaz = df_sefaz[[c for c in colunas if c in df_sefaz.columns]]
    
    for col in ["TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO"]:
        if col in df_sefaz.columns:
            df_sefaz[col] = pd.to_numeric(df_sefaz[col], errors="coerce").fillna(0)
    
    if "IMPORTAÇÃO" in df_sefaz.columns:
        em_andamento = df_sefaz[df_sefaz["IMPORTAÇÃO"].astype(str).str.upper() == "EM ANDAMENTO"].shape[0]
        outro_estado = df_sefaz[df_sefaz["IMPORTAÇÃO"].astype(str).str.upper() == "OUTRO ESTADO"].shape[0]
        sem_movimento = df_sefaz[df_sefaz["IMPORTAÇÃO"].astype(str).str.upper() == "SEM MOVIMENTO"].shape[0]
        concluido = df_sefaz[df_sefaz["IMPORTAÇÃO"].astype(str).str.upper() == "CONCLUÍDO"].shape[0]
    else:
        em_andamento = outro_estado = sem_movimento = concluido = 0
    
    st.markdown(f"<h2>SEFAZ</h2><p style='text-align:right; font-size:20px;'>"
                f"<b>Em andamento:</b> {em_andamento} | <b>Outro Estado:</b> {outro_estado} | "
                f"<b>Sem movimento:</b> {sem_movimento} | <b>Concluído:</b> {concluido} | "
                f"<b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
    
    time.sleep(1)
    exibe_aggrid(df_sefaz, height=400, grid_key="grid_sefaz")
    
    output = BytesIO()
    df_sefaz.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="sefaz.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_cnd_municipal():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    df_cnd = df[df["Situação"].astype(str).str.upper() == "ATIVA"] if "Situação" in df.columns else pd.DataFrame()
    if df_cnd.empty:
        st.warning("Nenhuma empresa ATIVA encontrada.")
        return
    
    colunas_solicitadas = ["Código", "Razão Social", "CNPJ", "Município", "Estado", 
                           "SITUAÇÃO CND MUNICIPAL", "VALIDADE", "LINK CND MUNICIPAL", "Situação"]
    colunas_existentes = [c for c in colunas_solicitadas if c in df_cnd.columns]
    df_cnd = df_cnd[colunas_existentes].copy()
    
    if "VALIDADE" in df_cnd.columns:
        df_cnd["VALIDADE"] = pd.to_datetime(df_cnd["VALIDADE"], errors='coerce').dt.strftime("%d/%m/%Y").fillna("")
    
    def check_pdf_link(link):
        return "Disponível" if pd.notna(link) and str(link).strip() != "" else "Indisponível"
    
    if "LINK CND MUNICIPAL" in df_cnd.columns:
        df_cnd["PDF"] = df_cnd["LINK CND MUNICIPAL"].apply(check_pdf_link)
    else:
        df_cnd["PDF"] = "Indisponível"
    
    if "SITUAÇÃO CND MUNICIPAL" in df_cnd.columns:
        situacao_upper = df_cnd["SITUAÇÃO CND MUNICIPAL"].astype(str).str.upper().str.strip()
        positivas = (situacao_upper == "POSITIVA").sum()
        negativas = (situacao_upper == "NEGATIVA").sum()
        positiva_efeito_negativa = (situacao_upper == "POSITIVA COM EFEITO NEGATIVA").sum()
    if "SITUAÇÃO CND MUNICIPAL" in df_cnd.columns:
        situacao_upper = df_cnd["SITUAÇÃO CND MUNICIPAL"].astype(str).str.upper().str.strip()
        positivas = (situacao_upper == "POSITIVA").sum()
        negativas = (situacao_upper == "NEGATIVA").sum()
        positiva_efeito_negativa = (situacao_upper == "POSITIVA COM EFEITO NEGATIVA").sum()
        nao_geradas = ((situacao_upper == "") | (situacao_upper == "NAN") | df_cnd["SITUAÇÃO CND MUNICIPAL"].isna()).sum()
    else:
        positivas = negativas = positiva_efeito_negativa = nao_geradas = 0
    
    total_geral = df_cnd.shape[0]
    
    if "visualizando_pdf" not in st.session_state:
        st.session_state.visualizando_pdf = False
        st.session_state.pdf_selecionado = None
    
    if st.session_state.visualizando_pdf and st.session_state.pdf_selecionado:
        col1, col2 = st.columns([6, 1])
        
        with col1:
            if st.button("← Voltar para a lista", type="primary"):
                st.session_state.visualizando_pdf = False
                st.session_state.pdf_selecionado = None
                st.rerun()
        
        with col2:
            row = st.session_state.pdf_selecionado
            link_pdf = row.get("LINK CND MUNICIPAL", "")
            
            if link_pdf and "drive.google.com" in str(link_pdf):
                if "/file/d/" in link_pdf:
                    file_id = link_pdf.split("/file/d/")[1].split("/")[0]
                    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
                    
                    st.markdown(
                        f'<a href="{download_url}" target="_blank">'
                        f'<button style="background-color:#1d3f77; color:white; padding:8px 16px; '
                        f'border:none; border-radius:4px; cursor:pointer; font-size:14px;">'
                        f'📥 Baixar PDF</button></a>',
                        unsafe_allow_html=True
                    )
        
        st.divider()
        row = st.session_state.pdf_selecionado
        cnpj = row.get("CNPJ", "")
        razao = row.get("Razão Social", "")
        link_pdf = row.get("LINK CND MUNICIPAL", "")
        status_pdf = row.get("PDF", "Indisponível")
        
        st.subheader(f"📄 {razao}")
        st.caption(f"CNPJ: {cnpj}")
        
        if status_pdf == "Disponível" and link_pdf:
            try:
                if "drive.google.com" in str(link_pdf):
                    if "/file/d/" in link_pdf:
                        file_id = link_pdf.split("/file/d/")[1].split("/")[0]
                        embed_url = f"https://drive.google.com/file/d/{file_id}/preview"
                        st.markdown(f'<iframe src="{embed_url}" width="100%" height="800" frameborder="0"></iframe>',
                                    unsafe_allow_html=True)
                    else:
                        st.error("❌ Formato de link do Google Drive não reconhecido")
                        st.info(f"Link: {link_pdf}")
                else:
                    st.markdown(f'<iframe src="{link_pdf}" width="100%" height="800" frameborder="0"></iframe>',
                                unsafe_allow_html=True)
            except Exception as e:
                st.error(f"❌ Erro ao carregar PDF: {e}")
                st.info(f"Link: {link_pdf}")
        else:
            st.error("❌ PDF não disponível")
            st.info("Link do PDF não foi encontrado na planilha (coluna LINK CND MUNICIPAL)")
    else:
        st.markdown(f"<h2>CND Municipal</h2><p style='text-align:right; font-size:20px;'>"
                    f"<b>Positivas:</b> {positivas} | <b>Negativas:</b> {negativas} | "
                    f"<b>Positiva c/ efeito negativa:</b> {positiva_efeito_negativa} | "
                    f"<b>Não geradas:</b> {nao_geradas} | <b>Total:</b> {total_geral} | "
                    f"<b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
        
        st.info("💡 Selecione uma linha na tabela para visualizar o PDF correspondente.")
        
        with st.container():
            grid_response = exibe_aggrid_com_oculta(df_cnd, height=400, grid_key="grid_cnd_municipal",
                                                     selection_mode='single',
                                                     colunas_ocultas=["Situação", "LINK CND MUNICIPAL"])
        
        selected_rows = grid_response.get('selected_rows', [])
        if selected_rows is not None and len(selected_rows) > 0:
            row = selected_rows.iloc[0].to_dict() if isinstance(selected_rows, pd.DataFrame) else selected_rows[0]
            st.session_state.pdf_selecionado = row
            st.session_state.visualizando_pdf = True
            st.rerun()
        
        output = BytesIO()
        df_cnd.to_excel(output, index=False)
        st.download_button("📥 Baixar Excel", data=output.getvalue(), file_name="cnd_municipal.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================================================================
# DASHBOARD PARALEGAL
# ============================================================================

def _formata_estado(s):
    """GO → G.O.  — impede Chrome de traduzir siglas"""
    s = str(s).strip()
    if len(s) == 2 and s.isalpha():
        return f"{s[0]}.{s[1]}."
    return s


def _estado_original(s):
    """G.O. → GO"""
    return s.replace(".", "")


@st.dialog("Detalhes das Empresas")
def _modal_dashboard(titulo, df_show, colunas):
    st.markdown(f"**{titulo}** — {df_show.shape[0]} empresa(s)")
    cols_ok = [c for c in colunas if c in df_show.columns]
    st.dataframe(df_show[cols_ok].reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.fragment
def pagina_dashboard_paralegal():
    import plotly.express as px
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    if "Situação" not in df.columns:
        st.error("Coluna 'Situação' não encontrada.")
        return

    df_ativas = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy()
    total_ativas = df_ativas.shape[0]

    st.markdown("<h2 style='color:#1d3f77;'>Dashboard — Departamento Paralegal</h2>",
                unsafe_allow_html=True)
    st.markdown(f"<p style='font-size:18px;'><b>Total de empresas ativas:</b> {total_ativas}</p>",
                unsafe_allow_html=True)
    st.divider()

    # ── controle de modal: guarda último clique de cada gráfico
    for k in ["ult_estado", "ult_municipio", "ult_cnd"]:
        if k not in st.session_state:
            st.session_state[k] = None

    modal_abrir = None   # apenas UM modal por execução

    # ── EMPRESAS POR ESTADO ──────────────────────────────────────────────────
    st.markdown("### Empresas por Estado")
    st.caption("Clique em uma barra para ver as empresas")

    if "Estado" in df_ativas.columns:
        df_est = df_ativas["Estado"].fillna("N/I").astype(str).str.strip()
        df_est_count = df_est.value_counts().reset_index()
        df_est_count.columns = ["Estado_orig", "Quantidade"]
        df_est_count["Estado"] = df_est_count["Estado_orig"].apply(_formata_estado)

        fig_est = px.bar(
            df_est_count, x="Estado", y="Quantidade",
            color="Quantidade",
            color_continuous_scale=[[0, "#4a90d9"], [1, "#1d3f77"]],
            text="Quantidade",
            custom_data=["Estado_orig"],
        )
        fig_est.update_traces(
            textposition="outside",
            hovertemplate="<b>%{customdata[0]}</b><br>Qtd: %{y}<extra></extra>",
        )
        fig_est.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            coloraxis_showscale=False,
            xaxis=dict(title="", tickfont=dict(size=13), showgrid=False),
            yaxis=dict(title="", showgrid=False, zeroline=False),
            margin=dict(t=30, b=20, l=10, r=10),
            height=350,
            clickmode="event+select",
        )

        ev_est = st.plotly_chart(fig_est, use_container_width=True,
                                  on_select="rerun", key="chart_estado")

        if ev_est and ev_est.selection and ev_est.selection.points:
            pt   = ev_est.selection.points[0]
            sel_raw = (pt.get("customdata") or [None])[0]
            if sel_raw is None:
                sel_raw = _estado_original(pt.get("x", ""))
            if sel_raw and sel_raw != st.session_state["ult_estado"]:
                st.session_state["ult_estado"] = sel_raw
                df_fil = df_ativas[
                    df_ativas["Estado"].fillna("N/I").astype(str).str.strip() == sel_raw
                ]
                modal_abrir = (f"Estado: {sel_raw}", df_fil,
                               ["Razão Social", "CNPJ"])
    else:
        st.warning("Coluna 'Estado' não encontrada.")

    st.divider()

    # ── EMPRESAS POR MUNICÍPIO ───────────────────────────────────────────────
    st.markdown("### Empresas por Município")
    st.caption("Clique em uma barra para ver as empresas")

    if "Município" in df_ativas.columns:
        df_mun_count = (df_ativas["Município"].fillna("N/I").astype(str).str.strip()
                        .value_counts().reset_index())
        df_mun_count.columns = ["Município", "Quantidade"]
        df_mun_count = df_mun_count.sort_values("Quantidade", ascending=True)

        altura_mun = max(400, len(df_mun_count) * 28)

        fig_mun = px.bar(
            df_mun_count, x="Quantidade", y="Município",
            orientation="h",
            color="Quantidade",
            color_continuous_scale=[[0, "#4a90d9"], [1, "#1d3f77"]],
            text="Quantidade",
        )
        fig_mun.update_traces(
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Qtd: %{x}<extra></extra>",
        )
        fig_mun.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            coloraxis_showscale=False,
            xaxis=dict(title="", showgrid=False, zeroline=False),
            yaxis=dict(title="", tickfont=dict(size=12), showgrid=False),
            margin=dict(t=20, r=60, b=20, l=10),
            height=altura_mun,
            clickmode="event+select",
        )

        ev_mun = st.plotly_chart(fig_mun, use_container_width=True,
                                  on_select="rerun", key="chart_municipio")

        if modal_abrir is None and ev_mun and ev_mun.selection and ev_mun.selection.points:
            sel_mun = ev_mun.selection.points[0].get("y")
            if sel_mun and sel_mun != st.session_state["ult_municipio"]:
                st.session_state["ult_municipio"] = sel_mun
                df_fil = df_ativas[
                    df_ativas["Município"].fillna("N/I").astype(str).str.strip() == sel_mun
                ]
                modal_abrir = (f"Município: {sel_mun}", df_fil,
                               ["Razão Social", "CNPJ"])
    else:
        st.warning("Coluna 'Município' não encontrada.")

    st.divider()

    # ── CND MUNICIPAL — SITUAÇÃO ─────────────────────────────────────────────
    st.markdown("### CND Municipal — Situação")
    st.caption("Clique em uma fatia para ver as empresas")

    if "SITUAÇÃO CND MUNICIPAL" in df_ativas.columns:
        df_cnd = df_ativas.copy()
        df_cnd["SIT"] = (df_cnd["SITUAÇÃO CND MUNICIPAL"]
                         .fillna("").astype(str).str.strip())
        df_cnd["CND_LABEL"] = df_cnd["SIT"].apply(
            lambda x: "Outros Municípios" if x == "" or x.upper() == "NAN" else x
        )

        df_cnd_count = df_cnd["CND_LABEL"].value_counts().reset_index()
        df_cnd_count.columns = ["Situação", "Quantidade"]

        color_map = {
            "NEGATIVA":                     "#27ae60",
            "POSITIVA":                     "#e74c3c",
            "POSITIVA COM EFEITO NEGATIVA": "#f39c12",
            "Outros Municípios":            "#bdc3c7",
        }

        # Ordena para barras menores no topo
        df_cnd_count = df_cnd_count.sort_values("Quantidade", ascending=True)

        fig_cnd = px.bar(
            df_cnd_count, x="Quantidade", y="Situação",
            orientation="h",
            text="Quantidade",
            color="Situação",
            color_discrete_map=color_map,
        )
        fig_cnd.update_traces(
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Qtd: %{x}<extra></extra>",
        )
        fig_cnd.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            showlegend=False,
            xaxis=dict(title="", showgrid=False, zeroline=False),
            yaxis=dict(title="", tickfont=dict(size=13), showgrid=False),
            margin=dict(t=20, r=80, b=20, l=10),
            height=max(200, len(df_cnd_count) * 60),
            clickmode="event+select",
        )

        ev_cnd = st.plotly_chart(fig_cnd, use_container_width=True,
                                  on_select="rerun", key="chart_cnd")

        if modal_abrir is None and ev_cnd and ev_cnd.selection and ev_cnd.selection.points:
            pt_cnd  = ev_cnd.selection.points[0]
            sel_cnd = pt_cnd.get("y")
            if sel_cnd and sel_cnd != st.session_state["ult_cnd"]:
                st.session_state["ult_cnd"] = sel_cnd
                df_fil = df_cnd[df_cnd["CND_LABEL"] == sel_cnd]
                modal_abrir = (f"Situação CND: {sel_cnd}", df_fil,
                               ["Razão Social", "CNPJ", "Município",
                                "SITUAÇÃO CND MUNICIPAL"])
    else:
        st.info("Coluna 'SITUAÇÃO CND MUNICIPAL' não encontrada.")

    # ── abre o modal (apenas um por execução) ────────────────────────────────
    if modal_abrir:
        _modal_dashboard(*modal_abrir)

# ============================================================================
# ROTEAMENTO
# ============================================================================

with st.session_state.main_container.container():
    if pagina == "Dashboard":
        pagina_dashboard_paralegal()
    elif pagina == "EMPRESAS":
        pagina_empresas()
    elif pagina == "SIMPLES NACIONAL":
        pagina_simples()
    elif pagina == "REINF":
        pagina_reinf()
    elif pagina == "DCTF WEB":
        pagina_dctf_web()
    elif pagina == "DMS":
        pagina_dms()
    elif pagina == "SERVIÇOS TOMADOS":
        pagina_rest()
    elif pagina == "SEFAZ":
        pagina_sefaz()
    elif pagina == "CND Municipal":
        pagina_cnd_municipal()