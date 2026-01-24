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
    padding-top: 18px;
    padding-bottom: 18px;
}
.sidebar-lt h2, .sidebar-lt p {
    color: white;
    text-align:center;
    margin:0;
    padding:0;
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
# SIDEBAR
# ============================================================================

st.sidebar.markdown("""
<div class="sidebar-lt">
    <h2>LuaTech</h2>
    <p>Automatização de processos</p>
    <hr style="border: 1px solid rgba(255,255,255,0.12); margin-top:12px;">
</div>
""", unsafe_allow_html=True)

pagina = st.sidebar.radio("", ["EMPRESAS", "SIMPLES NACIONAL", "REINF", "DCTF WEB", 
                                "DMS", "SERVIÇOS TOMADOS", "SEFAZ", "CND Municipal"],
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
        if st.button("← Voltar para a lista", type="primary"):
            st.session_state.visualizando_pdf = False
            st.session_state.pdf_selecionado = None
            st.rerun()
        
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
# ROTEAMENTO
# ============================================================================

with st.session_state.main_container.container():
    if pagina == "EMPRESAS":
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