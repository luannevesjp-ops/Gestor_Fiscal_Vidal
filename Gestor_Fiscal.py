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

# ============================================================================
# CONFIGURAÇÕES INICIAIS
# ============================================================================

# Configuração da página
st.set_page_config(page_title="LuaTech - Gestão Fiscal", layout="wide")

# URL do Google Sheets
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1bp7qtkKvsMHMvHjGznT6OwyX_YSQWMa3jVvylOJWSxM/export?format=xlsx"
SHEET_EMPRESAS = "GERAL"

# ============================================================================
# CSS E ESTILOS
# ============================================================================

# Script para suprimir erro removeChild
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

# CSS para cabeçalho e sidebar
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
.menu-item {
    color: white;
    font-size:18px;
    padding:10px 6px;
    margin-left:8px;
}
.menu-item:hover {
    opacity:0.85;
    cursor:pointer;
}
</style>
""", unsafe_allow_html=True)

# Container único para todos os grids
grid_container = st.empty()

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

@st.cache_data(ttl=600)
def le_planilha_google(url: str, aba: str):
    """Lê planilha do Google Sheets e retorna DataFrame"""
    try:
        resp = requests.get(url)
        resp.raise_for_status()
        df = pd.read_excel(BytesIO(resp.content), sheet_name=aba, engine='openpyxl')
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        return None


def exibe_aggrid(df, height=400, grid_key="grid"):
    """Exibe AgGrid com configurações padrão"""
    # Cria key única com timestamp para evitar conflitos
    unique_key = f"{grid_key}_{int(time.time() * 1000000)}"
    
    gb = GridOptionsBuilder.from_dataframe(df)
    
    # Configuração padrão para todas as colunas
    gb.configure_default_column(
        filter=True,
        sortable=True,
        editable=False,
        resizable=True
    )
    
    # Configura filtros corretos por tipo de coluna
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            gb.configure_column(col, filter="agNumberColumnFilter")
        else:
            gb.configure_column(col, filter="agTextColumnFilter")
    
    # Configurações gerais do grid com localização em português
    gb.configure_grid_options(
        domLayout="normal",
        floatingFilter=True,
        headerHeight=40,
        rowHeight=30,
        enableBrowserTooltips=True,
        enableCellTextSelection=True,
        suppressMenuHide=True,  # Mantém menu de filtro visível
        # Tradução para português
        localeText={
            'filterOoo': 'Filtrar...',
            'searchOoo': 'Pesquisar...',
            'contains': 'Contém',
            'notContains': 'Não contém',
            'equals': 'Igual',
            'notEqual': 'Diferente',
            'startsWith': 'Começa com',
            'endsWith': 'Termina com',
            'blank': 'Em branco',
            'notBlank': 'Não em branco',
            'andCondition': 'E',
            'orCondition': 'OU',
            'applyFilter': 'Aplicar',
            'resetFilter': 'Limpar',
            'clearFilter': 'Limpar filtro',
            'lessThan': 'Menor que',
            'greaterThan': 'Maior que',
            'lessThanOrEqual': 'Menor ou igual',
            'greaterThanOrEqual': 'Maior ou igual',
            'inRange': 'Entre',
            'pinColumn': 'Fixar coluna',
            'autosizeThiscolumn': 'Ajustar esta coluna',
            'autosizeAllColumns': 'Ajustar todas as colunas',
            'groupBy': 'Agrupar por',
            'resetColumns': 'Resetar colunas',
            'noRowsToShow': 'Nenhum registro para mostrar',
            'loadingOoo': 'Carregando...',
        }
    )
    
    grid_options = gb.build()
    
    # Renderiza o grid
    AgGrid(
        df,
        gridOptions=grid_options,
        height=height,
        key=unique_key,
        fit_columns_on_grid_load=True,
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.MODEL_CHANGED,  # Mantém o filtro ativo
        reload_data=False  # Não recarrega dados ao interagir
    )

# ============================================================================
# AUTENTICAÇÃO / LOGIN
# ============================================================================

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False


def tela_login():
    """Tela de login"""
    st.markdown("<h1 style='text-align:center; color:#0f4fa3;'>Gestão Fiscal</h1>", 
                unsafe_allow_html=True)
    senha = st.text_input("Senha", type="password", max_chars=20)
    entrar = st.button("Entrar")
    
    if entrar:
        if senha == "VIDAL":
            st.session_state["autenticado"] = True
        else:
            st.error("Senha incorreta. Tente novamente.")


if not st.session_state["autenticado"]:
    tela_login()
    st.stop()

# ============================================================================
# SIDEBAR / MENU
# ============================================================================

menu_html = """
<div class="sidebar-lt">
    <h2>LuaTech</h2>
    <p>Automatização de processos</p>
    <hr style="border: 1px solid rgba(255,255,255,0.12); margin-top:12px;">
</div>
"""
st.sidebar.markdown(menu_html, unsafe_allow_html=True)

pagina = st.sidebar.radio(
    "",
    ["EMPRESAS", "SIMPLES NACIONAL", "REINF", "DCTF WEB", "DMS", "SERVIÇOS TOMADOS", "SEFAZ"],
    index=0,
    label_visibility="collapsed"
)

# Controle de mudança de página
if "pagina_atual" not in st.session_state:
    st.session_state["pagina_atual"] = pagina

if st.session_state["pagina_atual"] != pagina:
    st.session_state["pagina_atual"] = pagina
    st.rerun()

# ============================================================================
# PÁGINAS
# ============================================================================

def pagina_empresas():
    """Página de listagem de empresas ativas"""
    st.empty()
    
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    # Competência
    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    # Filtra empresas ATIVAS
    if "Situação" in df.columns:
        df_empresas = df[df["Situação"].astype(str).str.upper() == "ATIVA"]
    else:
        st.error("Coluna 'Situação' não encontrada.")
        return
    
    # Seleção de colunas
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado", "Matriz / Filial", "Situação"]
    df_empresas = df_empresas[[c for c in colunas if c in df_empresas.columns]]
    
    total_empresas = df_empresas.shape[0]
    
    # Título
    st.subheader("Empresas - Apenas ATIVAS")
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Total:</b> {total_empresas} | <b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )
    
    # Exibe AgGrid
    with st.container():
        exibe_aggrid(df_empresas, height=400, grid_key="grid_empresas")
    
    # Download Excel
    output = BytesIO()
    df_empresas.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="empresas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def pagina_simples():
    """Página SIMPLES NACIONAL"""
    st.empty()
    
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    df_ativas = df[
        (df["Situação"].astype(str).str.upper() == "ATIVA") &
        (df["Regime"].astype(str).str.upper() == "SIMPLES NACIONAL")
    ] if "Situação" in df.columns and "Regime" in df.columns else pd.DataFrame()
    
    if df_ativas.empty:
        st.warning("Nenhuma empresa SIMPLES NACIONAL ATIVA encontrada.")
        return
    
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado", "SIMPLES GERADO", "Situação"]
    df_simples = df_ativas[[c for c in colunas if c in df_ativas.columns]]
    
    df_simples["SIMPLES GERADO"] = df_simples["SIMPLES GERADO"].apply(
        lambda x: "Filial" if str(x).upper() == "FILIAL" else ("Concluída" if pd.notna(x) else "Não")
    )
    
    concluidas = df_simples[df_simples["SIMPLES GERADO"].isin(["Concluída", "Filial"])].shape[0]
    nao_concluidas = df_simples[df_simples["SIMPLES GERADO"] == "Não"].shape[0]
    filial = df_simples[df_simples["SIMPLES GERADO"] == "Filial"].shape[0]
    
    st.markdown(
        f"<h2>SIMPLES NACIONAL</h2>"
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} | "
        f"<b>Filial:</b> {filial} | "
        f"<b>Não concluídas:</b> {nao_concluidas} | "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )
    
    time.sleep(1)
    exibe_aggrid(df_simples, height=400, grid_key="grid_simples")
    
    output = BytesIO()
    df_simples.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="simples_nacional.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def pagina_reinf():
    """Página REINF"""
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
            "OK": "Transmitida",
            "FILIAL": "FILIAL",
            "NAN": "Não",
            "": "Não"
        })
    else:
        df_reinf["TRANSMISSÃO"] = "Não"
    
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "TRANSMISSÃO", "Situação"]
    df_reinf = df_reinf[[c for c in colunas if c in df_reinf.columns]]
    
    total_filial = df_reinf[df_reinf["TRANSMISSÃO"] == "FILIAL"].shape[0]
    total_transmitida = df_reinf[df_reinf["TRANSMISSÃO"] == "Transmitida"].shape[0]
    total_nao = df_reinf[df_reinf["TRANSMISSÃO"] == "Não"].shape[0]
    
    st.markdown(
        f"<h2>REINF</h2>"
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Filial:</b> {total_filial} | "
        f"<b>Transmitida:</b> {total_transmitida} | "
        f"<b>Não transmitida:</b> {total_nao} | "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )
    
    time.sleep(1)
    exibe_aggrid(df_reinf, height=400, grid_key="grid_reinf")
    
    output = BytesIO()
    df_reinf.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="reinf.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def pagina_dctf_web():
    st.empty()  # Limpa renderizações anteriores

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)

    if df is None or df.empty:
        st.warning("Nenhum dado encontrado.")
        return

    # Remove None
    df = df.fillna("")

    # Filtra somente ATIVAS
    df = df[df["Situação"].astype(str).str.upper() == "ATIVA"]

    if df.empty:
        st.warning("Nenhuma empresa ATIVA encontrada.")
        return

    # =========================
    # COMPETÊNCIA (MM/YYYY)
    # =========================
    competencia = ""
    if "PERÍODO DE COMPETÊNCIA" in df.columns:
        try:
            competencia_dt = pd.to_datetime(
                df["PERÍODO DE COMPETÊNCIA"].iloc[0],
                errors="coerce"
            )
            if not pd.isna(competencia_dt):
                competencia = competencia_dt.strftime("%m/%Y")
        except Exception:
            competencia = ""

    # =========================
    # PERÍODO (MM-YYYY)
    # =========================
    if "PERÍODO" in df.columns:
        df["PERÍODO"] = (
            pd.to_datetime(df["PERÍODO"], errors="coerce")
            .dt.strftime("%m-%Y")
            .fillna("")
        )

    # =========================
    # DATAFRAME FINAL
    # =========================
    df_dctf = df[[
        "Código",
        "Razão Social",
        "CNPJ",
        "Regime",
        "PERÍODO",
        "ORIGEM",
        "TIPO",
        "SITUAÇÃO DCTF",
        "MATRIZ / FILIAL",
        "Situação"
    ]].copy()

    # =========================
    # TOTALIZADORES
    # =========================
    concluidas = df_dctf[
        df_dctf["SITUAÇÃO DCTF"].astype(str).str.upper() == "ATIVA"
    ].shape[0]

    sem_procuracao = df_dctf[
        df_dctf["SITUAÇÃO DCTF"].astype(str).str.upper() == "SEM PROCURAÇÃO"
    ].shape[0]

    nao_concluidas_total = df_dctf[
        ~df_dctf["SITUAÇÃO DCTF"].astype(str).str.upper().isin(
            ["ATIVA", "SEM PROCURAÇÃO"]
        )
    ].shape[0]

    filiais = df_dctf[
        df_dctf["MATRIZ / FILIAL"].astype(str).str.upper() == "FILIAL"
    ].shape[0]

    nao_concluidas = nao_concluidas_total - filiais
    if nao_concluidas < 0:
        nao_concluidas = 0

    # =========================
    # CABEÇALHO
    # =========================
    st.markdown(
        f"<h2>DCTF WEB</h2>"
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} | "
        f"<b>Sem Procuração:</b> {sem_procuracao} | "
        f"<b>Filiais:</b> {filiais} | "
        f"<b>Não concluídas:</b> {nao_concluidas} | "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )

    # =========================
    # GRID
    # =========================
    gb = GridOptionsBuilder.from_dataframe(df_dctf)
    gb.configure_default_column(resizable=True, filter=True, sortable=True)
    gb.configure_grid_options(domLayout="normal")

    AgGrid(
        df_dctf,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=True,
        height=600
    )
        # =========================
    # DOWNLOAD EXCEL
    # =========================
    output = BytesIO()
    df_dctf.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="dctf_web.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



def pagina_dms():
    """Página DMS"""
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
    
    # Formata valores monetários
    for col in ["FATURAMENTO SERVIÇOS", "BASE DE CÁLCULO ISS"]:
        if col in df_dms.columns:
            df_dms[col] = df_dms[col].fillna(0).astype(float).map(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
    
    if "DMS" in df_dms.columns:
        df_dms["DMS"] = df_dms["DMS"].fillna("")
    
    if "GUIA ISS DMS" not in df_dms.columns:
        df_dms["GUIA ISS DMS"] = "Não"
    else:
        df_dms["GUIA ISS DMS"] = df_dms["GUIA ISS DMS"].astype(str).str.upper().replace({
            "OK": "Guia salva",
            "NAN": "Não",
            "": "Não"
        })
    
    colunas = [
        "Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado",
        "FATURAMENTO SERVIÇOS", "BASE DE CÁLCULO ISS", "XML DMS", "DMS", "GUIA ISS DMS", "Situação"
    ]
    df_dms = df_dms[[c for c in colunas if c in df_dms.columns]]
    
    concluidas = df_dms[df_dms["DMS"].astype(str).str.upper() == "DMS SALVA"].shape[0]
    sem_acesso = df_dms[df_dms["DMS"].astype(str).str.upper() == "SEM ACESSO"].shape[0]
    nao_concluidas = df_dms[
        ~df_dms["DMS"].astype(str).str.upper().isin(["DMS SALVA", "SEM ACESSO"])
    ].shape[0]
    
    st.markdown(
        f"<h2>DMS</h2>"
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} | "
        f"<b>Sem acesso:</b> {sem_acesso} | "
        f"<b>Não concluídas:</b> {nao_concluidas} | "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )
    
    time.sleep(1)
    exibe_aggrid(df_dms, height=400, grid_key="grid_dms")
    
    output = BytesIO()
    df_dms.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="dms.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def pagina_rest():
    """Página SERVIÇOS TOMADOS"""
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
            "REST SALVA": "Concluído",
            "SEM ACESSO": "Sem acesso",
            "": "Não concluído"
        })
    
    colunas = ["Código", "Razão Social", "CNPJ", "REST", "XML REST", "GUIA ISS REST", "Situação"]
    df_rest = df_rest[[c for c in colunas if c in df_rest.columns]]
    
    concluidas = df_rest[df_rest["REST"] == "Concluído"].shape[0]
    sem_acesso = df_rest[df_rest["REST"] == "Sem acesso"].shape[0]
    nao_concluidas = df_rest[df_rest["REST"] == "Não concluído"].shape[0]
    
    st.markdown(
        f"<h2>SERVIÇOS TOMADOS</h2>"
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} | "
        f"<b>Sem acesso:</b> {sem_acesso} | "
        f"<b>Não concluídas:</b> {nao_concluidas} | "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )
    
    time.sleep(1)
    exibe_aggrid(df_rest, height=400, grid_key="grid_rest")
    
    output = BytesIO()
    df_rest.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="servicos_tomados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def pagina_sefaz():
    """Página SEFAZ"""
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
    
    colunas = [
        "Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual",
        "XML ENTRADA", "XML SAÍDA", "IMPORTAÇÃO",
        "TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO", "Situação"
    ]
    df_sefaz = df_sefaz[[c for c in colunas if c in df_sefaz.columns]]
    
    # Ajuste de tipo
    for col in ["TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO"]:
        if col in df_sefaz.columns:
            df_sefaz[col] = pd.to_numeric(df_sefaz[col], errors="coerce").fillna(0)
    
    # Totalizadores
    if "IMPORTAÇÃO" in df_sefaz.columns:
        em_andamento = df_sefaz[df_sefaz["IMPORTAÇÃO"].astype(str).str.upper() == "EM ANDAMENTO"].shape[0]
        outro_estado = df_sefaz[df_sefaz["IMPORTAÇÃO"].astype(str).str.upper() == "OUTRO ESTADO"].shape[0]
        sem_movimento = df_sefaz[df_sefaz["IMPORTAÇÃO"].astype(str).str.upper() == "SEM MOVIMENTO"].shape[0]
        concluido = df_sefaz[df_sefaz["IMPORTAÇÃO"].astype(str).str.upper() == "CONCLUÍDO"].shape[0]
    else:
        em_andamento = outro_estado = sem_movimento = concluido = 0
    
    st.markdown(
        f"<h2>SEFAZ</h2>"
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Em andamento:</b> {em_andamento} | "
        f"<b>Outro Estado:</b> {outro_estado} | "
        f"<b>Sem movimento:</b> {sem_movimento} | "
        f"<b>Concluído:</b> {concluido} | "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )
    
    time.sleep(1)
    exibe_aggrid(df_sefaz, height=400, grid_key="grid_sefaz")
    
    output = BytesIO()
    df_sefaz.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="sefaz.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ============================================================================
# ROTEAMENTO DE PÁGINAS
# ============================================================================

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