# Gestor_Fiscal.py
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from io import BytesIO
import requests
import time


# ---------------- SCRIPT PARA SUPRIMIR ERRO removeChild ----------------
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

# ---------------- Configuração da página --------------
st.set_page_config(page_title="LuaTech - Gestão Fiscal", layout="wide")

# ---------------- CSS cabeçalho ----------------
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
.sidebar-lt h2, .sidebar-lt p { color: white; text-align:center; margin:0; padding:0; }
.menu-item { color: white; font-size:18px; padding:10px 6px; margin-left:8px; }
.menu-item:hover { opacity:0.85; cursor:pointer; }
</style>
""", unsafe_allow_html=True)

# ---------------- Container único para todos os grids ----------------
grid_container = st.empty()

# ---------------- URL do Google Sheets ----------------
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1bp7qtkKvsMHMvHjGznT6OwyX_YSQWMa3jVvylOJWSxM/export?format=xlsx"
SHEET_EMPRESAS = "GERAL"

# ---------------- Função de leitura ----------------
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

# ---------------- Login ----------------
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

def tela_login():
    st.markdown("<h1 style='text-align:center; color:#0f4fa3;'>Gestão Fiscal</h1>", unsafe_allow_html=True)
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

# ---------------- Sidebar / Menu ----------------
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

# ---------------- Controle de mudança de página ----------------
if "pagina_atual" not in st.session_state:
    st.session_state["pagina_atual"] = pagina

# Se mudou de página, limpa e recarrega
if st.session_state["pagina_atual"] != pagina:
    st.session_state["pagina_atual"] = pagina
    st.rerun()

# ---------------- Função para exibir AgGrid ----------------
# ---------------- Função para exibir AgGrid ----------------
def exibe_aggrid(df, height=400, grid_key="grid"):
    from st_aggrid import AgGrid, GridOptionsBuilder
    from st_aggrid.shared import GridUpdateMode
    import pandas as pd
    import time

    # Cria key única com timestamp para evitar conflitos
    unique_key = f"{grid_key}_{int(time.time() * 1000000)}"
    
    gb = GridOptionsBuilder.from_dataframe(df)

    # ---------- Configuração padrão para todas as colunas ----------
    gb.configure_default_column(
        filter=True,
        sortable=True,
        editable=False,
        resizable=True
    )

    # ---------- Configura filtros corretos por tipo de coluna ----------
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            gb.configure_column(col, filter="agNumberColumnFilter")
        else:
            gb.configure_column(col, filter="agTextColumnFilter")

    # ---------- Configurações gerais do grid ----------
    gb.configure_grid_options(
        domLayout="normal",            # cria scroll interno e congela cabeçalho
        floatingFilter=True,           # ativa filtros flutuantes
        headerHeight=40,               # altura do cabeçalho
        rowHeight=30,                  # altura das linhas
        enableBrowserTooltips=True,    # tooltips do navegador
        enableCellTextSelection=True   # permitir seleção de texto
    )

    grid_options = gb.build()

    # ---------- Renderiza o grid com key única ----------
    AgGrid(
        df,
        gridOptions=grid_options,
        height=height,
        key=unique_key,  # Key única com timestamp
        fit_columns_on_grid_load=True,
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.NO_UPDATE
    )


# ---------------- Funções de cada página ----------------
def pagina_empresas():
    st.empty()  # Limpa renderizações anteriores
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    # ---------- Competência ----------
    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""

    # ---------- Filtra empresas ATIVAS ----------
    if "Situação" in df.columns:
        df_empresas = df[df["Situação"].astype(str).str.upper() == "ATIVA"]
    else:
        st.error("Coluna 'Situação' não encontrada.")
        return

    # ---------- Seleção de colunas ----------
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado", "Matriz / Filial", "Situação"]
    df_empresas = df_empresas[[c for c in colunas if c in df_empresas.columns]]

    # ---------- Total de empresas ----------
    total_empresas = df_empresas.shape[0]

    # ---------- Título ----------
    st.subheader("Empresas - Apenas ATIVAS")
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Total:</b> {total_empresas} | <b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )

    # ---------- Exibe AgGrid ----------
    grid_key = "grid_empresas"  # chave fixa para manter filtros
    with st.container():
        exibe_aggrid(df_empresas, height=400, grid_key=grid_key)

    # ---------- Download Excel ----------
    from io import BytesIO
    output = BytesIO()
    df_empresas.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="empresas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def pagina_simples():
    st.empty()  # Limpa renderizações anteriores
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

    # Pausa de 3 segundos antes de mostrar o grid
    time.sleep(1)

    exibe_aggrid(df_simples, height=400, grid_key="grid_simples")
    output = BytesIO()
    df_simples.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="simples_nacional.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_reinf():
    st.empty()  # Limpa renderizações anteriores
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
        df_reinf["TRANSMISSÃO"] = df_reinf["TRANSMISSÃO"].astype(str).str.upper().replace({"OK":"Transmitida","FILIAL":"FILIAL","NAN":"Não","": "Não"})
    else:
        df_reinf["TRANSMISSÃO"] = "Não"
    colunas = ["Código","Razão Social","CNPJ","Regime","TRANSMISSÃO","Situação"]
    df_reinf = df_reinf[[c for c in colunas if c in df_reinf.columns]]
    total_filial = df_reinf[df_reinf["TRANSMISSÃO"]=="FILIAL"].shape[0]
    total_transmitida = df_reinf[df_reinf["TRANSMISSÃO"]=="Transmitida"].shape[0]
    total_nao = df_reinf[df_reinf["TRANSMISSÃO"]=="Não"].shape[0]
    st.markdown(
        f"<h2>REINF</h2>"
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Filial:</b> {total_filial} | "
        f"<b>Transmitida:</b> {total_transmitida} | "
        f"<b>Não transmitida:</b> {total_nao} | "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )

    # Pausa de 3 segundos antes de mostrar o grid
    time.sleep(1)

    exibe_aggrid(df_reinf, height=400, grid_key="grid_reinf")
    output = BytesIO()
    df_reinf.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="reinf.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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
    st.empty()  # Limpa renderizações anteriores
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""

    df_dms = df[df["Situação"].astype(str).str.upper() == "ATIVA"] if "Situação" in df.columns else pd.DataFrame()
    if df_dms.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para DMS.")
        return

    for col in ["FATURAMENTO SERVIÇOS","BASE DE CÁLCULO ISS"]:
        if col in df_dms.columns:
            df_dms[col] = df_dms[col].fillna(0).astype(float).map(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    if "DMS" in df_dms.columns:
        df_dms["DMS"] = df_dms["DMS"].fillna("")
    if "GUIA ISS DMS" not in df_dms.columns:
        df_dms["GUIA ISS DMS"] = "Não"
    else:
        df_dms["GUIA ISS DMS"] = df_dms["GUIA ISS DMS"].astype(str).str.upper().replace({"OK":"Guia salva","NAN":"Não","": "Não"})

    colunas = ["Código","Razão Social","CNPJ","Regime","Município","Estado",
               "FATURAMENTO SERVIÇOS","BASE DE CÁLCULO ISS","XML DMS","DMS","GUIA ISS DMS","Situação"]
    df_dms = df_dms[[c for c in colunas if c in df_dms.columns]]

    concluidas = df_dms[df_dms["DMS"].astype(str).str.upper()=="DMS SALVA"].shape[0]
    sem_acesso = df_dms[df_dms["DMS"].astype(str).str.upper()=="SEM ACESSO"].shape[0]
    nao_concluidas = df_dms[~df_dms["DMS"].astype(str).str.upper().isin(["DMS SALVA","SEM ACESSO"])].shape[0]

    st.markdown(
        f"<h2>DMS</h2>"
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} | "
        f"<b>Sem acesso:</b> {sem_acesso} | "
        f"<b>Não concluídas:</b> {nao_concluidas} | "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True
    )

    # Pausa de 3 segundos antes de mostrar o grid
    time.sleep(1)

    # Exibe tabela
    exibe_aggrid(df_dms, height=400, grid_key="grid_dms")

    # Download Excel
    output = BytesIO()
    df_dms.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="dms.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def pagina_rest():
    st.empty()  # Limpa renderizações anteriores
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
    # Pausa de 3 segundos antes de mostrar o grid
    time.sleep(1)

    # Exibe tabela
    exibe_aggrid(df_rest, height=400, grid_key="grid_rest")

    # Download Excel
    output = BytesIO()
    df_rest.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="servicos_tomados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def pagina_sefaz():
    st.empty()  # Limpa renderizações anteriores
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""

    df_sefaz = df[df["Situação"].astype(str).str.upper() == "ATIVA"] if "Situação" in df.columns else pd.DataFrame()
    if df_sefaz.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para SEFAZ.")
        return

    # Seleciona as colunas
    colunas = [
        "Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual",
        "XML ENTRADA", "XML SAÍDA", "IMPORTAÇÃO",
        "TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO", "Situação"
    ]
    df_sefaz = df_sefaz[[c for c in colunas if c in df_sefaz.columns]]

    # =========================================
    # AJUSTE DE TIPO (SEM ALTERAR ESTRUTURA)
    # =========================================
    for col in ["TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO"]:
    if col in df_sefaz.columns:
        if pd.api.types.is_datetime64_any_dtype(df_sefaz[col]):
            df_sefaz[col] = df_sefaz[col].astype(str)

        df_sefaz[col] = (
            df_sefaz[col]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .str.replace(r"[^\d.-]", "", regex=True)
        )

        df_sefaz[col] = pd.to_numeric(df_sefaz[col], errors="coerce").fillna(0)


    # Calcula totalizadores baseados na coluna IMPORTAÇÃO
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

    # Pausa de 1 segundo antes de mostrar o grid
    time.sleep(1)

    # Exibe tabela
    exibe_aggrid(df_sefaz, height=400, grid_key="grid_sefaz")

    # Download Excel
    output = BytesIO()
    df_sefaz.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel",
        data=output.getvalue(),
        file_name="sefaz.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------- Roteamento das páginas ----------------
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