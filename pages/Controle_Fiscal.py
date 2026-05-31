# ============================================================================
# CONTROLE FISCAL - LUATECH
# ============================================================================

import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from st_aggrid.shared import JsCode

# ============================================================================
# CONFIGURAÇÃO
# ============================================================================

st.set_page_config(page_title="LUATECH - CONTROLE FISCAL", layout="wide")

# Oculta a navegação automática de páginas do Streamlit
st.markdown("""
<style>
[data-testid="stSidebarNav"] { display: none !important; }
</style>
""", unsafe_allow_html=True)

import tempfile

_PROJ_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_DB_LOCAL = os.path.join(_PROJ_DIR, "controle_fiscal.db")

# No Streamlit Cloud a pasta do projeto é somente leitura — usa /tmp como fallback
try:
    _teste = os.path.join(_PROJ_DIR, ".write_test")
    open(_teste, "w").close()
    os.remove(_teste)
    DB_PATH = _DB_LOCAL
except (IOError, OSError):
    DB_PATH = os.path.join(tempfile.gettempdir(), "controle_fiscal.db")

# ============================================================================
# BANCO DE DADOS
# ============================================================================

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS empresas_controle (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        cod             TEXT NOT NULL,
        razao_social    TEXT,
        cnpj            TEXT,
        regime          TEXT,
        matriz_filial   TEXT,
        ie              TEXT,
        im              TEXT,
        uf              TEXT,
        municipio       TEXT,
        grupo           TEXT,
        responsavel_fiscal TEXT,
        observacoes     TEXT,
        ativo           INTEGER DEFAULT 1,
        criado_em       TEXT    DEFAULT (datetime('now','localtime')),
        criado_por      TEXT
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS controle_municipal (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        competencia     TEXT NOT NULL,
        cod             TEXT NOT NULL,
        razao_social    TEXT, cnpj TEXT, regime TEXT,
        im TEXT, uf TEXT, municipio TEXT, responsavel TEXT,
        apuracao_dms    TEXT DEFAULT '',
        iss_proprio     TEXT DEFAULT '',
        fechado_dms     TEXT DEFAULT '',
        apuracao_rest   TEXT DEFAULT '',
        iss_retido      TEXT DEFAULT '',
        fechado_rest    TEXT DEFAULT '',
        observacao      TEXT DEFAULT '',
        conferencia     TEXT DEFAULT '',
        observacoes_geral TEXT DEFAULT '',
        status          TEXT DEFAULT '',
        motivo_pendencia  TEXT DEFAULT '',
        UNIQUE(competencia, cod)
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS controle_estadual (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        competencia     TEXT NOT NULL,
        cod             TEXT NOT NULL,
        razao_social    TEXT, cnpj TEXT, regime TEXT,
        ie TEXT, uf TEXT, municipio TEXT, responsavel TEXT,
        apuracao        TEXT DEFAULT '',
        guia_icms       TEXT DEFAULT '',
        guia_protege    TEXT DEFAULT '',
        fechado         TEXT DEFAULT '',
        sped_fiscal     TEXT DEFAULT '',
        conferencia     TEXT DEFAULT '',
        observacoes     TEXT DEFAULT '',
        status          TEXT DEFAULT '',
        motivo_pendencia  TEXT DEFAULT '',
        UNIQUE(competencia, cod)
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS controle_federal (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        competencia     TEXT NOT NULL,
        cod             TEXT NOT NULL,
        razao_social    TEXT, cnpj TEXT, regime TEXT,
        uf TEXT, municipio TEXT, responsavel TEXT,
        apuracao        TEXT DEFAULT '',
        guia            TEXT DEFAULT '',
        fechado         TEXT DEFAULT '',
        sped_contribuicoes TEXT DEFAULT '',
        dirbi           TEXT DEFAULT '',
        data_envio_dirbi  TEXT DEFAULT '',
        mit             TEXT DEFAULT '',
        data_envio_mit  TEXT DEFAULT '',
        reinf           TEXT DEFAULT '',
        observacao      TEXT DEFAULT '',
        relatorio_fiscal  TEXT DEFAULT '',
        diagnostico_fiscal TEXT DEFAULT '',
        situacao_pendencia TEXT DEFAULT '',
        status          TEXT DEFAULT '',
        motivo_pendencia  TEXT DEFAULT '',
        UNIQUE(competencia, cod)
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS controle_simples (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        competencia     TEXT NOT NULL,
        cod             TEXT NOT NULL,
        razao_social    TEXT, cnpj TEXT, ie TEXT, im TEXT,
        uf TEXT, municipio TEXT, responsavel TEXT,
        apuracao_dms    TEXT DEFAULT '',
        sn              TEXT DEFAULT '',
        das             TEXT DEFAULT '',
        fechado         TEXT DEFAULT '',
        reinf           TEXT DEFAULT '',
        conferencia     TEXT DEFAULT '',
        observacao      TEXT DEFAULT '',
        relatorio_fiscal  TEXT DEFAULT '',
        diagnostico_fiscal TEXT DEFAULT '',
        status          TEXT DEFAULT '',
        motivo_pendencia  TEXT DEFAULT '',
        UNIQUE(competencia, cod)
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS parcelamentos (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        cod TEXT, razao_social TEXT, cnpj TEXT,
        municipal       TEXT DEFAULT '',
        estadual        TEXT DEFAULT '',
        rfb_federal     TEXT DEFAULT '',
        pgfn_federal    TEXT DEFAULT '',
        status_parcelamento TEXT DEFAULT '',
        data_parcelamento   TEXT DEFAULT '',
        observacao      TEXT DEFAULT '',
        email           TEXT DEFAULT '',
        observacao2     TEXT DEFAULT ''
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS senhas_acessos (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        cod TEXT, razao_cnpj TEXT, cnpj TEXT,
        codigo_acesso   TEXT DEFAULT '',
        municipio       TEXT DEFAULT '',
        login           TEXT DEFAULT '',
        senha           TEXT DEFAULT '',
        observacoes     TEXT DEFAULT ''
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS alteracao_empresa (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        data_hora       TEXT DEFAULT (datetime('now','localtime')),
        tipo            TEXT,
        cod             TEXT,
        razao_social    TEXT,
        cnpj            TEXT,
        usuario         TEXT,
        observacao      TEXT
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS obrigacoes_prazos (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        competencia     TEXT NOT NULL,
        regime          TEXT NOT NULL,
        uf              TEXT NOT NULL,
        obrigacao       TEXT NOT NULL,
        responsavel     TEXT DEFAULT '',
        prazo           TEXT DEFAULT '',
        data_realizado  TEXT DEFAULT '',
        status          TEXT DEFAULT '',
        motivo_pendencia  TEXT DEFAULT '',
        UNIQUE(competencia, regime, uf, obrigacao)
    )""")

    conn.commit()
    conn.close()


init_db()

# ============================================================================
# AUTENTICAÇÃO
# ============================================================================

def check_auth():
    nivel = st.session_state.get("nivel_acesso")
    if nivel not in ("GESTOR", "FISCAL"):
        st.markdown("""
        <div style='max-width:400px; margin:80px auto; background:#f4f6fa;
                    padding:40px; border-radius:12px; box-shadow:0 2px 12px rgba(0,0,0,0.1);'>
        <h2 style='text-align:center; color:#1d3f77;'>Controle Fiscal</h2>
        </div>
        """, unsafe_allow_html=True)

        col_center = st.columns([1, 1, 1])[1]
        with col_center:
            st.markdown("#### Informe sua senha de acesso")
            senha = st.text_input("Senha", type="password", key="ctrl_senha_login")
            if st.button("Entrar", use_container_width=True, key="ctrl_btn_login"):
                if senha == "GESTOR":
                    st.session_state["nivel_acesso"] = "GESTOR"
                    st.session_state["autenticado"]  = True
                    st.rerun()
                elif senha == "FISCAL":
                    st.session_state["nivel_acesso"] = "FISCAL"
                    st.session_state["autenticado"]  = True
                    st.rerun()
                else:
                    st.error("Senha incorreta.")
            if st.button("Voltar ao Sistema", use_container_width=True, key="ctrl_btn_voltar_login"):
                st.switch_page("Gestor_Fiscal.py")
        st.stop()


check_auth()
_NIVEL = st.session_state.get("nivel_acesso", "FISCAL")
_GESTOR = _NIVEL == "GESTOR"

# ============================================================================
# HELPERS
# ============================================================================

_STATUS_OPCOES = ["", "Concluído", "Pendente"]

_ROW_STYLE = JsCode("""
function(params) {
    if (!params.data) return {};
    var s = params.data.status || '';
    if (s === 'Concluído') return {background: '#d5f5e3'};
    if (s === 'Pendente')  return {background: '#fadbd8'};
    return {};
}
""")


def _load(table, where="", params=()):
    conn = get_conn()
    sql = f"SELECT * FROM {table}"
    if where:
        sql += f" WHERE {where}"
    df = pd.read_sql_query(sql, conn, params=list(params))
    conn.close()
    return df


def _save_grid(df, table):
    conn = get_conn()
    c = conn.cursor()
    for _, row in df.iterrows():
        d = {k: (v if v is not None else "") for k, v in row.to_dict().items()}
        row_id = d.pop("id", None)
        if not row_id:
            continue
        d.pop("competencia", None)
        sets = ", ".join([f"{k}=?" for k in d])
        vals = list(d.values()) + [row_id]
        c.execute(f"UPDATE {table} SET {sets} WHERE id=?", vals)
    conn.commit()
    conn.close()


def _auto_populate(table, competencia, filtro_fn):
    conn = get_conn()
    existentes = {
        r[0] for r in conn.execute(
            f"SELECT cod FROM {table} WHERE competencia=?", (competencia,)
        ).fetchall()
    }
    df_emp = pd.read_sql_query(
        "SELECT * FROM empresas_controle WHERE ativo=1", conn
    )
    conn.close()
    if df_emp.empty:
        return
    novas = filtro_fn(df_emp)
    novas = novas[~novas["cod"].isin(existentes)]
    if novas.empty:
        return
    conn = get_conn()
    c = conn.cursor()
    for _, emp in novas.iterrows():
        _inserir_controle(c, table, competencia, emp)
    conn.commit()
    conn.close()


def _inserir_controle(c, table, comp, emp):
    e = lambda k: emp.get(k, "") or ""
    if table == "controle_municipal":
        c.execute("""INSERT OR IGNORE INTO controle_municipal
        (competencia,cod,razao_social,cnpj,regime,im,uf,municipio,responsavel)
        VALUES(?,?,?,?,?,?,?,?,?)""",
        (comp, e("cod"), e("razao_social"), e("cnpj"), e("regime"),
         e("im"), e("uf"), e("municipio"), e("responsavel_fiscal")))

    elif table == "controle_estadual":
        c.execute("""INSERT OR IGNORE INTO controle_estadual
        (competencia,cod,razao_social,cnpj,regime,ie,uf,municipio,responsavel)
        VALUES(?,?,?,?,?,?,?,?,?)""",
        (comp, e("cod"), e("razao_social"), e("cnpj"), e("regime"),
         e("ie"), e("uf"), e("municipio"), e("responsavel_fiscal")))

    elif table == "controle_federal":
        c.execute("""INSERT OR IGNORE INTO controle_federal
        (competencia,cod,razao_social,cnpj,regime,uf,municipio,responsavel)
        VALUES(?,?,?,?,?,?,?,?)""",
        (comp, e("cod"), e("razao_social"), e("cnpj"), e("regime"),
         e("uf"), e("municipio"), e("responsavel_fiscal")))

    elif table == "controle_simples":
        c.execute("""INSERT OR IGNORE INTO controle_simples
        (competencia,cod,razao_social,cnpj,ie,im,uf,municipio,responsavel)
        VALUES(?,?,?,?,?,?,?,?,?)""",
        (comp, e("cod"), e("razao_social"), e("cnpj"), e("ie"), e("im"),
         e("uf"), e("municipio"), e("responsavel_fiscal")))


def _build_grid(df, edit_cols=None, height=450, key="grid"):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        resizable=True, filter=True, sortable=True, editable=False, minWidth=100
    )
    for col in df.columns:
        if col == "id" or col == "competencia":
            gb.configure_column(col, hide=True)
        elif edit_cols and col in edit_cols:
            if col == "status":
                gb.configure_column(col, editable=True, width=140,
                                    cellEditor="agSelectCellEditor",
                                    cellEditorParams={"values": _STATUS_OPCOES})
            elif col == "motivo_pendencia":
                gb.configure_column(col, editable=True, width=220)
            else:
                gb.configure_column(col, editable=True)
        elif col in ("cod", "razao_social"):
            gb.configure_column(col, pinned="left", width=160 if col == "razao_social" else 80)

    gb.configure_grid_options(
        domLayout="normal", floatingFilter=True,
        getRowStyle=_ROW_STYLE,
        enableCellTextSelection=True, suppressMenuHide=True,
        localeText={"filterOoo": "Filtrar...", "noRowsToShow": "Nenhum registro"},
    )
    return AgGrid(
        df, gridOptions=gb.build(), height=height, key=key,
        fit_columns_on_grid_load=False, enable_enterprise_modules=False,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        allow_unsafe_jscode=True, reload_data=False,
    )


def _filtros_comp_resp(df, key):
    hoje = datetime.now()
    c1, c2 = st.columns(2)
    with c1:
        comp = st.text_input(
            "Competência (MM/AAAA)",
            value=f"{hoje.month:02d}/{hoje.year}",
            key=f"{key}_comp",
        )
    with c2:
        resps = ["Todos"] + sorted(
            df["responsavel"].dropna().astype(str).unique().tolist()
        ) if "responsavel" in df.columns else ["Todos"]
        resp = st.selectbox("Filtrar por Colaborador", resps, key=f"{key}_resp")
    return comp, resp


def _download_btn(df, nome, key):
    out = BytesIO()
    df.to_excel(out, index=False)
    st.download_button("📥 Baixar Excel", data=out.getvalue(),
                       file_name=f"{nome}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=f"dl_{key}")


def _menu_controle(titulo, table, filtro_fn, colunas, edit_gestor, edit_fiscal, key):
    st.markdown(f"<h2 style='color:#1d3f77;'>{titulo}</h2>", unsafe_allow_html=True)

    # Filtros
    df_vazio = pd.DataFrame(columns=["responsavel"])
    comp, resp = _filtros_comp_resp(df_vazio, key)

    # Auto-popula
    _auto_populate(table, comp, filtro_fn)

    # Carrega
    df = _load(table, "competencia=?", (comp,))
    if df.empty:
        st.info("Nenhuma empresa nesta competência. Verifique se há empresas cadastradas no menu **EMPRESAS** e selecione a competência correta.")
        return
    if resp != "Todos" and "responsavel" in df.columns:
        df = df[df["responsavel"].astype(str) == resp]

    # Colunas para exibição (inclui id oculto)
    cols_exib = ["id"] + [c for c in colunas if c in df.columns]
    df_exib = df[cols_exib].copy()
    df_exib = df_exib.fillna("")

    # Edição conforme nível
    edit_cols = edit_gestor if _GESTOR else edit_fiscal

    st.markdown(
        f"<p style='color:#555; font-size:13px;'>Total: <b>{len(df_exib)}</b> empresa(s) "
        f"| Acesso: <b style='color:{'#1e8449' if _GESTOR else '#ca6f1e'}'>{_NIVEL}</b></p>",
        unsafe_allow_html=True,
    )

    resp_grid = _build_grid(df_exib, edit_cols=edit_cols, key=f"grid_{key}")

    col_s, col_d = st.columns([1, 3])
    with col_s:
        if st.button("💾 Salvar Alterações", key=f"save_{key}", type="primary"):
            df_new = pd.DataFrame(resp_grid["data"])
            if not df_new.empty:
                _save_grid(df_new, table)
                st.success("Alterações salvas!")
                st.rerun()
    with col_d:
        _download_btn(df_exib.drop(columns=["id"], errors="ignore"),
                      f"{key}_{comp.replace('/','-')}", key)


# ============================================================================
# SIDEBAR
# ============================================================================

cor_nivel = "#27ae60" if _GESTOR else "#e67e22"
st.sidebar.markdown(f"""
<div style='background:#1d3f77; padding:12px; text-align:center; margin:-1rem -1rem 0;'>
    <span style='color:white; font-size:17px; font-weight:700;'>CONTROLE FISCAL</span><br>
    <span style='color:{cor_nivel}; font-size:13px; font-weight:600;'>{_NIVEL}</span>
</div>
""", unsafe_allow_html=True)
st.sidebar.markdown("<br>", unsafe_allow_html=True)

_MENUS = [
    "EMPRESAS",
    "CALENDÁRIO",
    "MUNICIPAL",
    "ESTADUAL",
    "FEDERAL",
    "SIMPLES NACIONAL",
    "PARCELAMENTOS",
    "SENHAS E ACESSOS",
    "OBRIGAÇÕES E PRAZOS",
    "PAINEL DE CONTROLE",
    "ALTERAÇÃO DE EMPRESA",
]

pagina = st.sidebar.radio("Menu", _MENUS, label_visibility="collapsed")

st.sidebar.markdown("<hr style='margin:8px 0;'>", unsafe_allow_html=True)

if st.sidebar.button("← Voltar ao Sistema", use_container_width=True):
    st.switch_page("Gestor_Fiscal.py")

if st.sidebar.button("Sair", use_container_width=True):
    st.session_state["nivel_acesso"] = None
    st.session_state["autenticado"]  = False
    st.switch_page("Gestor_Fiscal.py")

# ============================================================================
# MENU: EMPRESAS
# ============================================================================

def pagina_empresas_ctrl():
    st.markdown("<h2 style='color:#1d3f77;'>EMPRESAS</h2>", unsafe_allow_html=True)

    # --- Adicionar empresa (só GESTOR) ---
    if _GESTOR:
        with st.expander("➕ Nova Empresa", expanded=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                n_cod    = st.text_input("Código*", key="ne_cod")
                n_razao  = st.text_input("Razão Social*", key="ne_razao")
                n_cnpj   = st.text_input("CNPJ", key="ne_cnpj")
                n_regime = st.selectbox("Regime", ["", "SIMPLES NACIONAL",
                                        "LUCRO PRESUMIDO", "LUCRO REAL",
                                        "LUCRO ARBITRADO"], key="ne_regime")
            with c2:
                n_mf   = st.selectbox("Matriz/Filial", ["", "MATRIZ", "FILIAL"], key="ne_mf")
                n_ie   = st.text_input("I.E.", key="ne_ie")
                n_im   = st.text_input("I.M.", key="ne_im")
                n_uf   = st.text_input("UF", max_chars=2, key="ne_uf").upper()
            with c3:
                n_mun   = st.text_input("Município", key="ne_mun")
                n_grupo = st.text_input("Grupo", key="ne_grupo")
                n_resp  = st.text_input("Responsável Fiscal", key="ne_resp")
                n_obs   = st.text_input("Observações", key="ne_obs")

            if st.button("Adicionar", type="primary", key="btn_add_emp"):
                if not n_cod.strip() or not n_razao.strip():
                    st.error("Código e Razão Social são obrigatórios.")
                else:
                    conn = get_conn()
                    try:
                        conn.execute("""
                        INSERT INTO empresas_controle
                        (cod,razao_social,cnpj,regime,matriz_filial,ie,im,uf,
                         municipio,grupo,responsavel_fiscal,observacoes,criado_por)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """, (n_cod.strip(), n_razao.strip(), n_cnpj, n_regime, n_mf,
                              n_ie, n_im, n_uf, n_mun, n_grupo, n_resp, n_obs, _NIVEL))
                        conn.execute("""
                        INSERT INTO alteracao_empresa(tipo,cod,razao_social,cnpj,usuario)
                        VALUES('INCLUSÃO',?,?,?,?)
                        """, (n_cod.strip(), n_razao.strip(), n_cnpj, _NIVEL))
                        conn.commit()
                        st.success(f"Empresa '{n_razao}' adicionada!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Erro: {ex}")
                    finally:
                        conn.close()

    # --- Listagem ---
    df = _load("empresas_controle", "ativo=1")
    if df.empty:
        if _GESTOR:
            st.info("Nenhuma empresa cadastrada. Use o formulário acima para adicionar.")
        else:
            st.info("Nenhuma empresa cadastrada.")
        return

    COLS = ["id","cod","razao_social","cnpj","regime","matriz_filial",
            "ie","im","uf","municipio","grupo","responsavel_fiscal","observacoes"]
    df_show = df[[c for c in COLS if c in df.columns]].copy().fillna("")

    RENAME = {
        "cod":"CÓD","razao_social":"RAZÃO SOCIAL","cnpj":"CNPJ",
        "regime":"REGIME","matriz_filial":"MATRIZ/FILIAL","ie":"I.E",
        "im":"I.M","uf":"UF","municipio":"MUNICÍPIO","grupo":"GRUPO",
        "responsavel_fiscal":"RESPONSÁVEL FISCAL","observacoes":"OBSERVAÇÕES",
    }
    df_show = df_show.rename(columns=RENAME)

    edit_emp = list(RENAME.values()) if _GESTOR else []
    resp_grid = _build_grid(df_show, edit_cols=edit_emp, key="grid_emp_ctrl")

    col_s, col_e, col_d = st.columns([1, 1, 3])
    with col_s:
        if _GESTOR and st.button("💾 Salvar Alterações", type="primary", key="save_emp"):
            df_edited = pd.DataFrame(resp_grid["data"])
            df_edited_rev = df_edited.rename(columns={v: k for k, v in RENAME.items()})
            if "id" in df_edited_rev.columns:
                _save_grid(df_edited_rev, "empresas_controle")
            st.success("Salvo!")
            st.rerun()

    with col_e:
        if _GESTOR:
            sel = resp_grid.get("selected_rows", [])
            if sel is not None and len(sel) > 0:
                row = sel.iloc[0].to_dict() if hasattr(sel, "iloc") else sel[0]
                if st.button("🗑️ Excluir Selecionada", key="btn_exc_emp"):
                    cod_s   = row.get("CÓD", "")
                    razao_s = row.get("RAZÃO SOCIAL", "")
                    cnpj_s  = row.get("CNPJ", "")
                    conn = get_conn()
                    conn.execute("UPDATE empresas_controle SET ativo=0 WHERE cod=?", (cod_s,))
                    conn.execute("""INSERT INTO alteracao_empresa(tipo,cod,razao_social,cnpj,usuario)
                        VALUES('EXCLUSÃO',?,?,?,?)""", (cod_s, razao_s, cnpj_s, _NIVEL))
                    conn.commit()
                    conn.close()
                    st.success(f"Empresa '{razao_s}' excluída.")
                    st.rerun()

    with col_d:
        _download_btn(df_show.drop(columns=["id"], errors="ignore"), "empresas_controle", "emp")

# ============================================================================
# MENU: CALENDÁRIO
# ============================================================================

def pagina_calendario():
    st.markdown("<h2 style='color:#1d3f77;'>CALENDÁRIO</h2>", unsafe_allow_html=True)

    hoje = datetime.now()
    comp_sel = st.text_input("Competência (MM/AAAA)",
                             value=f"{hoje.month:02d}/{hoje.year}", key="cal_comp")

    st.markdown("### Resumo da Competência")

    tabelas = [
        ("Municipal",        "controle_municipal"),
        ("Estadual",         "controle_estadual"),
        ("Federal",          "controle_federal"),
        ("Simples Nacional", "controle_simples"),
    ]

    cols = st.columns(len(tabelas))
    for i, (nome, tbl) in enumerate(tabelas):
        df = _load(tbl, "competencia=?", (comp_sel,))
        total  = len(df)
        conc   = (df["status"] == "Concluído").sum() if not df.empty and "status" in df.columns else 0
        pend   = (df["status"] == "Pendente").sum()  if not df.empty and "status" in df.columns else 0
        outros = total - conc - pend
        pct_c  = round(conc / total * 100) if total else 0

        with cols[i]:
            cor_bg = "#d5f5e3" if pct_c >= 80 else "#fadbd8" if pct_c < 50 else "#fef9e7"
            cor_tx = "#186a3b" if pct_c >= 80 else "#922b21" if pct_c < 50 else "#7d6608"
            st.markdown(f"""
            <div style='background:{cor_bg}; border-radius:10px; padding:14px;
                        text-align:center; border:1px solid {cor_tx}22;'>
                <b style='color:#1d3f77; font-size:14px;'>{nome}</b><br>
                <span style='font-size:28px; font-weight:700; color:{cor_tx};'>{pct_c}%</span><br>
                <span style='font-size:12px; color:#555;'>
                    Concluído: {conc} | Pendente: {pend} | Total: {total}
                </span>
            </div>
            """, unsafe_allow_html=True)

    st.divider()
    st.markdown("### Pendências do Mês")

    pendentes = []
    for nome, tbl in tabelas:
        df = _load(tbl, "competencia=? AND status='Pendente'", (comp_sel,))
        if not df.empty:
            df["_menu"] = nome
            pendentes.append(df[["_menu","cod","razao_social","responsavel","motivo_pendencia"]])

    if pendentes:
        df_pend = pd.concat(pendentes, ignore_index=True)
        df_pend.columns = ["Menu","Cód","Razão Social","Responsável","Motivo"]
        st.dataframe(df_pend, use_container_width=True, hide_index=True)
    else:
        st.success("Nenhuma pendência registrada para esta competência!")

# ============================================================================
# MENUS: MUNICIPAL / ESTADUAL / FEDERAL / SIMPLES NACIONAL
# ============================================================================

_EDIT_GESTOR_MUNI  = ["apuracao_dms","iss_proprio","fechado_dms","apuracao_rest",
                      "iss_retido","fechado_rest","observacao","conferencia",
                      "observacoes_geral","responsavel","status","motivo_pendencia"]
_EDIT_FISCAL_MUNI  = ["status","motivo_pendencia","observacoes_geral"]

_EDIT_GESTOR_EST   = ["apuracao","guia_icms","guia_protege","fechado","sped_fiscal",
                      "conferencia","observacoes","responsavel","status","motivo_pendencia"]
_EDIT_FISCAL_EST   = ["status","motivo_pendencia","observacoes"]

_EDIT_GESTOR_FED   = ["apuracao","guia","fechado","sped_contribuicoes","dirbi",
                      "data_envio_dirbi","mit","data_envio_mit","reinf","observacao",
                      "relatorio_fiscal","diagnostico_fiscal","situacao_pendencia",
                      "responsavel","status","motivo_pendencia"]
_EDIT_FISCAL_FED   = ["status","motivo_pendencia","observacao"]

_EDIT_GESTOR_SN    = ["apuracao_dms","sn","das","fechado","reinf","conferencia",
                      "observacao","relatorio_fiscal","diagnostico_fiscal",
                      "responsavel","status","motivo_pendencia"]
_EDIT_FISCAL_SN    = ["status","motivo_pendencia","observacao"]

_COLS_MUNI = ["cod","razao_social","cnpj","regime","im","uf","municipio","responsavel",
              "apuracao_dms","iss_proprio","fechado_dms","apuracao_rest","iss_retido",
              "fechado_rest","observacao","conferencia","observacoes_geral",
              "status","motivo_pendencia"]

_COLS_EST  = ["cod","razao_social","cnpj","regime","ie","uf","municipio","responsavel",
              "apuracao","guia_icms","guia_protege","fechado","sped_fiscal",
              "conferencia","observacoes","status","motivo_pendencia"]

_COLS_FED  = ["cod","razao_social","cnpj","regime","uf","municipio","responsavel",
              "apuracao","guia","fechado","sped_contribuicoes","dirbi","data_envio_dirbi",
              "mit","data_envio_mit","reinf","observacao","relatorio_fiscal",
              "diagnostico_fiscal","situacao_pendencia","status","motivo_pendencia"]

_COLS_SN   = ["cod","razao_social","cnpj","ie","im","uf","municipio","responsavel",
              "apuracao_dms","sn","das","fechado","reinf","conferencia","observacao",
              "relatorio_fiscal","diagnostico_fiscal","status","motivo_pendencia"]

def _filtro_est(df):
    mask = (
        df["ie"].notna() & (df["ie"].astype(str).str.strip() != "") &
        (~df["regime"].astype(str).str.upper().str.contains("SIMPLES", na=False))
    )
    return df[mask]

def _filtro_fed(df):
    mask = (
        (~df["matriz_filial"].astype(str).str.upper().str.contains("FILIAL", na=False)) &
        (~df["regime"].astype(str).str.upper().str.contains("SIMPLES", na=False))
    )
    return df[mask]

def _filtro_sn(df):
    return df[df["regime"].astype(str).str.upper().str.contains("SIMPLES", na=False)]

# ============================================================================
# MENU: PARCELAMENTOS
# ============================================================================

def pagina_parcelamentos():
    st.markdown("<h2 style='color:#1d3f77;'>PARCELAMENTOS</h2>", unsafe_allow_html=True)

    if _GESTOR:
        with st.expander("➕ Novo Parcelamento", expanded=False):
            df_emp = _load("empresas_controle", "ativo=1")
            opcoes_emp = [""] + (df_emp["razao_social"].tolist() if not df_emp.empty else [])
            sel_emp = st.selectbox("Empresa", opcoes_emp, key="parc_emp")
            emp_row = df_emp[df_emp["razao_social"] == sel_emp].iloc[0] if sel_emp else None

            c1, c2, c3 = st.columns(3)
            with c1:
                p_mun  = st.text_input("Municipal", key="parc_mun")
                p_est  = st.text_input("Estadual", key="parc_est")
                p_rfb  = st.text_input("RFB (Federal)", key="parc_rfb")
            with c2:
                p_pgfn = st.text_input("PGFN (Federal)", key="parc_pgfn")
                p_sta  = st.text_input("Status", key="parc_sta")
                p_data = st.date_input("Data", key="parc_data")
            with c3:
                p_obs  = st.text_input("Observação", key="parc_obs")
                p_email= st.text_input("E-mail", key="parc_email")
                p_obs2 = st.text_input("Observação 2", key="parc_obs2")

            if st.button("Adicionar Parcelamento", type="primary", key="btn_add_parc"):
                if emp_row is not None:
                    conn = get_conn()
                    conn.execute("""
                    INSERT INTO parcelamentos(cod,razao_social,cnpj,municipal,estadual,
                    rfb_federal,pgfn_federal,status_parcelamento,data_parcelamento,
                    observacao,email,observacao2)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (emp_row.get("cod",""), emp_row.get("razao_social",""),
                          emp_row.get("cnpj",""), p_mun, p_est, p_rfb, p_pgfn,
                          p_sta, str(p_data), p_obs, p_email, p_obs2))
                    conn.commit()
                    conn.close()
                    st.success("Parcelamento adicionado!")
                    st.rerun()

    df = _load("parcelamentos")
    if df.empty:
        st.info("Nenhum parcelamento cadastrado.")
        return

    RENAME = {
        "cod":"CÓD","razao_social":"RAZÃO SOCIAL","cnpj":"CNPJ",
        "municipal":"MUNICIPAL","estadual":"ESTADUAL","rfb_federal":"RFB (FEDERAL)",
        "pgfn_federal":"PGFN (FEDERAL)","status_parcelamento":"STATUS",
        "data_parcelamento":"DATA","observacao":"OBSERVAÇÃO",
        "email":"E-MAIL","observacao2":"OBSERVAÇÃO 2",
    }
    df_show = df.rename(columns=RENAME).fillna("")

    edit_p = list(RENAME.values()) if _GESTOR else ["STATUS","OBSERVAÇÃO","OBSERVAÇÃO 2"]
    resp = _build_grid(df_show, edit_cols=edit_p, key="grid_parc")

    col_s, col_d = st.columns([1, 3])
    with col_s:
        if st.button("💾 Salvar", type="primary", key="save_parc"):
            df_ed = pd.DataFrame(resp["data"])
            df_ed = df_ed.rename(columns={v: k for k, v in RENAME.items()})
            _save_grid(df_ed, "parcelamentos")
            st.success("Salvo!")
            st.rerun()
    with col_d:
        _download_btn(df_show.drop(columns=["id"], errors="ignore"), "parcelamentos", "parc")

# ============================================================================
# MENU: SENHAS E ACESSOS
# ============================================================================

def pagina_senhas():
    st.markdown("<h2 style='color:#1d3f77;'>SENHAS E ACESSOS</h2>", unsafe_allow_html=True)

    if not _GESTOR:
        st.warning("Acesso restrito. Somente o nível GESTOR pode visualizar senhas.")
        return

    with st.expander("➕ Novo Registro", expanded=False):
        df_emp = _load("empresas_controle", "ativo=1")
        opcoes = [""] + (df_emp["razao_social"].tolist() if not df_emp.empty else [])
        sel = st.selectbox("Empresa", opcoes, key="sen_emp")
        emp_r = df_emp[df_emp["razao_social"] == sel].iloc[0] if sel else None

        c1, c2 = st.columns(2)
        with c1:
            s_cod_ac = st.text_input("Código de Acesso", key="sen_cod")
            s_mun    = st.text_input("Município", key="sen_mun")
            s_login  = st.text_input("Login", key="sen_login")
        with c2:
            s_senha  = st.text_input("Senha", type="password", key="sen_senha")
            s_obs    = st.text_input("Observações", key="sen_obs")

        if st.button("Adicionar", type="primary", key="btn_add_sen"):
            if emp_r is not None:
                conn = get_conn()
                conn.execute("""
                INSERT INTO senhas_acessos(cod,razao_cnpj,cnpj,codigo_acesso,municipio,login,senha,observacoes)
                VALUES(?,?,?,?,?,?,?,?)
                """, (emp_r.get("cod",""), emp_r.get("razao_social",""), emp_r.get("cnpj",""),
                      s_cod_ac, s_mun, s_login, s_senha, s_obs))
                conn.commit()
                conn.close()
                st.success("Registro adicionado!")
                st.rerun()

    df = _load("senhas_acessos")
    if df.empty:
        st.info("Nenhuma senha cadastrada.")
        return

    RENAME = {
        "cod":"CÓD","razao_cnpj":"RAZÃO SOCIAL","cnpj":"CNPJ",
        "codigo_acesso":"CÓDIGO ACESSO","municipio":"MUNICÍPIO",
        "login":"LOGIN","senha":"SENHA","observacoes":"OBSERVAÇÕES",
    }
    df_show = df.rename(columns=RENAME).fillna("")
    resp = _build_grid(df_show, edit_cols=list(RENAME.values()), key="grid_sen")

    col_s, col_d = st.columns([1, 3])
    with col_s:
        if st.button("💾 Salvar", type="primary", key="save_sen"):
            df_ed = pd.DataFrame(resp["data"])
            df_ed = df_ed.rename(columns={v: k for k, v in RENAME.items()})
            _save_grid(df_ed, "senhas_acessos")
            st.success("Salvo!")
            st.rerun()
    with col_d:
        _download_btn(df_show.drop(columns=["id"], errors="ignore"), "senhas_acessos", "sen")

# ============================================================================
# MENU: OBRIGAÇÕES E PRAZOS
# ============================================================================

_OBRIG_SIMPLES = {
    "TODOS": ["REST","DMS","GUIA","DIFAL USO","DIFAL COMERC.","PGDAS","DAS","REINF","DEFIS","SITUAÇÃO FISCAL"],
    "SP":    ["REST","DMS","GUIA","DIFAL USO","DIFAL COMERC.","PGDAS","DAS","REINF","DEFIS","SITUAÇÃO FISCAL"],
}

_OBRIG_PRESUMIDO = {
    "GO":  ["REST","DMS","GUIA","ICMS","GUIA ICMS","SPED FISCAL","ASSESSORIA ESTADUAL","PIS/COFINS","EFD CONTRIB.","REINF","MIT-DCTFWeb","DIRBI","SITUAÇÃO FISCAL"],
    "SC":  ["REST","DMS","GUIA","ICMS","GUIA ICMS","SPED FISCAL","DIME","PIS/COFINS","EFD CONTRIB.","REINF","MIT-DCTFWeb","DIRBI","SITUAÇÃO FISCAL"],
    "TO":  ["REST","DMS","GUIA","ICMS","GUIA ICMS","SPED FISCAL","GIAM","PIS/COFINS","EFD CONTRIB.","REINF","MIT-DCTFWeb","DIRBI","SITUAÇÃO FISCAL"],
    "MS":  ["REST","DMS","GUIA","ICMS","GUIA ICMS","SPED FISCAL","ASSESSORIA ESTADUAL","PIS/COFINS","EFD CONTRIB.","REINF","MIT-DCTFWeb","DIRBI","SITUAÇÃO FISCAL"],
    "MG":  ["REST","DMS","GUIA","ICMS","GUIA ICMS","SPED FISCAL","ASSESSORIA ESTADUAL","PIS/COFINS","EFD CONTRIB.","REINF","MIT-DCTFWeb","DIRBI","SITUAÇÃO FISCAL"],
}


def _auto_populate_obrigacoes(comp):
    conn = get_conn()
    for uf, obrs in _OBRIG_SIMPLES.items():
        for ob in obrs:
            conn.execute("""
            INSERT OR IGNORE INTO obrigacoes_prazos
            (competencia, regime, uf, obrigacao)
            VALUES (?, 'SIMPLES NACIONAL', ?, ?)
            """, (comp, uf, ob))
    for uf, obrs in _OBRIG_PRESUMIDO.items():
        for ob in obrs:
            conn.execute("""
            INSERT OR IGNORE INTO obrigacoes_prazos
            (competencia, regime, uf, obrigacao)
            VALUES (?, 'PRESUMIDO/REAL', ?, ?)
            """, (comp, uf, ob))
    conn.commit()
    conn.close()


def pagina_obrigacoes():
    st.markdown("<h2 style='color:#1d3f77;'>OBRIGAÇÕES E PRAZOS</h2>", unsafe_allow_html=True)

    hoje = datetime.now()
    comp = st.text_input("Competência (MM/AAAA)",
                         value=f"{hoje.month:02d}/{hoje.year}", key="obr_comp")

    _auto_populate_obrigacoes(comp)
    df_all = _load("obrigacoes_prazos", "competencia=?", (comp,))

    edit_obr = ["responsavel","prazo","data_realizado","status","motivo_pendencia"] \
               if _GESTOR else ["status","motivo_pendencia","data_realizado"]

    tab_sn, tab_pr = st.tabs(["SIMPLES NACIONAL", "PRESUMIDO / REAL"])

    with tab_sn:
        for uf in sorted(_OBRIG_SIMPLES.keys()):
            df_sn = df_all[(df_all["regime"]=="SIMPLES NACIONAL") & (df_all["uf"]==uf)].copy()
            if df_sn.empty:
                continue
            st.markdown(f"**Estado: {uf}**" if uf != "TODOS" else "**Todos os Estados**")
            RENAME_O = {
                "obrigacao":"OBRIGAÇÃO","responsavel":"RESPONSÁVEL","prazo":"PRAZO",
                "data_realizado":"REALIZADO","status":"STATUS","motivo_pendencia":"MOTIVO",
            }
            df_sn_show = df_sn[["id"] + list(RENAME_O.keys())].rename(columns=RENAME_O).fillna("")
            resp = _build_grid(df_sn_show,
                               edit_cols=[RENAME_O[k] for k in edit_obr if k in RENAME_O],
                               height=250, key=f"grid_obr_sn_{uf}")
            if st.button("💾 Salvar", key=f"save_obr_sn_{uf}", type="primary"):
                df_ed = pd.DataFrame(resp["data"]).rename(columns={v:k for k,v in RENAME_O.items()})
                _save_grid(df_ed, "obrigacoes_prazos")
                st.success("Salvo!")
                st.rerun()

    with tab_pr:
        for uf in sorted(_OBRIG_PRESUMIDO.keys()):
            df_pr = df_all[(df_all["regime"]=="PRESUMIDO/REAL") & (df_all["uf"]==uf)].copy()
            if df_pr.empty:
                continue
            st.markdown(f"**Estado: {uf}**")
            RENAME_O = {
                "obrigacao":"OBRIGAÇÃO","responsavel":"RESPONSÁVEL","prazo":"PRAZO",
                "data_realizado":"REALIZADO","status":"STATUS","motivo_pendencia":"MOTIVO",
            }
            df_pr_show = df_pr[["id"] + list(RENAME_O.keys())].rename(columns=RENAME_O).fillna("")
            resp = _build_grid(df_pr_show,
                               edit_cols=[RENAME_O[k] for k in edit_obr if k in RENAME_O],
                               height=250, key=f"grid_obr_pr_{uf}")
            if st.button("💾 Salvar", key=f"save_obr_pr_{uf}", type="primary"):
                df_ed = pd.DataFrame(resp["data"]).rename(columns={v:k for k,v in RENAME_O.items()})
                _save_grid(df_ed, "obrigacoes_prazos")
                st.success("Salvo!")
                st.rerun()

# ============================================================================
# MENU: PAINEL DE CONTROLE
# ============================================================================

def pagina_painel():
    import plotly.graph_objects as go
    st.markdown("<h2 style='color:#1d3f77;'>PAINEL DE CONTROLE</h2>", unsafe_allow_html=True)

    hoje = datetime.now()
    comp = st.text_input("Competência (MM/AAAA)",
                         value=f"{hoje.month:02d}/{hoje.year}", key="pain_comp")

    tabelas = [
        ("Municipal",        "controle_municipal"),
        ("Estadual",         "controle_estadual"),
        ("Federal",          "controle_federal"),
        ("Simples Nacional", "controle_simples"),
        ("Obrigações",       "obrigacoes_prazos"),
    ]

    dados = []
    for nome, tbl in tabelas:
        where = "competencia=?" if tbl != "parcelamentos" else ""
        params = (comp,) if where else ()
        df = _load(tbl, where, params)
        total = len(df)
        conc  = (df["status"] == "Concluído").sum() if not df.empty and "status" in df.columns else 0
        pend  = (df["status"] == "Pendente").sum()  if not df.empty and "status" in df.columns else 0
        dados.append({"Menu": nome, "Total": total, "Concluído": int(conc), "Pendente": int(pend),
                      "Sem Status": int(total - conc - pend)})

    st.markdown("### Resumo por Menu")
    CARD_COLS = st.columns(len(dados))
    for i, d in enumerate(dados):
        pct = round(d["Concluído"] / d["Total"] * 100) if d["Total"] else 0
        cor = "#27ae60" if pct >= 80 else "#e74c3c" if pct < 50 else "#e67e22"
        with CARD_COLS[i]:
            st.markdown(f"""
            <div style='background:#f4f6fa; border-radius:10px; padding:12px;
                        text-align:center; border-top:4px solid {cor};'>
                <b style='color:#1d3f77; font-size:13px;'>{d["Menu"]}</b><br>
                <span style='font-size:26px; font-weight:700; color:{cor};'>{pct}%</span><br>
                <span style='font-size:11px; color:#555;'>
                    ✅ {d["Concluído"]} | ⚠️ {d["Pendente"]} | — {d["Sem Status"]}
                </span>
            </div>
            """, unsafe_allow_html=True)

    st.divider()

    # Gráfico geral
    df_resumo = pd.DataFrame(dados)
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Concluído", x=df_resumo["Menu"],
                         y=df_resumo["Concluído"], marker_color="#27ae60"))
    fig.add_trace(go.Bar(name="Pendente",  x=df_resumo["Menu"],
                         y=df_resumo["Pendente"],  marker_color="#e74c3c"))
    fig.add_trace(go.Bar(name="Sem Status", x=df_resumo["Menu"],
                         y=df_resumo["Sem Status"], marker_color="#bdc3c7"))
    fig.update_layout(
        barmode="stack", height=350,
        paper_bgcolor="white", plot_bgcolor="white",
        legend=dict(orientation="h", y=1.1),
        margin=dict(t=20, b=20, l=10, r=10),
        xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
    )
    st.plotly_chart(fig, use_container_width=True, key="chart_painel")

    st.divider()

    # Lista de pendentes
    st.markdown("### Pendências Detalhadas")
    pendentes = []
    for nome, tbl in tabelas:
        where = "competencia=? AND status='Pendente'" if tbl != "parcelamentos" else "status_parcelamento='Pendente'"
        params = (comp,) if tbl != "parcelamentos" else ()
        df = _load(tbl, where, params)
        if not df.empty:
            df["_menu"] = nome
            cols_disp = [c for c in ["_menu","cod","razao_social","responsavel","motivo_pendencia"] if c in df.columns]
            pendentes.append(df[cols_disp])

    if pendentes:
        df_pend = pd.concat(pendentes, ignore_index=True)
        df_pend.columns = [c.replace("_menu","Menu").replace("cod","Cód")
                           .replace("razao_social","Razão Social")
                           .replace("responsavel","Responsável")
                           .replace("motivo_pendencia","Motivo") for c in df_pend.columns]
        st.dataframe(df_pend, use_container_width=True, hide_index=True)
    else:
        st.success("Nenhuma pendência encontrada para esta competência!")

# ============================================================================
# MENU: ALTERAÇÃO DE EMPRESA
# ============================================================================

def pagina_alteracao():
    st.markdown("<h2 style='color:#1d3f77;'>ALTERAÇÃO DE EMPRESA</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#666;'>Registro automático de todas as inclusões e exclusões de empresas.</p>",
                unsafe_allow_html=True)

    df = _load("alteracao_empresa")
    if df.empty:
        st.info("Nenhuma alteração registrada.")
        return

    RENAME = {
        "data_hora":"DATA/HORA","tipo":"TIPO","cod":"CÓD",
        "razao_social":"RAZÃO SOCIAL","cnpj":"CNPJ",
        "usuario":"USUÁRIO","observacao":"OBSERVAÇÃO",
    }
    df_show = df[[c for c in RENAME if c in df.columns]].rename(columns=RENAME).fillna("")
    df_show = df_show.sort_values("DATA/HORA", ascending=False)

    tipo_style = JsCode("""
    function(params) {
        if (params.data['TIPO'] === 'INCLUSÃO') return {background:'#d5f5e3', color:'#186a3b'};
        if (params.data['TIPO'] === 'EXCLUSÃO') return {background:'#fadbd8', color:'#922b21'};
        return {};
    }""")

    gb = GridOptionsBuilder.from_dataframe(df_show)
    gb.configure_default_column(resizable=True, filter=True, sortable=True, editable=False)
    gb.configure_grid_options(domLayout="normal", floatingFilter=True, getRowStyle=tipo_style)
    AgGrid(df_show, gridOptions=gb.build(), height=500,
           key="grid_alt_emp", fit_columns_on_grid_load=False,
           enable_enterprise_modules=False, update_mode=GridUpdateMode.NO_UPDATE,
           allow_unsafe_jscode=True)

    _download_btn(df_show, "alteracao_empresa", "alt_emp")

# ============================================================================
# ROTEAMENTO
# ============================================================================

if pagina == "EMPRESAS":
    pagina_empresas_ctrl()

elif pagina == "CALENDÁRIO":
    pagina_calendario()

elif pagina == "MUNICIPAL":
    _menu_controle(
        "MUNICIPAL", "controle_municipal",
        lambda df: df,  # todas as empresas ativas
        _COLS_MUNI, _EDIT_GESTOR_MUNI, _EDIT_FISCAL_MUNI, "municipal",
    )

elif pagina == "ESTADUAL":
    _menu_controle(
        "ESTADUAL", "controle_estadual",
        _filtro_est,  # com IE, não Simples
        _COLS_EST, _EDIT_GESTOR_EST, _EDIT_FISCAL_EST, "estadual",
    )

elif pagina == "FEDERAL":
    _menu_controle(
        "FEDERAL", "controle_federal",
        _filtro_fed,  # não filial, não Simples
        _COLS_FED, _EDIT_GESTOR_FED, _EDIT_FISCAL_FED, "federal",
    )

elif pagina == "SIMPLES NACIONAL":
    _menu_controle(
        "SIMPLES NACIONAL", "controle_simples",
        _filtro_sn,  # só Simples Nacional
        _COLS_SN, _EDIT_GESTOR_SN, _EDIT_FISCAL_SN, "simples",
    )

elif pagina == "PARCELAMENTOS":
    pagina_parcelamentos()

elif pagina == "SENHAS E ACESSOS":
    pagina_senhas()

elif pagina == "OBRIGAÇÕES E PRAZOS":
    pagina_obrigacoes()

elif pagina == "PAINEL DE CONTROLE":
    pagina_painel()

elif pagina == "ALTERAÇÃO DE EMPRESA":
    pagina_alteracao()
