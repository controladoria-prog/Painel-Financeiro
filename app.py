"""
Painel Analítico de Performance Estratégica — Controladoria B&A
=================================================================
Dashboard financeiro em Streamlit: consolida Orçado vs. Realizado,
DRE detalhada, histórico mensal e projeções de tendência.

Fontes de dados: Google Sheets (com fallback para arquivos locais em rede).
"""

import base64
import io
import os
import re
from datetime import datetime
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ============================================================================
st.set_page_config(
    page_title="Controladoria B&A - Painel Financeiro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

FUSO_BR = ZoneInfo("America/Sao_Paulo")

# ============================================================================
# 2. DESIGN SYSTEM — paleta única, usada tanto no CSS quanto nos gráficos
# ============================================================================
COLORS = {
    "bg": "#0B0E14",
    "sidebar_bg": "#080A10",
    "surface": "#141824",
    "surface_alt": "#1A1F2E",
    "border": "#232838",
    "border_soft": "#1D2230",
    "text": "#F1F5F9",
    "text_muted": "#8B95A5",
    "primary": "#4C8DFF",
    "primary_soft": "rgba(76, 141, 255, 0.14)",
    "secondary": "#5B6472",
    "positive": "#3ECF8E",
    "negative": "#F76E6E",
    "warning": "#F5A623",
    "muted_line": "#94A3B8",
}

FONT_STACK = "'Inter', 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif"

# Config fixa do Plotly: dashboard "travado" (sem zoom/pan/modebar) para uso executivo
CONFIG_PLOTLY_TRAVADO = {
    "staticPlot": False,
    "displayModeBar": False,
    "scrollZoom": False,
    "doubleClick": False,
    "responsive": True,
}


def cor_variacao(valor):
    """Retorna verde/vermelho conforme o sinal do valor (positivo/negativo)."""
    if pd.isna(valor):
        return COLORS["text_muted"]
    return COLORS["positive"] if valor >= 0 else COLORS["negative"]


def estilo_grafico(fig, height=400, **overrides):
    """Aplica o layout visual padrão do painel a uma figura Plotly."""
    layout = dict(
        paper_bgcolor=COLORS["surface"],
        plot_bgcolor=COLORS["surface"],
        font=dict(color=COLORS["text_muted"], family=FONT_STACK, size=12),
        margin=dict(l=20, r=20, t=30, b=60),
        height=height,
        hoverlabel=dict(
            bgcolor=COLORS["surface_alt"],
            bordercolor=COLORS["border"],
            font=dict(color=COLORS["text"], family=FONT_STACK, size=12),
        ),
    )
    layout.update(overrides)
    fig.update_layout(**layout)
    return fig


def kpi_card_html(label, value, value_color, subtext="", subtext_color=None,
                   progress_pct=None, icon="📌"):
    """Gera o HTML (string) de um cartão de KPI, sem renderizar."""
    subtext_color = subtext_color or COLORS["text_muted"]
    progress_html = ""
    if progress_pct is not None:
        pct = max(0.0, min(100.0, progress_pct))
        progress_html = (
            f'<div class="progress-container">'
            f'<div class="progress-bar" style="width:{pct:.1f}%;"></div>'
            f"</div>"
        )
    return (
        f'<div class="kpi-card" style="border-top-color:{value_color};">'
        f'<div class="kpi-top">'
        f'<span class="kpi-label">{label}</span>'
        f'<span class="kpi-icon" style="background:{value_color}22;">{icon}</span>'
        f"</div>"
        f'<div class="kpi-value" style="color:{value_color};">{value}</div>'
        f'<div class="kpi-subtext" style="color:{subtext_color};">{subtext}</div>'
        f"{progress_html}"
        f"</div>"
    )


def render_kpi_row(cards):
    """Renderiza uma linha de cartões de KPI em flexbox — usada no cabeçalho fixo (sticky)."""
    html = "".join(kpi_card_html(**c) for c in cards)
    return f'<div class="kpi-row">{html}</div>'


# ============================================================================
# 3. ESTILIZAÇÃO CSS GLOBAL
# ============================================================================
st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        html, body, [class*="css"] {{
            font-family: {FONT_STACK};
        }}

        .stApp {{
            background: linear-gradient(180deg, {COLORS["bg"]} 0%, #0D111A 100%);
            color: {COLORS["text"]};
        }}

        [data-testid="stSidebar"] {{
            background-color: {COLORS["sidebar_bg"]};
            border-right: 1px solid {COLORS["border"]};
        }}
        [data-testid="stSidebarUserContent"] {{
            padding-bottom: 80px !important;
        }}
        div[data-baseweb="popover"] {{
            z-index: 999999 !important;
        }}
        /* Garante que listas suspensas (ex: seletor de mês) sempre caibam na tela,
           com rolagem interna, em vez de serem cortadas pela barra de tarefas.
           Vários seletores cobrem diferentes versões do Streamlit/BaseWeb. */
        ul[data-baseweb="menu"],
        div[data-baseweb="menu"],
        div[data-testid="stSelectboxVirtualDropdown"],
        div[data-testid="stSelectboxVirtualDropdown"] ul,
        div[data-baseweb="popover"] div[role="listbox"],
        div[data-baseweb="popover"] ul[role="listbox"] {{
            max-height: 260px !important;
            overflow-y: auto !important;
        }}

        /* Header nativo do Streamlit — transparente */
        header[data-testid="stHeader"] {{
            background: transparent !important;
            pointer-events: none;
        }}
        header[data-testid="stHeader"] * {{
            pointer-events: auto;
        }}
        [data-testid="stAppDeployButton"] {{
            display: none !important;
        }}
        [data-testid="stHeaderActionElements"] {{
            position: fixed !important;
            top: 10px;
            right: 10px;
            z-index: 999999;
        }}

        /* Sidebar: marca + rótulos de filtro */
        .sidebar-brand {{
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 4px 0 16px 0;
            margin-bottom: 12px;
            border-bottom: 1px solid {COLORS["border"]};
        }}
        .sidebar-brand .dot {{
            width: 10px; height: 10px; border-radius: 50%;
            background: {COLORS["primary"]};
            box-shadow: 0 0 10px {COLORS["primary"]};
        }}
        .sidebar-brand span.title {{
            font-size: 14px; font-weight: 700; color: {COLORS["text"]};
            letter-spacing: 0.2px;
        }}
        .sidebar-brand span.subtitle {{
            display: block; font-size: 11px; color: {COLORS["text_muted"]};
        }}
        section[data-testid="stSidebar"] label p {{
            font-size: 12.5px !important;
            font-weight: 600 !important;
            color: {COLORS["text_muted"]} !important;
            text-transform: uppercase;
            letter-spacing: 0.4px;
        }}

        /* Cabeçalho principal */
        .main-header {{
            background: linear-gradient(135deg, {COLORS["surface"]} 0%, {COLORS["surface_alt"]} 100%);
            padding: 16px 22px;
            border-radius: 10px;
            border: 1px solid {COLORS["border"]};
            border-left: 3px solid {COLORS["primary"]};
            margin-bottom: 14px;
            box-shadow: 0 4px 18px rgba(0,0,0,0.25);
        }}
        .main-header h3 {{
            margin: 0 !important;
            padding: 0 !important;
            font-size: 19px !important;
            color: {COLORS["text"]} !important;
            font-weight: 700;
            letter-spacing: 0.2px;
        }}
        .main-header p {{
            margin: 5px 0 0 0 !important;
            font-size: 12.5px !important;
            color: {COLORS["text_muted"]} !important;
        }}
        .main-header .chip {{
            display: inline-block;
            background: {COLORS["primary_soft"]};
            color: {COLORS["primary"]};
            border-radius: 20px;
            padding: 1px 10px;
            font-weight: 600;
            font-size: 11.5px;
        }}

        /* Linha de cartões de KPI em flexbox (usada no cabeçalho fixo) */
        .kpi-row {{
            display: flex;
            gap: 14px;
            align-items: stretch;
        }}
        .kpi-row .kpi-card {{
            flex: 1 1 0;
            min-width: 0;
            margin-bottom: 0;
        }}

        /* Cartões de KPI */
        .kpi-card {{
            background-color: {COLORS["surface"]};
            padding: 16px 18px;
            border-radius: 10px;
            border: 1px solid {COLORS["border"]};
            margin-bottom: 10px;
            transition: transform 0.15s ease, border-color 0.15s ease;
        }}
        .kpi-card:hover {{
            transform: translateY(-2px);
            border-color: {COLORS["primary"]};
        }}
        .kpi-top {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 8px;
        }}
        .kpi-label {{
            font-size: 11px !important;
            font-weight: 700 !important;
            color: {COLORS["text_muted"]} !important;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .kpi-icon {{
            font-size: 13px;
            opacity: 0.95;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 26px;
            height: 26px;
            border-radius: 8px;
        }}
        .kpi-value {{
            font-size: 22px !important;
            font-weight: 700 !important;
            letter-spacing: -0.3px;
            line-height: 1.2;
        }}
        .kpi-subtext {{
            font-size: 12px;
            font-weight: 500;
            margin-top: 4px;
            color: {COLORS["muted_line"]};
        }}

        .progress-container {{
            width: 100%;
            background-color: {COLORS["border"]};
            border-radius: 3px;
            height: 6px;
            margin-top: 10px;
            overflow: hidden;
        }}
        .progress-bar {{
            height: 100%;
            background: linear-gradient(90deg, {COLORS["primary"]}, #7AB0FF);
            border-radius: 3px;
        }}

        /* Chart / section captions */
        .section-title {{
            font-size: 14px;
            font-weight: 700;
            color: {COLORS["text"]};
            margin-bottom: 2px;
        }}

        /* Abas */
        button[data-baseweb="tab"] {{
            background-color: {COLORS["surface"]};
            border-radius: 6px;
            color: {COLORS["text_muted"]} !important;
            padding: 7px 16px;
            margin-right: 4px;
            border: 1px solid {COLORS["border"]};
            font-size: 13px;
            font-weight: 500;
        }}
        button[data-baseweb="tab"][aria-selected="true"] {{
            background-color: {COLORS["primary_soft"]} !important;
            color: {COLORS["text"]} !important;
            font-weight: 700;
            border-color: {COLORS["primary"]} !important;
        }}
        hr {{
            border-color: {COLORS["border"]} !important;
        }}

        /* Travar primeira coluna em tabelas */
        div[data-testid="stDataFrame"] div[role="grid"] div[role="row"] div[role="gridcell"]:first-child,
        div[data-testid="stDataFrame"] div[role="grid"] div[role="row"] div[role="columnheader"]:first-child {{
            position: sticky;
            left: 0;
            background-color: {COLORS["surface"]} !important;
            z-index: 3;
            border-right: 1px solid {COLORS["border"]};
        }}

        /* Scrollbar discreta */
        ::-webkit-scrollbar {{ height: 8px; width: 8px; }}
        ::-webkit-scrollbar-track {{ background: {COLORS["bg"]}; }}
        ::-webkit-scrollbar-thumb {{ background: {COLORS["border"]}; border-radius: 4px; }}
        ::-webkit-scrollbar-thumb:hover {{ background: {COLORS["secondary"]}; }}

        .footer-note {{
            text-align: center;
            font-size: 11px;
            color: {COLORS["text_muted"]};
            margin-top: 28px;
            padding-top: 14px;
            border-top: 1px solid {COLORS["border"]};
        }}

        /* ==================== RESPONSIVIDADE MOBILE ==================== */
        @media only screen and (max-width: 768px) {{
            div[data-testid="column"] {{
                width: 100% !important;
                flex: 1 1 100% !important;
                min-width: 100% !important;
            }}
            .main-header {{ padding: 12px 16px !important; }}
            .main-header h3 {{ font-size: 16px !important; }}
            .main-header p {{ font-size: 11px !important; }}
            .kpi-card {{ padding: 12px !important; }}
            .kpi-value {{ font-size: 19px !important; }}
            div[data-baseweb="tab-list"] {{
                display: flex !important;
                overflow-x: auto !important;
                white-space: nowrap !important;
                padding-bottom: 5px !important;
            }}
            button[data-baseweb="tab"] {{
                font-size: 11px !important;
                padding: 5px 10px !important;
            }}
            div[data-testid="stDataFrame"] {{ overflow-x: auto !important; }}
        }}

        /* Tela de login (acesso restrito) */
        .login-wrapper {{
            display: flex;
            justify-content: center;
            margin-top: 8vh;
        }}
        .login-card {{
            background: linear-gradient(135deg, {COLORS["surface"]} 0%, {COLORS["surface_alt"]} 100%);
            border: 1px solid {COLORS["border"]};
            border-top: 3px solid {COLORS["primary"]};
            border-radius: 14px;
            padding: 34px 40px;
            text-align: center;
            box-shadow: 0 10px 40px rgba(0,0,0,0.35);
            max-width: 380px;
        }}
        .login-card .login-icon {{
            font-size: 30px;
            margin-bottom: 6px;
        }}
        .login-card .login-logo {{
            width: 84px;
            height: 84px;
            object-fit: contain;
            background: #FFFFFF;
            border-radius: 16px;
            padding: 10px;
            margin: 0 auto 14px auto;
            display: block;
            box-shadow: 0 4px 16px rgba(0,0,0,0.25);
        }}
        .login-card h2 {{
            margin: 4px 0 2px 0 !important;
            font-size: 19px !important;
            color: {COLORS["text"]} !important;
            font-weight: 700;
        }}
        .login-card p {{
            margin: 0 !important;
            font-size: 12.5px !important;
            color: {COLORS["text_muted"]} !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================================
# 3.1 CONTROLE DE ACESSO (login com e-mail / senha + perfis)
# ============================================================================
LOGO_BEEA_B64 = "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"

CHAVE_LOCALSTORAGE_LOGIN = "beea_login_v1"


def obter_usuarios_cadastrados():
    """Le a lista de usuarios dos Secrets, aceitando tres formas de configuracao
    (todas podem coexistir e sao somadas):

    1) Usuario "de topo" (campos soltos no inicio do arquivo de Secrets):
        email = "controladoria@grupobeea.com.br"
        senha = "Richards23*"
        papel = "admin"

    2) Lista de tabelas:
        [[users]]
        email = "..."
        senha = "..."
        papel = "admin"

    3) Sub-tabelas [usuarios.<qualquer_nome>], aceitando tanto a chave
       'papel' quanto 'perfil':
        [usuarios.coordenador_financeiro]
        email = "coordenador.financeiro@grupobeea.com.br"
        senha = "Montoia6037"
        perfil = "visualizacao"

    A chave do dicionario retornado e sempre o e-mail em minusculas. Se
    nenhum usuario for encontrado em nenhum desses formatos, retorna
    dicionario vazio (painel fica aberto, sem bloqueio de login -- util
    apenas para testes locais sem Secrets configurados)."""
    usuarios = {}

    def _add(email, senha, perfil):
        email = str(email or "").strip().lower()
        senha = str(senha or "")
        if not email or not senha:
            return
        usuarios[email] = {
            "email": email,
            "senha": senha,
            "perfil": str(perfil or "visualizacao").strip().lower(),
        }

    email_topo = st.secrets.get("email", None)
    senha_topo = st.secrets.get("senha", None)
    if email_topo and senha_topo:
        _add(email_topo, senha_topo, st.secrets.get("papel", st.secrets.get("perfil", "admin")))

    for u in st.secrets.get("users", []):
        _add(u.get("email", ""), u.get("senha", ""), u.get("papel", u.get("perfil", "visualizacao")))

    for _chave, dados in dict(st.secrets.get("usuarios", {})).items():
        _add(
            dados.get("email", ""),
            dados.get("senha", ""),
            dados.get("papel", dados.get("perfil", "visualizacao")),
        )

    return usuarios


def _salvar_credenciais_no_navegador(email, senha):
    """Grava e-mail/senha (ofuscados em base64 -- isso NAO e criptografia)
    no localStorage do navegador, para preencher o login sozinho da
    proxima vez que a pessoa abrir o painel neste mesmo navegador."""
    email_b64 = base64.b64encode(email.encode("utf-8")).decode("ascii")
    senha_b64 = base64.b64encode(senha.encode("utf-8")).decode("ascii")
    components.html(
        f"""
        <script>
        try {{
            const dados = {{ e: "{email_b64}", s: "{senha_b64}" }};
            window.top.localStorage.setItem('{CHAVE_LOCALSTORAGE_LOGIN}', JSON.stringify(dados));
        }} catch (e) {{}}
        </script>
        """,
        height=0,
        width=0,
    )


def _esquecer_credenciais_no_navegador():
    """Apaga o e-mail/senha salvos no navegador (usado ao clicar em Sair,
    ou quando as credenciais salvas nao sao mais validas)."""
    components.html(
        f"""
        <script>
        try {{ window.top.localStorage.removeItem('{CHAVE_LOCALSTORAGE_LOGIN}'); }} catch (e) {{}}
        </script>
        """,
        height=0,
        width=0,
    )


def _tentar_autologin_via_url():
    """Se a URL trouxer credenciais (?le=&ls=) vindas do localStorage, tenta
    logar automaticamente com elas. Retorna True se conseguiu logar."""
    le = st.query_params.get("le")
    ls = st.query_params.get("ls")
    st.query_params.clear()
    if not le or not ls:
        return False

    try:
        email_salvo = base64.b64decode(str(le)).decode("utf-8").strip().lower()
        senha_salva = base64.b64decode(str(ls)).decode("utf-8")
    except Exception:
        return False

    usuarios = obter_usuarios_cadastrados()
    usuario = usuarios.get(email_salvo)
    if usuario and senha_salva == usuario["senha"]:
        st.session_state["usuario_logado"] = {"email": usuario["email"], "perfil": usuario["perfil"]}
        return True

    # Credenciais salvas nao sao mais validas (ex.: senha trocada) -- apaga.
    _esquecer_credenciais_no_navegador()
    return False


def _pedir_autofill_via_localstorage():
    """Roda no maximo uma vez por sessao: verifica (via JS) se ha
    credenciais salvas no navegador e, se houver, recarrega a pagina com
    elas na URL para tentarmos o autologin."""
    components.html(
        f"""
        <script>
        try {{
            const salvo = window.top.localStorage.getItem('{CHAVE_LOCALSTORAGE_LOGIN}');
            if (salvo) {{
                const dados = JSON.parse(salvo);
                if (dados.e && dados.s) {{
                    const url = new URL(window.top.location.href);
                    if (!url.searchParams.get('le')) {{
                        url.searchParams.set('le', dados.e);
                        url.searchParams.set('ls', dados.s);
                        window.top.location.replace(url.toString());
                    }}
                }}
            }}
        }} catch (e) {{}}
        </script>
        """,
        height=0,
        width=0,
    )


def checar_login():
    """Exibe uma tela de login (e-mail + senha) e retorna True somente apos
    autenticacao valida. Guarda o usuario logado em
    st.session_state['usuario_logado']. Se a pessoa marcar "Lembrar de
    mim", salva e-mail/senha no navegador para nao precisar digitar de
    novo da proxima vez. Se nao houver usuarios configurados nos Secrets,
    o painel fica aberto (sem bloqueio)."""

    usuarios = obter_usuarios_cadastrados()
    if not usuarios:
        return True

    if st.session_state.get("usuario_logado"):
        return True

    if _tentar_autologin_via_url():
        st.rerun()

    if not st.session_state.get("_autofill_login_tentado"):
        st.session_state["_autofill_login_tentado"] = True
        _pedir_autofill_via_localstorage()

    def validar_login():
        email_digitado = st.session_state.get("campo_email", "").strip().lower()
        senha_digitada = st.session_state.get("campo_senha", "")
        lembrar = st.session_state.get("campo_lembrar", True)
        usuario = usuarios.get(email_digitado)
        if usuario and str(senha_digitada) == usuario["senha"]:
            st.session_state["usuario_logado"] = {
                "email": usuario["email"],
                "perfil": usuario["perfil"],
            }
            st.session_state["login_invalido"] = False
            if lembrar:
                st.session_state["_credenciais_para_salvar"] = (usuario["email"], senha_digitada)
            else:
                st.session_state["_esquecer_credenciais"] = True
        else:
            st.session_state["usuario_logado"] = None
            st.session_state["login_invalido"] = True

    _, col_centro, _ = st.columns([1, 1.1, 1])
    with col_centro:
        st.markdown(
            f"""
            <div class="login-wrapper">
                <div class="login-card">
                    <img class="login-logo" src="data:image/jpeg;base64,{LOGO_BEEA_B64}" alt="Grupo Beea" />
                    <h2>Controladoria B&amp;A</h2>
                    <p>Acesso restrito — Painel Financeiro 2026</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.text_input("E-mail", key="campo_email", placeholder="seu.email@grupobeea.com.br")
        st.text_input(
            "Senha",
            type="password",
            key="campo_senha",
            on_change=validar_login,
            placeholder="Digite sua senha",
        )
        st.checkbox("Lembrar de mim neste navegador", value=True, key="campo_lembrar")
        if st.button("Entrar", use_container_width=True):
            validar_login()
            st.rerun()
        if st.session_state.get("login_invalido", False):
            st.error("E-mail ou senha incorretos. Tente novamente.")

    return False


if not checar_login():
    st.stop()

# Apos um login bem-sucedido nesta mesma execucao: salva ou apaga as
# credenciais no navegador, conforme a caixa "Lembrar de mim".
_creds_pendentes = st.session_state.pop("_credenciais_para_salvar", None)
if _creds_pendentes:
    _salvar_credenciais_no_navegador(*_creds_pendentes)
if st.session_state.pop("_esquecer_credenciais", False):
    _esquecer_credenciais_no_navegador()

usuario_atual = st.session_state.get("usuario_logado") or {"email": "", "perfil": "admin"}
eh_admin = usuario_atual["perfil"] == "admin"

# ============================================================================
# 4. CARREGAMENTO DE DADOS (Google Sheets com fallback local em rede)
# ============================================================================
@st.cache_resource
def obter_caminhos_excel():
    url_orc = "https://docs.google.com/spreadsheets/d/1x68Eg_6LlSKeFJEGmfhyBfcGgheSrVsl/export?format=xlsx"
    url_real = "https://docs.google.com/spreadsheets/d/12I0vGpYU_KNhGxAHOMHWAQu3Xkz_EsUZ/export?format=xlsx"

    caminho_base = r"G:\Meu Drive\Grupo B&A\Escritorio\Financeiro\COORDENAÇÃO FINANCEIRA\ORÇAMENTO\ORÇAMENTO 2026\CONTROLADORIA"
    path_orc_local = os.path.join(caminho_base, "ORCAMENTO 2026 - REV.1.xlsx")
    path_real_local = os.path.join(caminho_base, "REALIZADO 2026.xlsx")

    try:
        xls_orc = pd.ExcelFile(url_orc)
        path_orc = url_orc
        path_real = url_real
    except Exception:
        path_orc = path_orc_local if os.path.exists(path_orc_local) else "ORCAMENTO 2026 - REV.1.xlsx"
        path_real = path_real_local if os.path.exists(path_real_local) else "REALIZADO 2026.xlsx"
        xls_orc = pd.ExcelFile(path_orc)

    abas_ignorar = ["Sint Ebt loja", "CONS 25X26 V.1", "CONS 25X26 V.2"]
    abas_validas = [sheet for sheet in xls_orc.sheet_names if sheet not in abas_ignorar]

    return abas_validas, path_orc, path_real


try:
    with st.spinner("Conectando às planilhas financeiras..."):
        abas_disponiveis, path_orc, path_real = obter_caminhos_excel()
except Exception as e:
    st.error(f"Erro ao carregar as planilhas: {e}")
    st.stop()


@st.cache_data(ttl=60)
def carregar_dados_abas(path_o, path_r, lista_abas):
    dfs_o = []
    dfs_r = []
    for aba in lista_abas:
        try:
            df_o = pd.read_excel(path_o, sheet_name=aba)
            df_r = pd.read_excel(path_r, sheet_name=aba)
            dfs_o.append(df_o)
            dfs_r.append(df_r)
        except Exception:
            continue
    return dfs_o, dfs_r


@st.cache_data(ttl=60)
def carregar_dados_por_loja(path_o, path_r, lista_lojas):
    """Carrega os dados de Orçado/Realizado de cada loja SEPARADAMENTE (uma aba
    por loja), para permitir a divisão por loja no relatório Excel — independente
    do modo de visão escolhido na barra lateral (Consolidado ou Unidades)."""
    dados_por_loja = {}
    for loja in lista_lojas:
        try:
            df_o = pd.read_excel(path_o, sheet_name=loja)
            df_r = pd.read_excel(path_r, sheet_name=loja)
            dados_por_loja[loja] = (df_o, df_r)
        except Exception:
            continue
    return dados_por_loja


# ---------------------------------------------------------------------------
# 4.1 PLANO DE CONTAS — fonte auxiliar para a aba "Plano de Contas" do relatório
# ---------------------------------------------------------------------------
@st.cache_data(ttl=300)
def carregar_tabela_contas(path_r):
    """Lê a aba Tabela_Contas da planilha Realizado 2026: relação entre
    Natureza Financeira (plano de contas) e Linha DRE."""
    try:
        df = pd.read_excel(path_r, sheet_name="Tabela_Contas")
    except Exception:
        return pd.DataFrame(columns=["Natureza Financeira", "Linha DRE"])
    df.columns = [str(c).strip() for c in df.columns]
    return df


def montar_mapa_planos_por_dre(df_tabela_contas):
    """A partir da Tabela_Contas, retorna {linha_dre: [planos_de_conta...]}."""
    mapa = {}
    if df_tabela_contas.empty or "Linha DRE" not in df_tabela_contas.columns or "Natureza Financeira" not in df_tabela_contas.columns:
        return mapa
    for linha_dre, grupo in df_tabela_contas.groupby("Linha DRE"):
        planos = grupo["Natureza Financeira"].dropna().astype(str).str.strip().unique().tolist()
        mapa[str(linha_dre).strip()] = planos
    return mapa


# ============================================================================
# 5. BARRA LATERAL — FILTROS
# ============================================================================
st.sidebar.markdown(
    """
    <div class="sidebar-brand">
        <div class="dot"></div>
        <div>
            <span class="title">Controladoria B&A</span>
            <span class="subtitle">Painel Financeiro 2026</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown("**🔎 Escopo da Análise**")
modo_visao = st.sidebar.radio(
    "Modo de Visão:",
    ["Visão Consolidada", "Selecionar Unidade"],
)

abas_consolidadas_permitidas = [
    "DRE CONSOLIDADO",
    "ABPR CONSOLIDADO",
    "VD CONSOLIDADO",
    "LJ CONSOLIDADO",
    "ABPR + VD",
    "LJ - G&A",
]

opcoes_consolidadas = [a for a in abas_disponiveis if a in abas_consolidadas_permitidas]
if not opcoes_consolidadas:
    opcoes_consolidadas = abas_disponiveis

opcoes_unidades = [a for a in abas_disponiveis if a not in abas_consolidadas_permitidas]
if not opcoes_unidades:
    opcoes_unidades = abas_disponiveis

if modo_visao == "Visão Consolidada":
    visao_sel = st.sidebar.selectbox("Selecione a Visão:", opcoes_consolidadas)
    abas_para_carregar = [visao_sel]
    label_visao = visao_sel
else:
    lojas_sel = st.sidebar.multiselect(
        "Selecione as Lojas/Unidades:",
        options=opcoes_unidades,
        default=opcoes_unidades[:3] if len(opcoes_unidades) >= 3 else opcoes_unidades,
    )
    if not lojas_sel:
        st.warning("Selecione ao menos uma unidade.")
        st.stop()
    abas_para_carregar = lojas_sel
    label_visao = f"Soma de {len(lojas_sel)} Unidades"

with st.spinner("Carregando dados das abas selecionadas..."):
    list_df_orc, list_df_real = carregar_dados_abas(path_orc, path_real, abas_para_carregar)

if not list_df_real or not list_df_orc:
    st.error("Erro ao ler dados das abas.")
    st.stop()

# 5.1 FILTRO DE PERÍODO / MESES
st.sidebar.markdown("---")
st.sidebar.markdown("**📅 Período**")

nomes_meses = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
]
meses_cols = [
    "01/2026", "02/2026", "03/2026", "04/2026", "05/2026", "06/2026",
    "07/2026", "08/2026", "09/2026", "10/2026", "11/2026", "12/2026",
]

df_ref = list_df_real[0]
colunas_validas = [m for m in meses_cols if m in df_ref.columns]
m_map = {
    m_nome: m_col
    for m_nome, m_col in zip(nomes_meses, meses_cols)
    if m_col in colunas_validas
}

tipo_periodo = st.sidebar.radio(
    "Modo do Período:",
    ["Mês Selecionado", "Múltiplos Meses", "ANO COMPLETO (2026)"],
)

if tipo_periodo == "ANO COMPLETO (2026)":
    cols_kpi = list(m_map.values())
    cols_graficos = list(m_map.values())
    label_periodo_kpi = "ANO COMPLETO (2026)"
    label_periodo_graf = "ANO COMPLETO (2026)"
elif tipo_periodo == "Mês Selecionado":
    idx_default = min(6, len(m_map) - 1)
    mes_ref = st.sidebar.selectbox("Mês Desejado:", list(m_map.keys()), index=idx_default)
    idx = list(m_map.keys()).index(mes_ref)

    cols_kpi = list(m_map.values())[: idx + 1]
    cols_graficos = [m_map[mes_ref]]

    label_periodo_kpi = f"Acumulado YTD até {mes_ref}"
    label_periodo_graf = f"Mês de {mes_ref}"
else:
    meses_mult = st.sidebar.multiselect(
        "Selecione os Meses:",
        list(m_map.keys()),
        default=list(m_map.keys())[: min(7, len(m_map))],
    )
    cols_kpi = [m_map[m] for m in meses_mult if m in m_map]
    cols_graficos = cols_kpi
    label_periodo_kpi = "Meses Selecionados"
    label_periodo_graf = "Meses Selecionados"

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Atualizar Dados", use_container_width=True):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

st.sidebar.caption(f"Última atualização: {datetime.now(FUSO_BR).strftime('%d/%m/%Y às %H:%M')}")

st.sidebar.markdown("---")
perfil_label = "Administrador" if eh_admin else "Visualização"
st.sidebar.caption(f"👤 {usuario_atual['email']}  ·  Perfil: **{perfil_label}**")
if st.sidebar.button("🚪 Sair", use_container_width=True):
    st.session_state["usuario_logado"] = None
    _esquecer_credenciais_no_navegador()
    st.rerun()


# ============================================================================
# 6. FUNÇÕES DE SUPORTE (cálculo e formatação)
# ============================================================================
def get_valor_consolidado_multi(list_dfs, termo_conta, colunas, exato_linha_sintetica=False):
    total = 0.0
    if not colunas:
        return total

    for df in list_dfs:
        if df.empty:
            continue
        col_nome = "Nome" if "Nome" in df.columns else df.columns[0]

        if exato_linha_sintetica:
            termo_limpo = str(termo_conta).strip().lower()
            mask = df[col_nome].astype(str).str.strip().str.lower() == termo_limpo
        else:
            termo_escapado = re.escape(str(termo_conta))
            mask = df[col_nome].astype(str).str.contains(termo_escapado, case=False, na=False, regex=True)

        sub_df = df[mask]
        if not sub_df.empty:
            cols_existentes = [c for c in colunas if c in sub_df.columns]
            dados_num = sub_df[cols_existentes].apply(pd.to_numeric, errors="coerce").fillna(0)
            total += dados_num.sum().sum()
    return total


def formata_brl(val):
    if pd.isna(val) or val == 0:
        return "R$ 0,00"
    return f"R$ {val:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def formata_m(val):
    if pd.isna(val) or val == 0:
        return "R$ 0M"
    return f"R$ {val / 1e6:,.1f}M".replace(".", ",")


def eh_grupo_sintetico(nome_linha):
    return bool(re.match(r"^\d+\s*-\s*", str(nome_linha).strip()))


def cor_valor(val):
    if pd.isna(val):
        return ""
    color = COLORS["positive"] if val >= 0 else COLORS["negative"]
    return f"color: {color}; font-weight: 500;"


# ============================================================================
# 6.1 GERAÇÃO DE RELATÓRIO EXCEL (formatado, para a aba de Emissão)
# ============================================================================
_THIN = Side(style="thin", color="FFD9DDE3")

EXCEL_STYLE = {
    "fill_title": PatternFill(fill_type="solid", start_color="FF0B0E14", end_color="FF0B0E14"),
    "fill_header": PatternFill(fill_type="solid", start_color="FF1A1F2E", end_color="FF1A1F2E"),
    "fill_zebra": PatternFill(fill_type="solid", start_color="FFF3F5F9", end_color="FFF3F5F9"),
    "fill_total": PatternFill(fill_type="solid", start_color="FFDCE8FF", end_color="FFDCE8FF"),
    "font_title": Font(color="FFFFFFFF", bold=True, size=13, name="Calibri"),
    "font_header": Font(color="FFFFFFFF", bold=True, size=11, name="Calibri"),
    "font_normal": Font(color="FF1F2937", size=10.5, name="Calibri"),
    "font_bold": Font(color="FF1F2937", bold=True, size=10.5, name="Calibri"),
    "font_pos": Font(color="FF1B8A5A", size=10.5, name="Calibri"),
    "font_neg": Font(color="FFC0392B", size=10.5, name="Calibri"),
    "font_caption": Font(color="FF6B7280", italic=True, size=9.5, name="Calibri"),
    "border": Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN),
}


def _fonte_por_valor(valor):
    return EXCEL_STYLE["font_pos"] if valor >= 0 else EXCEL_STYLE["font_neg"]


def _escrever_titulo(ws, texto, linha, n_colunas):
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=n_colunas)
    cell = ws.cell(row=linha, column=1, value=texto)
    cell.font = EXCEL_STYLE["font_title"]
    cell.fill = EXCEL_STYLE["fill_title"]
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[linha].height = 26


def _escrever_legenda(ws, texto, linha, n_colunas):
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=n_colunas)
    cell = ws.cell(row=linha, column=1, value=texto)
    cell.font = EXCEL_STYLE["font_caption"]
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[linha].height = 16


def _escrever_cabecalho_matriz(ws, linha, primeiro_rotulo, mapa_meses, fill_header=None):
    """Escreve um cabeçalho no formato: [primeiro_rotulo, Jan, Fev, ..., Total Ano]."""
    headers = [primeiro_rotulo] + [m.capitalize() for m in mapa_meses.keys()] + ["Total Ano"]
    fill = fill_header or EXCEL_STYLE["fill_zebra"]
    for col, texto in enumerate(headers, start=1):
        cell = ws.cell(row=linha, column=col, value=texto)
        cell.font = EXCEL_STYLE["font_bold"]
        cell.fill = fill
        cell.border = EXCEL_STYLE["border"]
        cell.alignment = Alignment(horizontal="left" if col == 1 else "center")
    return len(headers)


def _escrever_linha_matriz(ws, linha, rotulo, valores_mes, total, negrito=False, fill=None, colorir_por_sinal=False):
    """Escreve uma linha: [rotulo, valor_mes_1, valor_mes_2, ..., total]."""
    valores = [rotulo] + list(valores_mes) + [total]
    for col, val in enumerate(valores, start=1):
        cell = ws.cell(row=linha, column=col, value=val)
        cell.border = EXCEL_STYLE["border"]
        if fill:
            cell.fill = fill
        if col == 1:
            cell.font = EXCEL_STYLE["font_bold"] if negrito else EXCEL_STYLE["font_normal"]
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        else:
            valor_numerico = val if isinstance(val, (int, float)) else 0
            if colorir_por_sinal:
                cell.font = _fonte_por_valor(valor_numerico)
            else:
                cell.font = EXCEL_STYLE["font_bold"] if negrito else EXCEL_STYLE["font_normal"]
            cell.number_format = '"R$" #,##0.00'
            cell.alignment = Alignment(horizontal="right")


def montar_relatorio_excel(
    contas_sel,
    dfs_real,
    dfs_orc,
    mapa_meses,
    colunas_ano,
    escopo_label,
    dados_por_loja=None,
    mapa_planos_dre=None,
):
    """Gera um relatório Excel formatado com três planilhas:
    - Resumo: total do ano por conta, CONSOLIDADO (não muda com a divisão por loja).
    - Detalhe Mensal: contas da DRE nas linhas, meses nas colunas, dividido por loja.
    - Plano de Contas: composição (planos de contas) de cada linha da DRE, por loja
      (valores puxados diretamente da planilha Realizado 2026 de cada loja).
    """
    dados_por_loja = dados_por_loja or {}
    mapa_planos_dre = mapa_planos_dre or {}
    wb = Workbook()
    gerado_em = f"{escopo_label} · Gerado em {datetime.now(FUSO_BR).strftime('%d/%m/%Y às %H:%M')}"

    # ---------------- ABA "RESUMO" ----------------
    ws1 = wb.active
    ws1.title = "Resumo"

    _escrever_titulo(ws1, "Relatório de DRE — Orçado vs. Realizado (Resumo Anual)", 1, 5)
    _escrever_legenda(ws1, gerado_em, 2, 5)

    linha_header = 4
    headers = ["Conta / Linha DRE", "Realizado (R$)", "Orçado (R$)", "Desvio (R$)", "Desvio (%)"]
    for col, texto in enumerate(headers, start=1):
        cell = ws1.cell(row=linha_header, column=col, value=texto)
        cell.font = EXCEL_STYLE["font_header"]
        cell.fill = EXCEL_STYLE["fill_header"]
        cell.border = EXCEL_STYLE["border"]
        cell.alignment = Alignment(horizontal="left" if col == 1 else "center", vertical="center")
    ws1.row_dimensions[linha_header].height = 20
    ws1.freeze_panes = f"A{linha_header + 1}"

    linha = linha_header + 1
    for i, conta in enumerate(contas_sel):
        v_real = get_valor_consolidado_multi(dfs_real, conta, colunas_ano)
        v_orc = get_valor_consolidado_multi(dfs_orc, conta, colunas_ano)
        desvio = v_real - v_orc
        desvio_pct = (desvio / abs(v_orc) * 100) if v_orc != 0 else 0.0

        fill = EXCEL_STYLE["fill_zebra"] if i % 2 == 1 else None
        valores = [conta, v_real, v_orc, desvio, desvio_pct]
        for col, val in enumerate(valores, start=1):
            cell = ws1.cell(row=linha, column=col, value=val)
            cell.border = EXCEL_STYLE["border"]
            if fill:
                cell.fill = fill
            if col == 1:
                cell.font = EXCEL_STYLE["font_normal"]
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col in (2, 3):
                cell.font = EXCEL_STYLE["font_normal"]
                cell.number_format = '"R$" #,##0.00'
                cell.alignment = Alignment(horizontal="right")
            elif col == 4:
                cell.font = _fonte_por_valor(desvio)
                cell.number_format = '"R$" #,##0.00'
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.font = _fonte_por_valor(desvio)
                cell.number_format = '0.0"%"'
                cell.alignment = Alignment(horizontal="right")
        linha += 1

    for col, largura in zip(range(1, 6), [46, 18, 18, 18, 14]):
        ws1.column_dimensions[get_column_letter(col)].width = largura

    # ---------------- ABA "DETALHE MENSAL" (contas nas linhas, meses nas colunas, por loja) ----------------
    ws2 = wb.create_sheet("Detalhe Mensal")
    n_col_matriz = 1 + len(mapa_meses) + 1
    _escrever_titulo(ws2, "Detalhe Mensal — Orçado vs. Realizado, por Loja", 1, n_col_matriz)
    _escrever_legenda(ws2, gerado_em, 2, n_col_matriz)

    linha = 4
    lojas_ordenadas = sorted(dados_por_loja.keys()) if dados_por_loja else []

    if not lojas_ordenadas:
        ws2.cell(row=linha, column=1, value="Nenhuma loja/unidade disponível para divisão.").font = EXCEL_STYLE["font_normal"]
    for loja in lojas_ordenadas:
        df_o_loja, df_r_loja = dados_por_loja[loja]

        ws2.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=n_col_matriz)
        cell_loja = ws2.cell(row=linha, column=1, value=f"🏬 Loja: {loja}")
        cell_loja.font = EXCEL_STYLE["font_title"]
        cell_loja.fill = EXCEL_STYLE["fill_title"]
        cell_loja.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws2.row_dimensions[linha].height = 22
        linha += 1

        _escrever_cabecalho_matriz(ws2, linha, "Conta / Tipo", mapa_meses)
        linha += 1

        for conta in contas_sel:
            valores_real = [get_valor_consolidado_multi([df_r_loja], conta, [m_col]) for m_col in mapa_meses.values()]
            valores_orc = [get_valor_consolidado_multi([df_o_loja], conta, [m_col]) for m_col in mapa_meses.values()]
            valores_desvio = [vr - vo for vr, vo in zip(valores_real, valores_orc)]

            _escrever_linha_matriz(ws2, linha, f"{conta} — Realizado", valores_real, sum(valores_real))
            linha += 1
            _escrever_linha_matriz(ws2, linha, f"{conta} — Orçado", valores_orc, sum(valores_orc), fill=EXCEL_STYLE["fill_zebra"])
            linha += 1
            _escrever_linha_matriz(ws2, linha, f"{conta} — Desvio", valores_desvio, sum(valores_desvio), colorir_por_sinal=True)
            linha += 1

        linha += 1  # espaço entre lojas

    largura_mes = 14
    ws2.column_dimensions["A"].width = 40
    for col in range(2, n_col_matriz + 1):
        ws2.column_dimensions[get_column_letter(col)].width = largura_mes
    ws2.freeze_panes = "B4"

    # ---------------- ABA "PLANO DE CONTAS" (composição de cada linha da DRE, por loja) ----------------
    ws3 = wb.create_sheet("Plano de Contas")
    _escrever_titulo(ws3, "Plano de Contas — Composição das Linhas da DRE, por Loja", 1, n_col_matriz)
    _escrever_legenda(
        ws3,
        f"{gerado_em}  ·  Valores de cada plano de contas puxados diretamente da planilha Realizado 2026 de cada loja.",
        2, n_col_matriz,
    )

    linha = 4
    if not lojas_ordenadas:
        ws3.cell(row=linha, column=1, value="Nenhuma loja/unidade disponível para divisão.").font = EXCEL_STYLE["font_normal"]

    for loja in lojas_ordenadas:
        df_o_loja, df_r_loja = dados_por_loja[loja]

        ws3.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=n_col_matriz)
        cell_loja = ws3.cell(row=linha, column=1, value=f"🏬 Loja: {loja}")
        cell_loja.font = EXCEL_STYLE["font_title"]
        cell_loja.fill = EXCEL_STYLE["fill_title"]
        cell_loja.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws3.row_dimensions[linha].height = 22
        linha += 1

        for conta in contas_sel:
            ws3.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=n_col_matriz)
            cell_conta = ws3.cell(row=linha, column=1, value=f"Conta: {conta}")
            cell_conta.font = EXCEL_STYLE["font_header"]
            cell_conta.fill = EXCEL_STYLE["fill_header"]
            cell_conta.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws3.row_dimensions[linha].height = 20
            linha += 1

            _escrever_cabecalho_matriz(ws3, linha, "Plano de Contas", mapa_meses)
            linha += 1

            planos = mapa_planos_dre.get(str(conta).strip(), []) or [conta]

            soma_planos_mes = [0.0] * len(mapa_meses)
            for i_plano, plano in enumerate(planos):
                valores_plano = [
                    get_valor_consolidado_multi([df_r_loja], plano, [m_col]) for m_col in mapa_meses.values()
                ]
                soma_planos_mes = [s + v for s, v in zip(soma_planos_mes, valores_plano)]
                fill_plano = EXCEL_STYLE["fill_zebra"] if i_plano % 2 == 1 else None
                _escrever_linha_matriz(ws3, linha, plano, valores_plano, sum(valores_plano), fill=fill_plano)
                linha += 1

            _escrever_linha_matriz(
                ws3, linha, f"TOTAL — {conta}", soma_planos_mes, sum(soma_planos_mes),
                negrito=True, fill=EXCEL_STYLE["fill_total"],
            )
            linha += 2  # espaço entre contas

        linha += 1  # espaço entre lojas

    ws3.column_dimensions["A"].width = 46
    for col in range(2, n_col_matriz + 1):
        ws3.column_dimensions[get_column_letter(col)].width = largura_mes
    ws3.freeze_panes = "B4"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================================
# 7. CABEÇALHO PRINCIPAL
# ============================================================================
st.markdown(
    f"""
    <div class="main-header">
        <h3>PAINEL ANALÍTICO DE PERFORMANCE ESTRATÉGICA</h3>
        <p>
            <span class="chip">{label_visao}</span>
            &nbsp;·&nbsp; Período: <b>{label_periodo_kpi}</b>
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================================
# 8. ABAS
# ============================================================================
_nomes_abas = [
    "📊 Visão Geral & Charts",
    "📋 DRE Completa & Desvios",
    "📅 Histórico Mensal",
    "🔮 Previsões & Trends",
    "📤 Emitir Relatório",
]
if eh_admin:
    _nomes_abas.append("👥 Usuários")
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(_nomes_abas)
else:
    tab1, tab2, tab3, tab4, tab5 = st.tabs(_nomes_abas)

# ---------------------------------------------------------------------------
# ABA 1: VISÃO GERAL & CHARTS
# ---------------------------------------------------------------------------
with tab1:
    # ---- KPIs executivos da visão geral ----
    rec_liq_real_kpi = get_valor_consolidado_multi(list_df_real, "3 - Receita Operacional Liquida", cols_kpi)
    rec_liq_orc_kpi = get_valor_consolidado_multi(list_df_orc, "3 - Receita Operacional Liquida", cols_kpi)

    ebitda_real_kpi = get_valor_consolidado_multi(list_df_real, "11 - EBITDA", cols_kpi)
    ebitda_orc_kpi = get_valor_consolidado_multi(list_df_orc, "11 - EBITDA", cols_kpi)

    margem_ebitda_kpi = (ebitda_real_kpi / rec_liq_real_kpi * 100) if rec_liq_real_kpi != 0 else 0

    diff_ebitda_kpi = ebitda_real_kpi - ebitda_orc_kpi
    pct_ebitda_kpi = (diff_ebitda_kpi / abs(ebitda_orc_kpi)) * 100 if ebitda_orc_kpi != 0 else 0

    pct_vendas_prog = min(100.0, max(0.0, (rec_liq_real_kpi / rec_liq_orc_kpi * 100))) if rec_liq_orc_kpi > 0 else 0
    pct_lucro_prog = min(100.0, max(0.0, (ebitda_real_kpi / ebitda_orc_kpi * 100))) if ebitda_orc_kpi > 0 else 0

    cor_rec = cor_variacao(rec_liq_real_kpi)
    cor_ebitda = cor_variacao(ebitda_real_kpi)
    cor_diff_eb = cor_variacao(diff_ebitda_kpi)
    cor_mg_eb = cor_variacao(margem_ebitda_kpi)

    st.markdown(
        render_kpi_row([
            dict(label="RECEITA LÍQUIDA (YTD)", value=formata_brl(rec_liq_real_kpi), value_color=cor_rec,
                 subtext=f"Orçado: {formata_brl(rec_liq_orc_kpi)}", progress_pct=pct_vendas_prog, icon="💰"),
            dict(label="EBITDA (YTD)", value=formata_brl(ebitda_real_kpi), value_color=cor_ebitda,
                 subtext=f"Orçado: {formata_brl(ebitda_orc_kpi)}", progress_pct=pct_lucro_prog, icon="📈"),
            dict(label="VARIAÇÃO EBITDA", value=formata_brl(diff_ebitda_kpi), value_color=cor_diff_eb,
                 subtext=f"{pct_ebitda_kpi:+.1f}% vs Orçamento", subtext_color=cor_diff_eb, icon="⚖️"),
            dict(label="MARGEM EBITDA %", value=f"{margem_ebitda_kpi:.1f}%", value_color=cor_mg_eb,
                 subtext="Realizada no Período", icon="🎯"),
        ]),
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    st.caption(f"Visualização e Eficiência referente ao período: **{label_periodo_kpi}**")

    cg1, cg2 = st.columns(2)

    with cg1:
        st.markdown('<div class="section-title">Bridge de Performance (YTD)</div>', unsafe_allow_html=True)

        rec_bruta = get_valor_consolidado_multi(list_df_real, "1 - Receita Operacional Bruta", cols_kpi)
        deducoes = get_valor_consolidado_multi(list_df_real, "2 - Deduções da Receita Operacional Bruta", cols_kpi)
        rec_liq = get_valor_consolidado_multi(list_df_real, "3 - Receita Operacional Liquida", cols_kpi)

        cmv_bridge = get_valor_consolidado_multi(list_df_real, "4 - ", cols_kpi, exato_linha_sintetica=True)
        if cmv_bridge == 0:
            cmv_bridge = get_valor_consolidado_multi(list_df_real, "4 - Custo das Vendas", cols_kpi)

        margem_bruta = get_valor_consolidado_multi(list_df_real, "5 - Margem de Contribuição 1", cols_kpi)

        desp_var = get_valor_consolidado_multi(list_df_real, "6 - Despesas Variáveis", cols_kpi)
        if desp_var == 0:
            desp_var = get_valor_consolidado_multi(list_df_real, "Despesas Variáveis", cols_kpi)

        desp_op = get_valor_consolidado_multi(list_df_real, "8 - Despesas Operacionais", cols_kpi)
        sga_total = desp_var + desp_op

        deprec = get_valor_consolidado_multi(list_df_real, "13 - Depreciação e Amortização", cols_kpi)

        ebitda = get_valor_consolidado_multi(list_df_real, "11 - EBITDA", cols_kpi)
        ebit = margem_bruta - abs(sga_total) - abs(deprec)

        base_rec = rec_liq if rec_liq != 0 else 1.0

        p_rec_bruta = round((rec_bruta / base_rec) * 100)
        p_deducoes = round((deducoes / base_rec) * 100)
        p_rec_liq = 100
        p_cmv = round((-abs(cmv_bridge) / base_rec) * 100)
        p_mb = round((margem_bruta / base_rec) * 100)
        p_sga = round((-abs(sga_total) / base_rec) * 100)

        p_ebit = round((ebit / base_rec) * 100)
        p_deprec = round((-abs(deprec) / base_rec) * 100)
        p_ebitda = round((ebitda / base_rec) * 100)

        x_bridge = [
            "Receita Bruta", "Deduções", "Receita Líquida", "CMV",
            "Margem Bruta", "SG&A", "EBIT", "D&A", "EBITDA",
        ]
        measures = ["absolute", "relative", "total", "relative", "total", "relative", "total", "relative", "total"]
        y_bridge = [p_rec_bruta, p_deducoes, 0, p_cmv, 0, p_sga, 0, abs(p_deprec), 0]
        text_labels = [
            f"{p_rec_bruta}%", f"{p_deducoes}%", f"{p_rec_liq}%",
            f"{p_cmv}%", f"{p_mb}%", f"{p_sga}%",
            f"{p_ebit}%", f"{p_deprec}%", f"{p_ebitda}%",
        ]

        fig_waterfall = go.Figure(
            go.Waterfall(
                orientation="v",
                measure=measures,
                x=x_bridge,
                y=y_bridge,
                text=text_labels,
                textposition="outside",
                connector={"line": {"color": COLORS["border"], "width": 1}},
                decreasing={"marker": {"color": COLORS["muted_line"]}},
                increasing={"marker": {"color": COLORS["primary"]}},
                totals={"marker": {"color": COLORS["secondary"]}},
            )
        )
        estilo_grafico(
            fig_waterfall,
            height=400,
            xaxis=dict(tickangle=-45, gridcolor="rgba(0,0,0,0)", fixedrange=True),
            yaxis=dict(showticklabels=False, gridcolor="rgba(0,0,0,0)", fixedrange=True),
        )
        st.plotly_chart(fig_waterfall, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    with cg2:
        st.markdown('<div class="section-title">Real vs. Orçado (YTD)</div>', unsafe_allow_html=True)

        cats = ["CMV", "TRF/REM", "Margem Contrib. 2", "Despesas Fixas", "EBITDA"]

        NOME_LINHA_CMV_DETALHADO = "4.1 - Custo da Mercadoria Vendida - CMV"

        cmv_r = abs(get_valor_consolidado_multi(list_df_real, NOME_LINHA_CMV_DETALHADO, cols_kpi, exato_linha_sintetica=True))
        if cmv_r == 0:
            cmv_r = abs(get_valor_consolidado_multi(list_df_real, "4.1 - Custo da Mercadoria Vendida", cols_kpi))

        cmv_o = abs(get_valor_consolidado_multi(list_df_orc, NOME_LINHA_CMV_DETALHADO, cols_kpi, exato_linha_sintetica=True))
        if cmv_o == 0:
            cmv_o = abs(get_valor_consolidado_multi(list_df_orc, "4.1 - Custo da Mercadoria Vendida", cols_kpi))

        trf_r = abs(get_valor_consolidado_multi(list_df_real, "TRF/REM", cols_kpi))
        if trf_r == 0:
            trf_r = abs(get_valor_consolidado_multi(list_df_real, "TRF / REM", cols_kpi))

        trf_o = abs(get_valor_consolidado_multi(list_df_orc, "TRF/REM", cols_kpi))
        if trf_o == 0:
            trf_o = abs(get_valor_consolidado_multi(list_df_orc, "TRF / REM", cols_kpi))

        mc_r = get_valor_consolidado_multi(list_df_real, "Margem de Contribuição 2", cols_kpi)
        if mc_r == 0:
            mc_r = get_valor_consolidado_multi(list_df_real, "7 - Margem de Contribuição 2", cols_kpi)

        mc_o = get_valor_consolidado_multi(list_df_orc, "Margem de Contribuição 2", cols_kpi)
        if mc_o == 0:
            mc_o = get_valor_consolidado_multi(list_df_orc, "7 - Margem de Contribuição 2", cols_kpi)

        dfix_r = abs(get_valor_consolidado_multi(list_df_real, "8 - Despesas Operacionais", cols_kpi))
        dfix_o = abs(get_valor_consolidado_multi(list_df_orc, "8 - Despesas Operacionais", cols_kpi))

        eb_r = get_valor_consolidado_multi(list_df_real, "11 - EBITDA", cols_kpi)
        eb_o = get_valor_consolidado_multi(list_df_orc, "11 - EBITDA", cols_kpi)

        val_r = [cmv_r, trf_r, mc_r, dfix_r, eb_r]
        val_o = [cmv_o, trf_o, mc_o, dfix_o, eb_o]

        labels_r = [formata_m(v) for v in val_r]
        labels_o = [formata_m(v) for v in val_o]

        fig_bar = go.Figure(
            data=[
                go.Bar(name="Realizado (R$)", x=cats, y=val_r, text=labels_r, textposition="outside", marker_color=COLORS["primary"]),
                go.Bar(name="Orçado (R$)", x=cats, y=val_o, text=labels_o, textposition="outside", marker_color=COLORS["secondary"]),
            ]
        )
        estilo_grafico(
            fig_bar,
            height=400,
            barmode="group",
            xaxis=dict(gridcolor=COLORS["border"], zerolinecolor=COLORS["border"], fixedrange=True),
            yaxis=dict(showticklabels=False, gridcolor="rgba(0,0,0,0)", fixedrange=True),
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_bar, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    st.markdown("<br>", unsafe_allow_html=True)

    cg3, cg4 = st.columns([1.3, 0.7])

    with cg3:
        st.markdown('<div class="section-title">Evolução Mensal (Receita vs. Margem de Contribuição)</div>', unsafe_allow_html=True)

        rec_m, mc_m, rotulos_m = [], [], []
        for m_nome, c in m_map.items():
            v_rec = get_valor_consolidado_multi(list_df_real, "3 - Receita Operacional Liquida", [c])
            v_mc = get_valor_consolidado_multi(list_df_real, "7 - Margem de Contribuição 2", [c])
            if v_mc == 0:
                v_mc = get_valor_consolidado_multi(list_df_real, "Margem de Contribuição 2", [c])

            if v_rec != 0 or v_mc != 0:
                rec_m.append(v_rec)
                mc_m.append(v_mc)
                rotulos_m.append(m_nome.capitalize())

        labels_rec = [formata_m(v) for v in rec_m]
        labels_mc = [formata_m(v) for v in mc_m]

        fig_line = go.Figure()
        fig_line.add_trace(
            go.Scatter(
                x=rotulos_m, y=rec_m, mode="lines+markers+text", name="Receita (R$)",
                text=labels_rec, textposition="top center",
                line=dict(color=COLORS["primary"], width=2, shape="spline"),
                marker=dict(size=6, color=COLORS["surface"], line=dict(color=COLORS["primary"], width=2)),
                textfont=dict(color=COLORS["text"], size=11, family=FONT_STACK),
            )
        )
        fig_line.add_trace(
            go.Scatter(
                x=rotulos_m, y=mc_m, mode="lines+markers+text", name="Margem Contrib. 2 (R$)",
                text=labels_mc, textposition="bottom center",
                line=dict(color=COLORS["muted_line"], width=2, shape="spline"),
                marker=dict(size=6, color=COLORS["surface"], line=dict(color=COLORS["muted_line"], width=2)),
                textfont=dict(color=COLORS["text_muted"], size=11, family=FONT_STACK),
            )
        )
        estilo_grafico(
            fig_line,
            height=380,
            xaxis=dict(showgrid=False, zeroline=False, tickangle=-45, tickfont=dict(size=11, color=COLORS["text_muted"]), fixedrange=True),
            yaxis=dict(showgrid=False, showticklabels=False, zeroline=False, fixedrange=True),
            legend=dict(orientation="h", yanchor="top", y=-0.25, xanchor="center", x=0.5, font=dict(color=COLORS["text_muted"])),
        )
        st.plotly_chart(fig_line, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    with cg4:
        st.markdown('<div class="section-title">Composição dos Custos & Saídas</div>', unsafe_allow_html=True)

        cmv_real_kpi = abs(get_valor_consolidado_multi(list_df_real, "4 - ", cols_kpi, exato_linha_sintetica=True))
        if cmv_real_kpi == 0:
            cmv_real_kpi = abs(get_valor_consolidado_multi(list_df_real, "4 - Custo das Vendas", cols_kpi))

        desp_op_real = abs(get_valor_consolidado_multi(list_df_real, "8 - Despesas Operacionais", cols_kpi))

        v_cmv_pie = abs(cmv_real_kpi)
        v_desp_var_pie = abs(get_valor_consolidado_multi(list_df_real, "6 - Despesas Variáveis", cols_kpi))
        v_desp_op_pie = abs(desp_op_real)
        v_deprec_pie = abs(get_valor_consolidado_multi(list_df_real, "13 - Depreciação e Amortização", cols_kpi))

        total_pie = v_cmv_pie + v_desp_var_pie + v_desp_op_pie + v_deprec_pie

        fig_donut = go.Figure(
            data=[
                go.Pie(
                    labels=["CMV / Custo", "Despesas Var.", "Despesas Op. (OpEx)", "Depreciação/Amort."],
                    values=[v_cmv_pie, v_desp_var_pie, v_desp_op_pie, v_deprec_pie],
                    hole=0.62,
                    marker=dict(colors=[COLORS["primary"], COLORS["muted_line"], COLORS["secondary"], COLORS["border_soft"]],
                                line=dict(color=COLORS["surface"], width=2)),
                    textinfo="percent",
                    hoverinfo="label+value+percent",
                )
            ]
        )
        fig_donut.add_annotation(
            text=f"<b>{formata_m(total_pie)}</b><br><span style='font-size:10px;color:{COLORS['text_muted']}'>Total Saídas</span>",
            showarrow=False, font=dict(color=COLORS["text"], size=13, family=FONT_STACK),
        )
        estilo_grafico(
            fig_donut,
            height=380,
            legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_donut, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)


# ---------------------------------------------------------------------------
# ABA 2: DRE COMPLETA & DESVIOS
# ---------------------------------------------------------------------------
with tab2:
    st.markdown(f'<div class="section-title">📋 Análise de DRE e Desvios — {label_periodo_graf}</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    c1, _ = st.columns([2, 1])
    with c1:
        tipo_visao_dre = st.radio(
            "Filtro de Nível de Visão:",
            ["Apenas Grupos Principais (Sintética)", "Todas as Contas (Analítica)"],
            horizontal=True,
        )

    col_nome = "Nome" if "Nome" in df_ref.columns else df_ref.columns[0]
    linhas_dre = df_ref[col_nome].dropna().astype(str).unique()

    is_sintetica_dre = tipo_visao_dre == "Apenas Grupos Principais (Sintética)"
    if is_sintetica_dre:
        linhas_dre = [l for l in linhas_dre if eh_grupo_sintetico(l)]

    if not is_sintetica_dre:
        contas_filtradas_dre = st.multiselect(
            "🔍 Filtrar Contas Específicas (Estilo Excel):",
            options=linhas_dre,
            default=[],
            key="filtro_contas_dre",
        )
        if contas_filtradas_dre:
            linhas_dre = contas_filtradas_dre

    rec_bruta_real = get_valor_consolidado_multi(list_df_real, "1 - Receita Operacional Bruta", cols_graficos)
    rec_bruta_orc = get_valor_consolidado_multi(list_df_orc, "1 - Receita Operacional Bruta", cols_graficos)

    dados_dre = []
    for linha in linhas_dre:
        v_real = get_valor_consolidado_multi(list_df_real, linha, cols_graficos)
        v_orc = get_valor_consolidado_multi(list_df_orc, linha, cols_graficos)
        desvio_rs = v_real - v_orc

        av_real_pct = (v_real / rec_bruta_real * 100) if rec_bruta_real != 0 else 0.0
        av_orc_pct = (v_orc / rec_bruta_orc * 100) if rec_bruta_orc != 0 else 0.0
        ah_pct = (desvio_rs / abs(v_orc) * 100) if v_orc != 0 else 0.0

        dados_dre.append(
            {
                "Conta / Linha DRE": linha,
                "Realizado (R$)": v_real,
                "AV Real (%)": av_real_pct,
                "Orçado (R$)": v_orc,
                "AV Orçado (%)": av_orc_pct,
                "Desvio (R$)": desvio_rs,
                "AH (%)": ah_pct,
            }
        )

    df_dre_final = pd.DataFrame(dados_dre)

    # ---- KPIs contextuais de desvio ----
    if not df_dre_final.empty:
        n_favoravel = int((df_dre_final["Desvio (R$)"] > 0).sum())
        n_desfavoravel = int((df_dre_final["Desvio (R$)"] < 0).sum())
        idx_maior_desvio = df_dre_final["Desvio (R$)"].abs().idxmax()
        linha_maior_desvio = df_dre_final.loc[idx_maior_desvio, "Conta / Linha DRE"]
        valor_maior_desvio = df_dre_final.loc[idx_maior_desvio, "Desvio (R$)"]
        desvio_ebitda_dre = get_valor_consolidado_multi(list_df_real, "11 - EBITDA", cols_graficos) - \
            get_valor_consolidado_multi(list_df_orc, "11 - EBITDA", cols_graficos)

        st.markdown(
            render_kpi_row([
                dict(label="DESVIO EBITDA NO PERÍODO", value=formata_brl(desvio_ebitda_dre),
                     value_color=cor_variacao(desvio_ebitda_dre), subtext=label_periodo_graf, icon="📐"),
                dict(label="CONTAS COM DESVIO FAVORÁVEL", value=str(n_favoravel), value_color=COLORS["positive"],
                     subtext=f"de {len(df_dre_final)} linhas analisadas", icon="✅"),
                dict(label="CONTAS COM DESVIO DESFAVORÁVEL", value=str(n_desfavoravel), value_color=COLORS["negative"],
                     subtext=f"de {len(df_dre_final)} linhas analisadas", icon="⚠️"),
                dict(label="MAIOR DESVIO INDIVIDUAL", value=formata_brl(valor_maior_desvio),
                     value_color=cor_variacao(valor_maior_desvio), subtext=linha_maior_desvio[:38], icon="🔎"),
            ]),
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

    column_config_dre = {
        "Conta / Linha DRE": st.column_config.TextColumn("Conta / Linha DRE", width="large"),
    }

    ALTURA_17_LINHAS = 633
    cols_num_dre = ["Realizado (R$)", "AV Real (%)", "Orçado (R$)", "AV Orçado (%)", "Desvio (R$)", "AH (%)"]

    st.dataframe(
        df_dre_final.style.format(
            {
                "Realizado (R$)": formata_brl,
                "AV Real (%)": "{:.1f}%",
                "Orçado (R$)": formata_brl,
                "AV Orçado (%)": "{:.1f}%",
                "Desvio (R$)": formata_brl,
                "AH (%)": "{:.1f}%",
            }
        ).map(cor_valor, subset=cols_num_dre),
        column_config=column_config_dre,
        use_container_width=True,
        height=ALTURA_17_LINHAS,
        hide_index=True,
        on_select="rerun",
        selection_mode="multi-row",
    )


# ---------------------------------------------------------------------------
# ABA 3: HISTÓRICO MENSAL
# ---------------------------------------------------------------------------
with tab3:
    st.markdown('<div class="section-title">📅 Histórico Mensal Mês a Mês</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    ch1, ch2 = st.columns(2)
    with ch1:
        tipo_hist = st.radio("Base de Dados:", ["Realizado", "Orçado"], horizontal=True)
    with ch2:
        visao_hist_dre = st.radio("Detalhes:", ["Grupos Fechados (Sintético)", "Todas as Contas (Analítico)"], horizontal=True)

    linhas_hist = df_ref[col_nome].dropna().astype(str).unique()
    is_sintetica_hist = visao_hist_dre == "Grupos Fechados (Sintético)"
    if is_sintetica_hist:
        linhas_hist = [l for l in linhas_hist if eh_grupo_sintetico(l)]

    if not is_sintetica_hist:
        contas_filtradas_hist = st.multiselect(
            "🔍 Filtrar Contas Específicas (Estilo Excel):",
            options=linhas_hist,
            default=[],
            key="filtro_contas_hist",
        )
        if contas_filtradas_hist:
            linhas_hist = contas_filtradas_hist

    target_dfs = list_df_real if tipo_hist == "Realizado" else list_df_orc

    # ---- KPIs contextuais do histórico (referência: Receita Operacional Líquida) ----
    valores_ref_mensal = {
        m_nome: get_valor_consolidado_multi(target_dfs, "3 - Receita Operacional Liquida", [m_col])
        for m_nome, m_col in m_map.items()
    }
    meses_com_dado = {m: v for m, v in valores_ref_mensal.items() if v != 0}

    if meses_com_dado:
        mes_melhor = max(meses_com_dado, key=meses_com_dado.get)
        mes_pior = min(meses_com_dado, key=meses_com_dado.get)
        media_mensal_hist = sum(meses_com_dado.values()) / len(meses_com_dado)

        st.markdown(
            render_kpi_row([
                dict(label=f"MELHOR MÊS ({tipo_hist.upper()})", value=formata_brl(meses_com_dado[mes_melhor]),
                     value_color=COLORS["positive"], subtext=mes_melhor.capitalize(), icon="🏆"),
                dict(label=f"PIOR MÊS ({tipo_hist.upper()})", value=formata_brl(meses_com_dado[mes_pior]),
                     value_color=COLORS["negative"], subtext=mes_pior.capitalize(), icon="📉"),
                dict(label="MÉDIA MENSAL (RECEITA)", value=formata_brl(media_mensal_hist),
                     value_color=COLORS["text"], subtext=f"{len(meses_com_dado)} meses com dados", icon="📊"),
                dict(label="AMPLITUDE (MELHOR - PIOR)", value=formata_brl(meses_com_dado[mes_melhor] - meses_com_dado[mes_pior]),
                     value_color=COLORS["muted_line"], subtext="Variação entre extremos", icon="↕️"),
            ]),
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

    hist_data = []
    for linha in linhas_hist:
        row_dict = {"Conta / Linha DRE": linha}
        soma_linha = 0.0
        for m_nome, m_col in m_map.items():
            val_m = get_valor_consolidado_multi(target_dfs, linha, [m_col])
            row_dict[m_nome] = val_m
            soma_linha += val_m
        row_dict["Total Acumulado"] = soma_linha
        hist_data.append(row_dict)

    df_hist = pd.DataFrame(hist_data)

    col_config_hist = {
        "Conta / Linha DRE": st.column_config.TextColumn("Conta / Linha DRE", width="large", pinned=True),
    }

    colunas_numericas = list(m_map.keys()) + ["Total Acumulado"]
    format_dict_hist = {col: formata_brl for col in colunas_numericas}

    ALTURA_17_LINHAS = 633

    st.dataframe(
        df_hist.style.format(format_dict_hist).map(cor_valor, subset=colunas_numericas),
        column_config=col_config_hist,
        use_container_width=True,
        height=ALTURA_17_LINHAS,
        hide_index=True,
        on_select="rerun",
        selection_mode="multi-row",
    )


# ---------------------------------------------------------------------------
# ABA 4: PREVISÕES & TRENDS
# ---------------------------------------------------------------------------
with tab4:
    st.markdown('<div class="section-title">🔮 Painel Avançado de Previsões e Tendências 2026</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    c_f1, c_f2, c_f3 = st.columns([1.2, 1.2, 1.6])

    with c_f1:
        metrica_sel = st.selectbox("Métrica de Análise:", ["Receita Operacional Líquida", "EBITDA"], index=0)
        termo_metrica = "3 - Receita Operacional Liquida" if metrica_sel == "Receita Operacional Líquida" else "11 - EBITDA"

    with c_f2:
        modelo_proj = st.selectbox(
            "Modelo de Projeção Futura:",
            ["Média Histórica (Run-Rate)", "Manter Orçamento (Budget)", "Ajustado por Sazonalidade/Performance"],
            index=0,
            help="Define como os meses futuros não realizados serão calculados.",
        )

    with c_f3:
        sensibilidade = st.slider(
            "Ajuste Fino de Cenário / Estresse (%):",
            min_value=-20.0, max_value=20.0, value=0.0, step=1.0,
            help="Aplica uma variação percentual sobre os meses projetados.",
        )

    meses_todos = list(m_map.keys())

    meses_realizados_cols = (
        cols_kpi if tipo_periodo == "Mês Selecionado"
        else [m_map[m] for m in meses_todos if get_valor_consolidado_multi(list_df_real, termo_metrica, [m_map[m]]) != 0]
    )

    if not meses_realizados_cols:
        meses_realizados_cols = list(m_map.values())[:7]

    num_meses_realizados = len(meses_realizados_cols)

    val_real_acumulado = get_valor_consolidado_multi(list_df_real, termo_metrica, meses_realizados_cols)
    val_orc_acumulado = get_valor_consolidado_multi(list_df_orc, termo_metrica, meses_realizados_cols)

    val_orc_anual_total = get_valor_consolidado_multi(list_df_orc, termo_metrica, colunas_validas)

    media_mensal_real = val_real_acumulado / num_meses_realizados if num_meses_realizados > 0 else 0
    fator_performance = (val_real_acumulado / val_orc_acumulado) if val_orc_acumulado != 0 else 1.0

    fator_sensibilidade = 1.0 + (sensibilidade / 100.0)

    valores_finais_mes = []
    valores_orcado_mes = []
    valores_real_mes = []
    tipos_serie = []

    for idx_m, m_nome in enumerate(meses_todos):
        m_col = m_map[m_nome]
        v_orc = get_valor_consolidado_multi(list_df_orc, termo_metrica, [m_col])
        valores_orcado_mes.append(v_orc)

        if idx_m < num_meses_realizados:
            v_real = get_valor_consolidado_multi(list_df_real, termo_metrica, [m_col])
            valores_finais_mes.append(v_real)
            valores_real_mes.append(v_real)
            tipos_serie.append("Realizado")
        else:
            if modelo_proj == "Média Histórica (Run-Rate)":
                v_proj = media_mensal_real * fator_sensibilidade
            elif modelo_proj == "Manter Orçamento (Budget)":
                v_proj = v_orc * fator_sensibilidade
            else:
                v_proj = (v_orc * fator_performance) * fator_sensibilidade

            valores_finais_mes.append(v_proj)
            valores_real_mes.append(np.nan)
            tipos_serie.append("Projetado")

    projecao_total_anual = sum(valores_finais_mes)
    diff_anual = projecao_total_anual - val_orc_anual_total
    pct_atingimento_anual = (projecao_total_anual / val_orc_anual_total * 100) if val_orc_anual_total != 0 else 0

    c_proj = cor_variacao(projecao_total_anual)
    c_diff = cor_variacao(diff_anual)

    st.markdown("<br>", unsafe_allow_html=True)

    subtext_gap = f"{diff_anual / abs(val_orc_anual_total) * 100:+.1f}% vs Meta" if val_orc_anual_total != 0 else "—"

    st.markdown(
        render_kpi_row([
            dict(label=f"PROJEÇÃO ANUAL ({metrica_sel.upper()})", value=formata_brl(projecao_total_anual),
                 value_color=c_proj, subtext=f"Cenário: {modelo_proj.split(' ')[0]}", icon="🔮"),
            dict(label="META ANUAL (ORÇADO)", value=formata_brl(val_orc_anual_total), value_color=COLORS["text"],
                 subtext="Orçamento Fechado 2026", icon="🎯"),
            dict(label="GAP / DESVIO ANUAL", value=formata_brl(diff_anual), value_color=c_diff,
                 subtext=subtext_gap, subtext_color=c_diff, icon="📐"),
            dict(label="ATINGIMENTO ESTIMADO", value=f"{pct_atingimento_anual:.1f}%", value_color=c_diff,
                 subtext=f"Média Mensal Real: {formata_m(media_mensal_real)}", icon="📊"),
        ]),
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    df_trend = pd.DataFrame({
        "Mês": [m.capitalize() for m in meses_todos],
        "Valor Projetado/Real": valores_finais_mes,
        "Orçado": valores_orcado_mes,
        "Tipo": tipos_serie,
    })

    posicoes_meta = []
    posicoes_barras = []
    for val_p, val_o in zip(valores_finais_mes, valores_orcado_mes):
        if val_p >= val_o:
            posicoes_meta.append("bottom center")
            posicoes_barras.append("outside")
        else:
            posicoes_meta.append("top center")
            posicoes_barras.append("inside")

    fig_comb = go.Figure()

    df_real_bar = df_trend[df_trend["Tipo"] == "Realizado"]
    pos_bar_real = [posicoes_barras[i] for i in df_real_bar.index]
    fig_comb.add_trace(
        go.Bar(
            x=df_real_bar["Mês"], y=df_real_bar["Valor Projetado/Real"], name="Realizado",
            marker_color=COLORS["primary"],
            text=[formata_m(v) for v in df_real_bar["Valor Projetado/Real"]],
            textposition=pos_bar_real, textfont=dict(size=11, color=COLORS["text"]),
            cliponaxis=False,
        )
    )

    df_proj_bar = df_trend[df_trend["Tipo"] == "Projetado"]
    pos_bar_proj = [posicoes_barras[i] for i in df_proj_bar.index]
    fig_comb.add_trace(
        go.Bar(
            x=df_proj_bar["Mês"], y=df_proj_bar["Valor Projetado/Real"], name="Projetado (Tendência)",
            marker_color=COLORS["border_soft"],
            text=[formata_m(v) for v in df_proj_bar["Valor Projetado/Real"]],
            textposition=pos_bar_proj, textfont=dict(size=11, color=COLORS["text_muted"]),
            cliponaxis=False,
        )
    )

    fig_comb.add_trace(
        go.Scatter(
            x=df_trend["Mês"], y=df_trend["Orçado"], name="Orçado (Meta)",
            mode="lines+markers+text",
            text=[formata_m(v) for v in df_trend["Orçado"]],
            textposition=posicoes_meta, textfont=dict(size=10, color=COLORS["warning"]),
            line=dict(color=COLORS["warning"], width=2, dash="dash"),
            marker=dict(size=6, color=COLORS["warning"]),
            cliponaxis=False,
        )
    )

    max_val = max(
        max(df_trend["Valor Projetado/Real"].dropna(), default=0),
        max(df_trend["Orçado"].dropna(), default=0),
    )

    estilo_grafico(
        fig_comb,
        height=500,
        title=f"Evolução Mensal & Projeção Run-Rate: {metrica_sel}",
        xaxis=dict(gridcolor=COLORS["border"], zerolinecolor=COLORS["border"], fixedrange=True),
        yaxis=dict(
            showticklabels=False, gridcolor="rgba(0,0,0,0)",
            range=[0, max_val * 1.35] if max_val > 0 else None,
            fixedrange=True,
        ),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
        barmode="group",
    )
    st.plotly_chart(fig_comb, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    st.markdown('<div class="section-title">📋 Detalhamento da Projeção Mensal (R$)</div>', unsafe_allow_html=True)

    df_resumo_proj = pd.DataFrame({
        "Mês": [m.capitalize() for m in meses_todos],
        "Tipo de Dado": tipos_serie,
        "Valor Realizado / Projetado": valores_finais_mes,
        "Orçado Original": valores_orcado_mes,
        "Desvio (R$)": [v_p - v_o for v_p, v_o in zip(valores_finais_mes, valores_orcado_mes)],
    })

    ALTURA_12_LINHAS = 38 + len(df_resumo_proj) * 35

    st.dataframe(
        df_resumo_proj.style.format({
            "Valor Realizado / Projetado": formata_brl,
            "Orçado Original": formata_brl,
            "Desvio (R$)": formata_brl,
        }).map(cor_valor, subset=["Desvio (R$)"]),
        use_container_width=True,
        hide_index=True,
        height=ALTURA_12_LINHAS,
    )

# ---------------------------------------------------------------------------
# ABA 5: EMISSÃO DE RELATÓRIOS
# ---------------------------------------------------------------------------
with tab5:
    st.markdown('<div class="section-title">📤 Emissão de Relatórios em Excel</div>', unsafe_allow_html=True)
    st.caption(
        "Selecione as linhas da DRE desejadas e gere um relatório formatado com 3 planilhas: "
        "Resumo (consolidado), Detalhe Mensal (contas x meses, por loja) e Plano de Contas "
        "(composição de cada linha, por loja)."
    )
    st.markdown("<br>", unsafe_allow_html=True)

    linhas_relatorio = df_ref[col_nome].dropna().astype(str).unique()

    contas_relatorio = st.multiselect(
        "🔍 Selecione as linhas da DRE para o relatório:",
        options=linhas_relatorio,
        default=[],
        key="contas_relatorio",
    )

    col_btn, col_info = st.columns([1, 2])
    with col_btn:
        gerar_clicado = st.button(
            "📊 Gerar Relatório Excel",
            use_container_width=True,
            disabled=not contas_relatorio,
        )

    if gerar_clicado and contas_relatorio:
        with st.spinner("Carregando dados por loja e plano de contas..."):
            dados_por_loja_rel = carregar_dados_por_loja(path_orc, path_real, opcoes_unidades)
            df_tabela_contas = carregar_tabela_contas(path_real)
            mapa_planos_dre_rel = montar_mapa_planos_por_dre(df_tabela_contas)

        with st.spinner("Montando o relatório em Excel..."):
            excel_bytes = montar_relatorio_excel(
                contas_relatorio, list_df_real, list_df_orc, m_map, colunas_validas, label_visao,
                dados_por_loja=dados_por_loja_rel,
                mapa_planos_dre=mapa_planos_dre_rel,
            )
        st.session_state["relatorio_excel_bytes"] = excel_bytes
        st.session_state["relatorio_excel_nome"] = (
            f"Relatorio_DRE_{datetime.now(FUSO_BR).strftime('%Y%m%d_%H%M')}.xlsx"
        )
        st.success(f"Relatório gerado com {len(contas_relatorio)} conta(s) selecionada(s), pronto para download.")

    if not contas_relatorio:
        st.info("Selecione ao menos uma linha da DRE acima para habilitar a geração do relatório.")

    if st.session_state.get("relatorio_excel_bytes"):
        st.download_button(
            label="⬇️ Baixar Relatório (.xlsx)",
            data=st.session_state["relatorio_excel_bytes"],
            file_name=st.session_state.get("relatorio_excel_nome", "relatorio_dre.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("O relatório contém três planilhas: **Resumo** (total do ano por conta, consolidado), **Detalhe Mensal** (contas nas linhas x meses nas colunas, por loja) e **Plano de Contas** (composição de cada linha da DRE, por loja).")

# ---------------------------------------------------------------------------
# ABA 6: GESTÃO DE USUÁRIOS (somente administrador)
# ---------------------------------------------------------------------------
if eh_admin:
    with tab6:
        st.markdown('<div class="section-title">👥 Gestão de Usuários</div>', unsafe_allow_html=True)
        st.caption(
            "Cadastre novos usuários de **visualização** (eles só podem usar os filtros, "
            "visualizar o painel e gerar relatórios — não têm acesso a esta aba). "
            "Como o painel não tem banco de dados, o cadastro gera um bloco de texto que "
            "você precisa colar nos **Secrets** do app (Configurações do app → Secrets)."
        )
        st.markdown("<br>", unsafe_allow_html=True)

        usuarios_atuais = obter_usuarios_cadastrados()
        st.markdown("**Usuários atualmente configurados nos Secrets:**")
        lista_usuarios = [
            {"E-mail": u["email"], "Perfil": u["perfil"]}
            for u in usuarios_atuais.values()
        ]
        st.dataframe(pd.DataFrame(lista_usuarios), use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**➕ Novo usuário de visualização**")

        with st.form("form_novo_usuario"):
            col_email, col_senha = st.columns(2)
            with col_email:
                novo_email = st.text_input("E-mail do novo usuário")
            with col_senha:
                nova_senha = st.text_input("Senha do novo usuário", type="password")
            gerar_clicado_usuario = st.form_submit_button("Gerar bloco para os Secrets")

        if gerar_clicado_usuario:
            if not novo_email or not nova_senha:
                st.warning("Preencha e-mail e senha para gerar o cadastro.")
            elif novo_email.strip().lower() in usuarios_atuais:
                st.error("Já existe um usuário cadastrado com esse e-mail.")
            else:
                apelido = re.sub(r"[^a-z0-9]+", "_", novo_email.strip().lower().split("@")[0]).strip("_")
                senha_escapada = nova_senha.replace('"', '\\"')
                bloco_toml = (
                    f'[usuarios.{apelido}]\n'
                    f'email = "{novo_email.strip().lower()}"\n'
                    f'senha = "{senha_escapada}"\n'
                    f'perfil = "visualizacao"'
                )
                st.success("Copie o bloco abaixo e cole nos **Secrets** do app (em uma nova linha, mantendo os que já existem).")
                st.code(bloco_toml, language="toml")
                st.caption(
                    "Depois de colar e salvar nos Secrets, o app reinicia sozinho e o novo "
                    "usuário já consegue entrar com o e-mail e a senha cadastrados."
                )


# ============================================================================
# 9. RODAPÉ
# ============================================================================
st.markdown(
    f"""
    <div class="footer-note">
        Controladoria B&A · Painel Financeiro 2026 · Dados atualizados automaticamente a cada 60s
    </div>
    """,
    unsafe_allow_html=True,
)