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
import time
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
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.workbook.defined_name import DefinedName

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
        # Separadores no padrão numérico brasileiro (vírgula decimal, ponto de
        # milhar) -- afeta eixos, indicadores e qualquer número que o próprio
        # Plotly formatar automaticamente.
        separators=",.",
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


# ---------------------------------------------------------------------------
# Funções de cálculo/formatação (usadas em todo o painel, inclusive no Painel
# de TV, que é montado antes da barra lateral -- por isso ficam aqui perto do
# topo, e não mais lá embaixo na seção 6 original).
# ---------------------------------------------------------------------------
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


def eh_linha_custos_despesas(nome_linha):
    """Retorna True se a linha pertence a algum dos grupos de Custo das
    Vendas (4), Despesas Variáveis (6) ou Despesas Operacionais (8) --
    inclusive as sublinhas deles (ex.: "6.6 - ...", "8.3.1 - ..."). Usada na
    visão "Gerencial" (foco em custos e despesas) das abas de DRE e
    Histórico Mensal."""
    m = re.match(r"^(\d+)", str(nome_linha).strip())
    return bool(m) and m.group(1) in ("4", "6", "8")


def cor_valor(val):
    if pd.isna(val):
        return ""
    color = COLORS["positive"] if val >= 0 else COLORS["negative"]
    return f"color: {color}; font-weight: 500;"


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
        .sidebar-brand .brand-logo {{
            width: 30px; height: 30px; border-radius: 50%;
            object-fit: cover;
            box-shadow: 0 0 10px {COLORS["primary_soft"]};
            flex-shrink: 0;
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

        /* Faixa de contexto — fina, uma linha só, para não repetir um bloco
           grande e genérico em cima de todas as abas (libera espaço vertical
           pra cada aba usar com informação própria/diferente) */
        .top-status-strip {{
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 8px;
            padding: 6px 4px 10px 4px;
            border-bottom: 1px solid {COLORS["border"]};
            margin-bottom: 12px;
            font-size: 12px;
            color: {COLORS["text_muted"]};
        }}
        .top-status-strip b {{ color: {COLORS["text"]}; font-weight: 600; }}
        .top-status-strip .chip {{
            display: inline-block;
            background: {COLORS["primary_soft"]};
            color: {COLORS["primary"]};
            border-radius: 20px;
            padding: 1px 10px;
            font-weight: 700;
            font-size: 11px;
            letter-spacing: 0.2px;
        }}
        .top-status-strip .sep {{ color: {COLORS["border"]}; }}

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
            .top-status-strip {{ font-size: 11px !important; padding: 5px 2px 8px 2px !important; }}
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


# Trecho de JS reaproveitado em todos os componentes abaixo: tenta achar a
# janela "de verdade" (a aba do navegador) mesmo quando o painel roda dentro
# de um iframe (que é como o Streamlit renderiza os componentes -- e é
# também o motivo mais comum de "lembrar de mim" não funcionar de verdade:
# usar só `window.top` quebra sempre que o navegador bloqueia o acesso
# entre frames, o que muitos navegadores modernos fazem por padrão).
_JS_ACHAR_JANELA = """
function _janelaAlvo() {
    const candidatas = [window.top, window.parent, window];
    for (const w of candidatas) {
        try {
            if (w && w.localStorage) { w.localStorage.length; return w; }
        } catch (e) { /* sem acesso a essa janela, tenta a próxima */ }
    }
    return null;
}
function _lerCookie(nome) {
    const partes = ("; " + document.cookie).split("; " + nome + "=");
    if (partes.length === 2) return partes.pop().split(";").shift();
    return null;
}
"""


def _salvar_credenciais_no_navegador(email, senha):
    """Grava e-mail/senha (ofuscados em base64 -- isso NAO e criptografia) no
    navegador, para preencher o login sozinho na proxima vez que a pessoa
    abrir o painel. Usa DOIS mecanismos em paralelo (localStorage da janela
    "de verdade" + cookie de longa duração no documento do próprio
    componente) porque, dependendo do navegador/implantação, um dos dois
    pode estar bloqueado -- assim o recurso continua funcionando mesmo se
    um deles falhar."""
    email_b64 = base64.b64encode(email.encode("utf-8")).decode("ascii")
    senha_b64 = base64.b64encode(senha.encode("utf-8")).decode("ascii")
    components.html(
        f"""
        <script>
        {_JS_ACHAR_JANELA}
        const dados = JSON.stringify({{ e: "{email_b64}", s: "{senha_b64}" }});
        try {{
            const w = _janelaAlvo();
            if (w) w.localStorage.setItem('{CHAVE_LOCALSTORAGE_LOGIN}', dados);
        }} catch (e) {{}}
        try {{
            const validade = new Date();
            validade.setDate(validade.getDate() + 90);
            document.cookie = '{CHAVE_LOCALSTORAGE_LOGIN}=' + encodeURIComponent(dados) +
                '; expires=' + validade.toUTCString() + '; path=/; SameSite=Lax';
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
        {_JS_ACHAR_JANELA}
        try {{ const w = _janelaAlvo(); if (w) w.localStorage.removeItem('{CHAVE_LOCALSTORAGE_LOGIN}'); }} catch (e) {{}}
        try {{ document.cookie = '{CHAVE_LOCALSTORAGE_LOGIN}=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;'; }} catch (e) {{}}
        </script>
        """,
        height=0,
        width=0,
    )


def _tentar_autologin_via_url():
    """Se a URL trouxer credenciais (?le=&ls=) vindas do navegador, tenta
    logar automaticamente com elas. Retorna True se conseguiu logar.

    IMPORTANTE: remove só as chaves 'le'/'ls' da URL -- nunca usar
    st.query_params.clear() aqui, porque isso apagaria TAMBÉM outros
    parâmetros importantes da URL, como o "?modo=tv" do Painel de TV
    (era exatamente por isso que o link do Painel de TV caía sempre
    no painel normal: o parâmetro "modo" era apagado antes mesmo de
    ser lido)."""
    le = st.query_params.get("le")
    ls = st.query_params.get("ls")
    if "le" in st.query_params:
        del st.query_params["le"]
    if "ls" in st.query_params:
        del st.query_params["ls"]
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
    credenciais salvas no navegador (localStorage OU cookie, o que estiver
    disponível) e, se houver, recarrega a pagina com elas na URL para
    tentarmos o autologin."""
    components.html(
        f"""
        <script>
        {_JS_ACHAR_JANELA}
        try {{
            let salvo = null;
            const w = _janelaAlvo();
            if (w) {{
                try {{ salvo = w.localStorage.getItem('{CHAVE_LOCALSTORAGE_LOGIN}'); }} catch (e) {{}}
            }}
            if (!salvo) {{
                const viaCookie = _lerCookie('{CHAVE_LOCALSTORAGE_LOGIN}');
                if (viaCookie) salvo = decodeURIComponent(viaCookie);
            }}
            if (salvo) {{
                const dados = JSON.parse(salvo);
                if (dados.e && dados.s) {{
                    const janelaUrl = (w || window).location;
                    const url = new URL(janelaUrl.href);
                    if (!url.searchParams.get('le')) {{
                        url.searchParams.set('le', dados.e);
                        url.searchParams.set('ls', dados.s);
                        (w || window).location.replace(url.toString());
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

    xls_real = pd.ExcelFile(path_real)

    # Abas auxiliares/de referência: usadas para buscar dados (DE_PARA de
    # loja x centro de custo, lançamentos do DIÁRIO, mapeamento de plano de
    # contas x linha da DRE etc.), mas que NÃO são lojas/unidades -- por isso
    # nunca devem aparecer no filtro de lojas do painel nem do relatório.
    abas_ignorar = [
        "Sint Ebt loja", "CONS 25X26 V.1", "CONS 25X26 V.2",
        "DE_PARA", "DIÁRIO", "DIARIO", "DRE_MENSAL", "Tabela_Contas", "Tabela_Lojas",
    ]

    # IMPORTANTE: a lista de abas/lojas disponíveis precisa ser a UNIÃO das abas
    # dos dois arquivos (Orçado e Realizado), e não só do Orçado. Uma loja nova
    # (ex.: "LJ ASSAI 23157") pode já existir como aba no Realizado 2026 mas
    # ainda não ter sido criada no Orçado -- se olharmos só para o Orçado, essa
    # loja nunca aparece em lugar nenhum do painel nem do relatório, mesmo tendo
    # dados reais lançados. Mantemos a ORDEM das abas do Orçado primeiro (para
    # não bagunçar o layout já existente) e acrescentamos ao final quaisquer
    # abas que só existam no Realizado.
    abas_orc = list(xls_orc.sheet_names)
    abas_real = list(xls_real.sheet_names)
    abas_uniao = abas_orc + [a for a in abas_real if a not in abas_orc]

    abas_validas = [sheet for sheet in abas_uniao if sheet not in abas_ignorar]

    return abas_validas, path_orc, path_real


try:
    with st.spinner("Conectando às planilhas financeiras..."):
        abas_disponiveis, path_orc, path_real = obter_caminhos_excel()
except Exception as e:
    st.error(f"Erro ao carregar as planilhas: {e}")
    st.stop()


def _ler_aba_ou_vazio(path, aba, colunas_modelo=None):
    """Lê uma aba de uma planilha. Se a aba não existir naquele arquivo
    específico (ex.: loja nova que só tem aba no Realizado, ainda sem aba
    correspondente no Orçado), retorna um DataFrame vazio com as mesmas
    colunas do modelo de referência, em vez de derrubar a loja inteira do
    relatório/painel."""
    try:
        return pd.read_excel(path, sheet_name=aba)
    except Exception:
        if colunas_modelo is not None:
            return pd.DataFrame(columns=colunas_modelo)
        return pd.DataFrame()


@st.cache_data(ttl=60)
def carregar_dados_abas(path_o, path_r, lista_abas):
    """Carrega Orçado e Realizado de cada aba/loja. Cada lado é lido de forma
    INDEPENDENTE: se a loja não existir no Orçado (ex.: loja nova, ainda sem
    orçamento cadastrado) mas existir no Realizado (ou vice-versa), ainda
    assim o lado que existe é carregado normalmente -- o outro lado entra
    como zero. Antes, se faltasse em UM dos dois arquivos, a loja inteira
    era descartada dos dois lados, e por isso lojas novas como a
    "LJ ASSAI 23157" desapareciam de todo o relatório e do Resumo."""
    dfs_o = []
    dfs_r = []
    colunas_modelo_o, colunas_modelo_r = None, None
    for aba in lista_abas:
        df_o = _ler_aba_ou_vazio(path_o, aba, colunas_modelo_o)
        df_r = _ler_aba_ou_vazio(path_r, aba, colunas_modelo_r)
        if not df_o.empty and colunas_modelo_o is None:
            colunas_modelo_o = df_o.columns
        if not df_r.empty and colunas_modelo_r is None:
            colunas_modelo_r = df_r.columns
        dfs_o.append(df_o)
        dfs_r.append(df_r)
    return dfs_o, dfs_r


@st.cache_data(ttl=60)
def carregar_dados_por_loja(path_o, path_r, lista_lojas):
    """Carrega os dados de Orçado/Realizado de cada loja SEPARADAMENTE (uma aba
    por loja), para permitir a divisão por loja no relatório Excel — independente
    do modo de visão escolhido na barra lateral (Consolidado ou Unidades).
    Cada lado (Orçado/Realizado) é lido de forma independente -- uma loja sem
    aba no Orçado ainda entra no relatório com os valores do Realizado (e
    Orçado zerado), em vez de ser descartada por completo."""
    dados_por_loja = {}
    for loja in lista_lojas:
        df_o = _ler_aba_ou_vazio(path_o, loja)
        df_r = _ler_aba_ou_vazio(path_r, loja)
        if df_o.empty and df_r.empty:
            continue
        dados_por_loja[loja] = (df_o, df_r)
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


# ---------------------------------------------------------------------------
# 4.1.1 TABELA_LOJAS — DE_PARA entre o nome da Loja (aba) e o Centro de Custos
# usado na DIÁRIO. Essencial para o detalhamento de lançamentos: a DIÁRIO
# identifica a loja pelo "Centro de Custos", que normalmente NÃO é igual ao
# nome da aba da loja (ex.: aba "LJ ASSAI 23157" pode ter Centro de Custos
# "23157" ou outro código) -- sem esse DE_PARA, o filtro por loja no DIÁRIO
# nunca encontra nada, e por isso os lançamentos apareciam sempre vazios.
# ---------------------------------------------------------------------------
@st.cache_data(ttl=300)
def carregar_tabela_lojas(path_r):
    """Lê a aba Tabela_Lojas da planilha Realizado 2026: relação (DE_PARA)
    entre o nome da Loja e o Centro de Custos correspondente."""
    df = None
    for nome_aba in ("Tabela_Lojas", "Tabela Lojas", "TabelaLojas", "tabela_lojas", "DE_PARA"):
        try:
            df = pd.read_excel(path_r, sheet_name=nome_aba)
            break
        except Exception:
            continue
    if df is None:
        return pd.DataFrame(columns=["Loja", "Centro de Custos"])
    df.columns = [str(c).strip() for c in df.columns]
    return df


def montar_mapa_loja_centro_custo(df_tabela_lojas):
    """A partir da Tabela_Lojas (ou DE_PARA), retorna {nome_da_loja:
    centro_de_custo}. Tenta reconhecer as colunas pelo nome (tolerante a
    várias variações, inclusive um DE_PARA genérico "De"/"Para"); se não
    encontrar, cai para as duas primeiras colunas da aba. Monta o mapa nos
    DOIS sentidos (loja->CC e CC->loja) e devolve a união, para não depender
    de qual das duas colunas veio primeiro na planilha."""
    mapa = {}
    if df_tabela_lojas is None or df_tabela_lojas.empty:
        return mapa

    cols_lower = {str(c).strip().lower(): c for c in df_tabela_lojas.columns}
    nomes_loja = ["loja", "nome da loja", "unidade", "nome loja", "nome unidade", "de"]
    nomes_cc = [
        "centro de custos", "centro de custo", "cod centro de custo", "código centro de custo",
        "centro custo", "cc", "cod cc", "código cc", "para",
    ]
    col_loja = next((cols_lower[c] for c in nomes_loja if c in cols_lower), None)
    col_cc = next((cols_lower[c] for c in nomes_cc if c in cols_lower), None)

    if col_loja is None or col_cc is None:
        if len(df_tabela_lojas.columns) >= 2:
            col_loja, col_cc = df_tabela_lojas.columns[0], df_tabela_lojas.columns[1]
        else:
            return mapa

    for _, linha in df_tabela_lojas.iterrows():
        loja = str(linha[col_loja]).strip()
        centro_custo = str(linha[col_cc]).strip()
        if loja and centro_custo and loja.lower() != "nan" and centro_custo.lower() != "nan":
            mapa[loja] = centro_custo
            mapa.setdefault(centro_custo, loja)
    return mapa


# ---------------------------------------------------------------------------
# 4.2 DIÁRIO — lançamentos detalhados, fonte principal da aba "Plano de Contas"
# ---------------------------------------------------------------------------
@st.cache_data(ttl=300)
def carregar_diario(path_r):
    """Lê a aba DIÁRIO da planilha Realizado 2026: lançamentos detalhados,
    com Valor Bruto, Competência (data do lançamento), Plano de Contas,
    Centro de Custos (loja/unidade), Linha DRE (a linha da DRE a que aquele
    plano de contas pertence) e, quando disponíveis, Número, Cliente /
    Fornecedor e Histórico. É a fonte usada para montar as abas "Plano de
    Contas" e "Lançamentos" do relatório em Excel."""
    colunas_saida = [
        "Valor Bruto", "Competência", "Plano de Contas", "Centro de Custos", "Linha DRE",
        "Número", "Cliente / Fornecedor", "Histórico", "Mês",
    ]

    apelidos_coluna = {
        "valor bruto": "Valor Bruto",
        "competência": "Competência",
        "competencia": "Competência",
        "plano de contas": "Plano de Contas",
        "centro de custos": "Centro de Custos",
        "centro de custo": "Centro de Custos",
        "linha dre": "Linha DRE",
        "número": "Número",
        "numero": "Número",
        "nº": "Número",
        "n°": "Número",
        "num": "Número",
        "cliente / fornecedor": "Cliente / Fornecedor",
        "cliente/fornecedor": "Cliente / Fornecedor",
        "fornecedor / cliente": "Cliente / Fornecedor",
        "fornecedor/cliente": "Cliente / Fornecedor",
        "cliente": "Cliente / Fornecedor",
        "fornecedor": "Cliente / Fornecedor",
        "histórico": "Histórico",
        "historico": "Histórico",
        "descrição": "Histórico",
        "descricao": "Histórico",
    }

    df = None
    for nome_aba in ("DIÁRIO", "DIARIO", "Diário", "Diario"):
        try:
            df = pd.read_excel(path_r, sheet_name=nome_aba)
            break
        except Exception:
            continue
    if df is None:
        return pd.DataFrame(columns=colunas_saida)

    df.columns = [str(c).strip() for c in df.columns]
    renomeio = {c: apelidos_coluna[c.strip().lower()] for c in df.columns if c.strip().lower() in apelidos_coluna}
    df = df.rename(columns=renomeio)

    colunas_necessarias = ["Valor Bruto", "Competência", "Plano de Contas", "Centro de Custos", "Linha DRE"]
    if any(c not in df.columns for c in colunas_necessarias):
        return pd.DataFrame(columns=colunas_saida)

    # Colunas extras (opcionais) para dar mais contexto na aba "Lançamentos".
    # Se a DIÁRIO não tiver alguma delas, entra em branco -- não impede o resto.
    colunas_extras = ["Número", "Cliente / Fornecedor", "Histórico"]
    for col_extra in colunas_extras:
        if col_extra not in df.columns:
            df[col_extra] = ""

    df = df[colunas_necessarias + colunas_extras].copy()
    df["Valor Bruto"] = pd.to_numeric(df["Valor Bruto"], errors="coerce").fillna(0)
    df["Plano de Contas"] = df["Plano de Contas"].astype(str).str.strip()
    df["Centro de Custos"] = df["Centro de Custos"].astype(str).str.strip()
    df["Linha DRE"] = df["Linha DRE"].astype(str).str.strip()
    for col_extra in colunas_extras:
        df[col_extra] = df[col_extra].fillna("").astype(str).str.strip()

    # Competência normalmente vem como data (dd/mm/aaaa) -> convertemos para "mm/aaaa"
    # para casar com as colunas de mês (ex.: "01/2026") usadas no resto do painel.
    competencia_dt = pd.to_datetime(df["Competência"], errors="coerce", dayfirst=True)
    mes_formatado = competencia_dt.dt.strftime("%m/%Y")
    mask_sem_data = mes_formatado.isna()
    if mask_sem_data.any():
        # Se não deu para converter em data, usa o próprio texto da célula
        # (cobre o caso de a Competência já vir como "mm/aaaa" digitada).
        mes_formatado = mes_formatado.copy()
        mes_formatado[mask_sem_data] = df.loc[mask_sem_data, "Competência"].astype(str).str.strip()
    df["Mês"] = mes_formatado

    # IMPORTANTE: filtra só por Plano de Contas preenchido -- NÃO exige Linha
    # DRE preenchida, porque planos de contas "fora da DRE" (ex.: Mercadorias
    # no modelo de Compras) legitimamente não têm Linha DRE, e precisam
    # continuar aparecendo aqui para serem encontrados pelo relatório.
    df = df[df["Plano de Contas"] != ""]
    return df.reset_index(drop=True)


def _normalizar_texto(txt):
    """Normaliza texto para comparação tolerante (maiúsculas, sem espaços duplicados)."""
    return re.sub(r"\s+", " ", str(txt or "").strip().upper())


def _mask_loja_por_centro_custo(df_diario, candidatos_loja):
    """Retorna a máscara (True/False por linha) de correspondência da coluna
    "Centro de Custos" do DIÁRIO contra um ou mais candidatos de loja.

    Aceita tanto uma string única quanto uma lista de candidatos -- isso é
    usado para tentar, na ordem, o Centro de Custos resolvido pela
    Tabela_Lojas (DE_PARA) E o nome da própria loja/aba, sem depender de um
    único mapeamento estar 100% correto (se o DE_PARA não achar nada, ainda
    tentamos casar direto pelo nome da loja, e vice-versa)."""
    if isinstance(candidatos_loja, str):
        candidatos_loja = [candidatos_loja]
    candidatos_norm = [c for c in (_normalizar_texto(c) for c in candidatos_loja) if c]

    centros_norm = df_diario["Centro de Custos"].map(_normalizar_texto)

    # 1ª tentativa: correspondência exata com qualquer um dos candidatos.
    for cand in candidatos_norm:
        mask = centros_norm == cand
        if mask.any():
            return mask

    # 2ª tentativa: correspondência "contém" (num sentido ou no outro) com
    # qualquer um dos candidatos.
    for cand in candidatos_norm:
        mask = centros_norm.apply(lambda x: (cand in x) or (x in cand) if x else False)
        if mask.any():
            return mask

    return pd.Series(False, index=df_diario.index)


def _filtrar_diario_por_loja_e_conta(df_diario, loja, conta):
    """Filtra o DIÁRIO pelo Centro de Custos (loja) e pela Linha DRE (conta),
    com correspondência tolerante (exata primeiro, depois por contém).
    "loja" pode ser uma string ou uma lista de candidatos (ver
    _mask_loja_por_centro_custo)."""
    if df_diario is None or df_diario.empty:
        return df_diario

    conta_norm = _normalizar_texto(conta)

    mask_loja = _mask_loja_por_centro_custo(df_diario, loja)

    linhas_norm = df_diario["Linha DRE"].map(_normalizar_texto)
    mask_conta = linhas_norm == conta_norm
    if not mask_conta.any():
        mask_conta = linhas_norm.apply(lambda x: (conta_norm in x) or (x in conta_norm) if x else False)

    return df_diario[mask_loja & mask_conta]


def _filtrar_diario_por_loja_e_plano(df_diario, loja, plano):
    """Filtra o DIÁRIO pelo Centro de Custos (loja) e pelo nome exato do
    Plano de Contas, IGNORANDO a Linha DRE -- usado para planos de contas
    que fazem parte de um modelo de relatório mas não têm linha da DRE
    correspondente (ex.: "Mercadorias" no modelo de Compras). "loja" pode
    ser uma string ou uma lista de candidatos."""
    if df_diario is None or df_diario.empty:
        return df_diario

    plano_norm = _normalizar_texto(plano)

    mask_loja = _mask_loja_por_centro_custo(df_diario, loja)

    planos_norm = df_diario["Plano de Contas"].map(_normalizar_texto)
    mask_plano = planos_norm == plano_norm
    if not mask_plano.any():
        mask_plano = planos_norm.apply(lambda x: (plano_norm in x) or (x in plano_norm) if x else False)

    return df_diario[mask_loja & mask_plano]


def montar_composicao_plano_direto(df_diario, loja, plano, mapa_meses):
    """Retorna os valores mês a mês (na ordem de mapa_meses) de um Plano de
    Contas específico, direto do DIÁRIO, para uma loja -- sem passar pela
    Linha DRE."""
    df_filtrado = _filtrar_diario_por_loja_e_plano(df_diario, loja, plano)
    if df_filtrado is None or df_filtrado.empty:
        return [0.0] * len(mapa_meses)
    return [df_filtrado.loc[df_filtrado["Mês"] == m_col, "Valor Bruto"].sum() for m_col in mapa_meses.values()]


def montar_composicao_diario(df_diario, loja, conta, mapa_meses):
    """A partir do DIÁRIO, retorna {plano_de_contas: [valores_por_mes...]}
    para uma loja (Centro de Custos) e uma linha da DRE (conta) específicas,
    já na ordem de mapa_meses."""
    df_filtrado = _filtrar_diario_por_loja_e_conta(df_diario, loja, conta)
    if df_filtrado is None or df_filtrado.empty:
        return {}

    composicao = {}
    for plano, grupo in df_filtrado.groupby("Plano de Contas"):
        composicao[plano] = [grupo.loc[grupo["Mês"] == m_col, "Valor Bruto"].sum() for m_col in mapa_meses.values()]
    return composicao


# ============================================================================
# 4.3 PAINEL PARA TV — visão executiva "tech", cheia de KPIs, gauges e
# ranking de lojas, para deixar exibida numa tela/TV da empresa. Acessada por
# uma URL própria (?modo=tv), separada da sessão normal de quem está
# navegando no painel -- assim dá pra deixar aberta numa TV/monitor sem
# interferir (nem ser afetada) pelos filtros que um usuário está usando na
# sua própria sessão. Se auto-atualiza a cada 60s e tem botão de tela cheia.
# ============================================================================
ABAS_CONSOLIDADAS_TV = [
    "DRE CONSOLIDADO", "ABPR CONSOLIDADO", "VD CONSOLIDADO",
    "LJ CONSOLIDADO", "ABPR + VD", "LJ - G&A",
]


def _obter_aba_consolidada_padrao(lista_abas):
    prioridade = ["DRE CONSOLIDADO", "LJ CONSOLIDADO", "ABPR + VD", "LJ - G&A"]
    for nome in prioridade:
        if nome in lista_abas:
            return nome
    return lista_abas[0] if lista_abas else None


def renderizar_painel_tv(path_orc, path_real, abas_disponiveis):
    aba_escolhida = _obter_aba_consolidada_padrao(abas_disponiveis)
    if not aba_escolhida:
        st.error("Não foi possível carregar dados para o Painel de TV.")
        return

    list_df_real_tv, list_df_orc_tv = carregar_dados_abas(path_orc, path_real, [aba_escolhida])

    meses_cols_tv = [
        "01/2026", "02/2026", "03/2026", "04/2026", "05/2026", "06/2026",
        "07/2026", "08/2026", "09/2026", "10/2026", "11/2026", "12/2026",
    ]
    nomes_meses_tv = [
        "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
        "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
    ]
    df_ref_tv = list_df_real_tv[0] if list_df_real_tv else pd.DataFrame()
    colunas_validas_tv = [m for m in meses_cols_tv if m in df_ref_tv.columns]
    m_map_tv = {n: c for n, c in zip(nomes_meses_tv, meses_cols_tv) if c in colunas_validas_tv}

    agora = datetime.now(FUSO_BR)
    idx_mes_atual = min(max(agora.month - 1, 0), len(m_map_tv) - 1) if m_map_tv else 0
    cols_ytd = list(m_map_tv.values())[: idx_mes_atual + 1]

    # ---- CSS "quiosque tech": some com sidebar/header, grade de fundo, glow ----
    st.markdown(
        f"""
        <style>
            [data-testid="stSidebar"], header[data-testid="stHeader"], footer {{ display: none !important; }}
            .block-container {{ padding: 1rem 2.2rem 1rem 2.2rem !important; max-width: 100% !important; }}
            .stApp {{
                background:
                    radial-gradient(circle at 15% 0%, rgba(76,141,255,0.09) 0%, transparent 42%),
                    radial-gradient(circle at 90% 100%, rgba(62,207,142,0.06) 0%, transparent 45%),
                    linear-gradient(180deg, #0A0D16 0%, #05070c 100%) !important;
            }}
            @keyframes tv-pulse {{ 0%,100% {{ opacity: 1; }} 50% {{ opacity: 0.35; }} }}
            @keyframes tv-marquee {{ 0% {{ transform: translateX(100%); }} 100% {{ transform: translateX(-100%); }} }}
            .tv-header {{
                display: flex; justify-content: space-between; align-items: center;
                padding: 4px 4px 16px 4px; border-bottom: 1px solid {COLORS["border"]}; margin-bottom: 18px;
            }}
            .tv-header .brand {{ display:flex; align-items:center; gap:14px; }}
            .tv-header img.logo {{ width: 42px; height: 42px; border-radius: 50%; box-shadow: 0 0 14px rgba(76,141,255,0.35); }}
            .tv-header h1 {{ font-size: 26px; font-weight: 800; color: {COLORS["text"]}; margin: 0; letter-spacing: 0.3px; }}
            .tv-header .sub {{ color: {COLORS["text_muted"]}; font-size: 13px; margin-top: 3px; }}
            .tv-live-pill {{
                display: inline-flex; align-items: center; gap: 6px; background: rgba(62,207,142,0.12);
                border: 1px solid {COLORS["positive"]}; color: {COLORS["positive"]}; border-radius: 20px;
                padding: 3px 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.6px; margin-left: 12px;
            }}
            .tv-live-pill .dot {{ width: 7px; height: 7px; border-radius: 50%; background: {COLORS["positive"]}; animation: tv-pulse 1.4s infinite; }}
            .tv-clock {{ text-align: right; color: {COLORS["text_muted"]}; font-size: 12.5px; }}
            .tv-clock b {{
                color: {COLORS["primary"]}; font-size: 30px; display:block; letter-spacing: 2px;
                font-family: 'Consolas','Courier New',monospace;
            }}
            .tv-kpi-grid {{ display: flex; gap: 16px; margin-bottom: 20px; }}
            .tv-kpi {{
                flex: 1; background: {COLORS["surface"]};
                border: 1px solid {COLORS["border"]}; border-radius: 14px; padding: 18px 20px;
                border-top: 3px solid var(--tv-accent, {COLORS["primary"]});
            }}
            .tv-kpi .lbl {{ font-size: 11px; font-weight: 700; letter-spacing: 0.6px; text-transform: uppercase; color: {COLORS["text_muted"]}; }}
            .tv-kpi .val {{ font-size: 28px; font-weight: 800; margin-top: 6px; letter-spacing: -0.5px; }}
            .tv-kpi .sub {{ font-size: 12px; margin-top: 5px; color: {COLORS["muted_line"]}; }}
            .tv-section-title {{
                font-size: 13px; font-weight: 700; color: {COLORS["text_muted"]}; text-transform: uppercase;
                letter-spacing: 0.6px; margin: 4px 0 10px 2px; border-left: 3px solid {COLORS["primary"]}; padding-left: 8px;
            }}
            .tv-panel {{
                background: {COLORS["surface"]};
                border: 1px solid {COLORS["border"]}; border-radius: 14px; padding: 14px 16px 4px 16px;
                margin-bottom: 18px; height: 100%;
            }}
            .tv-rank-row {{ display:flex; align-items:center; gap:10px; padding: 7px 2px; border-bottom: 1px dashed {COLORS["border_soft"]}; }}
            .tv-rank-badge {{
                width:22px; height:22px; border-radius:50%; background:{COLORS["primary_soft"]}; color:{COLORS["primary"]};
                font-size:11px; font-weight:800; display:flex; align-items:center; justify-content:center; flex-shrink:0;
            }}
            .tv-rank-name {{ flex:1; font-size:12.5px; color:{COLORS["text"]}; font-weight:600; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
            .tv-rank-bar-bg {{ flex:1.4; background:{COLORS["border"]}; border-radius:4px; height:8px; overflow:hidden; }}
            .tv-rank-bar-fill {{ height:100%; border-radius:4px; background: linear-gradient(90deg, {COLORS["primary"]}, {COLORS["positive"]}); }}
            .tv-rank-val {{ font-size:11.5px; color:{COLORS["muted_line"]}; width: 78px; text-align:right; font-family:'Consolas','Courier New',monospace; }}
            .tv-ticker-wrap {{
                overflow: hidden; white-space: nowrap; border-top: 1px solid {COLORS["border"]};
                border-bottom: 1px solid {COLORS["border"]}; padding: 9px 0; margin-top: 6px; background: rgba(255,255,255,0.015);
            }}
            .tv-ticker {{ display:inline-block; padding-left: 100%; animation: tv-marquee 150s linear infinite; font-size: 16px; color: {COLORS["text_muted"]}; }}
            .tv-ticker b {{ color: {COLORS["text"]}; }}
            .tv-ticker .tv-tick-sep {{ color: {COLORS["primary"]}; margin: 0 30px; }}
            .tv-fullscreen-hint {{
                position: fixed; bottom: 18px; right: 22px; z-index: 999999;
                color: {COLORS["text_muted"]}; font-size: 11px; text-align: right;
                background: rgba(0,0,0,0.35); border-radius: 8px; padding: 6px 12px;
                border: 1px solid {COLORS["border"]};
            }}
        </style>
        <span class="tv-fullscreen-hint">⛶ F11 para tela cheia · Esc para sair</span>
        """,
        unsafe_allow_html=True,
    )

    # ---------------- Cálculo dos indicadores ----------------
    rec_bruta_real = get_valor_consolidado_multi(list_df_real_tv, "1 - Receita Operacional Bruta", cols_ytd)
    rec_liq_real = get_valor_consolidado_multi(list_df_real_tv, "3 - Receita Operacional Liquida", cols_ytd)
    rec_liq_orc = get_valor_consolidado_multi(list_df_orc_tv, "3 - Receita Operacional Liquida", cols_ytd)
    ebitda_real = get_valor_consolidado_multi(list_df_real_tv, "11 - EBITDA", cols_ytd)
    ebitda_orc = get_valor_consolidado_multi(list_df_orc_tv, "11 - EBITDA", cols_ytd)
    margem_ebitda = (ebitda_real / rec_liq_real * 100) if rec_liq_real else 0
    margem_ebitda_orc = (ebitda_orc / rec_liq_orc * 100) if rec_liq_orc else 0
    desvio_ebitda = ebitda_real - ebitda_orc
    pct_desvio_ebitda = (desvio_ebitda / abs(ebitda_orc) * 100) if ebitda_orc else 0
    pct_atingimento_rec = (rec_liq_real / rec_liq_orc * 100) if rec_liq_orc else 0
    pct_atingimento_eb = (ebitda_real / ebitda_orc * 100) if ebitda_orc else 0
    desp_op_tv_kpi = abs(get_valor_consolidado_multi(list_df_real_tv, "8 - Despesas Operacionais", cols_ytd))

    col_head_a, col_head_b = st.columns([3, 1])
    with col_head_a:
        st.markdown(
            f"""
            <div class="tv-header">
                <div class="brand">
                    <img class="logo" src="data:image/jpeg;base64,{LOGO_BEEA_B64}" alt="Grupo Beea" />
                    <div>
                        <h1>Grupo B&amp;A · Painel Executivo <span class="tv-live-pill"><span class="dot"></span>AO VIVO</span></h1>
                        <div class="sub">{aba_escolhida} · Acumulado até {nomes_meses_tv[idx_mes_atual].capitalize()}/2026 · Dados atualizados a cada 30 minutos</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_head_b:
        # Relógio de verdade "ao vivo": roda dentro de um componente (só ele
        # consegue executar JavaScript de fato -- HTML solto via st.markdown
        # não executa <script>) e atualiza a cada segundo com o horário real
        # do computador de quem está vendo, em vez de ficar parado até o
        # próximo ciclo de atualização dos dados (60s).
        components.html(
            f"""
            <div style="text-align:right; font-family:'Consolas','Courier New',monospace;
                        color:{COLORS['text_muted']}; padding-top:6px;">
                <div id="tvClockLive" style="color:{COLORS['primary']}; font-size:28px; font-weight:800; letter-spacing:2px;">--:--:--</div>
                <div id="tvDateLive" style="font-size:12px;"></div>
            </div>
            <script>
            function _tvAtualizarRelogio() {{
                var agora = new Date();
                function pad(n) {{ return String(n).padStart(2, '0'); }}
                var elC = document.getElementById('tvClockLive');
                var elD = document.getElementById('tvDateLive');
                if (elC) {{ elC.textContent = pad(agora.getHours()) + ':' + pad(agora.getMinutes()) + ':' + pad(agora.getSeconds()); }}
                if (elD) {{
                    var dias = ['Domingo','Segunda-feira','Terça-feira','Quarta-feira','Quinta-feira','Sexta-feira','Sábado'];
                    elD.textContent = dias[agora.getDay()] + ', ' + pad(agora.getDate()) + '/' + pad(agora.getMonth() + 1) + '/' + agora.getFullYear();
                }}
            }}
            _tvAtualizarRelogio();
            setInterval(_tvAtualizarRelogio, 1000);
            </script>
            """,
            height=70,
        )

    def _tv_kpi(cor_var, label, valor, sub, accent):
        return (
            f'<div class="tv-kpi" style="--tv-accent:{accent};">'
            f'<div class="lbl">{label}</div>'
            f'<div class="val" style="color:{cor_var};">{valor}</div>'
            f'<div class="sub">{sub}</div></div>'
        )

    st.markdown(
        '<div class="tv-kpi-grid">'
        + _tv_kpi(COLORS["text"], "Receita Bruta (YTD)", formata_m(rec_bruta_real),
                  "Antes de deduções", COLORS["muted_line"])
        + _tv_kpi(cor_variacao(rec_liq_real), "Receita Líquida (YTD)", formata_m(rec_liq_real),
                  f"{pct_atingimento_rec:.1f}% do orçado", COLORS["primary"])
        + _tv_kpi(cor_variacao(ebitda_real), "EBITDA (YTD)", formata_m(ebitda_real),
                  f"{pct_atingimento_eb:.1f}% do orçado", COLORS["positive"])
        + _tv_kpi(cor_variacao(desvio_ebitda), "Desvio de EBITDA", formata_m(desvio_ebitda),
                  f"{pct_desvio_ebitda:+.1f}% vs. Orçamento", COLORS["warning"])
        + _tv_kpi(cor_variacao(margem_ebitda), "Margem EBITDA", f"{margem_ebitda:.1f}%",
                  f"Orçado: {margem_ebitda_orc:.1f}%", COLORS["secondary"])
        + "</div>",
        unsafe_allow_html=True,
    )

    # ---------------- Gauges de atingimento (estilo instrumento de painel) ----------------
    cg_gauge1, cg_gauge2 = st.columns(2)
    for col, (titulo, valor_real, valor_orc, pct_atg, cor_gauge) in zip(
        (cg_gauge1, cg_gauge2),
        [
            ("Atingimento de Receita Líquida", rec_liq_real, rec_liq_orc, pct_atingimento_rec, COLORS["primary"]),
            ("Atingimento de EBITDA", ebitda_real, ebitda_orc, pct_atingimento_eb, COLORS["positive"]),
        ],
    ):
        with col:
            teto_gauge = max(abs(valor_orc) * 1.3, abs(valor_real) * 1.15, 1.0)
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=valor_real,
                number={"valueformat": ",.0f", "prefix": "R$ ", "font": {"size": 26, "color": COLORS["text"]}},
                title={"text": f"{titulo} · {pct_atg:.0f}% do orçado", "font": {"size": 13, "color": COLORS["text_muted"]}},
                gauge={
                    "axis": {"range": [0, teto_gauge], "tickfont": {"size": 9, "color": COLORS["text_muted"]}},
                    "bar": {"color": cor_gauge, "thickness": 0.35},
                    "bgcolor": COLORS["surface_alt"],
                    "borderwidth": 0,
                    "threshold": {"line": {"color": COLORS["warning"], "width": 3}, "thickness": 0.9, "value": abs(valor_orc)},
                },
            ))
            estilo_grafico(fig_gauge, height=180, margin=dict(l=24, r=24, t=40, b=6), separators=",.")
            st.plotly_chart(fig_gauge, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    # ---------------- Carrega dados por loja p/ ranking (top performers) ----------------
    lojas_tv = [a for a in abas_disponiveis if a not in ABAS_CONSOLIDADAS_TV]
    dados_por_loja_tv = carregar_dados_por_loja(path_orc, path_real, lojas_tv) if lojas_tv else {}

    ranking_lojas = []
    for loja, (df_o_loja, df_r_loja) in dados_por_loja_tv.items():
        v_rec = get_valor_consolidado_multi([df_r_loja], "3 - Receita Operacional Liquida", cols_ytd)
        ranking_lojas.append((loja, v_rec))
    ranking_lojas.sort(key=lambda x: x[1], reverse=True)
    top_lojas = ranking_lojas[:6]
    max_rank_val = max((v for _, v in top_lojas), default=1.0) or 1.0

    cgtv1, cgtv2, cgtv3 = st.columns([1.25, 1, 1])

    with cgtv1:
        st.markdown('<div class="tv-section-title">📈 Evolução Mensal — Receita vs. EBITDA</div>', unsafe_allow_html=True)
        rec_m_tv, eb_m_tv, rot_m_tv = [], [], []
        for m_nome, c in m_map_tv.items():
            rec_m_tv.append(get_valor_consolidado_multi(list_df_real_tv, "3 - Receita Operacional Liquida", [c]))
            eb_m_tv.append(get_valor_consolidado_multi(list_df_real_tv, "11 - EBITDA", [c]))
            rot_m_tv.append(m_nome.capitalize()[:3])
        fig_tv_line = go.Figure()
        fig_tv_line.add_trace(go.Scatter(
            x=rot_m_tv, y=rec_m_tv, name="Receita Líquida", mode="lines+markers+text",
            text=[formata_m(v) for v in rec_m_tv], textposition="top center",
            textfont=dict(size=10, color=COLORS["primary"]),
            line=dict(color=COLORS["primary"], width=3), marker=dict(size=7),
        ))
        fig_tv_line.add_trace(go.Scatter(
            x=rot_m_tv, y=eb_m_tv, name="EBITDA", mode="lines+markers+text",
            text=[formata_m(v) for v in eb_m_tv], textposition="bottom center",
            textfont=dict(size=10, color=COLORS["positive"]),
            line=dict(color=COLORS["positive"], width=3, dash="dot"), marker=dict(size=7),
        ))
        estilo_grafico(
            fig_tv_line, height=320,
            xaxis=dict(showgrid=False, fixedrange=True, tickfont=dict(size=11, color=COLORS["text_muted"])),
            yaxis=dict(showgrid=False, showticklabels=False, fixedrange=True),
            legend=dict(orientation="h", yanchor="bottom", y=-0.24, xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_tv_line, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    with cgtv2:
        st.markdown('<div class="tv-section-title">🥧 Composição de Custos & Saídas</div>', unsafe_allow_html=True)
        cmv_tv = abs(get_valor_consolidado_multi(list_df_real_tv, "4 - ", cols_ytd, exato_linha_sintetica=True)) or \
            abs(get_valor_consolidado_multi(list_df_real_tv, "4 - Custo das Vendas", cols_ytd))
        desp_var_tv = abs(get_valor_consolidado_multi(list_df_real_tv, "6 - Despesas Variáveis", cols_ytd))
        deprec_tv = abs(get_valor_consolidado_multi(list_df_real_tv, "13 - Depreciação e Amortização", cols_ytd))
        total_tv = cmv_tv + desp_var_tv + desp_op_tv_kpi + deprec_tv
        fig_tv_donut = go.Figure(data=[go.Pie(
            labels=["CMV", "Desp. Variáveis", "Desp. Operacionais", "Deprec./Amort."],
            values=[cmv_tv, desp_var_tv, desp_op_tv_kpi, deprec_tv], hole=0.62,
            marker=dict(colors=[COLORS["primary"], COLORS["muted_line"], COLORS["secondary"], COLORS["border_soft"]],
                        line=dict(color=COLORS["surface"], width=2)),
            textinfo="percent", textfont=dict(size=11),
        )])
        fig_tv_donut.add_annotation(
            text=f"<b>{formata_m(total_tv)}</b><br><span style='font-size:10px;color:{COLORS['text_muted']}'>Total Saídas</span>",
            showarrow=False, font=dict(color=COLORS["text"], size=13, family=FONT_STACK),
        )
        estilo_grafico(fig_tv_donut, height=320, legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="center", x=0.5, font=dict(size=11)))
        st.plotly_chart(fig_tv_donut, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    with cgtv3:
        st.markdown('<div class="tv-section-title">🏆 Ranking de Lojas — Receita Líquida (YTD)</div>', unsafe_allow_html=True)
        if top_lojas:
            linhas_rank = ['<div class="tv-panel" style="padding-top:6px;">']
            for i, (loja, valor) in enumerate(top_lojas, start=1):
                pct_barra = max(2, min(100, (valor / max_rank_val * 100))) if max_rank_val else 2
                linhas_rank.append(
                    '<div class="tv-rank-row">'
                    f'<div class="tv-rank-badge">{i}</div>'
                    f'<div class="tv-rank-name" title="{loja}">{loja}</div>'
                    f'<div class="tv-rank-bar-bg"><div class="tv-rank-bar-fill" style="width:{pct_barra:.0f}%;"></div></div>'
                    f'<div class="tv-rank-val">{formata_m(valor)}</div>'
                    "</div>"
                )
            linhas_rank.append("</div>")
            st.markdown("".join(linhas_rank), unsafe_allow_html=True)
        else:
            st.info("Sem lojas individuais disponíveis para ranking.")

    # ---------------- Ticker de destaques (rodapé animado) ----------------
    meses_com_receita = {
        m_nome: get_valor_consolidado_multi(list_df_real_tv, "3 - Receita Operacional Liquida", [c])
        for m_nome, c in m_map_tv.items()
    }
    meses_validos = {m: v for m, v in meses_com_receita.items() if v != 0}
    destaques = []
    if meses_validos:
        melhor_mes = max(meses_validos, key=meses_validos.get)
        destaques.append(f"🏆 Melhor mês: <b>{melhor_mes.capitalize()}</b> ({formata_m(meses_validos[melhor_mes])})")
        if len(meses_validos) > 1:
            pior_mes = min(meses_validos, key=meses_validos.get)
            destaques.append(f"📉 Pior mês: <b>{pior_mes.capitalize()}</b> ({formata_m(meses_validos[pior_mes])})")
    if top_lojas:
        destaques.append(f"🥇 Loja destaque: <b>{top_lojas[0][0]}</b> ({formata_m(top_lojas[0][1])})")
        if len(top_lojas) > 1:
            destaques.append(f"🥈 2º lugar: <b>{top_lojas[1][0]}</b> ({formata_m(top_lojas[1][1])})")
    destaques.append(f"📊 Margem EBITDA: <b>{margem_ebitda:.1f}%</b> (orçado: {margem_ebitda_orc:.1f}%)")
    destaques.append(f"🎯 Atingimento de receita: <b>{pct_atingimento_rec:.1f}%</b>")
    destaques.append(f"💹 Atingimento de EBITDA: <b>{pct_atingimento_eb:.1f}%</b>")
    destaques.append(f"⚖️ Desvio de EBITDA: <b>{formata_brl(desvio_ebitda)}</b> ({pct_desvio_ebitda:+.1f}%)")
    destaques.append(f"💰 Receita bruta: <b>{formata_m(rec_bruta_real)}</b>")
    destaques.append(f"🧾 Saídas: <b>{formata_m(total_tv)}</b> (CMV {formata_m(cmv_tv)} · OpEx {formata_m(desp_op_tv_kpi)})")

    ticker_html = f'<span class="tv-tick-sep">·</span>'.join(destaques)
    st.markdown(
        f"""
        <div class="tv-ticker-wrap"><div class="tv-ticker">{ticker_html}</div></div>
        <div style="text-align:center;margin-top:8px;color:{COLORS['text_muted']};font-size:11px;">
            Painel para exibição (somente leitura) · Atualiza automaticamente a cada 30 minutos ·
            <a href="?" style="color:{COLORS['text_muted']};">Sair do modo TV</a>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Atualização automática SEM recarregar a página (evita perder a tela
    # cheia e evita cair de novo na tela de login a cada ciclo, como
    # acontecia com o <meta refresh> -- esse rerun acontece dentro da mesma
    # sessão/aba, sem navegação de página).
    time.sleep(1800)
    st.rerun()


if st.query_params.get("modo") == "tv":
    renderizar_painel_tv(path_orc, path_real, abas_disponiveis)
    st.stop()

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
# 5. BARRA LATERAL — FILTROS
# ============================================================================
st.sidebar.markdown(
    f"""
    <div class="sidebar-brand">
        <img class="brand-logo" src="data:image/jpeg;base64,{LOGO_BEEA_B64}" alt="Grupo Beea" />
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
    if not m_map:
        # Nenhum mês com dados válidos para o escopo selecionado (ex.: uma
        # loja/aba sem nenhuma coluna de mês reconhecida) -- evita quebrar o
        # app tentando montar um selectbox vazio.
        st.sidebar.warning("Nenhum mês com dados disponível para este escopo.")
        cols_kpi = []
        cols_graficos = []
        label_periodo_kpi = "Sem dados no escopo selecionado"
        label_periodo_graf = "Sem dados no escopo selecionado"
    else:
        # Abre por padrão no mês atual (real), não num índice fixo.
        idx_mes_real = datetime.now(FUSO_BR).month - 1
        idx_default = min(max(idx_mes_real, 0), len(m_map) - 1)
        mes_ref = st.sidebar.selectbox("Mês Desejado:", list(m_map.keys()), index=idx_default)
        if mes_ref not in m_map:
            # Segurança extra: se por qualquer motivo (ex.: opções do
            # selectbox mudaram entre execuções) o valor devolvido não
            # estiver mais no mapa atual, cai para o padrão em vez de
            # quebrar com ValueError.
            mes_ref = list(m_map.keys())[idx_default]
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
st.sidebar.markdown("**🖥️ Painel para TV**")
st.sidebar.caption("Abre uma visão executiva (somente leitura, atualização automática) para deixar exibida numa tela/TV da empresa.")
st.sidebar.markdown(
    f"""
    <a href="?modo=tv" target="_blank" style="text-decoration:none;">
        <div style="
            background:{COLORS['primary_soft']}; color:{COLORS['primary']}; border:1px solid {COLORS['primary']};
            border-radius:8px; padding:8px 12px; text-align:center; font-weight:600; font-size:13.5px;
            margin-bottom: 6px;">
            📺 Abrir Painel de TV (nova aba)
        </div>
    </a>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown("---")
perfil_label = "Administrador" if eh_admin else "Visualização"
st.sidebar.caption(f"👤 {usuario_atual['email']}  ·  Perfil: **{perfil_label}**")
if st.sidebar.button("🚪 Sair", use_container_width=True):
    st.session_state["usuario_logado"] = None
    _esquecer_credenciais_no_navegador()
    st.rerun()


# ============================================================================
# 6. FUNÇÕES DE SUPORTE (cálculo e formatação) -- ver seção 2 para as funções
# get_valor_consolidado_multi / formata_brl / formata_m / eh_grupo_sintetico /
# cor_valor, que foram movidas para perto do topo do arquivo (logo após
# render_kpi_row) para ficarem disponíveis também para o Painel de TV,
# que é montado antes desta seção.
# ============================================================================

# ============================================================================
# 6.1 GERAÇÃO DE RELATÓRIO EXCEL (formatado, para a aba de Emissão)
# ============================================================================
_THIN = Side(style="thin", color="FFD9DDE3")

EXCEL_STYLE = {
    "fill_title": PatternFill(fill_type="solid", start_color="FF0B0E14", end_color="FF0B0E14"),
    "fill_header": PatternFill(fill_type="solid", start_color="FF1A1F2E", end_color="FF1A1F2E"),
    "fill_zebra": PatternFill(fill_type="solid", start_color="FFF3F5F9", end_color="FFF3F5F9"),
    "fill_group": PatternFill(fill_type="solid", start_color="FFEFF3FA", end_color="FFEFF3FA"),
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


def _escrever_linha_flat(ws, linha, campos, valores_mes, total, negrito=False, fill=None, colorir_por_sinal=False):
    """Escreve uma linha em formato tabular "achatado": [campo_1, campo_2,
    ..., valor_mes_1, ..., total]. `campos` é uma lista de rótulos (ex.:
    [loja, conta, tipo]) -- diferente de _escrever_linha_matriz, que só
    aceita um único rótulo. Usado nas abas que precisam de AutoFilter do
    Excel (uma "loja" por coluna própria, não misturada dentro do texto)."""
    n_campos = len(campos)
    valores = list(campos) + list(valores_mes) + [total]
    for col, val in enumerate(valores, start=1):
        cell = ws.cell(row=linha, column=col, value=val)
        cell.border = EXCEL_STYLE["border"]
        if fill:
            cell.fill = fill
        if col <= n_campos:
            cell.font = EXCEL_STYLE["font_bold"] if negrito else EXCEL_STYLE["font_normal"]
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            valor_numerico = val if isinstance(val, (int, float)) else 0
            if colorir_por_sinal:
                cell.font = _fonte_por_valor(valor_numerico)
            else:
                cell.font = EXCEL_STYLE["font_bold"] if negrito else EXCEL_STYLE["font_normal"]
            cell.number_format = '"R$" #,##0.00'
            cell.alignment = Alignment(horizontal="right")


_CARACTERES_INVALIDOS_ABA = set('[]:*?/\\')


def _nome_aba_seguro(texto, usados, max_len=31):
    """Gera um nome de aba do Excel válido (até 31 caracteres, sem os
    caracteres proibidos pelo Excel) e garante que não colida com nenhum
    outro nome de aba já usado no mesmo workbook (guardados em `usados`,
    em minúsculas)."""
    limpo = "".join(c for c in str(texto) if c not in _CARACTERES_INVALIDOS_ABA).strip()
    limpo = limpo[:max_len] if limpo else "Aba"
    base = limpo
    sufixo = 1
    while limpo.lower() in usados:
        sufixo += 1
        marcador = f"_{sufixo}"
        corte = max_len - len(marcador)
        limpo = f"{base[:corte]}{marcador}"
    usados.add(limpo.lower())
    return limpo


def _ref_aba(nome):
    """Formata o nome de uma aba para uso em fórmulas do Excel (entre aspas
    simples, com aspas simples internas escapadas)."""
    return "'" + str(nome).replace("'", "''") + "'"


def _filtrar_diario_completo(df_diario, loja, conta, plano):
    """Filtra o DIÁRIO por loja (Centro de Custos) + Linha DRE (conta) +
    Plano de Contas exato -- usado para montar a aba de detalhamento dos
    lançamentos ao "clicar" numa linha da aba Plano de Contas."""
    df_filtrado = _filtrar_diario_por_loja_e_conta(df_diario, loja, conta)
    if df_filtrado is None or df_filtrado.empty:
        return df_filtrado
    plano_norm = _normalizar_texto(plano)
    planos_norm = df_filtrado["Plano de Contas"].map(_normalizar_texto)
    mask = planos_norm == plano_norm
    if not mask.any():
        mask = planos_norm.apply(lambda x: (plano_norm in x) or (x in plano_norm) if x else False)
    return df_filtrado[mask]


def _criar_aba_lancamentos(wb, nomes_usados, titulo_bloco, df_lancamentos):
    """Cria uma aba nova com o detalhamento dos lançamentos do DIÁRIO que
    compõem um valor (usada para o "clique" que abre o detalhamento a partir
    da aba Detalhe Mensal ou Plano de Contas). É basicamente uma cópia da
    própria aba DIÁRIO, só que filtrada para aqueles lançamentos específicos
    -- por isso mantém as mesmas colunas dela (Competência, Centro de Custos,
    Linha DRE, Plano de Contas, Valor Bruto). Retorna o nome da aba criada."""
    nome_aba = _nome_aba_seguro(f"Lc_{titulo_bloco}", nomes_usados)
    ws = wb.create_sheet(nome_aba)
    ws.sheet_properties.tabColor = "FF4C8DFF"

    n_col = 5
    _escrever_titulo(ws, f"Lançamentos — {titulo_bloco}"[:250], 1, n_col)
    _escrever_legenda(ws, f"{len(df_lancamentos)} lançamento(s) encontrados na aba DIÁRIO (Realizado 2026).", 2, n_col)

    linha = 4
    headers = ["Competência", "Centro de Custos", "Linha DRE", "Plano de Contas", "Valor Bruto (R$)"]
    for col, texto in enumerate(headers, start=1):
        cell = ws.cell(row=linha, column=col, value=texto)
        cell.font = EXCEL_STYLE["font_header"]
        cell.fill = EXCEL_STYLE["fill_header"]
        cell.border = EXCEL_STYLE["border"]
        cell.alignment = Alignment(horizontal="left" if col in (1, 2, 3, 4) else "center")
    linha += 1

    df_ordenado = df_lancamentos.sort_values("Competência", na_position="last")
    for i, (_, row) in enumerate(df_ordenado.iterrows()):
        data_val = row["Competência"]
        data_txt = data_val.strftime("%d/%m/%Y") if pd.notna(data_val) and hasattr(data_val, "strftime") else str(data_val)
        valores = [data_txt, row["Centro de Custos"], row["Linha DRE"], row["Plano de Contas"], row["Valor Bruto"]]
        fill = EXCEL_STYLE["fill_zebra"] if i % 2 == 1 else None
        for col, val in enumerate(valores, start=1):
            cell = ws.cell(row=linha, column=col, value=val)
            cell.border = EXCEL_STYLE["border"]
            if fill:
                cell.fill = fill
            if col == 5:
                cell.font = EXCEL_STYLE["font_normal"]
                cell.number_format = '"R$" #,##0.00'
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.font = EXCEL_STYLE["font_normal"]
                cell.alignment = Alignment(horizontal="left")
        linha += 1

    total_valor = df_ordenado["Valor Bruto"].sum()
    cell_total_lbl = ws.cell(row=linha, column=1, value="TOTAL")
    cell_total_lbl.font = EXCEL_STYLE["font_bold"]
    cell_total = ws.cell(row=linha, column=5, value=total_valor)
    cell_total.font = EXCEL_STYLE["font_bold"]
    cell_total.number_format = '"R$" #,##0.00'
    cell_total.alignment = Alignment(horizontal="right")
    for col in range(1, n_col + 1):
        ws.cell(row=linha, column=col).fill = EXCEL_STYLE["fill_total"]
        ws.cell(row=linha, column=col).border = EXCEL_STYLE["border"]

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 34
    ws.column_dimensions["E"].width = 18
    ws.freeze_panes = "A5"
    return nome_aba


def montar_relatorio_excel(
    contas_sel,
    dfs_real,
    dfs_orc,
    mapa_meses,
    colunas_ano,
    escopo_label,
    dados_por_loja=None,
    mapa_planos_dre=None,
    df_diario=None,
    forcar_planos_contas=None,
    permitir_lancamento_manual=False,
    mapa_loja_centro_custo=None,
):
    """Gera um relatório Excel formatado com três planilhas:
    - Resumo: total do ano por conta, CONSOLIDADO (não muda com a divisão por loja).
    - Detalhe Mensal: contas da DRE nas linhas, meses nas colunas, dividido por loja.
    - Plano de Contas: composição (planos de contas) de cada linha da DRE, por loja,
      mês a mês. Fonte principal: aba DIÁRIO da planilha Realizado 2026 (df_diario).
      Se o DIÁRIO não estiver disponível para uma loja/conta, cai para o método
      antigo (Tabela_Contas + soma nas abas por loja).

    forcar_planos_contas: lista de nomes de Plano de Contas que devem ser
    incluídos na aba "Plano de Contas" à parte, puxados direto pelo nome
    (ignorando a Linha DRE) -- usado por modelos como o de Compras, onde
    "Mercadorias" faz parte do relatório mesmo sem linha da DRE correspondente.

    permitir_lancamento_manual: quando True, uma Linha DRE sem nenhum Plano
    de Contas correspondente no DIÁRIO/Tabela_Contas aparece como
    "Lançado Manualmente", com o valor da própria linha da DRE -- usado pelo
    modelo de RH, que tem várias linhas lançadas direto, sem plano de contas.

    mapa_loja_centro_custo: {nome_da_loja: centro_de_custo}, vindo da
    Tabela_Lojas (DE_PARA). A DIÁRIO identifica a loja pelo Centro de Custos,
    que normalmente é diferente do nome da aba da loja -- por isso, sempre
    que formos filtrar a DIÁRIO por loja, resolvemos primeiro o nome da loja
    para o Centro de Custos correspondente por esse mapa (usando o próprio
    nome da loja como alternativa, caso não haja correspondência no DE_PARA).
    """
    dados_por_loja = dados_por_loja or {}
    mapa_planos_dre = mapa_planos_dre or {}
    forcar_planos_contas = forcar_planos_contas or []
    mapa_loja_centro_custo = mapa_loja_centro_custo or {}
    diario_disponivel = df_diario is not None and not df_diario.empty

    def _cc(loja):
        """Resolve os candidatos de Centro de Custos para uma loja: o valor
        mapeado pela Tabela_Lojas (DE_PARA) E o próprio nome da loja/aba,
        nessa ordem. A busca no DIÁRIO tenta os dois, então funciona mesmo
        que o DE_PARA esteja incompleto/não encontrado para alguma loja, ou
        caso o Centro de Custos já seja igual ao nome da loja em alguns
        casos."""
        centro_mapeado = mapa_loja_centro_custo.get(loja)
        candidatos = [c for c in (centro_mapeado, loja) if c]
        return candidatos or [loja]
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

    # ========================================================================
    # Dados auxiliares: lojas individuais x visões consolidadas, e o filtro
    # único da DIÁRIO que alimenta tanto a aba "Lançamentos" quanto a aba
    # "Plano de Contas".
    # ========================================================================
    lojas_ordenadas = sorted(dados_por_loja.keys()) if dados_por_loja else []
    meses_chaves = list(mapa_meses.keys())
    n_meses = len(meses_chaves)
    n_col_matriz = 1 + n_meses + 1  # rótulo + meses + Total Ano

    LOJAS_CONSOLIDADAS = {
        "DRE CONSOLIDADO", "ABPR CONSOLIDADO", "VD CONSOLIDADO",
        "LJ CONSOLIDADO", "ABPR + VD", "LJ - G&A",
    }
    lojas_individuais = [l for l in lojas_ordenadas if l not in LOJAS_CONSOLIDADAS]

    def _lojas_do_grupo_consolidado(nome_grupo):
        """Para uma "loja" que na verdade é uma visão consolidada (ex.: "ABPR
        + VD"), devolve as lojas INDIVIDUAIS que fazem parte dela -- usado
        para somar corretamente a composição de Plano de Contas (que vem da
        DIÁRIO, que só reconhece lojas individuais, nunca uma visão
        consolidada). Grupos definidos explicitamente:
        - ABPR CONSOLIDADO = ABPR 23427 + ABPR ZNS 24527
        - VD CONSOLIDADO = VD - GUAJARA 23859 + VD - LESTE 21506 + VD - MACHAD 21691 +
          VD - MATRIZ 13967 + VD - VST ALEG 21497
        - ABPR + VD = ABPR CONSOLIDADO + VD CONSOLIDADO
        - LJ - G&A = as 13 lojas "LJ ..." (sem ESCRIT MATRIZ, sem lojas VD)
        - LJ CONSOLIDADO = ESCRIT MATRIZ 6037 + LJ - G&A (sem lojas VD)
        - DRE CONSOLIDADO = ABPR CONSOLIDADO + VD CONSOLIDADO + LJ CONSOLIDADO
          (na prática, todas as 21 lojas individuais)
        """
        grupo_abpr = ["ABPR 23427", "ABPR ZNS 24527"]
        grupo_vd = [
            "VD - GUAJARA 23859", "VD - LESTE 21506", "VD - MACHAD 21691",
            "VD - MATRIZ 13967", "VD - VST ALEG 21497",
        ]
        grupo_lj_ga = [
            "LJ ARAUJO 12606", "LJ ASSAI 23157", "LJ GUAJARA 23809", "LJ IG SHOP 20330",
            "LJ JATUARANA 6040", "LJ JK 12478", "LJ MACHAD 21462", "LJ MARECHAL 6039",
            "LJ NV ERA 18539", "LJ PVH1 11927", "LJ PVH2 14625", "LJ QDB 910332", "LJ SETE 6052",
        ]
        grupo_lj_consolidado = ["ESCRIT MATRIZ 6037"] + grupo_lj_ga
        grupo_dre_consolidado = grupo_abpr + grupo_vd + grupo_lj_consolidado

        mapa_grupos = {
            "ABPR CONSOLIDADO": grupo_abpr,
            "VD CONSOLIDADO": grupo_vd,
            "ABPR + VD": grupo_abpr + grupo_vd,
            "LJ - G&A": grupo_lj_ga,
            "LJ CONSOLIDADO": grupo_lj_consolidado,
            "DRE CONSOLIDADO": grupo_dre_consolidado,
        }
        lojas_definidas = mapa_grupos.get(nome_grupo.strip(), [])
        # Só devolve lojas que realmente existem entre as lojas individuais carregadas.
        return [l for l in lojas_definidas if l in lojas_individuais]


    # ---- Filtra a DIÁRIO para pegar só os lançamentos relevantes deste
    # relatório: linhas cuja "Linha DRE" bate com alguma conta selecionada,
    # OU cujo "Plano de Contas" bate com um dos planos forçados (ex.:
    # "Mercadorias", que não tem Linha DRE). Essa é a ÚNICA fonte usada tanto
    # pela aba "Lançamentos" quanto para montar a composição da aba
    # "Plano de Contas" -- ou seja, os dois batem entre si por construção. ----
    if diario_disponivel:
        linhas_norm = df_diario["Linha DRE"].map(_normalizar_texto)
        mask_conta = pd.Series(False, index=df_diario.index)
        for conta in contas_sel:
            conta_norm = _normalizar_texto(conta)
            m = linhas_norm == conta_norm
            if not m.any():
                m = linhas_norm.apply(lambda x: (conta_norm in x) or (x in conta_norm) if x else False)
            mask_conta |= m

        planos_norm = df_diario["Plano de Contas"].map(_normalizar_texto)
        mask_forcado = pd.Series(False, index=df_diario.index)
        for plano in forcar_planos_contas:
            plano_norm = _normalizar_texto(plano)
            m = planos_norm == plano_norm
            if not m.any():
                m = planos_norm.apply(lambda x: (plano_norm in x) or (x in plano_norm) if x else False)
            mask_forcado |= m

        df_lanc = df_diario[mask_conta | mask_forcado].copy()
        # Resolve o nome "bonito" da loja a partir do Centro de Custos (usa o
        # mapa nos dois sentidos, montado pela Tabela_Lojas/DE_PARA). Se não
        # achar correspondência, mantém o próprio Centro de Custos como
        # rótulo (mais transparente do que esconder a linha).
        df_lanc["Loja"] = df_lanc["Centro de Custos"].astype(str).str.strip().map(
            lambda cc: mapa_loja_centro_custo.get(cc, cc)
        )
    else:
        df_lanc = pd.DataFrame(
            columns=[
                "Competência", "Centro de Custos", "Linha DRE", "Plano de Contas", "Valor Bruto", "Mês",
                "Número", "Cliente / Fornecedor", "Histórico", "Loja",
            ]
        )

    # ---- Planos de contas GLOBAIS por conta (juntando TODAS as lojas) ----
    # A aba "Plano de Contas" precisa mostrar sempre o MESMO conjunto de
    # planos de contas para uma dada linha da DRE, em todas as lojas -- com
    # valor ou não. Sem isso, uma loja que não tem lançamento pra um plano
    # específico simplesmente não mostrava aquela linha, enquanto outra loja
    # (que tem) mostrava -- deixando a estrutura inconsistente entre visões.
    planos_por_conta_global = {}
    if not df_lanc.empty:
        linhas_norm_geral = df_lanc["Linha DRE"].map(_normalizar_texto)
        for conta in contas_sel:
            conta_norm_geral = _normalizar_texto(conta)
            mask_geral = linhas_norm_geral == conta_norm_geral
            if not mask_geral.any():
                mask_geral = linhas_norm_geral.apply(
                    lambda x: (conta_norm_geral in x) or (x in conta_norm_geral) if x else False
                )
            planos_conta = sorted(df_lanc.loc[mask_geral, "Plano de Contas"].dropna().unique().tolist())
            if planos_conta:
                planos_por_conta_global[conta] = planos_conta

    # ---------------- ABA "DETALHE MENSAL" (Orçado x Realizado, por loja) ----------------
    # Formato tabular "achatado" (Loja/Conta/Tipo em colunas próprias, uma
    # linha por combinação) -- isso permite usar o AutoFilter nativo do
    # Excel na coluna "Loja" para escolher qual visão ver, sem precisar abrir
    # um arquivo gigante com tudo misturado. Clique na setinha do cabeçalho
    # da coluna "Loja" (linha 4) para filtrar.
    ws2 = wb.create_sheet("Detalhe Mensal")
    n_campos_dm = 3  # Loja, Conta, Tipo
    n_col_dm = n_campos_dm + n_meses + 1
    _escrever_titulo(ws2, "Detalhe Mensal — Orçado vs. Realizado, por Loja", 1, n_col_dm)
    _escrever_legenda(
        ws2,
        f"{gerado_em} · Use o filtro (▾) no cabeçalho da coluna \"Loja\" para escolher qual visão ver.",
        2, n_col_dm,
    )

    linha_header_dm = 4
    headers_dm = ["Loja", "Conta", "Tipo"] + [m.capitalize() for m in mapa_meses.keys()] + ["Total Ano"]
    for col, texto in enumerate(headers_dm, start=1):
        cell = ws2.cell(row=linha_header_dm, column=col, value=texto)
        cell.font = EXCEL_STYLE["font_header"]
        cell.fill = EXCEL_STYLE["fill_header"]
        cell.border = EXCEL_STYLE["border"]
        cell.alignment = Alignment(horizontal="left" if col <= n_campos_dm else "center")
    ws2.row_dimensions[linha_header_dm].height = 20
    ws2.freeze_panes = f"D{linha_header_dm + 1}"

    linha = linha_header_dm + 1
    if not lojas_ordenadas:
        ws2.cell(row=linha, column=1, value="Nenhuma loja/unidade disponível para divisão.").font = EXCEL_STYLE["font_normal"]
        linha += 1
    for loja in lojas_ordenadas:
        df_o_loja, df_r_loja = dados_por_loja[loja]

        for conta in contas_sel:
            valores_real = [get_valor_consolidado_multi([df_r_loja], conta, [m_col]) for m_col in mapa_meses.values()]
            valores_orc = [get_valor_consolidado_multi([df_o_loja], conta, [m_col]) for m_col in mapa_meses.values()]
            valores_desvio = [vr - vo for vr, vo in zip(valores_real, valores_orc)]

            _escrever_linha_flat(ws2, linha, [loja, conta, "Realizado"], valores_real, sum(valores_real))
            linha += 1
            _escrever_linha_flat(ws2, linha, [loja, conta, "Orçado"], valores_orc, sum(valores_orc), fill=EXCEL_STYLE["fill_zebra"])
            linha += 1
            _escrever_linha_flat(ws2, linha, [loja, conta, "Desvio"], valores_desvio, sum(valores_desvio), colorir_por_sinal=True)
            linha += 1

    ultima_linha_dm = linha - 1
    largura_mes = 14
    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 40
    ws2.column_dimensions["C"].width = 12
    for col in range(n_campos_dm + 1, n_col_dm + 1):
        ws2.column_dimensions[get_column_letter(col)].width = largura_mes
    if ultima_linha_dm >= linha_header_dm + 1:
        ws2.auto_filter.ref = f"A{linha_header_dm}:{get_column_letter(n_col_dm)}{ultima_linha_dm}"

    # ---------------- ABA "PLANO DE CONTAS" (composição de cada linha da DRE, por loja) ----------------
    # Mesmo formato tabular achatado, com AutoFilter na coluna "Loja".
    ws3 = wb.create_sheet("Plano de Contas")
    n_campos_pc = 3  # Loja, Conta, Plano de Contas
    n_col_pc = n_campos_pc + n_meses + 1
    _escrever_titulo(ws3, "Plano de Contas — Composição das Linhas da DRE, por Loja", 1, n_col_pc)
    _escrever_legenda(
        ws3,
        f"{gerado_em} · Valores de cada plano de contas puxados da DIÁRIO (Realizado 2026), pela loja/Centro "
        f"de Custos. Visões consolidadas somam as lojas individuais correspondentes. Use o filtro (▾) no "
        f"cabeçalho da coluna \"Loja\" para escolher qual visão ver.",
        2, n_col_pc,
    )

    linha_header_pc = 4
    headers_pc = ["Loja", "Conta", "Plano de Contas"] + [m.capitalize() for m in mapa_meses.keys()] + ["Total Ano"]
    for col, texto in enumerate(headers_pc, start=1):
        cell = ws3.cell(row=linha_header_pc, column=col, value=texto)
        cell.font = EXCEL_STYLE["font_header"]
        cell.fill = EXCEL_STYLE["fill_header"]
        cell.border = EXCEL_STYLE["border"]
        cell.alignment = Alignment(horizontal="left" if col <= n_campos_pc else "center")
    ws3.row_dimensions[linha_header_pc].height = 20
    ws3.freeze_panes = f"D{linha_header_pc + 1}"

    linha = linha_header_pc + 1
    if not lojas_ordenadas:
        ws3.cell(row=linha, column=1, value="Nenhuma loja/unidade disponível para divisão.").font = EXCEL_STYLE["font_normal"]
        linha += 1

    indice_grupo = 0
    for loja in lojas_ordenadas:
        df_o_loja, df_r_loja = dados_por_loja[loja]

        # Lançamentos desta loja -- ou, se for uma visão consolidada, de
        # todas as lojas individuais que fazem parte dela.
        if loja in LOJAS_CONSOLIDADAS:
            lojas_do_grupo = _lojas_do_grupo_consolidado(loja)
            df_lanc_loja = df_lanc[df_lanc["Loja"].isin(lojas_do_grupo)] if lojas_do_grupo else df_lanc.iloc[0:0]
        else:
            df_lanc_loja = df_lanc[df_lanc["Loja"] == loja]

        for conta in contas_sel:
            # Cada conta (linha da DRE) forma um "grupo" -- as linhas de
            # plano de contas dela, mais o TOTAL, recebem o mesmo
            # sombreamento leve alternado, pra ficar visualmente claro onde
            # um grupo termina e o próximo começa (em vez de tudo num fundo
            # branco igual, que se perde fácil numa planilha grande).
            indice_grupo += 1
            fill_grupo = EXCEL_STYLE["fill_group"] if indice_grupo % 2 == 1 else None
            df_diario_conta = pd.DataFrame()
            if not df_lanc_loja.empty:
                conta_norm = _normalizar_texto(conta)
                linhas_norm_loja = df_lanc_loja["Linha DRE"].map(_normalizar_texto)
                mask_conta_loja = linhas_norm_loja == conta_norm
                if not mask_conta_loja.any():
                    mask_conta_loja = linhas_norm_loja.apply(lambda x: (conta_norm in x) or (x in conta_norm) if x else False)
                df_diario_conta = df_lanc_loja[mask_conta_loja]

            soma_planos_mes = [0.0] * n_meses
            planos_globais_conta = planos_por_conta_global.get(conta)
            if planos_globais_conta:
                # Fonte principal: DIÁRIO/Lançamentos. Usa o conjunto GLOBAL
                # de planos de contas dessa linha da DRE (calculado juntando
                # todas as lojas) -- assim TODA loja mostra as mesmas linhas
                # de plano de contas, mesmo que, para essa loja específica,
                # o valor seja zero (sem lançamento).
                for plano in planos_globais_conta:
                    grupo = df_diario_conta[df_diario_conta["Plano de Contas"] == plano] if not df_diario_conta.empty else df_diario_conta
                    valores_plano = (
                        [grupo.loc[grupo["Mês"] == m_col, "Valor Bruto"].sum() for m_col in mapa_meses.values()]
                        if not grupo.empty else [0.0] * n_meses
                    )
                    soma_planos_mes = [s + v for s, v in zip(soma_planos_mes, valores_plano)]
                    _escrever_linha_flat(ws3, linha, [loja, conta, plano], valores_plano, sum(valores_plano), fill=fill_grupo)
                    linha += 1
                n_planos_escritos = len(planos_globais_conta)
            elif permitir_lancamento_manual and not (mapa_planos_dre.get(str(conta).strip(), [])):
                # Linha DRE sem nenhum Plano de Contas correspondente (nem no
                # DIÁRIO, nem na Tabela_Contas) -- típico do modelo de RH.
                valores_manual = [
                    get_valor_consolidado_multi([df_r_loja], conta, [m_col]) for m_col in mapa_meses.values()
                ]
                soma_planos_mes = valores_manual
                _escrever_linha_flat(ws3, linha, [loja, conta, "Lançado Manualmente"], valores_manual, sum(valores_manual), fill=fill_grupo)
                linha += 1
                n_planos_escritos = 1
            else:
                # Fallback: Tabela_Contas + soma nas abas por loja (método antigo).
                planos = mapa_planos_dre.get(str(conta).strip(), []) or [conta]
                conta_norm_cmp = _normalizar_texto(conta)
                for plano in planos:
                    valores_plano = [
                        get_valor_consolidado_multi([df_r_loja], plano, [m_col]) for m_col in mapa_meses.values()
                    ]
                    soma_planos_mes = [s + v for s, v in zip(soma_planos_mes, valores_plano)]
                    # Quando não há um Plano de Contas distinto de verdade
                    # (caiu no fallback "ou [conta]", ou a Tabela_Contas só
                    # tem o próprio nome da linha da DRE como "plano"), o
                    # valor já vem certo, mas o rótulo repetia a conta --
                    # mostramos "Lançado Manualmente" para deixar claro que
                    # esse valor foi lançado direto na linha da DRE, sem
                    # composição por plano de contas. Vale para qualquer
                    # modelo de relatório, não só o de RH.
                    rotulo_plano = "Lançado Manualmente" if _normalizar_texto(plano) == conta_norm_cmp else plano
                    _escrever_linha_flat(ws3, linha, [loja, conta, rotulo_plano], valores_plano, sum(valores_plano), fill=fill_grupo)
                    linha += 1
                n_planos_escritos = len(planos)

            # A linha de TOTAL só faz sentido (e só é escrita) quando há MAIS
            # DE UM plano de contas para essa linha da DRE -- com um único
            # plano, o total seria idêntico à própria linha, repetindo o
            # mesmo valor à toa. O rótulo cita a própria conta ("TOTAL —
            # <conta>") para deixar explícito a quais linhas de plano de
            # contas, logo acima, aquele total se refere -- em vez de
            # depender só da cor de fundo pra não se perder na leitura.
            if n_planos_escritos > 1:
                _escrever_linha_flat(
                    ws3, linha, [loja, conta, f"TOTAL — {conta}"], soma_planos_mes, sum(soma_planos_mes),
                    negrito=True, fill=EXCEL_STYLE["fill_total"],
                )
                linha += 1

        if forcar_planos_contas:
            # Planos de contas que fazem parte do modelo mas não têm Linha DRE
            # correspondente (ex.: "Mercadorias" no modelo de Compras) --
            # puxados direto da DIÁRIO/Lançamentos pelo nome do plano,
            # ignorando a Linha DRE.
            indice_grupo += 1
            fill_grupo_forcado = EXCEL_STYLE["fill_group"] if indice_grupo % 2 == 1 else None
            soma_forcados_mes = [0.0] * n_meses
            for plano_forcado in forcar_planos_contas:
                df_plano_forcado = pd.DataFrame()
                if not df_lanc_loja.empty:
                    plano_norm = _normalizar_texto(plano_forcado)
                    planos_norm_loja = df_lanc_loja["Plano de Contas"].map(_normalizar_texto)
                    mask_plano = planos_norm_loja == plano_norm
                    if not mask_plano.any():
                        mask_plano = planos_norm_loja.apply(lambda x: (plano_norm in x) or (x in plano_norm) if x else False)
                    df_plano_forcado = df_lanc_loja[mask_plano]

                valores_plano = (
                    [df_plano_forcado.loc[df_plano_forcado["Mês"] == m_col, "Valor Bruto"].sum() for m_col in mapa_meses.values()]
                    if not df_plano_forcado.empty else [0.0] * n_meses
                )
                soma_forcados_mes = [s + v for s, v in zip(soma_forcados_mes, valores_plano)]
                _escrever_linha_flat(
                    ws3, linha, [loja, "(Fora da DRE)", plano_forcado], valores_plano, sum(valores_plano),
                    fill=fill_grupo_forcado,
                )
                linha += 1

            if len(forcar_planos_contas) > 1:
                _escrever_linha_flat(
                    ws3, linha, [loja, "(Fora da DRE)", "TOTAL — Fora da DRE"], soma_forcados_mes, sum(soma_forcados_mes),
                    negrito=True, fill=EXCEL_STYLE["fill_total"],
                )
                linha += 1

    ultima_linha_pc = linha - 1
    ws3.column_dimensions["A"].width = 26
    ws3.column_dimensions["B"].width = 40
    ws3.column_dimensions["C"].width = 42
    for col in range(n_campos_pc + 1, n_col_pc + 1):
        ws3.column_dimensions[get_column_letter(col)].width = largura_mes
    if ultima_linha_pc >= linha_header_pc + 1:
        ws3.auto_filter.ref = f"A{linha_header_pc}:{get_column_letter(n_col_pc)}{ultima_linha_pc}"

    # ---------------- ABA "LANÇAMENTOS" (cópia filtrada da DIÁRIO) ----------------
    ws4 = wb.create_sheet("Lançamentos")
    N_COL_LANC = 9
    _escrever_titulo(ws4, "Lançamentos — Cópia filtrada da aba DIÁRIO (Realizado 2026)", 1, N_COL_LANC)
    _escrever_legenda(
        ws4,
        f"{gerado_em} · {len(df_lanc)} lançamento(s): linhas cujo Plano de Contas pertence às linhas "
        f"da DRE deste relatório, ou aos planos adicionais do modelo (ex.: Mercadorias).",
        2, N_COL_LANC,
    )

    linha = 4
    headers_lanc = [
        "Loja", "Competência", "Número", "Centro de Custos", "Linha DRE",
        "Plano de Contas", "Cliente / Fornecedor", "Histórico", "Valor Bruto (R$)",
    ]
    col_valor_lanc = len(headers_lanc)
    for col, texto in enumerate(headers_lanc, start=1):
        cell = ws4.cell(row=linha, column=col, value=texto)
        cell.font = EXCEL_STYLE["font_header"]
        cell.fill = EXCEL_STYLE["fill_header"]
        cell.border = EXCEL_STYLE["border"]
        cell.alignment = Alignment(horizontal="left" if col < col_valor_lanc else "center")
    ws4.row_dimensions[linha].height = 20
    ws4.freeze_panes = f"A{linha + 1}"
    linha += 1

    if not df_lanc.empty:
        df_lanc_ordenado = df_lanc.sort_values(["Loja", "Competência"], na_position="last")
        for i, (_, row) in enumerate(df_lanc_ordenado.iterrows()):
            data_val = row["Competência"]
            data_txt = data_val.strftime("%d/%m/%Y") if pd.notna(data_val) and hasattr(data_val, "strftime") else str(data_val)
            valores = [
                row["Loja"], data_txt, row.get("Número", ""), row["Centro de Custos"], row["Linha DRE"],
                row["Plano de Contas"], row.get("Cliente / Fornecedor", ""), row.get("Histórico", ""), row["Valor Bruto"],
            ]
            fill = EXCEL_STYLE["fill_zebra"] if i % 2 == 1 else None
            for col, val in enumerate(valores, start=1):
                cell = ws4.cell(row=linha, column=col, value=val)
                cell.border = EXCEL_STYLE["border"]
                if fill:
                    cell.fill = fill
                if col == col_valor_lanc:
                    cell.font = EXCEL_STYLE["font_normal"]
                    cell.number_format = '"R$" #,##0.00'
                    cell.alignment = Alignment(horizontal="right")
                else:
                    cell.font = EXCEL_STYLE["font_normal"]
                    cell.alignment = Alignment(horizontal="left")
            linha += 1

        # AutoFilter nativo do Excel -- cobre só cabeçalho + linhas de dados
        # (não a linha de TOTAL, escrita logo abaixo).
        ws4.auto_filter.ref = f"A4:{get_column_letter(N_COL_LANC)}{linha - 1}"

        total_lanc = df_lanc["Valor Bruto"].sum()
        cell_total_lbl = ws4.cell(row=linha, column=1, value="TOTAL")
        cell_total_lbl.font = EXCEL_STYLE["font_bold"]
        cell_total = ws4.cell(row=linha, column=col_valor_lanc, value=total_lanc)
        cell_total.font = EXCEL_STYLE["font_bold"]
        cell_total.number_format = '"R$" #,##0.00'
        cell_total.alignment = Alignment(horizontal="right")
        for col in range(1, N_COL_LANC + 1):
            ws4.cell(row=linha, column=col).fill = EXCEL_STYLE["fill_total"]
            ws4.cell(row=linha, column=col).border = EXCEL_STYLE["border"]
    else:
        ws4.cell(
            row=linha, column=1,
            value="Nenhum lançamento encontrado (DIÁRIO indisponível ou sem correspondência para os filtros deste relatório).",
        ).font = EXCEL_STYLE["font_normal"]

    ws4.column_dimensions["A"].width = 24
    ws4.column_dimensions["B"].width = 14
    ws4.column_dimensions["C"].width = 14
    ws4.column_dimensions["D"].width = 20
    ws4.column_dimensions["E"].width = 30
    ws4.column_dimensions["F"].width = 30
    ws4.column_dimensions["G"].width = 28
    ws4.column_dimensions["H"].width = 40
    ws4.column_dimensions["I"].width = 18

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================================
# 7. FAIXA DE CONTEXTO (fina — não repete um bloco grande em toda aba)
# ============================================================================
st.markdown(
    f"""
    <div class="top-status-strip">
        <span class="chip">{label_visao}</span>
        <span class="sep">·</span>
        <span>Período: <b>{label_periodo_kpi}</b></span>
        <span class="sep">·</span>
        <span>Controladoria B&amp;A · Painel Financeiro 2026</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================================
# 8. ABAS
# ============================================================================
_nomes_abas = [
    "📊 Visão Geral & Charts",
    "📋 DRE Orçado X Realizado",
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
    st.markdown(
        '<div class="section-title">📊 Visão Geral & Charts — Indicadores Executivos e Composição do Resultado</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

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
    st.markdown(f'<div class="section-title">📋 Análise de DRE e Desvios — {label_visao} · {label_periodo_graf}</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    c1, _ = st.columns([2, 1])
    with c1:
        tipo_visao_dre = st.radio(
            "Filtro de Nível de Visão:",
            [
                "Apenas Grupos Principais (Sintética)",
                "Todas as Contas (Analítica)",
                "Visão Gerencial (Custos e Despesas)",
            ],
            horizontal=True,
        )

    col_nome = "Nome" if "Nome" in df_ref.columns else df_ref.columns[0]
    linhas_dre = df_ref[col_nome].dropna().astype(str).unique()

    is_sintetica_dre = tipo_visao_dre == "Apenas Grupos Principais (Sintética)"
    is_gerencial_dre = tipo_visao_dre == "Visão Gerencial (Custos e Despesas)"
    if is_sintetica_dre:
        linhas_dre = [l for l in linhas_dre if eh_grupo_sintetico(l)]
    elif is_gerencial_dre:
        linhas_dre = [l for l in linhas_dre if eh_linha_custos_despesas(l)]

    if not is_sintetica_dre and not is_gerencial_dre:
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

    if is_gerencial_dre:
        # Visão rápida e completa de custos e despesas: CMV, Despesas
        # Variáveis, Despesas Operacionais e o total dos três, cada um com
        # % sobre a Receita Líquida e o desvio vs. orçado -- é o resumo que
        # a visão Gerencial existe pra dar de cara, antes da tabela detalhada.
        rec_liq_ger = get_valor_consolidado_multi(list_df_real, "3 - Receita Operacional Liquida", cols_graficos)

        cmv_ger_r = abs(get_valor_consolidado_multi(list_df_real, "4 - ", cols_graficos, exato_linha_sintetica=True)) \
            or abs(get_valor_consolidado_multi(list_df_real, "4 - Custo das Vendas", cols_graficos))
        cmv_ger_o = abs(get_valor_consolidado_multi(list_df_orc, "4 - ", cols_graficos, exato_linha_sintetica=True)) \
            or abs(get_valor_consolidado_multi(list_df_orc, "4 - Custo das Vendas", cols_graficos))

        dvar_ger_r = abs(get_valor_consolidado_multi(list_df_real, "6 - Despesas Variáveis", cols_graficos))
        dvar_ger_o = abs(get_valor_consolidado_multi(list_df_orc, "6 - Despesas Variáveis", cols_graficos))

        dop_ger_r = abs(get_valor_consolidado_multi(list_df_real, "8 - Despesas Operacionais", cols_graficos))
        dop_ger_o = abs(get_valor_consolidado_multi(list_df_orc, "8 - Despesas Operacionais", cols_graficos))

        total_ger_r = cmv_ger_r + dvar_ger_r + dop_ger_r
        total_ger_o = cmv_ger_o + dvar_ger_o + dop_ger_o

        def _pct_receita(valor):
            return (valor / rec_liq_ger * 100) if rec_liq_ger else 0.0

        st.markdown('<div class="section-title">💸 Resumo Gerencial — Custos e Despesas</div>', unsafe_allow_html=True)
        st.markdown(
            render_kpi_row([
                dict(label="CMV (CUSTO DAS VENDAS)", value=formata_brl(cmv_ger_r),
                     value_color=cor_variacao(cmv_ger_o - cmv_ger_r),
                     subtext=f"{_pct_receita(cmv_ger_r):.1f}% da receita líquida · Orçado: {formata_brl(cmv_ger_o)}", icon="📦"),
                dict(label="DESPESAS VARIÁVEIS", value=formata_brl(dvar_ger_r),
                     value_color=cor_variacao(dvar_ger_o - dvar_ger_r),
                     subtext=f"{_pct_receita(dvar_ger_r):.1f}% da receita líquida · Orçado: {formata_brl(dvar_ger_o)}", icon="📉"),
                dict(label="DESPESAS OPERACIONAIS", value=formata_brl(dop_ger_r),
                     value_color=cor_variacao(dop_ger_o - dop_ger_r),
                     subtext=f"{_pct_receita(dop_ger_r):.1f}% da receita líquida · Orçado: {formata_brl(dop_ger_o)}", icon="🏢"),
                dict(label="TOTAL CUSTOS + DESPESAS", value=formata_brl(total_ger_r),
                     value_color=cor_variacao(total_ger_o - total_ger_r),
                     subtext=f"{_pct_receita(total_ger_r):.1f}% da receita líquida · Orçado: {formata_brl(total_ger_o)}", icon="🧾"),
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
    st.markdown(f'<div class="section-title">📅 Histórico Mensal Mês a Mês — {label_visao}</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    ch1, ch2 = st.columns(2)
    with ch1:
        tipo_hist = st.radio("Base de Dados:", ["Realizado", "Orçado"], horizontal=True)
    with ch2:
        visao_hist_dre = st.radio(
            "Detalhes:",
            [
                "Grupos Fechados (Sintético)",
                "Todas as Contas (Analítico)",
                "Visão Gerencial (Custos e Despesas)",
            ],
            horizontal=True,
        )

    linhas_hist = df_ref[col_nome].dropna().astype(str).unique()
    is_sintetica_hist = visao_hist_dre == "Grupos Fechados (Sintético)"
    is_gerencial_hist = visao_hist_dre == "Visão Gerencial (Custos e Despesas)"
    if is_sintetica_hist:
        linhas_hist = [l for l in linhas_hist if eh_grupo_sintetico(l)]
    elif is_gerencial_hist:
        linhas_hist = [l for l in linhas_hist if eh_linha_custos_despesas(l)]

    if not is_sintetica_hist and not is_gerencial_hist:
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
    if is_gerencial_hist:
        # Na visão Gerencial, a referência de "melhor/pior mês" é o total de
        # Custos + Despesas (CMV + Despesas Variáveis + Despesas
        # Operacionais) -- e, diferente da receita, aqui "melhor" é o mês
        # com o MENOR custo, não o maior.
        def _total_custos_mes(m_col):
            cmv_m = abs(get_valor_consolidado_multi(target_dfs, "4 - ", [m_col], exato_linha_sintetica=True)) \
                or abs(get_valor_consolidado_multi(target_dfs, "4 - Custo das Vendas", [m_col]))
            dvar_m = abs(get_valor_consolidado_multi(target_dfs, "6 - Despesas Variáveis", [m_col]))
            dop_m = abs(get_valor_consolidado_multi(target_dfs, "8 - Despesas Operacionais", [m_col]))
            return cmv_m + dvar_m + dop_m

        valores_ref_mensal = {m_nome: _total_custos_mes(m_col) for m_nome, m_col in m_map.items()}
        meses_com_dado = {m: v for m, v in valores_ref_mensal.items() if v != 0}

        if meses_com_dado:
            mes_menor_custo = min(meses_com_dado, key=meses_com_dado.get)
            mes_maior_custo = max(meses_com_dado, key=meses_com_dado.get)
            media_mensal_hist = sum(meses_com_dado.values()) / len(meses_com_dado)

            st.markdown(
                render_kpi_row([
                    dict(label=f"MENOR CUSTO MENSAL ({tipo_hist.upper()})", value=formata_brl(meses_com_dado[mes_menor_custo]),
                         value_color=COLORS["positive"], subtext=mes_menor_custo.capitalize(), icon="🏆"),
                    dict(label=f"MAIOR CUSTO MENSAL ({tipo_hist.upper()})", value=formata_brl(meses_com_dado[mes_maior_custo]),
                         value_color=COLORS["negative"], subtext=mes_maior_custo.capitalize(), icon="📈"),
                    dict(label="MÉDIA MENSAL (CUSTOS + DESPESAS)", value=formata_brl(media_mensal_hist),
                         value_color=COLORS["text"], subtext=f"{len(meses_com_dado)} meses com dados", icon="📊"),
                    dict(label="AMPLITUDE (MAIOR - MENOR)", value=formata_brl(meses_com_dado[mes_maior_custo] - meses_com_dado[mes_menor_custo]),
                         value_color=COLORS["muted_line"], subtext="Variação entre extremos", icon="↕️"),
                ]),
                unsafe_allow_html=True,
            )
            st.markdown("<br>", unsafe_allow_html=True)
    else:
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

    # -----------------------------------------------------------------------
    # Estresse por Linha da DRE — impacto em cascata
    # -----------------------------------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">🧪 Estresse de Cenário por Linha da DRE — Impacto em Cascata</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Escolha uma linha da DRE e aplique uma variação percentual para ver o impacto direto no "
        "EBITDA (e na Receita Líquida, se a linha escolhida fizer parte dela). Se outras linhas devem "
        "se mover junto (ex.: um custo variável que sobe quando as vendas sobem), marque-as em "
        "\"Linhas dependentes\" -- elas recebem a MESMA variação percentual da linha principal."
    )

    col_nome_stress = "Nome" if "Nome" in df_ref.columns else df_ref.columns[0]
    todas_linhas_dre = df_ref[col_nome_stress].dropna().astype(str).unique().tolist()

    ce1, ce2 = st.columns([1.3, 1.3])
    with ce1:
        idx_padrao_stress = (
            todas_linhas_dre.index("3 - Receita Operacional Liquida")
            if "3 - Receita Operacional Liquida" in todas_linhas_dre else 0
        )
        linha_estresse_sel = st.selectbox(
            "Linha da DRE a estressar:", todas_linhas_dre, index=idx_padrao_stress, key="linha_estresse_sel",
        )
        pct_estresse_linha = st.slider(
            f'Variação em "{linha_estresse_sel}" (%):',
            min_value=-50.0, max_value=50.0, value=5.0, step=1.0, key="pct_estresse_linha",
        )
    with ce2:
        linhas_dependentes = st.multiselect(
            "Linhas dependentes (variam junto, mesma variação %):",
            [l for l in todas_linhas_dre if l != linha_estresse_sel],
            key="linhas_dependentes_estresse",
            help='Ex.: se estressar "Vendas", marque aqui os custos variáveis que sobem/descem junto.',
        )

    periodo_estresse = colunas_validas  # ano completo com dados válidos no escopo atual

    def _valor_linha_stress(termo):
        return get_valor_consolidado_multi(list_df_real, termo, periodo_estresse, exato_linha_sintetica=True)

    LINHAS_RECEITA_STRESS = {"1 - Receita Operacional Bruta", "3 - Receita Operacional Liquida"}

    valor_original_linha = _valor_linha_stress(linha_estresse_sel)
    valor_estressado_linha = valor_original_linha * (1 + pct_estresse_linha / 100.0)
    delta_linha_principal = valor_estressado_linha - valor_original_linha

    delta_total_ebitda = delta_linha_principal
    delta_total_receita = delta_linha_principal if linha_estresse_sel in LINHAS_RECEITA_STRESS else 0.0

    detalhes_dependentes = []
    for linha_dep in linhas_dependentes:
        valor_dep = _valor_linha_stress(linha_dep)
        valor_dep_novo = valor_dep * (1 + pct_estresse_linha / 100.0)
        delta_dep = valor_dep_novo - valor_dep
        delta_total_ebitda += delta_dep
        if linha_dep in LINHAS_RECEITA_STRESS:
            delta_total_receita += delta_dep
        detalhes_dependentes.append((linha_dep, valor_dep, valor_dep_novo, delta_dep))

    ebitda_original_stress = _valor_linha_stress("11 - EBITDA")
    ebitda_novo_stress = ebitda_original_stress + delta_total_ebitda

    rec_liq_original_stress = _valor_linha_stress("3 - Receita Operacional Liquida")
    rec_liq_novo_stress = rec_liq_original_stress + delta_total_receita

    margem_original_stress = (ebitda_original_stress / rec_liq_original_stress * 100) if rec_liq_original_stress else 0
    margem_nova_stress = (ebitda_novo_stress / rec_liq_novo_stress * 100) if rec_liq_novo_stress else 0

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        render_kpi_row([
            dict(label=f'"{linha_estresse_sel[:28]}" — Original', value=formata_brl(valor_original_linha),
                 value_color=COLORS["text"], subtext="Valor realizado no período", icon="📍"),
            dict(label=f'"{linha_estresse_sel[:28]}" — Estressado', value=formata_brl(valor_estressado_linha),
                 value_color=cor_variacao(delta_linha_principal), subtext=f"{pct_estresse_linha:+.1f}% aplicado", icon="🧪"),
            dict(label="EBITDA — Impacto do Cenário", value=formata_brl(ebitda_novo_stress),
                 value_color=cor_variacao(delta_total_ebitda),
                 subtext=f"Original: {formata_brl(ebitda_original_stress)} · Delta: {formata_brl(delta_total_ebitda)}", icon="💹"),
            dict(label="Margem EBITDA — Impacto do Cenário", value=f"{margem_nova_stress:.1f}%",
                 value_color=cor_variacao(margem_nova_stress - margem_original_stress),
                 subtext=f"Original: {margem_original_stress:.1f}%", icon="📊"),
        ]),
        unsafe_allow_html=True,
    )

    if detalhes_dependentes:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-title">🔗 Linhas Dependentes — Detalhe do Impacto</div>', unsafe_allow_html=True)
        df_dependentes_stress = pd.DataFrame(
            [(l, v_o, v_n, d) for l, v_o, v_n, d in detalhes_dependentes],
            columns=["Linha da DRE", "Valor Original", "Valor Estressado", "Delta (R$)"],
        )
        st.dataframe(
            df_dependentes_stress.style.format({
                "Valor Original": formata_brl, "Valor Estressado": formata_brl, "Delta (R$)": formata_brl,
            }).map(cor_valor, subset=["Delta (R$)"]),
            use_container_width=True, hide_index=True,
            height=38 + len(df_dependentes_stress) * 35,
        )

    st.caption(
        "⚠️ Este cenário assume que todas as demais linhas da DRE permanecem constantes, exceto a linha "
        "estressada e as linhas dependentes marcadas -- é uma simulação isolada (\"what-if\"), não uma "
        "reprojeção completa do orçamento."
    )

# ---------------------------------------------------------------------------
# ABA 5: EMISSÃO DE RELATÓRIOS
# ---------------------------------------------------------------------------
MODELOS_RELATORIO = {
    "🛒 Relatório de Custos - Compras": {
        "linhas_dre": [
            "6.6 - Material de Embalagem",
            "6.11 - Catálogos e Revistas",
            "6.13 - Amostras",
            "6.14 - Flaconetes",
        ],
        # O plano de contas "Mercadorias" faz parte do modelo de Compras, mas
        # não tem uma "Linha DRE" correspondente na Tabela_Contas/DIÁRIO --
        # por isso ele é puxado à parte, direto pelo nome do Plano de Contas,
        # mesmo sem nenhuma linha da DRE selecionada apontar para ele.
        "forcar_planos_contas": ["Mercadorias"],
        "permitir_lancamento_manual": False,
    },
    "🚚 Relatório de Custos - Suprimentos": {
        "linhas_dre": [
            "6.8 - Serviço de Entrega",
            "8.1.3 - Limpeza e Conservação",
            "8.1.4 - Manutenção e Reparos",
            "8.5.3 - Combustível",
            "8.5.4 - Manutenção Veículos",
            "8.6.1 - Material de Escritório",
            "8.6.4 - Copa e Cozinha",
            "8.7.1 - Despesas com Viagens",
            "8.8.10 - Serviços de Transporte",
            "8.8.11 - Outros Serviços Terceirizados",
        ],
        "forcar_planos_contas": [],
        "permitir_lancamento_manual": False,
    },
    "👥 Relatório de Custos - RH": {
        "linhas_dre": [
            "6.1 - Comissões sobre Vendas",
            "8.3 - Pessoal",
            "8.3.1 - Salários",
            "8.3.1.1 - Salário",
            "8.3.1.2 - Adiantamento Salarial",
            "8.3.1.3 - 13º Salário",
            "8.3.1.4 - Hora Extra",
            "8.3.1.5 - DSR",
            "8.3.1.6 - Férias",
            "8.3.1.7 - Descontos Gerais Sobre a Folha",
            "8.3.2 - Encargos Sociais",
            "8.3.2.1 - INSS",
            "8.3.2.2 - FGTS",
            "8.3.2.3 - IRRF Salários",
            "8.3.2.4 - Contribuição Sindical / Assistencial",
            "8.3.2.5 - Desconto INSS",
            "8.3.2.6 - Desconto IRRF",
            "8.3.3 - Benefícios",
            "8.3.3.1 - Vale Transporte",
            "8.3.3.2 - Vale Refeição",
            "8.3.3.3 - Plano de Saúde",
            "8.3.3.4 - Ajuda de Custo para Funcionários",
            "8.3.3.5 - Outros Benefícios",
            "8.3.3.6 - Desconto sobre Beneficios",
            "8.3.3.7 - Prêmios / Bônus",
            "8.3.4 - Movimentação de Pessoal",
            "8.3.4.1 - Indenizações / Rescisões",
            "8.3.4.2 - Multa de FGTS",
            "8.3.4.3 - Exames Admissionais / Demissionais",
            "8.3.4.4 - Recrutamento e Seleção",
            "8.3.4.5 - Descontos Sobre Rescisões",
            "8.3.4.6 - Temporários / Estagiários",
            "8.3.5 - Uniformes",
            "8.3.6 - Treinamento",
            "8.3.7 - Pro Labore",
            "8.3.8 - Saúde e Segurança do Trabalho",
            "8.3.9 - Outras Despesas com Pessoal",
            "8.3.10 - Contratação Pessoa Jurídica",
            "8.3.11 - Encontro e Confraternização de Time",
            "8.6.6 - Outras Despesas Administrativas",
            "8.8.2 - Auditoria / Consultoria",
        ],
        "forcar_planos_contas": [],
        # Várias linhas da DRE do RH não têm um Plano de Contas correspondente
        # no DIÁRIO/Tabela_Contas (o valor é lançado direto na linha). Nesses
        # casos, em vez de deixar a composição vazia, mostramos uma linha
        # "Lançado Manualmente" com o valor da própria linha da DRE.
        "permitir_lancamento_manual": True,
    },
}

with tab5:
    st.markdown('<div class="section-title">📤 Emissão de Relatórios em Excel</div>', unsafe_allow_html=True)
    st.caption(
        "Escolha um modelo padrão ou selecione manualmente as linhas da DRE, e gere um relatório "
        "formatado com 4 planilhas: Resumo (consolidado), Detalhe Mensal (contas x meses, por loja), "
        "Plano de Contas (composição de cada linha, por loja, a partir do DIÁRIO) e Lançamentos "
        "(cópia filtrada da DIÁRIO). Dentro do Excel, use o filtro (▾) da coluna \"Loja\" para escolher "
        "qual visão ver sem precisar gerar o arquivo de novo."
    )
    st.markdown("<br>", unsafe_allow_html=True)

    linhas_relatorio = df_ref[col_nome].dropna().astype(str).unique()

    opcoes_modelo = ["Seleção manual"] + list(MODELOS_RELATORIO.keys())
    modelo_sel = st.selectbox("📁 Modelo de Relatório:", opcoes_modelo)

    default_contas = []
    termos_nao_encontrados = []
    if modelo_sel != "Seleção manual":
        for termo in MODELOS_RELATORIO[modelo_sel]["linhas_dre"]:
            termo_norm = termo.strip().lower()
            encontrados = [l for l in linhas_relatorio if termo_norm in str(l).strip().lower()]
            if encontrados:
                default_contas.extend(encontrados)
            else:
                termos_nao_encontrados.append(termo)
        default_contas = list(dict.fromkeys(default_contas))
        if termos_nao_encontrados:
            st.warning(
                "Não encontrei estas linhas do modelo na DRE atual (o texto pode estar um pouco diferente): "
                + "; ".join(termos_nao_encontrados)
                + ". Adicione-as manualmente no campo abaixo, se existirem."
            )

    contas_relatorio = st.multiselect(
        "🔍 Linhas da DRE incluídas no relatório:",
        options=linhas_relatorio,
        default=default_contas,
        key=f"contas_relatorio__{modelo_sel}",
    )

    opcoes_lojas_relatorio = list(abas_disponiveis)

    lojas_relatorio_sel = st.multiselect(
        "🏬 Lojas / Visões incluídas no relatório:",
        options=opcoes_lojas_relatorio,
        default=opcoes_lojas_relatorio,
        help=(
            "Por padrão, todas as lojas e visões consolidadas entram no relatório -- dentro do Excel "
            "gerado, use o filtro (▾) na coluna \"Loja\" das abas \"Detalhe Mensal\" e \"Plano de "
            "Contas\" para escolher qual visão ver, sem precisar gerar o arquivo de novo. Só reduza a "
            "seleção aqui se quiser um arquivo menor desde já."
        ),
    )

    col_btn, col_info = st.columns([1, 2])
    with col_btn:
        gerar_clicado = st.button(
            "📊 Gerar Relatório Excel",
            use_container_width=True,
            disabled=not contas_relatorio or not lojas_relatorio_sel,
        )

    if gerar_clicado and contas_relatorio:
        with st.spinner("Carregando dados por loja, plano de contas e DIÁRIO..."):
            # Carrega só as lojas/visões escolhidas no filtro acima -- evita
            # gerar um arquivo gigante com todas as 26 abas de uma vez.
            dados_por_loja_rel = carregar_dados_por_loja(path_orc, path_real, lojas_relatorio_sel)
            df_tabela_contas = carregar_tabela_contas(path_real)
            mapa_planos_dre_rel = montar_mapa_planos_por_dre(df_tabela_contas)
            df_diario_rel = carregar_diario(path_real)
            df_tabela_lojas_rel = carregar_tabela_lojas(path_real)
            mapa_loja_cc_rel = montar_mapa_loja_centro_custo(df_tabela_lojas_rel)

        info_modelo_sel = MODELOS_RELATORIO.get(modelo_sel, {})

        with st.spinner("Montando o relatório em Excel..."):
            excel_bytes = montar_relatorio_excel(
                contas_relatorio, list_df_real, list_df_orc, m_map, colunas_validas, label_visao,
                dados_por_loja=dados_por_loja_rel,
                mapa_planos_dre=mapa_planos_dre_rel,
                df_diario=df_diario_rel,
                forcar_planos_contas=info_modelo_sel.get("forcar_planos_contas", []),
                permitir_lancamento_manual=info_modelo_sel.get("permitir_lancamento_manual", False),
                mapa_loja_centro_custo=mapa_loja_cc_rel,
            )
        st.session_state["relatorio_excel_bytes"] = excel_bytes

        def _nome_arquivo_modelo(nome_modelo):
            """Usa o nome completo do modelo (igual aparece no seletor, só
            sem o emoji na frente) como base do nome do arquivo -- ex.:
            "🛒 Relatório de Custos - Compras" vira "Relatório de Custos -
            Compras". Só remove caracteres inválidos em nome de arquivo."""
            texto = re.sub(r"^[^\w]+", "", nome_modelo, flags=re.UNICODE).strip()
            texto = re.sub(r'[\\/*?:"<>|]', "", texto)
            return texto or "Relatório"

        st.session_state["relatorio_excel_nome"] = f"{_nome_arquivo_modelo(modelo_sel)}.xlsx"
        st.success(f"Relatório gerado com {len(contas_relatorio)} conta(s) selecionada(s), pronto para download.")

        if df_diario_rel is None or df_diario_rel.empty:
            st.warning(
                "Aba 'DIÁRIO' não encontrada (ou vazia/sem as colunas esperadas) no arquivo Realizado 2026 — "
                "a aba 'Plano de Contas' do relatório usou o método antigo (Tabela_Contas) como alternativa."
            )
        else:
            st.caption(f"📄 DIÁRIO conectado: {len(df_diario_rel)} lançamento(s) encontrados na aba do Realizado 2026.")

    if not contas_relatorio:
        st.info("Selecione um modelo padrão acima, ou escolha manualmente ao menos uma linha da DRE.")

    if st.session_state.get("relatorio_excel_bytes"):
        st.download_button(
            label="⬇️ Baixar Relatório (.xlsx)",
            data=st.session_state["relatorio_excel_bytes"],
            file_name=st.session_state.get("relatorio_excel_nome", "relatorio_dre.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("O relatório contém três planilhas: **Resumo** (total do ano por conta, consolidado), **Detalhe Mensal** (contas nas linhas x meses nas colunas, por loja) e **Plano de Contas** (composição de cada linha da DRE, por loja, a partir do DIÁRIO).")

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