import os
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Configuração da página
st.set_page_config(
    page_title="Controladoria B&A - Painel Financeiro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilização CSS Customizada (Clean & Muted / Dark Graphite Continuous + RESPONSIVIDADE MOBILE)
st.markdown(
    """
    <style>
        .stApp {
            background-color: #0D1117;
            color: #C9D1D9;
        }
        [data-testid="stSidebar"] {
            background-color: #010409;
            border-right: 1px solid #21262D;
        }
        
        /* Ajuste de Espaçamento e Menu Suspenso na Sidebar */
        [data-testid="stSidebarUserContent"] {
            padding-bottom: 80px !important;
        }
        div[data-baseweb="popover"] {
            z-index: 999999 !important;
        }

        /* Header Superior do Streamlit Transparente */
        header[data-testid="stHeader"] {
            background: transparent !important;
            background-color: rgba(0, 0, 0, 0) !important;
            pointer-events: none;
        }
        header[data-testid="stHeader"] * {
            pointer-events: auto;
        }

        /* Ocultar botão Deploy */
        [data-testid="stAppDeployButton"] {
            display: none !important;
        }

        [data-testid="stHeaderActionElements"] {
            position: fixed !important;
            top: 10px;
            right: 10px;
            z-index: 999999;
        }

        /* Header compacto */
        .main-header {
            background-color: #161B22;
            padding: 12px 20px;
            border-radius: 6px;
            border: 1px solid #21262D;
            margin-bottom: 15px;
        }
        .main-header h3 {
            margin: 0 !important;
            padding: 0 !important;
            font-size: 18px !important;
            color: #F0F6FC !important;
            font-weight: 600;
            letter-spacing: 0.3px;
        }
        .main-header p {
            margin: 2px 0 0 0 !important;
            font-size: 12px !important;
            color: #8B949E !important;
        }

        /* KPI Cards Executivos */
        .kpi-card {
            background-color: #161B22;
            padding: 14px 16px;
            border-radius: 6px;
            border: 1px solid #21262D;
            margin-bottom: 10px;
        }
        .kpi-label {
            font-size: 11px !important;
            font-weight: 600 !important;
            color: #8B949E !important;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
        }
        .kpi-value {
            font-size: 20px !important;
            font-weight: 600 !important;
            letter-spacing: -0.3px;
        }
        .kpi-subtext {
            font-size: 12px;
            font-weight: 500;
            margin-top: 4px;
            color: #A6B0C3;
        }

        /* Estilização da Barra de Progresso Integrada */
        .progress-container {
            width: 100%;
            background-color: #21262D;
            border-radius: 3px;
            height: 6px;
            margin-top: 10px;
            overflow: hidden;
        }
        .progress-bar {
            height: 100%;
            background-color: #2F80ED;
            border-radius: 3px;
        }

        button[data-baseweb="tab"] {
            background-color: #161B22;
            border-radius: 4px;
            color: #8B949E !important;
            padding: 6px 14px;
            margin-right: 4px;
            border: 1px solid #21262D;
            font-size: 13px;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            background-color: #21262D !important;
            color: #F0F6FC !important;
            font-weight: 600;
            border-color: #2F80ED !important;
        }
        hr {
            border-color: #21262D !important;
        }

        /* Travar primeira coluna em tabelas */
        div[data-testid="stDataFrame"] div[role="grid"] div[role="row"] div[role="gridcell"]:first-child,
        div[data-testid="stDataFrame"] div[role="grid"] div[role="row"] div[role="columnheader"]:first-child {
            position: sticky;
            left: 0;
            background-color: #161B22 !important;
            z-index: 3;
            border-right: 1px solid #21262D;
        }

        /* ========================================================= */
        /* REGRAS CSS EXCLUSIVAS PARA DISPOSITIVOS MÓVEIS (MOBILE)   */
        /* ========================================================= */
        @media only screen and (max-width: 768px) {
            /* Empilha colunas nativas do Streamlit no mobile */
            div[data-testid="column"] {
                width: 100% !important;
                flex: 1 1 100% !important;
                min-width: 100% !important;
            }

            /* Ajuste de fontes e padding do Header */
            .main-header {
                padding: 10px 14px !important;
            }
            .main-header h3 {
                font-size: 15px !important;
            }
            .main-header p {
                font-size: 11px !important;
            }

            /* Ajustes nos Cards de KPI */
            .kpi-card {
                padding: 12px !important;
            }
            .kpi-value {
                font-size: 18px !important;
            }

            /* Ajuste nas abas (tabs) para rolar horizontalmente no celular */
            div[data-baseweb="tab-list"] {
                display: flex !important;
                overflow-x: auto !important;
                white-space: nowrap !important;
                padding-bottom: 5px !important;
            }
            button[data-baseweb="tab"] {
                font-size: 11px !important;
                padding: 5px 10px !important;
            }

            /* Habilita scroll horizontal suave nas tabelas em telas pequenas */
            div[data-testid="stDataFrame"] {
                overflow-x: auto !important;
            }
        }
    </style>
""",
    unsafe_allow_html=True,
)

# Configuração global de travamento para o Streamlit/Plotly
CONFIG_PLOTLY_TRAVADO = {
    'staticPlot': False,
    'displayModeBar': False,
    'scrollZoom': False,
    'doubleClick': False,
    'responsive': True # Habilita responsividade no Plotly
}


# 1. CARREGAMENTO DOS DADOS COM CACHE SEGURO (NUVEM COM FALLBACK LOCAL)
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
    abas_validas = [
        sheet for sheet in xls_orc.sheet_names if sheet not in abas_ignorar
    ]

    return abas_validas, path_orc, path_real


try:
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


# 2. BARRA LATERAL - FILTROS PRIMEIRO
st.sidebar.title("Filtros do Painel")

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

opcoes_consolidadas = [
    a for a in abas_disponiveis if a in abas_consolidadas_permitidas
]
if not opcoes_consolidadas:
    opcoes_consolidadas = abas_disponiveis

opcoes_unidades = [
    a for a in abas_disponiveis if a not in abas_consolidadas_permitidas
]
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

list_df_orc, list_df_real = carregar_dados_abas(
    path_orc, path_real, abas_para_carregar
)

if not list_df_real or not list_df_orc:
    st.error("Erro ao ler dados das abas.")
    st.stop()

# 3. FILTRO DE PERÍODO / MESES
nomes_meses = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
]
meses_cols = [
    "01/2026", "02/2026", "03/2026", "04/2026", "05/2026", "06/2026",
    "07/2026", "08/2026", "09/2026", "10/2026", "11/2026", "12/2026"
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
    mes_ref = st.sidebar.selectbox(
        "Mês Desejado:", list(m_map.keys()), index=idx_default
    )
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


# 4. FUNÇÕES DE SUPORTE
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
    color = "#3FB950" if val >= 0 else "#F85149"
    return f"color: {color}; font-weight: 500;"


LAYOUT_NEUTRO = dict(
    paper_bgcolor="rgba(22, 27, 34, 1)",
    plot_bgcolor="rgba(22, 27, 34, 1)",
    font=dict(color="#8B949E", family="Segoe UI, sans-serif"),
    margin=dict(l=20, r=20, t=30, b=60),
)

# 5. CABEÇALHO COMPACTO
st.markdown(
    f"""
    <div class="main-header">
        <h3>PAINEL ANALÍTICO DE PERFORMANCE ESTRATÉGICA</h3>
        <p>Visão: <b>{label_visao}</b> &nbsp;|&nbsp; Período: <b>{label_periodo_kpi}</b></p>
    </div>
    """,
    unsafe_allow_html=True,
)

# KPIs Superiores
rec_liq_real_kpi = get_valor_consolidado_multi(list_df_real, "3 - Receita Operacional Liquida", cols_kpi)
rec_liq_orc_kpi = get_valor_consolidado_multi(list_df_orc, "3 - Receita Operacional Liquida", cols_kpi)

ebitda_real_kpi = get_valor_consolidado_multi(list_df_real, "11 - EBITDA", cols_kpi)
ebitda_orc_kpi = get_valor_consolidado_multi(list_df_orc, "11 - EBITDA", cols_kpi)

margem_ebitda_kpi = (ebitda_real_kpi / rec_liq_real_kpi * 100) if rec_liq_real_kpi != 0 else 0

diff_ebitda_kpi = ebitda_real_kpi - ebitda_orc_kpi
pct_ebitda_kpi = (diff_ebitda_kpi / abs(ebitda_orc_kpi)) * 100 if ebitda_orc_kpi != 0 else 0

pct_vendas_prog = min(100.0, max(0.0, (rec_liq_real_kpi / rec_liq_orc_kpi * 100))) if rec_liq_orc_kpi > 0 else 0
pct_lucro_prog = min(100.0, max(0.0, (ebitda_real_kpi / ebitda_orc_kpi * 100))) if ebitda_orc_kpi > 0 else 0

cor_rec = "#3FB950" if rec_liq_real_kpi >= 0 else "#F85149"
cor_ebitda = "#3FB950" if ebitda_real_kpi >= 0 else "#F85149"
cor_diff_eb = "#3FB950" if diff_ebitda_kpi >= 0 else "#F85149"
cor_mg_eb = "#3FB950" if margem_ebitda_kpi >= 0 else "#F85149"

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">RECEITA LÍQUIDA (YTD)</div>
            <div class="kpi-value" style="color: {cor_rec};">{formata_brl(rec_liq_real_kpi)}</div>
            <div class="kpi-subtext">Orçado: {formata_brl(rec_liq_orc_kpi)}</div>
            <div class="progress-container">
                <div class="progress-bar" style="width: {pct_vendas_prog:.1f}%;"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k2:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">EBITDA (YTD)</div>
            <div class="kpi-value" style="color: {cor_ebitda};">{formata_brl(ebitda_real_kpi)}</div>
            <div class="kpi-subtext">Orçado: {formata_brl(ebitda_orc_kpi)}</div>
            <div class="progress-container">
                <div class="progress-bar" style="width: {pct_lucro_prog:.1f}%;"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k3:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">VARIAÇÃO EBITDA</div>
            <div class="kpi-value" style="color: {cor_diff_eb};">{formata_brl(diff_ebitda_kpi)}</div>
            <div class="kpi-subtext" style="color: {cor_diff_eb};">{pct_ebitda_kpi:+.1f}% vs Orçamento</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k4:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">MARGEM EBITDA %</div>
            <div class="kpi-value" style="color: {cor_mg_eb};">{margem_ebitda_kpi:.1f}%</div>
            <div class="kpi-subtext">Realizada no Período</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ABAS
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📊 Visão Geral & Charts",
        "📋 DRE Completa & Desvios",
        "📅 Histórico Mensal",
        "🔮 Previsões & Trends",
    ]
)

# ABA 1: VISÃO GERAL & CHARTS
with tab1:
    st.caption(f"Visualização e Eficiência referente ao período: **{label_periodo_kpi}**")

    cg1, cg2 = st.columns(2)

    with cg1:
        st.markdown("##### **Bridge de Performance (YTD)**")

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
            "Margem Bruta", "SG&A", "EBIT", "D&A", "EBITDA"
        ]
        
        measures = ["absolute", "relative", "total", "relative", "total", "relative", "total", "relative", "total"]
        y_bridge = [p_rec_bruta, p_deducoes, 0, p_cmv, 0, p_sga, 0, abs(p_deprec), 0]
        text_labels = [
            f"{p_rec_bruta}%", f"{p_deducoes}%", f"{p_rec_liq}%", 
            f"{p_cmv}%", f"{p_mb}%", f"{p_sga}%", 
            f"{p_ebit}%", f"{p_deprec}%", f"{p_ebitda}%"
        ]

        fig_waterfall = go.Figure(
            go.Waterfall(
                orientation="v",
                measure=measures,
                x=x_bridge,
                y=y_bridge,
                text=text_labels,
                textposition="outside",
                connector={"line": {"color": "#30363D", "width": 1}},
                decreasing={"marker": {"color": "#A6B0C3"}},
                increasing={"marker": {"color": "#2F80ED"}},
                totals={"marker": {"color": "#484F58"}},
            )
        )
        fig_waterfall.update_layout(
            **LAYOUT_NEUTRO,
            height=400,
            xaxis=dict(tickangle=-45, gridcolor="rgba(0,0,0,0)", fixedrange=True),
            yaxis=dict(showticklabels=False, gridcolor="rgba(0,0,0,0)", fixedrange=True),
        )
        st.plotly_chart(fig_waterfall, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    with cg2:
        st.markdown("##### **Real vs. Orçado (YTD)**")

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
                go.Bar(name="Realizado (R$)", x=cats, y=val_r, text=labels_r, textposition="outside", marker_color="#2F80ED"),
                go.Bar(name="Orçado (R$)", x=cats, y=val_o, text=labels_o, textposition="outside", marker_color="#484F58"),
            ]
        )
        fig_bar.update_layout(
            barmode="group",
            **LAYOUT_NEUTRO,
            height=400,
            xaxis=dict(gridcolor="#21262D", zerolinecolor="#21262D", fixedrange=True),
            yaxis=dict(showticklabels=False, gridcolor="rgba(0,0,0,0)", fixedrange=True),
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_bar, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    st.markdown("<br>", unsafe_allow_html=True)

    cg3, cg4 = st.columns([1.3, 0.7])

    with cg3:
        st.markdown("##### **Evolução Mensal (Receita vs. Margem de Contribuição)**")
        
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
                x=rotulos_m,
                y=rec_m,
                mode="lines+markers+text",
                name="Receita (R$)",
                text=labels_rec,
                textposition="top center",
                line=dict(color="#2F80ED", width=2, shape="spline"),
                marker=dict(size=6, color="#161B22", line=dict(color="#2F80ED", width=2)),
                textfont=dict(color="#F0F6FC", size=11, family="Segoe UI")
            )
        )

        fig_line.add_trace(
            go.Scatter(
                x=rotulos_m,
                y=mc_m,
                mode="lines+markers+text",
                name="Margem Contrib. 2 (R$)",
                text=labels_mc,
                textposition="bottom center",
                line=dict(color="#A6B0C3", width=2, shape="spline"),
                marker=dict(size=6, color="#161B22", line=dict(color="#A6B0C3", width=2)),
                textfont=dict(color="#C9D1D9", size=11, family="Segoe UI")
            )
        )

        fig_line.update_layout(
            **LAYOUT_NEUTRO,
            height=380,
            xaxis=dict(
                showgrid=False,
                zeroline=False,
                tickangle=-45,
                tickfont=dict(size=11, color="#8B949E"),
                fixedrange=True
            ),
            yaxis=dict(
                showgrid=False,
                showticklabels=False,
                zeroline=False,
                fixedrange=True
            ),
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.25,
                xanchor="center",
                x=0.5,
                font=dict(color="#C9D1D9")
            )
        )

        st.plotly_chart(fig_line, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    with cg4:
        st.markdown("##### **Composição dos Custos & Saídas**")
        
        cmv_real_kpi = abs(get_valor_consolidado_multi(list_df_real, "4 - ", cols_kpi, exato_linha_sintetica=True))
        if cmv_real_kpi == 0:
            cmv_real_kpi = abs(get_valor_consolidado_multi(list_df_real, "4 - Custo das Vendas", cols_kpi))

        desp_op_real = abs(get_valor_consolidado_multi(list_df_real, "8 - Despesas Operacionais", cols_kpi))

        v_cmv_pie = abs(cmv_real_kpi)
        v_desp_var_pie = abs(get_valor_consolidado_multi(list_df_real, "6 - Despesas Variáveis", cols_kpi))
        v_desp_op_pie = abs(desp_op_real)
        v_deprec_pie = abs(get_valor_consolidado_multi(list_df_real, "13 - Depreciação e Amortização", cols_kpi))

        fig_donut = go.Figure(
            data=[
                go.Pie(
                    labels=["CMV / Custo", "Despesas Var.", "Despesas Op. (OpEx)", "Depreciação/Amort."],
                    values=[v_cmv_pie, v_desp_var_pie, v_desp_op_pie, v_deprec_pie],
                    hole=0.5,
                    marker=dict(colors=["#2F80ED", "#A6B0C3", "#484F58", "#30363D"]),
                    textinfo="percent",
                    hoverinfo="label+value+percent"
                )
            ]
        )
        fig_donut.update_layout(
            **LAYOUT_NEUTRO,
            height=380,
            legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_donut, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)


# ABA 2: DRE COMPLETA & DESVIOS
with tab2:
    st.markdown(f"##### 📋 **Análise de DRE e Desvios - {label_periodo_graf}**")

    c1, _ = st.columns([2, 1])
    with c1:
        tipo_visao_dre = st.radio(
            "Filtro de Nível de Visão:",
            ["Apenas Grupos Principais (Sintética)", "Todas as Contas (Analítica)"],
            horizontal=True
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
            key="filtro_contas_dre"
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

    column_config_dre = {
        "Conta / Linha DRE": st.column_config.TextColumn(
            "Conta / Linha DRE",
            width="large",
        )
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


# ABA 3: HISTÓRICO MENSAL
with tab3:
    st.markdown("##### 📅 **Histórico Mensal Mês a Mês**")

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
            key="filtro_contas_hist"
        )
        if contas_filtradas_hist:
            linhas_hist = contas_filtradas_hist

    target_dfs = list_df_real if tipo_hist == "Realizado" else list_df_orc

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
        "Conta / Linha DRE": st.column_config.TextColumn(
            "Conta / Linha DRE",
            width="large",
            pinned=True
        )
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


# ABA 4: PREVISÕES & TRENDS
with tab4:
    st.markdown("##### 🔮 **Painel Avançado de Previsões e Tendências 2026**")

    c_f1, c_f2, c_f3 = st.columns([1.2, 1.2, 1.6])

    with c_f1:
        metrica_sel = st.selectbox(
            "Métrica de Análise:",
            ["Receita Operacional Líquida", "EBITDA"],
            index=0
        )
        termo_metrica = "3 - Receita Operacional Liquida" if metrica_sel == "Receita Operacional Líquida" else "11 - EBITDA"

    with c_f2:
        modelo_proj = st.selectbox(
            "Modelo de Projeção Futura:",
            ["Média Histórica (Run-Rate)", "Manter Orçamento (Budget)", "Ajustado por Sazonalidade/Performance"],
            index=0,
            help="Define como os meses futuros não realizados serão calculados."
        )

    with c_f3:
        sensibilidade = st.slider(
            "Ajuste Fino de Cenário / Estresse (%):",
            min_value=-20.0,
            max_value=20.0,
            value=0.0,
            step=1.0,
            help="Aplica uma variação percentual sobre os meses projetados."
        )

    meses_todos = list(m_map.keys())
    
    meses_realizados_cols = cols_kpi if tipo_periodo == "Mês Selecionado" else [m_map[m] for m in meses_todos if get_valor_consolidado_multi(list_df_real, termo_metrica, [m_map[m]]) != 0]
    
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

    c_proj = "#3FB950" if projecao_total_anual >= 0 else "#F85149"
    c_diff = "#3FB950" if diff_anual >= 0 else "#F85149"

    st.markdown("<br>", unsafe_allow_html=True)

    kp1, kp2, kp3, kp4 = st.columns(4)

    with kp1:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">PROJEÇÃO ANUAL ({metrica_sel.upper()})</div>
                <div class="kpi-value" style="color: {c_proj};">{formata_brl(projecao_total_anual)}</div>
                <div class="kpi-subtext">Cenário: {modelo_proj.split(' ')[0]}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with kp2:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">META ANUAL (ORÇADO)</div>
                <div class="kpi-value" style="color: #F0F6FC;">{formata_brl(val_orc_anual_total)}</div>
                <div class="kpi-subtext">Orçamento Fechado 2026</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with kp3:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">GAP / DESVIO ANUAL</div>
                <div class="kpi-value" style="color: {c_diff};">{formata_brl(diff_anual)}</div>
                <div class="kpi-subtext" style="color: {c_diff};">{diff_anual / abs(val_orc_anual_total) * 100:+.1f}% vs Meta</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with kp4:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">ATINGIMENTO ESTIMADO</div>
                <div class="kpi-value" style="color: {c_diff};">{pct_atingimento_anual:.1f}%</div>
                <div class="kpi-subtext">Média Mensal Real: {formata_m(media_mensal_real)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    df_trend = pd.DataFrame({
        "Mês": [m.capitalize() for m in meses_todos],
        "Valor Projetado/Real": valores_finais_mes,
        "Orçado": valores_orcado_mes,
        "Tipo": tipos_serie
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
            x=df_real_bar["Mês"],
            y=df_real_bar["Valor Projetado/Real"],
            name="Realizado",
            marker_color="#2F80ED",
            text=[formata_m(v) for v in df_real_bar["Valor Projetado/Real"]],
            textposition=pos_bar_real,
            textfont=dict(size=11, color="#F0F6FC"),
            cliponaxis=False
        )
    )

    df_proj_bar = df_trend[df_trend["Tipo"] == "Projetado"]
    pos_bar_proj = [posicoes_barras[i] for i in df_proj_bar.index]
    fig_comb.add_trace(
        go.Bar(
            x=df_proj_bar["Mês"],
            y=df_proj_bar["Valor Projetado/Real"],
            name="Projetado (Tendência)",
            marker_color="#30363D",
            text=[formata_m(v) for v in df_proj_bar["Valor Projetado/Real"]],
            textposition=pos_bar_proj,
            textfont=dict(size=11, color="#8B949E"),
            cliponaxis=False
        )
    )

    fig_comb.add_trace(
        go.Scatter(
            x=df_trend["Mês"],
            y=df_trend["Orçado"],
            name="Orçado (Meta)",
            mode="lines+markers+text",
            text=[formata_m(v) for v in df_trend["Orçado"]],
            textposition=posicoes_meta,
            textfont=dict(size=10, color="#F2994A"),
            line=dict(color="#F2994A", width=2, dash="dash"),
            marker=dict(size=6, color="#F2994A"),
            cliponaxis=False
        )
    )

    max_val = max(
        max(df_trend["Valor Projetado/Real"].dropna(), default=0),
        max(df_trend["Orçado"].dropna(), default=0)
    )

    fig_comb.update_layout(
        **LAYOUT_NEUTRO,
        title=f"Evolução Mensal & Projeção Run-Rate: {metrica_sel}",
        height=500,
        xaxis=dict(gridcolor="#21262D", zerolinecolor="#21262D", fixedrange=True),
        yaxis=dict(
            showticklabels=False,
            gridcolor="rgba(0,0,0,0)",
            range=[0, max_val * 1.35] if max_val > 0 else None,
            fixedrange=True
        ),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
        barmode="group"
    )

    st.plotly_chart(fig_comb, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    st.markdown("##### 📋 **Detalhamento da Projeção Mensal (R$)**")
    
    df_resumo_proj = pd.DataFrame({
        "Mês": [m.capitalize() for m in meses_todos],
        "Tipo de Dado": tipos_serie,
        "Valor Realizado / Projetado": valores_finais_mes,
        "Orçado Original": valores_orcado_mes,
        "Desvio (R$)": [v_p - v_o for v_p, v_o in zip(valores_finais_mes, valores_orcado_mes)]
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
        height=ALTURA_12_LINHAS
    )