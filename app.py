import os
import re
import io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA & THEME ENGINE (DARK GRAPHITE EXEC)
# ==============================================================================
st.set_page_config(
    page_title="Controladoria B&A - Executive Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        /* Fundo Principal e Cores de Fonte */
        .stApp {
            background-color: #0B0E14;
            color: #C9D1D9;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        }
        
        [data-testid="stSidebar"] {
            background-color: #010409;
            border-right: 1px solid #161B22;
        }
        
        /* Ocultar elementos desnecessários da UI */
        header[data-testid="stHeader"] { background: transparent !important; }
        [data-testid="stAppDeployButton"] { display: none !important; }

        /* Estilização do Cabeçalho */
        .main-header {
            background: linear-gradient(135deg, #161B22 0%, #0D1117 100%);
            padding: 16px 22px;
            border-radius: 8px;
            border: 1px solid #21262D;
            margin-bottom: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }
        .main-header h3 {
            margin: 0 !important;
            font-size: 20px !important;
            color: #F0F6FC !important;
            font-weight: 700;
            letter-spacing: 0.5px;
        }
        .main-header p {
            margin: 4px 0 0 0 !important;
            font-size: 13px !important;
            color: #8B949E !important;
        }

        /* Executive KPI Cards Modernos */
        .kpi-card {
            background-color: #161B22;
            padding: 16px;
            border-radius: 8px;
            border: 1px solid #21262D;
            transition: transform 0.2s, border-color 0.2s;
        }
        .kpi-card:hover {
            border-color: #388BFD;
        }
        .kpi-label {
            font-size: 11px !important;
            font-weight: 700 !important;
            color: #8B949E !important;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            margin-bottom: 8px;
        }
        .kpi-value {
            font-size: 22px !important;
            font-weight: 700 !important;
            letter-spacing: -0.5px;
        }
        .kpi-subtext {
            font-size: 12px;
            font-weight: 500;
            margin-top: 6px;
            color: #8B949E;
        }

        /* Barra de Progresso Customizada */
        .progress-container {
            width: 100%;
            background-color: #21262D;
            border-radius: 4px;
            height: 6px;
            margin-top: 10px;
            overflow: hidden;
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #1F6FEB 0%, #388BFD 100%);
            border-radius: 4px;
        }

        /* Estilização das Abas (Tabs) */
        button[data-baseweb="tab"] {
            background-color: #161B22;
            border-radius: 6px 6px 0 0;
            color: #8B949E !important;
            padding: 8px 18px;
            margin-right: 4px;
            border: 1px solid #21262D;
            font-size: 13px;
            font-weight: 500;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            background-color: #21262D !important;
            color: #F0F6FC !important;
            font-weight: 600;
            border-bottom: 2px solid #388BFD !important;
        }

        /* Responsividade para Dispositivos Móveis */
        @media only screen and (max-width: 768px) {
            div[data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; }
            .kpi-value { font-size: 18px !important; }
        }
    </style>
""",
    unsafe_allow_html=True,
)

CONFIG_PLOTLY_TRAVADO = {
    'staticPlot': False,
    'displayModeBar': False,
    'scrollZoom': False,
    'doubleClick': False,
    'responsive': True
}

# ==============================================================================
# 2. CARREGAMENTO E TRATAMENTO DE DADOS (CACHE & VECTORIZATION)
# ==============================================================================
@st.cache_resource
def obter_caminhos_excel():
    url_orc = "https://docs.google.com/spreadsheets/d/1x68Eg_6LlSKeFJEGmfhyBfcGgheSrVsl/export?format=xlsx"
    url_real = "https://docs.google.com/spreadsheets/d/12I0vGpYU_KNhGxAHOMHWAQu3Xkz_EsUZ/export?format=xlsx"

    caminho_base = r"G:\Meu Drive\Grupo B&A\Escritorio\Financeiro\COORDENAÇÃO FINANCEIRA\ORÇAMENTO\ORÇAMENTO 2026\CONTROLADORIA"
    path_orc_local = os.path.join(caminho_base, "ORCAMENTO 2026 - REV.1.xlsx")
    path_real_local = os.path.join(caminho_base, "REALIZADO 2026.xlsx")

    try:
        xls_orc = pd.ExcelFile(url_orc)
        path_orc, path_real = url_orc, url_real
    except Exception:
        path_orc = path_orc_local if os.path.exists(path_orc_local) else "ORCAMENTO 2026 - REV.1.xlsx"
        path_real = path_real_local if os.path.exists(path_real_local) else "REALIZADO 2026.xlsx"
        xls_orc = pd.ExcelFile(path_orc)

    abas_ignorar = ["Sint Ebt loja", "CONS 25X26 V.1", "CONS 25X26 V.2"]
    abas_validas = [s for s in xls_orc.sheet_names if s not in abas_ignorar]

    return abas_validas, path_orc, path_real

try:
    abas_disponiveis, path_orc, path_real = obter_caminhos_excel()
except Exception as e:
    st.error(f"Erro no carregamento das planilhas de origem: {e}")
    st.stop()

@st.cache_data(ttl=60)
def carregar_dados_abas(path_o, path_r, lista_abas):
    dfs_o, dfs_r = [], []
    for aba in lista_abas:
        try:
            dfs_o.append(pd.read_excel(path_o, sheet_name=aba))
            dfs_r.append(pd.read_excel(path_r, sheet_name=aba))
        except Exception:
            continue
    return dfs_o, dfs_r

# ==============================================================================
# 3. FILTROS E NAVEGAÇÃO LATERAL
# ==============================================================================
st.sidebar.markdown("### 🎛️ Filtros Executivos")

modo_visao = st.sidebar.radio("Modo de Visão:", ["Visão Consolidada", "Selecionar Unidade"])

abas_consolidadas_permitidas = [
    "DRE CONSOLIDADO", "ABPR CONSOLIDADO", "VD CONSOLIDADO", "LJ CONSOLIDADO", "ABPR + VD", "LJ - G&A"
]

opcoes_consolidadas = [a for a in abas_disponiveis if a in abas_consolidadas_permitidas] or abas_disponiveis
opcoes_unidades = [a for a in abas_disponiveis if a not in abas_consolidadas_permitidas] or abas_disponiveis

if modo_visao == "Visão Consolidada":
    visao_sel = st.sidebar.selectbox("Visão:", opcoes_consolidadas)
    abas_para_carregar = [visao_sel]
    label_visao = visao_sel
else:
    lojas_sel = st.sidebar.multiselect("Unidades:", options=opcoes_unidades, default=opcoes_unidades[:3])
    if not lojas_sel:
        st.warning("Selecione ao menos uma unidade para prosseguir.")
        st.stop()
    abas_para_carregar = lojas_sel
    label_visao = f"Soma de {len(lojas_sel)} Unidades"

list_df_orc, list_df_real = carregar_dados_abas(path_orc, path_real, abas_para_carregar)

if not list_df_real or not list_df_orc:
    st.error("Falha ao carregar as tabelas financeiras selecionadas.")
    st.stop()

# Configuração de Período do Ano
nomes_meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
meses_cols = [f"{i:02d}/2026" for i in range(1, 13)]

df_ref = list_df_real[0]
colunas_validas = [m for m in meses_cols if m in df_ref.columns]
m_map = {nome: col for nome, col in zip(nomes_meses, meses_cols) if col in colunas_validas}

tipo_periodo = st.sidebar.radio("Período de Análise:", ["Mês Selecionado", "Múltiplos Meses", "ANO COMPLETO (2026)"])

if tipo_periodo == "ANO COMPLETO (2026)":
    cols_kpi = list(m_map.values())
    cols_graficos = list(m_map.values())
    label_periodo_kpi = "Ano Completo (2026)"
    label_periodo_graf = "Ano Completo (2026)"
elif tipo_periodo == "Mês Selecionado":
    idx_default = min(6, len(m_map) - 1)
    mes_ref = st.sidebar.selectbox("Mês de Referência:", list(m_map.keys()), index=idx_default)
    idx = list(m_map.keys()).index(mes_ref)
    cols_kpi = list(m_map.values())[:idx + 1]
    cols_graficos = [m_map[mes_ref]]
    label_periodo_kpi = f"Acumulado YTD até {mes_ref}"
    label_periodo_graf = f"Mês de {mes_ref}"
else:
    meses_mult = st.sidebar.multiselect("Selecione os Meses:", list(m_map.keys()), default=list(m_map.keys())[:min(7, len(m_map))])
    cols_kpi = [m_map[m] for m in meses_mult if m in m_map]
    cols_graficos = cols_kpi
    label_periodo_kpi = "Período Personalizado"
    label_periodo_graf = "Período Personalizado"

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Atualizar Cache", use_container_width=True):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

# ==============================================================================
# 4. FUNÇÕES DE SUPORTE E CALCULO FINANCEIRO (VETORIZADO)
# ==============================================================================
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
            total += dados_num.to_numpy().sum()
    return float(total)

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
    return f"color: {color}; font-weight: 600;"

LAYOUT_NEUTRO = dict(
    paper_bgcolor="rgba(22, 27, 34, 0)",
    plot_bgcolor="rgba(22, 27, 34, 0)",
    font=dict(color="#8B949E", family="Segoe UI, sans-serif"),
    margin=dict(l=20, r=20, t=30, b=50),
)

# ==============================================================================
# 5. HEADER & KPIS EXECUTIVOS
# ==============================================================================
st.markdown(
    f"""
    <div class="main-header">
        <h3>CONTROLADORIA EXECUTIVA B&A</h3>
        <p>Visão: <b>{label_visao}</b> &nbsp;|&nbsp; Período de Análise: <b>{label_periodo_kpi}</b></p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Cálculo dos Métricas do Topo
rec_liq_real = get_valor_consolidado_multi(list_df_real, "3 - Receita Operacional Liquida", cols_kpi)
rec_liq_orc = get_valor_consolidado_multi(list_df_orc, "3 - Receita Operacional Liquida", cols_kpi)

ebitda_real = get_valor_consolidado_multi(list_df_real, "11 - EBITDA", cols_kpi)
ebitda_orc = get_valor_consolidado_multi(list_df_orc, "11 - EBITDA", cols_kpi)

margem_ebitda = (ebitda_real / rec_liq_real * 100) if rec_liq_real != 0 else 0.0
diff_ebitda = ebitda_real - ebitda_orc
pct_diff_ebitda = (diff_ebitda / abs(ebitda_orc) * 100) if ebitda_orc != 0 else 0.0

# Novas Métricas de Eficiência Operacional
rec_bruta_real = get_valor_consolidado_multi(list_df_real, "1 - Receita Operacional Bruta", cols_kpi)
cmv_real = abs(get_valor_consolidado_multi(list_df_real, "4.1 - Custo da Mercadoria", cols_kpi))
if cmv_real == 0:
    cmv_real = abs(get_valor_consolidado_multi(list_df_real, "4 - Custo das Vendas", cols_kpi))

margem_bruta_rs = rec_liq_real - cmv_real
pct_margem_bruta = (margem_bruta_rs / rec_liq_real * 100) if rec_liq_real != 0 else 0.0

desp_op_real = abs(get_valor_consolidado_multi(list_df_real, "8 - Despesas Operacionais", cols_kpi))
opex_ratio = (desp_op_real / rec_liq_real * 100) if rec_liq_real != 0 else 0.0

pct_vendas_prog = min(100.0, max(0.0, (rec_liq_real / rec_liq_orc * 100))) if rec_liq_orc > 0 else 0.0
pct_lucro_prog = min(100.0, max(0.0, (ebitda_real / ebitda_orc * 100))) if ebitda_orc > 0 else 0.0

k1, k2, k3, k4, k5 = st.columns(5)

with k1:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">RECEITA LÍQUIDA</div>
            <div class="kpi-value" style="color: {'#3FB950' if rec_liq_real >= 0 else '#F85149'};">{formata_brl(rec_liq_real)}</div>
            <div class="kpi-subtext">Meta: {formata_brl(rec_liq_orc)}</div>
            <div class="progress-container"><div class="progress-bar" style="width: {pct_vendas_prog:.1f}%;"></div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k2:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">EBITDA REALIZADO</div>
            <div class="kpi-value" style="color: {'#3FB950' if ebitda_real >= 0 else '#F85149'};">{formata_brl(ebitda_real)}</div>
            <div class="kpi-subtext">Meta: {formata_brl(ebitda_orc)}</div>
            <div class="progress-container"><div class="progress-bar" style="width: {pct_lucro_prog:.1f}%;"></div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k3:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">MARGEM EBITDA %</div>
            <div class="kpi-value" style="color: {'#3FB950' if margem_ebitda >= 0 else '#F85149'};">{margem_ebitda:.1f}%</div>
            <div class="kpi-subtext" style="color: {'#3FB950' if diff_ebitda >= 0 else '#F85149'};">{pct_diff_ebitda:+.1f}% vs Orçamento</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k4:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">MARGEM BRUTA %</div>
            <div class="kpi-value" style="color: #388BFD;">{pct_margem_bruta:.1f}%</div>
            <div class="kpi-subtext">Lucro Bruto: {formata_m(margem_bruta_rs)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k5:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">OPEX RATIO %</div>
            <div class="kpi-value" style="color: #F2994A;">{opex_ratio:.1f}%</div>
            <div class="kpi-subtext">Eficiência Estrutural</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ==============================================================================
# 6. MODULOS / TABS
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Dashboards & Indicadores",
    "📋 DRE Gerencial & Desvios",
    "📅 Visão Mensal Comparativa",
    "🔮 Projeções & Trend Analytics"
])

# ------------------------------------------------------------------------------
# TAB 1: DASHBOARDS
# ------------------------------------------------------------------------------
with tab1:
    cg1, cg2 = st.columns(2)

    with cg1:
        st.markdown("##### **Demonstrativo em Ponte (Waterfall) - YTD**")

        rec_bruta = get_valor_consolidado_multi(list_df_real, "1 - Receita Operacional Bruta", cols_kpi)
        deducoes = get_valor_consolidado_multi(list_df_real, "2 - Deduções da Receita", cols_kpi)
        rec_liq = get_valor_consolidado_multi(list_df_real, "3 - Receita Operacional Liquida", cols_kpi)
        cmv_b = get_valor_consolidado_multi(list_df_real, "4 - ", cols_kpi, exato_linha_sintetica=True) or get_valor_consolidado_multi(list_df_real, "4 - Custo das Vendas", cols_kpi)
        mb_b = get_valor_consolidado_multi(list_df_real, "5 - Margem de Contribuição 1", cols_kpi)
        desp_v = get_valor_consolidado_multi(list_df_real, "6 - Despesas Variáveis", cols_kpi)
        desp_o = get_valor_consolidado_multi(list_df_real, "8 - Despesas Operacionais", cols_kpi)
        deprec = get_valor_consolidado_multi(list_df_real, "13 - Depreciação e Amortização", cols_kpi)
        ebitda_b = get_valor_consolidado_multi(list_df_real, "11 - EBITDA", cols_kpi)

        base = rec_liq if rec_liq != 0 else 1.0
        p_rb, p_ded, p_cmv, p_mb = round((rec_bruta/base)*100), round((deducoes/base)*100), round((-abs(cmv_b)/base)*100), round((mb_b/base)*100)
        p_sga = round((-(abs(desp_v)+abs(desp_o))/base)*100)
        p_ebitda_b = round((ebitda_b/base)*100)

        fig_waterfall = go.Figure(
            go.Waterfall(
                orientation="v",
                measure=["absolute", "relative", "total", "relative", "total", "relative", "total"],
                x=["Receita Bruta", "Deduções", "Receita Líquida", "CMV", "Margem Bruta", "SG&A", "EBITDA"],
                y=[p_rb, p_ded, 0, p_cmv, 0, p_sga, 0],
                text=[f"{p_rb}%", f"{p_ded}%", "100%", f"{p_cmv}%", f"{p_mb}%", f"{p_sga}%", f"{p_ebitda_b}%"],
                textposition="outside",
                connector={"line": {"color": "#21262D"}},
                decreasing={"marker": {"color": "#F85149"}},
                increasing={"marker": {"color": "#388BFD"}},
                totals={"marker": {"color": "#238636"}},
            )
        )
        fig_waterfall.update_layout(**LAYOUT_NEUTRO, height=380, xaxis=dict(fixedrange=True), yaxis=dict(showticklabels=False, fixedrange=True))
        st.plotly_chart(fig_waterfall, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

    with cg2:
        st.markdown("##### **Realizado vs. Orçado por Grupo Financeiro**")

        cats = ["CMV", "Margem Contrib 2", "OpEx / Despesas", "EBITDA"]
        cmv_r = abs(get_valor_consolidado_multi(list_df_real, "4.1 - Custo", cols_kpi) or get_valor_consolidado_multi(list_df_real, "4 - Custo", cols_kpi))
        cmv_o = abs(get_valor_consolidado_multi(list_df_orc, "4.1 - Custo", cols_kpi) or get_valor_consolidado_multi(list_df_orc, "4 - Custo", cols_kpi))
        
        mc_r = get_valor_consolidado_multi(list_df_real, "7 - Margem de Contribuição 2", cols_kpi)
        mc_o = get_valor_consolidado_multi(list_df_orc, "7 - Margem de Contribuição 2", cols_kpi)

        opex_r = abs(get_valor_consolidado_multi(list_df_real, "8 - Despesas Operacionais", cols_kpi))
        opex_o = abs(get_valor_consolidado_multi(list_df_orc, "8 - Despesas Operacionais", cols_kpi))

        eb_r = get_valor_consolidado_multi(list_df_real, "11 - EBITDA", cols_kpi)
        eb_o = get_valor_consolidado_multi(list_df_orc, "11 - EBITDA", cols_kpi)

        fig_bar = go.Figure(data=[
            go.Bar(name="Realizado", x=cats, y=[cmv_r, mc_r, opex_r, eb_r], text=[formata_m(v) for v in [cmv_r, mc_r, opex_r, eb_r]], textposition="outside", marker_color="#1F6FEB"),
            go.Bar(name="Orçado", x=cats, y=[cmv_o, mc_o, opex_o, eb_o], text=[formata_m(v) for v in [cmv_o, mc_o, opex_o, eb_o]], textposition="outside", marker_color="#30363D")
        ])
        fig_bar.update_layout(barmode="group", **LAYOUT_NEUTRO, height=380, legend=dict(orientation="h", y=-0.2, x=0.3), yaxis=dict(showticklabels=False, fixedrange=True))
        st.plotly_chart(fig_bar, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)

# ------------------------------------------------------------------------------
# TAB 2: DRE GERENCIAL & DESVIOS COM EXPORTAÇÃO
# ------------------------------------------------------------------------------
with tab2:
    st.markdown(f"##### 📋 **Demonstrativo do Resultado do Exercício - {label_periodo_graf}**")

    c_f1, c_f2 = st.columns([2, 1])
    with c_f1:
        tipo_visao_dre = st.radio("Visão da DRE:", ["Apenas Sintética (Grupos)", "Analítica Completa"], horizontal=True)

    col_nome = "Nome" if "Nome" in df_ref.columns else df_ref.columns[0]
    linhas_dre = df_ref[col_nome].dropna().astype(str).unique()

    if tipo_visao_dre == "Apenas Sintética (Grupos)":
        linhas_dre = [l for l in linhas_dre if eh_grupo_sintetico(l)]

    dados_dre = []
    for linha in linhas_dre:
        v_real = get_valor_consolidado_multi(list_df_real, linha, cols_graficos)
        v_orc = get_valor_consolidado_multi(list_df_orc, linha, cols_graficos)
        desvio = v_real - v_orc
        av_real = (v_real / rec_bruta_real * 100) if rec_bruta_real != 0 else 0.0
        av_orc = (v_orc / rec_liq_orc * 100) if rec_liq_orc != 0 else 0.0
        ah = (desvio / abs(v_orc) * 100) if v_orc != 0 else 0.0

        dados_dre.append({
            "Estrutura DRE": linha,
            "Realizado (R$)": v_real,
            "AV Real (%)": av_real,
            "Orçado (R$)": v_orc,
            "AV Orçado (%)": av_orc,
            "Desvio Absoluto (R$)": desvio,
            "Variação AH (%)": ah,
        })

    df_dre_final = pd.DataFrame(dados_dre)

    # Botão de Exportação para Excel
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_dre_final.to_excel(writer, sheet_name='DRE_Gerencial', index=False)
    
    st.download_button(
        label="📥 Exportar DRE para Excel",
        data=output_excel.getvalue(),
        file_name=f"DRE_BA_Controladoria_{label_periodo_graf}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    cols_num_dre = ["Realizado (R$)", "AV Real (%)", "Orçado (R$)", "AV Orçado (%)", "Desvio Absoluto (R$)", "Variação AH (%)"]

    st.dataframe(
        df_dre_final.style.format({
            "Realizado (R$)": formata_brl,
            "AV Real (%)": "{:.1f}%",
            "Orçado (R$)": formata_brl,
            "AV Orçado (%)": "{:.1f}%",
            "Desvio Absoluto (R$)": formata_brl,
            "Variação AH (%)": "{:.1f}%",
        }).map(cor_valor, subset=cols_num_dre),
        use_container_width=True,
        height=550,
        hide_index=True
    )

# ------------------------------------------------------------------------------
# TAB 3: VISÃO MENSAL COMPARATIVA
# ------------------------------------------------------------------------------
with tab3:
    st.markdown("##### 📅 **Evolução do Desempenho Mensal Mês a Mês**")

    tipo_hist = st.radio("Selecione a Base:", ["Realizado", "Orçado"], horizontal=True)
    linhas_hist = [l for l in df_ref[col_nome].dropna().astype(str).unique() if eh_grupo_sintetico(l)]

    target_dfs = list_df_real if tipo_hist == "Realizado" else list_df_orc
    hist_data = []

    for linha in linhas_hist:
        row_dict = {"Estrutura DRE": linha}
        soma_linha = 0.0
        for m_nome, m_col in m_map.items():
            val_m = get_valor_consolidado_multi(target_dfs, linha, [m_col])
            row_dict[m_nome] = val_m
            soma_linha += val_m
        row_dict["Acumulado"] = soma_linha
        hist_data.append(row_dict)

    df_hist = pd.DataFrame(hist_data)
    cols_num_hist = list(m_map.keys()) + ["Acumulado"]

    st.dataframe(
        df_hist.style.format({col: formata_brl for col in cols_num_hist}).map(cor_valor, subset=cols_num_hist),
        use_container_width=True,
        height=550,
        hide_index=True
    )

# ------------------------------------------------------------------------------
# TAB 4: PROJEÇÕES & TREND ANALYTICS (SIMULADOR WHAT-IF)
# ------------------------------------------------------------------------------
with tab4:
    st.markdown("##### 🔮 **Modelo Preditivo & Simulador de Cenários ("What-If")**")

    c_f1, c_f2, c_f3 = st.columns([1.2, 1.2, 1.6])
    with c_f1:
        metrica_sel = st.selectbox("Indicador Projetado:", ["Receita Operacional Líquida", "EBITDA"])
        termo_metrica = "3 - Receita Operacional Liquida" if metrica_sel == "Receita Operacional Líquida" else "11 - EBITDA"

    with c_f2:
        modelo_proj = st.selectbox("Algoritmo de Projeção:", ["Run-Rate Histórico", "Manter Budget Original", "Ajustado por Performance"])

    with c_f3:
        sensibilidade = st.slider("Ajuste Fino / Estresse de Cenário (%):", min_value=-25.0, max_value=25.0, value=0.0, step=1.0)

    meses_todos = list(m_map.keys())
    meses_realizados = list(m_map.values())[:len(cols_kpi)]
    num_m_realizados = len(meses_realizados)

    val_real_acum = get_valor_consolidado_multi(list_df_real, termo_metrica, meses_realizados)
    val_orc_acum = get_valor_consolidado_multi(list_df_orc, termo_metrica, meses_realizados)
    val_orc_anual = get_valor_consolidado_multi(list_df_orc, termo_metrica, colunas_validas)

    media_mensal_real = val_real_acum / num_m_realizados if num_m_realizados > 0 else 0.0
    fator_perf = (val_real_acum / val_orc_acum) if val_orc_acum != 0 else 1.0
    fator_sens = 1.0 + (sensibilidade / 100.0)

    valores_finais, valores_orcado, tipos_serie = [], [], []

    for idx_m, m_nome in enumerate(meses_todos):
        m_col = m_map[m_nome]
        v_orc = get_valor_consolidado_multi(list_df_orc, termo_metrica, [m_col])
        valores_orcado.append(v_orc)

        if idx_m < num_m_realizados:
            valores_finais.append(get_valor_consolidado_multi(list_df_real, termo_metrica, [m_col]))
            tipos_serie.append("Realizado")
        else:
            if modelo_proj == "Run-Rate Histórico":
                v_proj = media_mensal_real * fator_sens
            elif modelo_proj == "Manter Budget Original":
                v_proj = v_orc * fator_sens
            else:
                v_proj = (v_orc * fator_perf) * fator_sens
            valores_finais.append(v_proj)
            tipos_serie.append("Projetado")

    projecao_anual = sum(valores_finais)
    diff_anual = projecao_anual - val_orc_anual

    fig_comb = go.Figure()
    fig_comb.add_trace(go.Bar(x=[m.capitalize() for m in meses_todos], y=valores_finais, name="Projetado / Real", marker_color="#1F6FEB"))
    fig_comb.add_trace(go.Scatter(x=[m.capitalize() for m in meses_todos], y=valores_orcado, name="Meta (Orçamento)", line=dict(color="#F2994A", width=2, dash="dash")))
    
    fig_comb.update_layout(**LAYOUT_NEUTRO, height=400, barmode="group", legend=dict(orientation="h", y=-0.2, x=0.4))
    st.plotly_chart(fig_comb, use_container_width=True, config=CONFIG_PLOTLY_TRAVADO)