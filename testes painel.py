"""
Testes do Painel Controladoria/Financeiro — Grupo B&A
=====================================================

Rodar fora do Streamlit, na pasta do projeto:

    python testes_painel.py            # usa app.py
    python testes_painel.py outro.py   # ou o arquivo que quiser

O que estes testes protegem, em uma frase: os erros que NÃO aparecem na tela.
Cada bloco aqui nasceu de um problema real que passou despercebido porque o
painel continuou abrindo normalmente, só com o número errado.

Como funciona: as funções são extraídas do arquivo do app pela árvore de
sintaxe (sem importar o módulo, o que exigiria o Streamlit rodando) e
executadas com dados montados à mão.
"""
import ast
import base64
import hashlib
import hmac
import io
import re
import sys
import time
import unittest
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd

CAMINHO_APP = sys.argv[1] if len(sys.argv) > 1 else "app.py"
try:
    FONTE = open(CAMINHO_APP, encoding="utf-8").read()
except FileNotFoundError:
    print(f"Nao encontrei {CAMINHO_APP}. Rode na pasta do projeto ou passe o caminho.")
    raise SystemExit(1)

ARVORE = ast.parse(FONTE)
FUSO_BR = ZoneInfo("America/Sao_Paulo")


def carregar(nomes_funcoes, nomes_constantes=(), extras=None):
    """Executa apenas as funcoes e constantes pedidas, num espaco isolado."""
    pedacos = []
    for no in ARVORE.body:
        if isinstance(no, ast.FunctionDef) and no.name in nomes_funcoes:
            pedacos.append(ast.get_source_segment(FONTE, no))
        elif isinstance(no, ast.Assign):
            for alvo in no.targets:
                if isinstance(alvo, ast.Name) and alvo.id in nomes_constantes:
                    pedacos.append(ast.get_source_segment(FONTE, no))
    espaco = {
        "pd": pd, "re": re, "io": io, "time": time, "datetime": datetime,
        "base64": base64, "hashlib": hashlib, "hmac": hmac, "FUSO_BR": FUSO_BR,
    }
    espaco.update(extras or {})
    exec("\n\n".join(pedacos), espaco)
    faltando = [n for n in nomes_funcoes if n not in espaco]
    if faltando:
        raise SystemExit(f"Nao encontrei no {CAMINHO_APP}: {', '.join(faltando)}")
    return espaco


# ============================================================================
# 1. LEITURA DE DATAS - o erro que trocou dia por mes em 187 mil lancamentos
# ============================================================================
class TesteOrdemDaData(unittest.TestCase):
    """O CSV do fluxo vem em mes/dia/ano. Quem adivinha pelo primeiro valor
    erra quando ele e ambiguo (1/2/2026), e a coluna inteira vira outra coisa
    sem nenhum aviso na tela."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(["_ordem_data_fin"])

    def test_mes_primeiro_quando_segundo_componente_passa_de_12(self):
        serie = pd.Series(["12/31/2026", "1/2/2026", "3/20/2026"])
        self.assertFalse(self.ns["_ordem_data_fin"](serie))

    def test_dia_primeiro_quando_primeiro_componente_passa_de_12(self):
        serie = pd.Series(["31/12/2026", "02/01/2026", "20/03/2026"])
        self.assertTrue(self.ns["_ordem_data_fin"](serie))

    def test_valor_ambiguo_no_comeco_nao_engana(self):
        serie = pd.Series(["1/2/2026", "1/5/2026", "12/29/2025"])
        self.assertFalse(self.ns["_ordem_data_fin"](serie))

    def test_tudo_ambiguo_mantem_comportamento_antigo(self):
        serie = pd.Series(["1/2/2026", "3/4/2026"])
        self.assertTrue(self.ns["_ordem_data_fin"](serie))

    def test_serie_vazia_nao_quebra(self):
        self.assertTrue(self.ns["_ordem_data_fin"](pd.Series([], dtype=object)))


# ============================================================================
# 2. SALDO INICIAL DO FLUXO DIARIO
# ============================================================================
class TesteSaldoInicialDiario(unittest.TestCase):
    """Dia sem caixa/banco preenchido herda o fechamento do dia anterior; dia
    com posicao preenchida zera o saldo inicial (senao o dinheiro e contado
    duas vezes)."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(["calcular_saldo_inicial_diario"])

    def _base(self):
        def linha(dia, tipo, valor):
            return {"DiaOrd": pd.Timestamp(f"2026-08-{dia:02d}"),
                    "Tipo Movimento": tipo, "Valor.1": valor}
        return pd.DataFrame([
            linha(13, "saldo", 4_000_000.0), linha(13, "entrada", 300_000.0),
            linha(13, "saida", -55_330.09),
            linha(14, "saldo", 4_236_000.0), linha(14, "entrada", 481_290.60),
            linha(14, "saida", -236_433.50),
            linha(15, "saldo", 0.0), linha(15, "saida", -65_906.77),
            linha(16, "saldo", 0.0), linha(16, "saida", -9_195.12),
        ])

    def test_dia_com_posicao_zera_o_saldo_inicial(self):
        ini, _ = self.ns["calcular_saldo_inicial_diario"](self._base(), "Valor.1")
        self.assertEqual(ini[pd.Timestamp("2026-08-13")], 0.0)
        self.assertEqual(ini[pd.Timestamp("2026-08-14")], 0.0)

    def test_dia_sem_posicao_herda_o_fechamento_anterior(self):
        ini, tot = self.ns["calcular_saldo_inicial_diario"](self._base(), "Valor.1")
        self.assertAlmostEqual(ini[pd.Timestamp("2026-08-15")],
                               tot[pd.Timestamp("2026-08-14")], places=2)
        self.assertAlmostEqual(ini[pd.Timestamp("2026-08-16")],
                               tot[pd.Timestamp("2026-08-15")], places=2)

    def test_total_do_dia_soma_saldo_inicial_e_movimento(self):
        ini, tot = self.ns["calcular_saldo_inicial_diario"](self._base(), "Valor.1")
        d15 = pd.Timestamp("2026-08-15")
        self.assertAlmostEqual(tot[d15], ini[d15] - 65_906.77, places=2)

    def test_acumulo_atravessa_meses(self):
        linhas = [{"DiaOrd": pd.Timestamp("2026-08-14"), "Tipo Movimento": "saldo",
                   "Valor.1": 1_000_000.0}]
        for mes, dias in ((8, range(15, 32)), (9, range(1, 31))):
            for dia in dias:
                linhas.append({"DiaOrd": pd.Timestamp(2026, mes, dia),
                               "Tipo Movimento": "saida", "Valor.1": -1_000.0})
        ini, tot = self.ns["calcular_saldo_inicial_diario"](pd.DataFrame(linhas), "Valor.1")
        self.assertAlmostEqual(ini[pd.Timestamp("2026-09-01")],
                               tot[pd.Timestamp("2026-08-31")], places=2)


# ============================================================================
# 3. ALERTAS DO FLUXO DE CAIXA
# ============================================================================
class TesteAlertasDoFluxo(unittest.TestCase):
    """Alerta que dispara sempre vira ruido e passa a ser ignorado - por isso
    cada regra e testada no cenario que deve disparar E no que nao deve."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(
            ["_avaliar_alertas_fluxo", "_saldo_posicao_atual_fin", "formata_brl"],
            ["META_RESERVA_PADRAO", "HORIZONTE_ALERTA_SEMANAS", "COL_FIN_VALOR",
             "COL_FIN_VENCIMENTO", "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_LIQ_AMPLA",
             "COL_FIN_LIQ_DIARIO", "COL_FIN_CANAL"],
        )
        cls.hoje = pd.Timestamp(datetime.now(FUSO_BR).date())

    def _linha(self, canal, tipo, dias, valor):
        data = self.hoje + pd.Timedelta(days=dias)
        return {"Canal.1": canal, "Tipo Movimento": tipo, "Data Efetiva": data,
                "Valor.1": valor, "Vencimento.1": data, "Data Liquidação": pd.NaT}

    def _rodar(self, linhas, **kw):
        df = pd.DataFrame(linhas)
        df["Data Liquidação"] = pd.to_datetime(df["Data Liquidação"])
        p = dict(meta=30, vencido=50_000, concentracao=60, horizonte=30)
        p.update(kw)
        return self.ns["_avaliar_alertas_fluxo"](
            df, "Valor.1", p["meta"], p["vencido"], p["concentracao"],
            horizonte_canal_dias=p["horizonte"],
        )

    def test_empresa_saudavel_nao_dispara_nada(self):
        linhas = [self._linha("LOJA", "saldo", 0, 8_000_000.0),
                  self._linha("LOJA", "entrada", 3, 2_000_000.0)]
        linhas += [self._linha("LOJA", "saida", d, -60_000.0) for d in range(1, 13)]
        self.assertEqual(self._rodar(linhas), [])

    def test_vencido_em_aberto_dispara(self):
        linhas = [self._linha("LOJA", "saldo", 0, 5_000_000.0)]
        linhas += [self._linha("LOJA", "saida", -20, -40_000.0) for _ in range(3)]
        titulos = [a["titulo"] for a in self._rodar(linhas)]
        self.assertTrue(any("vencidos" in t for t in titulos), titulos)

    def test_vencido_abaixo_do_limite_nao_dispara(self):
        linhas = [self._linha("LOJA", "saldo", 0, 5_000_000.0),
                  self._linha("LOJA", "saida", -20, -1_000.0)]
        titulos = [a["titulo"] for a in self._rodar(linhas)]
        self.assertFalse(any("vencidos" in t for t in titulos), titulos)

    def test_pagamentos_da_semana_acima_do_caixa(self):
        linhas = [self._linha("LOJA", "saldo", 0, 100_000.0),
                  self._linha("LOJA", "saida", 3, -500_000.0)]
        niveis = {a["nivel"] for a in self._rodar(linhas)}
        self.assertIn("critico", niveis)

    def test_concentracao_exige_minimo_de_dias(self):
        linhas = [self._linha("LOJA", "saldo", 0, 9_000_000.0)]
        linhas += [self._linha("LOJA", "saida", d, -50_000.0) for d in range(1, 4)]
        titulos = [a["titulo"] for a in self._rodar(linhas)]
        self.assertFalse(any("desembolso do" in t for t in titulos), titulos)

    def test_canal_sem_caixa_proprio_e_atencao_quando_a_empresa_cobre(self):
        linhas = [self._linha("LOJA", "saldo", 0, 50_000.0),
                  self._linha("VENDA DIRETA", "saldo", 0, 5_000_000.0),
                  self._linha("LOJA", "saida", 5, -400_000.0)]
        canal = [a for a in self._rodar(linhas) if "sem recurso" in a["titulo"]]
        self.assertEqual(len(canal), 1)
        self.assertEqual(canal[0]["nivel"], "atencao")

    def test_canal_e_empresa_sem_caixa_e_critico(self):
        linhas = [self._linha("LOJA", "saldo", 0, 50_000.0),
                  self._linha("VENDA DIRETA", "saldo", 0, 80_000.0),
                  self._linha("LOJA", "saida", 5, -400_000.0)]
        canal = [a for a in self._rodar(linhas) if "sem recurso" in a["titulo"]]
        self.assertEqual(len(canal), 1)
        self.assertEqual(canal[0]["nivel"], "critico")

    def test_canal_sem_posicao_propria_nao_e_avaliado(self):
        linhas = [self._linha("VENDA DIRETA", "saldo", 0, 9_000_000.0),
                  self._linha("HUB LOGISTICO", "saida", 4, -80_000.0)]
        canal = [a for a in self._rodar(linhas) if "sem recurso" in a["titulo"]]
        self.assertEqual(canal, [])


# ============================================================================
# 4. SESSAO QUE SOBREVIVE AO F5
# ============================================================================
class TesteBilheteDeSessao(unittest.TestCase):
    """O bilhete viaja na URL. Se pudesse ser forjado ou ter o prazo esticado
    a mao, qualquer pessoa entraria sem senha."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(
            ["_assinar_sessao", "gerar_bilhete_sessao", "ler_bilhete_sessao"],
            ["MINUTOS_INATIVIDADE", "CHAVE_SESSAO_URL"],
            extras={"_segredo_sessao": lambda: "segredo-de-teste"},
        )

    def test_bilhete_valido_devolve_o_email(self):
        b = self.ns["gerar_bilhete_sessao"]("Controladoria@GrupoBeea.com.br ")
        self.assertEqual(self.ns["ler_bilhete_sessao"](b), "controladoria@grupobeea.com.br")

    def test_bilhete_forjado_e_recusado(self):
        falso = base64.urlsafe_b64encode(b"outro@empresa.com|9999999999|" + b"0" * 32).decode()
        self.assertIsNone(self.ns["ler_bilhete_sessao"](falso))

    def test_prazo_esticado_a_mao_e_recusado(self):
        b = self.ns["gerar_bilhete_sessao"]("controladoria@grupobeea.com.br")
        email, exp, assinatura = base64.urlsafe_b64decode(b).decode().rsplit("|", 2)
        esticado = base64.urlsafe_b64encode(
            f"{email}|{int(exp) + 99999}|{assinatura}".encode()).decode()
        self.assertIsNone(self.ns["ler_bilhete_sessao"](esticado))

    def test_bilhete_vencido_e_recusado(self):
        corpo = f"controladoria@grupobeea.com.br|{int(time.time()) - 60}"
        vencido = base64.urlsafe_b64encode(
            f"{corpo}|{self.ns['_assinar_sessao'](corpo)}".encode()).decode()
        self.assertIsNone(self.ns["ler_bilhete_sessao"](vencido))

    def test_lixo_e_vazio_nao_quebram(self):
        for valor in ("abc123", "", None, "!!!"):
            self.assertIsNone(self.ns["ler_bilhete_sessao"](valor))

    def test_bilhete_nao_carrega_a_senha(self):
        b = self.ns["gerar_bilhete_sessao"]("controladoria@grupobeea.com.br")
        self.assertNotIn("senha", base64.urlsafe_b64decode(b).decode().lower())


# ============================================================================
# 5. RELATORIOS: PLANOS DE CONTAS
# ============================================================================
class TestePlanosDeContas(unittest.TestCase):
    """Cada modelo so pode oferecer os planos que pertencem a ele - oferecer a
    lista inteira seria dar a um departamento acesso a conta de outro."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(["planos_do_diario", "resolver_planos_forcados",
                           "_normalizar_texto"])
        cls.diario = pd.DataFrame({
            "Plano de Contas": ["Mercadorias", "Material de Embalagem", "Flaconetes",
                                "Benfeitorias em Imóveis de Terceiros", "Combustível"],
            "Linha DRE": ["", "6.6 - Material de Embalagem", "6.14 - Flaconetes",
                          "", "8.5.3 - Combustível"],
            "Valor Bruto": [-1.0] * 5,
        })

    def test_lista_restrita_as_linhas_do_modelo(self):
        planos = self.ns["planos_do_diario"](self.diario, ["6.14 - Flaconetes"])
        self.assertEqual(planos, ["Flaconetes"])

    def test_sem_linhas_traz_todos(self):
        self.assertEqual(len(self.ns["planos_do_diario"](self.diario)), 5)

    def test_plano_sem_linha_da_dre_e_achado_pelo_nome(self):
        achados, faltando = self.ns["resolver_planos_forcados"](self.diario, ["Mercadorias"])
        self.assertEqual(achados, ["Mercadorias"])
        self.assertEqual(faltando, [])

    def test_busca_tolera_caixa_e_nome_parcial(self):
        for termo in ("mercadorias", "MERCADORIAS", "Benfeitorias em Imóveis"):
            achados, _ = self.ns["resolver_planos_forcados"](self.diario, [termo])
            self.assertTrue(achados, termo)

    def test_termo_inexistente_e_reportado(self):
        achados, faltando = self.ns["resolver_planos_forcados"](self.diario, ["Conta Que Sumiu"])
        self.assertEqual(achados, [])
        self.assertEqual(faltando, ["Conta Que Sumiu"])


# ============================================================================
# 5b. LEITURA DE LINHA DA DRE POR CODIGO
# ============================================================================
class TesteSomaComFilhas(unittest.TestCase):
    """A linha-mae nem sempre esta preenchida: em algumas abas o valor so
    existe nas filhas. Ler so a mae devolvia zero em silencio."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(["_valor_linha_por_codigo", "_somar_codigo_com_filhas"])

    def _aba(self, linhas):
        return pd.DataFrame([{"Nome": n, "01/2026": v} for n, v in linhas])

    def test_usa_a_mae_quando_ela_existe(self):
        aba = self._aba([("6.24.2 - Marketing", -90_000.0), ("6.24.2.1 - Midia", -50_000.0)])
        self.assertEqual(
            self.ns["_somar_codigo_com_filhas"](aba, "6.24.2", ["01/2026"]), -90_000.0)

    def test_desce_para_as_filhas_quando_a_mae_falta(self):
        aba = self._aba([("6.24.2.1 - Midia", -30_000.0), ("6.24.2.2 - Eventos", -20_000.0)])
        self.assertEqual(
            self.ns["_somar_codigo_com_filhas"](aba, "6.24.2", ["01/2026"]), -50_000.0)

    def test_nao_soma_neto_junto_com_filho(self):
        aba = self._aba([("6.24.2.1 - Midia", -30_000.0), ("6.24.2.1.1 - Radio", -10_000.0)])
        self.assertEqual(
            self.ns["_somar_codigo_com_filhas"](aba, "6.24.2", ["01/2026"]), -30_000.0)

    def test_sem_a_conta_devolve_zero(self):
        aba = self._aba([("9 - Resultado Operacional", 0.0)])
        self.assertEqual(self.ns["_somar_codigo_com_filhas"](aba, "6.24.2", ["01/2026"]), 0.0)


# ============================================================================
# 5c. GERENCIA COMERCIAL
# ============================================================================
class TesteGerenciaComercial(unittest.TestCase):
    """O bloco reproduz o relatorio que a Controladoria enviava na mao: o
    Quadro de Metas (3 linhas) e o grupo 6 aberto por sublinha."""

    NOME = "📈 Relatório de Custos - Gerência Comercial"

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(
            ["_normalizar_texto", "_normalizar_nome_aba", "get_valor_consolidado_multi",
             "_nome_sem_numero_dre", "_resolver_termo_departamento", "cor_valor",
             "cor_variacao", "formata_brl", "formata_valor_curto", "faixa_metricas_html",
             "html_compacto", "_total_dre", "_tabela_departamento", "_painel_dept_comercial",
             "_numero_linha_dre", "_linha_pertence_ao_grupo", "_linhas_raiz_do_conjunto"],
            ["COLORS", "FONTE_MONO", "LINHA_RECEITA_LIQUIDA", "LINHA_DESPESAS_VARIAVEIS",
             "MODELOS_RELATORIO"],
            extras={"st": type("st", (), {
                "markdown": staticmethod(lambda *a, **k: None),
                "caption": staticmethod(lambda *a, **k: None),
                "dataframe": staticmethod(lambda d, **k: TesteGerenciaComercial.tabelas.append(d)),
            })()},
        )

    def setUp(self):
        TesteGerenciaComercial.tabelas = []

    def _resolver(self, universo):
        """Separa o que a gerencia gere (Quadro de Metas) do que vai junto so
        para conhecimento (grupo 6), como o painel faz."""
        modelo = self.ns["MODELOS_RELATORIO"][self.NOME]
        geridas, informativas = [], []
        for termo in modelo["linhas_dre"]:
            geridas.extend(self.ns["_resolver_termo_departamento"](termo, universo))
        for termo in modelo.get("linhas_informativas", []):
            informativas.extend(self.ns["_resolver_termo_departamento"](termo, universo))
        return list(dict.fromkeys(geridas)), list(dict.fromkeys(informativas))

    def _rodar(self):
        # Valores consolidados (jan a maio) do relatorio enviado pelo gestor.
        real = {"6 - Despesas Variáveis": -6_108_698.41,
                "6.24.2.5 - Encontro de Ciclo": -4_519.40,
                "6.24.2.6 - Outras Despesas de Marketing": -156_203.19,
                "8.3.3.7 - Prêmios / Bônus": -176_115.22,
                "6.1 - Comissões sobre Vendas": -700_432.11,
                "3 - Receita Operacional Liquida": 75_737_289.82}
        orc = {"6 - Despesas Variáveis": -7_000_276.88,
               "6.24.2.5 - Encontro de Ciclo": -28_400.00,
               "6.24.2.6 - Outras Despesas de Marketing": -393_600.00,
               "8.3.3.7 - Prêmios / Bônus": -128_357.03,
               "6.1 - Comissões sobre Vendas": -801_680.80,
               "3 - Receita Operacional Liquida": 91_012_257.49}
        aba = lambda d: pd.DataFrame([{"Nome": n, "01/2026": v} for n, v in d.items()])
        universo = list(real.keys())
        geridas, informativas = self._resolver(universo)
        self.ns["_painel_dept_comercial"]({
            "departamento": self.NOME, "dfs_real": [aba(real)], "dfs_orc": [aba(orc)],
            "colunas": ["01/2026"], "linhas_resolvidas": geridas,
            "linhas_raiz": self.ns["_linhas_raiz_do_conjunto"](geridas),
            "linhas_todas": universo, "linhas_informativas": informativas,
            "path_orc": "o", "path_real": "r"})
        return [t.data for t in TesteGerenciaComercial.tabelas]

    def test_quadro_de_metas_bate_com_o_relatorio(self):
        metas = self._rodar()[0]
        total = metas[metas["Linha da DRE"].str.startswith("TOTAL")].iloc[0]
        self.assertAlmostEqual(total["Realizado (R$)"], 336_837.81, places=2)
        self.assertAlmostEqual(total["Orçado (R$)"], 550_357.03, places=2)

    def test_quadro_de_metas_tem_as_tres_linhas(self):
        metas = self._rodar()[0]
        self.assertEqual(len(metas), 4)  # 3 linhas + TOTAL
        self.assertTrue(metas.iloc[-1]["Linha da DRE"].startswith("TOTAL"))

    def test_grupo_6_nao_soma_no_total_do_departamento(self):
        """O grupo 6 vai junto so para conhecimento. Se ele entrasse nas
        linhas geridas, as duas linhas de marketing (netas da 6) seriam
        contadas duas vezes e o total do departamento nao bateria com o
        'Quadro de Metas GER.COM.' do relatorio."""
        universo = ["6 - Despesas Variáveis", "6.1 - Comissões sobre Vendas",
                    "6.24.2.5 - Encontro de Ciclo",
                    "6.24.2.6 - Outras Despesas de Marketing",
                    "8.3.3.7 - Prêmios / Bônus"]
        geridas, informativas = self._resolver(universo)
        self.assertEqual(len(geridas), 3, geridas)
        self.assertNotIn("6 - Despesas Variáveis", geridas)
        self.assertIn("6 - Despesas Variáveis", informativas)

    def test_grupo_6_fecha_com_o_total(self):
        grupo = self._rodar()[1]
        total = grupo[grupo["Linha da DRE"].str.startswith("TOTAL")].iloc[0]
        self.assertAlmostEqual(total["Realizado (R$)"], 6_108_698.41, places=2)
        self.assertAlmostEqual(total["Orçado (R$)"], 7_000_276.88, places=2)

    def test_modelo_seleciona_o_recorte_do_relatorio(self):
        """Por padrao o relatorio sai com o grupo 6 e as sublinhas diretas
        (6.1 a 6.25) mais as tres linhas do Quadro de Metas - o mesmo recorte
        do arquivo que a Controladoria envia. A 6.24.2 (mae das duas linhas de
        marketing) NAO entra: ela nao esta no relatorio."""
        ns = carregar(["_numero_linha_dre", "_linha_pertence_ao_grupo",
                       "_resolver_termo_departamento", "_linhas_raiz_do_conjunto"])
        dre = ["1 - Receita Operacional Bruta", "6 - Despesas Variáveis"]
        dre += [f"6.{i} - Linha {i}" for i in range(1, 26)]
        dre += ["6.24.2 - Marketing Regional - Gestão CP", "6.24.2.5 - Encontro de Ciclo",
                "6.24.2.6 - Outras Despesas de Marketing", "8.3.3.7 - Prêmios / Bônus"]
        modelo = ["FILHAS:6 - Despesas Variáveis", "6.24.2.5 - Encontro de Ciclo",
                  "6.24.2.6 - Outras Despesas de Marketing", "8.3.3.7 - Prêmios / Bônus"]
        sel = []
        for termo in modelo:
            sel.extend(ns["_resolver_termo_departamento"](termo, dre))
        sel = list(dict.fromkeys(sel))
        self.assertIn("6 - Despesas Variáveis", sel)
        self.assertEqual(sum(1 for l in sel if l.startswith("6.") and l.count(".") == 1), 25)
        self.assertFalse(any(l.startswith("6.24.2 ") for l in sel),
                         "a linha-mae 6.24.2 nao faz parte do relatorio da gerencia")
        self.assertIn("8.3.3.7 - Prêmios / Bônus", sel)
        self.assertEqual(
            sorted(ns["_linhas_raiz_do_conjunto"](sel)),
            sorted(["6 - Despesas Variáveis", "8.3.3.7 - Prêmios / Bônus"]),
            "os totais do painel contariam o mesmo custo duas vezes")

    def test_bloco_do_departamento_avisa_o_que_ja_mostrou(self):
        """Cada bloco personalizado devolve o que ja desenhou, e o bloco
        generico pula isso. Sem essa combinacao, a Gerencia Comercial exibia
        as mesmas duas tabelas duas vezes na mesma tela."""
        ns = carregar(
            ["_normalizar_texto", "_normalizar_nome_aba", "get_valor_consolidado_multi",
             "_nome_sem_numero_dre", "_resolver_termo_departamento", "_numero_linha_dre",
             "_linha_pertence_ao_grupo", "_linhas_raiz_do_conjunto", "cor_valor",
             "cor_variacao", "formata_brl", "formata_valor_curto", "faixa_metricas_html",
             "html_compacto", "_total_dre", "_tabela_departamento", "_painel_dept_comercial"],
            ["COLORS", "FONTE_MONO", "LINHA_RECEITA_LIQUIDA", "LINHA_DESPESAS_VARIAVEIS",
             "MODELOS_RELATORIO"],
            extras={"st": type("st", (), {
                "markdown": staticmethod(lambda *a, **k: None),
                "caption": staticmethod(lambda *a, **k: None),
                "dataframe": staticmethod(lambda *a, **k: None),
            })()},
        )
        nome = "📈 Relatório de Custos - Gerência Comercial"
        valores = {"6 - Despesas Variáveis": -9_363_906.38,
                   "6.1 - Comissões sobre Vendas": -935_381.52,
                   "6.24.2.5 - Encontro de Ciclo": -4_519.40,
                   "6.24.2.6 - Outras Despesas de Marketing": -415_547.83,
                   "8.3.3.7 - Prêmios / Bônus": -301_779.55,
                   "3 - Receita Operacional Liquida": 75_000_000}
        aba = pd.DataFrame([{"Nome": n, "01/2026": v} for n, v in valores.items()])
        universo = list(valores.keys())
        modelo = ns["MODELOS_RELATORIO"][nome]
        geridas, informativas = [], []
        for termo in modelo["linhas_dre"]:
            geridas.extend(ns["_resolver_termo_departamento"](termo, universo))
        for termo in modelo.get("linhas_informativas", []):
            informativas.extend(ns["_resolver_termo_departamento"](termo, universo))
        cobertura = ns["_painel_dept_comercial"]({
            "departamento": nome, "dfs_real": [aba], "dfs_orc": [aba], "colunas": ["01/2026"],
            "linhas_resolvidas": geridas, "linhas_raiz": ns["_linhas_raiz_do_conjunto"](geridas),
            "linhas_todas": universo, "linhas_informativas": informativas,
            "path_orc": "o", "path_real": "r"})
        self.assertEqual(cobertura, {"linhas_dept", "informativas"})
        self.assertIn('if "linhas_dept" not in _ja_mostrado_dept:', FONTE)
        self.assertIn('"informativas" not in _ja_mostrado_dept', FONTE)

    def test_relatorio_agrupa_o_quadro_de_metas(self):
        """No Excel o Quadro de Metas tem de sair como um bloco proprio, com
        subtotal e as tres linhas recuadas abaixo -- foi assim que o
        relatorio sempre foi enviado. Sem isso, o Excel lista tudo numa
        sequencia so e o total do quadro nao aparece em lugar nenhum."""
        self.assertIn('"bloco_relatorio"', FONTE)
        self.assertIn("Despesa Variavel - Quadro de Metas GER.COM.", FONTE)
        i = FONTE.index("def montar_relatorio_excel(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("blocos_agrupados", trecho,
                      "o gerador do Excel voltou a ignorar os blocos do modelo")
        self.assertIn("contas_em_bloco", trecho)

    def test_relatorio_tem_opcao_de_puxar_a_dre_inteira(self):
        self.assertIn("Puxar todas as linhas da DRE", FONTE)
        self.assertIn("puxar_dre_completa", FONTE)

    def test_modelo_e_email_cadastrados(self):
        self.assertIn(self.NOME, FONTE)
        self.assertIn("gerente.comercial@grupobeea.com.br", FONTE)
        self.assertIn('"linhas_informativas"', FONTE)


# ============================================================================
# 6. FORMATACAO
# ============================================================================
class TesteFormatacao(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(["formata_brl", "formata_valor_curto", "formata_sigma_curto"])

    def test_reais_no_padrao_brasileiro(self):
        self.assertEqual(self.ns["formata_brl"](1234567.89), "R$ 1.234.567,89")

    def test_valor_curto_muda_de_escala(self):
        self.assertEqual(self.ns["formata_valor_curto"](1_200_000), "R$ 1,2M")
        self.assertEqual(self.ns["formata_valor_curto"](12_000), "R$ 12 mil")
        self.assertEqual(self.ns["formata_valor_curto"](480), "R$ 480")

    def test_sigma_tem_teto(self):
        self.assertEqual(self.ns["formata_sigma_curto"](25), "+10σ+")
        # o sinal negativo aqui e o MENOS tipografico (U+2212), nao o hifen
        self.assertEqual(self.ns["formata_sigma_curto"](-25), "\u221210\u03c3+")
        self.assertEqual(self.ns["formata_sigma_curto"](5.4), "+5,4σ")


# ============================================================================
# 7. TRAVAS ESTRUTURAIS - as regressoes que ja aconteceram
# ============================================================================
class TesteTravasEstruturais(unittest.TestCase):
    """Estes testes nao olham resultado: olham o codigo. Cada um trava uma
    decisao que ja foi desfeita sem querer e custou horas para achar."""

    def test_login_falha_fechado(self):
        # Olha SO o bloco do "if not usuarios": e ali que estava o return True
        # que abria o painel quando a leitura dos Secrets falhava.
        i = FONTE.index("def checar_login(")
        trecho = FONTE[i:i + 6000]
        inicio = trecho.index("if not usuarios:")
        # o bloco termina na proxima linha com 4 espacos de recuo
        resto = trecho[inicio:]
        fim = len(resto)
        for linha_inicio in range(1, len(resto)):
            if resto[linha_inicio - 1] == "\n" and resto[linha_inicio:linha_inicio + 5] == "    i":
                fim = linha_inicio
                break
        bloco = resto[:fim]
        self.assertNotIn("return True", bloco,
                         "checar_login voltou a liberar acesso quando nao ha usuarios")
        self.assertIn("return False", bloco,
                      "o bloco de 'sem usuarios' precisa barrar o acesso")

    def test_colunas_do_csv_usam_a_ordem_detectada(self):
        i = FONTE.index("df[COL_FIN_DATA_LIQUIDACAO] = pd.to_datetime(")
        self.assertIn("dayfirst=dayfirst_liq", FONTE[i:i + 300])
        j = FONTE.index("df[COL_FIN_VENCIMENTO] = pd.to_datetime(")
        self.assertIn("dayfirst=dayfirst_venc", FONTE[j:j + 300])

    def test_baixa_da_diario_nao_entra_na_coluna_do_csv(self):
        self.assertNotIn("df[COL_FIN_DATA_LIQUIDACAO] = df[COL_FIN_LIQ_DIARIO]", FONTE)

    def test_caches_de_dados_tem_validade(self):
        for funcao in ("obter_caminhos_excel", "obter_dados_fluxo_caixa",
                       "preparar_fluxo_caixa", "carregar_liquidacoes_diario"):
            i = FONTE.index(f"def {funcao}(")
            decorador = FONTE[max(0, i - 260):i]
            self.assertIn("ttl=", decorador, f"{funcao} ficou sem ttl no cache")

    def test_fluxo_tenta_mais_de_uma_fonte(self):
        i = FONTE.index("def fontes_csv_fluxo(")
        trecho = FONTE[i:i + 2000]
        self.assertIn("pub?output=csv", trecho)
        self.assertIn("FLUXO_CAIXA_FILE_ID", trecho)
        # A validacao evita aceitar em silencio uma fonte que devolveu
        # outra aba - o pior desfecho possivel: numero errado sem aviso.
        i_leitura = FONTE.index("def obter_dados_fluxo_caixa(")
        leitura = FONTE[i_leitura:i_leitura + 2200]
        self.assertIn("COLUNAS_MINIMAS_FLUXO", leitura,
                      "a leitura do fluxo parou de conferir as colunas da fonte")
        self.assertIn("veio outra aba", leitura)

    def test_chamadas_depreciadas_do_streamlit_nao_voltam(self):
        self.assertNotIn("use_container_width", FONTE,
                         "use_container_width foi removido do Streamlit em 31/12/2025")
        self.assertLessEqual(
            FONTE.count("components.html("), 2,
            "components.html so pode aparecer dentro de html_embutido")

    def test_orcado_por_canal_desce_para_as_linhas_filhas(self):
        """Em VD CONSOLIDADO o orcamento de marketing esta lancado nas linhas
        filhas, nao na 6.24.2. Ler so a mae devolvia zero e o canal aparecia
        sem orcamento - com o desvio inteiro contado como estouro."""
        i = FONTE.index("def _painel_dept_mkt(")
        trecho = FONTE[i:i + 3500]
        self.assertIn("_somar_codigo_com_filhas(df_o, CODIGO_MKT_GESTAO_CP", trecho,
                      "o orcado por canal voltou a ler so a linha-mae")
        self.assertIn("_somar_codigo_com_filhas(df_r, CODIGO_MKT_GESTAO_CP", trecho)

    def test_tabela_por_canal_fecha_com_o_cabecalho(self):
        """Somar a coluna da tabela por canal dava um numero diferente do
        cartao 'Desvio vs. Orcado' do topo - a diferenca era o que sai do 1%,
        descontado do investido da Loja mas presente no total do departamento.
        A linha 'Fora do 1%' e a linha TOTAL fazem os dois blocos conciliarem."""
        i = FONTE.index("def _painel_dept_mkt(")
        trecho = FONTE[i:i + 5000]
        self.assertIn('"Canal": "Fora do 1% (Loja)"', trecho,
                      "sumiu a linha que concilia o que sai do 1%")
        self.assertIn('"Canal": "TOTAL"', trecho, "sumiu a linha de total da tabela por canal")

    def test_nomes_de_parametro_nao_vazam_entre_funcoes(self):
        """Uma troca de nome em massa ja renomeou height=/width= para
        altura=/largura= em 29 chamadas de grafico e tabela, que esperam os
        nomes em ingles - o painel so quebrava ao abrir a aba afetada."""
        arvore = ast.parse(FONTE)

        def nome(no):
            f = no.func
            return f.id if isinstance(f, ast.Name) else (f.attr if isinstance(f, ast.Attribute) else "")

        errados = []
        for no in ast.walk(arvore):
            if not isinstance(no, ast.Call):
                continue
            args = [kw.arg for kw in no.keywords]
            if nome(no) == "html_embutido":
                if "height" in args or "width" in args:
                    errados.append(f"linha {no.lineno}: html_embutido com height/width")
            elif "altura" in args or "largura" in args:
                errados.append(f"linha {no.lineno}: {nome(no)}(...) com altura/largura")
        self.assertEqual(errados, [], "argumentos com o nome trocado: " + "; ".join(errados[:5]))

    def test_regra_do_1pct_usa_o_orcado_do_canal_loja(self):
        i = FONTE.index("def _painel_dept_mkt(")
        # a janela cobre a funcao inteira: ela ja cresceu duas vezes
        fim = FONTE.index("\ndef ", i + 10)
        trecho = FONTE[i:fim]
        self.assertIn('teto_loja = loja["Orçado (R$)"]', trecho,
                      "o teto do 1% voltou a ser calculado sobre a receita")


if __name__ == "__main__":
    print(f"Testando: {CAMINHO_APP}\n" + "=" * 62)
    unittest.main(argv=[sys.argv[0]], verbosity=2, exit=True)