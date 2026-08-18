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


# Dependencias internas: quem pede a funcao da esquerda precisa das da
# direita para rodar, mesmo sem saber disso. Sem esta tabela, cada teste teria
# de conhecer as tripas do app -- e quebrava sozinho quando o app mudava por
# dentro sem mudar de comportamento.
DEPENDENCIAS = {
    "_linhas_raiz_do_conjunto": ["_eh_linha_de_resultado", "_normalizar_texto"],
    "_eh_linha_de_resultado": ["_normalizar_texto"],
    "resolver_planos_forcados": ["_planos_sem_linha_dre", "_normalizar_texto"],
    "_planos_sem_linha_dre": ["_normalizar_texto"],
}
CONSTANTES_DE_DEPENDENCIA = {
    "_eh_linha_de_resultado": ["PALAVRAS_LINHA_DE_RESULTADO"],
    "_planos_sem_linha_dre": ["MARCAS_SEM_LINHA_DRE"],
    "resolver_planos_forcados": ["MODELOS_RELATORIO"],
}


def carregar(nomes_funcoes, nomes_constantes=(), extras=None):
    """Executa apenas as funcoes e constantes pedidas, num espaco isolado."""
    nomes_funcoes = list(nomes_funcoes)
    nomes_constantes = list(nomes_constantes)
    # Fila, e nao um laco unico: a dependencia de uma dependencia tambem
    # precisa entrar. Sem isso, o detector entrava mas a constante dele nao.
    fila = list(nomes_funcoes)
    while fila:
        nome = fila.pop()
        for extra in DEPENDENCIAS.get(nome, []):
            if extra not in nomes_funcoes:
                nomes_funcoes.append(extra)
                fila.append(extra)
        for extra in CONSTANTES_DE_DEPENDENCIA.get(nome, []):
            if extra not in nomes_constantes:
                nomes_constantes.append(extra)
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
# 5d. ADM/FINANCEIRO
# ============================================================================
class TesteAdmFinanceiro(unittest.TestCase):
    """O ADM/Financeiro fica com o que sobra da DRE. O recorte e calculado,
    nao listado a mao -- entao precisa de trava, senao um departamento novo
    passa a levar linhas embora sem ninguem perceber."""

    DRE = [
        "1 - Receita Operacional Bruta", "2 - Deduções da Receita Operacional Bruta",
        "3 - Receita Operacional Liquida", "4 - Custo das Vendas",
        "4.1 - Custo da Mercadoria Vendida - CMV", "5 - Margem de Contribuição 1",
        "6 - Despesas Variáveis", "6.1 - Comissões sobre Vendas",
        "6.2 - Taxa com Cartão de Crédito / Débito", "6.5 - Taxa de Emissão de Boleto",
        "6.6 - Material de Embalagem", "6.8 - Serviço de Entrega",
        "6.11 - Catálogos e Revistas", "6.13 - Amostras", "6.14 - Flaconetes",
        "6.16 - Vitrines", "6.24 - Esforços de Marketing",
        "6.24.1 - Marketing Regional - Gestão GB", "6.24.2 - Marketing Regional - Gestão CP",
        "6.24.2.5 - Encontro de Ciclo", "7 - Margem de Contribuição 2",
        "8 - Despesas Operacionais", "8.1 - Ocupação", "8.1.1 - Aluguel",
        "8.1.2 - Energia Elétrica", "8.1.3 - Limpeza e Conservação",
        "8.1.4 - Manutenção e Reparos", "8.2 - Tributos e Taxas", "8.3 - Pessoal",
        "8.3.1 - Salários", "8.3.3.7 - Prêmios / Bônus", "8.5.3 - Combustível",
        "8.6.1 - Material de Escritório", "8.6.6 - Outras Despesas Administrativas",
        "8.6.7 - Tarifas Bancárias", "8.8.1 - Contabilidade",
        "8.8.2 - Auditoria / Consultoria", "11 - EBITDA",
        "12 - Resultado Financeiro", "12.1 - Despesas Financeiras",
        "13 - Depreciação e Amortização",
        "14 - Outras Receitas e Despesas não Operacionais",
        "14.3 - Outras Receitas não Operacionais",
        "15 - Resultado Antes do Imposto", "16 - Impostos sobre o Lucro", "16.1 - IRPJ",
        "17 - Resultado Gerencial do Período",
    ]

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(
            ["_normalizar_texto", "_numero_linha_dre", "_linha_pertence_ao_grupo",
             "eh_linha_custos_despesas", "_resolver_termo_departamento",
             "_linhas_com_dono_de_departamento", "_linhas_restantes_da_dre",
             "_linhas_raiz_do_conjunto",
             "resolver_planos_forcados"],
            ["MODELOS_RELATORIO", "PALAVRAS_LINHA_DE_RESULTADO"],
        )

    def _geridas(self):
        return self.ns["_resolver_termo_departamento"]("RESTANTE", self.DRE)

    def test_nao_leva_linha_de_outro_departamento(self):
        geridas = self._geridas()
        for alheia in ["6.1 - Comissões sobre Vendas", "6.6 - Material de Embalagem",
                       "6.8 - Serviço de Entrega", "8.1.3 - Limpeza e Conservação",
                       "8.3 - Pessoal", "8.3.1 - Salários", "8.5.3 - Combustível",
                       "8.6.1 - Material de Escritório", "8.8.2 - Auditoria / Consultoria",
                       "6.24.2 - Marketing Regional - Gestão CP",
                       "6.24.1 - Marketing Regional - Gestão GB"]:
            self.assertNotIn(alheia, geridas, f"{alheia} ja tem dono")

    def test_nao_leva_linha_ancestral(self):
        """A linha de grupo carrega o custo dos outros departamentos dentro:
        se ela entrasse, a folha do RH viraria despesa do ADM."""
        geridas = self._geridas()
        for ancestral in ["6 - Despesas Variáveis", "8 - Despesas Operacionais",
                          "8.1 - Ocupação", "6.24 - Esforços de Marketing"]:
            self.assertNotIn(ancestral, geridas, f"{ancestral} soma custo de outra area")

    def test_fica_com_o_que_sobra(self):
        geridas = self._geridas()
        for propria in ["4 - Custo das Vendas", "6.2 - Taxa com Cartão de Crédito / Débito",
                        "6.5 - Taxa de Emissão de Boleto", "6.16 - Vitrines",
                        "8.1.1 - Aluguel", "8.1.2 - Energia Elétrica",
                        "8.2 - Tributos e Taxas", "8.6.7 - Tarifas Bancárias",
                        "8.8.1 - Contabilidade"]:
            self.assertIn(propria, geridas, f"{propria} nao tem dono e sumiu do ADM")

    def test_departamento_leva_a_dre_inteira_que_sobra(self):
        """Nao operacionais, impostos sobre o lucro e ate as linhas de
        resultado fazem parte do departamento -- em tela e no relatorio."""
        geridas = self._geridas()
        for linha in ["13 - Depreciação e Amortização", "12.1 - Despesas Financeiras",
                      "14 - Outras Receitas e Despesas não Operacionais",
                      "14.3 - Outras Receitas não Operacionais",
                      "16 - Impostos sobre o Lucro", "16.1 - IRPJ",
                      "15 - Resultado Antes do Imposto",
                      "17 - Resultado Gerencial do Período",
                      "1 - Receita Operacional Bruta", "11 - EBITDA"]:
            self.assertIn(linha, geridas, f"{linha} sumiu do ADM/Financeiro")

    def test_linha_calculada_nunca_entra_numa_soma(self):
        """15 e 17 sao a DRE inteira fechada: soma-las junto com as contas
        que as formam multiplicaria o total do departamento."""
        calculadas = ["1 - Receita Operacional Bruta",
                      "2 - Deduções da Receita Operacional Bruta",
                      "3 - Receita Operacional Liquida", "5 - Margem de Contribuição 1",
                      "11 - EBITDA", "12 - Resultado Financeiro",
                      "15 - Resultado Antes do Imposto",
                      "17 - Resultado Gerencial do Período"]
        for linha in calculadas:
            self.assertTrue(self.ns["_eh_linha_de_resultado"](linha), linha)
        raizes = self.ns["_linhas_raiz_do_conjunto"](self._geridas())
        for linha in calculadas:
            self.assertNotIn(linha, raizes, f"{linha} entrou no total")

    def test_imposto_sobre_o_lucro_e_conta_de_verdade(self):
        """A palavra "lucro" ja esteve na lista de linhas calculadas e
        derrubava os impostos do total do departamento."""
        self.assertFalse(self.ns["_eh_linha_de_resultado"]("16 - Impostos sobre o Lucro"))
        raizes = self.ns["_linhas_raiz_do_conjunto"](self._geridas())
        self.assertIn("16 - Impostos sobre o Lucro", raizes)
        self.assertNotIn("16.1 - IRPJ", raizes, "a filha nao pode somar junto com a mae")

    def test_gerencia_comercial_nao_tira_linha_do_adm(self):
        """As tres linhas da Gerencia Comercial sao acompanhamento sobre ramos
        de outras areas -- desconta-las abriria um buraco na DRE."""
        donas = self.ns["_linhas_com_dono_de_departamento"](self.DRE)
        self.assertNotIn("6.16 - Vitrines", donas)
        modelo = self.ns["MODELOS_RELATORIO"]["📈 Relatório de Custos - Gerência Comercial"]
        self.assertFalse(modelo.get("linhas_exclusivas", False))

    def test_planos_que_sobram(self):
        diario = pd.DataFrame([
            {"Plano de Contas": "Mercadorias", "Linha DRE": ""},
            {"Plano de Contas": "Benfeitorias em imóvel próprio", "Linha DRE": ""},
            {"Plano de Contas": "Aplicação Financeira", "Linha DRE": ""},
            {"Plano de Contas": "Energia Elétrica", "Linha DRE": "8.1.2 - Energia Elétrica"},
        ])
        achados, _ = self.ns["resolver_planos_forcados"](diario, ["RESTANTE"])
        self.assertIn("Aplicação Financeira", achados)
        self.assertNotIn("Mercadorias", achados, "Mercadorias e de Compras")
        self.assertNotIn("Benfeitorias em imóvel próprio", achados, "benfeitorias sao de Suprimentos")
        self.assertNotIn("Energia Elétrica", achados, "esse ja tem linha da DRE")

    def test_mapa_de_email_aponta_para_departamento_existente(self):
        """Um e-mail apontando para um departamento com nome errado nao da
        erro: a pessoa simplesmente abre na Controladoria e ninguem descobre
        o porque. Por isso a conferencia e automatica."""
        trecho = FONTE[FONTE.index("MAPA_EMAIL_DEPARTAMENTO = {"):]
        trecho = trecho[:trecho.index("\n}")]
        mapeados = re.findall(r'"([^"]+@[^"]+)":\s*"([^"]+)"', trecho)
        self.assertGreaterEqual(len(mapeados), 8)
        for email, departamento in mapeados:
            self.assertIn(departamento, self.ns["MODELOS_RELATORIO"],
                          f"{email} aponta para um departamento que nao existe")
        for email, departamento in [
            ("coordenador.financeiro@grupobeea.com.br", "🏦 Relatório de Custos - ADM/Financeiro"),
            ("coordenador.loja@grupobeea.com.br", "🏬 Relatório de Custos - Coordenação de Loja"),
            ("coordenador.vd@grupobeea.com.br", "🚗 Relatório de Custos - Coordenação de VD"),
        ]:
            self.assertIn((email, departamento), mapeados)

    def test_celula_vazia_de_linha_dre_e_reconhecida(self):
        """A leitura converte a coluna inteira para texto, entao o NaN do
        pandas vira a palavra "nan". Testar so por string vazia deixava
        quase todos os planos fora da DRE passarem batido."""
        ns = carregar(["_normalizar_texto", "_planos_sem_linha_dre"],
                      ["MARCAS_SEM_LINHA_DRE"])
        diario = pd.DataFrame({
            "Plano de Contas": ["Aplicação Financeira", "Empréstimos", "Energia Elétrica",
                                "Distribuição de Lucros", "Transferência"],
            "Linha DRE": [float("nan"), "", "8.1.2 - Energia Elétrica", "-", "nan"],
        })
        diario["Linha DRE"] = diario["Linha DRE"].astype(str).str.strip()
        achados = ns["_planos_sem_linha_dre"](diario)
        self.assertIn("Aplicação Financeira", achados)
        self.assertIn("Distribuição de Lucros", achados)
        self.assertIn("Transferência", achados)
        self.assertNotIn("Energia Elétrica", achados, "esse tem linha da DRE")

    def test_planos_de_outro_departamento_ficam_de_fora(self):
        ns = carregar(["_normalizar_texto", "_planos_sem_linha_dre",
                       "resolver_planos_forcados"],
                      ["MARCAS_SEM_LINHA_DRE", "MODELOS_RELATORIO"])
        diario = pd.DataFrame({
            "Plano de Contas": ["Aplicação Financeira", "Mercadorias",
                                "Adiantamento de Benfeitorias em Imóvel Próprio"],
            "Linha DRE": ["", "", ""],
        })
        achados, _ = ns["resolver_planos_forcados"](diario, ["RESTANTE"])
        self.assertIn("Aplicação Financeira", achados)
        self.assertNotIn("Mercadorias", achados, "Mercadorias e de Compras")
        self.assertNotIn("Adiantamento de Benfeitorias em Imóvel Próprio", achados,
                         "benfeitorias sao de Suprimentos")

    def test_bloco_de_planos_aparece_no_painel(self):
        self.assertIn("Planos de Contas Fora da DRE", FONTE)
        i = FONTE.index("def _painel_dept_adm(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("_planos_sem_linha_dre", trecho)
        self.assertNotIn('"% do bloco"', trecho,
                         "percentual sobre total com sinais opostos nao significa nada")

    def test_todos_passam_pela_tela_de_escolha(self):
        """A tela de escolha entre Controladoria e Financeiro aparece para
        TODO usuario logado. O que muda por pessoa e so o acesso ao
        Financeiro: quem nao esta na lista ve o cartao bloqueado, com o
        botao desligado. Nao pode existir atalho que mande alguem direto
        para um painel sem passar por aqui."""
        i = FONTE.index('if st.session_state["painel_escolhido"] is None:')
        trecho = FONTE[i:i + 12000]
        # A unica condicao que governa a tela e "ainda nao escolheu" -- nao
        # pode haver desvio por e-mail antes dela.
        antes = FONTE[:i]
        self.assertNotIn('st.session_state["painel_escolhido"] = "controladoria"', antes,
                         "alguem esta escolhendo o painel antes da tela aparecer")
        self.assertIn("_pode_financeiro = _email_hub in EMAILS_FINANCEIRO_PERMITIDOS", trecho)
        self.assertIn("disabled=not _pode_financeiro", trecho,
                      "o botao do Financeiro precisa nascer desligado para quem nao tem acesso")
        self.assertIn("st.stop()", trecho, "a tela precisa parar o resto do app")

    def test_acesso_ao_financeiro_e_por_lista(self):
        i = FONTE.index("EMAILS_FINANCEIRO_PERMITIDOS = {")
        trecho = FONTE[i:FONTE.index("}", i)]
        emails = re.findall(r'"([^"]+@[^"]+)"', trecho)
        self.assertIn("controladoria@grupobeea.com.br", emails)
        for de_fora in ["coordenador.loja@grupobeea.com.br",
                        "coordenador.vd@grupobeea.com.br",
                        "gerente.comercial@grupobeea.com.br"]:
            self.assertNotIn(de_fora, emails,
                             f"{de_fora} nao foi autorizado para o Painel Financeiro")

    def test_modelo_e_painel_cadastrados(self):
        self.assertIn("🏦 Relatório de Custos - ADM/Financeiro", self.ns["MODELOS_RELATORIO"])
        self.assertIn("_painel_dept_adm", FONTE)
        self.assertIn('"🏦 Relatório de Custos - ADM/Financeiro": _painel_dept_adm', FONTE)


# ============================================================================
# 5d-bis. COORDENACOES DE LOJA E DE VD
# ============================================================================
class TesteCoordenacoes(unittest.TestCase):
    """As duas coordenacoes leem da receita ao EBITDA e respondem so pelas
    abas delas."""

    DRE = ["1 - Receita Operacional Bruta", "2 - Deduções da Receita Operacional Bruta",
           "3 - Receita Operacional Liquida", "4 - Custo das Vendas",
           "5 - Margem de Contribuição 1", "6 - Despesas Variáveis",
           "6.1 - Comissões sobre Vendas", "7 - Margem de Contribuição 2",
           "8 - Despesas Operacionais", "9 - Resultado Operacional", "11 - EBITDA",
           "12 - Resultado Financeiro", "13 - Depreciação e Amortização",
           "15 - Resultado Antes do Imposto", "16 - Impostos sobre o Lucro",
           "17 - Resultado Gerencial do Período"]
    LOJA = "🏬 Relatório de Custos - Coordenação de Loja"
    VD = "🚗 Relatório de Custos - Coordenação de VD"

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(
            ["_normalizar_texto", "_numero_linha_dre", "_linha_pertence_ao_grupo",
             "_linhas_ate_ebitda", "_resolver_termo_departamento"],
            ["MODELOS_RELATORIO", "GRUPO_EBITDA_DRE"],
        )

    def test_para_no_ebitda(self):
        """Abaixo do 11 fica o que a operacao da loja nao decide: resultado
        financeiro, depreciacao, imposto sobre o lucro."""
        achadas = self.ns["_resolver_termo_departamento"]("ATE_EBITDA", self.DRE)
        numeros = [self.ns["_numero_linha_dre"](l) for l in achadas]
        self.assertIn("11", numeros)
        self.assertIn("6.1", numeros, "sublinha do grupo 6 tem de entrar")
        for abaixo in ["12", "13", "15", "16", "17"]:
            self.assertNotIn(abaixo, numeros, f"a linha {abaixo} vem depois do EBITDA")

    def test_cada_coordenacao_so_ve_as_abas_dela(self):
        loja = self.ns["MODELOS_RELATORIO"][self.LOJA]
        vd = self.ns["MODELOS_RELATORIO"][self.VD]
        self.assertEqual(len(loja["unidades_permitidas"]), 13)
        self.assertEqual(len(vd["unidades_permitidas"]), 7)  # 5 VD + 2 ABPR
        self.assertNotIn("DRE CONSOLIDADO", loja["visoes_permitidas"])
        self.assertNotIn("DRE CONSOLIDADO", vd["visoes_permitidas"])
        # Nenhuma unidade pode aparecer nas duas coordenacoes.
        self.assertFalse(
            set(loja["unidades_permitidas"]) & set(vd["unidades_permitidas"]),
            "a mesma loja apareceria nas duas coordenacoes",
        )

    def test_abas_batem_com_os_grupos_do_relatorio(self):
        """As listas do modelo tem de ser as MESMAS que o gerador do Excel usa
        para montar as visoes consolidadas -- se uma loja nova entrar so num
        lugar, o consolidado e o ranking param de fechar."""
        i = FONTE.index("def _lojas_do_grupo_consolidado(")
        trecho = FONTE[i:i + 2000]
        for loja in self.ns["MODELOS_RELATORIO"][self.LOJA]["unidades_permitidas"]:
            self.assertIn(loja, trecho, f"{loja} nao esta no grupo do relatorio")
        for unidade in self.ns["MODELOS_RELATORIO"][self.VD]["unidades_permitidas"]:
            self.assertIn(unidade, trecho, f"{unidade} nao esta no grupo do relatorio")

    def test_linha_de_percentual_do_ebitda(self):
        for nome in (self.LOJA, self.VD):
            self.assertTrue(self.ns["MODELOS_RELATORIO"][nome].get("linha_percentual_ebitda"))
        i = FONTE.index("if percentual_ebitda is not None:")
        trecho = FONTE[i:i + 900]
        self.assertIn("11 - EBITDA (%)", trecho)
        self.assertIn("pontos percentuais", FONTE,
                      "a legenda precisa avisar que o desvio da linha e em p.p.")

    def test_barra_lateral_filtra_as_abas(self):
        self.assertIn("_filtrar_abas_permitidas", FONTE)
        self.assertIn('_modelo_dept_ativo.get("visoes_permitidas")', FONTE)
        self.assertIn('_modelo_dept_ativo.get("unidades_permitidas")', FONTE)


# ============================================================================
# 5e. PRAZOS DO FLUXO
# ============================================================================
class TestePrazosDoFluxo(unittest.TestCase):
    """A baixa dos titulos nao vem no CSV do fluxo; ela e buscada na aba
    DIARIO pelo numero do documento. O cruzamento precisa cobrir os DOIS
    lados -- quando olhava so o "pagar", o atraso medio no recebimento
    aparecia vazio na tela."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(["_normalizar_coluna_fin", "_chave_numero_fin",
                           "_classificar_movimento_fin"])

    def _cruzar(self):
        df = pd.DataFrame({
            "Movimento": ["4 - Contas a Pagar", "2 - Contas a Receber", "1 - Banco"],
            "Número": ["1001", "5001", ""],
            "Data Liquidação": [pd.NaT, pd.NaT, pd.NaT],
            "Vencimento.1": pd.to_datetime(["2026-08-05", "2026-08-03", None]),
        })
        diario = pd.DataFrame({
            "chave_num": ["1001", "5001"],
            "data_liq": pd.to_datetime(["2026-08-07", "2026-08-06"]),
        })
        pagar = df["Movimento"].astype(str).str.contains("pagar", case=False, na=False)
        receber = df["Movimento"].astype(str).str.contains("receber", case=False, na=False)
        faltando = (pagar | receber) & df["Data Liquidação"].isna()
        a_cruzar = pd.DataFrame({
            "_idx": df.index[faltando],
            "chave_num": self.ns["_chave_numero_fin"](df.loc[faltando, "Número"]).values,
        })
        casados = a_cruzar.merge(diario, on="chave_num", how="left").dropna(subset=["data_liq"])
        df["Liquidação Efetiva"] = pd.NaT
        df.loc[casados["_idx"].values, "Liquidação Efetiva"] = casados["data_liq"].values
        df["Tipo Movimento"] = df["Movimento"].map(self.ns["_classificar_movimento_fin"])
        return df

    def test_o_receber_tambem_ganha_data_de_baixa(self):
        df = self._cruzar()
        entradas = df[df["Tipo Movimento"] == "entrada"]
        self.assertTrue(entradas["Liquidação Efetiva"].notna().all(),
                        "titulo a receber ficou sem data de baixa")

    def test_atraso_sai_dos_dois_lados(self):
        df = self._cruzar()
        prazo = df[df["Liquidação Efetiva"].notna() & df["Vencimento.1"].notna()].copy()
        prazo["DiasAteLiquidar"] = (prazo["Liquidação Efetiva"] - prazo["Vencimento.1"]).dt.days
        for tipo in ("saida", "entrada"):
            lado = prazo[prazo["Tipo Movimento"] == tipo]
            self.assertFalse(lado.empty, f"nenhum titulo de {tipo} com prazo calculado")
            self.assertEqual(lado["DiasAteLiquidar"].mean(), 2 if tipo == "saida" else 3)

    def test_cruzamento_olha_pagar_e_receber(self):
        """Trava estrutural: se o filtro voltar a ser so o "pagar", o
        recebimento fica sem data e o indicador zera de novo."""
        i = FONTE.index("mask_receber = df[COL_FIN_MOVIMENTO]")
        trecho = FONTE[i:i + 400]
        self.assertIn("(mask_pagar | mask_receber)", trecho)

    def test_vencimento_e_a_chave_mais_forte(self):
        """As duas abas trazem o MESMO vencimento -- a diferenca entre elas e
        so a liquidacao. Cruzar por numero + vencimento + valor evita casar
        dois titulos do mesmo fornecedor com numeros parecidos."""
        i = FONTE.index("tentativas = [")
        trecho = FONTE[i:i + 700]
        ordem = [n for n in ["por_num_venc_valor", "por_num_venc", "por_num_valor",
                             "por_venc_valor", "por_num"] if n in trecho]
        self.assertEqual(ordem[0], "por_num_venc_valor", "a chave mais especifica saiu da frente")
        self.assertEqual(ordem[-1], "por_num", "so o numero tem de ser a ultima tentativa")

    def test_par_ambiguo_nao_recebe_data(self):
        """Dois titulos que vencem no mesmo dia pelo mesmo valor, pagos em
        datas diferentes, nao tem como ser distinguidos: ficam em aberto em
        vez de receber um chute que estragaria o atraso medio."""
        diario = pd.DataFrame({
            "chave_num": ["1005", "1006", "1007"],
            "venc": pd.to_datetime(["2026-08-20", "2026-08-20", "2026-08-21"]),
            "valor_abs": ["950.00", "950.00", "300.00"],
            "data_liq": pd.to_datetime(["2026-08-22", "2026-08-25", "2026-08-21"]),
        })
        contagem = diario.groupby(["venc", "valor_abs"])["data_liq"].nunique()
        unicos = contagem[contagem == 1].index
        candidatos = diario.set_index(["venc", "valor_abs"])
        sobrando = candidatos.loc[candidatos.index.isin(unicos)]
        self.assertEqual(len(sobrando), 1)
        self.assertEqual(sobrando["chave_num"].iloc[0], "1007")

    def test_titulo_sem_baixa_na_diario_fica_em_aberto(self):
        """Quem nao foi pago esta sem data na DIARIO tambem -- e tem de
        continuar aparecendo como em aberto, nao virar pago sem data."""
        base = pd.DataFrame({
            "chave_num": ["1001", "1003"],
            "data_liq": [pd.Timestamp("2026-08-07"), pd.NaT],
        })
        validos = base[base["data_liq"].notna()]
        self.assertEqual(list(validos["chave_num"]), ["1001"])

    def test_indicador_nao_se_chama_prazo_medio(self):
        """O calculo e liquidacao menos vencimento, ou seja ATRASO. Chamar de
        prazo medio faz ler "+5 dias" como se a empresa pagasse em 5 dias."""
        self.assertIn("ATRASO MÉDIO NO PAGAMENTO", FONTE)
        self.assertIn("ATRASO MÉDIO NO RECEBIMENTO", FONTE)
        self.assertNotIn('label="PRAZO MÉDIO DE PAGAMENTO"', FONTE)
        self.assertNotIn('label="PRAZO MÉDIO DE RECEBIMENTO"', FONTE)


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