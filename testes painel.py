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
import json
import hashlib
from datetime import date, datetime
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
    "_assinatura_coluna_fin": ["_normalizar_coluna_fin"],
    "resolver_colunas_fluxo": ["_assinatura_coluna_fin", "_normalizar_coluna_fin"],
    "guardar_memoria": ["memoria_em_uso_mb"],
    "meta_diaria_que_ainda_falta": [],
    "_tabela_departamento": ["_cor_valor_invertido", "cor_valor", "formata_brl"],
    "_cor_valor_invertido": ["cor_valor"],
}
CONSTANTES_DE_DEPENDENCIA_CONST = {
    # Constante que depende de outra constante para ser avaliada.
    "VISOES_CONSOLIDADAS": ["GRUPO_ABPR", "GRUPO_VD", "GRUPO_LJ_GA", "GRUPO_LJ_CONSOLIDADO"],
    "RENOMEAR_MOVIMENTO_FIN": ["MOV_RECEBER_META", "MOV_RECEBER_AVENCER",
                               "MOV_RECEBER_LIQUIDADO", "MOV_PAGAR"],
}
CONSTANTES_DE_DEPENDENCIA = {
    "_eh_linha_de_resultado": ["PALAVRAS_LINHA_DE_RESULTADO"],
    "_planos_sem_linha_dre": ["MARCAS_SEM_LINHA_DRE"],
    "_assinatura_coluna_fin": ["LIGACOES_NOME_COLUNA", "_ACENTOS_FIN"],
    "_normalizar_coluna_fin": ["_ACENTOS_FIN"],
    "tabela_selecionavel": ["COLORS", "FONTE_MONO", "FONTE_PADRAO_TABELA",
                            "TETO_LINHAS_TABELA", "ALTURA_LINHA_TABELA_PX",
                            "ALTURA_CABECALHO_TABELA_PX", "ALTURA_BARRA_SOMA_PX",
                            "ALTURA_BARRA_ROLAGEM_PX", "COLUNAS_SEM_ROLAGEM_HORIZONTAL",
                            "FUNDO_TABELA_FLUXO", "PARAM_LINHAS_ABERTAS",
                            "SEPARADOR_LINHAS_ABERTAS", "PREFIXO_BOTAO_ABRIR"],
    "botoes_de_abrir": ["PREFIXO_BOTAO_ABRIR"],
    "guardar_memoria": ["LIMITE_MEMORIA_LIMPEZA_MB"],
    "linhas_abertas_da_url": ["PARAM_LINHAS_ABERTAS", "SEPARADOR_LINHAS_ABERTAS"],
    "resolver_planos_forcados": ["MODELOS_RELATORIO"],
}


def dubla_html_embutido(capturado):
    """Dublê de html_embutido com a assinatura COMPLETA da função de
    verdade. Cada teste tinha o seu, escrito com os argumentos do dia -- e
    quando a função ganhou `redimensionavel`, catorze testes quebraram de
    uma vez por um motivo que não era o que eles testavam."""

    def _capturar(codigo, altura=0, largura=None, redimensionavel=False):
        capturado.update(codigo=codigo, altura=altura, largura=largura,
                         redimensionavel=redimensionavel)

    return _capturar


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
    for nome in list(nomes_constantes):
        for extra in CONSTANTES_DE_DEPENDENCIA_CONST.get(nome, []):
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

    def test_abas_batem_com_as_visoes_consolidadas(self):
        """As listas do modelo tem de bater com VISOES_CONSOLIDADAS, que e a
        fonte unica usada pelo painel e pelo Excel -- se uma loja nova entrar
        so num lugar, o consolidado e o ranking param de fechar."""
        ns = carregar([], ["VISOES_CONSOLIDADAS"])
        visoes = ns["VISOES_CONSOLIDADAS"]
        loja = self.ns["MODELOS_RELATORIO"][self.LOJA]
        vd = self.ns["MODELOS_RELATORIO"][self.VD]
        self.assertEqual(sorted(loja["unidades_permitidas"]), sorted(visoes["LJ - G&A"]))
        self.assertEqual(sorted(vd["unidades_permitidas"]),
                         sorted(visoes["VD CONSOLIDADO"] + visoes["ABPR CONSOLIDADO"]))
        for nome_visao in loja["visoes_permitidas"] + vd["visoes_permitidas"]:
            self.assertIn(nome_visao, visoes, f"{nome_visao} nao existe no mapa de visoes")

    def test_diario_e_recortada_pela_visao(self):
        """Numero vindo da DIARIO (Mercadorias, benfeitorias, planos fora da
        DRE) tem de respeitar a visao escolhida -- senao aparece o valor da
        empresa inteira ao lado de linhas da DRE de uma loja so."""
        ns = carregar(["_normalizar_nome_aba", "_normalizar_texto",
                       "_lojas_individuais_das_abas", "_recortar_diario_por_loja"],
                      ["VISOES_CONSOLIDADAS"])
        self.assertEqual(ns["_lojas_individuais_das_abas"](["LJ PVH1 11927"]), ["LJ PVH1 11927"])
        self.assertEqual(len(ns["_lojas_individuais_das_abas"](["LJ - G&A"])), 13)
        self.assertEqual(ns["_lojas_individuais_das_abas"]([]), [])
        diario = pd.DataFrame({
            "Loja": ["LJ PVH1 11927", "LJ SETE 6052", "ABPR 23427"],
            "Plano de Contas": ["Mercadorias"] * 3,
            "Valor Bruto": [-100.0, -200.0, -300.0],
        })
        recorte = ns["_recortar_diario_por_loja"](diario, ["LJ PVH1 11927"])
        self.assertEqual(len(recorte), 1)
        self.assertEqual(recorte["Valor Bruto"].sum(), -100.0)
        # Nome que nao existe na DIARIO: devolve tudo, nunca zero.
        self.assertEqual(len(ns["_recortar_diario_por_loja"](diario, ["LOJA QUE NAO EXISTE"])), 3)

    def test_desvio_e_realizado_menos_orcado(self):
        """Um so significado para "Desvio" na tela inteira: realizado menos
        orcado, igual ao relatorio em Excel. Positivo = gastou mais."""
        self.assertIn("Realizado menos orçado · positivo = gastou mais", FONTE)
        self.assertNotIn('subtext="Positivo = gastou menos que o orçado"', FONTE)
        self.assertIn("def _cor_valor_invertido(", FONTE)
        # Nenhum bloco pode ter voltado para orcado menos realizado.
        for invertido in ['"Desvio (R$)": orc - real,', '"Desvio (R$)": orcado - valor,',
                          '"Desvio (R$)": v_o - v_r,']:
            self.assertNotIn(invertido, FONTE, f"convencao antiga de volta: {invertido}")

    def test_linha_de_percentual_do_ebitda(self):
        for nome in (self.LOJA, self.VD):
            self.assertTrue(self.ns["MODELOS_RELATORIO"][nome].get("linha_percentual_ebitda"))
        i = FONTE.index("if percentual_ebitda is not None:")
        trecho = FONTE[i:i + 900]
        self.assertIn("11 - EBITDA (%)", trecho)
        self.assertIn("pontos percentuais", FONTE,
                      "a legenda precisa avisar que o desvio da linha e em p.p.")

    def test_uma_aba_anual_por_visao_consolidada(self):
        """Loja tem 2 visoes consolidadas, VD tem 3 -- e cada uma vira uma aba
        de resumo anual no Excel."""
        loja = self.ns["MODELOS_RELATORIO"][self.LOJA]
        vd = self.ns["MODELOS_RELATORIO"][self.VD]
        self.assertTrue(loja.get("resumos_anuais_por_visao"))
        self.assertTrue(vd.get("resumos_anuais_por_visao"))
        self.assertEqual(len(loja["visoes_permitidas"]), 2)
        self.assertEqual(len(vd["visoes_permitidas"]), 3)
        self.assertIn("def _criar_aba_resumo_anual(", FONTE)
        i = FONTE.index("def _criar_aba_resumo_anual(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("ORÇADO", trecho)
        self.assertIn("REALIZADO", trecho)
        self.assertIn("EBITDA (%)", trecho)

    def test_percentual_anual_e_mes_a_mes(self):
        """O EBITDA % de cada mes sai do proprio mes, e o total sai do
        acumulado -- media dos meses daria outro numero quando o faturamento
        e desigual entre eles."""
        ebitda = [419_362.50, 382_194.63, 355_756.86]
        receita = [1_934_279.55, 1_708_908.50, 2_097_982.47]
        por_mes = [e / r * 100 for e, r in zip(ebitda, receita)]
        total = sum(ebitda) / sum(receita) * 100
        self.assertAlmostEqual(por_mes[0], 21.68, places=1)
        self.assertNotAlmostEqual(total, sum(por_mes) / 3, places=1)
        i = FONTE.index("def _criar_aba_resumo_anual(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("pares.append((sum(ebitda[indice]), sum(receita[indice])))", trecho)

    def test_relatorio_so_traz_plano_com_linha_da_dre(self):
        for nome in (self.LOJA, self.VD):
            self.assertTrue(
                self.ns["MODELOS_RELATORIO"][nome].get("apenas_planos_com_linha_dre"))
        i = FONTE.index("if apenas_planos_com_linha_dre")
        trecho = FONTE[i:i + 400]
        self.assertIn("MARCAS_SEM_LINHA_DRE", trecho,
                      "o filtro precisa usar o mesmo helper de celula vazia")

    def test_relatorio_so_oferece_as_abas_do_departamento(self):
        i = FONTE.index("opcoes_lojas_relatorio = _filtrar_abas_permitidas")
        trecho = FONTE[i - 600:i + 200]
        self.assertIn('_modelo_rel.get("visoes_permitidas"', trecho)
        self.assertIn('_modelo_rel.get("unidades_permitidas"', trecho)

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

    def test_as_duas_pontas_olham_a_mesma_coluna(self):
        """A pagar e a receber precisam ler a MESMA coluna de baixa; se uma
        delas olhar outra fonte, os dois lados param de conversar."""
        i = FONTE.index("mask_receber = df[COL_FIN_MOVIMENTO]")
        trecho = FONTE[i - 300:i + 600]
        self.assertIn("pagar_sem_liq_no_csv", trecho)
        self.assertIn("receber_sem_liq_no_csv", trecho)
        self.assertIn("COL_FIN_DATA_LIQUIDACAO", trecho)

    def test_baixa_vem_da_propria_aba_do_fluxo(self):
        """A aba Fluxo de Caixa 2026 passou a trazer a Data de Liquidacao
        preenchida, entao o confronto vencimento x liquidacao e feito ali
        mesmo. O cruzamento com a aba DIARIO -- que casava por numero,
        vencimento e valor -- foi removido e nao pode voltar sem decisao."""
        for resto in ("carregar_liquidacoes_diario", "COL_FIN_LIQ_DIARIO",
                      "liq_do_diario", "por_num_venc_valor", "pagar_sem_match"):
            self.assertNotIn(resto, FONTE, f"sobrou pedaco do cruzamento: {resto}")
        i = FONTE.index("df[COL_FIN_LIQ_EFETIVA] = df[COL_FIN_DATA_LIQUIDACAO]")
        trecho = FONTE[i:i + 300]
        self.assertIn("COL_FIN_LIQ_AMPLA", trecho,
                      "a leitura tolerante da MESMA coluna continua valendo")

    def test_titulo_sem_baixa_fica_no_vencimento(self):
        """Sem data de liquidacao o titulo esta em aberto: continua
        posicionado no vencimento, que e a programacao."""
        df = pd.DataFrame({
            "Movimento": ["4 - Contas a Pagar"] * 3,
            "Vencimento.1": pd.to_datetime(["2026-08-31"] * 3),
            "Data Liquidação": pd.to_datetime(["2026-08-24", None, None]),
            "Valor.1": [-300_000.0, -60_000.0, -40_000.0],
        })
        df["Liquidação Efetiva"] = df["Data Liquidação"]
        df["Data Efetiva"] = df["Vencimento.1"]
        com_baixa = df["Liquidação Efetiva"].notna()
        df.loc[com_baixa, "Data Efetiva"] = df.loc[com_baixa, "Liquidação Efetiva"]
        por_dia = df.groupby(df["Data Efetiva"].dt.date)["Valor.1"].sum()
        self.assertAlmostEqual(por_dia[date(2026, 8, 24)], -300_000.0, places=2)
        self.assertAlmostEqual(por_dia[date(2026, 8, 31)], -100_000.0, places=2)

    def test_indicador_nao_se_chama_prazo_medio(self):
        """O calculo e liquidacao menos vencimento, ou seja ATRASO. Chamar de
        prazo medio faz ler "+5 dias" como se a empresa pagasse em 5 dias."""
        self.assertIn("ATRASO MÉDIO NO PAGAMENTO", FONTE)
        self.assertIn("ATRASO MÉDIO NO RECEBIMENTO", FONTE)
        self.assertNotIn('label="PRAZO MÉDIO DE PAGAMENTO"', FONTE)
        self.assertNotIn('label="PRAZO MÉDIO DE RECEBIMENTO"', FONTE)


# ============================================================================
# 5f. REVISAO DOS PAINEIS (17/08/2026)
# ============================================================================
class TesteRevisaoDosPaineis(unittest.TestCase):
    """Travas do que a revisao de 17/08 corrigiu -- sao regras que somem
    facil na proxima mexida e que produzem numero errado sem dar erro."""

    def test_orcado_do_mes_corrente_entra_proporcional(self):
        ns = carregar(["_fator_proporcional_mes_corrente", "_escalar_orcado_mes_corrente"])
        orc = [pd.DataFrame({"Nome": ["11 - EBITDA"], "07/2026": [100_000.0],
                             "08/2026": [310_000.0]})]
        cols = ["07/2026", "08/2026"]
        saida, aviso = ns["_escalar_orcado_mes_corrente"](orc, cols, cols, date(2026, 8, 18))
        self.assertEqual(saida[0]["07/2026"].iloc[0], 100_000.0, "mes fechado nao pode encolher")
        self.assertAlmostEqual(saida[0]["08/2026"].iloc[0], 310_000 * 18 / 31, places=2)
        self.assertIn("18 de 31 dias", aviso)

    def test_periodo_fechado_nao_mexe_no_orcado(self):
        ns = carregar(["_fator_proporcional_mes_corrente", "_escalar_orcado_mes_corrente"])
        orc = [pd.DataFrame({"Nome": ["11 - EBITDA"], "07/2026": [100_000.0]})]
        saida, aviso = ns["_escalar_orcado_mes_corrente"](orc, ["07/2026"], ["07/2026"],
                                                          date(2026, 8, 18))
        self.assertIs(saida, orc)
        self.assertEqual(aviso, "")

    def test_mkt_acumula_o_fora_do_1pct(self):
        """Se um segundo canal passar a descontar, o valor do primeiro nao
        pode sumir -- o bug seria silencioso."""
        i = FONTE.index("def _painel_dept_mkt(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("fora_do_1pct += fora_do_canal", trecho)
        self.assertNotIn("fora_do_1pct = sum(", trecho)

    def test_reguas_do_rh_estao_declaradas(self):
        ns = carregar([], ["LIMITE_HORA_EXTRA_PCT_FOLHA", "LIMITE_RESCISOES_PCT_FOLHA"])
        self.assertEqual(ns["LIMITE_HORA_EXTRA_PCT_FOLHA"], 5.0)
        self.assertEqual(ns["LIMITE_RESCISOES_PCT_FOLHA"], 3.0)
        i = FONTE.index("def _painel_dept_rh(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertNotIn("* 100 - 5)", trecho, "limite voltou a ficar cravado no codigo")
        self.assertIn("régua", trecho, "a regua precisa aparecer na tela")

    def test_rh_nao_zera_o_orcado_do_resto(self):
        i = FONTE.index("def _painel_dept_rh(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("outros_orc = folha_orc_total", trecho)
        self.assertNotIn('"Orçado (R$)": 0.0,', trecho)

    def test_adm_ranqueia_por_valor_e_por_percentual(self):
        """So por valor, o custo da venda lidera sempre e enterra as
        distorcoes cronicas pequenas."""
        i = FONTE.index("def _painel_dept_adm(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("por_percentual", trecho)
        self.assertIn("piso_relevante", trecho)

    def test_bloco_que_carrega_planilha_usa_a_mesma_regua_de_orcado(self):
        """MKT (tabela por canal) e as coordenacoes (ranking por unidade)
        carregam planilha por conta propria. Se nao aplicarem o mesmo ajuste
        de mes corrente dos cartoes do topo, a tabela mostra o orcamento
        cheio e o cabecalho a parte decorrida -- foi o que apareceu na tela
        em 18/08/2026, com R$ 64 mil de diferenca."""
        self.assertIn('"escalar_orcado": lambda dfs: _escalar_orcado_mes_corrente(', FONTE)
        i = FONTE.index("def _painel_dept_mkt(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn('escalar = ctx.get("escalar_orcado")', trecho)
        self.assertIn("df_o = escalar([df_o])[0]", trecho)
        j = FONTE.index("def _painel_dept_coordenacao(")
        corpo = FONTE[j:FONTE.index("\ndef ", j + 10)]
        self.assertIn('dfs_orc_un = ctx["escalar_orcado"](dfs_orc_un)', corpo)

    def test_mkt_avisa_quando_a_tabela_nao_fecha(self):
        """A tabela por canal e os cartoes do topo saem por caminhos
        diferentes e por isso um valida o outro. Quando divergem, a tela tem
        de dizer -- melhor que descobrir na reuniao."""
        i = FONTE.index("def _painel_dept_mkt(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn('_total_linha = next((l for l in linhas_tabela if l["Canal"] == "TOTAL")', trecho)
        self.assertIn("não está fechando com os cartões do topo", trecho)
        self.assertIn("st.warning(", trecho)

    def test_soma_por_plano_passa_pelo_recorte(self):
        """A funcao que soma planos da DIARIO tem de aplicar o recorte por
        loja. Sem esta trava, tirar a chamada nao quebra teste nenhum: o
        numero fica so maior, e maior parece plausivel."""
        i = FONTE.index("def _total_planos(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("_recortar_diario_por_loja(df_diario, lojas)", trecho)
        # E os tres blocos que a usam precisam passar as lojas do contexto.
        for bloco in ("_painel_dept_compras", "_painel_dept_suprimentos"):
            j = FONTE.index(f"def {bloco}(")
            corpo = FONTE[j:FONTE.index("\ndef ", j + 10)]
            self.assertIn('ctx.get("lojas")', corpo, f"{bloco} soma a empresa inteira")
        j = FONTE.index("def _painel_dept_adm(")
        corpo = FONTE[j:FONTE.index("\ndef ", j + 10)]
        self.assertIn("_recortar_diario_por_loja(df_diario_adm", corpo)

    def test_coordenacao_compara_com_a_mediana(self):
        """Mediana, nao media: com uma unidade fora da curva a media se
        desloca e todas as outras parecem boas."""
        i = FONTE.index("def _painel_dept_coordenacao(")
        trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("p.p. vs mediana", trecho)
        self.assertNotIn("sum(margens) / len(margens)", trecho, "virou media")


# ============================================================================
# 5g. FLUXO MENSAL — SALDO DE ABERTURA
# ============================================================================
class TesteSaldoDeAberturaMensal(unittest.TestCase):
    """Em "Movimentos por Mes", caixa e banco mostram o saldo do PRIMEIRO dia
    do mes. Todo o resto do painel segue com a posicao de fechamento."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(["_normalizar_texto", "_classificar_movimento_fin",
                           "_pivot_fluxo_fin"])
        linhas = []
        for dia, saldo in [(1, 500_000.0), (15, 620_000.0), (31, 480_000.0)]:
            linhas.append({"Data Efetiva": pd.Timestamp(2026, 7, dia),
                           "Movimento": "1 - Banco", "Valor.1": saldo})
        for dia, saldo in [(1, 480_000.0), (10, 700_000.0), (18, 655_000.0)]:
            linhas.append({"Data Efetiva": pd.Timestamp(2026, 8, dia),
                           "Movimento": "1 - Banco", "Valor.1": saldo})
        for mes in (7, 8):
            for dia in (5, 20):
                linhas.append({"Data Efetiva": pd.Timestamp(2026, mes, dia),
                               "Movimento": "4 - Contas a Pagar", "Valor.1": -250_000.0})
        cls.df = pd.DataFrame(linhas)
        cls.df["PeriodoMes"] = cls.df["Data Efetiva"].dt.to_period("M")
        cls.meses = sorted(cls.df["PeriodoMes"].unique())

    def _pivo(self, posicao):
        return self.ns["_pivot_fluxo_fin"](
            self.df, "PeriodoMes", "Valor.1", "Movimento", self.meses, posicao_saldo=posicao)

    def test_saldo_do_primeiro_dia_do_mes(self):
        abertura = self._pivo("primeira")
        self.assertEqual(abertura.loc["1 - Banco"].iloc[0], 500_000.0)
        self.assertEqual(abertura.loc["1 - Banco"].iloc[1], 480_000.0)

    def test_padrao_continua_sendo_o_fechamento(self):
        """Quem nao pedir nada tem de continuar recebendo a posicao do ultimo
        dia -- o resto do painel depende disso."""
        fechamento = self.ns["_pivot_fluxo_fin"](
            self.df, "PeriodoMes", "Valor.1", "Movimento", self.meses)
        self.assertEqual(fechamento.loc["1 - Banco"].iloc[0], 480_000.0)
        self.assertEqual(fechamento.loc["1 - Banco"].iloc[1], 655_000.0)

    def test_movimentacao_nao_muda_de_base(self):
        """A troca vale so para linha de SALDO: a pagar e a receber continuam
        somando o mes inteiro nas duas leituras."""
        for posicao in ("primeira", "ultima"):
            pivo = self._pivo(posicao)
            self.assertEqual(pivo.loc["4 - Contas a Pagar"].iloc[0], -500_000.0)

    def test_coluna_final_do_mensal_traz_a_ultima_posicao(self):
        """Caixa e banco na coluna final mostram a posicao MAIS RECENTE. A
        versao anterior pegava o primeiro mes com saldo e exibia julho
        quando agosto ja tinha posicao. Meses de previsao vem zerados, e
        zero e ausencia de dado, nao posicao -- por isso nao contam."""
        caixa = pd.Series({"JUL": 14_184.65, "AGO": 20_318.50, "SET": 0.0, "OUT": 0.0})
        nao_zerados = caixa[caixa != 0]
        self.assertEqual(nao_zerados.iloc[-1], 20_318.50)
        i = FONTE.index("totais_finais_m = {}")
        trecho = FONTE[i:i + 900]
        self.assertIn("nao_zerados.iloc[-1]", trecho,
                      "a coluna final tem de pegar a ultima posicao, nao a primeira")
        self.assertNotIn("nao_zerados.iloc[0]", trecho)

    def test_escrita_do_mapa_nunca_sai_vazia_em_silencio(self):
        """O mapa chegava VAZIO ao navegador porque a escrita falhava e o
        erro morria numa variavel. Agora ha uma ultima defesa para tipos que
        o JSON nao conhece, e o erro viaja ate a tela."""
        i = FONTE.index("def tabela_selecionavel(")
        corpo = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("default=_valor_para_json", corpo,
                      "sem isto, um tipo desconhecido zera o mapa inteiro")
        self.assertIn('data-erro-mapa=', corpo, "o erro precisa chegar na tela")
        self.assertIn("erroDoMapa", corpo)

    def test_saldo_inicial_mensal_so_vale_para_previsao(self):
        """Mes que ja passou, e o mes corrente, tem as posicoes reais de caixa
        e banco DENTRO deles: somar um saldo de abertura por cima contaria o
        mesmo dinheiro duas vezes. Da frente em diante e previsao, e ai o mes
        parte do que sobrou do anterior."""
        mes_corrente = pd.Period("2026-08", "M")
        meses = [pd.Period(m, "M") for m in ("2026-07", "2026-08", "2026-09", "2026-10")]
        movimentos = {"JUL": 4_200.0, "AGO": 5_037.0, "SET": -300.0, "OUT": -400.0}
        colunas = list(movimentos)

        saldos, acumulado = {}, 0.0
        for periodo, coluna in zip(meses, colunas):
            if periodo <= mes_corrente:
                saldos[coluna] = 0.0
                acumulado = movimentos[coluna]
            else:
                saldos[coluna] = acumulado
                acumulado += movimentos[coluna]

        self.assertEqual(saldos["JUL"], 0.0, "mes realizado nao tem abertura")
        self.assertEqual(saldos["AGO"], 0.0, "mes corrente nao tem abertura")
        self.assertEqual(saldos["SET"], 5_037.0, "previsao parte do fechamento anterior")
        self.assertEqual(saldos["OUT"], 4_737.0, "e segue em cadeia")

        i = FONTE.index("📋 Movimentos por Mês")
        trecho = FONTE[max(0, i - 6000):i + 3000]
        self.assertIn("if _periodo <= _mes_corrente_m:", trecho,
                      "a regra precisa distinguir passado/corrente de previsao")
        self.assertIn('index=["SALDO INICIAL"]', trecho)

    def test_so_a_tabela_mensal_usa_a_abertura(self):
        # Janela ampla: o bloco cresceu quando o SALDO INICIAL entrou, e uma
        # janela curta faz o teste falhar por posicao, nao por defeito.
        i = FONTE.index("📋 Movimentos por Mês")
        trecho = FONTE[max(0, i - 6000):i + 3000]
        self.assertIn('posicao_saldo="primeira"', trecho)
        self.assertIn("pivot_m_fechamento = _pivot_fluxo_fin(", trecho)
        # A reserva de caixa nao pode ter passado a ler a abertura.
        j = FONTE.index("Reserva de Caixa — sobra depois de pagar tudo")
        reserva = FONTE[max(0, j - 1800):j]
        self.assertIn("serie_total_geral = pivot_m_fechamento", reserva)
        # E nenhuma outra chamada do pivo pode ter mudado de base.
        self.assertEqual(FONTE.count('posicao_saldo="primeira"'), 1)


# ============================================================================
# 5h. METAS DE RECEBIMENTO E AS TRES LINHAS DE A RECEBER
# ============================================================================
class TesteMetasDeRecebimento(unittest.TestCase):
    """A meta e balizador: aparece como QUANTO FALTA e nunca soma em total."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(
            ["_normalizar_texto", "_classificar_movimento_fin", "_dias_da_meta_no_mes",
             "montar_linhas_de_meta", "_aplicar_meta_como_falta", "_total_geral_sem_meta",
             "_ordenar_movimentos_fin", "_peso_ordem_movimento_fin"],
            ["METAS_RECEBER", "DIAS_DA_SEMANA_POR_MODALIDADE", "MOV_RECEBER_META",
             "MOV_RECEBER_AVENCER", "MOV_RECEBER_LIQUIDADO", "RENOMEAR_MOVIMENTO_FIN",
             "COL_FIN_MOVIMENTO", "COL_FIN_CANAL",
             "COL_FIN_MODALIDADE", "COL_FIN_VALOR", "COL_FIN_VENCIMENTO",
             "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_LIQ_EFETIVA"],
        )

    def test_meta_mensal_bate_com_a_planilha_do_gestor(self):
        """A soma dos dias tem de devolver a meta do mes -- os centavos do
        arredondamento vao no ultimo dia, entao nada se perde."""
        metas = self.ns["METAS_RECEBER"]
        esperado = {("HUB LOGISTICO", 9): 2_850_343.42, ("LOJA", 9): 2_713_551.41,
                    ("VENDA DIRETA", 9): 5_318_893.86, ("LOJA", 7): 1_982_539.27}
        for (canal, mes), alvo in esperado.items():
            soma = sum(v.get((2026, mes), 0.0) for v in metas[canal].values())
            self.assertAlmostEqual(soma, alvo, delta=0.02, msg=f"{canal} mes {mes}")

    def test_soma_dos_dias_devolve_o_mes(self):
        colunas = [self.ns["COL_FIN_MOVIMENTO"], self.ns["COL_FIN_CANAL"],
                   self.ns["COL_FIN_MODALIDADE"], self.ns["COL_FIN_VALOR"], "Data Efetiva",
                   self.ns["COL_FIN_VENCIMENTO"], self.ns["COL_FIN_DATA_LIQUIDACAO"],
                   self.ns["COL_FIN_LIQ_EFETIVA"], "Liquidado", "Tipo Movimento"]
        diarias = self.ns["montar_linhas_de_meta"](colunas)
        setembro = diarias[diarias["Data Efetiva"].dt.to_period("M") == pd.Period("2026-09")]
        self.assertAlmostEqual(setembro[self.ns["COL_FIN_VALOR"]].sum(), 10_882_788.68, delta=0.05)

    def test_boleto_garantido_so_cai_em_terca_e_quinta(self):
        dias = self.ns["_dias_da_meta_no_mes"](2026, 9, "Boleto Garantido")
        self.assertEqual({d.weekday() for d in dias}, {1, 3})
        # As demais modalidades usam dia util, sem fim de semana.
        uteis = self.ns["_dias_da_meta_no_mes"](2026, 9, "Débito")
        self.assertEqual({d.weekday() for d in uteis}, {0, 1, 2, 3, 4})

    def test_meta_e_um_tipo_proprio_fora_dos_totais(self):
        """O nome contem "receber"; sem tratamento proprio ela seria
        classificada como entrada e entraria em todo indicador do painel."""
        self.assertEqual(self.ns["_classificar_movimento_fin"]("2 - Contas a Receber Meta"), "meta")
        self.assertEqual(self.ns["_classificar_movimento_fin"]("3 - Contas a Receber"), "entrada")
        self.assertEqual(
            self.ns["_classificar_movimento_fin"]("3.1 - Contas a Receber Liquidado"), "entrada")

    def _pivo(self, avencer, liquidado, meta=10_882_788.68):
        return pd.DataFrame({"Setembro": {
            self.ns["MOV_RECEBER_META"]: meta,
            self.ns["MOV_RECEBER_AVENCER"]: avencer,
            self.ns["MOV_RECEBER_LIQUIDADO"]: liquidado,
            "4 - Contas a Pagar": -3_500_000.00,
        }})

    def test_meta_mostra_quanto_falta(self):
        ajustado, cheia = self.ns["_aplicar_meta_como_falta"](
            self._pivo(4_000_000.0, 6_000_000.0))
        self.assertAlmostEqual(ajustado.loc[self.ns["MOV_RECEBER_META"], "Setembro"],
                               882_788.68, places=2)
        self.assertAlmostEqual(cheia["Setembro"], 10_882_788.68, places=2)

    def test_meta_zera_quando_batida_e_nao_fica_negativa(self):
        batida, _ = self.ns["_aplicar_meta_como_falta"](self._pivo(4_000_000.0, 6_882_788.68))
        self.assertAlmostEqual(batida.loc[self.ns["MOV_RECEBER_META"], "Setembro"], 0.0, places=2)
        # Um real abaixo da meta: e um real que tem de aparecer.
        quase, _ = self.ns["_aplicar_meta_como_falta"](self._pivo(4_000_000.0, 6_882_787.68))
        self.assertAlmostEqual(quase.loc[self.ns["MOV_RECEBER_META"], "Setembro"], 1.0, places=2)
        # Passando da meta continua zero, nunca negativo.
        passou, _ = self.ns["_aplicar_meta_como_falta"](self._pivo(4_000_000.0, 9_000_000.0))
        self.assertEqual(passou.loc[self.ns["MOV_RECEBER_META"], "Setembro"], 0.0)

    def test_total_geral_ignora_a_meta(self):
        pivo = self._pivo(4_000_000.0, 6_000_000.0)
        total = self.ns["_total_geral_sem_meta"](pivo)["Setembro"]
        self.assertAlmostEqual(total, 4_000_000.0 + 6_000_000.0 - 3_500_000.0, places=2)

    def test_ordem_de_leitura_das_linhas(self):
        """Sequencia fixa: o dinheiro que ja esta, o que deve entrar, o que
        sai. Os nomes vem da planilha como "1.1.Caixa" e "1.Banco", entao
        ordenar pelo numero colocaria o Banco na frente do Caixa."""
        embaralhado = ["2 - Contas a Receber Meta", "3 - Contas a Receber",
                       "3.1 - Contas a Receber Liquidado", "1.1.Caixa", "1.Banco",
                       "4 - Contas a Pagar"]
        self.assertEqual(
            self.ns["_ordenar_movimentos_fin"](embaralhado),
            ["1.1.Caixa", "1.Banco", "2 - Contas a Receber Meta", "3 - Contas a Receber",
             "3.1 - Contas a Receber Liquidado", "4 - Contas a Pagar"],
        )

    def test_ordem_tolera_outras_grafias(self):
        """A ordenacao olha o conteudo do nome, nao o texto exato -- se a
        planilha mudar "1.Banco" para "1 - Banco Bradesco", a ordem segue."""
        ordem = self.ns["_ordenar_movimentos_fin"](
            ["4 - Contas a Pagar", "9 - Aplicação Financeira", "1 - Banco Bradesco",
             "1.1 - Caixa Geral", "3 - Contas a Receber"])
        self.assertEqual(ordem[0], "1.1 - Caixa Geral")
        self.assertEqual(ordem[1], "1 - Banco Bradesco")
        self.assertEqual(ordem[-1], "9 - Aplicação Financeira",
                         "movimento desconhecido tem de cair no fim, nunca no meio")

    def test_liquidado_nao_troca_de_lugar_com_a_vencer(self):
        """O nome do liquidado tambem contem "receber": se a checagem vier
        na ordem errada, as duas linhas trocam de posicao."""
        self.assertEqual(self.ns["_peso_ordem_movimento_fin"]("3 - Contas a Receber"), 3)
        self.assertEqual(
            self.ns["_peso_ordem_movimento_fin"]("3.1 - Contas a Receber Liquidado"), 4)

    def test_diario_ordena_a_meta_junto_com_as_outras(self):
        """No diario a meta e montada a parte, entao ela poderia ser
        empilhada por cima de caixa e banco. Este teste roda a MESMA
        expressao de ordenacao do painel e confere onde a meta cai."""
        achado = re.search(
            r"linhas_do_canal\.sort\(\s*key=(lambda item:.+?)\s*\n\s*\)", FONTE, re.S)
        self.assertIsNotNone(achado, "o diario deixou de ordenar as linhas do canal")
        chave = eval(  # noqa: S307 - a expressao vem do proprio app
            achado.group(1),
            {"_peso_ordem_movimento_fin": self.ns["_peso_ordem_movimento_fin"], "str": str},
        )
        # Nomes SEM numero de propósito: com "1.1.Caixa" e "1.Banco" a ordem
        # alfabetica coincide com a certa por acaso, e o teste passaria mesmo
        # se a ordenacao voltasse a ser por nome. Aqui alfabetico poria Banco
        # antes de Caixa e "A pagar" antes de tudo.
        linhas = [("Contas a Pagar", None), ("Banco Itaú", None),
                  ("Contas a Receber", None), ("Caixa Geral", None),
                  (self.ns["MOV_RECEBER_META"], None),
                  ("Contas a Receber Liquidado", None)]
        linhas.sort(key=chave)
        self.assertEqual(
            [nome for nome, _ in linhas],
            ["Caixa Geral", "Banco Itaú", "2 - Contas a Receber Meta", "Contas a Receber",
             "Contas a Receber Liquidado", "Contas a Pagar"],
        )

    def test_as_duas_linhas_de_a_receber_nao_se_repetem(self):
        """Titulo com baixa vai para a linha de liquidado, na data da baixa;
        sem baixa fica na linha a vencer, na data de vencimento. Se as duas
        seguissem o mesmo eixo, o mesmo titulo apareceria duas vezes."""
        i = FONTE.index("_e_receber = df[COL_FIN_MOVIMENTO] == MOV_RECEBER_AVENCER")
        trecho = FONTE[i:i + 900]
        self.assertIn("MOV_RECEBER_LIQUIDADO", trecho)
        self.assertIn('_e_receber & _tem_baixa, "Data Efetiva"', trecho)
        self.assertIn('_e_receber & ~_tem_baixa, "Data Efetiva"', trecho)

    def test_meta_da_planilha_e_descartada(self):
        i = FONTE.index("_metas_antigas = int(")
        trecho = FONTE[i:i + 300]
        self.assertIn("df = df[df[COL_FIN_MOVIMENTO] != MOV_RECEBER_META]", trecho)
        self.assertIn("montar_linhas_de_meta(df.columns)", FONTE)

    def test_tabela_permite_selecionar_celulas_soltas(self):
        """O st.dataframe so seleciona linha ou coluna inteira. A area precisa
        do comportamento do Excel: clicar numa celula, segurar Ctrl, clicar em
        outras e ver a soma. Por isso a tabela e HTML propria."""
        ns_local = carregar(["formata_brl", "tabela_selecionavel"],
                            ["COLORS", "FONTE_MONO", "ALTURA_LINHA_TABELA_PX",
                             "ALTURA_CABECALHO_TABELA_PX", "ALTURA_BARRA_SOMA_PX"])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        df = pd.DataFrame([[10.0, -20.0], [30.0, 40.0]],
                          index=["1.1.Caixa", "4 - Contas a Pagar"], columns=["18/08", "19/08"])
        ns_local["tabela_selecionavel"](df, chave="t", tipos_linha=["movimento", "movimento"],
                                        linhas_visiveis=21)
        codigo = capturado["codigo"]
        self.assertEqual(len(re.findall(r'data-v="', codigo)), 4, "faltou celula clicavel")
        self.assertIn("evento.ctrlKey || evento.metaKey", codigo, "sem Ctrl+clique")
        self.assertIn("evento.shiftKey", codigo, "sem Shift+clique para intervalo")
        for identificador in ("soma", "media", "qtd"):
            self.assertIn(f'id="{identificador}"', codigo)

    def _altura_para(self, n_linhas):
        ns_local = carregar(["formata_brl", "tabela_selecionavel"],
                            ["COLORS", "FONTE_MONO", "FONTE_PADRAO_TABELA",
                             "TETO_LINHAS_TABELA", "ALTURA_LINHA_TABELA_PX",
                             "ALTURA_CABECALHO_TABELA_PX", "ALTURA_BARRA_SOMA_PX",
                             "ALTURA_BARRA_ROLAGEM_PX"])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        # Colunas suficientes para haver rolagem horizontal, que e o caso das
        # tabelas de verdade (14 a 31 dias no diario).
        colunas = [f"C{i}" for i in range(20)]
        df = pd.DataFrame([[1.0] * len(colunas)] * n_linhas,
                          index=[f"L{i}" for i in range(n_linhas)], columns=colunas)
        ns_local["tabela_selecionavel"](df, chave="t")
        return capturado, ns_local

    def _medir(self, n_linhas, n_colunas):
        ns_local = carregar(["formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        df = pd.DataFrame([[1.0] * n_colunas] * n_linhas,
                          index=[f"L{i}" for i in range(n_linhas)],
                          columns=[f"C{i}" for i in range(n_colunas)])
        ns_local["tabela_selecionavel"](df, chave="t")
        return capturado, ns_local

    def test_altura_cabe_a_tabela_inteira_sem_rolagem(self):
        """A barra de rolagem HORIZONTAL fica dentro da area que rola e come
        altura: sem reservar esse espaco sobra sempre um filete de rolagem
        vertical, que foi o que apareceu na tela."""
        for n_linhas in (6, 21, 25):
            capturado, ns_local = self._altura_para(n_linhas)
            esperado = (ns_local["ALTURA_CABECALHO_TABELA_PX"]
                        + ns_local["ALTURA_LINHA_TABELA_PX"] * n_linhas
                        + ns_local["ALTURA_BARRA_ROLAGEM_PX"] + 2
                        + ns_local["ALTURA_BARRA_SOMA_PX"] + 10)
            self.assertEqual(capturado["altura"], esperado, f"{n_linhas} linhas")

    def test_tabela_pequena_nao_ganha_espaco_vazio(self):
        """Altura fixa em 21 deixaria 15 linhas de vazio numa tabela de 6."""
        pequena, _ = self._altura_para(6)
        grande, _ = self._altura_para(21)
        self.assertLess(pequena["altura"], grande["altura"])

    def test_altura_tem_teto_para_o_caso_extremo(self):
        capturado, ns_local = self._altura_para(60)
        teto = ns_local["TETO_LINHAS_TABELA"]
        esperado = (ns_local["ALTURA_CABECALHO_TABELA_PX"]
                    + ns_local["ALTURA_LINHA_TABELA_PX"] * teto
                    + ns_local["ALTURA_BARRA_ROLAGEM_PX"] + 2
                    + ns_local["ALTURA_BARRA_SOMA_PX"] + 10)
        self.assertEqual(capturado["altura"], esperado)

    def test_altura_da_caixa_fecha_com_o_conteudo(self):
        """A conta em Python e o CSS tem de falar a mesma altura de
        cabecalho. Enquanto o cabecalho herdava os 34px das linhas e a conta
        somava 40, sobrava um filete vazio embaixo da ultima linha."""
        for n_linhas, n_colunas, folga_esperada in [(7, 7, 0), (21, 31, 18)]:
            capturado, ns_local = self._medir(n_linhas, n_colunas)
            css = capturado["codigo"]
            altura_cabecalho = int(
                re.search(r"th\.cabecalho, th\.canto \{ height:(\d+)px", css).group(1))
            altura_linha = int(re.search(r"th, td \{ height:(\d+)px", css).group(1))
            altura_caixa = int(re.search(r"\.rolagem \{ height:(\d+)px", css).group(1))
            conteudo = altura_cabecalho + altura_linha * n_linhas + 2
            self.assertEqual(altura_cabecalho, ns_local["ALTURA_CABECALHO_TABELA_PX"],
                             "CSS e conta discordam da altura do cabecalho")
            self.assertEqual(altura_caixa - conteudo, folga_esperada,
                             f"{n_linhas} linhas x {n_colunas} colunas")

    @staticmethod
    def _fundo_do_seletor(css, seletor):
        """Le o background de UM bloco do CSS. Casar por regex solta pega a
        regra vizinha e da falso positivo -- foi o que aconteceu quando este
        teste dizia que o cabecalho era preto sem ele ser."""
        # Tira os comentarios antes: um deles tem virgula, e a virgula e o
        # separador de seletores -- sem limpar, o bloco certo nao e achado.
        css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
        for bloco in css.split("}"):
            if "{" not in bloco:
                continue
            alvos = [s.strip() for s in bloco.split("{")[0].split(",")]
            if seletor not in alvos:
                continue
            achado = re.search(r"background:([^;]+);", bloco)
            if achado:
                return achado.group(1).strip()
        return None

    def test_corpo_da_tabela_e_preto(self):
        """O painel tem fundo azulado. Deixar a tabela transparente fazia as
        linhas comuns puxarem esse azul e o destaque das consolidacoes se
        perder no meio -- o corpo tem de ser preto."""
        capturado, ns_local = self._medir(5, 5)
        css = capturado["codigo"].split("<div")[0]
        preto = ns_local["FUNDO_TABELA_FLUXO"]
        self.assertEqual(preto, "#000000")
        self.assertEqual(self._fundo_do_seletor(css, ".rolagem"), preto)
        self.assertEqual(self._fundo_do_seletor(css, ".linha-movimento td"), preto)
        self.assertEqual(self._fundo_do_seletor(css, "th.rotulo"), preto)

    def test_regras_de_destaque_encontram_alvo_no_html(self):
        """Ler o CSS nao basta: a regra pode estar escrita e nao valer para
        nada. Foi o que aconteceu -- a classe do tipo estava na CELULA e o
        seletor procurava "td DENTRO de algo com essa classe", entao nenhuma
        linha de consolidacao ficava destacada."""
        from html.parser import HTMLParser

        ns_local = carregar(["formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        df = pd.DataFrame([[1.0, -2.0]] * 5,
                          index=["SALDO INICIAL", "HUB", "1.1.Caixa", "1.Banco", "TOTAL GERAL"],
                          columns=["18/08", "19/08"])
        ns_local["tabela_selecionavel"](
            df, chave="t",
            tipos_linha=["saldo_inicial", "canal", "movimento", "movimento", "total"])

        class Leitor(HTMLParser):
            def __init__(self):
                super().__init__()
                self.linha = None
                self.pares = []

            def handle_starttag(self, tag, attrs):
                classes = dict(attrs).get("class", "")
                if tag == "tr":
                    self.linha = classes
                elif tag in ("td", "th") and self.linha is not None:
                    self.pares.append((self.linha, tag))

        leitor = Leitor()
        leitor.feed(capturado["codigo"])
        for classe in ("linha-saldo_inicial", "linha-canal", "linha-total", "linha-movimento"):
            self.assertIn((classe, "td"), leitor.pares,
                          f"nenhuma celula dentro de tr.{classe} -- a regra do CSS fica morta")
            self.assertIn((classe, "th"), leitor.pares, f"rotulo fora de tr.{classe}")

    def test_consolidacoes_e_cabecalho_ficam_no_tom_claro(self):
        capturado, ns_local = self._medir(5, 5)
        css = capturado["codigo"].split("<div")[0]
        claro = ns_local["COLORS"]["surface_alt"]
        self.assertEqual(self._fundo_do_seletor(css, "th.cabecalho"), claro)
        self.assertEqual(self._fundo_do_seletor(css, "th.canto"), claro)
        for destaque in (".linha-canal td", ".linha-total td", ".linha-saldo_inicial td"):
            self.assertEqual(self._fundo_do_seletor(css, destaque), claro, destaque)

    def test_destaque_nao_pinta_a_coluna_de_nomes(self):
        """O preenchimento fica nos VALORES. Pintar tambem a coluna de nomes
        criava uma faixa clara atravessando a tabela, competindo com os
        numeros -- que sao o que se le ali. O nome se distingue pelo negrito."""
        capturado, _ = self._medir(5, 5)
        css = capturado["codigo"].split("<div")[0]
        for tipo in ("canal", "total", "saldo_inicial"):
            seletor = f".linha-{tipo} th.rotulo"
            self.assertIsNone(self._fundo_do_seletor(css, seletor),
                              f"{seletor} nao pode ter preenchimento")
            self.assertIn(f"{seletor} {{", re.sub(r"/\*.*?\*/", "", css, flags=re.S),
                          f"{seletor} precisa existir para manter o negrito")

    def test_folga_da_rolagem_so_quando_ha_rolagem(self):
        """Com poucas colunas nao aparece barra horizontal, e reservar o
        espaco dela deixava uma faixa vazia embaixo da ultima linha."""
        ns_local = carregar(["formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)

        def altura(n_colunas):
            df = pd.DataFrame([[1.0] * n_colunas] * 7,
                              index=[f"L{i}" for i in range(7)],
                              columns=[f"C{i}" for i in range(n_colunas)])
            ns_local["tabela_selecionavel"](df, chave=f"t{n_colunas}")
            return capturado["altura"]

        poucas = altura(7)   # mensal
        muitas = altura(31)  # diario
        self.assertEqual(muitas - poucas, ns_local["ALTURA_BARRA_ROLAGEM_PX"])

    def test_tabela_usa_a_mesma_fonte_das_outras(self):
        """O iframe nao herda a fonte da pagina. Com monoespacada nos valores
        esta tabela destoava de todas as outras da tela."""
        capturado, _ = self._altura_para(3)
        codigo = capturado["codigo"]
        self.assertNotIn("monospace", codigo, "voltou a fonte monoespacada")
        self.assertIn("Source Sans Pro", codigo, "sem a fonte do app declarada no iframe")
        corpo = codigo.split(".barra")[0]
        self.assertNotIn("text-transform:uppercase", corpo,
                         "cabecalho da tabela nao e caixa alta nas outras telas")

    def test_rotulo_nao_mostra_o_espaco_invisivel(self):
        """O indice do diario usa espacos invisiveis para nao repetir rotulo;
        eles nao podem chegar na tela."""
        ns_local = carregar(["formata_brl", "tabela_selecionavel"],
                            ["COLORS", "FONTE_MONO", "ALTURA_LINHA_TABELA_PX",
                             "ALTURA_CABECALHO_TABELA_PX", "ALTURA_BARRA_SOMA_PX"])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        df = pd.DataFrame([[1.0]], index=["\u200b\u200b    1.1.Caixa"], columns=["18/08"])
        ns_local["tabela_selecionavel"](df, chave="t")
        self.assertNotIn("\u200b", capturado["codigo"])
        self.assertIn(">1.1.Caixa<", capturado["codigo"])


# ============================================================================
# 5m. NOMES DAS LINHAS E A VIRADA DA META
# ============================================================================
class TesteNomesEVirada(unittest.TestCase):

    def test_nomes_das_linhas_do_fluxo(self):
        ns = carregar([], ["MOV_RECEBER_META", "MOV_RECEBER_AVENCER",
                           "MOV_RECEBER_LIQUIDADO", "MOV_PAGAR",
                           "RENOMEAR_MOVIMENTO_FIN"])
        self.assertEqual(ns["MOV_RECEBER_META"], "2 - Contas a Receber Meta")
        self.assertEqual(ns["MOV_RECEBER_AVENCER"], "2.1 - Contas a Receber")
        self.assertEqual(ns["MOV_RECEBER_LIQUIDADO"], "2.2 - Contas a Receber Liquidado")
        self.assertEqual(ns["MOV_PAGAR"], "3 - Contas a Pagar")
        # A planilha continua escrevendo do jeito antigo: a traducao precisa
        # existir, senao o contas a pagar entra com o numero velho.
        self.assertEqual(ns["RENOMEAR_MOVIMENTO_FIN"]["4 - contas a pagar"], "3 - Contas a Pagar")

    def test_ordem_continua_valendo_com_os_nomes_novos(self):
        ns = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                       "_ordenar_movimentos_fin"])
        self.assertEqual(
            ns["_ordenar_movimentos_fin"](
                ["3 - Contas a Pagar", "2.2 - Contas a Receber Liquidado", "1.Banco",
                 "2.1 - Contas a Receber", "2 - Contas a Receber Meta", "1.1.Caixa"]),
            ["1.1.Caixa", "1.Banco", "2 - Contas a Receber Meta", "2.1 - Contas a Receber",
             "2.2 - Contas a Receber Liquidado", "3 - Contas a Pagar"])

    def _falta(self, meta, realizado, dias_visiveis=None, meta_alvo=None):
        """Roda o par de funcoes como as telas rodam: o fator sai da EMPRESA
        e a meta mostrada pode ser a de um canal."""
        ns = carregar(["fracao_da_meta_ainda_nao_coberta", "meta_diaria_que_ainda_falta"])
        fator = ns["fracao_da_meta_ainda_nao_coberta"](meta, realizado)
        return ns["meta_diaria_que_ainda_falta"](
            meta if meta_alvo is None else meta_alvo, fator, dias_visiveis)

    def test_recebivel_do_fim_do_mes_ja_cobre_os_dias_anteriores(self):
        """O contas a receber e posicionado pelo VENCIMENTO. Comparando o
        acumulado ATE o dia, um titulo que vence dia 28 so contava no dia 28
        e os dias anteriores seguiam cobrando meta com o mes ja garantido --
        era o que deixava valor futuro na tela."""
        dias = pd.date_range("2026-08-01", "2026-08-31", freq="D")
        meta = pd.Series([100.0] * 31, index=dias)           # meta do mes: 3.100
        real = pd.Series(0.0, index=dias)
        real[pd.Timestamp("2026-08-28")] = 3500.0            # tudo vence no fim
        falta = self._falta(meta, real)
        self.assertEqual((falta > 0).sum(), 0, "o mes inteiro esta coberto")

    def test_mes_parcialmente_coberto_cobra_o_que_sobra(self):
        """Cobrindo 1.250 de 3.100, os primeiros dias zeram, um fica parcial
        e o resto continua cheio -- e a soma bate com o que falta."""
        dias = pd.date_range("2026-08-01", "2026-08-31", freq="D")
        meta = pd.Series([100.0] * 31, index=dias)
        real = pd.Series(0.0, index=dias)
        real[pd.Timestamp("2026-08-28")] = 1250.0
        falta = self._falta(meta, real)
        self.assertEqual(falta[pd.Timestamp("2026-08-01")], 0.0)
        self.assertEqual(falta[pd.Timestamp("2026-08-12")], 0.0)
        self.assertAlmostEqual(falta[pd.Timestamp("2026-08-13")], 50.0, places=2)
        self.assertAlmostEqual(falta[pd.Timestamp("2026-08-14")], 100.0, places=2)
        self.assertAlmostEqual(falta.sum(), 3100.0 - 1250.0, places=2)

    def test_canal_que_nao_bateu_zera_quando_a_empresa_cobriu(self):
        """A decisao e da EMPRESA: um canal pode nao ter batido a sua parte
        enquanto a empresa ja cobriu o mes."""
        dias = pd.date_range("2026-08-01", "2026-08-31", freq="D")
        meta_empresa = pd.Series([100.0] * 31, index=dias)
        real_empresa = pd.Series(0.0, index=dias)
        real_empresa[pd.Timestamp("2026-08-28")] = 3500.0
        meta_loja = pd.Series([40.0] * 31, index=dias)
        falta_loja = self._falta(meta_empresa, real_empresa, meta_alvo=meta_loja)
        self.assertEqual(falta_loja.sum(), 0.0)

    def test_empresa_longe_da_meta_ninguem_zera(self):
        dias = pd.date_range("2026-08-01", "2026-08-10", freq="D")
        meta = pd.Series([100.0] * 10, index=dias)
        real = pd.Series(0.0, index=dias)
        real.iloc[2] = 50.0
        falta = self._falta(meta, real)
        self.assertAlmostEqual(falta.sum(), 950.0, places=2)

    def test_conta_usa_o_mes_inteiro_e_nao_o_recorte(self):
        """Com o periodo comecando no dia 18, olhar so o recorte esconde o
        que entrou do dia 1 ao 17."""
        dias = pd.date_range("2026-08-01", "2026-08-31", freq="D")
        meta = pd.Series([100.0] * 31, index=dias)
        real = pd.Series(0.0, index=dias)
        real[pd.Timestamp("2026-08-05")] = 3500.0
        visiveis = pd.date_range("2026-08-18", "2026-08-31", freq="D")

        so_recorte = self._falta(meta[visiveis], real[visiveis])
        self.assertTrue((so_recorte > 0).all(),
                        "o cenario precisa mesmo falhar sem o mes inteiro")
        com_mes = self._falta(meta, real, visiveis)
        self.assertEqual(list(com_mes.index), list(visiveis))
        self.assertFalse(com_mes.any())

    def test_cada_mes_e_avaliado_por_si(self):
        """Agosto coberto nao pode zerar setembro."""
        dias_ago = pd.date_range("2026-08-01", "2026-08-31", freq="D")
        dias_set = pd.date_range("2026-09-01", "2026-09-30", freq="D")
        meta = pd.concat([pd.Series([100.0] * 31, index=dias_ago),
                          pd.Series([200.0] * 30, index=dias_set)])
        real = pd.Series(0.0, index=list(dias_ago) + list(dias_set))
        real[pd.Timestamp("2026-08-05")] = 3500.0
        falta = self._falta(meta, real)
        self.assertFalse(falta[:pd.Timestamp("2026-08-31")].any(), "agosto esta coberto")
        self.assertTrue((falta[pd.Timestamp("2026-09-01"):] > 0).all(),
                        "setembro ainda tem meta a cobrar")

    def test_diario_decide_o_fator_uma_vez_no_geral(self):
        i = FONTE.index("with tab_fin_diario:")
        trecho = FONTE[i:FONTE.index("with tab_fin_consolidado:")]
        self.assertIn("_fator_meta_d = fracao_da_meta_ainda_nao_coberta(", trecho)
        self.assertIn("serie_meta, _fator_meta_d, dias_ordenados_d", trecho)
        j = trecho.index("_fator_meta_d = fracao_da_meta_ainda_nao_coberta(")
        self.assertIn("_agrega_no_mes_cheio(df_d_metas)", trecho[j:j + 300],
                      "o fator precisa usar as metas de TODOS os canais")

    def test_as_duas_telas_montam_a_meta_com_o_mes_cheio(self):
        """Trava estrutural: se qualquer uma voltar a somar so o recorte, o
        numero volta a discordar do mensal."""
        i = FONTE.index("with tab_fin_diario:")
        diario = FONTE[i:FONTE.index("with tab_fin_consolidado:")]
        self.assertIn("_agrega_no_mes_cheio(df_meta_canal)", diario)
        self.assertIn("serie_meta, _fator_meta_d, dias_ordenados_d", diario)
        j = FONTE.index("with tab_fin_consolidado:")
        consolidado = FONTE[j:FONTE.index("# ---------------- TESOURARIA", j)]
        self.assertIn("_por_dia_mes_cheio(", consolidado)
        self.assertIn("fracao_da_meta_ainda_nao_coberta(", consolidado)
        self.assertIn("dias_dc)", consolidado,
                      "o consolidado precisa cortar o resultado nos dias da tela")

    def test_linha_da_meta_some_quando_zerada_no_intervalo(self):
        """Se no recorte visivel a meta ja foi batida em todos os dias, a
        linha nao acrescenta nada -- e some. Volta se o periodo alcancar um
        mes em que ainda falta."""
        i = FONTE.index("with tab_fin_consolidado:")
        trecho = FONTE[i:FONTE.index("# ---------------- TESOURARIA", i)]
        self.assertIn("serie_falta_dc = falta if falta.any() else None", trecho)
        self.assertIn("if serie_falta_dc is not None:", trecho)
        j = FONTE.index("falta_meta_canal = meta_diaria_que_ainda_falta(")
        self.assertIn("if falta_meta_canal.any():", FONTE[j:j + 500])


# ============================================================================
# 5l. NOME DAS COLUNAS DA PLANILHA
# ============================================================================
class TesteNomeDasColunas(unittest.TestCase):
    """A planilha muda de escrita de vez em quando ("Data de Liquidação" no
    lugar de "Data Liquidação"). Para o painel e a mesma coluna, e uma
    preposicao a mais nao pode derrubar a tela inteira."""

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(
            ["_normalizar_coluna_fin", "_assinatura_coluna_fin", "resolver_colunas_fluxo"],
            ["_ACENTOS_FIN", "LIGACOES_NOME_COLUNA", "COL_FIN_VALOR", "COL_FIN_MODALIDADE",
             "COL_FIN_CANAL", "COL_FIN_MOVIMENTO", "COL_FIN_DATA_LIQUIDACAO",
             "COL_FIN_VENCIMENTO"],
        )
        cls.esperadas = [cls.ns[k] for k in (
            "COL_FIN_VALOR", "COL_FIN_MODALIDADE", "COL_FIN_CANAL", "COL_FIN_MOVIMENTO",
            "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_VENCIMENTO")]

    def _resolver(self, colunas):
        df = pd.DataFrame({c: [1] for c in colunas})
        return self.ns["resolver_colunas_fluxo"](df, self.esperadas)

    def test_aceita_as_escritas_que_a_planilha_ja_usou(self):
        base = ["Valor.1", "Modalidade", "Canal.1", "Movimento", "Vencimento.1"]
        for escrita in ["Data Liquidação", "Data de Liquidação", "DATA DE LIQUIDACAO",
                        "  Data  de  Liquidação "]:
            _saida, faltando, _ren = self._resolver(base + [escrita])
            self.assertEqual(faltando, [], f"nao reconheceu: {escrita!r}")

    def test_coluna_que_falta_de_verdade_continua_faltando(self):
        """Tolerancia nao pode virar adivinhacao: sem a coluna, o painel tem
        de dizer o que falta, e nao inventar um substituto."""
        _saida, faltando, _ren = self._resolver(
            ["Valor.1", "Modalidade", "Canal.1", "Movimento", "Vencimento.1"])
        self.assertEqual(faltando, [self.ns["COL_FIN_DATA_LIQUIDACAO"]])

    def test_nao_confunde_valor_com_valor_ponto_um(self):
        """A planilha tem pares como "Valor" e "Valor.1". Trocar uma pela
        outra em silencio e bem pior do que a tela de erro."""
        saida, faltando, _ren = self._resolver(
            ["Valor", "Valor.1", "Modalidade", "Canal.1", "Movimento",
             "Data de Liquidação", "Vencimento.1"])
        self.assertEqual(faltando, [])
        self.assertIn("Valor.1", saida.columns)
        self.assertIn("Valor", saida.columns)

    def test_reconhece_os_nomes_antes_de_descartar_colunas(self):
        """ORDEM IMPORTA. A economia de memoria descarta as colunas que o
        painel nao usa; se ela rodar ANTES do reconhecimento, joga fora
        justamente a coluna que ainda nao tem o nome canonico ("Data de
        Liquidacao") e o painel acusa que ela esta faltando. Foi o que
        aconteceu em 20/08/2026."""
        i = FONTE.index("df_fluxo, faltando, colunas_renomeadas = resolver_colunas_fluxo(")
        j = FONTE.index("_descartar = [c for c in df_fluxo.columns")
        self.assertLess(i, j, "descartar colunas antes de reconhecer os nomes quebra o painel")

    def test_a_ordem_certa_salva_a_coluna_com_outra_escrita(self):
        """Prova pelo resultado, e nao so pela posicao no arquivo."""
        colunas_csv = ["Movimento", "Valor.1", "Vencimento.1", "Data de Liquidação",
                       "Canal.1", "Modalidade", "GRUPO DESPESA", "Coluna Extra"]
        uteis = set(self.esperadas) | {"GRUPO DESPESA"}
        df = pd.DataFrame({c: [1] for c in colunas_csv})

        # Ordem errada: descarta primeiro.
        errada = df.drop(columns=[c for c in df.columns if c not in uteis])
        _s, faltando_errada, _r = self.ns["resolver_colunas_fluxo"](errada, self.esperadas)
        self.assertEqual(faltando_errada, [self.ns["COL_FIN_DATA_LIQUIDACAO"]],
                         "o cenario precisa mesmo falhar na ordem errada")

        # Ordem certa: reconhece primeiro.
        certa, faltando_certa, _r = self.ns["resolver_colunas_fluxo"](df, self.esperadas)
        certa = certa.drop(columns=[c for c in certa.columns if c not in uteis])
        self.assertEqual(faltando_certa, [])
        self.assertIn(self.ns["COL_FIN_DATA_LIQUIDACAO"], certa.columns)
        self.assertNotIn("Coluna Extra", certa.columns, "a economia tem de continuar valendo")

    def test_checagem_inicial_tambem_tolera_a_escrita(self):
        i = FONTE.index("ausentes = [")
        trecho = FONTE[i - 400:i + 400]
        self.assertIn("_assinatura_coluna_fin(c) not in _assinaturas", trecho)

    def test_nome_exato_tem_prioridade(self):
        """Se a coluna certa existe com o nome exato, ninguem mexe nela."""
        _saida, faltando, renomeadas = self._resolver(
            ["Valor.1", "Modalidade", "Canal.1", "Movimento",
             "Data Liquidação", "Vencimento.1"])
        self.assertEqual(faltando, [])
        self.assertEqual(renomeadas, {}, "renomeou uma coluna que ja estava certa")


# ============================================================================
# 5j. CONTAS A PAGAR — PROGRAMADO NO FUTURO, EFETIVO NO PASSADO
# ============================================================================
class TesteContasAPagarEfetivo(unittest.TestCase):
    """Titulo ja pago entra no dia do pagamento; em aberto fica no
    vencimento. Sem isso, um dia com R$ 1 milhao vencendo aparecia com o
    milhao inteiro mesmo tendo sido pago dias antes."""

    COL_VAL, COL_VENC, COL_LIQ = "Valor.1", "Vencimento.1", "Data Liquidação"
    COL_EFET, COL_MOV = "Liquidação Efetiva", "Movimento"

    def _base(self):
        return pd.DataFrame({
            self.COL_MOV: ["4 - Contas a Pagar"] * 5,
            self.COL_VENC: pd.to_datetime(["2026-07-31"] * 4 + ["2026-09-15"]),
            self.COL_LIQ: pd.to_datetime(["2026-07-31", "2026-07-28", "2026-07-24", None, None]),
            self.COL_VAL: [-290_000.0, -350_000.0, -300_000.0, -60_000.0, -400_000.0],
        })

    def _aplicar(self, df):
        df[self.COL_EFET] = df[self.COL_LIQ]
        df["Data Efetiva"] = df[self.COL_VENC].fillna(df[self.COL_LIQ])
        e_pagar = df[self.COL_MOV].str.contains("pagar", case=False)
        com_baixa = e_pagar & df[self.COL_EFET].notna()
        df.loc[com_baixa, "Data Efetiva"] = df.loc[com_baixa, self.COL_EFET]
        return df.groupby(df["Data Efetiva"].dt.date)[self.COL_VAL].sum()

    def test_pago_antes_sai_do_dia_do_vencimento(self):
        por_dia = self._aplicar(self._base())
        self.assertAlmostEqual(por_dia[date(2026, 7, 31)], -350_000.0, places=2)
        self.assertAlmostEqual(por_dia[date(2026, 7, 28)], -350_000.0, places=2)
        self.assertAlmostEqual(por_dia[date(2026, 7, 24)], -300_000.0, places=2)

    def test_em_aberto_continua_no_vencimento(self):
        por_dia = self._aplicar(self._base())
        self.assertAlmostEqual(por_dia[date(2026, 9, 15)], -400_000.0, places=2)

    def test_total_do_periodo_nao_muda(self):
        """Mover de dia nao pode criar nem sumir com dinheiro."""
        df = self._base()
        total_antes = df[self.COL_VAL].sum()
        self.assertAlmostEqual(self._aplicar(df).sum(), total_antes, places=2)

    def test_regra_esta_no_preparo_e_usa_a_liquidacao_efetiva(self):
        i = FONTE.index("_pagar_com_baixa = _e_pagar")
        trecho = FONTE[i - 900:i + 600]
        self.assertIn('df.loc[_pagar_com_baixa, "Data Efetiva"] = df.loc[_pagar_com_baixa, COL_FIN_LIQ_EFETIVA]',
                      trecho)
        self.assertIn("COL_FIN_LIQ_EFETIVA].notna()", trecho)
        # A liquidacao efetiva vale para o painel todo, nao so na aba
        # Analises, e junta as duas leituras da coluna do CSV.
        j = FONTE.index("df[COL_FIN_LIQ_EFETIVA] = df[COL_FIN_DATA_LIQUIDACAO]")
        montagem = FONTE[j:j + 400]
        self.assertIn("COL_FIN_LIQ_AMPLA", montagem)

    def test_a_receber_usa_a_mesma_fonte_de_baixa(self):
        """As duas pontas precisam olhar a mesma coluna; se o a receber
        olhasse so o CSV, perderia as baixas que vieram da DIARIO."""
        i = FONTE.index("_e_receber = df[COL_FIN_MOVIMENTO] == MOV_RECEBER_AVENCER")
        trecho = FONTE[i:i + 700]
        self.assertIn("_tem_baixa = df[COL_FIN_LIQ_EFETIVA].notna()", trecho)


# ============================================================================
# 5k. DIARIO CONSOLIDADO — HIERARQUIA COM EXPANDIR
# ============================================================================
class TesteDiarioConsolidado(unittest.TestCase):
    """A empresa inteira dia a dia, com as linhas abrindo em modalidade,
    grupo de despesa ou canal."""

    INDICES = ["SALDO INICIAL", "4 - Contas a Pagar", "Frete", "Energia", "1.1.Caixa",
               "3 - Contas a Receber", "Débito", "Boleto Garantido", "TOTAL GERAL"]
    PAIS = [None, None, 1, 1, None, None, 5, 5, None]
    TIPOS = ["saldo_inicial", "movimento", "movimento", "movimento", "movimento",
             "movimento", "movimento", "movimento", "total"]

    @classmethod
    def setUpClass(cls):
        cls.ns = carregar(["_normalizar_texto", "_classificar_movimento_fin",
                           "_peso_ordem_movimento_fin", "_rotulo_unico_tabela",
                           "_ordenar_com_filhas", "_pais_reordenados"])

    def test_filha_fica_logo_abaixo_da_mae(self):
        """Ordenar so as maes embaralharia as filhas: a filha iria para junto
        de outra linha e passaria a compor um total que nao e o dela."""
        ordem = self.ns["_ordenar_com_filhas"](self.INDICES, self.PAIS, self.TIPOS)
        nomes = [self.INDICES[p] for p in ordem]
        self.assertEqual(nomes[0], "SALDO INICIAL")
        self.assertEqual(nomes[-1], "TOTAL GERAL")
        self.assertEqual(nomes[nomes.index("3 - Contas a Receber") + 1:][:2],
                         ["Débito", "Boleto Garantido"])
        self.assertEqual(nomes[nomes.index("4 - Contas a Pagar") + 1:][:2],
                         ["Frete", "Energia"])

    def test_referencia_da_mae_acompanha_a_reordenacao(self):
        """Depois de reordenar, o indice da mae muda de lugar: sem reescrever
        a referencia, o clique de abrir mostraria as filhas de outra linha."""
        ordem = self.ns["_ordenar_com_filhas"](self.INDICES, self.PAIS, self.TIPOS)
        novos = self.ns["_pais_reordenados"](self.PAIS, ordem)
        for nova, antiga in enumerate(ordem):
            if novos[nova] is None:
                self.assertIsNone(self.PAIS[antiga])
            else:
                mae_antes = self.INDICES[self.PAIS[antiga]]
                mae_depois = self.INDICES[ordem[novos[nova]]]
                self.assertEqual(mae_antes, mae_depois, self.INDICES[antiga])

    def test_rotulo_repetido_nao_quebra_o_indice(self):
        """A mesma modalidade pode aparecer sob duas maes; rotulo repetido
        quebra o indice do DataFrame."""
        rotulos = [self.ns["_rotulo_unico_tabela"]("Débito", i) for i in range(3)]
        self.assertEqual(len(set(rotulos)), 3)
        self.assertTrue(all(r.replace("\u200b", "") == "Débito" for r in rotulos))

    def _montar(self, abertas=()):
        """Monta a tabela como a aba monta: as filhas so entram para as
        linhas marcadas como abertas."""
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "_ordenar_com_filhas",
                             "_pais_reordenados", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        filhas_de = {"4 - Contas a Pagar": ["Frete", "Energia"],
                     "3 - Contas a Receber": ["Débito", "Boleto Garantido"]}
        indices, tipos, pais = [], [], []
        for mae in ["SALDO INICIAL", "1.1.Caixa", "3 - Contas a Receber",
                    "4 - Contas a Pagar", "TOTAL GERAL"]:
            posicao = len(indices)
            indices.append(mae)
            tipos.append("saldo_inicial" if mae == "SALDO INICIAL"
                         else "total" if mae == "TOTAL GERAL" else "movimento")
            pais.append(None)
            if mae in abertas:
                for filha in filhas_de.get(mae, []):
                    indices.append(filha)
                    tipos.append("movimento")
                    pais.append(posicao)
        ordem = ns_local["_ordenar_com_filhas"](indices, pais, tipos)
        df = pd.DataFrame([[1.0] * 31] * len(indices),
                          index=[ns_local["_rotulo_unico_tabela"](indices[p], p) for p in ordem],
                          columns=[f"D{i}" for i in range(31)])
        ns_local["tabela_selecionavel"](
            df, chave="t", tipos_linha=[tipos[p] for p in ordem],
            pais=ns_local["_pais_reordenados"](pais, ordem), filhas_abertas=True,
            comandos_abrir={mae: (mae in abertas, i)
                            for i, mae in enumerate(filhas_de)})
        return capturado, ns_local

    def test_nenhum_metodo_de_teste_duplicado(self):
        """Dois metodos com o mesmo nome numa classe: o ultimo apaga o
        primeiro em silencio. Foi o que aconteceu com um `_falta` antigo,
        que passou a responder pelas chamadas do novo."""
        arvore = ast.parse(open(__file__, encoding="utf-8").read())
        repetidos = []
        for no in ast.walk(arvore):
            if not isinstance(no, ast.ClassDef):
                continue
            vistos = set()
            for item in no.body:
                if isinstance(item, ast.FunctionDef):
                    if item.name in vistos:
                        repetidos.append(f"{no.name}.{item.name}")
                    vistos.add(item.name)
        self.assertEqual(repetidos, [], "metodo definido duas vezes na mesma classe")

    def test_desempacotar_tupla_bate_com_o_que_foi_guardado(self):
        """Uma lista passou a guardar 3 itens por linha e um leitor continuou
        desempacotando em 2. O Python compila, a suite passa, e o app quebra
        NO AR com ValueError -- o erro so existe em tempo de execucao.
        Aconteceu em 20/08/2026 com estilo_linhas_d."""
        from collections import defaultdict

        arvore = ast.parse(FONTE)
        guardados = defaultdict(set)
        for no in ast.walk(arvore):
            if (isinstance(no, ast.Call) and isinstance(no.func, ast.Attribute)
                    and no.func.attr == "append"
                    and isinstance(no.func.value, ast.Name)
                    and len(no.args) == 1 and isinstance(no.args[0], ast.Tuple)):
                guardados[no.func.value.id].add(len(no.args[0].elts))

        problemas = []
        for no in ast.walk(arvore):
            if not isinstance(no, (ast.comprehension, ast.For)):
                continue
            alvo, fonte_iter = no.target, no.iter
            if isinstance(fonte_iter, ast.Call) and fonte_iter.args:
                fonte_iter = fonte_iter.args[0]          # enumerate(lista)
                if isinstance(alvo, ast.Tuple) and len(alvo.elts) == 2:
                    alvo = alvo.elts[1]
            if not isinstance(fonte_iter, ast.Name) or not isinstance(alvo, ast.Tuple):
                continue
            tamanhos = guardados.get(fonte_iter.id)
            if tamanhos and len(alvo.elts) not in tamanhos:
                problemas.append(
                    f"{fonte_iter.id}: guarda {sorted(tamanhos)}, "
                    f"desempacota em {len(alvo.elts)} (linha {no.iter.lineno})")
        self.assertEqual(sorted(set(problemas)), [],
                         "desempacotamento que nao bate com o que foi guardado")

    def test_toda_funcao_chamada_existe(self):
        """Teste que so procura o TEXTO da chamada passa mesmo quando a
        funcao nunca foi definida -- e o app quebra no ar com NameError.
        Aconteceu com _agrega_no_mes_cheio em 19/08/2026: o script que a
        criaria abortou, so a chamada entrou, e a trava estrutural (que
        procurava a chamada) continuou verde."""
        import builtins

        arvore = ast.parse(FONTE)
        definidos = set(dir(builtins))
        for no in ast.walk(arvore):
            if isinstance(no, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                definidos.add(no.name)
                definidos.update(a.arg for a in getattr(no.args, "args", []))
                definidos.update(a.arg for a in getattr(no.args, "kwonlyargs", []))
            elif isinstance(no, ast.Name) and isinstance(no.ctx, ast.Store):
                definidos.add(no.id)
            elif isinstance(no, (ast.Import, ast.ImportFrom)):
                for alias in no.names:
                    definidos.add((alias.asname or alias.name).split(".")[0])
            elif isinstance(no, ast.ExceptHandler) and no.name:
                definidos.add(no.name)

        faltando = {}
        for no in ast.walk(arvore):
            if isinstance(no, ast.Call) and isinstance(no.func, ast.Name):
                if no.func.id not in definidos:
                    faltando.setdefault(no.func.id, no.lineno)
        self.assertEqual(
            faltando, {},
            "funcao chamada e nunca definida: "
            + ", ".join(f"{n}() na linha {l}" for n, l in faltando.items()))

    def test_javascript_gerado_nao_tem_erro_de_sintaxe(self):
        """Um erro de sintaxe no JS nao quebra so o trecho errado: impede o
        script INTEIRO de rodar, e a tabela vira uma imagem -- sem clique,
        sem soma, sem seta. Foi o que aconteceu quando dois blocos escritos
        em momentos diferentes declararam `const ALTURA_LINHA` no mesmo
        escopo. Nenhum teste de Python pegava isso, porque o Python estava
        certo; o defeito nascia no texto que ele montava."""
        for descricao, argumentos in [
            ("tabela simples", dict(tipos_linha=["movimento", "movimento", "total"])),
            ("com hierarquia", dict(tipos_linha=["movimento", "movimento", "total"],
                                    pais=[None, 0, None], filhas_abertas=True,
                                    comandos_abrir={"1.1.Caixa": (True, 0)})),
        ]:
            ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                                 "_rotulo_unico_tabela", "formata_brl",
                                 "tabela_selecionavel"], [])
            capturado = {}
            ns_local["html_embutido"] = dubla_html_embutido(capturado)
            df = pd.DataFrame([[1.0, -2.0]] * 3,
                              index=["1.1.Caixa", "HUB", "TOTAL GERAL"],
                              columns=["18/08", "19/08"])
            ns_local["tabela_selecionavel"](df, chave="t", **argumentos)
            codigo = capturado["codigo"]
            # Do PRIMEIRO <script> ao ULTIMO </script>: a pagina de detalhe
            # que o codigo monta contem um <script> dentro dela, e recortar
            # no primeiro fechamento truncava o texto no meio.
            js = codigo[codigo.index("<script>") + len("<script>"):
                        codigo.rindex("</script>")]

            self.assertEqual(js.count("{"), js.count("}"), f"{descricao}: chaves")
            self.assertEqual(js.count("("), js.count(")"), f"{descricao}: parenteses")

    # Uma string aberta e nao fechada na mesma linha e erro de sintaxe -- e
            # nenhuma contagem de chaves pega isso. Foi o que aconteceu com um
            # "\\r\\n" que o Python transformou em quebra de linha DE VERDADE dentro
            # do JavaScript: a tabela inteira parou de responder e o verificador
            # continuou dizendo que estava tudo certo.
            #
            # Precisa ser um leitor de verdade, e nao contagem de aspas: aspa dentro
            # de aspa e comum ('"' + valor) e a contagem sozinha da falso positivo.
            dentro_de_crase = False
            for numero, linha in enumerate(js.split("\n"), 1):
                aberta = None
                pulo = False
                saltar_ate = -1
                for posicao, caractere in enumerate(linha):
                    if posicao <= saltar_ate:
                        continue
                    if pulo:
                        pulo = False
                        continue
                    if caractere == "\\":
                        pulo = True
                        continue
                    if dentro_de_crase:
                        if caractere == "`":
                            dentro_de_crase = False
                        continue
                    if aberta:
                        if caractere == aberta:
                            aberta = None
                        continue
                    if caractere == "`":
                        dentro_de_crase = True
                    elif caractere in "'\"":
                        aberta = caractere
                    elif caractere == "/":
                        if linha[posicao + 1:posicao + 2] == "/":
                            break
                        # Expressao regular como /"/g: as aspas ali dentro nao abrem
                        # string nenhuma. Sem pular isso, o leitor acusa erro onde
                        # nao ha -- e um verificador que mente e pior que nenhum.
                        anterior = linha[:posicao].rstrip()
                        if not anterior or anterior[-1] in "(,=:[!&|?{;+":
                            fecha = posicao + 1
                            while fecha < len(linha):
                                if linha[fecha] == "\\":
                                    fecha += 2
                                    continue
                                if linha[fecha] == "/":
                                    break
                                fecha += 1
                            saltar_ate = fecha
                if aberta:
                    self.fail(
                        f"{descricao}: linha {numero} abre uma string com {aberta} e nao fecha "
                        f"-- {linha.strip()[:60]}")


            declaracoes = {}
            for linha in js.split("\n"):
                achado = re.match(r"(\s*)(?:const|let)\s+([A-Za-z_$][\w$]*)\s*=", linha)
                if achado:
                    chave = (len(achado.group(1)), achado.group(2))
                    declaracoes[chave] = declaracoes.get(chave, 0) + 1
            repetidas = [nome for (_nivel, nome), vezes in declaracoes.items() if vezes > 1]
            self.assertEqual(repetidas, [],
                             f"{descricao}: declarado mais de uma vez no mesmo escopo "
                             f"-- isso e erro de sintaxe e mata o script inteiro")

    def test_uma_conta_de_altura_so_no_javascript(self):
        """Duas versoes da mesma conta no mesmo script foi o que gerou a
        duplicacao. Uma so, e com nome unico."""
        i = FONTE.index("def tabela_selecionavel(")
        corpo = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertEqual(corpo.count("function ajustarCaixa()"), 1)
        self.assertNotIn("function avisarAltura()", corpo,
                         "voltou a segunda funcao de altura")

    def _mapa_de_lancamentos(self):
        ns = carregar(["montar_lancamentos_por_celula"],
                      ["TETO_LANCAMENTOS_POR_CELULA", "COL_FIN_CANAL", "COL_FIN_MODALIDADE",
                       "COL_FIN_GRUPO_DESPESA", "COL_FIN_VENCIMENTO",
                       "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_VALOR", "COL_FIN_MOVIMENTO",
                       "COL_FIN_NUMERO", "COL_FIN_HISTORICO",
                       "TETO_LANCAMENTOS_NA_PAGINA"])
        dia = pd.Timestamp("2026-08-21")
        df = pd.DataFrame({
            "Data Efetiva": [dia] * 3,
            ns["COL_FIN_MOVIMENTO"]: ["3 - Contas a Pagar"] * 3,
            ns["COL_FIN_NUMERO"]: ["NF 1001", "NF 1002", "NF 1003"],
            ns["COL_FIN_HISTORICO"]: ["FORNECEDOR A LTDA"] * 3,
            ns["COL_FIN_CANAL"]: ["LOJA", "HUB LOGISTICO", "LOJA"],
            ns["COL_FIN_MODALIDADE"]: ["", "", ""],
            ns["COL_FIN_GRUPO_DESPESA"]: ["Ativo Permanente"] * 3,
            ns["COL_FIN_VENCIMENTO"]: [dia] * 3,
            ns["COL_FIN_DATA_LIQUIDACAO"]: [dia, pd.NaT, dia],
            ns["COL_FIN_VALOR"]: [-10_000.00, -4_426.75, -10_000.00],
        })
        mapa, _cortou = ns["montar_lancamentos_por_celula"](
            df, {dia: "21/08"}, df[ns["COL_FIN_GRUPO_DESPESA"]].astype(str))
        return mapa, ns

    def test_lancamentos_da_celula_somam_o_valor_mostrado(self):
        """A conta da aba nova tem de bater com a celula clicada -- e nao
        adianta ela abrir se mostrar outro numero."""
        mapa, _ns = self._mapa_de_lancamentos()
        chave = "Ativo Permanente||21/08"
        self.assertIn(chave, mapa)
        self.assertEqual(len(mapa[chave]), 3)
        self.assertAlmostEqual(sum(l["Valor"] for l in mapa[chave]), -24_426.75, places=2)
        # Procura pelo documento, e nao pela posicao: a lista vem ordenada
        # por valor, para o corte por teto preservar o que explica a celula.
        em_aberto = [l for l in mapa[chave] if l["Número"] == "NF 1002"]
        self.assertEqual(len(em_aberto), 1)
        self.assertEqual(em_aberto[0]["Liquidação"], "",
                         "titulo em aberto tem de aparecer sem data, nao como NaT")

    def test_lancamento_de_outro_dia_nao_entra_na_celula(self):
        ns = carregar(["montar_lancamentos_por_celula"],
                      ["TETO_LANCAMENTOS_POR_CELULA", "COL_FIN_CANAL", "COL_FIN_MODALIDADE",
                       "COL_FIN_GRUPO_DESPESA", "COL_FIN_VENCIMENTO",
                       "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_VALOR", "COL_FIN_MOVIMENTO",
                       "COL_FIN_NUMERO", "COL_FIN_HISTORICO",
                       "TETO_LANCAMENTOS_NA_PAGINA"])
        dias = [pd.Timestamp("2026-08-21"), pd.Timestamp("2026-08-22")]
        df = pd.DataFrame({
            "Data Efetiva": dias,
            ns["COL_FIN_MOVIMENTO"]: ["3 - Contas a Pagar"] * 2,
            ns["COL_FIN_NUMERO"]: ["NF 1", "NF 2"],
            ns["COL_FIN_HISTORICO"]: ["FORNECEDOR A", "FORNECEDOR B"],
            ns["COL_FIN_CANAL"]: ["LOJA", "LOJA"],
            ns["COL_FIN_GRUPO_DESPESA"]: ["Ativo Permanente"] * 2,
            ns["COL_FIN_VENCIMENTO"]: dias,
            ns["COL_FIN_DATA_LIQUIDACAO"]: dias,
            ns["COL_FIN_VALOR"]: [-100.0, -200.0],
        })
        # Só o dia 21 está na tela: o lançamento do 22 fica de fora.
        mapa, _cortou = ns["montar_lancamentos_por_celula"](
            df, {dias[0]: "21/08"}, df[ns["COL_FIN_GRUPO_DESPESA"]].astype(str))
        self.assertEqual(list(mapa), ["Ativo Permanente||21/08"])
        self.assertEqual(len(mapa["Ativo Permanente||21/08"]), 1)

    def test_teto_de_lancamentos_por_celula(self):
        """Eles viajam junto com a pagina: sem teto, uma celula com milhares
        de lancamentos pesaria no carregamento."""
        ns = carregar(["montar_lancamentos_por_celula"],
                      ["TETO_LANCAMENTOS_POR_CELULA", "COL_FIN_CANAL", "COL_FIN_MODALIDADE",
                       "COL_FIN_GRUPO_DESPESA", "COL_FIN_VENCIMENTO",
                       "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_VALOR", "COL_FIN_MOVIMENTO",
                       "COL_FIN_NUMERO", "COL_FIN_HISTORICO",
                       "TETO_LANCAMENTOS_NA_PAGINA"])
        teto = ns["TETO_LANCAMENTOS_POR_CELULA"]
        dia = pd.Timestamp("2026-08-21")
        n = teto + 50
        df = pd.DataFrame({
            "Data Efetiva": [dia] * n,
            ns["COL_FIN_MOVIMENTO"]: ["3 - Contas a Pagar"] * n,
            ns["COL_FIN_NUMERO"]: [f"NF {i}" for i in range(n)],
            ns["COL_FIN_HISTORICO"]: ["FORNECEDOR A"] * n,
            ns["COL_FIN_CANAL"]: ["LOJA"] * n,
            ns["COL_FIN_GRUPO_DESPESA"]: ["Frete"] * n,
            ns["COL_FIN_VENCIMENTO"]: [dia] * n,
            ns["COL_FIN_DATA_LIQUIDACAO"]: [dia] * n,
            ns["COL_FIN_VALOR"]: [-1.0] * n,
        })
        mapa, cortou = ns["montar_lancamentos_por_celula"](
            df, {dia: "21/08"}, df[ns["COL_FIN_GRUPO_DESPESA"]].astype(str))
        self.assertEqual(len(mapa["Frete||21/08"]), teto)
        self.assertTrue(cortou, "a tela precisa saber que cortou, para avisar")

    def test_nenhuma_celula_fica_sem_chave(self):
        """Uma celula sem chave e a falha mais irritante possivel: o valor
        continua sublinhado prometendo detalhe e o duplo clique nao abre.
        Ja aconteceu duas vezes por causa de teto de pagina cortando linhas
        soltas -- por isso o unico limite agora e POR CELULA."""
        ns = carregar(["montar_lancamentos_por_celula"],
                      ["TETO_LANCAMENTOS_POR_CELULA", "COL_FIN_CANAL", "COL_FIN_MODALIDADE",
                       "COL_FIN_GRUPO_DESPESA", "COL_FIN_VENCIMENTO",
                       "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_VALOR", "COL_FIN_MOVIMENTO",
                       "COL_FIN_NUMERO", "COL_FIN_HISTORICO"])
        self.assertNotIn("TETO_LANCAMENTOS_NA_PAGINA", FONTE,
                         "o teto de pagina ja apagou celulas duas vezes")
        n = 6000
        grupos = [f"Grupo {i}" for i in range(40)]
        dias = pd.to_datetime("2026-08-01") + pd.to_timedelta(
            [i % 13 for i in range(n)], unit="D")
        # O "Grupo 0" so tem centavos: por valor, seria o primeiro a cair.
        df = pd.DataFrame({
            "Data Efetiva": dias,
            ns["COL_FIN_MOVIMENTO"]: "3 - Contas a Pagar",
            ns["COL_FIN_NUMERO"]: [f"NF {i}" for i in range(n)],
            ns["COL_FIN_HISTORICO"]: "FORN",
            ns["COL_FIN_CANAL"]: "LOJA",
            ns["COL_FIN_GRUPO_DESPESA"]: [grupos[i % 40] for i in range(n)],
            ns["COL_FIN_VENCIMENTO"]: dias,
            ns["COL_FIN_DATA_LIQUIDACAO"]: dias,
            ns["COL_FIN_VALOR"]: [-0.01 if i % 40 == 0 else -1000.0 - i for i in range(n)],
        })
        rotulos = {d: pd.Timestamp(d).strftime("%d/%m")
                   for d in df["Data Efetiva"].dt.normalize().unique()}
        mapa, _cortou = ns["montar_lancamentos_por_celula"](
            df, rotulos, df[ns["COL_FIN_GRUPO_DESPESA"]].astype(str))
        nos_dados = df.groupby(
            [df[ns["COL_FIN_GRUPO_DESPESA"]], df["Data Efetiva"].dt.normalize()]).ngroups
        self.assertEqual(len(mapa), nos_dados,
                         "toda celula que existe nos dados precisa ter chave")

    def test_detalhe_cobre_linha_fechada_e_linha_aberta(self):
        """A linha aberta mostra o grupo de despesa; a fechada mostra o
        movimento. O detalhe precisa existir para os DOIS rotulos -- antes
        ele dependia de saber quais linhas estavam abertas, e quando essa
        informacao nao chegava, a linha filha ficava sublinhada e o duplo
        clique nao abria nada."""
        ns = carregar(["_normalizar_texto", "_classificar_movimento_fin",
                       "montar_lancamentos_por_celula"],
                      ["TETO_LANCAMENTOS_POR_CELULA", "TETO_LANCAMENTOS_NA_PAGINA",
                       "COL_FIN_CANAL", "COL_FIN_MODALIDADE", "COL_FIN_GRUPO_DESPESA",
                       "COL_FIN_VENCIMENTO", "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_VALOR",
                       "COL_FIN_MOVIMENTO", "COL_FIN_NUMERO", "COL_FIN_HISTORICO"])
        dia = pd.Timestamp("2026-08-25")
        df = pd.DataFrame({
            "Data Efetiva": [dia] * 2,
            ns["COL_FIN_MOVIMENTO"]: ["3 - Contas a Pagar"] * 2,
            ns["COL_FIN_NUMERO"]: ["NF 1", "NF 2"],
            ns["COL_FIN_HISTORICO"]: ["FORN A", "FORN B"],
            ns["COL_FIN_CANAL"]: ["LOJA", "HUB"],
            ns["COL_FIN_MODALIDADE"]: ["", ""],
            ns["COL_FIN_GRUPO_DESPESA"]: ["Custo s/ Venda", "Custo - Mercadoria"],
            ns["COL_FIN_VENCIMENTO"]: [dia] * 2,
            ns["COL_FIN_DATA_LIQUIDACAO"]: [dia] * 2,
            ns["COL_FIN_VALOR"]: [-3000.0, -1_070_186.40],
        })
        mapa = {}
        por_movimento = df[ns["COL_FIN_MOVIMENTO"]].astype(str)
        for rotulagem in (por_movimento,
                          df[ns["COL_FIN_GRUPO_DESPESA"]].astype(str).str.strip()):
            parcial, _c = ns["montar_lancamentos_por_celula"](df, {dia: "25/08"}, rotulagem)
            for chave, linhas in parcial.items():
                mapa.setdefault(chave, linhas)
        # Fechada: o movimento inteiro. Aberta: cada grupo.
        self.assertIn("3 - Contas a Pagar||25/08", mapa)
        self.assertIn("Custo s/ Venda||25/08", mapa)
        self.assertIn("Custo - Mercadoria||25/08", mapa)
        self.assertAlmostEqual(
            sum(l["Valor"] for l in mapa["Custo s/ Venda||25/08"]), -3000.0, places=2)

    def test_codigo_monta_o_detalhe_para_as_duas_rotulagens(self):
        i = FONTE.index("with tab_fin_consolidado:")
        trecho = FONTE[i:FONTE.index("# ---------------- TESOURARIA", i)]
        self.assertIn("_rotulagens = [_por_movimento]", trecho)
        self.assertIn("lancamentos_dc.setdefault(_chave, _linhas)", trecho)
        self.assertIn('lancamentos_dc.pop("||", None)', trecho,
                      "a chave vazia (linhas sem grupo) nao pode virar celula")

    def test_categoria_com_valor_ausente_nao_derruba_o_mapa(self):
        """astype(str) numa coluna CATEGORIA devolve o ausente como nan de
        verdade, nao como o texto "nan" -- entao um replace depois nao pega,
        e esse nan derruba a escrita do mapa INTEIRO. Foi a mensagem que o
        painel finalmente mostrou: "Out of range float values are not JSON
        compliant: nan". Canal, Modalidade e Grupo de Despesa sao categorias."""
        ns = carregar(["montar_lancamentos_por_celula"],
                      ["TETO_LANCAMENTOS_POR_CELULA", "COL_FIN_CANAL", "COL_FIN_MODALIDADE",
                       "COL_FIN_GRUPO_DESPESA", "COL_FIN_VENCIMENTO",
                       "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_VALOR", "COL_FIN_MOVIMENTO",
                       "COL_FIN_NUMERO", "COL_FIN_HISTORICO"])
        dia = pd.Timestamp("2026-08-24")
        df = pd.DataFrame({
            "Data Efetiva": [dia] * 2,
            ns["COL_FIN_MOVIMENTO"]: pd.Categorical(["3 - Contas a Pagar"] * 2),
            ns["COL_FIN_NUMERO"]: ["NF 1", None],
            ns["COL_FIN_HISTORICO"]: pd.Categorical(["FORN A", None]),
            ns["COL_FIN_CANAL"]: pd.Categorical(["LOJA", None]),
            ns["COL_FIN_MODALIDADE"]: pd.Categorical([None, None]),
            ns["COL_FIN_GRUPO_DESPESA"]: pd.Categorical(["Frete", None]),
            ns["COL_FIN_VENCIMENTO"]: [dia, pd.NaT],
            ns["COL_FIN_DATA_LIQUIDACAO"]: [pd.NaT] * 2,
            ns["COL_FIN_VALOR"]: [-79_843.01, float("nan")],
        })
        mapa, _cortou = ns["montar_lancamentos_por_celula"](
            df, {dia: "24/08"}, df[ns["COL_FIN_MOVIMENTO"]].astype(str))
        # allow_nan=False e o que o navegador faz: recusa nan.
        texto = json.dumps(mapa, ensure_ascii=False, allow_nan=False)
        self.assertIn("3 - Contas a Pagar||24/08", json.loads(texto))
        for lancamento in json.loads(texto)["3 - Contas a Pagar||24/08"]:
            for campo, valor in lancamento.items():
                self.assertNotEqual(valor, "nan", f"{campo} virou o texto 'nan'")

    def test_fluxo_mensal_nao_tem_coluna_de_total(self):
        """So os meses. Uma coluna que ora soma (a receber, a pagar) ora
        mostra posicao (caixa, banco) confunde mais do que ajuda."""
        self.assertNotIn('pivot_m["TOTAL / SALDO DE ABERTURA"]', FONTE)
        i = FONTE.index("📋 Movimentos por Mês")
        trecho = FONTE[max(0, i - 6000):i + 2000]
        self.assertNotIn("TOTAL / SALDO DE ABERTURA", trecho)

    def test_mapa_e_sempre_legivel_pelo_navegador(self):
        """UM lancamento com valor vazio bastava para o navegador nao
        conseguir ler o mapa INTEIRO -- Python escreve NaN, que nao e JSON
        valido -- e ai TODA celula respondia "sem detalhe". Foi o defeito
        mais caro desta sessao: o servidor mostrava as chaves certas
        enquanto o navegador estava com o mapa vazio."""
        ns = carregar(["montar_lancamentos_por_celula"],
                      ["TETO_LANCAMENTOS_POR_CELULA", "COL_FIN_CANAL", "COL_FIN_MODALIDADE",
                       "COL_FIN_GRUPO_DESPESA", "COL_FIN_VENCIMENTO",
                       "COL_FIN_DATA_LIQUIDACAO", "COL_FIN_VALOR", "COL_FIN_MOVIMENTO",
                       "COL_FIN_NUMERO", "COL_FIN_HISTORICO"])
        dia = pd.Timestamp("2026-08-24")
        df = pd.DataFrame({
            "Data Efetiva": [dia] * 3,
            ns["COL_FIN_MOVIMENTO"]: ["3 - Contas a Pagar"] * 3,
            ns["COL_FIN_NUMERO"]: ["NF 1", "NF 2", "NF 3"],
            ns["COL_FIN_HISTORICO"]: ["FORN A", "FORN B", "FORN C"],
            ns["COL_FIN_CANAL"]: ["LOJA"] * 3,
            ns["COL_FIN_MODALIDADE"]: [""] * 3,
            ns["COL_FIN_GRUPO_DESPESA"]: ["Frete"] * 3,
            ns["COL_FIN_VENCIMENTO"]: [dia] * 3,
            ns["COL_FIN_DATA_LIQUIDACAO"]: [dia] * 3,
            ns["COL_FIN_VALOR"]: [-79843.01, float("nan"), float("inf")],
        })
        mapa, _cortou = ns["montar_lancamentos_por_celula"](
            df, {dia: "24/08"}, df[ns["COL_FIN_MOVIMENTO"]].astype(str))
        # allow_nan=False e o que o navegador faz: recusa NaN e Infinity.
        texto = json.dumps(mapa, ensure_ascii=False, allow_nan=False)
        self.assertIn("3 - Contas a Pagar||24/08", json.loads(texto))
        self.assertNotIn("NaN", texto)
        self.assertNotIn("Infinity", texto)

    def test_codigo_recusa_escrever_mapa_ilegivel(self):
        """E, ao escrever, o Python tem de RECUSAR NaN em vez de gerar um
        texto que o navegador nao le. Se escapar, a tela avisa."""
        i = FONTE.index("def tabela_selecionavel(")
        corpo = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("allow_nan=False", corpo)
        self.assertIn("não consegui ler os lançamentos", corpo,
                      "o navegador precisa avisar quando a leitura falha")

    def test_celula_com_detalhe_e_clicavel_no_html(self):
        mapa, _ns = self._mapa_de_lancamentos()
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        tabela = pd.DataFrame([[-24_426.75]], index=["Ativo Permanente"], columns=["21/08"])
        ns_local["tabela_selecionavel"](tabela, chave="t", detalhes_por_celula=mapa)
        html = capturado["codigo"]
        # A marca tem de estar na PRÓPRIA célula, não só na folha de estilo:
        # sem ela, nada indica que aquele número abre alguma coisa.
        self.assertRegex(html, r'<td class="[^"]*com-detalhe[^"]*"[^>]*data-k=',
                         "a celula com detalhe precisa se mostrar clicavel")
        self.assertIn('data-k="Ativo Permanente||21/08"', html)
        self.assertIn("dblclick", html, "um clique so continua servindo para somar")
        self.assertIn("painel-detalhe", html)
        embutido = json.loads(html.split('id="detalhes">')[1].split("</script>")[0])
        self.assertIn("Ativo Permanente||21/08", embutido)

    def test_aba_de_lancamentos_segue_o_padrao_do_painel(self):
        """A leitura ali e longa: linhas alternadas para o olho nao pular de
        linha, cabecalho discreto em vez de branco forte, realce sob o
        cursor e o total fixo no rodape."""
        mapa, _ns = self._mapa_de_lancamentos()
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        tabela = pd.DataFrame([[-24_426.75]], index=["Ativo Permanente"], columns=["21/08"])
        ns_local["tabela_selecionavel"](tabela, chave="t", detalhes_por_celula=mapa)
        html = capturado["codigo"]
        for marca, motivo in [
            ("nth-child(even)", "linhas alternadas"),
            ("tbody tr:hover", "realce sob o cursor"),
            ("td.pos", "valor colorido por sinal"),
            ("td.neg", "valor colorido por sinal"),
            ("white-space:normal", "o historico precisa poder quebrar linha"),
            ("position:sticky; bottom:0", "total fixo no rodape"),
            ("FONTE_PADRAO_TABELA" if False else "Source Sans Pro", "mesma fonte do painel"),
        ]:
            self.assertIn(marca, html, motivo)

    def test_aba_tem_um_unico_botao_de_download(self):
        """Um botao so, e o CSV no formato que o Excel em portugues abre sem
        pedir nada: ponto e virgula, virgula decimal e BOM para os acentos."""
        mapa, _ns = self._mapa_de_lancamentos()
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        tabela = pd.DataFrame([[-24_426.75]], index=["Ativo Permanente"], columns=["21/08"])
        ns_local["tabela_selecionavel"](tabela, chave="t", detalhes_por_celula=mapa)
        html = capturado["codigo"]
        self.assertEqual(html.count('id="baixar"'), 1, "tem de ser um botao so")
        self.assertIn("join(';')", html, "separador ponto e virgula")
        self.assertIn("replace('.', ',')", html, "virgula decimal")
        self.assertIn("\ufeff", html, "BOM: sem ele os acentos saem trocados no Excel")
        self.assertIn("URL.createObjectURL", html)
        self.assertIn("link.download", html)

    def test_duplo_clique_abre_a_aba_com_window_open(self):
        """ABRIR JANELA e permitido; LEVAR PARA OUTRA PAGINA nao e. O console
        acusou "Unsafe attempt to initiate navigation" quando troquei
        window.open por um link -- eu confundi as duas coisas e quebrei o que
        funcionava. Este teste RODA o JavaScript para garantir que o caminho
        e o permitido."""
        import shutil
        import subprocess
        import tempfile

        node = shutil.which("node") or shutil.which("nodejs")
        if not node:
            self.skipTest("sem motor JavaScript nesta maquina")

        mapa, _ns = self._mapa_de_lancamentos()
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        tabela = pd.DataFrame([[-24_426.75]], index=["Ativo Permanente"], columns=["21/08"])
        ns_local["tabela_selecionavel"](tabela, chave="t", detalhes_por_celula=mapa)
        html = capturado["codigo"]
        inicio = html.index("<script>", html.index('id="detalhes"')) + len("<script>")
        js = html[inicio:html.index("</script>", inicio)]

        # ABRIR JANELA e permitido ("allow-popups"); LEVAR PARA OUTRA PAGINA
        # nao e ("allow-top-navigation" fica de fora). Sao coisas diferentes,
        # e confundi-las custou varios dias: window.open funciona, link e
        # window.parent.open viram navegacao e o navegador barra.
        self.assertIn("window.open('', '_blank')", js,
                      "abrir janela e o caminho permitido")
        # Sem comentarios: eles CITAM os padroes proibidos de proposito, para
        # que o proximo a mexer saiba por que nao usar.
        sem_comentario = re.sub(r"//[^\n]*", "", js)
        for proibido in ("location.href =", "target = '_blank'", "window.parent.open("):
            self.assertNotIn(proibido, sem_comentario,
                             f"{proibido} vira navegacao e o navegador barra")
        # Ler o documento de fora NAO e navegacao -- e o que faz a seta
        # funcionar. Por isso nao entra na lista acima.

        roteiro = """
        let ouvinteDaCelula = null;
        const celula = {dataset:{k:'Ativo Permanente||21/08',v:'-1',l:'0',c:'0'},
          style:{}, classList:{add(){},remove(){},toggle(){},contains(){return false}},
          addEventListener(t,f){ if (t === 'dblclick') ouvinteDaCelula = f; },
          closest(s){return s==='td[data-k]'?this:null}};
        const jsonTag = {textContent: JSON.stringify(
          {'Ativo Permanente||21/08':[{'Valor':-1}]})};
        const doc = {}; const corpo = {innerHTML:'', style:{cssText:''}};
        const elemento = () => ({style:{cssText:''}, textContent:'', dataset:{},
          classList:{add(){},remove(){},toggle(){},contains(){return false}},
          addEventListener(){}, click(){}, appendChild(){}, querySelector(){return null}});
        global.document = {addEventListener(t,f){doc[t]=f},
          getElementById(id){return id==='detalhes'?jsonTag:elemento()},
          querySelector(){return elemento()},
          querySelectorAll(s){return (s.includes('data-v')||s.includes('data-k'))?[celula]:[]},
          createElement: elemento, body: corpo, documentElement:{scrollHeight:1}};
        let abriuJanela = false;
        global.window = {parent:{postMessage(){}}, top:null, frameElement:null,
          location:{href:'http://x/', reload(){}},
          open(){abriuJanela = true; return {document:{write(){},close(){}}}}};
        global.URL={createObjectURL(){return 'blob:'},revokeObjectURL(){}};
        global.Blob=function(){};
        eval(JS_DA_TABELA);
        ouvinteDaCelula({target: celula, preventDefault(){}});
        console.log(abriuJanela ? 'abriu janela' : 'nao abriu');
        """
        with tempfile.TemporaryDirectory() as pasta:
            caminho = f"{pasta}/t.js"
            with open(caminho, "w", encoding="utf-8") as arquivo:
                arquivo.write("const JS_DA_TABELA = " + json.dumps(js) + ";\n" + roteiro)
            saida = subprocess.run([node, caminho], capture_output=True, text=True)
        self.assertEqual(saida.returncode, 0, f"o JavaScript nao rodou: {saida.stderr[:300]}")
        self.assertEqual(saida.stdout.strip(), "abriu janela",
                         "o duplo clique tem de abrir a aba com os lancamentos")

    def test_abertura_da_aba_e_o_minimo_possivel(self):
        """Tres linhas: abre, escreve, fecha -- exatamente a versao em que a
        aba abria. Cada verificacao que acrescentei aqui (document.open
        antes, reconferir se a janela seguia aberta) foi uma chance a mais de
        o caminho ser abandonado no meio, e nenhuma resolveu nada."""
        mapa, _ns = self._mapa_de_lancamentos()
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        tabela = pd.DataFrame([[-24_426.75]], index=["Ativo Permanente"], columns=["21/08"])
        ns_local["tabela_selecionavel"](tabela, chave="t", detalhes_por_celula=mapa)
        html = capturado["codigo"]
        # CAMINHO COMPROVADO POR PRINT (20/08/2026): a aba abriu com o
        # endereco "about:blank" -- janela em branco com a pagina escrita
        # dentro. Cada variacao que tentei aqui quebrou o que funcionava:
        #   - arquivo em memoria: devolve janela mesmo quando o navegador
        #     barra o conteudo, entao o codigo nunca cai no jeito certo
        #   - link ou window.parent.open: viram navegacao da pagina, que o
        #     quadro nao tem permissao de fazer
        #   - document.open() antes de escrever, reconferir se a janela
        #     seguia aberta: so criam caminhos para desistir no meio
        self.assertIn("const aba = window.open('', '_blank');", html)
        self.assertIn("aba.document.write(pagina);", html)
        for proibido in ("window.open(enderecoPagina", "aba.document.open()",
                         "!aba.closed", "window.parent.open("):
            self.assertNotIn(proibido, html,
                             f"{proibido} ja quebrou a abertura da aba antes")

    def test_duplo_clique_tem_caminho_alternativo(self):
        """O navegador pode recusar a abertura da aba vinda de dentro do
        quadro -- e recusa em SILENCIO. Sem um plano B, o duplo clique
        simplesmente nao fazia nada e nao havia como saber por que."""
        mapa, _ns = self._mapa_de_lancamentos()
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        tabela = pd.DataFrame([[-24_426.75]], index=["Ativo Permanente"], columns=["21/08"])
        ns_local["tabela_selecionavel"](tabela, chave="t", detalhes_por_celula=mapa)
        html = capturado["codigo"]
        self.assertIn("window._corpoOriginal", html)
        self.assertIn("Voltar para a tabela", html, "precisa de volta para a tabela")
        self.assertIn("URL.createObjectURL", html, "o CSV continua disponivel")

    def test_duplo_clique_usa_delegacao_e_nunca_fica_mudo(self):
        """Ouvinte por celula depende do momento em que o script roda; a
        delegacao no documento nao. E o caminho de "sem detalhe" precisa
        DIZER alguma coisa: o silencio ali atrasou o diagnostico por dois
        turnos seguidos."""
        mapa, _ns = self._mapa_de_lancamentos()
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        tabela = pd.DataFrame([[-24_426.75]], index=["Ativo Permanente"], columns=["21/08"])
        ns_local["tabela_selecionavel"](tabela, chave="t", detalhes_por_celula=mapa)
        html = capturado["codigo"]
        self.assertIn("querySelectorAll('td[data-k]')", html,
                      "um ouvinte por celula, como na versao em que a aba abria")
        self.assertIn("sem detalhe para", html,
                      "o caminho sem detalhe precisa avisar, nao ficar mudo")

    def test_tela_mostra_quantas_celulas_tem_detalhe(self):
        """Quando o duplo clique nao responde, a primeira pergunta e se o
        detalhe chegou a ser montado."""
        i = FONTE.index("with tab_fin_consolidado:")
        trecho = FONTE[i:FONTE.index("# ---------------- TESOURARIA", i)]
        self.assertIn("Detalhe por célula — diagnóstico", trecho)
        # E so para quem administra: e uma ferramenta de investigacao, nao
        # informacao para quem usa o painel no dia a dia.
        posicao = trecho.index("Detalhe por célula — diagnóstico")
        self.assertIn("if eh_admin:", trecho[max(0, posicao - 300):posicao])
        self.assertIn("células com detalhe:", trecho)
        self.assertIn("com dia reconhecido:", trecho,
                      "o diagnostico precisa dizer onde o detalhe se perde")

    def test_celula_sem_detalhe_nao_finge_ser_clicavel(self):
        ns_local = carregar(["_normalizar_texto", "_peso_ordem_movimento_fin",
                             "_rotulo_unico_tabela", "formata_brl", "tabela_selecionavel"], [])
        capturado = {}
        ns_local["html_embutido"] = dubla_html_embutido(capturado)
        tabela = pd.DataFrame([[1.0]], index=["Sem detalhe"], columns=["21/08"])
        ns_local["tabela_selecionavel"](tabela, chave="t", detalhes_por_celula={})
        self.assertNotIn("com-detalhe\"", capturado["codigo"])
        self.assertNotIn("data-k=", capturado["codigo"])

    def test_seta_fica_na_propria_linha(self):
        """A escolha do que abrir fica DENTRO da tabela, na setinha da
        linha -- nao num campo separado acima dela."""
        capturado, _ = self._montar()
        html = capturado["codigo"]
        self.assertEqual(len(re.findall(r'class="seta comando', html)), 2,
                         "so as maes com detalhe podem ter seta")
        self.assertIn("data-linha=", html)
        i = FONTE.index("with tab_fin_consolidado:")
        trecho = FONTE[i:FONTE.index("# ---------------- TESOURARIA", i)]
        self.assertNotIn("st.multiselect", trecho,
                         "voltou o campo separado para escolher o que abrir")

    def test_clique_na_seta_faz_o_quadro_crescer(self):
        """A altura do quadro e decidida ao desenhar: um pedido de crescer
        vindo de dentro dele e ignorado. Por isso o clique navega a pagina e
        a tabela volta redesenhada -- com as filhas contando na altura."""
        fechado, ns_local = self._montar()
        aberto, _ = self._montar({"4 - Contas a Pagar"})
        self.assertEqual(aberto["altura"] - fechado["altura"],
                         ns_local["ALTURA_LINHA_TABELA_PX"] * 2,
                         "abrir uma linha com 2 filhas tem de crescer 2 linhas")
        self.assertNotIn('style="display:none"', aberto["codigo"],
                         "filha de linha aberta nao pode nascer escondida")
        self.assertIn("alvo.click()", aberto["codigo"],
                      "o clique precisa acionar o botao da pagina")

    def test_seta_da_linha_aberta_vem_girada(self):
        aberto, _ = self._montar({"4 - Contas a Pagar"})
        self.assertIn("seta comando aberta", aberto["codigo"])

    def test_separador_da_url_nao_quebra_nome_com_virgula(self):
        """Nome de modalidade tem virgula; com virgula de separador, uma
        linha aberta viraria duas e nenhuma seria encontrada."""
        ns_local = carregar(["linhas_abertas_da_url"],
                            ["PARAM_LINHAS_ABERTAS", "SEPARADOR_LINHAS_ABERTAS"])
        self.assertNotEqual(ns_local["SEPARADOR_LINHAS_ABERTAS"], ",")
        ns_local["st"] = type("st", (), {"query_params": {
            ns_local["PARAM_LINHAS_ABERTAS"]:
                "3 - Contas a Receber~Crédito à Vista, parcelado"}})()
        self.assertEqual(ns_local["linhas_abertas_da_url"](),
                         ["3 - Contas a Receber", "Crédito à Vista, parcelado"])

    def test_seta_aciona_um_botao_da_pagina(self):
        """O quadro da tabela NAO pode trocar o endereco da pagina -- o
        navegador recusa a navegacao sem avisar, e foi por isso que a seta
        ficou muda mesmo com o resto do script funcionando. O que ele pode e
        apertar um botao que ja existe na pagina, porque as duas sao da
        mesma origem."""
        ns_local = carregar(["botoes_de_abrir"], ["PREFIXO_BOTAO_ABRIR"])
        estado, estilos = {}, []

        class FakeST:
            session_state = estado
            clicado = None

            def markdown(self, html, **kw):
                estilos.append(html)

            def button(self, rotulo, key=None, **kw):
                return key == FakeST.clicado

        ns_local["st"] = FakeST()
        linhas = ["1.1.Caixa", "4 - Contas a Pagar"]

        comandos = ns_local["botoes_de_abrir"](linhas, "dc")
        self.assertEqual(comandos, {"1.1.Caixa": (False, 0), "4 - Contas a Pagar": (False, 1)})
        self.assertIn("left:-9999px", estilos[0], "o botao tem de ficar fora da tela")

        FakeST.clicado = "dcabrir_0"
        comandos = ns_local["botoes_de_abrir"](linhas, "dc")
        self.assertTrue(comandos["1.1.Caixa"][0], "o clique tinha de abrir a linha")
        comandos = ns_local["botoes_de_abrir"](linhas, "dc")
        self.assertFalse(comandos["1.1.Caixa"][0], "clicar de novo tinha de fechar")

    def test_js_procura_o_botao_e_tem_plano_b(self):
        i = FONTE.index("def tabela_selecionavel(")
        corpo = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("{PREFIXO_BOTAO_ABRIR}' + indice + ' button'", corpo,
                      "o JS precisa achar o botao pela classe da chave")
        self.assertIn("innerText.trim() ===", corpo,
                      "sem a classe por chave, o plano B e achar pelo texto")
        self.assertIn("alvo.click()", corpo)

    def test_botao_e_desenhado_antes_da_tabela(self):
        """O Streamlit conta o clique no ponto em que o botao aparece: se ele
        vier depois, a tabela e montada com o estado velho e o clique so faz
        efeito na proxima interacao."""
        i = FONTE.index("with tab_fin_consolidado:")
        trecho = FONTE[i:FONTE.index("# ---------------- TESOURARIA", i)]
        posicao_botoes = trecho.index("botoes_de_abrir(")
        posicao_tabela = trecho.index("tabela_selecionavel(")
        self.assertLess(posicao_botoes, posicao_tabela)

    def test_periodo_sobrevive_ao_clique_na_seta(self):
        """Abrir uma linha recarrega a pagina, e recarregar apaga o que
        estava nos campos. Sem guardar o periodo na URL, cada abertura
        jogaria as datas de volta para o padrao."""
        ns_local = carregar(["guardar_periodo_na_url"], [])
        guardadas = {}
        ns_local["st"] = type("st", (), {"query_params": guardadas})()
        ns_local["datetime"] = datetime
        ns_local["guardar_periodo_na_url"]("dc", date(2026, 7, 20), date(2026, 8, 15))
        self.assertEqual(guardadas, {"dc_ini": "2026-07-20", "dc_fim": "2026-08-15"})
        self.assertEqual(ns_local["guardar_periodo_na_url"]("dc"),
                         (date(2026, 7, 20), date(2026, 8, 15)))
        guardadas.clear()
        self.assertEqual(ns_local["guardar_periodo_na_url"]("dc"), (None, None),
                         "URL sem o periodo nao pode explodir")
        i = FONTE.index("with tab_fin_consolidado:")
        trecho = FONTE[i:FONTE.index("# ---------------- TESOURARIA", i)]
        self.assertIn('guardar_periodo_na_url("dc", data_ini_dc, data_fim_dc)', trecho)
        # Sem esta condicao, TODAS as linhas abririam sempre -- a tabela
        # nasceria com 60 linhas e o clique na seta nao faria diferenca.
        self.assertIn("if movimento in abertas_dc and coluna_abertura in recorte.columns:",
                      trecho, "a abertura deixou de respeitar a URL")
        self.assertIn("if MOV_RECEBER_META in abertas_dc else []", trecho)

    def test_total_do_fluxo_diario_soma_o_mes_e_nao_o_recorte(self):
        """Esta aba mostra UM mes de cada vez (misturar meses e o papel do
        Diario Consolidado). A leitura da coluna final e "quanto este
        movimento soma no mes", nao "quanto soma nos dias que sobraram na
        tela": olhando de 19 a 31/08, o total traz agosto inteiro."""
        i = FONTE.index("with tab_fin_diario:")
        trecho = FONTE[i:FONTE.index("with tab_fin_consolidado:")]
        self.assertIn("def _total_do_mes_d(", trecho)
        self.assertIn("_total_do_mes_d(canal_da_linha, nome_limpo)", trecho,
                      "a linha de movimento tem de somar o mes")
        self.assertIn("_fluxo_do_mes_por_canal_d", trecho, "a linha de canal tambem")
        # A base e a COMPLETA, e o calculo e UM agrupamento -- nao um filtro
        # da base inteira por linha da tabela, que foi como fiz primeiro.
        self.assertIn("_mes_cheio_d = df_d_completo[", trecho)
        self.assertIn("_soma_por_canal_e_movimento_d = _mes_cheio_d.groupby(", trecho)
        j = trecho.index("def _total_do_mes_d(")
        corpo = trecho[j:j + 900]
        self.assertNotIn("df_d_completo[", corpo,
                         "filtrar a base dentro da funcao volta a varrer tudo por linha")
        self.assertIn("TOTAL DO MÊS / ÚLT. POSIÇÃO", trecho,
                      "o titulo precisa dizer que e do mes, senao engana")

    def test_saldo_continua_sendo_posicao_e_nao_soma(self):
        """Saldo de caixa e banco e POSICAO num dia: somar dias diferentes
        nao faz sentido, entao essas linhas seguem mostrando a ultima."""
        i = FONTE.index("with tab_fin_diario:")
        trecho = FONTE[i:FONTE.index("with tab_fin_consolidado:")]
        self.assertIn('if _classificar_movimento_fin(nome_limpo) in ("saldo", "aplicacao"):',
                      trecho)
        posicao = trecho.index('in ("saldo", "aplicacao"):')
        self.assertIn("nao_zerados.iloc[-1]", trecho[posicao:posicao + 400],
                      "saldo tem de continuar pegando a ultima posicao")

    def test_as_duas_abas_diarias_abrem_no_mesmo_dia(self):
        """Ontem, nas duas. O movimento do dia corrente costuma estar
        incompleto; e, se as abas abrissem em dias diferentes, a mesma
        pergunta daria respostas diferentes conforme a aba."""
        for ancora, variavel in [("with tab_fin_diario:", "data_ontem_d"),
                                 ("with tab_fin_consolidado:", "_ontem_dc")]:
            i = FONTE.index(ancora)
            trecho = FONTE[i:i + 2600]
            self.assertIn(f"{variavel} = ", trecho, ancora)
            self.assertIn("Timedelta(days=1)", trecho, f"{ancora}: nao comeca em ontem")
            self.assertIn("MonthEnd(1)", trecho, f"{ancora}: nao vai ate o fim do mes seguinte")
            # Nao basta calcular "ontem": o seletor tem de RECEBER esse
            # valor. Calcular e nao usar passaria despercebido.
            self.assertRegex(trecho, rf"padrao_ini=[^,]*{re.escape(variavel)}",
                             f"{ancora}: o padrao nao usa o dia de ontem")

    def test_periodo_por_datas_nas_duas_abas_diarias(self):
        """As duas abas diarias usam o mesmo seletor de duas datas, que pode
        atravessar meses -- julho e agosto na mesma tela."""
        self.assertIn("def seletor_periodo_dias(", FONTE)
        for ancora in ("with tab_fin_diario:", "with tab_fin_consolidado:"):
            i = FONTE.index(ancora)
            trecho = FONTE[i:i + 3000]
            self.assertIn("seletor_periodo_dias(", trecho, ancora)
        i = FONTE.index("def seletor_periodo_dias(")
        corpo = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn("min_value=limite_ini", corpo)
        self.assertIn("data_ini, data_fim = data_fim, data_ini", corpo,
                      "datas invertidas tem de ser ordenadas, nao virar erro")

    def test_aba_existe_com_as_duas_formas_de_abrir(self):
        self.assertIn("📆 Diário Consolidado", FONTE)
        i = FONTE.index("with tab_fin_consolidado:")
        trecho = FONTE[i:FONTE.index("# ---------------- TESOURARIA", i)]
        self.assertIn("COL_FIN_GRUPO_DESPESA", trecho, "contas a pagar abre por grupo")
        self.assertIn("COL_FIN_MODALIDADE", trecho, "contas a receber abre por modalidade")
        self.assertIn("COL_FIN_CANAL", trecho, "falta a opcao de abrir por canal")
        self.assertIn('"Detalhe da linha", "Canal"', trecho)
        # A meta nao pode entrar no total desta aba tambem.
        self.assertIn('rotulo != MOV_RECEBER_META', trecho)


# ============================================================================
# 5i. MEMORIA E LEITURA DAS PLANILHAS
# ============================================================================
class TesteMemoriaELeitura(unittest.TestCase):
    """O servidor derruba o app quando a memoria estoura, e a tela e um erro
    generico sem causa. Estas travas protegem as duas frentes: ler menos e
    soltar memoria antes do limite."""

    def test_abas_sao_lidas_com_a_planilha_aberta_uma_vez(self):
        """Cada `pd.read_excel(caminho, sheet_name=...)` reprocessa o arquivo
        INTEIRO. Carregar 13 lojas dos dois arquivos eram 26 leituras
        completas -- 26 picos de memoria."""
        for funcao in ("carregar_dados_abas", "carregar_dados_por_loja"):
            i = FONTE.index(f"def {funcao}(")
            trecho = FONTE[i:FONTE.index("\ndef ", i + 10)]
            self.assertIn("_planilha_aberta(path_o)", trecho, funcao)
            self.assertIn("_planilha_aberta(path_r)", trecho, funcao)
            self.assertNotIn("_ler_aba_ou_vazio(path_o", trecho,
                             f"{funcao} voltou a reabrir o arquivo por aba")

    def test_caches_pesados_tem_teto_de_entradas(self):
        """Sem max_entries, cada combinacao de abas guarda uma copia inteira
        dos dados e elas se acumulam ate o app cair."""
        import re as _re
        for funcao in ("carregar_dados_abas", "carregar_dados_por_loja", "carregar_diario",
                       "preparar_fluxo_caixa"):
            i = FONTE.index(f"def {funcao}(")
            decorador = FONTE[max(0, i - 400):i]
            achado = _re.search(r"@st\.cache_\w+\([^)]*\)\s*$", decorador)
            self.assertIsNotNone(achado, f"{funcao} sem decorador de cache")
            self.assertIn("max_entries", achado.group(0), f"{funcao} sem teto de entradas")

    def test_csv_do_fluxo_e_lido_em_formato_economico(self):
        """O servidor corta o app perto de 1 GB. As colunas de texto que se
        repetem ao longo das ~650 mil linhas entram como CATEGORIA: cada
        valor distinto e guardado uma vez. Medido em 19/08/2026: 57% menos
        memoria por copia do fluxo."""
        ns = carregar([], ["TIPOS_ECONOMICOS_FLUXO"])
        for coluna in ("Movimento", "Canal.1", "Modalidade", "GRUPO DESPESA"):
            self.assertEqual(ns["TIPOS_ECONOMICOS_FLUXO"].get(coluna), "category", coluna)
        self.assertIn("dtype=TIPOS_ECONOMICOS_FLUXO", FONTE)
        self.assertIn("low_memory=False", FONTE)

    def test_colunas_sem_uso_sao_descartadas(self):
        i = FONTE.index("_descartar = [c for c in df_fluxo.columns")
        self.assertIn("df_fluxo.drop(columns=_descartar)", FONTE[i:i + 300])

    def test_nao_converte_coluna_inteira_para_texto(self):
        """`astype(str)` numa coluna categoria desfaz a economia: cria uma
        copia inteira em texto. Dentro de um laco por canal isso acontece
        dezenas de vezes por tela."""
        self.assertNotIn(".dropna().astype(str).unique()", FONTE,
                         "converte a coluna toda antes do unique() -- inverta a ordem")
        i = FONTE.index("with tab_fin_diario:")
        bloco = FONTE[i:FONTE.index("# ---------------- TESOURARIA", i)]
        self.assertNotIn("].astype(str) ==", bloco,
                         "comparacao com categoria nao precisa virar texto")
        self.assertNotIn("].astype(str).isin(", bloco)

    def test_plano_de_contas_continua_fora(self):
        """Nenhuma tela do fluxo usa o Plano de Contas, e ele e um texto
        diferente por lancamento, 650 mil vezes."""
        i = FONTE.index("_uteis = {")
        bloco = FONTE[i:FONTE.index("}", i)]
        self.assertNotIn("COL_FIN_PLANO_CONTAS", bloco,
                         "voltou a ser carregada sem nenhuma tela usar")
        self.assertLessEqual(FONTE.count("COL_FIN_PLANO_CONTAS"), 1,
                             "passou a ser usada -- precisa voltar para _uteis")

    def test_historico_e_numero_so_existem_para_o_detalhamento(self):
        """Elas custam memoria (cerca de 39 MB por copia) e so se justificam
        por causa do detalhamento por celula, onde respondem "de quem e este
        valor". Se o uso sumir, precisam sair de novo."""
        i = FONTE.index("_uteis = {")
        bloco = FONTE[i:FONTE.index("}", i)]
        for coluna in ("COL_FIN_HISTORICO", "COL_FIN_NUMERO"):
            self.assertIn(coluna, bloco, f"{coluna} precisa ser carregada para o detalhamento")
        j = FONTE.index("def montar_lancamentos_por_celula(")
        corpo = FONTE[j:FONTE.index("\ndef ", j + 10)]
        self.assertIn("COL_FIN_NUMERO", corpo, "o Numero precisa aparecer no detalhamento")
        self.assertIn("COL_FIN_HISTORICO", corpo)
        self.assertIn('"Fornecedor / Histórico"', corpo)

    def test_historico_entra_como_categoria_e_numero_nao(self):
        """O mesmo fornecedor aparece em centenas de lancamentos (42 MB viram
        1,4 MB como categoria). O Numero e unico por lancamento, e ai a
        categoria custa MAIS que o texto puro."""
        ns = carregar([], ["TIPOS_ECONOMICOS_FLUXO"])
        self.assertEqual(ns["TIPOS_ECONOMICOS_FLUXO"].get("Histórico"), "category")
        self.assertNotIn("Número", ns["TIPOS_ECONOMICOS_FLUXO"])

    def test_diario_guarda_as_repetitivas_como_categoria(self):
        i = FONTE.index("def carregar_diario(")
        corpo = FONTE[i:FONTE.index("\ndef ", i + 10)]
        self.assertIn('astype("category")', corpo)
        for coluna in ("Plano de Contas", "Centro de Custos", "Linha DRE"):
            self.assertIn(f'"{coluna}"', corpo)

    def test_caches_pesados_guardam_poucas_copias(self):
        """Cada entrada guarda uma copia inteira dos dados de 13 lojas. Oito
        entradas eram oito copias."""
        import re as _re
        for funcao in ("carregar_dados_abas", "carregar_dados_por_loja"):
            i = FONTE.index(f"def {funcao}(")
            decorador = FONTE[max(0, i - 400):i]
            achado = _re.search(r"max_entries=(\d+)", decorador)
            self.assertIsNotNone(achado, f"{funcao} sem teto")
            self.assertLessEqual(int(achado.group(1)), 3, f"{funcao} guarda copias demais")

    def test_guarda_de_memoria_limpa_antes_do_corte(self):
        """850 MB era tarde: o servidor corta perto de 1 GB, manda e-mail e
        bloqueia o acesso. Reler uma planilha custa segundos."""
        ns = carregar([], ["LIMITE_MEMORIA_ALERTA_MB", "LIMITE_MEMORIA_LIMPEZA_MB"])
        self.assertLessEqual(ns["LIMITE_MEMORIA_LIMPEZA_MB"], 700)
        self.assertLess(ns["LIMITE_MEMORIA_ALERTA_MB"], ns["LIMITE_MEMORIA_LIMPEZA_MB"])

    def test_guarda_de_memoria_existe_e_tem_limite(self):
        ns = carregar(["memoria_em_uso_mb"],
                      ["LIMITE_MEMORIA_ALERTA_MB", "LIMITE_MEMORIA_LIMPEZA_MB"])
        self.assertLess(ns["LIMITE_MEMORIA_ALERTA_MB"], ns["LIMITE_MEMORIA_LIMPEZA_MB"],
                        "o alerta tem de vir antes da limpeza")
        # A medicao nao pode derrubar nada onde nao houver /proc.
        valor = ns["memoria_em_uso_mb"]()
        self.assertTrue(valor is None or valor > 0)

    def test_guarda_so_limpa_acima_do_limite(self):
        ns = carregar(["guardar_memoria"], ["LIMITE_MEMORIA_LIMPEZA_MB"])
        limpezas = []
        ns["st"] = type("st", (), {"cache_data": type("c", (), {
            "clear": staticmethod(lambda: limpezas.append(1))})()})()
        ns["gc"] = type("gc", (), {"collect": staticmethod(lambda: None)})()

        ns["memoria_em_uso_mb"] = lambda: 100.0
        mb, limpou = ns["guardar_memoria"]()
        self.assertFalse(limpou, "limpou com a memoria baixa")

        ns["memoria_em_uso_mb"] = lambda: ns["LIMITE_MEMORIA_LIMPEZA_MB"] + 50
        mb, limpou = ns["guardar_memoria"]()
        self.assertTrue(limpou, "nao limpou mesmo acima do limite")
        self.assertEqual(len(limpezas), 1)

        # Sem medicao possivel, nao age -- e nao quebra.
        ns["memoria_em_uso_mb"] = lambda: None
        self.assertEqual(ns["guardar_memoria"](), (None, False))


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
                       "preparar_fluxo_caixa"):
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
        # Funcoes escritas por nos, cujos parametros sao em portugues.
        FUNCOES_EM_PORTUGUES = {"html_embutido", "tabela_selecionavel"}

        def nome(no):
            f = no.func
            return f.id if isinstance(f, ast.Name) else (f.attr if isinstance(f, ast.Attribute) else "")

        errados = []
        for no in ast.walk(arvore):
            if not isinstance(no, ast.Call):
                continue
            args = [kw.arg for kw in no.keywords]
            if nome(no) in FUNCOES_EM_PORTUGUES:
                # Funcoes nossas: aqui o portugues e o certo, e o ingles e
                # que seria o erro.
                if "height" in args or "width" in args:
                    errados.append(f"linha {no.lineno}: {nome(no)} com height/width")
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