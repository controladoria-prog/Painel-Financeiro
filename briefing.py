# -*- coding: utf-8 -*-
"""Briefing executivo das 7h -- roda no GitHub Actions e envia por e-mail.

Reaproveita as funções do app.py SEM importar o Streamlit: extrai o fonte de
cada função e constante por AST (o mesmo truque da suíte de testes), fecha as
dependências automaticamente e executa tudo num espaço onde `st` é um objeto
de mentira que aceita qualquer chamada e não faz nada. Resultado: a tela e o
e-mail contam a MESMA história, porque saem das MESMAS funções -- nada foi
duplicado, e quando o app mudar a regra, o e-mail muda junto.

Uso:
    python briefing.py                    calcula e envia (precisa das variáveis abaixo)
    python briefing.py --teste            calcula e imprime o e-mail em texto, sem enviar
    python briefing.py --salvar b.html    grava o HTML do e-mail para conferir no navegador

Variáveis de ambiente (no GitHub ficam em Settings > Secrets and variables > Actions):
    SMTP_USUARIO        e-mail que envia (Gmail)
    SMTP_SENHA          SENHA DE APP do Gmail (não é a senha da conta)
    EMAIL_DESTINO       quem recebe; vários separados por vírgula
    FECHAMENTO_CSV_URL  link da planilha de Fechamento Mensal (o mesmo do app; opcional)
    LINK_PAINEL         link do painel para o rodapé do e-mail (opcional)
    URL_ORCADO, URL_REALIZADO, ABA_DRE, SMTP_SERVIDOR, SMTP_PORTA   opcionais (têm padrão)
"""
import ast
import base64
import contextlib
import gc
import hashlib
import hmac
import io
import math
import os
import re
import smtplib
import ssl
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timedelta
from email.message import EmailMessage
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

CAMINHO_APP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")

# Os mesmos endereços do obter_caminhos_excel do app; podem ser trocados por
# variável de ambiente sem mexer no código.
URL_ORCADO_PADRAO = "https://docs.google.com/spreadsheets/d/1x68Eg_6LlSKeFJEGmfhyBfcGgheSrVsl/export?format=xlsx"
URL_REALIZADO_PADRAO = "https://docs.google.com/spreadsheets/d/12I0vGpYU_KNhGxAHOMHWAQu3Xkz_EsUZ/export?format=xlsx"

# O que o briefing precisa do app. As dependências de cada uma (funções e
# constantes que elas chamam) são descobertas sozinhas pela árvore do fonte.
SEMENTES = [
    "carregar_dados_abas", "get_valor_consolidado_multi",
    "fatos_do_ritmo", "montar_fatos_executivos", "montar_narrativa_executiva",
    "narrativa_em_texto", "projetar_margem_fechamento",
    "_deltas_pendentes_do_fechamento", "url_csv_do_fechamento",
    "carregar_planilha_fechamento", "formata_valor_curto", "_pct_br",
    "DIAS_DEFASAGEM_DADOS", "FUSO_BR",
]

NOMES_MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
               "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
DIAS_SEMANA = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
               "sexta-feira", "sábado", "domingo"]


class _StreamlitDeMentira:
    """Aceita qualquer atributo, chamada, índice ou iteração e não faz nada.

    Serve para o `st` (e o `components`, e o `go`) dentro das funções
    extraídas: um `@st.cache_data(ttl=None)` vira decorador que devolve a
    própria função; um `st.warning(...)` dentro de um loader vira nada. Só o
    que o briefing PRECISA acontece -- ler planilha e fazer conta."""

    def __getattr__(self, nome):
        return self

    def __call__(self, *args, **kwargs):
        if len(args) == 1 and not kwargs and callable(args[0]):
            return args[0]          # uso como decorador: devolve a função intacta
        return self

    def __getitem__(self, chave):
        return self

    def __contains__(self, item):
        return False

    def __bool__(self):
        return False

    def __iter__(self):
        return iter(())

    def __str__(self):
        return ""


def carregar_funcoes_do_app(caminho=CAMINHO_APP, sementes=SEMENTES):
    """Extrai do app.py as funções/constantes pedidas MAIS tudo o que elas
    usam, e executa num espaço isolado. Devolve o espaço (dict nome -> objeto).

    A busca de dependências é por NOME: qualquer identificador dentro de uma
    função que seja função ou constante de nível de módulo entra também. Pode
    trazer um pouco a mais (nome local homônimo), nunca a menos."""
    with open(caminho, encoding="utf-8") as arquivo:
        fonte = arquivo.read()
    arvore = ast.parse(fonte)
    # Entram no catálogo as FUNÇÕES e as CONSTANTES (nome em maiúsculas, com
    # ou sem _ na frente). Variável de execução do app -- list_df_real, m_map,
    # df_ref -- fica de fora de propósito: ela também é nome de parâmetro
    # das funções, e incluí-la arrastaria meio app (e o Streamlit) junto.
    eh_constante = re.compile(r"_?[A-Z][A-Z0-9_]*").fullmatch
    catalogo = {}
    for no in arvore.body:
        if isinstance(no, ast.FunctionDef):
            catalogo[no.name] = no
        elif isinstance(no, ast.Assign):
            for alvo in no.targets:
                if isinstance(alvo, ast.Name) and eh_constante(alvo.id):
                    catalogo[alvo.id] = no
        elif (isinstance(no, ast.AnnAssign) and isinstance(no.target, ast.Name)
              and eh_constante(no.target.id)):
            catalogo[no.target.id] = no
    faltam = [s for s in sementes if s not in catalogo]
    if faltam:
        raise SystemExit(f"O app.py não tem mais: {', '.join(faltam)} -- o briefing precisa acompanhar.")
    incluidos, fila = set(), list(sementes)
    while fila:
        nome = fila.pop()
        if nome in incluidos:
            continue
        incluidos.add(nome)
        for sub in ast.walk(catalogo[nome]):
            if isinstance(sub, ast.Name) and sub.id in catalogo and sub.id not in incluidos:
                fila.append(sub.id)
    nos = sorted((catalogo[n] for n in incluidos), key=lambda n: n.lineno)
    st = _StreamlitDeMentira()
    espaco = {
        "pd": pd, "np": np, "re": re, "io": io, "os": os, "math": math, "time": time,
        "base64": base64, "hashlib": hashlib, "hmac": hmac, "contextlib": contextlib,
        "datetime": datetime, "timedelta": timedelta, "ZoneInfo": ZoneInfo,
        "gc": gc, "urllib": urllib, "Workbook": Workbook, "Font": Font, "PatternFill": PatternFill,
        "Alignment": Alignment, "Border": Border, "Side": Side,
        "st": st, "components": st, "go": st,
    }
    exec("\n\n".join(ast.unparse(no) for no in nos), espaco)   # noqa: S102 -- fonte do próprio repositório
    return espaco


def montar_briefing(ns, hoje=None, aba=None, url_orc=None, url_real=None, url_fech=""):
    """Carrega as planilhas e devolve (fatos, itens_da_narrativa, contexto)."""
    hoje = hoje or datetime.now(ns["FUSO_BR"]).date()
    hoje_dados = hoje - timedelta(days=ns["DIAS_DEFASAGEM_DADOS"])
    aba = aba or os.environ.get("ABA_DRE", "DRE CONSOLIDADO")
    url_orc = url_orc or os.environ.get("URL_ORCADO", URL_ORCADO_PADRAO)
    url_real = url_real or os.environ.get("URL_REALIZADO", URL_REALIZADO_PADRAO)

    list_df_orc, list_df_real = ns["carregar_dados_abas"](url_orc, url_real, [aba])
    df_ref = next((d for d in list_df_real if d is not None and not d.empty), None)
    if df_ref is None:
        raise SystemExit(f"Não consegui ler a aba '{aba}' do Realizado.")
    meses_cols = [f"{m:02d}/{hoje.year}" for m in range(1, 13)]
    m_map = {n: c for n, c in zip(NOMES_MESES, meses_cols) if c in df_ref.columns}
    cols_kpi = [c for c in m_map.values() if int(c[:2]) <= hoje.month]
    if not cols_kpi:
        raise SystemExit("A planilha não tem as colunas de mês deste ano.")
    col_nome = "Nome" if "Nome" in df_ref.columns else df_ref.columns[0]
    linhas = list(df_ref[col_nome].dropna().unique().astype(str))
    gv = ns["get_valor_consolidado_multi"]

    def valor(lado, linha, cols, exato=False):
        return gv(list_df_orc if lado == "orc" else list_df_real, linha, cols,
                  exato_linha_sintetica=exato)

    ritmo = ns["fatos_do_ritmo"](list_df_real, list_df_orc, cols_kpi, meses_cols, m_map, hoje_dados)

    pendencias, margem_proj = [], None
    if url_fech:
        df_fech, erro = ns["carregar_planilha_fechamento"](ns["url_csv_do_fechamento"](url_fech))
        if not erro and df_fech is not None and not df_fech.empty:
            d_rec, d_eb, pendencias = ns["_deltas_pendentes_do_fechamento"](
                df_fech, cols_kpi, meses_cols, (hoje.year, hoje.month),
                lambda linha, col: gv(list_df_real, linha, [col]))
            margem_proj = ns["projetar_margem_fechamento"](
                valor("real", "3 - Receita Operacional Liquida", cols_kpi),
                valor("real", "11 - EBITDA", cols_kpi), d_rec, d_eb)

    col_corrente = f"{hoje.month:02d}/{hoje.year}"
    nome_corrente = next((n for n, c in m_map.items() if c == col_corrente), col_corrente)
    rotulo = f"Acumulado YTD até {nome_corrente}"
    fatos = ns["montar_fatos_executivos"](
        valor, linhas, cols_kpi, col_corrente, nome_corrente,
        pendencias=pendencias, ritmo=ritmo, margem_proj=margem_proj, rotulo_periodo=rotulo)
    itens = ns["montar_narrativa_executiva"](fatos)
    contexto = {"hoje": hoje, "hoje_dados": hoje_dados, "aba": aba, "rotulo": rotulo}
    return fatos, itens, contexto


def montar_email(itens, fatos, hoje, link_painel="", ns=None):
    """Devolve (html, texto). O HTML é claro e simples de propósito: cliente
    de e-mail no celular não respeita tema escuro nem CSS esperto."""
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    pct = (ns or {}).get("_pct_br", lambda v, casas=1: f"{v:.{casas}f}".replace(".", ","))
    tirar_tags = lambda t: re.sub(r"<[^>]+>", "", t)   # noqa: E731
    dia = f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}"
    cores = {"negativo": "#c0392b", "positivo": "#1e8449", "alerta": "#b9770e", "neutro": "#7f8c8d"}

    kpis = []
    if fatos.get("rec_real") is not None:
        kpis.append(("Receita líquida YTD", fmt(fatos["rec_real"]),
                     f"orçado {fmt(fatos['rec_orc'])}" if fatos.get("rec_orc") else ""))
    if fatos.get("ebitda_real") is not None:
        kpis.append(("EBITDA YTD", fmt(fatos["ebitda_real"]),
                     f"orçado {fmt(fatos['ebitda_orc'])}" if fatos.get("ebitda_orc") else ""))
    if fatos.get("rec_real"):
        margem = fatos["ebitda_real"] / fatos["rec_real"] * 100
        sub = (f"fecha em {pct(fatos['margem_proj'])}% com os lançamentos pendentes"
               if fatos.get("margem_proj") is not None else "realizada no período")
        kpis.append(("Margem EBITDA", f"{pct(margem)}%", sub))
    r = fatos.get("ritmo")
    if r:
        sub = (f"chance de bater a meta: {r['chance'] * 100:.0f}%" if r.get("chance") is not None
               else f"dia {r['dia']} de {r['dias']}")
        kpis.append((f"Ritmo de {str(r['mes']).capitalize()}", f"{r['pct']:.0f}%", sub))

    celulas_kpi = "".join(
        '<td style="padding:10px 12px; border:1px solid #e3e6ea; border-radius:6px; vertical-align:top; width:25%;">'
        f'<div style="font-size:10px; letter-spacing:0.6px; color:#7f8c8d; text-transform:uppercase;">{rotulo}</div>'
        f'<div style="font-size:20px; font-weight:700; color:#2c3e50; margin:2px 0;">{valor}</div>'
        f'<div style="font-size:11px; color:#7f8c8d;">{sub}</div></td>'
        for rotulo, valor, sub in kpis)
    linhas_narrativa = "".join(
        '<tr><td style="padding:9px 10px 9px 0; border-top:1px solid #e3e6ea; vertical-align:top; width:110px;'
        f' font-size:10px; letter-spacing:0.8px; text-transform:uppercase; color:{cores.get(i["tom"], "#7f8c8d")};">'
        f'{i["rotulo"]}</td>'
        f'<td style="padding:9px 0; border-top:1px solid #e3e6ea; font-size:14px; line-height:1.5; color:#2c3e50;">'
        f'{i["texto"]}</td></tr>'
        for i in itens)
    rodape_link = (f'<a href="{link_painel}" style="color:#2874a6;">Abrir o painel</a> · '
                   if link_painel else "")
    html = (
        '<div style="font-family:Segoe UI, Arial, sans-serif; max-width:760px; margin:0 auto; color:#2c3e50;">'
        '<div style="font-size:11px; letter-spacing:1px; text-transform:uppercase; color:#7f8c8d;">'
        'Controladoria B&amp;A · briefing executivo</div>'
        f'<h2 style="margin:4px 0 14px 0; font-size:20px;">{fatos.get("periodo", "")} · {dia}</h2>'
        f'<table cellspacing="6" cellpadding="0" style="width:100%; border-collapse:separate;"><tr>{celulas_kpi}</tr></table>'
        '<div style="font-size:11px; letter-spacing:0.8px; text-transform:uppercase; color:#7f8c8d; margin:18px 0 4px 0;">'
        'O que aconteceu e por quê</div>'
        f'<table cellspacing="0" cellpadding="0" style="width:100%;">{linhas_narrativa}</table>'
        f'<div style="font-size:11px; color:#7f8c8d; margin-top:18px;">{rodape_link}'
        'Gerado automaticamente pelo painel a partir das planilhas de Orçado, Realizado e Fechamento · '
        'lançamentos chegam D+2 · a chance de bater a meta é uma estimativa a partir do histórico do ano.</div>'
        '</div>')
    texto = (f"Controladoria B&A · briefing executivo\n{fatos.get('periodo', '')} · {dia}\n\n"
             + "\n".join(f"{rotulo}: {tirar_tags(valor)} ({sub})" if sub else f"{rotulo}: {tirar_tags(valor)}"
                         for rotulo, valor, sub in kpis)
             + "\n\nO que aconteceu e por quê\n"
             + (ns or {}).get("narrativa_em_texto", lambda it: "\n".join(
                 f"{i['rotulo']}: {tirar_tags(i['texto'])}" for i in it))(itens)
             + (f"\n\nPainel: {link_painel}" if link_painel else "")
             + "\n\nGerado automaticamente pelo painel. Lançamentos chegam D+2.")
    return html, texto


def enviar_email(assunto, html, texto):
    usuario = os.environ["SMTP_USUARIO"]
    senha = os.environ["SMTP_SENHA"]
    destinos = [d.strip() for d in os.environ["EMAIL_DESTINO"].split(",") if d.strip()]
    if not destinos:
        raise SystemExit("EMAIL_DESTINO está vazio.")
    servidor = os.environ.get("SMTP_SERVIDOR", "smtp.gmail.com")
    porta = int(os.environ.get("SMTP_PORTA", "465"))
    mensagem = EmailMessage()
    mensagem["Subject"] = assunto
    mensagem["From"] = usuario
    mensagem["To"] = ", ".join(destinos)
    mensagem.set_content(texto)
    mensagem.add_alternative(html, subtype="html")
    with smtplib.SMTP_SSL(servidor, porta, context=ssl.create_default_context()) as conexao:
        conexao.login(usuario, senha)
        conexao.send_message(mensagem)
    return destinos


def main(argv):
    ns = carregar_funcoes_do_app()
    fatos, itens, ctx = montar_briefing(ns, url_fech=os.environ.get("FECHAMENTO_CSV_URL", ""))
    if not itens:
        raise SystemExit("A narrativa saiu vazia -- confira as planilhas.")
    html, texto = montar_email(itens, fatos, ctx["hoje"], os.environ.get("LINK_PAINEL", ""), ns)
    assunto = f"Briefing executivo · {ctx['rotulo']} · {ctx['hoje'].strftime('%d/%m')}"
    if "--salvar" in argv:
        caminho = argv[argv.index("--salvar") + 1]
        with open(caminho, "w", encoding="utf-8") as arquivo:
            arquivo.write(html)
        print(f"HTML gravado em {caminho}")
    if "--teste" in argv:
        print(assunto)
        print(texto)
        return
    destinos = enviar_email(assunto, html, texto)
    print(f"Briefing enviado para {', '.join(destinos)}: {assunto}")


if __name__ == "__main__":
    main(sys.argv[1:])