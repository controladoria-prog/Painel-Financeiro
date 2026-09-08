# -*- coding: utf-8 -*-
"""Board pack -- a apresentação do fechamento mensal, pronta no dia 5.

Gera um PPTX de seis slides (capa, semáforo e cartões, receita por mês,
EBITDA por mês, estouros e folgas, narrativa) com os MESMOS números e a
MESMA narrativa da tela e do e-mail, e manda por e-mail como anexo. A
reunião de resultados começa com o material pronto -- e igual todo mês.

Uso:
    python board_pack.py                    gera e envia
    python board_pack.py --teste            gera e grava board_pack.pptx, sem enviar
    python board_pack.py --salvar arq.pptx  grava com outro nome

Precisa de: pip install python-pptx (além de pandas, numpy, openpyxl).
"""
import io
import os
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

from briefing import (DIAS_SEMANA, carregar_funcoes_do_app, enviar_email, montar_briefing,
                      responsaveis_por_conta, series_mensais, status_geral)

NAVY = RGBColor(0x1B, 0x2A, 0x41)
CINZA = RGBColor(0x6B, 0x72, 0x80)
TEXTO = RGBColor(0x1F, 0x29, 0x37)
VERDE = RGBColor(0x1E, 0x84, 0x49)
VERMELHO = RGBColor(0xC0, 0x39, 0x2B)
AMBAR = RGBColor(0xB9, 0x77, 0x0E)
AZUL = RGBColor(0x2F, 0x6F, 0xD6)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
TONS = {"negativo": VERMELHO, "positivo": VERDE, "alerta": AMBAR, "neutro": CINZA}


def _texto(slide, x, y, w, h, texto, tamanho=14, cor=TEXTO, negrito=False, alinhamento=None):
    caixa = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    quadro = caixa.text_frame
    quadro.word_wrap = True
    paragrafo = quadro.paragraphs[0]
    paragrafo.text = str(texto)
    paragrafo.font.size = Pt(tamanho)
    paragrafo.font.bold = negrito
    paragrafo.font.color.rgb = cor
    if alinhamento is not None:
        paragrafo.alignment = alinhamento
    return caixa


def _retangulo(slide, x, y, w, h, cor):
    from pptx.enum.shapes import MSO_SHAPE
    forma = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    forma.fill.solid()
    forma.fill.fore_color.rgb = cor
    forma.line.fill.background()
    return forma


def _cabecalho(slide, titulo, subtitulo=""):
    _retangulo(slide, 0, 0, 13.333, 0.9, NAVY)
    _texto(slide, 0.5, 0.18, 9, 0.5, titulo, 22, BRANCO, True)
    if subtitulo:
        _texto(slide, 9.3, 0.3, 3.6, 0.4, subtitulo, 11, RGBColor(0x9F, 0xB3, 0xD1))


def _grafico_colunas(slide, titulo, rotulos, real, orc, y):
    dados = CategoryChartData()
    dados.categories = rotulos
    dados.add_series("Realizado", [v / 1e6 for v in real])
    dados.add_series("Orçado", [v / 1e6 for v in orc])
    grafico = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(y),
                                     Inches(12.3), Inches(5.3), dados).chart
    grafico.has_legend = True
    grafico.legend.position = XL_LEGEND_POSITION.BOTTOM
    grafico.legend.include_in_layout = False
    grafico.value_axis.has_major_gridlines = False
    grafico.value_axis.minimum_scale = 0   # eixo do zero: recorte inflava a diferença
    grafico.value_axis.tick_labels.font.size = Pt(10)
    grafico.category_axis.tick_labels.font.size = Pt(11)
    grafico.plots[0].series[0].format.fill.solid()
    grafico.plots[0].series[0].format.fill.fore_color.rgb = AZUL
    grafico.plots[0].series[1].format.fill.solid()
    grafico.plots[0].series[1].format.fill.fore_color.rgb = RGBColor(0xC9, 0xD1, 0xDB)
    grafico.plots[0].has_data_labels = True
    grafico.plots[0].data_labels.number_format = '0.0"M"'
    grafico.plots[0].data_labels.font.size = Pt(9)
    return grafico


def causas_do_gap(fatos, modo, ns=None):
    """Organiza o que os dados sabem sobre o gap em categorias (as espinhas do
    Ishikawa). Cada item traz o valor; nada aqui é opinião."""
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    pct = (ns or {}).get("_pct_br", lambda v, casas=1: f"{v:.{casas}f}".replace(".", ","))
    cat = {}
    if modo == "departamento":
        efeito = (f"Gasto {fmt(fatos.get('gasto_real', 0))} contra {fmt(fatos.get('gasto_orc', 0))} orçados "
                  f"({fatos.get('departamento', '')})")
        cat["Contas acima do orçado"] = [f"{e['conta']}: +{fmt(e['desvio'])}" + (f" (+{e['pct']:.0f}%)" if e.get("pct") is not None else " (sem orçamento)")
                                        for e in (fatos.get("estouros") or [])[:4]]
        cat["Contas com folga"] = [f"{c['conta']}: −{fmt(c['folga'])}" for c in (fatos.get("folgas") or [])[:3]]
        cat["Cobertura"] = [f"{fatos.get('n_acima', 0)} de {fatos.get('n_contas', 0)} contas acima do orçado"]
        return efeito, {k: v for k, v in cat.items() if v}
    fech = fatos.get("fechado") or {}
    eb_r, eb_o = fech.get("ebitda_real", fatos.get("ebitda_real", 0)), fech.get("ebitda_orc", fatos.get("ebitda_orc", 0))
    rec_r, rec_o = fech.get("rec_real", fatos.get("rec_real", 0)), fech.get("rec_orc", fatos.get("rec_orc", 0))
    gap = (eb_r or 0) - (eb_o or 0)
    efeito = f"EBITDA {fmt(eb_r or 0)} contra {fmt(eb_o or 0)} orçados: {'+' if gap >= 0 else '−'}{fmt(abs(gap))}"
    if rec_o and eb_o:
        margem_orc = eb_o / rec_o
        efeito_rec = ((rec_r or 0) - rec_o) * margem_orc
        efeito_custo = gap - efeito_rec
        cat["Receita"] = [f"Receita {fmt(rec_r or 0)} vs {fmt(rec_o)} orçados ({pct(abs((rec_r or 0) / rec_o - 1) * 100)}% "
                          f"{'abaixo' if (rec_r or 0) < rec_o else 'acima'})",
                          f"Efeito no EBITDA, na margem orçada de {pct(margem_orc * 100)}%: {'+' if efeito_rec >= 0 else '−'}{fmt(abs(efeito_rec))}"]
        cat["Custos e despesas"] = [f"Efeito total no EBITDA: {'+' if efeito_custo >= 0 else '−'}{fmt(abs(efeito_custo))} "
                                   f"({'acima' if efeito_custo < 0 else 'abaixo'} do orçado)"]
    var, oper = [], []
    for e in (fatos.get("estouros") or [])[:6]:
        txt = f"{e['conta']}: +{fmt(e['desvio'])}" + (f" (+{e['pct']:.0f}%)" if e.get("pct") is not None else " (sem orçamento)")
        (var if str(e.get("linha", "")).startswith("6") else oper).append(txt)
    if var:
        cat["Despesas variáveis"] = var[:3]
    if oper:
        cat["Despesas operacionais"] = oper[:3]
    lg = fatos.get("lojas_gap")
    if lg and lg.get("lojas"):
        cat["Lojas"] = [f"{l['loja']}: −{fmt(abs(l['desvio']))}" for l in lg["lojas"][:3]] + [
            f"{lg['n']} de {lg['total_lojas']} lojas explicam {lg['fracao'] * 100:.0f}% do gap"]
    if fatos.get("artefatos"):
        cat["Fechamento pendente"] = [f"{a['conta']}: {fmt(a['pendente'])} ainda por lançar (não é economia)"
                                     for a in fatos["artefatos"][:2]]
    return efeito, cat


def _slide_ishikawa(apresentacao, em_branco, efeito, categorias, dia):
    from pptx.enum.shapes import MSO_CONNECTOR
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "Ishikawa · causas do gap organizadas pelos dados", dia)
    # espinha e efeito
    espinha = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.8), Inches(4.05), Inches(10.6), Inches(4.05))
    espinha.line.color.rgb = NAVY
    espinha.line.width = Pt(3)
    _retangulo(slide, 10.6, 3.35, 2.5, 1.4, NAVY)
    _texto(slide, 10.7, 3.45, 2.3, 1.2, efeito, 11, BRANCO, True)
    nomes = list(categorias.keys())[:6]
    posicoes_x = [1.6, 4.6, 7.6]
    for i, nome in enumerate(nomes):
        em_cima = i < 3
        x = posicoes_x[i % 3]
        y_topo, y_base = (1.15, 4.05) if em_cima else (4.05, 6.95)
        osso = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y_topo if em_cima else y_base),
                                          Inches(x + 1.2), Inches(y_base if em_cima else y_topo))
        osso.line.color.rgb = CINZA
        osso.line.width = Pt(1.5)
        y_rot = 1.0 if em_cima else 6.55
        _texto(slide, x - 0.3, y_rot, 2.9, 0.35, nome.upper(), 10, NAVY, True)
        y_itens = (1.4 if em_cima else 4.45)
        for j, item in enumerate(categorias[nome][:4]):
            _texto(slide, x + 0.15, y_itens + j * 0.5, 2.9, 0.5, "• " + item, 9, TEXTO)
    _texto(slide, 0.5, 7.05, 12.3, 0.35,
           "As causas são as que os números mostram (quanto e onde). O motivo de negócio de cada uma é de quem lança a conta.",
           9, CINZA)


def _slide_5w2h(apresentacao, em_branco, fatos, modo, hoje, dia, ns=None):
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "5W2H · plano a partir do que os dados apontam", dia)
    responsaveis = fatos.get("responsaveis") or {}
    onde = fatos.get("departamento") or "Consolidado"
    quando = f"meses fechados · visto em {hoje.strftime('%d/%m')}"
    linhas = [("O quê", "Por quê (dado)", "Onde", "Quando", "Quem", "Como", "Quanto")]
    for e in (fatos.get("estouros") or [])[:4]:
        porque = (f"+{e['pct']:.0f}% sobre o orçado" if e.get("pct") is not None else "conta sem orçamento")
        quem = fatos.get("departamento") or responsaveis.get(e["conta"], "a definir")
        linhas.append((f"{e['conta']} passou do orçado", porque, onde, quando, quem,
                       "Abrir os lançamentos mês a mês contra o orçado e decidir: renegociar, cortar ou reorçar",
                       f"+{fmt(e['desvio'])}"))
    for a in (fatos.get("artefatos") or [])[:1]:
        linhas.append((f"{a['conta']} com lançamento pendente", "fechamento ainda não entrou", onde, quando,
                       fatos.get("departamento") or responsaveis.get(a["conta"], "Fechamento"),
                       "Concluir o lançamento no fechamento para a margem real aparecer", fmt(a["pendente"])))
    lg = fatos.get("lojas_gap")
    if lg and lg.get("lojas") and modo != "departamento":
        pior = lg["lojas"][0]
        linhas.append((f"{pior['loja']} concentra o gap de EBITDA", f"−{fmt(abs(pior['desvio']))} vs orçado",
                       pior["loja"], quando, "Coordenação da loja",
                       "Rever receita e despesas da unidade linha a linha", f"−{fmt(abs(pior['desvio']))}"))
    if len(linhas) == 1:
        linhas.append(("Nada fora do orçado nos dados", "—", onde, quando, "—", "Manter acompanhamento", "—"))
    tabela = slide.shapes.add_table(len(linhas), 7, Inches(0.4), Inches(1.15), Inches(12.5),
                                    Inches(0.55 * len(linhas))).table
    larguras = (2.4, 1.6, 1.4, 1.5, 1.6, 2.8, 1.2)
    for k, w in enumerate(larguras):
        tabela.columns[k].width = Inches(w)
    for i, linha in enumerate(linhas):
        for j, valor in enumerate(linha):
            celula = tabela.cell(i, j)
            celula.text = str(valor)
            paragrafo = celula.text_frame.paragraphs[0]
            paragrafo.font.size = Pt(10 if i else 10)
            paragrafo.font.bold = i == 0
            paragrafo.font.color.rgb = BRANCO if i == 0 else TEXTO
            celula.fill.solid()
            celula.fill.fore_color.rgb = NAVY if i == 0 else (RGBColor(0xF4, 0xF7, 0xFB) if i % 2 else BRANCO)
    _texto(slide, 0.5, 6.9, 12.3, 0.5,
           "'Quem' vem do mapa de departamentos do painel; 'Como' é o encaminhamento padrão -- a decisão é da reunião.",
           9, CINZA)


def _slide_5_porques(apresentacao, em_branco, fatos, modo, dia, ns=None):
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    pct = (ns or {}).get("_pct_br", lambda v, casas=1: f"{v:.{casas}f}".replace(".", ","))
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "5 Porquês · até onde os dados respondem", dia)
    passos = []
    if modo == "departamento":
        g_r, g_o = fatos.get("gasto_real", 0), fatos.get("gasto_orc", 0)
        passos.append(("Por que o gasto do departamento ficou fora do orçado?",
                       f"{fmt(g_r)} realizados contra {fmt(g_o)} orçados ({'+' if g_r >= g_o else '−'}{fmt(abs(g_r - g_o))})."))
        est = fatos.get("estouros") or []
        if est:
            passos.append(("Por que passou?", "As contas que mais pesaram: " + "; ".join(
                f"{e['conta']} +{fmt(e['desvio'])}" for e in est[:3]) + "."))
            e0 = est[0]
            passos.append((f"Por que {e0['conta']} passou?",
                           (f"Ficou {e0['pct']:.0f}% acima do orçado nos meses fechados." if e0.get("pct") is not None
                            else "Não tinha orçamento previsto -- gasto novo ou não orçado.")))
    else:
        fech = fatos.get("fechado") or {}
        eb_r, eb_o = fech.get("ebitda_real", 0), fech.get("ebitda_orc", 0)
        rec_r, rec_o = fech.get("rec_real", 0), fech.get("rec_orc", 0)
        gap = (eb_r or 0) - (eb_o or 0)
        passos.append(("Por que o EBITDA ficou fora do orçado?",
                       f"{fmt(eb_r or 0)} contra {fmt(eb_o or 0)} orçados: {'+' if gap >= 0 else '−'}{fmt(abs(gap))} nos meses fechados."))
        if rec_o and eb_o:
            margem_orc = eb_o / rec_o
            ef_rec = ((rec_r or 0) - rec_o) * margem_orc
            ef_custo = gap - ef_rec
            passos.append(("Por que? (de onde vem o gap)",
                           f"A receita {'menor' if ef_rec < 0 else 'maior'} explica {fmt(abs(ef_rec))} na margem orçada de "
                           f"{pct(margem_orc * 100)}%; custos e despesas {'acima' if ef_custo < 0 else 'abaixo'} do orçado, {fmt(abs(ef_custo))}."))
        lg = fatos.get("lojas_gap")
        if lg and lg.get("lojas"):
            passos.append(("Por que a receita/EBITDA ficou abaixo? (onde)",
                           f"{lg['n']} de {lg['total_lojas']} lojas explicam {lg['fracao'] * 100:.0f}%: " + ", ".join(
                               f"{l['loja']} −{fmt(abs(l['desvio']))}" for l in lg["lojas"][:3]) + "."))
        est = fatos.get("estouros") or []
        if est:
            passos.append(("Por que custos e despesas pesaram? (quais contas)", "; ".join(
                f"{e['conta']} +{fmt(e['desvio'])}" + (f" (+{e['pct']:.0f}%)" if e.get("pct") is not None else " (sem orçamento)")
                for e in est[:3]) + "."))
    responsaveis = fatos.get("responsaveis") or {}
    donos = sorted({responsaveis.get(e["conta"]) for e in (fatos.get("estouros") or []) if responsaveis.get(e["conta"])})
    passos.append(("Por que essas contas se comportaram assim?",
                   "Aqui os dados param: eles dizem quanto e onde, não o motivo de negócio. "
                   + (f"Quem responde: {', '.join(donos)}." if donos else "Trazer os responsáveis para a reunião.")))
    y = 1.2
    for n, (pergunta, resposta) in enumerate(passos[:5], start=1):
        _retangulo(slide, 0.5, y, 0.5, 0.95, NAVY if n < len(passos[:5]) else AMBAR)
        _texto(slide, 0.5, y + 0.25, 0.5, 0.5, str(n), 16, BRANCO, True)
        _texto(slide, 1.15, y - 0.02, 11.6, 0.4, pergunta, 12, NAVY, True)
        _texto(slide, 1.15, y + 0.33, 11.6, 0.65, resposta, 11, TEXTO)
        y += 1.1


def montar_board_pack(fatos, itens, series, hoje, ns=None, modo="consolidado"):
    """Devolve os bytes do PPTX. `modo="departamento"` troca receita/EBITDA
    por gasto realizado x orçado do departamento (fatos_do_departamento)."""
    if modo == "departamento":
        return _board_pack_departamento(fatos, itens, series, hoje, ns)
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    pct = (ns or {}).get("_pct_br", lambda v, casas=1: f"{v:.{casas}f}".replace(".", ","))
    apresentacao = Presentation()
    apresentacao.slide_width = Inches(13.333)
    apresentacao.slide_height = Inches(7.5)
    em_branco = apresentacao.slide_layouts[6]
    dia = f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}"
    status, _cor_status, frase_status = status_geral(fatos)

    # 1) capa
    slide = apresentacao.slides.add_slide(em_branco)
    _retangulo(slide, 0, 0, 13.333, 7.5, NAVY)
    _texto(slide, 0.8, 2.3, 11, 0.5, "CONTROLADORIA B&A · FECHAMENTO", 13, RGBColor(0x9F, 0xB3, 0xD1))
    _texto(slide, 0.8, 2.8, 11, 1.2, fatos.get("periodo", ""), 40, BRANCO, True)
    _texto(slide, 0.8, 4.1, 11, 0.6, f"{frase_status} · {dia}" if frase_status else dia, 16,
           RGBColor(0x9F, 0xB3, 0xD1))

    # 2) semáforo e cartões
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "Resultado do período", dia)
    cor_status = TONS.get({"NO ORÇADO": "positivo", "OBSERVAR": "alerta"}.get(status, "negativo"))
    _retangulo(slide, 0.5, 1.2, 12.3, 0.6, cor_status)
    _texto(slide, 0.7, 1.3, 12, 0.4, f"{status} · {frase_status}", 14, BRANCO, True)
    cartoes = []
    if fatos.get("rec_real") is not None:
        var = (fatos["rec_real"] / fatos["rec_orc"] - 1) * 100 if fatos.get("rec_orc") else None
        cartoes.append(("Receita líquida", fmt(fatos["rec_real"]),
                        f"{'▲' if var >= 0 else '▼'} {pct(abs(var))}% vs orçado" if var is not None else ""))
    if fatos.get("ebitda_real") is not None:
        var = (fatos["ebitda_real"] / fatos["ebitda_orc"] - 1) * 100 if fatos.get("ebitda_orc") else None
        cartoes.append(("EBITDA", fmt(fatos["ebitda_real"]),
                        f"{'▲' if var >= 0 else '▼'} {pct(abs(var))}% vs orçado" if var is not None else ""))
    if fatos.get("rec_real"):
        margem = fatos["ebitda_real"] / fatos["rec_real"] * 100
        sub = (f"fecha em {pct(fatos['margem_proj'])}% com os lançamentos pendentes"
               if fatos.get("margem_proj") is not None else "realizada no período")
        cartoes.append(("Margem EBITDA", f"{pct(margem)}%", sub))
    r = fatos.get("ritmo")
    if r:
        cartoes.append((f"Ritmo de {str(r['mes']).capitalize()}", f"{r['pct']:.0f}%",
                        f"dia {r['dia']} de {r['dias']}" + (f" · chance {r['chance'] * 100:.0f}%" if r.get("chance") is not None else "")))
    largura = 12.3 / max(len(cartoes), 1)
    for i, (rotulo, valor, sub) in enumerate(cartoes):
        x = 0.5 + i * largura
        _retangulo(slide, x + 0.05, 2.2, largura - 0.1, 0.06, NAVY)
        _texto(slide, x + 0.15, 2.35, largura - 0.3, 0.4, rotulo.upper(), 10, CINZA)
        _texto(slide, x + 0.15, 2.7, largura - 0.3, 0.8, valor, 30, TEXTO, True)
        _texto(slide, x + 0.15, 3.5, largura - 0.3, 0.8, sub, 11, CINZA)
    _texto(slide, 0.5, 4.8, 12.3, 1.2,
           "Base da análise: meses fechados; o mês corrente entra só como explicação do gap. "
           "Lançamentos chegam D+2; a chance de bater a meta é uma estimativa a partir do histórico do ano.",
           11, CINZA)

    # 3) receita por mês
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "Receita líquida por mês · realizado vs. orçado (R$ milhões)", dia)
    _grafico_colunas(slide, "Receita", series["rotulos"], series["rec_real"], series["rec_orc"], 1.2)

    # 4) EBITDA por mês
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "EBITDA por mês · realizado vs. orçado (R$ milhões)", dia)
    _grafico_colunas(slide, "EBITDA", series["rotulos"], series["eb_real"], series["eb_orc"], 1.2)

    # 5) estouros e folgas
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "Onde passou e onde sobrou (meses fechados)", dia)
    linhas_tab = [("Conta", "Situação", "Valor", "%")]
    for e in fatos.get("estouros") or []:
        linhas_tab.append((e["conta"], "acima do orçado", f"+{fmt(e['desvio'])}",
                           f"+{e['pct']:.0f}%" if e.get("pct") is not None else "s/ orç."))
    for c in fatos.get("folgas") or []:
        linhas_tab.append((c["conta"], "folga real", f"−{fmt(c['folga'])}",
                           f"−{c['pct']:.0f}%" if c.get("pct") is not None else ""))
    for a in fatos.get("artefatos") or []:
        linhas_tab.append((a["conta"], "lançamento pendente (não é economia)", f"−{fmt(a['folga'])}",
                           f"{fmt(a['pendente'])} por entrar"))
    tabela = slide.shapes.add_table(len(linhas_tab), 4, Inches(0.5), Inches(1.2), Inches(12.3),
                                    Inches(0.4 * len(linhas_tab))).table
    for i, linha in enumerate(linhas_tab):
        for j, valor in enumerate(linha):
            celula = tabela.cell(i, j)
            celula.text = str(valor)
            paragrafo = celula.text_frame.paragraphs[0]
            paragrafo.font.size = Pt(12 if i else 11)
            paragrafo.font.bold = i == 0
            paragrafo.font.color.rgb = BRANCO if i == 0 else TEXTO
            celula.fill.solid()
            celula.fill.fore_color.rgb = NAVY if i == 0 else (RGBColor(0xF4, 0xF7, 0xFB) if i % 2 else BRANCO)
    tabela.columns[0].width = Inches(5.3)
    tabela.columns[1].width = Inches(3.5)
    tabela.columns[2].width = Inches(1.8)
    tabela.columns[3].width = Inches(1.7)

    # 6-8) Ishikawa, 5W2H e 5 Porquês com os dados que existem
    efeito, categorias = causas_do_gap(fatos, "consolidado", ns)
    _slide_ishikawa(apresentacao, em_branco, efeito, categorias, dia)
    _slide_5w2h(apresentacao, em_branco, fatos, "consolidado", hoje, dia, ns)
    _slide_5_porques(apresentacao, em_branco, fatos, "consolidado", dia, ns)

    # 6) narrativa
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "O que aconteceu e por quê", dia)
    y = 1.2
    for item in itens:
        texto_limpo = item["texto"].replace("<b>", "").replace("</b>", "")
        _retangulo(slide, 0.5, y + 0.08, 0.06, 0.55, TONS.get(item.get("tom"), CINZA))
        _texto(slide, 0.7, y, 1.6, 0.4, item["rotulo"].upper(), 10, TONS.get(item.get("tom"), CINZA), True)
        _texto(slide, 2.3, y - 0.02, 10.5, 0.95, texto_limpo, 13, TEXTO)
        y += 0.95
    saida = io.BytesIO()
    apresentacao.save(saida)
    return saida.getvalue()


def _board_pack_departamento(fatos, itens, series, hoje, ns=None):
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    pct = (ns or {}).get("_pct_br", lambda v, casas=1: f"{v:.{casas}f}".replace(".", ","))
    apresentacao = Presentation()
    apresentacao.slide_width = Inches(13.333)
    apresentacao.slide_height = Inches(7.5)
    em_branco = apresentacao.slide_layouts[6]
    dia = f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}"
    gasto_real, gasto_orc = fatos.get("gasto_real", 0.0), fatos.get("gasto_orc", 0.0)
    gap = (gasto_real / gasto_orc - 1) if gasto_orc else 0.0
    status = "NO ORÇADO" if gap <= 0 else ("OBSERVAR" if gap <= 0.05 else "ATENÇÃO")
    cor_status = {"NO ORÇADO": VERDE, "OBSERVAR": AMBAR}.get(status, VERMELHO)
    frase = f"Gasto {pct(abs(gap) * 100)}% {'acima' if gap > 0 else 'abaixo'} do orçado" if gasto_orc else ""
    nome = str(fatos.get("departamento", ""))

    slide = apresentacao.slides.add_slide(em_branco)
    _retangulo(slide, 0, 0, 13.333, 7.5, NAVY)
    _texto(slide, 0.8, 2.3, 11, 0.5, "CONTROLADORIA B&A · RELATÓRIO DE CUSTOS", 13, RGBColor(0x9F, 0xB3, 0xD1))
    _texto(slide, 0.8, 2.8, 11.5, 1.2, nome, 34, BRANCO, True)
    _texto(slide, 0.8, 4.1, 11, 0.6, f"{fatos.get('periodo', '')} · {frase} · {dia}", 15, RGBColor(0x9F, 0xB3, 0xD1))

    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, f"Gasto do departamento · {fatos.get('periodo', '')}", dia)
    _retangulo(slide, 0.5, 1.2, 12.3, 0.6, cor_status)
    _texto(slide, 0.7, 1.3, 12, 0.4, f"{status} · {frase}", 14, BRANCO, True)
    cartoes = [("Gasto realizado", fmt(gasto_real), "meses fechados"),
               ("Orçado", fmt(gasto_orc), "mesmo período"),
               ("Desvio", f"{'+' if gasto_real >= gasto_orc else '−'}{fmt(abs(gasto_real - gasto_orc))}",
                f"{pct(abs(gap) * 100)}% {'acima' if gap > 0 else 'abaixo'}" if gasto_orc else ""),
               ("Contas acima do orçado", f"{fatos.get('n_acima', 0)} de {fatos.get('n_contas', 0)}", "no escopo do departamento")]
    largura = 12.3 / 4
    for i, (rotulo, valor, sub) in enumerate(cartoes):
        x = 0.5 + i * largura
        _retangulo(slide, x + 0.05, 2.2, largura - 0.1, 0.06, NAVY)
        _texto(slide, x + 0.15, 2.35, largura - 0.3, 0.4, rotulo.upper(), 10, CINZA)
        _texto(slide, x + 0.15, 2.7, largura - 0.3, 0.8, valor, 28, TEXTO, True)
        _texto(slide, x + 0.15, 3.5, largura - 0.3, 0.8, sub, 11, CINZA)
    _texto(slide, 0.5, 4.8, 12.3, 1.0,
           "Base: meses fechados; o mês corrente fica fora. Lançamentos chegam D+2. "
           "Folga de conta pode ser lançamento de fechamento ainda por entrar.", 11, CINZA)

    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "Gasto por mês · realizado vs. orçado (R$ milhões)", dia)
    _grafico_colunas(slide, "Gasto", series["rotulos"], series["gasto_real"], series["gasto_orc"], 1.2)

    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "Contas: onde passou e onde sobrou", dia)
    linhas_tab = [("Conta", "Situação", "Valor", "%")]
    for e in fatos.get("estouros") or []:
        linhas_tab.append((e["conta"], "acima do orçado", f"+{fmt(e['desvio'])}",
                           f"+{e['pct']:.0f}%" if e.get("pct") is not None else "s/ orç."))
    for c in fatos.get("folgas") or []:
        linhas_tab.append((c["conta"], "abaixo do orçado", f"−{fmt(c['folga'])}",
                           f"−{c['pct']:.0f}%" if c.get("pct") is not None else ""))
    tabela = slide.shapes.add_table(len(linhas_tab), 4, Inches(0.5), Inches(1.2), Inches(12.3),
                                    Inches(0.4 * len(linhas_tab))).table
    for i, linha in enumerate(linhas_tab):
        for j, valor in enumerate(linha):
            celula = tabela.cell(i, j)
            celula.text = str(valor)
            paragrafo = celula.text_frame.paragraphs[0]
            paragrafo.font.size = Pt(12 if i else 11)
            paragrafo.font.bold = i == 0
            paragrafo.font.color.rgb = BRANCO if i == 0 else TEXTO
            celula.fill.solid()
            celula.fill.fore_color.rgb = NAVY if i == 0 else (RGBColor(0xF4, 0xF7, 0xFB) if i % 2 else BRANCO)
    for k, w in enumerate((5.3, 3.5, 1.8, 1.7)):
        tabela.columns[k].width = Inches(w)

    efeito, categorias = causas_do_gap(fatos, "departamento", ns)
    _slide_ishikawa(apresentacao, em_branco, efeito, categorias, dia)
    _slide_5w2h(apresentacao, em_branco, fatos, "departamento", hoje, dia, ns)
    _slide_5_porques(apresentacao, em_branco, fatos, "departamento", dia, ns)
    slide = apresentacao.slides.add_slide(em_branco)
    _cabecalho(slide, "O que aconteceu e por quê", dia)
    y = 1.2
    for item in itens:
        _retangulo(slide, 0.5, y + 0.08, 0.06, 0.55, TONS.get(item.get("tom"), CINZA))
        _texto(slide, 0.7, y, 1.6, 0.4, item["rotulo"].upper(), 10, TONS.get(item.get("tom"), CINZA), True)
        _texto(slide, 2.3, y - 0.02, 10.5, 0.95, item["texto"].replace("<b>", "").replace("</b>", ""), 13, TEXTO)
        y += 0.95
    saida = io.BytesIO()
    apresentacao.save(saida)
    return saida.getvalue()


def main(argv):
    ns = carregar_funcoes_do_app()
    fatos, itens, ctx = montar_briefing(ns, url_fech=os.environ.get("FECHAMENTO_CSV_URL", ""))
    hoje = ctx["hoje"]
    series = series_mensais(ns, ctx["list_df_real"], ctx["list_df_orc"], ctx["meses_cols"], ctx["m_map"],
                            ate_mes=hoje.month - 1 if hoje.month > 1 else 12)
    fatos["responsaveis"] = responsaveis_por_conta(ns, ctx.get("linhas", []))
    pptx_bytes = montar_board_pack(fatos, itens, series, hoje, ns)
    nome = f"board_pack_{hoje.strftime('%Y-%m')}.pptx"
    if "--salvar" in argv:
        nome = argv[argv.index("--salvar") + 1]
    if "--teste" in argv or "--salvar" in argv:
        with open(nome, "wb") as arquivo:
            arquivo.write(pptx_bytes)
        print(f"Board pack gravado em {nome} ({len(pptx_bytes) // 1024} KB)")
        return
    _status, _cor, frase = status_geral(fatos)
    assunto = f"Board pack · {ctx['rotulo']}" + (f" · {frase}" if frase else "")
    texto = (f"Segue o board pack do fechamento ({ctx['rotulo']}), gerado automaticamente pelo painel.\n\n"
             + "\n".join(f"{i['rotulo']}: {i['texto'].replace('<b>', '').replace('</b>', '')}" for i in itens))
    html = ("<p>Segue o board pack do fechamento, gerado automaticamente pelo painel.</p><ul>"
            + "".join(f"<li><b>{i['rotulo']}:</b> {i['texto']}</li>" for i in itens) + "</ul>")
    destinos = enviar_email(assunto, html, texto, str(ns.get("LOGO_BEEA_B64") or ""),
                            anexos=[(nome, pptx_bytes, "application",
                                     "vnd.openxmlformats-officedocument.presentationml.presentation")])
    print(f"Board pack enviado para {', '.join(destinos)}: {assunto}")


if __name__ == "__main__":
    main(sys.argv[1:])
