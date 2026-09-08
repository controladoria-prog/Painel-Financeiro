# -*- coding: utf-8 -*-
"""Pódio semanal das lojas -- segunda-feira, 10h de Porto Velho.

Ranking das lojas pelo ritmo do mês (receita realizada contra a meta até a
data dos dados) e pelo EBITDA acumulado nos meses fechados. Um e-mail para a
diretoria com a tabela inteira e, se houver o mapa de gerentes (secret
EMAILS_LOJAS, JSON {"LJ PVH2 14625": "gerente@..."}), um e-mail por loja com
a posição dela. Competição saudável é a alavanca mais barata do varejo.

Uso:
    python podio.py            calcula e envia
    python podio.py --teste    calcula e imprime, sem enviar
"""
import os
import sys

from briefing import (CID_LOGO, CORES, FONTE, DIAS_SEMANA, carregar_funcoes_do_app, emails_do_departamento,
                      enviar_email, lojas_do_departamento, lojas_oficiais, moldura_email, montar_briefing,
                      urls_das_planilhas)


def ranking_das_lojas(ns, dados_por_loja, ritmo, cols_fechados):
    """[{loja, rec_mes, meta_mes, ritmo_pct, ebitda_real, ebitda_orc, desvio}] ordenado pelo ritmo."""
    gv = ns["get_valor_consolidado_multi"]
    receita, ebitda = "3 - Receita Operacional Liquida", "11 - EBITDA"
    saida = []
    oficiais = set(lojas_oficiais(ns))
    for loja, (df_o, df_r) in dados_por_loja.items():
        if oficiais and loja not in oficiais:
            continue
        lr = [df_r] if df_r is not None and not df_r.empty else []
        lo = [df_o] if df_o is not None and not df_o.empty else []
        rec_mes = gv(lr, receita, [ritmo["col"]]) if lr and ritmo else 0.0
        meta_mes = (gv(lo, receita, [ritmo["col"]]) * ritmo["frac"]) if lo and ritmo else 0.0
        eb_r = gv(lr, ebitda, cols_fechados) if lr else 0.0
        eb_o = gv(lo, ebitda, cols_fechados) if lo else 0.0
        if not rec_mes and not eb_r and not eb_o:
            continue
        saida.append({"loja": loja, "rec_mes": rec_mes, "meta_mes": meta_mes,
                      "ritmo_pct": (rec_mes / meta_mes * 100) if meta_mes else None,
                      "ebitda_real": eb_r, "ebitda_orc": eb_o, "desvio": eb_r - eb_o})
    # Sem meta de receita (o escritório) não se compete: fica fora do pódio.
    saida = [r for r in saida if r["ritmo_pct"] is not None]
    saida.sort(key=lambda r: -r["ritmo_pct"])
    for i, r in enumerate(saida, start=1):
        r["posicao"] = i
    return saida


def montar_email_podio(ranking, ritmo, hoje, link="", logo_src="", ns=None, destaque=None, periodo_fechado=""):
    """Cada célula se explica sozinha: valor principal em destaque, o
    complemento em letra menor logo abaixo, cabeçalho com o período."""
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    dia = f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}"
    mes = str(ritmo["mes"]).capitalize() if ritmo else ""
    data_dados = ritmo["data_dados"] if ritmo else ""
    selo = {1: ("1º", "#C9A227"), 2: ("2º", "#8E9AAF"), 3: ("3º", "#B87333")}
    sub = f'style="font-size:11px; color:{CORES["apagado"]}; margin-top:2px;"'
    linhas = ""
    for r in ranking:
        pct = r["ritmo_pct"]
        cor_r = (CORES["positivo"] if pct >= 100 else CORES["alerta"] if pct >= 90 else CORES["negativo"])
        cor_e = CORES["positivo"] if r["desvio"] >= 0 else CORES["negativo"]
        fundo = "#FFF6DA" if destaque and r["loja"] == destaque else ("#F4F7FB" if r["posicao"] % 2 == 0 else "#FFFFFF")
        rotulo, cor_selo = selo.get(r["posicao"], (str(r["posicao"]), CORES["borda"]))
        cor_txt_selo = "#FFFFFF" if r["posicao"] <= 3 else CORES["texto"]
        linhas += (
            f'<tr style="background:{fundo};">'
            f'<td style="padding:9px 6px 9px 8px; font-family:{FONTE};">'
            f'<span style="display:inline-block; min-width:26px; padding:3px 6px; border-radius:11px; background:{cor_selo}; '
            f'color:{cor_txt_selo}; font-size:11px; font-weight:700; text-align:center;">{rotulo}</span></td>'
            f'<td style="padding:9px 8px; font-family:{FONTE}; font-size:13px; font-weight:600; color:{CORES["texto"]};">{r["loja"]}</td>'
            f'<td align="right" style="padding:9px 8px; font-family:{FONTE};">'
            f'<div style="font-size:15px; font-weight:700; color:{cor_r};">{pct:.0f}%</div>'
            f'<div {sub}>{"acima" if pct >= 100 else "abaixo"} da meta</div></td>'
            f'<td align="right" style="padding:9px 8px; font-family:{FONTE};">'
            f'<div style="font-size:13px; font-weight:700; color:{CORES["texto"]};">{fmt(r["rec_mes"])}</div>'
            f'<div {sub}>meta até {data_dados}: {fmt(r["meta_mes"])}</div></td>'
            f'<td align="right" style="padding:9px 8px; font-family:{FONTE};">'
            f'<div style="font-size:13px; font-weight:700; color:{cor_e};">{"+" if r["desvio"] >= 0 else "−"}{fmt(abs(r["desvio"]))} '
            f'{"acima" if r["desvio"] >= 0 else "abaixo"}</div>'
            f'<div {sub}>{fmt(r["ebitda_real"])} vs {fmt(r["ebitda_orc"])} orçado</div></td></tr>')
    cab = "".join(
        f'<th align="{al}" style="padding:6px 8px; font-family:{FONTE}; font-size:10px; letter-spacing:1px; line-height:1.4; '
        f'text-transform:uppercase; color:{CORES["apagado"]}; border-bottom:2px solid {CORES["marca"]};">{t}</th>'
        for t, al in (("#", "left"), ("Loja", "left"), (f"Ritmo de {mes}", "right"),
                      (f"Receita de {mes}<br>até {data_dados}", "right"), (f"EBITDA {periodo_fechado}<br>vs orçado", "right")))
    legenda = (
        f'<div style="font-size:12px; line-height:1.6; color:{CORES["apagado"]}; margin-bottom:10px;">'
        f'<b style="color:{CORES["texto"]};">Ritmo</b>: quanto a loja já vendeu em {mes.lower()} em relação à meta proporcional '
        f'até {data_dados} (100% = no ritmo da meta; os dados chegam D+2).<br>'
        f'<b style="color:{CORES["texto"]};">Receita</b>: o vendido no mês e a meta até essa data.<br>'
        f'<b style="color:{CORES["texto"]};">EBITDA</b>: acumulado dos meses fechados ({periodo_fechado}) contra o orçado do mesmo período. '
        'A ordem do ranking é pelo ritmo.</div>')
    corpo = legenda + ('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;">'
                       f"<tr>{cab}</tr>{linhas}</table>")
    titulo = f"Pódio das lojas · {mes}" + (f" · {destaque}" if destaque else "")
    lider = ranking[0]["loja"] if ranking else ""
    html = moldura_email(titulo, dia, "RANKING", CORES["marca"], corpo, link, logo_src,
                         f"Gerado automaticamente pelo painel toda segunda-feira. Quem lidera o ritmo de {mes.lower()}: {lider}.")
    texto = f"{titulo} · {dia}\n\n" + "\n".join(
        f"{r['posicao']}. {r['loja']} · ritmo {r['ritmo_pct']:.0f}% (vendeu {fmt(r['rec_mes'])} de meta {fmt(r['meta_mes'])} até {data_dados})"
        f" · EBITDA {periodo_fechado} {'+' if r['desvio'] >= 0 else '-'}{fmt(abs(r['desvio']))} vs orçado" for r in ranking)
    return html, texto


def main(argv):
    ns = carregar_funcoes_do_app()
    fatos, _itens, ctx = montar_briefing(ns, url_fech=os.environ.get("FECHAMENTO_CSV_URL", ""))
    url_orc, url_real = urls_das_planilhas()
    lojas = lojas_oficiais(ns)
    hoje = ctx["hoje"]
    cols_fech = [c for c in ctx["m_map"].values() if int(c[:2]) < hoje.month]
    nomes_fech = [n for n, c in ctx["m_map"].items() if c in cols_fech]
    periodo_fechado = (f"{nomes_fech[0][:3].lower()}–{nomes_fech[-1][:3].lower()}" if len(nomes_fech) > 1
                       else (nomes_fech[0][:3].lower() if nomes_fech else ""))
    dados = ns["carregar_dados_por_loja"](url_orc, url_real, lojas)
    ranking = ranking_das_lojas(ns, dados, fatos.get("ritmo"), cols_fech)
    if not ranking:
        raise SystemExit("Nenhuma loja com dados.")
    logo_b64 = str(ns.get("LOGO_BEEA_B64") or "")
    logo_src = f"cid:{CID_LOGO}" if logo_b64 else ""
    link = os.environ.get("LINK_PAINEL", "")
    html, texto = montar_email_podio(ranking, fatos.get("ritmo"), hoje, link, logo_src, ns, periodo_fechado=periodo_fechado)
    if "--teste" in argv:
        print(texto)
        return
    assunto = f"Pódio das lojas {hoje.strftime('%d/%m')} · {ranking[0]['loja']} lidera"
    destinos = enviar_email(assunto, html, texto, logo_b64)
    print(f"Pódio enviado para {', '.join(destinos)}.")
    for departamento in (ns.get("MODELOS_RELATORIO") or {}):
        emails = emails_do_departamento(departamento, ns.get("MAPA_EMAIL_DEPARTAMENTO"),
                                        ns.get("EMAILS_TRAVADOS_NO_DEPARTAMENTO"))
        minhas = set(lojas_do_departamento(departamento, ns))
        recorte = [r for r in ranking if r["loja"] in minhas]
        if not emails or not recorte or len(recorte) == len(ranking) and "Comercial" not in departamento:
            continue
        curto = departamento.split(" - ")[-1].strip()
        html_d, texto_d = montar_email_podio(recorte, fatos.get("ritmo"), hoje, link, logo_src, ns, destaque=curto,
                                             periodo_fechado=periodo_fechado)
        enviar_email(f"Pódio das lojas {hoje.strftime('%d/%m')} · {curto} · {recorte[0]['loja']} lidera", html_d, texto_d,
                     logo_b64, destinos=emails)
        print(f"Pódio de {curto} enviado para {', '.join(emails)}.")

if __name__ == "__main__":
    main(sys.argv[1:])
