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
import json
import os
import sys

from briefing import (CID_LOGO, CORES, FONTE, DIAS_SEMANA, carregar_funcoes_do_app, enviar_email,
                      lojas_do_workbook, moldura_email, montar_briefing, urls_das_planilhas)


def ranking_das_lojas(ns, dados_por_loja, ritmo, cols_fechados):
    """[{loja, rec_mes, meta_mes, ritmo_pct, ebitda_real, ebitda_orc, desvio}] ordenado pelo ritmo."""
    gv = ns["get_valor_consolidado_multi"]
    receita, ebitda = "3 - Receita Operacional Liquida", "11 - EBITDA"
    saida = []
    for loja, (df_o, df_r) in dados_por_loja.items():
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
    saida.sort(key=lambda r: -(r["ritmo_pct"] if r["ritmo_pct"] is not None else -1))
    for i, r in enumerate(saida, start=1):
        r["posicao"] = i
    return saida


def montar_email_podio(ranking, ritmo, hoje, link="", logo_src="", ns=None, destaque=None):
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    dia = f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}"
    mes = str(ritmo["mes"]).capitalize() if ritmo else ""
    medalha = {1: "🥇", 2: "🥈", 3: "🥉"}
    linhas = ""
    for r in ranking:
        cor = (CORES["positivo"] if (r["ritmo_pct"] or 0) >= 100 else
               CORES["alerta"] if (r["ritmo_pct"] or 0) >= 90 else CORES["negativo"])
        fundo = "#FFF6DA" if destaque and r["loja"] == destaque else ("#F4F7FB" if r["posicao"] % 2 == 0 else "#FFFFFF")
        ritmo_txt = f"{r['ritmo_pct']:.0f}%" if r["ritmo_pct"] is not None else "—"
        linhas += (f'<tr style="background:{fundo};">'
                   f'<td style="padding:7px 8px; font-family:{FONTE}; font-size:13px; color:{CORES["texto"]};">{medalha.get(r["posicao"], r["posicao"])}</td>'
                   f'<td style="padding:7px 8px; font-family:{FONTE}; font-size:13px; font-weight:600; color:{CORES["texto"]};">{r["loja"]}</td>'
                   f'<td align="right" style="padding:7px 8px; font-family:{FONTE}; font-size:13px; font-weight:700; color:{cor};">{ritmo_txt}</td>'
                   f'<td align="right" style="padding:7px 8px; font-family:{FONTE}; font-size:12px; color:{CORES["apagado"]};">{fmt(r["rec_mes"])} / {fmt(r["meta_mes"])}</td>'
                   f'<td align="right" style="padding:7px 8px; font-family:{FONTE}; font-size:13px; font-weight:700; '
                   f'color:{CORES["positivo"] if r["desvio"] >= 0 else CORES["negativo"]};">{"+" if r["desvio"] >= 0 else "−"}{fmt(abs(r["desvio"]))}</td></tr>')
    cabecalho = "".join(f'<th align="{al}" style="padding:6px 8px; font-family:{FONTE}; font-size:10px; letter-spacing:1px; '
                        f'text-transform:uppercase; color:{CORES["apagado"]}; border-bottom:2px solid {CORES["marca"]};">{t}</th>'
                        for t, al in (("#", "left"), ("Loja", "left"), (f"Ritmo de {mes}", "right"),
                                      ("Realizado / meta", "right"), ("EBITDA vs orçado", "right")))
    corpo = (f'<div style="font-size:13px; color:{CORES["apagado"]}; margin-bottom:8px;">Ritmo = receita do mês contra a meta até '
             f'{ritmo["data_dados"] if ritmo else ""} (dados D+2). EBITDA = acumulado dos meses fechados.</div>'
             f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;">'
             f"<tr>{cabecalho}</tr>{linhas}</table>")
    titulo = f"Pódio das lojas · {mes}" + (f" · {destaque}" if destaque else "")
    lider = ranking[0]["loja"] if ranking else ""
    html = moldura_email(titulo, dia, "RANKING", CORES["marca"], corpo, link, logo_src,
                         "Gerado automaticamente pelo painel toda segunda-feira. Quem lidera hoje: " + lider + ".")
    texto = f"{titulo} · {dia}\n\n" + "\n".join(
        f"{r['posicao']}. {r['loja']} · ritmo {r['ritmo_pct']:.0f}% · EBITDA {'+' if r['desvio'] >= 0 else '-'}{fmt(abs(r['desvio']))}"
        if r["ritmo_pct"] is not None else f"{r['posicao']}. {r['loja']}" for r in ranking)
    return html, texto


def main(argv):
    ns = carregar_funcoes_do_app()
    fatos, _itens, ctx = montar_briefing(ns, url_fech=os.environ.get("FECHAMENTO_CSV_URL", ""))
    url_orc, url_real = urls_das_planilhas()
    lojas = [l.strip() for l in os.environ.get("LOJAS", "").split(",") if l.strip()] or lojas_do_workbook(url_real, ns)
    hoje = ctx["hoje"]
    cols_fech = [c for c in ctx["m_map"].values() if int(c[:2]) < hoje.month]
    dados = ns["carregar_dados_por_loja"](url_orc, url_real, lojas)
    ranking = ranking_das_lojas(ns, dados, fatos.get("ritmo"), cols_fech)
    if not ranking:
        raise SystemExit("Nenhuma loja com dados.")
    logo_b64 = str(ns.get("LOGO_BEEA_B64") or "")
    logo_src = f"cid:{CID_LOGO}" if logo_b64 else ""
    link = os.environ.get("LINK_PAINEL", "")
    html, texto = montar_email_podio(ranking, fatos.get("ritmo"), hoje, link, logo_src, ns)
    if "--teste" in argv:
        print(texto)
        return
    assunto = f"Pódio das lojas {hoje.strftime('%d/%m')} · {ranking[0]['loja']} lidera"
    destinos = enviar_email(assunto, html, texto, logo_b64)
    print(f"Pódio enviado para {', '.join(destinos)}.")
    try:
        gerentes = json.loads(os.environ.get("EMAILS_LOJAS", "") or "{}")
    except ValueError:
        gerentes = {}
    for r in ranking:
        email = gerentes.get(r["loja"])
        if not email:
            continue
        html_l, texto_l = montar_email_podio(ranking, fatos.get("ritmo"), hoje, link, logo_src, ns, destaque=r["loja"])
        enviar_email(f"Pódio das lojas {hoje.strftime('%d/%m')} · {r['loja']} está em {r['posicao']}º", html_l, texto_l,
                     logo_b64, destinos=[e.strip() for e in str(email).split(",") if e.strip()])
        print(f"Pódio de {r['loja']} enviado para {email}.")


if __name__ == "__main__":
    main(sys.argv[1:])
