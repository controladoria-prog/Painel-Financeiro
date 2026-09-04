# -*- coding: utf-8 -*-
"""Alerta imediato -- roda a cada duas horas em dia útil e SÓ manda e-mail
quando algo cruzou uma linha. Silêncio quando está tudo bem: é por isso que,
quando chega, é lido.

Reaproveita o briefing inteiro (carga das planilhas, fatos, e-mail) e
acrescenta as regras e uma memória do que já foi avisado -- o mesmo alerta
não chega duas vezes.

Uso:
    python alertas.py            avalia e envia (se houver alerta novo)
    python alertas.py --teste    avalia e imprime, sem enviar nem lembrar

Regras (todas sobre a DRE, que é o que roda sem o app):
    1. Conta acima do orçado na base fechada, com estouro relevante
       (>= R$ 100 mil OU >= 20%). Um aviso por conta e por mês.
    2. Ritmo do mês abaixo de 80% a partir do dia 15. Um aviso por mês.
    3. Por GESTOR: as contas do departamento dele (escopo resolvido pelos
       modelos do próprio app) acima do orçado -- e-mail só para ele.

LINHA DE BASE: na primeira execução (sem dados/estado_alertas.json) tudo o
que já passou da linha é apenas REGISTRADO, sem e-mail. Muita coisa já tinha
passado antes do robô existir, e avisar o passado todo de uma vez é ruído.
A partir daí, só o que cruzar a linha de novo é avisado -- uma vez.
"""
import json
import os
import sys
from datetime import datetime

from briefing import (CID_LOGO, CORES, FONTE, DIAS_SEMANA, carregar_funcoes_do_app,
                      emails_do_departamento, enviar_email, fatos_do_departamento,
                      linhas_do_departamento, montar_briefing, urls_das_planilhas)

CAMINHO_ESTADO = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "dados", "estado_alertas.json")
LIMITE_ESTADO = 500
ESTOURO_MINIMO_REAIS = 100_000
ESTOURO_MINIMO_PCT = 20
ESTOURO_MINIMO_REAIS_DEPARTAMENTO = 20_000
RITMO_MINIMO_PCT = 80
DIA_MINIMO_RITMO = 15


def carregar_estado(caminho=CAMINHO_ESTADO):
    try:
        with open(caminho, encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
        return dados if isinstance(dados, list) else []
    except (OSError, ValueError):
        return []


def salvar_estado(chaves, caminho=CAMINHO_ESTADO):
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "w", encoding="utf-8") as arquivo:
        json.dump(chaves[-LIMITE_ESTADO:], arquivo, ensure_ascii=False, indent=1)


def avaliar_alertas(fatos, hoje, fmt=None):
    """Lista de {chave, titulo, detalhe, tom}. A chave é o que impede o mesmo
    aviso de sair duas vezes."""
    fmt = fmt or (lambda v: f"R$ {v:,.0f}")
    alertas = []
    mc = fatos.get("mes_corrente") or {}
    mes_ref = str(mc.get("col") or hoje.strftime("%m/%Y"))
    for e in fatos.get("estouros") or []:
        pct = e.get("pct")
        if e.get("desvio", 0) >= ESTOURO_MINIMO_REAIS or (pct is not None and pct >= ESTOURO_MINIMO_PCT):
            alertas.append({
                "chave": f"estouro:{e['conta']}:{mes_ref}",
                "titulo": f"{e['conta']} passou do orçado",
                "detalhe": (f"+{fmt(e['desvio'])} acima do orçado"
                            + (f" (+{pct:.0f}%)" if pct is not None else " (conta sem orçamento)")
                            + " nos meses fechados."),
                "tom": "negativo",
            })
    r = fatos.get("ritmo") or {}
    if r and (r.get("dia") or 0) >= DIA_MINIMO_RITMO and r.get("pct") is not None and r["pct"] < RITMO_MINIMO_PCT:
        detalhe = (f"Realizado {fmt(r.get('rec_real', 0))} contra meta de {fmt(r.get('rec_orc_prop', 0))} "
                   f"até {r.get('data_dados', '')} (D+2).")
        if r.get("chance") is not None:
            detalhe += f" Chance de bater a meta do mês: {r['chance'] * 100:.0f}%."
        alertas.append({
            "chave": f"ritmo:{r.get('col')}",
            "titulo": f"{str(r.get('mes', '')).capitalize()} corre a {r['pct']:.0f}% do esperado no dia {r['dia']}",
            "detalhe": detalhe,
            "tom": "negativo" if r["pct"] < 70 else "alerta",
        })
    return alertas


def alertas_dos_departamentos(ns, ctx, fmt=None):
    """[(departamento, e-mails, alertas)] -- um bloco por gestor com e-mail
    cadastrado e escopo resolvível fora do app."""
    fmt = fmt or (lambda v: f"R$ {v:,.0f}")
    gv = ns["get_valor_consolidado_multi"]
    hoje = ctx["hoje"]
    col_corrente = f"{hoje.month:02d}/{hoje.year}"
    cols_kpi = [c for c in ctx["m_map"].values() if int(c[:2]) <= hoje.month]
    url_orc, url_real = urls_das_planilhas()
    saida = []
    for departamento, modelo in (ns.get("MODELOS_RELATORIO") or {}).items():
        emails = emails_do_departamento(departamento, ns.get("MAPA_EMAIL_DEPARTAMENTO"),
                                        ns.get("EMAILS_TRAVADOS_NO_DEPARTAMENTO"))
        if not emails:
            continue
        visoes = modelo.get("visoes_permitidas") or []
        if visoes and visoes[0] != ctx["aba"]:
            list_df_orc, list_df_real = ns["carregar_dados_abas"](url_orc, url_real, [visoes[0]])
            df_ref = next((d for d in list_df_real if d is not None and not d.empty), None)
            if df_ref is None:
                continue
            col_nome = "Nome" if "Nome" in df_ref.columns else df_ref.columns[0]
            linhas_visao = list(df_ref[col_nome].dropna().unique().astype(str))
        else:
            list_df_orc, list_df_real, linhas_visao = ctx["list_df_orc"], ctx["list_df_real"], ctx["linhas"]
        linhas = linhas_do_departamento(modelo, linhas_visao, ns)
        if not linhas:
            continue

        def valor(lado, linha, cols, exato=False, _r=list_df_real, _o=list_df_orc):
            return gv(_o if lado == "orc" else _r, linha, cols, exato_linha_sintetica=exato)

        f = fatos_do_departamento(valor, linhas, cols_kpi, col_corrente, departamento, ns)
        alertas = []
        for e in f["estouros"]:
            pct = e.get("pct")
            if e["desvio"] >= ESTOURO_MINIMO_REAIS_DEPARTAMENTO or (pct is not None and pct >= ESTOURO_MINIMO_PCT):
                alertas.append({
                    "chave": f"dep:{departamento}:{e['conta']}:{col_corrente}",
                    "titulo": f"{e['conta']} passou do orçado",
                    "detalhe": (f"+{fmt(e['desvio'])} acima do orçado"
                                + (f" (+{pct:.0f}%)" if pct is not None else " (conta sem orçamento)")
                                + f" nos meses fechados · {departamento}"),
                    "tom": "negativo",
                })
        saida.append((departamento, emails, alertas))
    return saida


def eh_primeira_execucao(caminho=CAMINHO_ESTADO):
    return not os.path.exists(caminho)


def novos_alertas(alertas, chaves_enviadas):
    ja = set(chaves_enviadas or [])
    return [a for a in alertas if a["chave"] not in ja]


def montar_email_alerta(alertas, fatos, hoje, link_painel="", logo_src=""):
    dia = f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')} · {datetime.now().strftime('%H:%M')} UTC"
    cor = {"negativo": CORES["negativo"], "alerta": CORES["alerta"]}
    linhas = "".join(
        "<tr>"
        f'<td width="4" bgcolor="{cor.get(a["tom"], CORES["alerta"])}" style="width:4px; background:{cor.get(a["tom"], CORES["alerta"])}; font-size:0;">&nbsp;</td>'
        f'<td style="padding:12px 0 12px 14px; font-family:{FONTE}; border-bottom:1px solid {CORES["borda"]};">'
        f'<div style="font-size:15px; font-weight:700; color:{CORES["texto"]};">{a["titulo"]}</div>'
        f'<div style="font-size:13px; color:{CORES["apagado"]}; margin-top:3px;">{a["detalhe"]}</div></td></tr>'
        for a in alertas)
    selo = (f'<img src="{logo_src}" width="40" height="40" alt="Grupo B&amp;A" style="display:block; border-radius:20px;">'
            if logo_src else "")
    botao = (f'<a href="{link_painel}" style="display:inline-block; margin-top:14px; padding:10px 22px; background:{CORES["marca"]}; '
             f'color:#FFFFFF; font-family:{FONTE}; font-size:13px; font-weight:600; text-decoration:none; border-radius:6px;">Abrir o painel &rarr;</a>'
             if link_painel else "")
    html = (
        '<meta charset="utf-8">'
        f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" bgcolor="{CORES["fundo"]}" style="background:{CORES["fundo"]};">'
        '<tr><td align="center" style="padding:24px 12px;">'
        '<table role="presentation" width="640" cellpadding="0" cellspacing="0" style="width:640px; max-width:100%;">'
        f'<tr><td bgcolor="{CORES["negativo"]}" style="background:{CORES["negativo"]}; padding:18px 24px; border-radius:10px 10px 0 0;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;"><tr>'
        f'<td width="40">{selo}</td>'
        f'<td style="padding-left:12px; font-family:{FONTE};">'
        '<div style="font-size:10px; letter-spacing:1.6px; text-transform:uppercase; color:rgba(255,255,255,0.8);">Controladoria B&amp;A · alerta</div>'
        f'<div style="font-size:19px; font-weight:700; color:#FFFFFF; margin-top:2px;">{len(alertas)} ponto{"s" if len(alertas) != 1 else ""} de atenção</div>'
        f'<div style="font-size:12px; color:rgba(255,255,255,0.8); margin-top:2px;">{dia}</div></td></tr></table></td></tr>'
        f'<tr><td align="center" style="background:{CORES["cartao"]}; padding:8px 26px 22px 26px; border-radius:0 0 10px 10px;">'
        f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;">{linhas}</table>'
        f'{botao}'
        f'<div style="font-family:{FONTE}; font-size:11px; color:{CORES["apagado"]}; margin-top:14px;">'
        'Este aviso só é enviado quando algo cruza uma linha, e cada ponto é avisado uma vez. '
        'Base: DRE realizada e orçada, meses fechados; lançamentos chegam D+2.</div>'
        "</td></tr></table></td></tr></table>")
    texto = (f"Controladoria B&A · alerta · {dia}\n\n"
             + "\n".join(f"- {a['titulo']}: {a['detalhe']}" for a in alertas)
             + (f"\n\nPainel: {link_painel}" if link_painel else ""))
    return html, texto


def main(argv):
    ns = carregar_funcoes_do_app()
    fatos, _itens, ctx = montar_briefing(ns, url_fech=os.environ.get("FECHAMENTO_CSV_URL", ""))
    fmt = ns.get("formata_valor_curto")
    geral = avaliar_alertas(fatos, ctx["hoje"], fmt)
    por_departamento = alertas_dos_departamentos(ns, ctx, fmt)
    todas_as_chaves = [x["chave"] for x in geral] + [x["chave"] for _, _, al in por_departamento for x in al]
    if eh_primeira_execucao() and "--teste" not in argv:
        salvar_estado(todas_as_chaves)
        print(f"Linha de base registrada: {len(todas_as_chaves)} ponto(s) que já tinham passado da linha "
              "ficam sem aviso. A partir de agora só o que cruzar de novo é avisado.")
        return
    estado = carregar_estado()
    novos_geral = novos_alertas(geral, estado)
    novos_dep = [(d, emails, novos_alertas(al, estado)) for d, emails, al in por_departamento]
    novos_dep = [(d, emails, al) for d, emails, al in novos_dep if al]
    if "--teste" in argv:
        print(f"Geral: {len(geral)} avaliado(s), {len(novos_geral)} novo(s).")
        for x in novos_geral:
            print(f"- {x['titulo']}: {x['detalhe']}")
        for d, emails, al in novos_dep:
            print(f"{d} -> {', '.join(emails)}: {len(al)} novo(s)")
            for x in al:
                print(f"- {x['titulo']}: {x['detalhe']}")
        return
    if not novos_geral and not novos_dep:
        print("Nada novo para avisar.")
        return
    logo_b64 = str(ns.get("LOGO_BEEA_B64") or "")
    link = os.environ.get("LINK_PAINEL", "")
    logo_src = f"cid:{CID_LOGO}" if logo_b64 else ""
    enviados = []
    if novos_geral:
        html, texto = montar_email_alerta(novos_geral, fatos, ctx["hoje"], link, logo_src)
        assunto = f"Alerta {ctx['hoje'].strftime('%d/%m')} · " + novos_geral[0]["titulo"] + (
            f" (+{len(novos_geral) - 1})" if len(novos_geral) > 1 else "")
        destinos = enviar_email(assunto, html, texto, logo_b64)
        print(f"Alerta geral enviado para {', '.join(destinos)}: {assunto}")
        enviados += [x["chave"] for x in novos_geral]
    for departamento, emails, al in novos_dep:
        html, texto = montar_email_alerta(al, fatos, ctx["hoje"], link, logo_src)
        assunto = f"Alerta {ctx['hoje'].strftime('%d/%m')} · {departamento} · " + al[0]["titulo"]
        enviar_email(assunto, html, texto, logo_b64, destinos=emails)
        print(f"Alerta de {departamento} enviado para {', '.join(emails)}")
        enviados += [x["chave"] for x in al]
    salvar_estado(estado + enviados)


if __name__ == "__main__":
    main(sys.argv[1:])
