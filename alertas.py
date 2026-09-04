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
    python alertas.py --exemplo  manda um alerta ficticio, so para ver a cara do e-mail

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
                      linhas_do_departamento, moldura_email, montar_briefing, urls_das_planilhas)

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


def montar_email_alerta(alertas, fatos, hoje, link_painel="", logo_src="", departamento=None):
    """A mesma moldura do briefing; o selo vermelho e as linhas com barra
    dizem que é alerta."""
    dia = f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')} · {datetime.now().strftime('%H:%M')} UTC"
    cor = {"negativo": CORES["negativo"], "alerta": CORES["alerta"]}
    linhas = ('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;">' + "".join(
        "<tr>"
        f'<td width="4" bgcolor="{cor.get(x["tom"], CORES["alerta"])}" style="width:4px; background:{cor.get(x["tom"], CORES["alerta"])}; font-size:0;">&nbsp;</td>'
        f'<td style="padding:12px 0 12px 14px; font-family:{FONTE}; border-bottom:1px solid {CORES["borda"]};">'
        f'<div style="font-size:15px; font-weight:700; color:{CORES["texto"]};">{x["titulo"]}</div>'
        f'<div style="font-size:13px; color:{CORES["apagado"]}; margin-top:3px;">{x["detalhe"]}</div></td></tr>'
        for x in alertas) + "</table>")
    n = len(alertas)
    titulo = f"{n} ponto{'s' if n != 1 else ''} de atenção" + (f" · {departamento}" if departamento else "")
    html = moldura_email(titulo, dia, "ALERTA", CORES["negativo"], linhas, link_painel, logo_src,
                         "Este aviso só é enviado quando algo cruza uma linha, e cada ponto é avisado uma vez. "
                         "Base: DRE realizada e orçada, meses fechados; lançamentos chegam D+2.")
    texto = (f"Controladoria B&A · alerta · {dia}" + (f" · {departamento}" if departamento else "") + "\n\n"
             + "\n".join(f"- {x['titulo']}: {x['detalhe']}" for x in alertas)
             + (f"\n\nPainel: {link_painel}" if link_painel else ""))
    return html, texto


def enviar_exemplo(ns, hoje, link):
    """Alertas de mentira, só para ver a cara do e-mail: o geral vai para o
    EMAIL_DESTINO; o de cada departamento vai SÓ para a cópia da controladoria,
    com o nome do gestor a quem iria -- exemplo fictício não entra na caixa
    dos gestores. A memória dos alertas fica intocada."""
    logo_b64 = str(ns.get("LOGO_BEEA_B64") or "")
    logo_src = f"cid:{CID_LOGO}" if logo_b64 else ""
    geral = [
        {"chave": "exemplo:1", "titulo": "Taxa de Emissão de Boleto passou do orçado",
         "detalhe": "+R$ 832 mil acima do orçado (+103%) nos meses fechados. (EXEMPLO)", "tom": "negativo"},
        {"chave": "exemplo:2", "titulo": "Setembro corre a 76% do esperado no dia 17",
         "detalhe": "Realizado R$ 6,1M contra meta de R$ 8,0M até 15/09 (D+2). Chance de bater a meta: 18%. (EXEMPLO)",
         "tom": "alerta"},
    ]
    html, texto = montar_email_alerta(geral, {}, hoje, link, logo_src)
    destinos = enviar_email(f"[EXEMPLO] Alerta {hoje.strftime('%d/%m')} · como o aviso geral chega", html, texto, logo_b64)
    print(f"Exemplo geral enviado para {', '.join(destinos)}.")
    copia = [e.strip() for e in os.environ.get("EMAIL_COPIA_DEPARTAMENTOS", os.environ.get("SMTP_USUARIO", "")).split(",") if e.strip()]
    for departamento in (ns.get("MODELOS_RELATORIO") or {}):
        gestores = emails_do_departamento(departamento, ns.get("MAPA_EMAIL_DEPARTAMENTO"),
                                          ns.get("EMAILS_TRAVADOS_NO_DEPARTAMENTO"))
        if not gestores or not copia:
            continue
        exemplo = [{"chave": "exemplo:dep", "titulo": "Serviços de Terceiros passou do orçado",
                    "detalhe": f"+R$ 120 mil acima do orçado (+20%) nos meses fechados · {departamento}. (EXEMPLO)",
                    "tom": "negativo"}]
        html, texto = montar_email_alerta(exemplo, {}, hoje, link, logo_src, departamento)
        enviar_email(f"[EXEMPLO] Alerta {hoje.strftime('%d/%m')} · {departamento} · iria para {', '.join(gestores)}",
                     html, texto, logo_b64, destinos=copia)
        print(f"Exemplo de {departamento} enviado para {', '.join(copia)} (no real iria para {', '.join(gestores)}).")


def main(argv):
    ns = carregar_funcoes_do_app()
    if "--exemplo" in argv:
        enviar_exemplo(ns, datetime.now(ns["FUSO_BR"]).date(), os.environ.get("LINK_PAINEL", ""))
        return
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
    # A controladoria recebe cópia de todo alerta de departamento (EMAIL_COPIA_DEPARTAMENTOS;
    # por padrão, a própria conta que envia).
    copia = [e.strip() for e in os.environ.get("EMAIL_COPIA_DEPARTAMENTOS", os.environ.get("SMTP_USUARIO", "")).split(",") if e.strip()]
    for departamento, emails, al in novos_dep:
        html, texto = montar_email_alerta(al, fatos, ctx["hoje"], link, logo_src, departamento)
        assunto = f"Alerta {ctx['hoje'].strftime('%d/%m')} · {departamento} · " + al[0]["titulo"]
        destinos = sorted(set(emails) | set(copia))
        enviar_email(assunto, html, texto, logo_b64, destinos=destinos)
        print(f"Alerta de {departamento} enviado para {', '.join(destinos)}")
        enviados += [x["chave"] for x in al]
    salvar_estado(estado + enviados)


if __name__ == "__main__":
    main(sys.argv[1:])
