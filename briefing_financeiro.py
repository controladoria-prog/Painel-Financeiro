# -*- coding: utf-8 -*-
"""Briefing FINANCEIRO (fluxo de caixa) -- mesma engrenagem do briefing da
Controladoria, sobre o painel financeiro: saldo de hoje (D+0: a planilha é
atualizada cedo), entradas e saídas do mês, vencidos, e os alertas do próprio
motor do painel (_avaliar_alertas_fluxo) com os limites padrão da barra
lateral. E-mail na mesma moldura, "desde ontem" pelo histórico.

Uso:
    python briefing_financeiro.py            calcula e envia
    python briefing_financeiro.py --teste    imprime, sem enviar
    python briefing_financeiro.py --forcar   reenvia mesmo que o de hoje já tenha saído
    python briefing_financeiro.py --alertas  modo alerta (uma vez por dia): só manda e-mail se houver ponto crítico novo
"""
import json
import os
import sys
from datetime import datetime

import pandas as pd

from briefing import (CID_LOGO, CORES, DIAS_SEMANA, FONTE, carregar_funcoes_do_app, enviar_email,
                      linhas_de_narrativa_html, moldura_email)

SEMENTES_FIN = [
    "obter_dados_fluxo_caixa", "preparar_fluxo_caixa", "_saldo_posicao_atual_fin", "_avaliar_alertas_fluxo",
    "COL_FIN_VALOR", "META_RESERVA_PADRAO", "FUSO_BR", "LOGO_BEEA_B64", "formata_valor_curto", "_pct_br",
]
LIMITE_VENCIDO_PADRAO = 50_000
LIMITE_CONCENTRACAO_PADRAO = 30
HORIZONTE_CANAL_PADRAO = 30
CAMINHO_HISTORICO_FIN = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "historico_financeiro.json")
CAMINHO_ESTADO_FIN = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "estado_alertas_financeiro.json")


def _ler_json(caminho, padrao):
    try:
        with open(caminho, encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
        return dados if isinstance(dados, type(padrao)) else padrao
    except (OSError, ValueError):
        return padrao


def _gravar_json(caminho, dados):
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=1)


def carregar_fluxo(ns):
    """Base preparada do fluxo, como o painel a usa. Levanta SystemExit com
    a mensagem do próprio loader se a planilha não vier."""
    base, erro, origem = ns["obter_dados_fluxo_caixa"]()
    if base is None:
        raise SystemExit(f"Fluxo de caixa indisponível: {erro}")
    df, erro_prep, *_ = ns["preparar_fluxo_caixa"](base)
    if df is None or getattr(df, "empty", True):
        raise SystemExit(f"Fluxo de caixa vazio: {erro_prep}")
    return df, origem


def fatos_do_caixa(ns, df, hoje):
    """Números do dia: saldo (D+0), entradas/saídas do mês até hoje e os
    alertas do motor do painel com os limites padrão da barra lateral."""
    col_valor = ns["COL_FIN_VALOR"]
    saldo, data_saldo = ns["_saldo_posicao_atual_fin"](df, col_valor)
    col_data = next((c for c in ("Data Efetiva", "Data Liquidação", "Data Vencimento", "Data") if c in df.columns), None)
    entradas = saidas = 0.0
    if col_data is not None:
        datas = pd.to_datetime(df[col_data], errors="coerce")
        no_mes = df[(datas.dt.year == hoje.year) & (datas.dt.month == hoje.month) & (datas.dt.date <= hoje)]
        entradas = float(no_mes.loc[no_mes["Tipo Movimento"] == "entrada", col_valor].abs().sum())
        saidas = float(no_mes.loc[no_mes["Tipo Movimento"] == "saida", col_valor].abs().sum())
    alertas = ns["_avaliar_alertas_fluxo"](df, col_valor, ns.get("META_RESERVA_PADRAO", 20), LIMITE_VENCIDO_PADRAO,
                                           LIMITE_CONCENTRACAO_PADRAO, horizonte_canal_dias=HORIZONTE_CANAL_PADRAO)
    return {
        "saldo": float(saldo or 0.0),
        "data_saldo": (pd.Timestamp(data_saldo).strftime("%d/%m") if data_saldo is not None and not pd.isna(data_saldo)
                       else hoje.strftime("%d/%m")),
        "entradas_mes": entradas, "saidas_mes": saidas, "liquido_mes": entradas - saidas,
        "alertas": [{"nivel": a.get("nivel", "atencao"), "titulo": str(a.get("titulo", "")),
                     "detalhe": str(a.get("detalhe", ""))} for a in (alertas or [])],
        "mes": hoje.strftime("%m/%Y"),
    }


def narrativa_financeira(f, ns=None):
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    itens = [{"rotulo": "Caixa", "tom": "positivo" if f["saldo"] >= 0 else "negativo",
              "texto": f"Saldo de <b>{fmt(f['saldo'])}</b> na posição de {f['data_saldo']}."}]
    if f.get("entradas_mes") or f.get("saidas_mes"):
        liq = f["liquido_mes"]
        itens.append({"rotulo": "Mês", "tom": "positivo" if liq >= 0 else "negativo",
                      "texto": (f"Até aqui no mês: entradas de <b>{fmt(f['entradas_mes'])}</b> e saídas de "
                                f"<b>{fmt(f['saidas_mes'])}</b> — líquido {'+' if liq >= 0 else '−'}{fmt(abs(liq))}.")})
    criticos = [a for a in f["alertas"] if a["nivel"] == "critico"]
    atencao = [a for a in f["alertas"] if a["nivel"] != "critico"]
    if criticos or atencao:
        partes = [f"<b>{a['titulo']}</b>" for a in criticos[:3]] + [a["titulo"] for a in atencao[:2]]
        itens.append({"rotulo": "Alertas", "tom": "negativo" if criticos else "alerta",
                      "texto": (f"{len(criticos)} crítico(s) e {len(atencao)} de atenção: " + "; ".join(partes) + ".")})
    else:
        itens.append({"rotulo": "Alertas", "tom": "positivo", "texto": "Nenhum alerta pelo motor do painel (reserva, vencidos e concentração dentro dos limites)."})
    return itens


def desde_ontem_fin(atual, anterior, fmt):
    if not anterior:
        return []
    d = atual["saldo"] - anterior.get("saldo", atual["saldo"])
    frases = [f"saldo {'+' if d >= 0 else '−'}{fmt(abs(d))} (era {fmt(anterior.get('saldo', 0))})"]
    novos = [a["titulo"] for a in atual["alertas"] if a["titulo"] not in anterior.get("titulos", [])]
    frases.append("alerta novo: " + ", ".join(novos) if novos else "nenhum alerta novo")
    return frases


def montar_email_financeiro(f, itens, hoje, link="", logo_src="", ns=None, desde=None, data_anterior=None):
    fmt = (ns or {}).get("formata_valor_curto", lambda v: f"R$ {v:,.0f}")
    dia = f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}"
    criticos = [a for a in f["alertas"] if a["nivel"] == "critico"]
    status, cor = ("ATENÇÃO", CORES["negativo"]) if criticos else (("OBSERVAR", CORES["alerta"]) if f["alertas"] else ("EM ORDEM", CORES["positivo"]))
    cartoes = [("Saldo em caixa", fmt(f["saldo"]), f"posição de {f['data_saldo']}"),
               ("Entradas do mês", fmt(f["entradas_mes"]), "até hoje"),
               ("Saídas do mês", fmt(f["saidas_mes"]), "até hoje"),
               ("Alertas", str(len(f["alertas"])), f"{len(criticos)} crítico(s)")]
    celulas = "".join(
        f'<td width="25%" valign="top" style="padding:6px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        f'style="width:100%; border:1px solid {CORES["borda"]}; border-top:3px solid {CORES["marca"]}; border-radius:6px;">'
        f'<tr><td style="padding:12px 12px; font-family:{FONTE};"><div style="font-size:10px; letter-spacing:1px; text-transform:uppercase; color:{CORES["apagado"]};">{r}</div>'
        f'<div style="font-size:20px; font-weight:700; color:{CORES["texto"]}; margin-top:3px;">{v}</div>'
        f'<div style="font-size:11px; color:{CORES["apagado"]}; margin-top:4px;">{s}</div></td></tr></table></td>'
        for r, v, s in cartoes)
    corpo = f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;"><tr>{celulas}</tr></table>'
    if desde:
        rot = f"Desde o briefing de {data_anterior[8:10]}/{data_anterior[5:7]}" if data_anterior else "Desde o último briefing"
        corpo += (f'<div style="background:#F4F7FB; border:1px solid {CORES["borda"]}; border-radius:6px; padding:10px 14px; margin:8px 6px 4px 6px;">'
                  f'<div style="font-size:10px; letter-spacing:1.2px; text-transform:uppercase; color:{CORES["apagado"]};">{rot}</div>'
                  f'<div style="font-size:13px; color:{CORES["texto"]}; margin-top:3px;">{" · ".join(desde)}</div></div>')
    corpo += '<div style="height:10px;"></div>' + linhas_de_narrativa_html(itens)
    if f["alertas"]:
        corpo += (f'<div style="font-size:10px; letter-spacing:1.4px; text-transform:uppercase; color:{CORES["apagado"]}; margin:16px 0 4px 0;">Alertas do motor do painel</div>'
                  + "".join(f'<div style="font-size:13px; line-height:1.5; color:{CORES["texto"]}; padding:6px 0; border-top:1px solid {CORES["borda"]};">'
                            f'<b style="color:{CORES["negativo"] if a["nivel"] == "critico" else CORES["alerta"]};">{a["titulo"]}</b> — {a["detalhe"]}</div>'
                            for a in f["alertas"][:6]))
    html = moldura_email("Briefing financeiro · caixa", f"{dia} · saldo de hoje (D+0)", status, cor, corpo, link, logo_src,
                         "Gerado automaticamente pelo painel financeiro a partir das planilhas de fluxo. Limites de alerta: os padrões do painel.")
    texto = (f"Briefing financeiro · {dia} · {status}\n\n" + "\n".join(f"{r}: {v} ({s})" for r, v, s in cartoes)
             + (f"\n\nDesde o último briefing: {' · '.join(desde)}" if desde else "") + "\n\n"
             + "\n".join(f"{i['rotulo']}: {i['texto'].replace('<b>', '').replace('</b>', '')}" for i in itens)
             + ("\n\nAlertas:\n" + "\n".join(f"- [{a['nivel']}] {a['titulo']}: {a['detalhe']}" for a in f["alertas"]) if f["alertas"] else ""))
    return html, texto


def main(argv):
    ns = carregar_funcoes_do_app(sementes=SEMENTES_FIN)
    hoje = datetime.now(ns["FUSO_BR"]).date()
    df, _origem = carregar_fluxo(ns)
    f = fatos_do_caixa(ns, df, hoje)
    fmt = ns.get("formata_valor_curto")
    logo_b64 = str(ns.get("LOGO_BEEA_B64") or "")
    logo_src = f"cid:{CID_LOGO}" if logo_b64 else ""
    link = os.environ.get("LINK_PAINEL", "")
    if "--alertas" in argv:
        # Modo alerta: linha de base na primeira execução; depois só o que for novo, uma vez por mês.
        chaves = [f"fin:{a['titulo']}:{f['mes']}" for a in f["alertas"] if a["nivel"] == "critico"]
        if not os.path.exists(CAMINHO_ESTADO_FIN):
            _gravar_json(CAMINHO_ESTADO_FIN, chaves)
            print(f"Linha de base dos alertas financeiros: {len(chaves)} ponto(s) registrados sem aviso.")
            return
        estado = _ler_json(CAMINHO_ESTADO_FIN, [])
        novos = [a for a in f["alertas"] if a["nivel"] == "critico" and f"fin:{a['titulo']}:{f['mes']}" not in estado]
        if not novos:
            print("Nada novo no caixa.")
            return
        linhas = "".join(f'<div style="font-family:{FONTE}; padding:10px 0; border-bottom:1px solid {CORES["borda"]};">'
                         f'<div style="font-size:15px; font-weight:700; color:{CORES["texto"]};">{a["titulo"]}</div>'
                         f'<div style="font-size:13px; color:{CORES["apagado"]}; margin-top:3px;">{a["detalhe"]}</div></div>' for a in novos)
        html = moldura_email(f"{len(novos)} ponto(s) de atenção no caixa", f"{hoje.strftime('%d/%m/%Y')}", "ALERTA", CORES["negativo"],
                             linhas, link, logo_src, "Cada ponto é avisado uma vez por mês. Limites: os padrões do painel financeiro.")
        texto = "\n".join(f"- {a['titulo']}: {a['detalhe']}" for a in novos)
        if "--teste" in argv:
            print(texto)
            return
        enviar_email(f"Alerta de caixa {hoje.strftime('%d/%m')} · {novos[0]['titulo']}", html, texto, logo_b64)
        _gravar_json(CAMINHO_ESTADO_FIN, estado + [f"fin:{a['titulo']}:{f['mes']}" for a in novos])
        print(f"Alerta de caixa enviado ({len(novos)}).")
        return
    historico = _ler_json(CAMINHO_HISTORICO_FIN, [])
    if any(h.get("data") == hoje.isoformat() for h in historico) and "--forcar" not in argv and "--teste" not in argv:
        print("O briefing financeiro de hoje já foi enviado (use --forcar para reenviar).")
        return
    anteriores = [h for h in historico if h.get("data", "") < hoje.isoformat()]
    anterior = anteriores[-1] if anteriores else None
    itens = narrativa_financeira(f, ns)
    desde = desde_ontem_fin(f, anterior, fmt)
    html, texto = montar_email_financeiro(f, itens, hoje, link, logo_src, ns, desde, (anterior or {}).get("data"))
    assunto = f"Briefing financeiro {hoje.strftime('%d/%m')} · saldo {fmt(f['saldo'])}" + (
        f" · {sum(1 for a in f['alertas'] if a['nivel'] == 'critico')} alerta(s) crítico(s)" if any(a["nivel"] == "critico" for a in f["alertas"]) else "")
    if "--teste" in argv:
        print(assunto)
        print(texto)
        return
    destinos = enviar_email(assunto, html, texto, logo_b64)
    print(f"Briefing financeiro enviado para {', '.join(destinos)}: {assunto}")
    entrada = {"data": hoje.isoformat(), "saldo": f["saldo"], "titulos": [a["titulo"] for a in f["alertas"]]}
    historico = [h for h in historico if h.get("data") != entrada["data"]] + [entrada]
    _gravar_json(CAMINHO_HISTORICO_FIN, historico[-400:])


if __name__ == "__main__":
    main(sys.argv[1:])
