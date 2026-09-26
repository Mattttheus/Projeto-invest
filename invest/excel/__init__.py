"""Monta a planilha: entradas -> cálculos -> Painel."""
import pandas as pd
from openpyxl import Workbook

from .abas import D0, aba_ativos, aba_carteira, aba_lancamentos, aba_premissas, aba_proventos, aba_verificar
from .estilo import ABAS, AZUL, CINZA_TXT, MODO, VERMELHO
from .dinamicas import aba_analises
from .painel import painel


VAGAS_GOOGLE = 15                    # linhas livres no catálogo para ticker novo (versão Google, sem regerar)


def construir(ativos, lanc, prov, cfg, avisos, pendentes=None, google=False):
    MODO["google"] = google
    try:
        return _construir(ativos, lanc, prov, cfg, avisos, pendentes, google)
    finally:
        MODO["google"] = False


def _construir(ativos, lanc, prov, cfg, avisos, pendentes, google):
    wb = Workbook()
    wb.remove(wb.active)
    if google:
        vagas = pd.DataFrame([{c: None for c in ativos.columns}] * VAGAS_GOOGLE)
        ativos = pd.concat([ativos, vagas], ignore_index=True)
        ativos["ticker"] = ativos["ticker"].fillna("")
    N = len(ativos)
    T = D0 + N                        # linha de total (ativos ocupam D0..T-1)

    aba_ativos(wb, ativos, T)
    _, LN = aba_lancamentos(wb, lanc)
    _, PN = aba_proventos(wb, prov)
    _, classes = aba_premissas(wb, cfg)
    wsC, sub = aba_carteira(wb, N, T, LN, PN)
    aba_verificar(wb, avisos)
    aba_analises(wb)
    from ..cadastro import aba_cadastro
    aba_cadastro(wb, f"Ativos!$A${D0}:$A${T - 1}", N, D0, pendentes)
    painel(wb, wsC, cfg, N, D0, T, sub, classes)
    wb["Ativos"].sheet_state = "hidden"      # base do catálogo; a lista visível é uma só: Carteira

    ordem = {nome: i for i, nome in enumerate(ABAS + ["Base Painel"])}
    wb._sheets.sort(key=lambda s: ordem[s.title])
    cores = {"Painel": AZUL, "Carteira": AZUL, "Análises": AZUL, "Cadastro": "16A34A",
             "Verificar": VERMELHO}
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = cores.get(ws.title, CINZA_TXT)
    wb.active = 0
    wb.properties.creator = "Gestão de Investimentos"
    return wb


def gerar(ativos, lanc, prov, cfg, avisos, saida, dinamicas=True, pendentes=None, macros=True):
    """Gera a planilha: openpyxl monta tudo; o Excel acrescenta tabelas dinâmicas, macros e botões (.xlsm).
    Retorna a lista de mensagens de status."""
    from ..cadastro import LANCAMENTO, PROVENTO, status_linha
    from .automacao import finalizar
    saida.parent.mkdir(parents=True, exist_ok=True)
    rascunho = saida.with_name(f"~rascunho_{saida.stem}.xlsx")
    construir(ativos, lanc, prov, cfg, avisos, pendentes).save(rascunho)
    if not (dinamicas or macros):
        rascunho.replace(saida.with_suffix(".xlsx"))
        return ["planilha sem etapa do Excel (.xlsx)"]
    botoes = {"lc": status_linha(LANCAMENTO) + 2, "pv": status_linha(PROVENTO) + 2}
    msgs, ok = finalizar(rascunho.resolve(), saida.resolve(), botoes, dinamicas, macros)
    if ok:
        rascunho.unlink(missing_ok=True)
    else:                                   # sem Excel: entrega a versão sem macros para não ficar sem planilha
        rascunho.replace(saida.with_suffix(".xlsx"))
        msgs.append(f"entregue sem macros: {saida.with_suffix('.xlsx').name}")
    return msgs
