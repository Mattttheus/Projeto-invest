"""Monta a planilha: entradas -> cálculos -> análises -> Painel."""
from openpyxl import Workbook

from .abas import D0, aba_ativos, aba_carteira, aba_lancamentos, aba_premissas, aba_proventos, aba_verificar
from .cadastro import LANCAMENTO, PROVENTO, aba_cadastro, status_linha
from .dinamicas import aba_analises
from .estilo import ABAS, AZUL, CINZA_TXT, VERDE, VERMELHO, modo_google
from .meta import aba_meta
from .painel import painel
from .quantitativo import aba_modelos, aba_projecao, aba_quant

VAGAS_GOOGLE = 15                    # linhas livres no catálogo para ticker novo (versão Google, sem regerar)
CORES_ABAS = {"Painel": AZUL, "Carteira": AZUL, "Análises": AZUL, "Meta": AZUL, "Quant": AZUL, "Modelos": AZUL,
              "Projeção": AZUL, "Cadastro": VERDE, "Verificar": VERMELHO}


def construir(base, pendentes=None, google=False):
    """Workbook completo a partir de uma BaseDados. google=True: versão sem VBA, com vagas no catálogo."""
    with modo_google(google):
        return _construir(base, pendentes, google)


def _com_vagas(ativos):
    """Catálogo + linhas vazias para tickers novos lançados direto no Google Planilhas."""
    out = ativos.reindex(range(len(ativos) + VAGAS_GOOGLE))
    out["ticker"] = out["ticker"].fillna("")
    return out


def _construir(base, pendentes, google):
    wb = Workbook()
    wb.remove(wb.active)
    ativos = _com_vagas(base.ativos) if google else base.ativos
    N = len(ativos)
    T = D0 + N                        # linha de total (ativos ocupam D0..T-1)

    aba_ativos(wb, ativos, T)
    _, LN = aba_lancamentos(wb, base.lanc)
    _, PN = aba_proventos(wb, base.prov)
    _, classes = aba_premissas(wb, base.cfg)
    _, sub = aba_carteira(wb, N, T, LN, PN)
    aba_verificar(wb, base.avisos)
    aba_analises(wb)
    aba_quant(wb, ativos, base.analise, T)
    aba_modelos(wb, base.analise)
    aba_projecao(wb, base.analise, T)
    aba_meta(wb, N, T)
    aba_cadastro(wb, f"Ativos!$A${D0}:$A${T - 1}", N, D0, pendentes)
    painel(wb, base.cfg, N, D0, T, sub, classes)
    wb["Ativos"].sheet_state = "hidden"      # base do catálogo; a lista visível é uma só: Carteira

    ordem = {nome: i for i, nome in enumerate(ABAS + ["Base Painel"])}
    wb._sheets.sort(key=lambda s: ordem[s.title])
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = CORES_ABAS.get(ws.title, CINZA_TXT)
    wb.active = 0
    wb.properties.creator = "Gestão de Investimentos"
    return wb


def gerar(base, saida, dinamicas=True, macros=True, pendentes=None):
    """Gera a planilha: openpyxl monta tudo; o Excel acrescenta tabelas dinâmicas, macros e botões (.xlsm).
    Retorna a lista de mensagens de status."""
    from .automacao import finalizar
    saida.parent.mkdir(parents=True, exist_ok=True)
    rascunho = saida.with_name(f"~rascunho_{saida.stem}.xlsx")
    construir(base, pendentes).save(rascunho)
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
