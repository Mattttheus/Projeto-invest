"""Carteira a partir dos valores internos: posições, fluxos de caixa e taxa interna de retorno."""
import numpy as np
import pandas as pd

OPERACOES = ("COMPRA", "VENDA")
TIPOS_ATIVO = ("AÇÃO", "FII")
TIPOS_PROVENTO = ("DIVIDENDO", "JCP", "RENDIMENTO", "AMORTIZAÇÃO")


def sinal(operacao):
    return -1 if operacao == "VENDA" else 1


def posicoes(ativos, lanc):
    """Quantidade atual de cada ticker do cadastro (compras - vendas)."""
    if lanc.empty:
        return pd.Series(0.0, index=ativos["ticker"])
    q = (lanc["quantidade"] * lanc["operacao"].map(sinal)).groupby(lanc["ticker"]).sum()
    return q.reindex(ativos["ticker"]).fillna(0.0)


def valor_posicoes(ativos, qtd):
    precos = ativos.set_index("ticker")["preco"]
    return (qtd * precos.reindex(qtd.index)).clip(lower=0)


def fluxos_internos(ativos, lanc, prov, qtd, hoje):
    """Compras (-), vendas (+), proventos (+) e o valor da posição em `hoje` (+)."""
    f = []
    for x in lanc.itertuples():
        bruto = x.quantidade * x.preco
        f.append((x.data, -(bruto + x.custos) if x.operacao == "COMPRA" else bruto - x.custos))
    for x in prov.itertuples():
        f.append((x.data, x.valor_cota * x.quantidade))
    f.append((hoje, float(valor_posicoes(ativos, qtd).sum())))
    return f


def tir(fluxos, iteracoes=200):
    """Taxa interna de retorno anual de [(data, valor)] por bissecção. None se não houver solução."""
    fluxos = [(d, v) for d, v in fluxos if pd.notna(d) and v]
    if len(fluxos) < 2 or all(v > 0 for _, v in fluxos) or all(v < 0 for _, v in fluxos):
        return None
    t0 = min(d for d, _ in fluxos)
    anos = np.array([(d - t0).days / 365.25 for d, _ in fluxos])
    val = np.array([v for _, v in fluxos])

    def vpl(taxa):
        return (val / (1 + taxa) ** anos).sum()

    lo, hi = -0.99, 10.0
    if vpl(lo) * vpl(hi) > 0:
        return None
    for _ in range(iteracoes):
        meio = (lo + hi) / 2
        if vpl(lo) * vpl(meio) <= 0:
            hi = meio
        else:
            lo = meio
    return (lo + hi) / 2
