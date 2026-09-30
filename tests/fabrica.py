"""Dados sintéticos para os testes (nada de internet nem dos arquivos reais em dados/)."""
import numpy as np
import pandas as pd

HOJE = pd.Timestamp("2026-09-29")


def ativos(**extra):
    df = pd.DataFrame({"ticker": ["AAAA3", "BBBB11", "CCCC4"], "tipo": ["AÇÃO", "FII", "AÇÃO"],
                       "segmento": ["X", "Y", "Z"], "preco": [10.0, 100.0, 20.0],
                       "provento_anual_cota": [0.8, 10.0, 1.0], "peso_ideal": [0.4, 0.3, 0.3],
                       "atualizado_em": ["29/09/2026"] * 3})
    return df.assign(**extra)


def lancamentos(linhas=()):
    colunas = ["data", "ticker", "operacao", "quantidade", "preco", "custos"]
    df = pd.DataFrame(list(linhas), columns=colunas)
    df["data"] = pd.to_datetime(df["data"], format="%d/%m/%Y")
    for c in ("quantidade", "preco", "custos"):
        df[c] = df[c].astype(float)
    return df


def proventos(linhas=()):
    df = pd.DataFrame(list(linhas), columns=["data", "ticker", "tipo", "valor_cota", "quantidade"])
    df["data"] = pd.to_datetime(df["data"], format="%d/%m/%Y")
    for c in ("valor_cota", "quantidade"):
        df[c] = df[c].astype(float)
    return df


def historico(tickers=("AAAA3", "BBBB11", "CCCC4", "^BVSP"), semanas=260, semente=1):
    """Preços semanais em passeio aleatório com tendência, proventos trimestrais."""
    rng = np.random.default_rng(semente)
    datas = pd.date_range(end=HOJE, periods=semanas, freq="W-MON")
    partes = []
    for i, t in enumerate(tickers):
        r = rng.normal(0.002 + 0.001 * i, 0.03, semanas)
        preco = 20 * np.exp(np.cumsum(r))
        prov = np.where(np.arange(semanas) % 13 == 0, preco * 0.02, 0.0) if not t.startswith("^") else 0.0
        partes.append(pd.DataFrame({"data": datas, "ticker": t, "fechamento": preco, "provento": prov}))
    if not partes:
        return pd.DataFrame(columns=["data", "ticker", "fechamento", "provento"])
    return pd.concat(partes, ignore_index=True)


CFG = {"meta_mensal_por_ativo": 1000, "aporte_mensal": 3000, "peso_ideal_renda_variavel": 0.32,
       "outras_classes": [{"classe": "Renda fixa", "valor": 1000, "peso_ideal": 0.5}],
       "analise": {"cdi_anual": 0.14, "horizonte_anos": 3, "simulacoes_monte_carlo": 500}}
