"""Histórico semanal de cotações e proventos (Yahoo Finance) — base da análise quantitativa.

Gravado em dados/historico.csv (gerado automaticamente, não edite): permite rodar a análise offline.
Inclui o Ibovespa (^BVSP) como referência de mercado. Se um ticker falhar, o histórico anterior é mantido.
"""
import pandas as pd

from ..config import DADOS
from ..dominio.analise import IBOV
from .yahoo import historico_semanal

ARQ = DADOS / "historico.csv"
COLUNAS = ["data", "ticker", "fechamento", "provento"]


def carregar():
    if not ARQ.exists():
        return pd.DataFrame(columns=COLUNAS)
    df = pd.read_csv(ARQ, sep=";", decimal=",", encoding="utf-8-sig", dtype={"ticker": str})
    df["data"] = pd.to_datetime(df["data"], format="%d/%m/%Y")
    return df


def atualizar(tickers, anos=5):
    antigo = carregar()
    partes, ok = [], 0
    for t in list(tickers) + [IBOV]:
        try:
            partes.append(historico_semanal(t, anos))
            ok += 1
        except Exception as e:  # rede fora, ticker novo sem negociação etc.
            print(f"  {t:7} histórico mantido (falha: {e})")
            partes.append(antigo[antigo["ticker"] == t])
    df = pd.concat([p for p in partes if len(p)], ignore_index=True).sort_values(["ticker", "data"])
    saida = df.assign(data=df["data"].dt.strftime("%d/%m/%Y"))
    saida.to_csv(ARQ, sep=";", decimal=",", encoding="utf-8-sig", index=False)
    print(f"  histórico semanal: {ok} de {len(tickers) + 1} séries atualizadas ({anos} anos).")
    return ok
