"""Atualiza cotação e proventos dos últimos 12 meses em dados/ativos.csv (fonte: Yahoo Finance).

Grava direto nas colunas preco, provento_anual_cota e atualizado_em — não cria cópia dos dados.
Se um ticker falhar, o valor anterior é mantido. Só os tickers são enviados ao Yahoo.
Para ações, o Yahoo informa JCP pelo valor bruto (antes dos 15% de IR).
"""
import pandas as pd

from .caminhos import DADOS

ARQ = DADOS / "ativos.csv"


def buscar(ticker):
    import yfinance as yf

    k = yf.Ticker(f"{ticker}.SA")
    fech = k.history(period="1mo")["Close"].dropna()
    if fech.empty:
        raise ValueError("sem cotação")
    div = k.dividends
    if len(div):
        inicio = pd.Timestamp.now(tz=div.index.tz) - pd.DateOffset(years=1)
        div = div[div.index >= inicio]
    return round(float(fech.iloc[-1]), 2), round(float(div.sum()), 4), fech.index[-1].date()


def atualizar():
    df = pd.read_csv(ARQ, sep=";", decimal=",", encoding="utf-8-sig", dtype={"ticker": str})
    df = df.drop(columns=["fonte"], errors="ignore")      # redundante: atualizado_em já indica a origem
    if "atualizado_em" not in df:
        df["atualizado_em"] = ""
    df["atualizado_em"] = df["atualizado_em"].astype("object")
    ok = 0
    for i, t in df["ticker"].items():
        try:
            preco, prov, dia = buscar(t.strip().upper())
        except Exception as e:  # rede fora, ticker inexistente etc.
            print(f"  {t:7} mantido (falha: {e})")
            continue
        antes = df.at[i, "preco"]
        df.at[i, "preco"], df.at[i, "provento_anual_cota"] = preco, prov
        df.at[i, "atualizado_em"] = dia.strftime("%d/%m/%Y")
        ok += 1
        print(f"  {t:7} R$ {antes:>8.2f} -> R$ {preco:>8.2f}   proventos 12m R$ {prov:.4f}")
    df.to_csv(ARQ, sep=";", decimal=",", encoding="utf-8-sig", index=False)
    print(f"  {ok} de {len(df)} ativos atualizados.")
    return ok
