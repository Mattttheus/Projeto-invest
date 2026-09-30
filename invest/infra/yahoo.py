"""Cliente do Yahoo Finance (yfinance): único ponto do projeto que acessa a internet.

Só os tickers são enviados. Ações brasileiras usam o sufixo .SA; índices (^BVSP) vão como estão.
Para ações, o Yahoo informa JCP pelo valor bruto (antes dos 15% de IR).
"""
import pandas as pd


def _simbolo(ticker):
    return ticker if ticker.startswith("^") else f"{ticker}.SA"


def _sem_fuso(idx):
    return (idx.tz_localize(None) if idx.tz is not None else idx).normalize()


def cotacao_e_proventos_12m(ticker):
    """(último fechamento, soma dos proventos por cota nos últimos 12 meses, data do fechamento)."""
    import yfinance as yf

    k = yf.Ticker(_simbolo(ticker))
    fech = k.history(period="1mo")["Close"].dropna()
    if fech.empty:
        raise ValueError("sem cotação")
    div = k.dividends
    if len(div):
        inicio = pd.Timestamp.now(tz=div.index.tz) - pd.DateOffset(years=1)
        div = div[div.index >= inicio]
    return round(float(fech.iloc[-1]), 2), round(float(div.sum()), 4), fech.index[-1].date()


def historico_semanal(ticker, anos):
    """DataFrame [data, ticker, fechamento, provento]: fechamento semanal (sem ajuste de proventos)
    e proventos pagos em cada semana."""
    import yfinance as yf

    k = yf.Ticker(_simbolo(ticker))
    h = k.history(period=f"{anos}y", interval="1wk", auto_adjust=False)["Close"].dropna()
    if h.empty:
        raise ValueError("sem histórico")
    semanas = _sem_fuso(h.index)
    prov = pd.Series(0.0, index=semanas)
    div = k.dividends
    if len(div):
        pos = semanas.searchsorted(_sem_fuso(div.index), side="right") - 1     # semana em que o provento caiu
        for p, v in zip(pos, div.values):
            if p >= 0:
                prov.iloc[p] += float(v)
    return pd.DataFrame({"data": semanas, "ticker": ticker, "fechamento": h.values.round(4),
                         "provento": prov.values.round(6)})
