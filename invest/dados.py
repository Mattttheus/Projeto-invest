"""Leitura e validação da base única (dados/)."""
import json

import pandas as pd

from .caminhos import DADOS


def ler_csv(nome, datas=()):
    df = pd.read_csv(DADOS / nome, sep=";", decimal=",", encoding="utf-8-sig", dtype={"ticker": str})
    df.columns = df.columns.str.strip()
    for c in df.select_dtypes("object"):
        df[c] = df[c].str.strip()
    for c in datas:
        df[c] = pd.to_datetime(df[c], format="%d/%m/%Y", errors="coerce")
    return df


def carregar():
    ativos = ler_csv("ativos.csv")
    ativos["ticker"] = ativos["ticker"].str.upper()
    ativos["tipo"] = ativos["tipo"].str.upper()
    lanc = ler_csv("lancamentos.csv", datas=["data"])
    lanc["ticker"] = lanc["ticker"].str.upper()
    lanc["operacao"] = lanc["operacao"].str.upper()
    lanc["custos"] = lanc["custos"].fillna(0)
    prov = ler_csv("proventos.csv", datas=["data"])
    prov["ticker"] = prov["ticker"].str.upper()
    prov["tipo"] = prov["tipo"].str.upper()
    cfg = json.loads((DADOS / "config.json").read_text(encoding="utf-8"))
    return ativos, lanc, prov, cfg


def validar(ativos, lanc, prov):
    """Retorna lista de (gravidade, ativo, mensagem). ERRO impede a geração."""
    out = []
    for t in ativos.loc[ativos["ticker"].duplicated(keep=False), "ticker"].unique():
        out.append(("ERRO", t, "Ticker cadastrado mais de uma vez em ativos.csv."))
    conhecidos = set(ativos["ticker"])
    for nome, df in (("lancamentos.csv", lanc), ("proventos.csv", prov)):
        for t in sorted(set(df["ticker"]) - conhecidos):
            out.append(("ERRO", t, f"Ticker em {nome} que não existe em ativos.csv."))
        for _, r in df[df.duplicated(keep="first")].iterrows():
            out.append(("AVISO", r["ticker"], f"Linha repetida em {nome} (mesma data, valores e quantidade) — confira se não foi lançada duas vezes."))
        sem_data = df[df["data"].isna()]["ticker"].tolist()
        if sem_data:
            out.append(("AVISO", ", ".join(sem_data), f"Registros sem data em {nome}"
                        + (" — não entram nos proventos de 12 meses." if nome == "proventos.csv" else ".")))
    ops = set(lanc["operacao"]) - {"COMPRA", "VENDA"}
    if ops:
        out.append(("ERRO", "-", f"Operação inválida em lancamentos.csv: {', '.join(ops)} (use COMPRA ou VENDA)."))
    if abs(ativos["peso_ideal"].sum() - 1) > 0.001:
        out.append(("AVISO", "-", f"Pesos ideais somam {ativos['peso_ideal'].sum():.0%} (deveriam somar 100%)."))

    # coerência dos números
    comp = lanc[lanc["operacao"] == "COMPRA"]
    g = comp.assign(total=comp["quantidade"] * comp["preco"]).groupby("ticker")
    pm = g["total"].sum() / g["quantidade"].sum()
    investido = g["total"].sum()
    for t, v in investido.items():
        if len(investido) > 2 and v > 0.5 * investido.sum():
            out.append(("AVISO", t, f"Concentra {v / investido.sum():.0%} de todo o valor investido (R$ {v:,.0f}). Confira a quantidade."))
    for _, a in ativos.iterrows():
        t = a["ticker"]
        if t in pm and a["preco"] > 0 and abs(a["preco"] / pm[t] - 1) > 0.5:
            out.append(("AVISO", t, f"Cotação (R$ {a['preco']:.2f}) difere mais de 50% do preço médio (R$ {pm[t]:.2f}). Confira ticker/cotação."))
        if a["preco"] > 0 and a["provento_anual_cota"] / a["preco"] > 0.25:
            out.append(("AVISO", t, f"Provento estimado dá DY de {a['provento_anual_cota'] / a['preco']:.0%} ao ano — improvável."))
    precos = ativos.set_index("ticker")["preco"]
    for _, p in prov.iterrows():
        preco = precos.get(p["ticker"], 0)
        if preco > 0 and p["valor_cota"] / preco > 0.25:
            out.append(("AVISO", p["ticker"], f"Provento de R$ {p['valor_cota']:.2f}/cota num pagamento só é mais de 25% da cotação — provável erro de digitação."))
    # mesma compra (data, qtd, preço) em tickers diferentes: sinal de cópia
    chave = comp.dropna(subset=["data"]).groupby(["data", "quantidade", "preco"])["ticker"].apply(list)
    for ts in chave[chave.str.len() > 1]:
        out.append(("AVISO", " / ".join(ts), "Compras idênticas (data, quantidade e preço) em ativos diferentes — pode ser cópia."))
    return out
