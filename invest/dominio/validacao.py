"""Checagens de consistência da base (cadastro, lançamentos e proventos)."""
from .avisos import AVISO, ERRO, Aviso
from .carteira import OPERACOES

LIMITE_CONCENTRACAO = 0.5     # fração do valor investido num só ativo
LIMITE_DESVIO_PRECO = 0.5     # cotação x preço médio
LIMITE_DY = 0.25              # dividend yield anual improvável
TOLERANCIA_PESOS = 0.001


def validar(ativos, lanc, prov):
    """Lista de Aviso(gravidade, ativo, mensagem). ERRO impede a geração da planilha."""
    return (_numeros(ativos, lanc, prov) + _cadastro(ativos) + _referencias(ativos, lanc, prov) + _operacoes(lanc)
            + _coerencia_precos(ativos, lanc, prov) + _copias(lanc))


def _numeros(ativos, lanc, prov):
    """Valor que não é número (ex.: '1.234,5' com ponto de milhar) chega vazio do repositório."""
    out = []
    for nome, df, colunas in (("ativos.csv", ativos, ("preco", "provento_anual_cota", "peso_ideal")),
                              ("lancamentos.csv", lanc, ("quantidade", "preco")),
                              ("proventos.csv", prov, ("valor_cota", "quantidade"))):
        for c in colunas:
            for t in df.loc[df[c].isna(), "ticker"]:
                out.append(Aviso(ERRO, t, f"Valor vazio ou não numérico na coluna '{c}' de {nome} "
                                          "(use vírgula decimal e sem ponto de milhar)."))
    return out


def _cadastro(ativos):
    out = [Aviso(ERRO, t, "Ticker cadastrado mais de uma vez em ativos.csv.")
           for t in ativos.loc[ativos["ticker"].duplicated(keep=False), "ticker"].unique()]
    soma = ativos["peso_ideal"].sum()
    if abs(soma - 1) > TOLERANCIA_PESOS:
        out.append(Aviso(AVISO, "-", f"Pesos ideais somam {soma:.0%} (deveriam somar 100%)."))
    return out


def _referencias(ativos, lanc, prov):
    out = []
    conhecidos = set(ativos["ticker"])
    for nome, df in (("lancamentos.csv", lanc), ("proventos.csv", prov)):
        for t in sorted(set(df["ticker"].dropna()) - conhecidos):
            out.append(Aviso(ERRO, t, f"Ticker em {nome} que não existe em ativos.csv."))
        for _, r in df[df.duplicated(keep="first")].iterrows():
            out.append(Aviso(AVISO, r["ticker"], f"Linha repetida em {nome} (mesma data, valores e quantidade) — "
                                                 "confira se não foi lançada duas vezes."))
        sem_data = df.loc[df["data"].isna(), "ticker"].astype(str).tolist()
        if sem_data:
            out.append(Aviso(AVISO, ", ".join(sem_data), f"Registros sem data em {nome}"
                             + (" — não entram nos proventos de 12 meses." if nome == "proventos.csv" else ".")))
    return out


def _operacoes(lanc):
    out = []
    invalidas = set(lanc["operacao"].astype(str)) - set(OPERACOES)
    if invalidas:
        out.append(Aviso(ERRO, "-", f"Operação inválida em lancamentos.csv: {', '.join(sorted(invalidas))} "
                                    "(use COMPRA ou VENDA)."))
    for _, r in lanc[(lanc["quantidade"] <= 0) | (lanc["preco"] <= 0)].iterrows():
        out.append(Aviso(AVISO, r["ticker"], "Lançamento com quantidade ou preço zerado/negativo — confira a linha."))
    # posição acumulada negativa em algum momento: venda maior do que as cotas que havia
    validos = lanc[lanc["operacao"].isin(OPERACOES)].sort_values("data", kind="stable")
    sinal = validos["operacao"].map({"COMPRA": 1, "VENDA": -1})
    acumulado = (validos["quantidade"] * sinal).groupby(validos["ticker"]).cumsum()
    for t in sorted(validos.loc[acumulado < -1e-9, "ticker"].unique()):
        out.append(Aviso(AVISO, t, "Venda maior do que as cotas que você tinha na data — confira os lançamentos."))
    return out


def _coerencia_precos(ativos, lanc, prov):
    out = []
    comp = lanc[lanc["operacao"] == "COMPRA"]
    g = comp.assign(total=comp["quantidade"] * comp["preco"]).groupby("ticker")
    investido = g["total"].sum()
    pm = investido / g["quantidade"].sum()
    for t, v in investido.items():
        if len(investido) > 2 and v > LIMITE_CONCENTRACAO * investido.sum():
            out.append(Aviso(AVISO, t, f"Concentra {v / investido.sum():.0%} de todo o valor investido "
                                       f"(R$ {v:,.0f}). Confira a quantidade."))
    for _, a in ativos.iterrows():
        t, preco = a["ticker"], a["preco"]
        if t in pm and preco > 0 and abs(preco / pm[t] - 1) > LIMITE_DESVIO_PRECO:
            out.append(Aviso(AVISO, t, f"Cotação (R$ {preco:.2f}) difere mais de 50% do preço médio "
                                       f"(R$ {pm[t]:.2f}). Confira ticker/cotação."))
        if preco > 0 and a["provento_anual_cota"] / preco > LIMITE_DY:
            out.append(Aviso(AVISO, t, f"Provento estimado dá DY de {a['provento_anual_cota'] / preco:.0%} ao ano — improvável."))
    precos = ativos.set_index("ticker")["preco"]
    for _, p in prov.iterrows():
        preco = precos.get(p["ticker"], 0)
        if preco > 0 and p["valor_cota"] / preco > LIMITE_DY:
            out.append(Aviso(AVISO, p["ticker"], f"Provento de R$ {p['valor_cota']:.2f}/cota num pagamento só é mais "
                                                 "de 25% da cotação — provável erro de digitação."))
    return out


def _copias(lanc):
    """Mesma compra (data, quantidade e preço) em tickers diferentes: sinal de linha copiada."""
    comp = lanc[lanc["operacao"] == "COMPRA"].dropna(subset=["data"])
    grupos = comp.groupby(["data", "quantidade", "preco"])["ticker"].apply(list)
    return [Aviso(AVISO, " / ".join(ts), "Compras idênticas (data, quantidade e preço) em ativos diferentes — pode ser cópia.")
            for ts in grupos if len(ts) > 1]
