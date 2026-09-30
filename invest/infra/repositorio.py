"""Acesso à base única (dados/*.csv): leitura tipada e acréscimo de registros.

CSV no padrão brasileiro: separador ";", vírgula decimal, datas dd/mm/aaaa, UTF-8 com BOM (abre certo no Excel).
"""
import pandas as pd

from ..config import DADOS, ler_config

ATIVOS, LANCAMENTOS, PROVENTOS, METAS = "ativos.csv", "lancamentos.csv", "proventos.csv", "metas.csv"
FORMATO_DATA = "%d/%m/%Y"
NUMERICAS = {ATIVOS: ("preco", "provento_anual_cota", "peso_ideal"),
             LANCAMENTOS: ("quantidade", "preco", "custos"),
             PROVENTOS: ("valor_cota", "quantidade"),
             METAS: ("meta_mensal", "prazo_meses")}


def ler_csv(nome, datas=()):
    df = pd.read_csv(DADOS / nome, sep=";", decimal=",", encoding="utf-8-sig", dtype={"ticker": str})
    df.columns = df.columns.str.strip()
    for c in df.select_dtypes("object"):
        df[c] = df[c].str.strip()
    for c in datas:
        df[c] = pd.to_datetime(df[c], format=FORMATO_DATA, errors="coerce")
    for c in NUMERICAS.get(nome, ()):         # CSV vazio ou valor digitado errado não viram texto
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    return df


def _maiusculas(df, *colunas):
    for c in colunas:
        df[c] = df[c].str.upper()
    return df


def carregar():
    """(ativos, lançamentos, proventos, config) prontos para o domínio."""
    ativos = _maiusculas(ler_csv(ATIVOS), "ticker", "tipo")
    lanc = _maiusculas(ler_csv(LANCAMENTOS, datas=["data"]), "ticker", "operacao")
    lanc["custos"] = lanc["custos"].fillna(0)
    prov = _maiusculas(ler_csv(PROVENTOS, datas=["data"]), "ticker", "tipo")
    return ativos, lanc, prov, ler_config()


def carregar_metas():
    """Metas personalizadas por ativo (opcional: sem o arquivo, todos usam o padrão do config.json)."""
    from ..dominio.metas import vazias
    if not (DADOS / METAS).exists():
        return vazias()
    return _maiusculas(ler_csv(METAS), "ticker")


def salvar_metas(metas):
    """Reescreve metas.csv só com os ativos personalizados (campo vazio = padrão)."""
    linhas = ["ticker;meta_mensal;prazo_meses"]
    for m in metas:
        meta = "" if m["meta_mensal"] is None else br(m["meta_mensal"])
        prazo = "" if m["prazo_meses"] is None else str(m["prazo_meses"])
        linhas.append(f"{m['ticker']};{meta};{prazo}")
    (DADOS / METAS).write_text("\n".join(linhas) + "\n", encoding="utf-8-sig")


def tickers_cadastrados():
    return ler_csv(ATIVOS)["ticker"].str.upper().tolist()


def br(x, casas=2):
    """Número no formato do CSV (vírgula decimal)."""
    return f"{x:.{casas}f}".replace(".", ",")


def br_qtd(x):
    return br(x, 0) if float(x).is_integer() else br(x, 6)


def acrescentar(nome, linhas):
    """Acrescenta linhas (listas de textos já formatados) ao fim do CSV, sem reescrever o que existe."""
    if not linhas:
        return
    arq = DADOS / nome
    texto = arq.read_text(encoding="utf-8-sig")
    if texto and not texto.endswith("\n"):
        texto += "\n"
    texto += "".join(";".join(l) + "\n" for l in linhas)
    arq.write_text(texto, encoding="utf-8-sig")


def acrescentar_lancamentos(lancamentos):
    acrescentar(LANCAMENTOS, [[l.data.strftime(FORMATO_DATA), l.ticker, l.operacao, br_qtd(l.quantidade),
                               br(l.preco), br(l.custos)] for l in lancamentos])


def acrescentar_proventos(proventos):
    acrescentar(PROVENTOS, [[p.data.strftime(FORMATO_DATA), p.ticker, p.tipo, br(p.valor_cota, 4), br_qtd(p.quantidade)]
                            for p in proventos])


def cadastrar_ativos(novos):
    """novos: {ticker: (tipo, segmento)}. Cotação e proventos ficam 0 até a próxima atualização online."""
    acrescentar(ATIVOS, [[tk, tipo, seg or "-", "0", "0", "0", "manual"] for tk, (tipo, seg) in novos.items()])
