"""Caso de uso: importar compras/proventos preenchidos no Cadastro da planilha anterior (uso sem macros).

Recebe os campos já lidos do formulário, valida cada registro (domínio) e acrescenta aos CSV (infraestrutura).
O que não passar volta com a observação do erro, para reaparecer no formulário.
"""
import datetime as dt

from ..dominio import metas as regras_metas
from ..dominio.carteira import sinal
from ..dominio.lancamentos import validar_lancamento, validar_provento
from ..infra import repositorio

def _cotacao_atual(ativos):
    """Cotação do cadastro; ticker novo (sem cotação) é buscado no Yahoo."""
    precos = ativos.set_index("ticker")["preco"]

    def cotacao(tk):
        if precos.get(tk, 0) > 0:
            return float(precos[tk])
        try:
            from ..infra.yahoo import cotacao_e_proventos_12m
            return cotacao_e_proventos_12m(tk)[0]
        except Exception:  # sem internet ou ticker inexistente: o usuário informa o preço
            return None
    return cotacao


def importar(compras, proventos, hoje=None):
    """compras/proventos: listas de dicionários com os campos do formulário.
    Retorna (resumo, compras pendentes, proventos pendentes)."""
    hoje = hoje or dt.date.today()
    ativos, lanc, _, _ = repositorio.carregar()
    conhecidos = set(ativos["ticker"])
    cotas = (lanc["quantidade"] * lanc["operacao"].map(sinal)).groupby(lanc["ticker"]).sum().to_dict()
    cotacao = _cotacao_atual(ativos)
    novos, lanc_ok, prov_ok, pend_c, pend_p = {}, [], [], [], []

    for d in compras:
        try:
            lanc_ok.append(validar_lancamento(d, conhecidos, novos, cotas, cotacao, hoje))
        except (ValueError, TypeError) as e:
            pend_c.append({**d, "Observação": f"Não importado: {e}"})
    for d in proventos:
        try:
            prov_ok.append(validar_provento(d, conhecidos, novos, cotas, hoje))
        except (ValueError, TypeError) as e:
            pend_p.append({**d, "Observação": f"Não importado: {e}"})

    repositorio.cadastrar_ativos(novos)
    repositorio.acrescentar_lancamentos(lanc_ok)
    repositorio.acrescentar_proventos(prov_ok)
    return _resumo(lanc_ok, prov_ok, novos, pend_c + pend_p), pend_c, pend_p


def _resumo(lanc_ok, prov_ok, novos, erros):
    partes = [f"{len(lanc_ok)} lançamento(s)", f"{len(prov_ok)} provento(s)"]
    if novos:
        partes.append("ticker(s) novo(s): " + ", ".join(novos))
    if erros:
        partes.append(f"{len(erros)} com erro: " + " | ".join(f"{d.get('Ticker') or '?'} — {d['Observação']}" for d in erros))
    return "importado: " + " • ".join(partes)


def atualizar_metas(entradas):
    """entradas: [(ticker, meta digitada, prazo digitado)] lidas da aba Meta. Grava em metas.csv só os
    personalizados; vazio ou fórmula = padrão. Retorna (resumo, erros)."""
    salvas, erros = [], []
    for ticker, meta, prazo in entradas:
        try:
            m = regras_metas.entrada(ticker, meta, prazo)
        except (ValueError, TypeError) as e:
            erros.append(str(e))
            continue
        if m["meta_mensal"] is not None or m["prazo_meses"] is not None:
            salvas.append(m)
    repositorio.salvar_metas(salvas)
    resumo = f"{len(salvas)} meta(s) personalizada(s)" + (f" • com erro: {'; '.join(erros)}" if erros else "")
    return resumo, erros
