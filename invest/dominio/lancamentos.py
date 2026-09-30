"""Regras para aceitar uma compra/venda ou um provento digitado no formulário do Cadastro.

Entrada: dicionários com os campos já lidos (texto ou número). Saída: registros normalizados ou ValueError
com a mensagem que volta para o formulário. Sem acesso a arquivos: quem grava é a camada de aplicação.
"""
import datetime as dt
from dataclasses import dataclass

from .carteira import OPERACOES, TIPOS_ATIVO, TIPOS_PROVENTO


@dataclass
class Lancamento:
    data: dt.date
    ticker: str
    operacao: str
    quantidade: float
    preco: float
    custos: float


@dataclass
class Provento:
    data: dt.date
    ticker: str
    tipo: str
    valor_cota: float
    quantidade: float


def texto(v):
    return str(v).strip().upper() if v not in (None, "") else ""


def numero(v):
    """Aceita número ou texto no formato brasileiro (R$ 1.234,56) ou com ponto decimal (1234.56)."""
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("R$", "").strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return float(s)


def data(v, hoje):
    if v in (None, ""):
        return hoje
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return dt.datetime.strptime(str(v).strip(), "%d/%m/%Y").date()


def validar_lancamento(d, conhecidos, novos, cotas, cotacao, hoje):
    """d: campos do formulário. novos: {ticker: (tipo, segmento)} cadastrados nesta importação (atualizado aqui).
    cotas: {ticker: quantidade} (atualizado aqui). cotacao(ticker) -> preço atual ou None."""
    tk = texto(d.get("Ticker"))
    op = texto(d.get("Operação")) or "COMPRA"
    qtd, custos = numero(d.get("Quantidade de cotas")), numero(d.get("Custos (R$)")) or 0
    if not tk:
        raise ValueError("informe o ticker")
    if op not in OPERACOES:
        raise ValueError("operação deve ser COMPRA ou VENDA")
    if not qtd or qtd <= 0:
        raise ValueError("informe a quantidade de cotas")
    if tk not in conhecidos and tk not in novos:
        tipo = texto(d.get("Tipo do ativo"))
        if tipo not in TIPOS_ATIVO:
            raise ValueError("ticker novo: informe o Tipo do ativo (AÇÃO ou FII)")
        novos[tk] = (tipo, str(d.get("Segmento") or "").strip())
    preco = numero(d.get("Preço por cota (R$)")) or cotacao(tk)
    if not preco or preco <= 0:
        raise ValueError("sem cotação para este ticker — informe o preço por cota")
    if op == "VENDA" and qtd > cotas.get(tk, 0):
        raise ValueError(f"venda maior que as cotas que você tem ({cotas.get(tk, 0):g})")
    cotas[tk] = cotas.get(tk, 0) + (-qtd if op == "VENDA" else qtd)
    return Lancamento(data(d.get("Data"), hoje), tk, op, qtd, preco, custos)


def validar_provento(d, conhecidos, novos, cotas, hoje):
    tk = texto(d.get("Ticker"))
    tipo = texto(d.get("Tipo de provento")) or "DIVIDENDO"
    valor = numero(d.get("Valor por cota (R$)"))
    qtd = numero(d.get("Qtd de cotas")) or cotas.get(tk, 0)
    if not tk or (tk not in conhecidos and tk not in novos):
        raise ValueError("ticker não cadastrado — registre a compra primeiro")
    if tipo not in TIPOS_PROVENTO:
        raise ValueError("tipo de provento inválido")
    if not valor or valor <= 0:
        raise ValueError("informe o valor por cota")
    if not qtd or qtd <= 0:
        raise ValueError("você não tem cotas deste ativo — informe a quantidade")
    return Provento(data(d.get("Data pagamento"), hoje), tk, tipo, valor, qtd)
