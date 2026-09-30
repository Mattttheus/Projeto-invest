"""Avisos de qualidade dos dados: (gravidade, ativo, mensagem). ERRO impede gerar a planilha."""
from typing import NamedTuple

ERRO, AVISO, INFO = "ERRO", "AVISO", "INFO"


class Aviso(NamedTuple):
    gravidade: str
    ativo: str
    mensagem: str


def tem_erro(avisos):
    return any(a[0] == ERRO for a in avisos)
