"""Tudo o que a apresentação precisa para montar as planilhas, já carregado, validado e analisado."""
from dataclasses import dataclass, field

import pandas as pd

from .analise import parametros


@dataclass
class BaseDados:
    ativos: pd.DataFrame
    lanc: pd.DataFrame
    prov: pd.DataFrame
    cfg: dict
    avisos: list = field(default_factory=list)
    analise: dict = None

    def __post_init__(self):
        if self.analise is None:                     # sem histórico: abas de análise mostram só o aviso
            self.analise = {"vazio": True, "cfg": parametros(self.cfg), "avisos": []}
