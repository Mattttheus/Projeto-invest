"""Tudo o que a apresentação precisa para montar as planilhas, já carregado, validado e analisado."""
from dataclasses import dataclass, field

import pandas as pd

from .analise import parametros
from .metas import resolver, vazias


@dataclass
class BaseDados:
    ativos: pd.DataFrame
    lanc: pd.DataFrame
    prov: pd.DataFrame
    cfg: dict
    avisos: list = field(default_factory=list)
    analise: dict = None
    metas: pd.DataFrame = None             # meta e prazo efetivos por ticker (dominio/metas.resolver)

    def __post_init__(self):
        if self.metas is None:
            self.metas = resolver(self.ativos, vazias(), self.cfg)
        if self.analise is None:                     # sem histórico: abas de análise mostram só o aviso
            self.analise = {"vazio": True, "cfg": parametros(self.cfg), "avisos": []}
