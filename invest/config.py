"""Caminhos do projeto e leitura de dados/config.json (premissas do usuário)."""
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DADOS = BASE / "dados"          # entrada — única fonte de informação
SAIDA = BASE / "saida"          # planilha gerada
ARQUIVO_SAIDA_PADRAO = "Gestao de investimentos.xlsm"


def ler_config():
    return json.loads((DADOS / "config.json").read_text(encoding="utf-8"))


def caminho_saida(cfg, alternativo=None):
    return Path(alternativo) if alternativo else SAIDA / cfg.get("arquivo_saida", ARQUIVO_SAIDA_PADRAO)
