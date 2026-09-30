"""Casos de uso encadeados pelo comando gestao.py: atualizar o mercado e conferir/analisar a base."""
from ..dominio import metas as regras_metas
from ..dominio.analise import analisar
from ..dominio.base import BaseDados
from ..dominio.validacao import validar
from ..infra import historico, repositorio


def atualizar_mercado(cfg):
    """Cotações e proventos 12m (ativos.csv) + histórico semanal (historico.csv). Exige internet."""
    from ..infra import mercado
    mercado.atualizar()
    historico.atualizar(repositorio.tickers_cadastrados(), cfg.get("analise", {}).get("historico_anos", 5))


def conferir_e_analisar(hoje=None):
    """Carrega dados/, valida e roda a análise quantitativa. Os avisos da análise entram na conferência."""
    ativos, lanc, prov, cfg = repositorio.carregar()
    personalizadas = repositorio.carregar_metas()
    metas = regras_metas.resolver(ativos, personalizadas, cfg)
    avisos = validar(ativos, lanc, prov) + regras_metas.validar(personalizadas, ativos)
    resultado = analisar(ativos, lanc, prov, cfg, historico.carregar(), hoje, metas)
    return BaseDados(ativos, lanc, prov, cfg, avisos + resultado["avisos"], resultado, metas)
