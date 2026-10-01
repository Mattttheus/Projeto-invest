"""Gera a demonstração pública (GitHub Pages) em docs/ a partir de dados SIMULADOS (demo/dados/).

    python demo/gerar_demo.py

Nunca usa dados/ (seus dados reais). Cotações e histórico são públicos (Yahoo Finance); lançamentos,
proventos e outras classes são fictícios. A página fica em modo "demo": nada é gravado, as alterações
feitas no navegador ficam só nele. Também gera a planilha de exemplo (.xlsx, sem macros) para download.
"""
import os
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
os.environ["INVEST_DADOS"] = str(RAIZ / "demo" / "dados")       # antes de importar o pacote invest
sys.path.insert(0, str(RAIZ))
sys.dont_write_bytecode = True

from invest.aplicacao.pipeline import conferir_e_analisar      # noqa: E402
from invest.apresentacao import excel, web                       # noqa: E402
from invest.infra import historico                               # noqa: E402

DOCS = RAIZ / "docs"
PLANILHA = "Gestao de investimentos - demonstracao.xlsx"


def main():
    hist = historico.carregar()
    hoje = hist["data"].max() if not hist.empty else None        # data fixa: demo reproduzível
    base = conferir_e_analisar(hoje)
    for g, t, m in base.avisos:
        print(f"  [{g}] {t}: {m}")
    DOCS.mkdir(exist_ok=True)
    planilha = DOCS / PLANILHA
    excel.construir(base).save(planilha)
    for m in web.gerar(base, DOCS, "demo", hist, planilha):
        print(m)
    (DOCS / ".nojekyll").touch()                                 # GitHub Pages serve os arquivos como estão


if __name__ == "__main__":
    main()
