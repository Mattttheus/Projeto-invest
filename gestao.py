"""Gestão de investimentos — comando único.

    python gestao.py            importa o Cadastro, atualiza cotações/proventos e gera a planilha
    python gestao.py --offline  gera sem acessar a internet
    python gestao.py --google   também gera a versão Google Planilhas (pasta google/)
    python gestao.py --google --drive   ...e copia para G:\Meu Drive\InvestERP

Entrada: dados/ (única fonte) + aba Cadastro da planilha.  Saída: saida/Gestao de investimentos.xlsm
(com macros: botões Registrar, Limpar, Novo lançamento e Atualizar cotações).
"""
import argparse
import json
import os
import sys
import time
from pathlib import Path

sys.dont_write_bytecode = True     # sem pastas __pycache__ no projeto

from invest import excel
from invest.cadastro import importar
from invest.caminhos import DADOS, SAIDA
from invest.dados import carregar, validar


def aberta(arquivo: Path):
    """True se a planilha estiver aberta no Excel (arquivo bloqueado para escrita)."""
    if not arquivo.exists():
        return False
    try:
        with open(arquivo, "r+b"):
            return False
    except PermissionError:
        return True


def main():
    p = argparse.ArgumentParser(description="Gera a planilha de investimentos a partir de dados/.")
    p.add_argument("--offline", action="store_true", help="não busca cotações na internet")
    p.add_argument("--saida", help="caminho alternativo para a planilha")
    p.add_argument("--sem-dinamicas", action="store_true", help="não cria as tabelas dinâmicas")
    p.add_argument("--sem-macros", action="store_true", help="não insere macros/botões")
    p.add_argument("--esperar", action="store_true", help="espera a planilha ser fechada (usado pelo botão Atualizar)")
    p.add_argument("--abrir", action="store_true", help="abre a planilha no Excel ao terminar")
    p.add_argument("--google", action="store_true", help="também gera a versão Google Planilhas em google/")
    p.add_argument("--drive", nargs="?", const="padrao", help="copia a versão Google para o Google Drive")
    a = p.parse_args()

    cfg = json.loads((DADOS / "config.json").read_text(encoding="utf-8"))
    saida = Path(a.saida) if a.saida else SAIDA / cfg.get("arquivo_saida", "Gestao de investimentos.xlsm")
    anterior = saida if saida.exists() else saida.with_suffix(".xlsx")   # transição da versão sem macros
    if a.esperar:
        for _ in range(60):
            if not (aberta(saida) or aberta(anterior)):
                break
            time.sleep(1)
    for arq in {saida, anterior}:   # antes de importar: evita gravar o Cadastro e não conseguir limpar o formulário
        if aberta(arq):
            sys.exit(f"A planilha está aberta no Excel. Salve, feche '{arq.name}' e rode de novo.")

    print("1. Importando a aba Cadastro...")
    resumo, pendentes = importar(anterior)
    print(f"  {resumo}")

    if not a.offline:
        print("2. Atualizando cotações e proventos (Yahoo Finance)...")
        try:
            from invest.mercado import atualizar
            atualizar()
        except ImportError:
            print("  yfinance não instalado — rode: pip install -r requirements.txt")

    print("3. Conferindo os dados...")
    ativos, lanc, prov, cfg = carregar()
    avisos = validar(ativos, lanc, prov)
    for g, t, m in avisos:
        print(f"  [{g}] {t}: {m}")
    if any(g == "ERRO" for g, _, _ in avisos):
        sys.exit("Corrija os erros acima em dados/ e rode de novo.")

    try:
        msgs = excel.gerar(ativos, lanc, prov, cfg, avisos, saida, dinamicas=not a.sem_dinamicas,
                           pendentes=pendentes, macros=not a.sem_macros)
    except PermissionError:
        sys.exit(f"\nA planilha está aberta no Excel. Feche '{saida.name}' e rode de novo.")
    print(f"4. Planilha gerada: {saida}")
    for m in msgs:
        print(f"   {m}")
    if saida.suffix == ".xlsm" and saida.exists() and anterior != saida and anterior.exists():
        anterior.unlink()                                 # versão antiga sem macros (dados já importados)
    if a.google or a.drive:
        from invest import google
        destino = None
        if a.drive:
            destino = google.DRIVE_PADRAO if a.drive == "padrao" else Path(a.drive)
            if not destino.parent.exists():
                sys.exit(f"Pasta do Google Drive não encontrada: {destino.parent}")
        for m in google.gerar(ativos, lanc, prov, cfg, avisos, destino):
            print(f"5. {m}")
    if a.abrir and saida.exists():
        os.startfile(saida)


if __name__ == "__main__":
    main()
