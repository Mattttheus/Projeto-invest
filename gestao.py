r"""Gestão de investimentos — comando único (raiz de composição: liga as camadas do pacote invest).

    python gestao.py            importa o Cadastro, atualiza cotações/proventos e gera a planilha
    python gestao.py --offline  gera sem acessar a internet
    python gestao.py --google   também gera a versão Google Planilhas (pasta google/)
    python gestao.py --google --drive   ...e copia para G:\Meu Drive\InvestERP

Entrada: dados/ (única fonte) + aba Cadastro da planilha.  Saída: saida/Gestao de investimentos.xlsm
(com macros: botões Registrar, Limpar, Novo lançamento e Atualizar cotações).

Camadas (invest/): dominio (regras puras) <- infra (arquivos, Yahoo) <- aplicacao (casos de uso)
<- apresentacao (Excel, Google). Só este arquivo conhece todas.
"""
import argparse
import os
import sys
import time
from pathlib import Path

sys.dont_write_bytecode = True     # sem pastas __pycache__ no projeto

from invest.aplicacao.importacao import importar
from invest.aplicacao.pipeline import atualizar_mercado, conferir_e_analisar
from invest.apresentacao import excel
from invest.apresentacao.excel.cadastro import ler_formulario_preenchido, para_formulario
from invest.config import caminho_saida, ler_config
from invest.dominio.avisos import tem_erro

ESPERA_FECHAR = 60                 # segundos aguardando a planilha ser fechada (botão Atualizar)


def argumentos():
    p = argparse.ArgumentParser(description="Gera a planilha de investimentos a partir de dados/.")
    p.add_argument("--offline", action="store_true", help="não busca cotações na internet")
    p.add_argument("--saida", help="caminho alternativo para a planilha")
    p.add_argument("--sem-dinamicas", action="store_true", help="não cria as tabelas dinâmicas")
    p.add_argument("--sem-macros", action="store_true", help="não insere macros/botões")
    p.add_argument("--esperar", action="store_true", help="espera a planilha ser fechada (usado pelo botão Atualizar)")
    p.add_argument("--abrir", action="store_true", help="abre a planilha no Excel ao terminar")
    p.add_argument("--google", action="store_true", help="também gera a versão Google Planilhas em google/")
    p.add_argument("--drive", nargs="?", const="padrao", help="copia a versão Google para o Google Drive")
    return p.parse_args()


def aberta(arquivo: Path):
    """True se a planilha estiver aberta no Excel (arquivo bloqueado para escrita)."""
    if not arquivo.exists():
        return False
    try:
        with open(arquivo, "r+b"):
            return False
    except PermissionError:
        return True


def garantir_fechada(arquivos, esperar):
    if esperar:
        for _ in range(ESPERA_FECHAR):
            if not any(aberta(a) for a in arquivos):
                break
            time.sleep(1)
    for arq in arquivos:   # antes de importar: evita gravar o Cadastro e não conseguir limpar o formulário
        if aberta(arq):
            sys.exit(f"A planilha está aberta no Excel. Salve, feche '{arq.name}' e rode de novo.")


def importar_cadastro(planilha):
    """Lê o formulário da planilha anterior e grava o que for válido; o resto volta ao formulário."""
    compras, proventos, mensagem = ler_formulario_preenchido(planilha)
    if mensagem:
        return mensagem, para_formulario([], [])
    resumo, pend_compras, pend_proventos = importar(compras, proventos)
    return resumo, para_formulario(pend_compras, pend_proventos)


def atualizar_online(cfg):
    try:
        atualizar_mercado(cfg)
    except ImportError:
        print("  yfinance não instalado — rode: pip install -r requirements.txt")


def conferir():
    base = conferir_e_analisar()
    for g, t, m in base.avisos:
        print(f"  [{g}] {t}: {m}")
    if tem_erro(base.avisos):
        sys.exit("Corrija os erros acima em dados/ e rode de novo.")
    return base


def gerar_google(base, drive):
    from invest.apresentacao import google
    destino = None
    if drive:
        destino = google.DRIVE_PADRAO if drive == "padrao" else Path(drive)
        if not destino.parent.exists():
            sys.exit(f"Pasta do Google Drive não encontrada: {destino.parent}")
    for m in google.gerar(base, destino):
        print(f"5. {m}")


def main():
    a = argumentos()
    cfg = ler_config()
    saida = caminho_saida(cfg, a.saida)
    anterior = saida if saida.exists() else saida.with_suffix(".xlsx")   # transição da versão sem macros
    garantir_fechada({saida, anterior}, a.esperar)

    print("1. Importando a aba Cadastro...")
    resumo, pendentes = importar_cadastro(anterior)
    print(f"  {resumo}")
    if not a.offline:
        print("2. Atualizando cotações, proventos e histórico (Yahoo Finance)...")
        atualizar_online(cfg)
    print("3. Conferindo e analisando os dados...")
    base = conferir()

    try:
        msgs = excel.gerar(base, saida, dinamicas=not a.sem_dinamicas, macros=not a.sem_macros, pendentes=pendentes)
    except PermissionError:
        sys.exit(f"\nA planilha está aberta no Excel. Feche '{saida.name}' e rode de novo.")
    print(f"4. Planilha gerada: {saida}")
    for m in msgs:
        print(f"   {m}")
    if saida.suffix == ".xlsm" and saida.exists() and anterior != saida and anterior.exists():
        anterior.unlink()                                 # versão antiga sem macros (dados já importados)
    if a.google or a.drive:
        gerar_google(base, a.drive)
    if a.abrir and saida.exists():
        os.startfile(saida)


if __name__ == "__main__":
    main()
