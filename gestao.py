"""Gestão de investimentos — comando único (raiz de composição: liga as camadas do pacote invest).

    python gestao.py            importa o Cadastro, atualiza cotações/proventos e gera a planilha
    python gestao.py --offline  gera sem acessar a internet
    python gestao.py --google   também gera a versão Google Planilhas (pasta google/)
    python gestao.py --google --drive   ...e copia para G:\Meu Drive\InvestERP
    python gestao.py --web      também gera a página web em web/ (WAMP: http://localhost:8081)

Entrada: dados/ (única fonte) + aba Cadastro da planilha.  Saída: saida/Gestao de investimentos.xlsm
(com macros: botões Registrar, Limpar, Novo lançamento e Atualizar cotações).

Camadas (invest/): dominio (regras puras) <- infra (arquivos, Yahoo) <- aplicacao (casos de uso)
<- apresentacao (Excel, Google). Só este arquivo conhece todas.
"""
import argparse
import json
import os
import sys
import time
from pathlib import Path

sys.dont_write_bytecode = True     # sem pastas __pycache__ no projeto

from invest.aplicacao.importacao import atualizar_metas, importar
from invest.aplicacao.pipeline import atualizar_mercado, conferir_e_analisar
from invest.apresentacao import excel
from invest.apresentacao.excel.cadastro import ler_formulario_preenchido, para_formulario
from invest.apresentacao.excel.meta import ler_metas_preenchidas
from invest.config import BASE, caminho_saida, ler_config
from invest.dominio.avisos import tem_erro

ESPERA_FECHAR = 60                 # segundos aguardando a planilha ser fechada (botão Atualizar)
WEB = BASE / "web"                 # página servida pelo WAMP (dados reais: só localhost e rede local)


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
    p.add_argument("--web", action="store_true", help="também gera a página web em web/")
    p.add_argument("--web-atualizar", action="store_true", help="atualiza só a página web (usado pelo api.php)")
    p.add_argument("--web-registrar", metavar="JSON", help="grava lançamentos/proventos/metas vindos da página web")
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


def importar_metas(planilha):
    """Metas e prazos digitados na aba Meta -> dados/metas.csv. Planilha sem a aba nova: nada muda."""
    entradas = ler_metas_preenchidas(planilha)
    if not entradas:
        return "aba Meta sem campos editáveis — metas mantidas"
    return atualizar_metas(entradas)[0]


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


def gerar_web(base, planilha=None):
    """Página web em web/ (servida pelo WAMP) + servidor.php: o Python que o api.php deve chamar."""
    import site
    from invest.apresentacao import web
    from invest.infra import historico
    if planilha is None:
        saida = caminho_saida(base.cfg)
        planilha = saida if saida.exists() else saida.with_suffix(".xlsx")
    for m in web.gerar(base, WEB, "local", historico.carregar(), planilha):
        print(f"6. {m}")
    cfg = json.dumps({"python": sys.executable, "raiz": str(BASE), "site_usuario": site.getusersitepackages()})
    (WEB / "servidor.php").write_text("<?php // gerado por gestao.py --web: Python usado pelo api.php\n"
                                      f"return json_decode(<<<'JSON'\n{cfg}\nJSON, true);\n", encoding="utf-8")


def web_atualizar(a, cfg):
    """Botão ⟳ da página: cotações (se online) + página. Não mexe na planilha (pode estar aberta no Excel)."""
    if not a.offline:
        atualizar_online(cfg)
    gerar_web(conferir())


def web_registrar(arquivo):
    """Formulários da página web -> mesmas regras do Cadastro da planilha. Responde em JSON (para o api.php)."""
    pedido = json.loads(Path(arquivo).read_text(encoding="utf-8"))
    resumo, pend_c, pend_p = importar(pedido.get("lancamentos", []), pedido.get("proventos", []))
    erros = [d["Observação"] for d in pend_c + pend_p]
    if pedido.get("metas"):
        r_metas, e_metas = atualizar_metas(pedido["metas"])
        resumo, erros = f"{resumo} • {r_metas}", erros + e_metas
    base = conferir_e_analisar()
    if not tem_erro(base.avisos):
        gerar_web(base)
    return {"ok": not erros, "resumo": resumo, "erros": erros}


def main():
    a = argumentos()
    cfg = ler_config()
    if a.web_registrar:
        print("@@" + json.dumps(web_registrar(a.web_registrar), ensure_ascii=False))
        return
    if a.web_atualizar:
        return web_atualizar(a, cfg)
    saida = caminho_saida(cfg, a.saida)
    anterior = saida if saida.exists() else saida.with_suffix(".xlsx")   # transição da versão sem macros
    garantir_fechada({saida, anterior}, a.esperar)

    print("1. Importando a aba Cadastro...")
    resumo, pendentes = importar_cadastro(anterior)
    print(f"  {resumo}")
    print(f"  metas: {importar_metas(anterior)}")
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
    if a.web:
        gerar_web(base, saida if saida.exists() else saida.with_suffix(".xlsx"))
    if a.abrir and saida.exists():
        os.startfile(saida)


if __name__ == "__main__":
    main()
