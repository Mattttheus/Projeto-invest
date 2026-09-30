"""Atualiza cotação e proventos dos últimos 12 meses em dados/ativos.csv (fonte: Yahoo Finance).

Grava direto nas colunas preco, provento_anual_cota e atualizado_em — não cria cópia dos dados.
Se um ticker falhar, o valor anterior é mantido.
"""
import pandas as pd

from ..config import DADOS
from .yahoo import cotacao_e_proventos_12m

ARQ = DADOS / "ativos.csv"


def atualizar():
    df = pd.read_csv(ARQ, sep=";", decimal=",", encoding="utf-8-sig", dtype={"ticker": str})
    df = df.drop(columns=["fonte"], errors="ignore")      # redundante: atualizado_em já indica a origem
    if "atualizado_em" not in df:
        df["atualizado_em"] = ""
    df["atualizado_em"] = df["atualizado_em"].astype("object")
    ok = 0
    for i, t in df["ticker"].items():
        try:
            preco, prov, dia = cotacao_e_proventos_12m(t.strip().upper())
        except Exception as e:  # rede fora, ticker inexistente etc.
            print(f"  {t:7} mantido (falha: {e})")
            continue
        antes = df.at[i, "preco"]
        df.at[i, "preco"], df.at[i, "provento_anual_cota"] = preco, prov
        df.at[i, "atualizado_em"] = dia.strftime("%d/%m/%Y")
        ok += 1
        print(f"  {t:7} R$ {antes:>8.2f} -> R$ {preco:>8.2f}   proventos 12m R$ {prov:.4f}")
    df.to_csv(ARQ, sep=";", decimal=",", encoding="utf-8-sig", index=False)
    print(f"  {ok} de {len(df)} ativos atualizados.")
    return ok
