"""Aba Meta: você decide quanto cada ação/FII deve pagar por mês e em quanto tempo; a aba calcula valor da
cota, quantas cotas comprar, quanto investir no total e por mês para cumprir o prazo.

Campos amarelos (Meta mensal, Prazo): vêm com o padrão (Premissas) como fórmula; o número digitado por cima
vira meta personalizada, salva em dados/metas.csv na próxima atualização (ler_metas_preenchidas).
O resto é fórmula, lendo a Carteira na mesma linha de cada ativo.
"""
import warnings
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter as L

from .abas import D0, H
from .estilo import (BRL, BRL0, BRL4, F, MESES, PCT, QTD, TEAL, VERDE, VERMELHO, cabecalho, cel, f_bold, f_input,
                     f_link, f_norm, faixa, fill_input, fill_total, fill_zebra, larguras, linha_total, notas, tabela)

ABA = "Meta"
META, PRAZO = "Meta mensal (R$)", "Prazo (meses)"          # cabeçalhos dos campos editáveis (usados na leitura)
F_PERSONALIZADA = Font(name=F, size=10, bold=True, color="0000FF")      # meta digitada pelo usuário (≠ padrão)

# (cabeçalho, fórmula da linha r, formato, tipo) — tipo: "link" (Carteira), "entrada" (amarelo) ou "calc"
COLUNAS = [
    ("Ticker", '=Carteira!A{r}', None, "link"),                                                    # A
    ("Tipo", '=Carteira!B{r}', None, "link"),                                                      # B
    ("Segmento", '=Carteira!AB{r}', None, "link"),                                                 # C
    ("Valor da cota (R$)", '=Carteira!F{r}', BRL, "link"),                                         # D
    ("Provento mensal por cota (R$)", '=Carteira!U{r}', BRL4, "link"),                             # E
    ("Rendimento mensal", '=IF(D{r}>0,E{r}/D{r},0)', '0.00%', "calc"),                             # F
    (META, '=IF(A{r}="","",META_MENSAL)', BRL, "entrada"),                                         # G
    (PRAZO, '=IF(A{r}="","",PRAZO_META)', '0', "entrada"),                                         # H
    ("Cotas necessárias", '=IF(N(G{r})=0,"",IF(E{r}>0,ROUNDUP(G{r}/E{r},0),"sem proventos"))', QTD, "calc"),  # I
    ("Valor total a investir (R$)", '=IF(ISNUMBER(I{r}),I{r}*D{r},"")', BRL0, "calc"),            # J
    ("Cotas que você tem", '=Carteira!C{r}', QTD, "link"),                                         # K
    ("Preço médio pago (R$)", '=Carteira!D{r}', BRL, "link"),                                      # L
    ("Valor investido hoje (R$)", '=Carteira!E{r}', BRL0, "link"),                                 # M
    ("Renda mensal atual (R$)", '=Carteira!V{r}', BRL, "link"),                                    # N
    ("% da meta", '=IF(N(G{r})>0,MIN(1,N{r}/G{r}),0)', PCT, "calc"),                              # O
    ("Cotas faltantes", '=IF(ISNUMBER(I{r}),MAX(0,I{r}-K{r}),"")', QTD, "calc"),                   # P
    ("Valor faltante (R$)", '=IF(ISNUMBER(P{r}),P{r}*D{r},"")', BRL0, "calc"),                     # Q
    ("Cotas por mês", '=IF(AND(ISNUMBER(P{r}),N(H{r})>0),ROUNDUP(P{r}/H{r},0),"")', QTD, "calc"),  # R
    ("Aporte mensal no prazo (R$)", '=IF(AND(ISNUMBER(Q{r}),N(H{r})>0),Q{r}/H{r},"")', BRL0, "calc"),   # S
    ("Aporte mensal reinvestindo proventos (R$)",                                                  # T
     '=IF(NOT(ISNUMBER(S{r})),"",IF(F{r}>0,Q{r}*F{r}/((1+F{r})^H{r}-1),S{r}))', BRL0, "calc"),
    ("Data prevista", '=IF(N(H{r})>0,EDATE(TODAY(),H{r}),"")', "mm/yyyy", "calc"),                # U
]
SOMAS = {"G": BRL, "J": BRL0, "M": BRL0, "N": BRL, "P": QTD, "Q": BRL0, "R": QTD, "S": BRL0, "T": BRL0}


def aba_meta(wb, ativos, metas, T):
    """ativos: catálogo (mesma ordem das linhas da Carteira); metas: resolver() — personalizadas viram número."""
    ws = wb.create_sheet(ABA)
    n = len(COLUNAS)
    faixa(ws, "Meta de Renda por Ativo", "Amarelo = você decide: quanto cada ativo deve pagar por mês e em quantos meses "
          "(negrito = personalizado; vazio = padrão das Premissas). O resto é calculado.", n)
    cabecalho(ws, H, [c[0] for c in COLUNAS])
    ws.row_dimensions[H].height = 44
    for i, tk in enumerate(ativos["ticker"]):
        r = D0 + i
        z = fill_zebra if i % 2 else None
        pers = metas.loc[tk] if tk in metas.index and metas.loc[tk, "personalizada"] else None
        for col, (_, formula, fmt, tipo) in enumerate(COLUNAS, 1):
            valor = formula.format(r=r)
            if tipo == "entrada":
                if pers is not None:                     # personalizada: número (senão, fórmula do padrão)
                    valor = float(pers["meta_mensal"]) if col == 7 else int(pers["prazo_meses"])
                cel(ws, r, col, valor, fmt, F_PERSONALIZADA if pers is not None else f_input, fill_input)
            else:
                cel(ws, r, col, valor, fmt, f_link if tipo == "link" else f_norm, z)
    tabela(ws, "tMeta", H, T - 1, n)
    _totais(ws, T)
    fim = _resumo(ws, T)
    _formatacao(ws, T)
    notas(ws, fim + 2, [
        "Cotas necessárias = meta mensal ÷ provento mensal por cota (proventos dos últimos 12 meses ÷ 12), arredondado para cima.",
        "Valor total a investir = cotas necessárias × valor da cota hoje. Valor faltante desconta as cotas que você já tem.",
        "Aporte mensal no prazo = valor faltante ÷ prazo. Reinvestindo proventos: cada cota comprada passa a pagar e o "
        "provento compra mais cotas (juros compostos à taxa do rendimento mensal) — o aporte necessário fica menor.",
        "Para personalizar, digite por cima do campo amarelo; apague para voltar ao padrão. Salva em dados/metas.csv "
        "ao clicar em ⟳ Atualizar (ou rodar atualizar.bat).",
        "Cálculo com a cotação e o provento de hoje; mudanças de preço e de proventos alteram o resultado.",
    ])
    larguras(ws, [10, 7, 20, 12, 13, 11, 13, 10, 12, 16, 11, 12, 15, 13, 10, 11, 15, 10, 15, 17, 11])
    ws.freeze_panes = "B5"
    return ws


def _totais(ws, T):
    linha_total(ws, T, len(COLUNAS))
    for c, fmt in SOMAS.items():
        ws[f"{c}{T}"] = f"=SUM({c}{D0}:{c}{T - 1})"
        ws[f"{c}{T}"].number_format = fmt
    ws[f"O{T}"] = f"=IF(G{T}>0,N{T}/G{T},0)"
    ws[f"O{T}"].number_format = PCT


def _resumo(ws, T):
    """Cabe no seu aporte? Compara o aporte que os prazos exigem com o aporte mensal das Premissas."""
    linhas = [
        ("Aporte mensal para cumprir todos os prazos", f"=S{T}", BRL0),
        ("… reinvestindo os proventos", f"=T{T}", BRL0),
        ("Aporte mensal disponível (Premissas)", "=APORTE_MENSAL", BRL0),
        ("Sobra (+) ou falta (−) por mês, reinvestindo", f"=APORTE_MENSAL-T{T}", BRL0),
        ("Subtotal ações — renda desejada / a investir", f'=SUMIFS(G{D0}:G{T - 1},B{D0}:B{T - 1},"AÇÃO")', BRL0),
        ("Subtotal FIIs — renda desejada / a investir", f'=SUMIFS(G{D0}:G{T - 1},B{D0}:B{T - 1},"FII")', BRL0),
    ]
    r0 = T + 2
    for k, (rotulo, formula, fmt) in enumerate(linhas):
        r = r0 + k
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        cel(ws, r, 1, rotulo, None, f_bold)
        cel(ws, r, 7, formula, fmt, f_bold, fill_total)
    for k, tipo in ((4, "AÇÃO"), (5, "FII")):              # valor a investir ao lado da renda desejada
        cel(ws, r0 + k, 10, f'=SUMIFS(J{D0}:J{T - 1},B{D0}:B{T - 1},"{tipo}")', BRL0, f_bold, fill_total)
    sobra = f"G{r0 + 3}"
    ws.conditional_formatting.add(sobra, CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color=VERMELHO)))
    ws.conditional_formatting.add(sobra, CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(name=F, bold=True, color=VERDE)))
    return r0 + len(linhas)


def _formatacao(ws, T):
    ws.conditional_formatting.add(f"O{D0}:O{T - 1}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                                 end_value=1, color=TEAL))


def ler_metas_preenchidas(planilha: Path):
    """[(ticker, meta digitada, prazo digitado)] da aba Meta de uma planilha salva; [] se não houver a aba.
    O ticker vem da aba Ativos (mesma linha): na aba Meta ele é fórmula."""
    if not planilha.exists():
        return []
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = load_workbook(planilha, data_only=False)
    if ABA not in wb.sheetnames or "Ativos" not in wb.sheetnames:
        return []
    ws, cat = wb[ABA], wb["Ativos"]
    cols = {ws.cell(H, c).value: c for c in range(1, ws.max_column + 1)}
    if META not in cols or PRAZO not in cols:
        return []                                         # layout antigo, sem campos editáveis
    out = []
    for r in range(D0, ws.max_row + 1):
        tk = cat.cell(r, 1).value
        if not isinstance(tk, str) or not tk.strip() or tk == "TOTAL":
            break
        out.append((tk.strip().upper(), ws.cell(r, cols[META]).value, ws.cell(r, cols[PRAZO]).value))
    return out
