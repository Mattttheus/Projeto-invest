"""Barra de navegação no topo (estilo ERP): um item por módulo na linha 1 de todas as abas.

Os itens são formas com hiperlink para a aba (funcionam sem macros); o módulo aberto fica em destaque.
Inseridos pelo Excel em automacao.py, antes das tabelas dinâmicas.
"""
from .estilo import AZUL, F

# (aba, rótulo no menu)
ITENS = [("Painel", "Painel"), ("Carteira", "Carteira"), ("Análises", "Análises"), ("Cadastro", "✚ Lançar"),
         ("Lançamentos", "Lançamentos"), ("Proventos", "Proventos"), ("Premissas", "Premissas"),
         ("Verificar", "Verificar")]
INICIO = 118          # pontos: espaço reservado à marca "◆ InvestERP"
ESPACO = 4


def _bgr(hexa):
    return int(hexa[4:6] + hexa[2:4] + hexa[0:2], 16)


def _largura(texto):
    return 16 + 6.2 * len(texto)


def barra(ws, ativa):
    linha = ws.Rows(1)
    topo, alt = linha.Top + 5, linha.Height - 10
    x = INICIO
    for aba, rotulo in ITENS:
        w = _largura(rotulo)
        shp = ws.Shapes.AddShape(5, x, topo, w, alt)            # retângulo arredondado
        shp.Name = f"nav_{aba}"
        shp.Line.Visible = False
        if aba == ativa:
            shp.Fill.ForeColor.RGB = _bgr(AZUL)
        else:
            shp.Fill.Visible = False
        tr = shp.TextFrame2.TextRange
        tr.Text = rotulo
        tr.Font.Name, tr.Font.Size, tr.Font.Bold = F, 9, aba == ativa
        tr.Font.Fill.ForeColor.RGB = _bgr("FFFFFF" if aba == ativa else "CBD5E1")
        tr.ParagraphFormat.Alignment = 2                        # centralizado
        tf = shp.TextFrame2
        tf.VerticalAnchor = 3
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tf.WordWrap = False
        shp.Placement = 3                                       # não move nem redimensiona com as células
        destino = f"'{aba}'!A1" if not aba.isalnum() else f"{aba}!A1"
        ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=destino, ScreenTip=f"Ir para {aba}")
        x += w + ESPACO


def inserir(wb):
    """Cria a barra em todas as abas visíveis que tenham a linha 1 no padrão do sistema."""
    abas = {ws.Name for ws in wb.Worksheets}
    feitas = 0
    for ws in wb.Worksheets:
        if ws.Visible != -1 or ws.Name not in {a for a, _ in ITENS}:
            continue
        barra(ws, ws.Name)
        if ws.Name not in ("Painel", "Cadastro"):          # esses já definem a área de impressão
            # imprime só as colunas com dados (a faixa escura da linha 1 é mais larga para caber o menu)
            ult_col = ws.Cells(4, ws.Columns.Count).End(-4159).Column          # xlToLeft
            ult_lin = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(ult_lin, max(ult_col, 3))).Address
        feitas += 1
    faltando = [a for a, _ in ITENS if a not in abas]
    return f"barra de navegação em {feitas} abas" + (f" (sem aba: {', '.join(faltando)})" if faltando else "")
