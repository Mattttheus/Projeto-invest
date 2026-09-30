"""Menu de abas (padrão Casa Organizada): um item por módulo na linha 2 de todas as abas.

Os itens são formas com hiperlink para a aba (funcionam sem macros); o módulo aberto fica em negrito, sublinhado em dourado.
Inseridos pelo Excel em automacao.py, antes das tabelas dinâmicas.
"""
from .estilo import DOURADO, F, TOPO, bgr

# (aba, rótulo no menu)
ITENS = [("Painel", "Painel"), ("Carteira", "Carteira"), ("Análises", "Análises"), ("Meta", "Meta"), ("Quant", "Quant"),
         ("Modelos", "Modelos"), ("Projeção", "Projeção"), ("Cadastro", "✚ Lançar"),
         ("Lançamentos", "Lançamentos"), ("Proventos", "Proventos"), ("Premissas", "Premissas"),
         ("Verificar", "Verificar")]
INICIO = 10           # pontos: o menu começa alinhado à marca da faixa de cima
ESPACO = 6
LINHA = 2             # linha do menu (a 1 é a faixa azul-marinho com marca e título)


def _largura(texto):
    return 18 + 6.2 * len(texto)


def barra(ws, ativa):
    linha = ws.Rows(LINHA)
    topo, alt = linha.Top + 1, linha.Height - 2
    x = INICIO
    for aba, rotulo in ITENS:
        w = _largura(rotulo)
        shp = ws.Shapes.AddShape(1, x, topo, w, alt)            # retângulo sem fundo: só o texto clicável
        shp.Name = f"nav_{aba}"
        shp.Line.Visible = False
        shp.Fill.Visible = False
        tr = shp.TextFrame2.TextRange
        tr.Text = rotulo
        tr.Font.Name, tr.Font.Size, tr.Font.Bold = F, 9.5, aba == ativa
        tr.Font.Fill.ForeColor.RGB = bgr(TOPO if aba == ativa else "64748B")
        tr.ParagraphFormat.Alignment = 2                        # centralizado
        tf = shp.TextFrame2
        tf.VerticalAnchor = 3
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tf.WordWrap = False
        shp.Placement = 3                                       # não move nem redimensiona com as células
        destino = f"'{aba}'!A1" if not aba.isalnum() else f"{aba}!A1"
        ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=destino, ScreenTip=f"Ir para {aba}")
        if aba == ativa:                                        # sublinhado dourado da aba aberta
            sub = ws.Shapes.AddShape(1, x + 4, linha.Top + linha.Height - 3, w - 8, 3)
            sub.Name = "nav_ativa"
            sub.Line.Visible = False
            sub.Fill.ForeColor.RGB = bgr(DOURADO)
            sub.Placement = 3
        x += w + ESPACO


def inserir(wb):
    """Cria a barra em todas as abas visíveis que tenham o cabeçalho no padrão do sistema."""
    abas = {ws.Name for ws in wb.Worksheets}
    feitas = 0
    for ws in wb.Worksheets:
        if ws.Visible != -1 or ws.Name not in {a for a, _ in ITENS}:
            continue
        barra(ws, ws.Name)
        if ws.Name not in ("Painel", "Cadastro"):          # esses já definem a área de impressão
            # imprime só as colunas com dados (a faixa do topo e o menu são mais largos que a tabela)
            ult_col = ws.Cells(4, ws.Columns.Count).End(-4159).Column          # xlToLeft
            ult_lin = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(ult_lin, max(ult_col, 3))).Address
        feitas += 1
    faltando = [a for a, _ in ITENS if a not in abas]
    return f"barra de navegação em {feitas} abas" + (f" (sem aba: {', '.join(faltando)})" if faltando else "")
