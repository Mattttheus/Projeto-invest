"""Aba Análises: tabelas dinâmicas, gráfico dinâmico e segmentação de dados.

O openpyxl não cria tabelas dinâmicas; elas são montadas pelo próprio Excel (automação COM, Windows,
em automacao.py) depois que a planilha é salva. Sem Excel/pywin32 a aba fica só com o aviso e o resto da planilha funciona.
As tabelas leem as tabelas do Excel (tCarteira, tProventos, tLancamentos) e se atualizam ao abrir o arquivo.
"""
from openpyxl.styles import Alignment, Font

from .estilo import MODO, AZUL, CINZA_TXT, F, TEXTO, bgr, faixa, larguras

XL_DATABASE, XL_ROW, XL_COLUMN, XL_PAGE, XL_SUM = 1, 1, 2, 3, -4157
XL_TABULAR, XL_BAR_CLUSTERED = 1, 57
ESTILO = "PivotStyleLight16"          # azul/cinza, alinhado à paleta executiva


def aba_analises(wb):
    """Cria a aba (com layout padrão); o conteúdo dinâmico é inserido por montar()."""
    ws = wb.create_sheet("Análises")
    faixa(ws, "Análises Dinâmicas", "Tabelas dinâmicas sobre Carteira, Proventos e Lançamentos — use os filtros e a segmentação.", 16)
    if MODO["google"]:
        consultas(ws)
    else:
        aviso = ws.cell(5, 1, "As tabelas dinâmicas são criadas pelo Excel ao rodar gestao.py (requer Excel no Windows).")
        aviso.font = Font(name=F, size=9, italic=True, color=CINZA_TXT)
    larguras(ws, [18, 16, 16, 16, 16, 3, 18, 14, 14, 14, 3, 14, 14, 14, 14, 14])
    return ws


def consultas(ws):
    """Google Planilhas: análises com QUERY (equivalente às tabelas dinâmicas, recalculadas sozinhas)."""
    blocos = [
        ("A5", "Carteira por tipo de ativo", "A7",
         "=IFERROR(QUERY(Carteira!A5:AB200,\"select B, sum(E), sum(G), sum(H), sum(J) where C > 0 group by B "
         "label B 'Tipo', sum(E) 'Investido', sum(G) 'Posição', sum(H) 'Lucro', sum(J) 'Proventos'\",0),"
         "\"Nenhuma ação ou FII na carteira ainda\")"),
        ("G5", "Posição por ativo", "G7",
         "=IFERROR(QUERY(Carteira!A5:G200,\"select A, G where C > 0 order by G desc label A 'Ticker', G 'Posição'\",0),"
         "\"Sem posições\")"),
        ("A20", "Proventos por ativo", "A22",
         "=IFERROR(QUERY(Proventos!A5:F5000,\"select B, sum(F) where B is not null group by B order by sum(F) desc "
         "label B 'Ticker', sum(F) 'Proventos'\",0),\"Sem proventos registrados\")"),
        ("G20", "Lançamentos por ativo e operação", "G22",
         "=IFERROR(QUERY('Lançamentos'!A5:G5000,\"select B, sum(G) where B is not null group by B pivot C "
         "label B 'Ticker'\",0),\"Sem lançamentos\")"),
    ]
    for tit_cel, titulo, cel, formula in blocos:
        ws[tit_cel] = titulo
        ws[tit_cel].font = Font(name=F, size=11, bold=True, color=TEXTO)
        ws[cel] = formula
    ws["A3"] = "Consultas QUERY recalculadas automaticamente pelo Google Planilhas."


def _titulo(ws, celula, texto):
    c = ws.Range(celula)
    c.Value = texto
    c.Font.Name, c.Font.Size, c.Font.Bold, c.Font.Color = F, 11, True, bgr(TEXTO)


def _dados(pt, campos, fmt):
    for origem, nome in campos:
        df = pt.AddDataField(pt.PivotFields(origem), nome, XL_SUM)
        df.NumberFormat = fmt


def _pivot(wb, ws, fonte, destino, nome):
    pc = wb.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=fonte, Version=6)
    pc.RefreshOnFileOpen = True
    pt = pc.CreatePivotTable(TableDestination=ws.Range(destino), TableName=nome, DefaultVersion=6)
    pt.RowAxisLayout(XL_TABULAR)
    pt.TableStyle2 = ESTILO
    pt.ShowDrillIndicators = True
    return pc, pt


def criar(xl, wb):
    """Cria as análises na planilha já aberta pela automação (ver automacao.py). Retorna mensagem de status."""
    import pythoncom
    try:
        ws = wb.Worksheets("Análises")
        ws.Range("A5").ClearContents()
        # campos de tabela dinâmica usam o formato no idioma do Excel instalado (ex.: pt-BR: "R$" #.##0)
        brl = f'"R$" #{xl.International[3]}##0'   # índice 3 = xlThousandsSeparator (4)

        # 1. carteira por tipo + gráfico dinâmico + segmentação
        _titulo(ws, "A5", "Carteira por tipo de ativo")
        pc, pt = _pivot(wb, ws, "tCarteira", "A7", "ptCarteira")
        pt.PivotFields("Tipo").Orientation = XL_ROW
        _dados(pt, [("Valor investido (R$)", "Total investido"), ("Posição atual (R$)", "Posição"),
                    ("Lucro / prejuízo (R$)", "Lucro / prejuízo"), ("Proventos recebidos (R$)", "Proventos")], brl)
        pt.DataPivotField.Orientation = XL_COLUMN

        grafico = pc.CreatePivotTable(TableDestination=ws.Range("R7"), TableName="ptGrafico", DefaultVersion=6)
        grafico.RowAxisLayout(XL_TABULAR)
        grafico.PivotFields("Ticker").Orientation = XL_ROW
        _dados(grafico, [("Posição atual (R$)", "Posição atual")], brl)
        grafico.PivotFields("Ticker").AutoSort(2, "Posição atual")         # decrescente
        forma = ws.Shapes.AddChart2(-1, XL_BAR_CLUSTERED, ws.Range("G5").Left, ws.Range("G5").Top, 470, 300)
        ch = forma.Chart
        ch.SetSourceData(grafico.TableRange1)
        ch.HasTitle = True
        ch.ChartTitle.Text = "Posição por ativo (filtre pelo tipo)"
        ch.HasLegend = False
        ch.ShowAllFieldButtons = False
        ch.Axes(1).ReversePlotOrder = True                                # maior no topo
        ch.Axes(2).Delete()                                               # valores já estão nas barras
        ch.PlotVisibleOnly = False                                        # base do gráfico fica oculta
        ws.Range("R:S").EntireColumn.Hidden = True
        ch.FullSeriesCollection(1).Format.Fill.ForeColor.RGB = bgr(AZUL)
        ch.FullSeriesCollection(1).HasDataLabels = True
        ch.FullSeriesCollection(1).DataLabels().NumberFormat = '"R$ "#,##0,"k"'

        sc = wb.SlicerCaches.Add2(pt, "Tipo")
        sc.PivotTables.AddPivotTable(grafico)                              # um filtro para tabela e gráfico
        sl = sc.Slicers.Add(ws, pythoncom.Missing, "sTipo", "Tipo")
        sl.Top, sl.Left, sl.Width, sl.Height = ws.Range("A13").Top, ws.Range("A13").Left, 220, 95
        sl.Style = "SlicerStyleLight1"

        # 2. proventos por ativo, com filtro por tipo de provento
        _titulo(ws, "A27", "Proventos por ativo")
        _, pv = _pivot(wb, ws, "tProventos", "A31", "ptProventos")
        pv.PivotFields("Tipo").Orientation = XL_PAGE
        pv.PivotFields("Ticker").Orientation = XL_ROW
        _dados(pv, [("Total (R$)", "Proventos recebidos")], brl)
        pv.PivotFields("Ticker").AutoSort(2, "Proventos recebidos")

        # 3. lançamentos por ativo e operação
        _titulo(ws, "G27", "Lançamentos por ativo")
        _, lc = _pivot(wb, ws, "tLancamentos", "G31", "ptLancamentos")
        lc.PivotFields("Ticker").Orientation = XL_ROW
        lc.PivotFields("Operação").Orientation = XL_COLUMN
        _dados(lc, [("Total (R$)", "Valor movimentado")], brl)

        ws.Range("A3").Value = "Clique numa tabela para ver a lista de campos • o filtro 'Tipo' atualiza a tabela e o gráfico."
        return "tabelas dinâmicas criadas na aba Análises"
    except Exception as e:  # a planilha continua válida sem elas
        return f"tabelas dinâmicas não criadas ({e})"
