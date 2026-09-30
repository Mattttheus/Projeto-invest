"""Macros (VBA) e botões da planilha — inseridos pelo Excel em automacao.py.

Botões:
  Cadastro  ✔ Registrar lançamento / Limpar      ✔ Registrar provento / Limpar
  Painel    ✚ Novo lançamento   ⟳ Atualizar   filtros Todos / Ações / FIIs e Limpar
"Registrar" grava na base (dados/*.csv — fonte única) e na tabela da aba Lançamentos/Proventos,
limpa o formulário e atualiza fórmulas e tabelas dinâmicas na hora.
"Atualizar cotações" salva, fecha e roda gestao.py (busca cotações, gera de novo e reabre a planilha).
"""
from .estilo import AZUL, CINZA_TXT, DOURADO, F, NAVY, TOPO, bgr

VBA = r'''
Option Explicit
' InvestERP — gerado por gestao.py. Não edite aqui: as alterações são perdidas ao gerar a planilha de novo.

Private Function Campo(ByVal nome As String) As Range
    Set Campo = ThisWorkbook.Names(nome).RefersToRange
End Function

Private Function Vazio(ByVal nome As String) As Boolean
    Vazio = Len(Trim$(CStr(Campo(nome).Value))) = 0
End Function

Private Function Numero(ByVal nome As String) As Double
    If Not Vazio(nome) And IsNumeric(Campo(nome).Value) Then Numero = CDbl(Campo(nome).Value)
End Function

Private Sub Mensagem(ByVal nome As String, ByVal texto As String, ByVal ok As Boolean)
    With Campo(nome)
        .Value = texto
        .Font.Color = IIf(ok, RGB(22, 163, 74), RGB(220, 38, 38))
    End With
End Sub

Private Function PastaDados() As String
    PastaDados = CreateObject("Scripting.FileSystemObject").GetParentFolderName(ThisWorkbook.Path) & "\dados\"
End Function

Private Function Num(ByVal v As Double, ByVal casas As Integer) As String
    ' número no formato do CSV (vírgula decimal), independente do idioma do Windows
    Dim s As String
    s = Trim$(Str$(Round(v, casas)))
    If Left$(s, 1) = "." Then s = "0" & s
    Num = Replace(s, ".", ",")
End Function

Private Sub AcrescentarCSV(ByVal arquivo As String, ByVal linha As String)
    Dim st As Object, txt As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile PastaDados() & arquivo
    txt = st.ReadText
    If Len(txt) > 0 And Right$(txt, 1) <> vbLf Then txt = txt & vbCrLf
    st.Position = 0
    st.SetEOS
    st.WriteText txt & linha & vbCrLf
    st.SaveToFile PastaDados() & arquivo, 2
    st.Close
End Sub

Private Function NovaLinha(ByVal lo As ListObject) As Range
    ' reaproveita a linha vazia deixada quando a tabela ainda não tem registros
    If lo.ListRows.Count = 1 And Len(CStr(lo.DataBodyRange.Cells(1, 2).Value)) = 0 Then
        Set NovaLinha = lo.DataBodyRange.Rows(1)
    Else
        Set NovaLinha = lo.ListRows.Add.Range
    End If
End Function

Private Function TickerCadastrado(ByVal tk As String) As Boolean
    TickerCadastrado = Not IsError(Application.Match(tk, ThisWorkbook.Worksheets("Ativos").Range("A:A"), 0))
End Function

Private Sub Limpar(ByVal nomes As String)
    Dim n As Variant
    For Each n In Split(nomes, ",")
        Campo(CStr(n)).ClearContents
    Next n
End Sub

Private Sub Foco(ByVal nome As String)
    ' leva o cursor ao campo só se a aba Cadastro estiver na tela
    On Error Resume Next
    If ActiveSheet.Name = Campo(nome).Worksheet.Name Then Campo(nome).Select
End Sub

Private Sub Atualizar()
    ' recalcula fórmulas e tabelas dinâmicas; uma falha aqui não desfaz o registro já gravado
    On Error Resume Next
    Application.Calculate
    Dim ws As Worksheet, pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
        Next pt
    Next ws
    On Error GoTo 0
End Sub

Public Sub LimparLancamento()
    Limpar "lc_data,lc_ticker,lc_op,lc_qtd,lc_preco,lc_custos,lc_tipo,lc_seg"
    Campo("lc_status").MergeArea.ClearContents
    Foco "lc_ticker"
End Sub

Public Sub LimparProvento()
    Limpar "pv_data,pv_ticker,pv_tipo,pv_valor,pv_qtd"
    Campo("pv_status").MergeArea.ClearContents
    Foco "pv_ticker"
End Sub

Public Sub RegistrarLancamento()
    Dim tk As String, op As String, tipo As String, seg As String
    Dim qtd As Double, preco As Double, custos As Double, tem As Double, dt As Date, novo As Boolean
    On Error GoTo Falha
    tk = UCase$(Trim$(CStr(Campo("lc_ticker").Value)))
    op = UCase$(Trim$(CStr(Campo("lc_op").Value)))
    If op = "" Then op = "COMPRA"
    qtd = Numero("lc_qtd")
    If tk = "" Then Mensagem "lc_status", "Informe o ticker.", False: Exit Sub
    If op <> "COMPRA" And op <> "VENDA" Then Mensagem "lc_status", "Operação deve ser COMPRA ou VENDA.", False: Exit Sub
    If qtd <= 0 Then Mensagem "lc_status", "Informe a quantidade de cotas.", False: Exit Sub
    novo = Not TickerCadastrado(tk)
    If novo Then
        tipo = UCase$(Trim$(CStr(Campo("lc_tipo").Value)))
        seg = Trim$(CStr(Campo("lc_seg").Value))
        If tipo <> "AÇÃO" And tipo <> "FII" Then
            Mensagem "lc_status", "Ticker novo: informe o Tipo do ativo (AÇÃO ou FII).", False: Exit Sub
        End If
    End If
    preco = Numero("lc_preco")
    If preco <= 0 Then preco = Numero("lc_cot")          ' valor unitário já estabelecido no cadastro do ativo
    If preco <= 0 Then Mensagem "lc_status", "Sem cotação para este ticker: informe o preço por cota.", False: Exit Sub
    custos = Numero("lc_custos")
    tem = Numero("lc_tem")
    If op = "VENDA" And qtd > tem Then
        Mensagem "lc_status", "Venda maior que as cotas que você tem (" & tem & ").", False: Exit Sub
    End If
    If IsDate(Campo("lc_data").Value) Then dt = CDate(Campo("lc_data").Value) Else dt = Date

    ' 1) base de dados (fonte única)
    If novo Then AcrescentarCSV "ativos.csv", tk & ";" & tipo & ";" & IIf(seg = "", "-", seg) & ";" & Num(preco, 2) & ";0;0;manual"
    AcrescentarCSV "lancamentos.csv", Format$(dt, "dd\/mm\/yyyy") & ";" & tk & ";" & op & ";" & _
                   Num(qtd, 6) & ";" & Num(preco, 2) & ";" & Num(custos, 2)
    ' 2) planilha (aparece na hora na aba Lançamentos, Carteira e Painel)
    Dim r As Range
    Set r = NovaLinha(ThisWorkbook.Worksheets("Lançamentos").ListObjects("tLancamentos"))
    r.Cells(1, 1).Value = dt
    r.Cells(1, 2).Value = tk
    r.Cells(1, 3).Value = op
    r.Cells(1, 4).Value = qtd
    r.Cells(1, 5).Value = preco
    r.Cells(1, 6).Value = custos
    r.Cells(1, 7).FormulaR1C1 = "=RC[-3]*RC[-2]+RC[-1]"

    Limpar "lc_data,lc_ticker,lc_op,lc_qtd,lc_preco,lc_custos,lc_tipo,lc_seg"
    Atualizar
    Mensagem "lc_status", ChrW(10004) & "  Registrado: " & op & " de " & qtd & " " & tk & " a R$ " & Format$(preco, "#,##0.00") & _
             IIf(novo, "  •  ticker novo: clique em Atualizar cotações para incluí-lo na carteira.", ""), True
    Foco "lc_ticker"
    Exit Sub
Falha:
    Mensagem "lc_status", "Erro ao registrar: " & Err.Description, False
End Sub

Public Sub RegistrarProvento()
    Dim tk As String, tipo As String, valor As Double, qtd As Double, dt As Date
    On Error GoTo Falha
    tk = UCase$(Trim$(CStr(Campo("pv_ticker").Value)))
    tipo = UCase$(Trim$(CStr(Campo("pv_tipo").Value)))
    If tipo = "" Then tipo = "DIVIDENDO"
    valor = Numero("pv_valor")
    qtd = Numero("pv_qtd")
    If qtd <= 0 Then qtd = Numero("pv_tem")
    If tk = "" Or Not TickerCadastrado(tk) Then Mensagem "pv_status", "Ticker não cadastrado: registre a compra primeiro.", False: Exit Sub
    If InStr(",DIVIDENDO,JCP,RENDIMENTO,AMORTIZAÇÃO,", "," & tipo & ",") = 0 Then Mensagem "pv_status", "Tipo de provento inválido.", False: Exit Sub
    If valor <= 0 Then Mensagem "pv_status", "Informe o valor por cota.", False: Exit Sub
    If qtd <= 0 Then Mensagem "pv_status", "Você não tem cotas deste ativo: informe a quantidade.", False: Exit Sub
    If IsDate(Campo("pv_data").Value) Then dt = CDate(Campo("pv_data").Value) Else dt = Date

    AcrescentarCSV "proventos.csv", Format$(dt, "dd\/mm\/yyyy") & ";" & tk & ";" & tipo & ";" & Num(valor, 4) & ";" & Num(qtd, 6)
    Dim r As Range
    Set r = NovaLinha(ThisWorkbook.Worksheets("Proventos").ListObjects("tProventos"))
    r.Cells(1, 1).Value = dt
    r.Cells(1, 2).Value = tk
    r.Cells(1, 3).Value = tipo
    r.Cells(1, 4).Value = valor
    r.Cells(1, 5).Value = qtd
    r.Cells(1, 6).FormulaR1C1 = "=RC[-2]*RC[-1]"

    Limpar "pv_data,pv_ticker,pv_tipo,pv_valor,pv_qtd"
    Atualizar
    Mensagem "pv_status", ChrW(10004) & "  Registrado: " & tipo & " de " & tk & " — R$ " & Format$(valor * qtd, "#,##0.00"), True
    Foco "pv_ticker"
    Exit Sub
Falha:
    Mensagem "pv_status", "Erro ao registrar: " & Err.Description, False
End Sub

Private Sub Filtrar(ByVal tipo As String)
    ThisWorkbook.Names("FILTRO_TIPO").RefersToRange.Value = tipo
    PintarFiltros
End Sub

Public Sub PintarFiltros()
    ' botão do tipo selecionado em azul (como o mês escolhido no filtro), os demais em branco
    Dim atual As String, par As Variant, shp As Shape
    atual = CStr(ThisWorkbook.Names("FILTRO_TIPO").RefersToRange.Value)
    On Error Resume Next
    For Each par In Array(Array("flt_Todos", "Todos"), Array("flt_ACAO", "AÇÃO"), Array("flt_FII", "FII"))
        Set shp = ThisWorkbook.Worksheets("Painel").Shapes(par(0))
        If par(1) = atual Then
            shp.Fill.ForeColor.RGB = RGB(46, 109, 180)
            shp.Line.ForeColor.RGB = RGB(46, 109, 180)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        Else
            shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
            shp.Line.ForeColor.RGB = RGB(203, 213, 225)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(51, 65, 85)
        End If
    Next par
End Sub

Public Sub FiltrarTodos()
    Filtrar "Todos"
End Sub

Public Sub FiltrarAcoes()
    Filtrar "AÇÃO"
End Sub

Public Sub FiltrarFIIs()
    Filtrar "FII"
End Sub

Public Sub IrCadastro()
    ThisWorkbook.Worksheets("Cadastro").Activate
    Foco "lc_ticker"
End Sub

Public Sub AtualizarTudo()
    Dim raiz As String
    If MsgBox("Buscar cotações atualizadas e gerar a planilha de novo?" & vbLf & vbLf & _
              "Ela será salva, fechada e reaberta em seguida (leva cerca de 1 minuto).", _
              vbYesNo + vbQuestion, "InvestERP") = vbNo Then Exit Sub
    raiz = CreateObject("Scripting.FileSystemObject").GetParentFolderName(ThisWorkbook.Path)
    ThisWorkbook.Save
    Shell "cmd /c chcp 65001>nul & cd /d """ & raiz & """ & set PYTHONIOENCODING=utf-8& python gestao.py --esperar --abrir", vbNormalFocus
    ThisWorkbook.Close SaveChanges:=False
End Sub
'''


def botao(ws, endereco, texto, macro, cor=AZUL, fonte="FFFFFF", borda=None, tamanho=10):
    """Botão arredondado sobre um intervalo de células, ligado a uma macro."""
    rng = ws.Range(endereco)
    shp = ws.Shapes.AddShape(5, rng.Left + 2, rng.Top + 3, rng.Width - 4, rng.Height - 6)   # 5 = retângulo arredondado
    shp.Name = f"btn_{macro}"
    try:
        shp.Adjustments.SetItem(1, 0.25)              # cantos levemente arredondados
    except Exception:
        pass
    shp.Fill.ForeColor.RGB = bgr(cor)
    if borda:
        shp.Line.ForeColor.RGB = bgr(borda)
        shp.Line.Weight = 1
    else:
        shp.Line.Visible = False
    tr = shp.TextFrame2.TextRange
    tr.Text = texto
    tr.Font.Name, tr.Font.Size, tr.Font.Bold = F, tamanho, True
    tr.Font.Fill.ForeColor.RGB = bgr(fonte)
    tr.ParagraphFormat.Alignment = 2                  # centralizado
    shp.TextFrame2.VerticalAnchor = 3                 # meio
    shp.TextFrame2.MarginLeft = shp.TextFrame2.MarginRight = 2
    shp.Placement = 2                                 # move com as células, sem redimensionar
    shp.OnAction = macro
    return shp


def inserir(wb, linhas_botoes):
    """Grava o módulo VBA e cria os botões. linhas_botoes: {'lc': linha, 'pv': linha} na aba Cadastro."""
    try:
        modulo = wb.VBProject.VBComponents.Add(1)     # 1 = módulo padrão
    except Exception as e:
        return f"macros não inseridas — o Excel bloqueou o acesso ao projeto VBA ({e})"
    modulo.Name = "InvestERP"
    modulo.CodeModule.AddFromString(VBA)

    cad = wb.Worksheets("Cadastro")
    rl, rp = linhas_botoes["lc"], linhas_botoes["pv"]
    botao(cad, f"B{rl}", "✔  Registrar lançamento", "RegistrarLancamento")
    botao(cad, f"C{rl}", "Limpar", "LimparLancamento", cor="FFFFFF", fonte=CINZA_TXT, borda="CBD5E1")
    botao(cad, f"E{rp}", "✔  Registrar provento", "RegistrarProvento", cor=NAVY)
    botao(cad, f"F{rp}", "Limpar", "LimparProvento", cor="FFFFFF", fonte=CINZA_TXT, borda="CBD5E1")

    from .painel import BOTAO_ATUALIZAR, BOTAO_LIMPAR, BOTAO_NOVO, FILTROS
    painel = wb.Worksheets("Painel")
    botao(painel, BOTAO_NOVO, "✚  Novo lançamento", "IrCadastro", cor=DOURADO, fonte=TOPO)
    botao(painel, BOTAO_ATUALIZAR, "⟳  Atualizar", "AtualizarTudo", cor="FFFFFF", fonte=TOPO, borda="CBD5E1")
    botao(painel, BOTAO_LIMPAR, "Limpar", "FiltrarTodos", cor="FFFFFF", fonte="2E6DB4", borda="CBD5E1", tamanho=8)
    for tipo, (rng, macro) in FILTROS.items():
        b = botao(painel, rng, {"AÇÃO": "Ações", "FII": "FIIs"}.get(tipo, tipo), macro, cor="FFFFFF", fonte="334155",
                  borda="CBD5E1", tamanho=9)
        b.Name = "flt_" + {"AÇÃO": "ACAO"}.get(tipo, tipo)
    try:
        wb.Application.Run(f"'{wb.Name}'!PintarFiltros")      # destaca o filtro atual (Todos)
    except Exception:
        pass
    return "macros e botões inseridos (Registrar, Limpar, Novo lançamento, Atualizar, filtros do Painel)"
