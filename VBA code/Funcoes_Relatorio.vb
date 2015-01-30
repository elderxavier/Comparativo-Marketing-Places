'*****************************************************
'* Relatorios
'*****************************************************


'*****************************************************
'* Funcao Cria nova Planilha
'* Input': WorkbookName
'* Output: Patch e nome do arquivo
'*****************************************************

Function CreatWorkbook(WorkbookName As String) As String
Dim oApp As Excel.Application
Dim oWks As Excel.Workbook
Dim strRep As String

strRep = WorkbookName & Now()
strRep = Replace(strRep, "/", "-")
strRep = Replace(strRep, ".", "-")
strRep = Replace(strRep, ":", "-")
strRep = Replace(strRep, ",", "-")
strRep = Replace(strRep, " ", "_")
strRep = strRep & ".xls"

WorkbookName = ThisWorkbook.Path & "\" & strRep

If Dir(WorkbookName) <> vbNullString Then
        'CaminhoExiste = True
        Kill WorkbookName
        Set oWks = oApp.Workbooks
        oWks.Close
End If
        'CaminhoExiste = False
    
    
Set oApp = New Excel.Application
Set oWks = oApp.Workbooks.Add
'oWks.SaveAs "C:\developer\Projetos\1-JANEIRO-2015\Comparativo Marketing Places\TuaVariavel.xls"
oWks.SaveAs WorkbookName
oWks.Close
CreatWorkbook = strRep

End Function

'*****************************************************




'*****************************************************
'* Funcao para criar bordas
'* Input': Range
'* Output: N/A
'*****************************************************

Function PaintBorder(Name As String, RangeSelect As String)
'Workbooks(Name).Worksheets(1).Activate
 Range(RangeSelect).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
        
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    'With Selection.Borders(xlInsideVertical)
     '   .LineStyle = xlContinuous
      '  .ColorIndex = 0
       ' .TintAndShade = 0
        '.Weight = xlThin
    'End With
    'With Selection.Borders(xlInsideHorizontal)
     '   .LineStyle = xlContinuous
      '  .ColorIndex = 0
       ' .TintAndShade = 0
        '.Weight = xlThin
    'End With
    
End Function

'*****************************************************

'*****************************************************
'* Funcao para criar bordas medias
'* Input': Range
'* Output: N/A
'*****************************************************

Function PaintBorderxlThick(Name As String, RangeSelect As String)
'Workbooks(Name).Worksheets(1).Activate
 Range(RangeSelect).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
        
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    'With Selection.Borders(xlInsideVertical)
     '   .LineStyle = xlContinuous
      '  .ColorIndex = 0
       ' .TintAndShade = 0
        '.Weight = xlThick
    'End With
    'With Selection.Borders(xlInsideHorizontal)
     '   .LineStyle = xlContinuous
      '  .ColorIndex = 0
       ' .TintAndShade = 0
        '.Weight = xlThick
    'End With
    
End Function



'*****************************************************
'* Extra
'*****************************************************
'*****************************************************
'* Funcao Cria Cabeçalho Relatorio
'* Input': WorkbookName
'* Output: N/A
'*****************************************************

Function CreatHeaderExtra(WorkbookName As String)
   
   Workbooks.Open (ThisWorkbook.Path & "\" & WorkbookName)
   'Workbooks(WorkbookName).Worksheets(1).Activate
    Workbooks(WorkbookName).Worksheets(1).Name = "Relatorio Extra"

    'TAMANHO LINHAS
    Workbooks(WorkbookName).Worksheets(1).Rows(1).RowHeight = 26.25
    Workbooks(WorkbookName).Worksheets(1).Rows(2).RowHeight = 26.25
    Workbooks(WorkbookName).Worksheets(1).Rows(3).RowHeight = 26.25
    Workbooks(WorkbookName).Worksheets(1).Rows(10).RowHeight = 30
    Workbooks(WorkbookName).Worksheets(1).Rows(11).RowHeight = 30.75
    
    ' Largura das colunas
    Workbooks(WorkbookName).Worksheets(1).Columns(1).ColumnWidth = 29.57 'A
    Workbooks(WorkbookName).Worksheets(1).Columns(2).ColumnWidth = 5.29 'B
    Workbooks(WorkbookName).Worksheets(1).Columns(3).ColumnWidth = 27.57 ' C
    Workbooks(WorkbookName).Worksheets(1).Columns(4).ColumnWidth = 26 ' D
    Workbooks(WorkbookName).Worksheets(1).Columns(5).ColumnWidth = 5.29 ' E
    Workbooks(WorkbookName).Worksheets(1).Columns(6).ColumnWidth = 27.43 'F
    Workbooks(WorkbookName).Worksheets(1).Columns(7).ColumnWidth = 27.43 'G
    Workbooks(WorkbookName).Worksheets(1).Columns(8).ColumnWidth = 11.57 'H
    Workbooks(WorkbookName).Worksheets(1).Columns(9).ColumnWidth = 5.29 'I
    Workbooks(WorkbookName).Worksheets(1).Columns(10).ColumnWidth = 23.43 'J
    Workbooks(WorkbookName).Worksheets(1).Columns(11).ColumnWidth = 12.57 'K
    Workbooks(WorkbookName).Worksheets(1).Columns(12).ColumnWidth = 8.43 'L
    Workbooks(WorkbookName).Worksheets(1).Columns(13).ColumnWidth = 5.29 'M
    Workbooks(WorkbookName).Worksheets(1).Columns(14).ColumnWidth = 27.86 'N
    Workbooks(WorkbookName).Worksheets(1).Columns(15).ColumnWidth = 23.29  'O
    Workbooks(WorkbookName).Worksheets(1).Columns(16).ColumnWidth = 5.29 'P
    Workbooks(WorkbookName).Worksheets(1).Columns(17).ColumnWidth = 13.71 'Q
    Workbooks(WorkbookName).Worksheets(1).Columns(18).ColumnWidth = 13.71 'R
    Workbooks(WorkbookName).Worksheets(1).Columns(19).ColumnWidth = 13.71 'S
    Workbooks(WorkbookName).Worksheets(1).Columns(20).ColumnWidth = 5.29 'T
    Workbooks(WorkbookName).Worksheets(1).Columns(21).ColumnWidth = 13.29 'U
    Workbooks(WorkbookName).Worksheets(1).Columns(22).ColumnWidth = 13.29 'V
    Workbooks(WorkbookName).Worksheets(1).Columns(23).ColumnWidth = 13.29 'W
    
    ' Mescla celulas
    ' TITULO
    Range("A1:B2").Merge 'Logo'
    Range("A3:B3").Merge ''
    Range("C1:G2").Merge 'titulo'
    Range("C3:G3").Merge 'DESCRICAO'
    Range("H1:I1").Merge ''
    Range("H2:I2").Merge ''
    Range("H3:I3").Merge ''
    'ESTATISTICAS
    Range("N1:S1").Merge 'TITULO'
    Range("P2:R2").Merge ''
    Range("P3:R3").Merge ''
    Range("P4:R4").Merge ''
    Range("P5:R5").Merge ''
    Range("P6:R6").Merge ''
    'CABECALHO 1
    Range("C9:L9").Merge ''
    Range("C10:D10").Merge ''
    Range("F10:H10").Merge ''
    Range("J10:L10").Merge ''
    
    'CABECALHO 2
    Range("N9:W9").Merge ''
    Range("N10:O10").Merge ''
    Range("Q10:S10").Merge ''
    Range("U10:W10").Merge ''
    
    'BORDAS
    PaintBorderxlThick WorkbookName, "A1:B2"
    PaintBorderxlThick WorkbookName, "A3:B3"
    PaintBorderxlThick WorkbookName, "C1:G2"
    PaintBorderxlThick WorkbookName, "C3:G3"
    PaintBorderxlThick WorkbookName, "H1:J2"
    PaintBorderxlThick WorkbookName, "H3:J3"
    PaintBorderxlThick WorkbookName, "N1:S1"
    PaintBorderxlThick WorkbookName, "P6:S6"
    
    
    'normal
    PaintBorder WorkbookName, "N2:O5"
    PaintBorder WorkbookName, "P2:S5"
    'ITENS 1
    PaintBorder WorkbookName, "A10"
    PaintBorder WorkbookName, "A11"
    PaintBorder WorkbookName, "A12:A9999"
    
    'ITENS 2
    PaintBorder WorkbookName, "C10:D10"
    PaintBorder WorkbookName, "C11"
    PaintBorder WorkbookName, "D11"
    PaintBorder WorkbookName, "C12:C9999"
    PaintBorder WorkbookName, "D12:D9999"
    
    'ITENS 3
    PaintBorder WorkbookName, "F10:H10"
    PaintBorder WorkbookName, "F11"
    PaintBorder WorkbookName, "G11"
    PaintBorder WorkbookName, "H11"
    PaintBorder WorkbookName, "F12:F9999"
    PaintBorder WorkbookName, "G12:G9999"
    PaintBorder WorkbookName, "H12:H9999"
    
    'ITENS 4
    PaintBorder WorkbookName, "J10:L10"
    PaintBorder WorkbookName, "J11"
    PaintBorder WorkbookName, "K11"
    PaintBorder WorkbookName, "L11"
    PaintBorder WorkbookName, "J12:J9999"
    PaintBorder WorkbookName, "K12:K9999"
    PaintBorder WorkbookName, "L12:L9999"
    
    'ITENS 5
    PaintBorder WorkbookName, "N10:O10"
    PaintBorder WorkbookName, "N11"
    PaintBorder WorkbookName, "O11"
    PaintBorder WorkbookName, "N12:N9999"
    PaintBorder WorkbookName, "O12:O9999"
    
    'ITENS 6
    PaintBorder WorkbookName, "Q10:S10"
    PaintBorder WorkbookName, "Q11"
    PaintBorder WorkbookName, "R11"
    PaintBorder WorkbookName, "S11"
    PaintBorder WorkbookName, "Q12:Q9999"
    PaintBorder WorkbookName, "R12:R9999"
    PaintBorder WorkbookName, "S12:S9999"
    
    'ITENS 7
    PaintBorder WorkbookName, "U10:W10"
    PaintBorder WorkbookName, "U11"
    PaintBorder WorkbookName, "V11"
    PaintBorder WorkbookName, "W11"
    PaintBorder WorkbookName, "U12:U9999"
    PaintBorder WorkbookName, "V12:V9999"
    PaintBorder WorkbookName, "W12:W9999"
    
    PaintBorderxlThick WorkbookName, "C9:L11"
    PaintBorderxlThick WorkbookName, "N9:W11"
    
    
    'alinhamento,fonte econteudo
    
    'TITULO
    Range("C1:G2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial Black"
        .Font.Size = 16
        .Font.Bold = True
        .Value = "SITUAÇÃO DE PRODUTOS MARKETING PLACE EXTRA"
    End With
    
    'DESCRIÇÃO
    Range("A3:B3").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "DESCRIÇÃO:"
    End With
    
    Range("C3:G3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial Narrow"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "RELAÇÃO E CONTROLE DE PRODUTOS DA EMPRESA E-LUSTRE NO EXTRA.COM"
    End With
    
    'DATA
    Range("H1:I1").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "DATA / HORA:"
    End With
    
    Range("J1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = Now()
    End With
    
    'CONTROLE
    Range("H2:I2").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "CONTROLE:"
    End With
    
    Dim strRep As String
    strRep = Now()
    strRep = Replace(strRep, "/", "")
    strRep = Replace(strRep, ".", "")
    strRep = Replace(strRep, ":", "")
    strRep = Replace(strRep, ",", "")
    strRep = Replace(strRep, " ", "")
    
    Range("J2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "RE-EX-" & strRep & "-A"
    End With
    
    'REVISAO
    Range("H3:I3").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "REVISÃO:"
    End With
    
    Range("J3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "A"
    End With
    
    'ESTATISTICAS
    Range("N1:S1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "ESTATÍSTICAS"
    End With
    
    Range("N2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "ESTOQUE - NÃO CADASTRADOS EXTRA:"
    End With
    Range("O2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(C12:C9998)"
        .Formula = "=CONT.VALORES(C12:C9998)"
    End With
    
    Range("N3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "ESTOQUE - NÃO CADASTRADOS MAGENTO:"
    End With
    Range("O3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(D12:D9998)"
        .Formula = "=CONT.VALORES(D12:D9998)"
    End With
    
    
    Range("N4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "ESTOQUE -DIVERGE DE QTD:"
    End With
    Range("O4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(F12:F9998)"
        .Formula = "=CONT.VALORES(F12:F9998)"
    End With
    
    Range("N5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "ESTOQUE - DIVERGE HABILITADO:"
    End With
    Range("O5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(J12:J9998)"
        .Formula = "=CONT.VALORES(J12:J9998)"
    End With
    
    Dim val As Variant
    
    Range("P2:R2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO - NÃO CADASTRADOS EXTRA:"
    End With
    Range("S2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(N12:N9998)"
        .Formula = "=CONT.VALORES(N12:N9998)"
        
        'val = Format(.Value, "##.##0,00")
        'val = CDbl(.Value)
        'val = FormatNumber(val, 2)
        '.Value = val
    End With
    
    
    Range("P3:R3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO - NÃO CADASTRADOS MAGENTO:"
    End With
    Range("S3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(O12:O9998)"
        .Formula = "=CONT.VALORES(O12:O9998)"
    End With
    
    
    Range("P4:R4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO -DIVERGE DE PREÇO:"
    End With
    Range("S4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(Q12:Q9998)"
        .Formula = "=CONT.VALORES(Q12:Q9998)"
    End With
    
    Range("P5:R5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO - DIVERGE PROMOÇÃO:"
    End With
    Range("S5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(U12:U9998)"
        .Formula = "=CONT.VALORES(U12:U9998)"
    End With
    
    Range("P5:R5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO - DIVERGE PROMOÇÃO:"
    End With
    Range("S5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(U12:U9998)"
        .Formula = "=CONT.VALORES(U12:U9998)"
    End With
    
    
    Range("P6:R6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        .Value = "TOTAL DE PRODUTOS MAGENTO:"
    End With
    Range("S6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(A12:A99998)"
        .Formula = "=CONT.VALORES(A12:A99998)"
    End With
    
    'COLUNAS
    Range("A10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "RELAÇAO GERAL DE PRODUTOS" & vbCrLf & "E-LUSTRE"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("A11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    '2
    Range("C9:L9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "COMPARATIVO PRODUTOS LOJA E-LUSTRE X EXTRA.COM - TABELA ESTOQUE"
        .Interior.Color = RGB(255, 89, 71)
    End With
    
    Range("C10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA DE PRODUTOS CADASTRADOS"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("C11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "PRODUTOS NÃO CADASTRADO" & vbCrLf & "NO EXTRA"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("D11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "PRODUTOS EXTRA NÃO" & vbCrLf & "CADASTRADO MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    Range("E10:E11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = ""
        .Interior.Color = RGB(255, 89, 71)
    End With
    
    Range("F10:I10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA NA QUANTIDADE DE  PRODUTOS "
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("F11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("G11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
     Range("H11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "EXTRA"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    Range("I10:I11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = ""
        .Interior.Color = RGB(255, 89, 71)
    End With
    
    Range("J10:L10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA EM  PRODUTOS" & vbCrLf & "HABILITADOS/DESABILITADOS"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("J11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("K11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
     Range("L11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "EXTRA"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    
    
    '3
    Range("N9:W9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "COMPARATIVO PRODUTOS LOJA E-LUSTRE X EXTRA.COM - TABELA ESTOQUE"
        .Interior.Color = RGB(112, 153, 202)
    End With
    
    Range("N10:O10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA DE PRODUTOS CADASTRADOS"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("N11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "PRODUTOS NÃO CADASTRADO" & vbCrLf & "NO EXTRA"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("O11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "PRODUTOS EXTRA NÃO" & vbCrLf & "CADASTRADO MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    Range("P10:P11").Select
    With Selection
        .Interior.Color = RGB(112, 153, 202)
    End With
    
    Range("Q10:S10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA NO PREÇO DOS PRODUTOS "
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("Q11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("R11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
     Range("S11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "EXTRA"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    Range("T10:T11").Select
    With Selection
        .Interior.Color = RGB(112, 153, 202)
    End With
    
    Range("U10:W10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA EM  PREÇO PROMOCIONAIS" & vbCrLf & "DOS PRODUTOS"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("U11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("V11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
     Range("W11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "EXTRA"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    'congela paineis
    Rows("1:11").Select
    ActiveWindow.FreezePanes = True
    
    'adiciona filtros
    ActiveSheet.Range("A11:W9999").AutoFilter
    'ActiveSheet.Range("A12:A9999").AutoFilter
    
    'inserir logo
    ActiveSheet.Shapes.AddPicture Filename:=ThisWorkbook.Path & "\images\logo.jpg", linktofile:=msoFalse, _
            savewithdocument:=msoCTrue, Left:=30, Top:=3, Width:=134, Height:=49
    
    'cota linhas da grade
    ActiveWindow.DisplayGridlines = False
     Range("A11").Select
  


    
End Function

'*****************************************************
    'extra

'*****************************************************
'* Funcao Compara os dados e inclui no relatorio - Extra
'* Input': WorkbookName
'* Output: Patch e nome do arquivo
'*****************************************************
Function CompDataForRel(WorkbookRelatorio As String)
    Dim WorkbookThis As String
    Dim UlinhaEste As Long
    Dim UlinhaComp As Long
    Dim UlinhaRel As Long
    Dim primeiraLinha As Integer
    Dim I As Long
    Dim J As Long
    Dim valor

    WorkbookThis = "Comparativo de tabelas.xlsm" 'nome da planilha principal
    primeiraLinha = 2 ' seta primeira linah apos cabeçalho

    Workbooks(WorkbookThis).Activate
    
    '***************PLANILHAS DE ESTOQUE******************
    ' * Verifica total de produtos - ESTOQUE*
    UlinhaEste = Workbooks(WorkbookThis).Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Magento Estoque
    UlinhaComp = Workbooks(WorkbookThis).Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Extra Estoque
        
        
   
    For I = primeiraLinha To UlinhaComp
            ' * Compara Produtos se cadastrados Extra - ESTOQUE*
            valor = Workbooks(WorkbookThis).Worksheets(3).Cells(I, 1).Value
            If LineExist(2, 1, valor, UlinhaEste) = False Then
               UlinhaRel = Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 4).End(xlUp).Offset(1, 0).Row '
                Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(UlinhaRel), 4).Value = Workbooks(WorkbookThis).Worksheets(3).Cells(I, 1).Value
            End If
                ' * Fim Compara Produtos se cadastrados extra*
    Next I
        
     For I = primeiraLinha To UlinhaEste
        'total de produtos magento
        Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row), 1).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value
        'fim total de produtos magento
          ' * Compara Produtos se cadastrados Magento - ESTOQUE*
            valor = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value
            If LineExist(3, 1, valor, UlinhaComp) = False Then
            UlinhaRel = Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 3).End(xlUp).Offset(1, 0).Row '
                Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(UlinhaRel), 3).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value
            End If
        ' * Fim Compara Produtos se cadastrados Magento*
        
        For J = primeiraLinha To UlinhaComp
            
            If UCase(Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value) = UCase(Workbooks(WorkbookThis).Worksheets(3).Cells(J, 1).Value) Then
            ' * Verifica divergencia de habilitados*
                If Workbooks(WorkbookThis).Worksheets(2).Cells(I, 2).Value <> Workbooks(WorkbookThis).Worksheets(3).Cells(J, 2).Value Then
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 6).End(xlUp).Offset(1, 0).Row), 6).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value ' SKU
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 7).End(xlUp).Offset(1, 0).Row), 7).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 2).Value ' MAGENTO
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 8).End(xlUp).Offset(1, 0).Row), 8).Value = Workbooks(WorkbookThis).Worksheets(3).Cells(J, 2).Value ' EXTRA
                End If
                ' * Verifica divergencia quantidade*
                If Workbooks(WorkbookThis).Worksheets(2).Cells(I, 3).Value <> Workbooks(WorkbookThis).Worksheets(3).Cells(J, 3).Value Then
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 10).End(xlUp).Offset(1, 0).Row), 10).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value ' SKU
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 11).End(xlUp).Offset(1, 0).Row), 11).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 3).Value ' MAGENTO
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 12).End(xlUp).Offset(1, 0).Row), 12).Value = Workbooks(WorkbookThis).Worksheets(3).Cells(J, 3).Value ' EXTRA
                End If
            End If
        Next J
    Next I
   
   '***************FIM PLANILHAS DE ESTOQUE******************
   
   
   
   
   '***************PLANILHAS DE PREÇO******************
    UlinhaEste = Workbooks(WorkbookThis).Worksheets(4).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Magento Estoque
    UlinhaComp = Workbooks(WorkbookThis).Worksheets(5).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Extra Estoque
        
    ' * Compara Produtos se cadastrados Extra - PRECO*
    For I = primeiraLinha To UlinhaComp
            valor = Workbooks(WorkbookThis).Worksheets(5).Cells(I, 1).Value
            If LineExist(4, 1, valor, UlinhaEste) = False Then
               Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 15).End(xlUp).Offset(1, 0).Row), 15).Value = Workbooks(WorkbookThis).Worksheets(5).Cells(I, 1).Value
            End If
    Next I
    ' * Fim Compara Produtos se cadastrados extra*
     For I = primeiraLinha To UlinhaEste
      ' * Compara Produtos se cadastrados Magento - PRECO*
            valor = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value
            If LineExist(5, 1, valor, UlinhaComp) = False Then
                Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 14).End(xlUp).Offset(1, 0).Row), 14).Value = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value
            End If
            ' * Fim Compara Produtos se cadastrados Magento*
        For J = primeiraLinha To UlinhaComp
            If UCase(Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value) = UCase(Workbooks(WorkbookThis).Worksheets(5).Cells(J, 1).Value) Then
            ' * Verifica divergencia de habilitados*
                If CDbl(Workbooks(WorkbookThis).Worksheets(4).Cells(I, 2).Value) <> CDbl(Workbooks(WorkbookThis).Worksheets(5).Cells(J, 2).Value) Then
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 17).End(xlUp).Offset(1, 0).Row), 17).Value = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value ' SKU
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 18).End(xlUp).Offset(1, 0).Row), 18).Value = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 2).Value ' MAGENTO
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 19).End(xlUp).Offset(1, 0).Row), 19).Value = Workbooks(WorkbookThis).Worksheets(5).Cells(J, 2).Value ' EXTRA
                End If
                ' * Verifica divergencia quantidade*
                If CDbl(Workbooks(WorkbookThis).Worksheets(4).Cells(I, 3).Value) <> CDbl(Workbooks(WorkbookThis).Worksheets(5).Cells(J, 3).Value) Then
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 21).End(xlUp).Offset(1, 0).Row), 21).Value = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value ' SKU
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 22).End(xlUp).Offset(1, 0).Row), 22).Value = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 3).Value ' MAGENTO
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 23).End(xlUp).Offset(1, 0).Row), 23).Value = Workbooks(WorkbookThis).Worksheets(5).Cells(J, 3).Value ' EXTRA
                End If
            End If
        Next J
    Next I
  
  
End Function












'*****************************************************
'* B2W
'*****************************************************
'*****************************************************
'* Funcao Cria Cabeçalho Relatorio B2W
'* Input': WorkbookName
'* Output: N/A
'*****************************************************

Function CreatHeaderB2W(WorkbookName As String)
   
   Workbooks.Open (ThisWorkbook.Path & "\" & WorkbookName)
   'Workbooks(WorkbookName).Worksheets(1).Activate
    Workbooks(WorkbookName).Worksheets(1).Name = "Relatorio B2W"

    'TAMANHO LINHAS
    Workbooks(WorkbookName).Worksheets(1).Rows(1).RowHeight = 26.25
    Workbooks(WorkbookName).Worksheets(1).Rows(2).RowHeight = 26.25
    Workbooks(WorkbookName).Worksheets(1).Rows(3).RowHeight = 26.25
    Workbooks(WorkbookName).Worksheets(1).Rows(10).RowHeight = 30
    Workbooks(WorkbookName).Worksheets(1).Rows(11).RowHeight = 30.75
    
    ' Largura das colunas
    Workbooks(WorkbookName).Worksheets(1).Columns(1).ColumnWidth = 29.57 'A
    Workbooks(WorkbookName).Worksheets(1).Columns(2).ColumnWidth = 5.29 'B
    Workbooks(WorkbookName).Worksheets(1).Columns(3).ColumnWidth = 27.57 ' C
    Workbooks(WorkbookName).Worksheets(1).Columns(4).ColumnWidth = 26 ' D
    Workbooks(WorkbookName).Worksheets(1).Columns(5).ColumnWidth = 5.29 ' E
    Workbooks(WorkbookName).Worksheets(1).Columns(6).ColumnWidth = 27.43 'F
    Workbooks(WorkbookName).Worksheets(1).Columns(7).ColumnWidth = 27.43 'G
    Workbooks(WorkbookName).Worksheets(1).Columns(8).ColumnWidth = 11.57 'H
    Workbooks(WorkbookName).Worksheets(1).Columns(9).ColumnWidth = 5.29 'I
    Workbooks(WorkbookName).Worksheets(1).Columns(10).ColumnWidth = 23.43 'J
    Workbooks(WorkbookName).Worksheets(1).Columns(11).ColumnWidth = 12.57 'K
    Workbooks(WorkbookName).Worksheets(1).Columns(12).ColumnWidth = 8.43 'L
    Workbooks(WorkbookName).Worksheets(1).Columns(13).ColumnWidth = 5.29 'M
    Workbooks(WorkbookName).Worksheets(1).Columns(14).ColumnWidth = 27.86 'N
    Workbooks(WorkbookName).Worksheets(1).Columns(15).ColumnWidth = 23.29  'O
    Workbooks(WorkbookName).Worksheets(1).Columns(16).ColumnWidth = 5.29 'P
    Workbooks(WorkbookName).Worksheets(1).Columns(17).ColumnWidth = 13.71 'Q
    Workbooks(WorkbookName).Worksheets(1).Columns(18).ColumnWidth = 13.71 'R
    Workbooks(WorkbookName).Worksheets(1).Columns(19).ColumnWidth = 13.71 'S
    Workbooks(WorkbookName).Worksheets(1).Columns(20).ColumnWidth = 5.29 'T
    Workbooks(WorkbookName).Worksheets(1).Columns(21).ColumnWidth = 13.29 'U
    Workbooks(WorkbookName).Worksheets(1).Columns(22).ColumnWidth = 13.29 'V
    Workbooks(WorkbookName).Worksheets(1).Columns(23).ColumnWidth = 13.29 'W
    
    ' Mescla celulas
    ' TITULO
    Range("A1:B2").Merge 'Logo'
    Range("A3:B3").Merge ''
    Range("C1:G2").Merge 'titulo'
    Range("C3:G3").Merge 'DESCRICAO'
    Range("H1:I1").Merge ''
    Range("H2:I2").Merge ''
    Range("H3:I3").Merge ''
    'ESTATISTICAS
    Range("N1:S1").Merge 'TITULO'
    Range("P2:R2").Merge ''
    Range("P3:R3").Merge ''
    Range("P4:R4").Merge ''
    Range("P5:R5").Merge ''
    Range("P6:R6").Merge ''
    'CABECALHO 1
    Range("C9:L9").Merge ''
    Range("C10:D10").Merge ''
    Range("F10:H10").Merge ''
    Range("J10:L10").Merge ''
    
    'CABECALHO 2
    Range("N9:W9").Merge ''
    Range("N10:O10").Merge ''
    Range("Q10:S10").Merge ''
    Range("U10:W10").Merge ''
    
    'BORDAS
    PaintBorderxlThick WorkbookName, "A1:B2"
    PaintBorderxlThick WorkbookName, "A3:B3"
    PaintBorderxlThick WorkbookName, "C1:G2"
    PaintBorderxlThick WorkbookName, "C3:G3"
    PaintBorderxlThick WorkbookName, "H1:J2"
    PaintBorderxlThick WorkbookName, "H3:J3"
    PaintBorderxlThick WorkbookName, "N1:S1"
    PaintBorderxlThick WorkbookName, "P6:S6"
    
    
    'normal
    PaintBorder WorkbookName, "N2:O5"
    PaintBorder WorkbookName, "P2:S5"
    'ITENS 1
    PaintBorder WorkbookName, "A10"
    PaintBorder WorkbookName, "A11"
    PaintBorder WorkbookName, "A12:A9999"
    
    'ITENS 2
    PaintBorder WorkbookName, "C10:D10"
    PaintBorder WorkbookName, "C11"
    PaintBorder WorkbookName, "D11"
    PaintBorder WorkbookName, "C12:C9999"
    PaintBorder WorkbookName, "D12:D9999"
    
    'ITENS 3
    PaintBorder WorkbookName, "F10:H10"
    PaintBorder WorkbookName, "F11"
    PaintBorder WorkbookName, "G11"
    PaintBorder WorkbookName, "H11"
    PaintBorder WorkbookName, "F12:F9999"
    PaintBorder WorkbookName, "G12:G9999"
    PaintBorder WorkbookName, "H12:H9999"
    
    'ITENS 4
    PaintBorder WorkbookName, "J10:L10"
    PaintBorder WorkbookName, "J11"
    PaintBorder WorkbookName, "K11"
    PaintBorder WorkbookName, "L11"
    PaintBorder WorkbookName, "J12:J9999"
    PaintBorder WorkbookName, "K12:K9999"
    PaintBorder WorkbookName, "L12:L9999"
    
    'ITENS 5
    PaintBorder WorkbookName, "N10:O10"
    PaintBorder WorkbookName, "N11"
    PaintBorder WorkbookName, "O11"
    PaintBorder WorkbookName, "N12:N9999"
    PaintBorder WorkbookName, "O12:O9999"
    
    'ITENS 6
    PaintBorder WorkbookName, "Q10:S10"
    PaintBorder WorkbookName, "Q11"
    PaintBorder WorkbookName, "R11"
    PaintBorder WorkbookName, "S11"
    PaintBorder WorkbookName, "Q12:Q9999"
    PaintBorder WorkbookName, "R12:R9999"
    PaintBorder WorkbookName, "S12:S9999"
    
    'ITENS 7
    PaintBorder WorkbookName, "U10:W10"
    PaintBorder WorkbookName, "U11"
    PaintBorder WorkbookName, "V11"
    PaintBorder WorkbookName, "W11"
    PaintBorder WorkbookName, "U12:U9999"
    PaintBorder WorkbookName, "V12:V9999"
    PaintBorder WorkbookName, "W12:W9999"
    
    PaintBorderxlThick WorkbookName, "C9:L11"
    PaintBorderxlThick WorkbookName, "N9:W11"
    
    
    'alinhamento,fonte econteudo
    
    'TITULO
    Range("C1:G2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial Black"
        .Font.Size = 16
        .Font.Bold = True
        .Value = "SITUAÇÃO DE PRODUTOS MARKETING PLACE B2W"
    End With
    
    'DESCRIÇÃO
    Range("A3:B3").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "DESCRIÇÃO:"
    End With
    
    Range("C3:G3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial Narrow"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "RELAÇÃO E CONTROLE DE PRODUTOS DA EMPRESA E-LUSTRE NA B2W"
    End With
    
    'DATA
    Range("H1:I1").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "DATA / HORA:"
    End With
    
    Range("J1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = Now()
    End With
    
    'CONTROLE
    Range("H2:I2").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "CONTROLE:"
    End With
    
    Dim strRep As String
    strRep = Now()
    strRep = Replace(strRep, "/", "")
    strRep = Replace(strRep, ".", "")
    strRep = Replace(strRep, ":", "")
    strRep = Replace(strRep, ",", "")
    strRep = Replace(strRep, " ", "")
    
    Range("J2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "RE-BW-" & strRep & "-A"
    End With
    
    'REVISAO
    Range("H3:I3").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "REVISÃO:"
    End With
    
    Range("J3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "A"
    End With
    
    'ESTATISTICAS
    Range("N1:S1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "ESTATÍSTICAS"
    End With
    
    Range("N2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "ESTOQUE - NÃO CADASTRADOS B2W:"
    End With
    Range("O2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(C12:C9998)"
        .Formula = "=CONT.VALORES(C12:C9998)"
    End With
    
    Range("N3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "ESTOQUE - NÃO CADASTRADOS MAGENTO:"
    End With
    Range("O3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(D12:D9998)"
        .Formula = "=CONT.VALORES(D12:D9998)"
    End With
    
    
    Range("N4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "ESTOQUE -DIVERGE DE QTD:"
    End With
    Range("O4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(F12:F9998)"
        .Formula = "=CONT.VALORES(F12:F9998)"
    End With
    
    Range("N5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "ESTOQUE - DIVERGE HABILITADO:"
    End With
    Range("O5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(J12:J9998)"
        .Formula = "=CONT.VALORES(J12:J9998)"
    End With
    
    Dim val As Variant
    
    Range("P2:R2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO - NÃO CADASTRADOS B2W:"
    End With
    Range("S2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(N12:N9998)"
        .Formula = "=CONT.VALORES(N12:N9998)"
        
        'val = Format(.Value, "##.##0,00")
        'val = CDbl(.Value)
        'val = FormatNumber(val, 2)
        '.Value = val
    End With
    
    
    Range("P3:R3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO - NÃO CADASTRADOS MAGENTO:"
    End With
    Range("S3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(O12:O9998)"
        .Formula = "=CONT.VALORES(O12:O9998)"
    End With
    
    
    Range("P4:R4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO -DIVERGE DE PREÇO:"
    End With
    Range("S4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(Q12:Q9998)"
        .Formula = "=CONT.VALORES(Q12:Q9998)"
    End With
    
    Range("P5:R5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO - DIVERGE PROMOÇÃO:"
    End With
    Range("S5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(U12:U9998)"
        .Formula = "=CONT.VALORES(U12:U9998)"
    End With
    
    Range("P5:R5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Value = "PREÇO - DIVERGE PROMOÇÃO:"
    End With
    Range("S5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(U12:U9998)"
        .Formula = "=CONT.VALORES(U12:U9998)"
    End With
    
    
    Range("P6:R6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        .Value = "TOTAL DE PRODUTOS MAGENTO:"
    End With
    Range("S6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = True
        '.Value = "=CONT.VALORES(A12:A99998)"
        .Formula = "=CONT.VALORES(A12:A99998)"
    End With
    
    'COLUNAS
    Range("A10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "RELAÇAO GERAL DE PRODUTOS" & vbCrLf & "E-LUSTRE"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("A11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    '2
    Range("C9:L9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "COMPARATIVO PRODUTOS LOJA E-LUSTRE X B2W - TABELA ESTOQUE"
        .Interior.Color = RGB(255, 89, 71)
    End With
    
    Range("C10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA DE PRODUTOS CADASTRADOS"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("C11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "PRODUTOS NÃO CADASTRADO" & vbCrLf & "NO B2W"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("D11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "PRODUTOS B2W NÃO" & vbCrLf & "CADASTRADO MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    Range("E10:E11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = ""
        .Interior.Color = RGB(255, 89, 71)
    End With
    
    Range("F10:I10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA NA QUANTIDADE DE  PRODUTOS "
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("F11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("G11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
     Range("H11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "B2W"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    Range("I10:I11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = ""
        .Interior.Color = RGB(255, 89, 71)
    End With
    
    Range("J10:L10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA EM  PRODUTOS" & vbCrLf & "HABILITADOS/DESABILITADOS"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("J11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("K11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
     Range("L11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "B2W"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    
    
    '3
    Range("N9:W9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "COMPARATIVO PRODUTOS LOJA E-LUSTRE X B2W - TABELA ESTOQUE"
        .Interior.Color = RGB(112, 153, 202)
    End With
    
    Range("N10:O10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA DE PRODUTOS CADASTRADOS"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("N11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "PRODUTOS NÃO CADASTRADO" & vbCrLf & "NO B2W"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("O11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "PRODUTOS B2W NÃO" & vbCrLf & "CADASTRADO MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    Range("P10:P11").Select
    With Selection
        .Interior.Color = RGB(112, 153, 202)
    End With
    
    Range("Q10:S10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA NO PREÇO DOS PRODUTOS "
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("Q11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("R11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
     Range("S11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "B2W"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    Range("T10:T11").Select
    With Selection
        .Interior.Color = RGB(112, 153, 202)
    End With
    
    Range("U10:W10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Value = "DIVERGÊNCIA EM  PREÇO PROMOCIONAIS" & vbCrLf & "DOS PRODUTOS"
        .Interior.Color = RGB(242, 242, 242)
    End With
    Range("U11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(191, 191, 191)
    End With
    Range("V11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "MAGENTO"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
     Range("W11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "B2W"
        .Interior.Color = RGB(191, 191, 191)
    End With
    
    'congela paineis
    Rows("1:11").Select
    ActiveWindow.FreezePanes = True
    
    'adiciona filtros
    ActiveSheet.Range("A11:W9999").AutoFilter
    'ActiveSheet.Range("A12:A9999").AutoFilter
    
    'inserir logo
    ActiveSheet.Shapes.AddPicture Filename:=ThisWorkbook.Path & "\images\logo.jpg", linktofile:=msoFalse, _
            savewithdocument:=msoCTrue, Left:=30, Top:=3, Width:=134, Height:=49
    
    'cota linhas da grade
    ActiveWindow.DisplayGridlines = False
     Range("A11").Select
  


    
End Function

'*****************************************************






'*****************************************************
'* Funcao Compara os dados e inclui no relatorio - B2W
'* Input': WorkbookName
'* Output: Patch e nome do arquivo
'*****************************************************
Function CompDataForRelB2W(WorkbookRelatorio As String)
    Dim WorkbookThis As String
    Dim UlinhaEste As Long
    Dim UlinhaComp As Long
    Dim UlinhaRel As Long
    Dim primeiraLinha As Integer
    Dim I As Long
    Dim J As Long
    Dim valor
    Dim compara As String
    WorkbookThis = "Comparativo de tabelas.xlsm" 'nome da planilha principal
    primeiraLinha = 2 ' seta primeira linah apos cabeçalho da planilha pesquisada

    Workbooks(WorkbookThis).Activate
    
    
    '***************PLANILHAS DE ESTOQUE******************
    ' * Verifica total de produtos - ESTOQUE*
    UlinhaEste = Workbooks(WorkbookThis).Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Magento Estoque
    UlinhaComp = Workbooks(WorkbookThis).Worksheets(6).Cells(Rows.Count, 4).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Extra Estoque
    
    
    
        
    ' * Fim Verifica produtos não cadastrados B2W*
    For I = primeiraLinha To UlinhaComp
    ' * Compara Produtos se cadastrados B2W - ESTOQUE*
            valor = UCase(Workbooks(WorkbookThis).Worksheets(6).Cells(I, 4).Value)
            If LineExist(2, 1, valor, UlinhaEste) = False Then
                UlinhaRel = Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 4).End(xlUp).Offset(1, 0).Row '
                Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(UlinhaRel), 4).Value = Workbooks(WorkbookThis).Worksheets(6).Cells(I, 4).Value
            End If
            ' * fim Compara Produtos se cadastrados B2W - ESTOQUE*
            ' * Compara Produtos se cadastrados B2W - PRECO*
            valor = UCase(Workbooks(WorkbookThis).Worksheets(6).Cells(I, 4).Value)
            If LineExist(4, 1, valor, UlinhaEste) = False Then
               Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 15).End(xlUp).Offset(1, 0).Row), 15).Value = Workbooks(WorkbookThis).Worksheets(6).Cells(I, 4).Value
            End If
            
            ' *Fim Compara Produtos se cadastrados B2W - PRECO*
    Next I
     
    
    
     For I = primeiraLinha To UlinhaEste
        '*Total de produtos Magento*'
        Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row), 1).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value
        'Workbooks(WorkbookRelatorio).Worksheets(1).Cells((I + 10), 1).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value
        ' * Compara Produtos se cadastrados Magento - ESTOQUE*
        valor = UCase(Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value)
            If LineExist(6, 4, valor, UlinhaComp) = False Then
                UlinhaRel = Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 3).End(xlUp).Offset(1, 0).Row '
                Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(UlinhaRel), 3).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value
            End If
        ' * Fim Compara Produtos se cadastrados Magento*
        
        For J = primeiraLinha To UlinhaComp
            
            ' * Verifica divergencia de habilitados*
            If UCase(Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value) = UCase(Workbooks(WorkbookThis).Worksheets(6).Cells(J, 4).Value) Then
                If Workbooks(WorkbookThis).Worksheets(2).Cells(I, 2).Value <> Workbooks(WorkbookThis).Worksheets(6).Cells(J, 7).Value Then
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 6).End(xlUp).Offset(1, 0).Row), 6).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value ' SKU
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 7).End(xlUp).Offset(1, 0).Row), 7).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 2).Value ' MAGENTO
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 8).End(xlUp).Offset(1, 0).Row), 8).Value = Workbooks(WorkbookThis).Worksheets(6).Cells(J, 7).Value ' B2W
                End If
                ' * Verifica divergencia habilitados*
                '
                If Workbooks(WorkbookThis).Worksheets(2).Cells(I, 3).Value <> "Y" And Workbooks(WorkbookThis).Worksheets(6).Cells(J, 9).Value = "ATIVO" Then
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 10).End(xlUp).Offset(1, 0).Row), 10).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 1).Value ' SKU
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 11).End(xlUp).Offset(1, 0).Row), 11).Value = Workbooks(WorkbookThis).Worksheets(2).Cells(I, 3).Value ' MAGENTO
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 12).End(xlUp).Offset(1, 0).Row), 12).Value = Workbooks(WorkbookThis).Worksheets(6).Cells(J, 9).Value ' B2W
                End If
            End If
        Next J
    Next I
   
   '***************FIM PLANILHAS DE ESTOQUE******************
   
   
   
   
   '***************PLANILHAS DE PREÇO******************
    UlinhaEste = Workbooks(WorkbookThis).Worksheets(4).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Magento Estoque
    UlinhaComp = Workbooks(WorkbookThis).Worksheets(6).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Extra Estoque

    'For I = primeiraLinha To UlinhaEste
        
    'Next I
        ' * Fim Compara Produtos se cadastrados Magento*

    
    
    ' * Fim Compara Produtos se cadastrados extra*
     For I = primeiraLinha To UlinhaEste
        'Ajusta dados planilha B2W
        valor = Workbooks(WorkbookThis).Worksheets(6).Cells(I, 5).Value
        valor = Replace(valor, ".", ",")
        Workbooks(WorkbookThis).Worksheets(6).Cells(I, 5).Value = valor
        
        
        valor = Workbooks(WorkbookThis).Worksheets(6).Cells(I, 6).Value
        valor = Replace(valor, ".", ",")
        Workbooks(WorkbookThis).Worksheets(6).Cells(I, 6).Value = valor
        
        'fim Ajusta dados planilha B2W
        
        ' * Compara Produtos se cadastrados Magento - PRECO*
        
            valor = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value
            If LineExist(6, 4, valor, UlinhaComp) = False Then
                Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 14).End(xlUp).Offset(1, 0).Row), 14).Value = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value
            End If
            ' * Fim Compara Preco de Produtos B2W*
        
        For J = primeiraLinha To UlinhaComp
            
            
            If UCase(Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value) = UCase(Workbooks(WorkbookThis).Worksheets(6).Cells(J, 4).Value) Then
            ' * Verifica divergencia de PRECOS*
                If CDbl(Workbooks(WorkbookThis).Worksheets(4).Cells(I, 2).Value) <> CDbl(Workbooks(WorkbookThis).Worksheets(6).Cells(J, 5).Value) Then
                                         
                     'Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 18).End(xlUp).Offset(1, 0).Row), 18).NumberFormat = "#0.00"
                     'Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 19).End(xlUp).Offset(1, 0).Row), 19).NumberFormat = "#0.00"
                     
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 17).End(xlUp).Offset(1, 0).Row), 17).Value = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value ' SKU
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 18).End(xlUp).Offset(1, 0).Row), 18).Value = CDbl(Workbooks(WorkbookThis).Worksheets(4).Cells(I, 2).Value) ' MAGENTO
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 19).End(xlUp).Offset(1, 0).Row), 19).Value = CDbl(Workbooks(WorkbookThis).Worksheets(6).Cells(J, 5).Value) ' EXTRA
                End If
                ' * Verifica divergencia PROMOCAO*
                If CDbl(Workbooks(WorkbookThis).Worksheets(4).Cells(I, 3).Value) <> CDbl(Workbooks(WorkbookThis).Worksheets(6).Cells(J, 6).Value) Then
                    
                     'Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 21).End(xlUp).Offset(1, 0).Row), 21).NumberFormat = "#0.00"
                     'Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 22).End(xlUp).Offset(1, 0).Row), 22).NumberFormat = "#0.00"
                    
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 21).End(xlUp).Offset(1, 0).Row), 21).Value = Workbooks(WorkbookThis).Worksheets(4).Cells(I, 1).Value ' SKU
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 22).End(xlUp).Offset(1, 0).Row), 22).Value = CDbl(Workbooks(WorkbookThis).Worksheets(4).Cells(I, 3).Value) ' MAGENTO
                     Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 23).End(xlUp).Offset(1, 0).Row), 23).Value = CDbl(Workbooks(WorkbookThis).Worksheets(6).Cells(J, 6).Value) ' EXTRA
                End If
            End If
        Next J
    Next I
  
  
End Function



'*****************************************************
'* Relatorio Geral de Produtos
'*****************************************************
'*****************************************************
'* Funcao Cria Cabeçalho Relatorio Relatorio Geral de Produtos Magento
'* Input': WorkbookName
'* Output: N/A
'*****************************************************

Function CreatHeaderRelGeral(WorkbookName As String)
   
   Workbooks.Open (ThisWorkbook.Path & "\" & WorkbookName)
   'Workbooks(WorkbookName).Worksheets(1).Activate
    Workbooks(WorkbookName).Worksheets(1).Name = "Relatorio B2W"

   'TAMANHO LINHAS
    Workbooks(WorkbookName).Worksheets(1).Rows(1).RowHeight = 26.25
    Workbooks(WorkbookName).Worksheets(1).Rows(2).RowHeight = 26.25
    Workbooks(WorkbookName).Worksheets(1).Rows(14).RowHeight = 28.5
    
    
    ' Largura das colunas
    Workbooks(WorkbookName).Worksheets(1).Columns(1).ColumnWidth = 8.43 'A
    Workbooks(WorkbookName).Worksheets(1).Columns(2).ColumnWidth = 21 'B
    Workbooks(WorkbookName).Worksheets(1).Columns(3).ColumnWidth = 16.86 ' C
    Workbooks(WorkbookName).Worksheets(1).Columns(4).ColumnWidth = 21.14 ' D
    Workbooks(WorkbookName).Worksheets(1).Columns(5).ColumnWidth = 18.57 ' E
    Workbooks(WorkbookName).Worksheets(1).Columns(6).ColumnWidth = 18.43 'F
    Workbooks(WorkbookName).Worksheets(1).Columns(7).ColumnWidth = 18.43 'G
    Workbooks(WorkbookName).Worksheets(1).Columns(8).ColumnWidth = 28.71 'H
    Workbooks(WorkbookName).Worksheets(1).Columns(9).ColumnWidth = 23.14 'I
    Workbooks(WorkbookName).Worksheets(1).Columns(10).ColumnWidth = 8.43 'J
    Workbooks(WorkbookName).Worksheets(1).Columns(11).ColumnWidth = 18
  
    
    ' Mescla celulas
    ' TITULO
    Range("A1:B2").Merge 'Logo'
    Range("A3:B3").Merge ''
    Range("C1:G2").Merge 'titulo'
    Range("C3:G3").Merge 'DESCRICAO'
    Range("H1:I1").Merge ''
    Range("H2:I2").Merge ''
    Range("H3:I3").Merge ''
    Range("J1:K1").Merge ''
    Range("J2:K2").Merge ''
    Range("J3:K3").Merge ''
    
    'LEGENDAS
    Range("A5:C5").Merge ''
    Range("A6:C6").Merge ''
    Range("A7:C7").Merge ''
    Range("A8:C8").Merge ''
    Range("A9:C9").Merge ''
    Range("A10:C10").Merge ''
    
    'ESTATISTICAS
    Range("E5:J5").Merge ''
    Range("E6:F6").Merge ''
    Range("E7:F7").Merge ''
    Range("E8:F8").Merge ''
    Range("E9:F9").Merge ''
    Range("E10:F10").Merge ''
    Range("H6:I6").Merge ''
    Range("H7:I7").Merge ''
    Range("H8:I8").Merge ''
    Range("H9:I9").Merge ''
    
    'BORDAS
    PaintBorderxlThick WorkbookName, "A1:B2"
    PaintBorderxlThick WorkbookName, "A3:B3"
    PaintBorderxlThick WorkbookName, "C1:G2"
    PaintBorderxlThick WorkbookName, "C3:G3"
    PaintBorderxlThick WorkbookName, "H1:K2"
    PaintBorderxlThick WorkbookName, "H3:K3"
    
    
    
    'normal
    'ESTATISTICAS
    PaintBorder WorkbookName, "A5:C7"
    PaintBorder WorkbookName, "A8:C10"
    
    PaintBorder WorkbookName, "E5:J5"
    PaintBorder WorkbookName, "E6:F6"
    PaintBorder WorkbookName, "G6"
    PaintBorder WorkbookName, "E7:F7"
    PaintBorder WorkbookName, "G7"
    PaintBorder WorkbookName, "E8:F8"
    PaintBorder WorkbookName, "G8"
    PaintBorder WorkbookName, "E9:F9"
    PaintBorder WorkbookName, "G9"
    PaintBorder WorkbookName, "E10:F10"
    PaintBorder WorkbookName, "G10"
    PaintBorder WorkbookName, "H6:I6"
    PaintBorder WorkbookName, "J6"
    PaintBorder WorkbookName, "H7:I7"
    PaintBorder WorkbookName, "J7"
    PaintBorder WorkbookName, "H8:I8"
    PaintBorder WorkbookName, "J8"
    PaintBorder WorkbookName, "H9:I9"
    PaintBorder WorkbookName, "J9"
    
    'TITULO COLUNAS
    PaintBorder WorkbookName, "A14"
    PaintBorder WorkbookName, "B14"
    PaintBorder WorkbookName, "C14"
    PaintBorder WorkbookName, "D14"
    PaintBorder WorkbookName, "E14"
    'COLUNAS
    PaintBorder WorkbookName, "A15:A9999"
    PaintBorder WorkbookName, "B15:B9999"
    PaintBorder WorkbookName, "C15:C9999"
    PaintBorder WorkbookName, "D15:D9999"
    PaintBorder WorkbookName, "E15:E9999"
    
    
    
    'alinhamento,fonte econteudo
    
    'TITULO
    Range("C1:G2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial Black"
        .Font.Size = 16
        .Font.Bold = True
        .Value = "RELATÓRIO GERAL DE ESTOQUE"
    End With
    
    'DESCRIÇÃO
    Range("A3:B3").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "DESCRIÇÃO:"
    End With
    
    Range("C3:G3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial Narrow"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "RELAÇÃO E CONTROLE DE DE ESTOQUE DA EMPRESA E-LUSTE"
    End With
    
    'DATA
    Range("H1:I1").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "DATA / HORA:"
    End With
    
    Range("J1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = Now()
    End With
    
    'CONTROLE
    Range("H2:I2").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "CONTROLE:"
    End With
    
    Dim strRep As String
    strRep = Now()
    strRep = Replace(strRep, "/", "")
    strRep = Replace(strRep, ".", "")
    strRep = Replace(strRep, ":", "")
    strRep = Replace(strRep, ",", "")
    strRep = Replace(strRep, " ", "")
    
    Range("J2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "RE-GE-" & strRep & "-A"
    End With
    
    'REVISAO
    Range("H3:I3").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "REVISÃO:"
    End With
    
    Range("J3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "A"
    End With
    
    'LEGENDAS
    Range("A5:C5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "STATUS"
        .Interior.Color = RGB(216, 216, 216)
    End With
    
    Range("A6:C6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "2 - Produto desabilitado"
        .Interior.Color = RGB(255, 192, 0)
    End With
    
    Range("A7:C7").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "1 - Produto habilitado"
        .Interior.Color = RGB(149, 179, 215)
    End With
    
    Range("A8:C8").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "ESTOQUE"
        .Interior.Color = RGB(216, 216, 216)
    End With
    
    Range("A9:C9").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "0 - Produto esgotado"
    End With
    
    Range("A10:C10").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "0 - Produto esgotado"
    End With
    
    'ESTATISTICAS
    Range("E5:J5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "ESTOQUE"
        .Interior.Color = RGB(216, 216, 216)
    End With
    
    Range("E6:F6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "2 - Quantidade de produtos desabilitados"
        .Interior.Color = RGB(255, 192, 0)
    End With
    
    Range("G6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        '.Value = "=CONT.SE(C13:C9999;2)"
        .Formula = "=CONT.SE(C15:C9999,2)"
        .Interior.Color = RGB(255, 192, 0)
    End With
    
    Range("E7:F7").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "Quantidade de produtos habilitados"
        .Interior.Color = RGB(149, 179, 215)
    End With
    
    Range("G7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        '.Value = "=CONT.SE(Plan1!C13:C9999;1)"
        .Formula = "=CONT.SE(C15:C9999,1)"
        .Interior.Color = RGB(149, 179, 215)
    End With
    
    Range("E8:F8").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "Produtos estoque desabilitado"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("G8").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "=CONT.SE(E15:E10000,0)"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("E9:F9").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Formula = "Produtos estoque habilitado"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("G9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Formula = "=CONT.SE(E15:E10000,1)"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    
    Range("E10:F10").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Value = "TOTAL DE PRODUTOS"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("G10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .Formula = "=CONT.VALORES(B15:B9999)"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("H6:I6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "Produtos desabilitados e desabilitados no estoque"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("J6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Formula = "=CONT.SES(C15:C9999,2,E15:E9999,0)"
    End With
    
    Range("H7:I7").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "Quantidade de produtos  habilitados e habilitados no estoque"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("J7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Formula = "=CONT.SES(C15:C9999,1,E15:E9999,1)"
        
    End With
    
     Range("H8:I8").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "Quantidade de produtos  habilitados e desabilitados no estoque"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("J8").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Formula = "=CONT.SES(C15:C9999,1,E15:E9999,0)"
        
        
    End With
    
    Range("H9:I9").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Value = "Quantidade de produtos  desabilitados e habilitados no estoque"
        .Interior.Color = RGB(221, 217, 195)
    End With
    
    Range("J9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Formula = "=CONT.SES(C15:C9999,2,E15:E9999,1)"
        
    End With
    
    'colunas
    Range("A14").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "ID"
        .Interior.Color = RGB(216, 216, 216)
    End With
    
    'colunas
    Range("B14").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "SKU"
        .Interior.Color = RGB(216, 216, 216)
    End With
    
    'colunas
    Range("C14").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "STATUS"
        .Interior.Color = RGB(216, 216, 216)
    End With
    
    'colunas
    Range("D14").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "QUANTIDADE"
        .Interior.Color = RGB(216, 216, 216)
    End With
    
    'colunas
    Range("E14").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .Value = "EM ESTOQUE"
        .Interior.Color = RGB(216, 216, 216)
    End With
    
    
    
    'congela paineis
    Range("A1").Select
    ActiveWindow.FreezePanes = True
    
    'adiciona filtros
    ActiveSheet.Range("A14:E9999").AutoFilter
    'ActiveSheet.Range("A12:A9999").AutoFilter
    
    'inserir logo
    ActiveSheet.Shapes.AddPicture Filename:=ThisWorkbook.Path & "\images\logo.jpg", linktofile:=msoFalse, _
            savewithdocument:=msoCTrue, Left:=10, Top:=3, Width:=134, Height:=49
    
     Range("A11").Select
    'ocuta linhas da grade
    ActiveWindow.DisplayGridlines = False


    
End Function

'*****************************************************


Function DataForRelGeral(WorkbookRelatorio As String)
    Dim WorkbookThis As String
    Dim UlinhaEste As Long
    Dim primeiraLinha As Integer
    Dim I As Long
    Dim J As Long
    
    WorkbookThis = "Comparativo de tabelas.xlsm" 'nome da planilha principal
    primeiraLinha = 15 ' seta primeira linah apos cabeçalho

    Workbooks(WorkbookThis).Activate
    
    
    '***************PLANILHAS DE ESTOQUE******************
    ' * Verifica total de produtos - ESTOQUE*
    UlinhaEste = Workbooks(WorkbookThis).Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' seta linha maxima para omparativo de tabelas.xlsm ->plan2-> coluna A - Magento Estoque
    For I = 2 To UlinhaEste
        Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row), 1).Value = I - 1
        Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row), 2).Value = Workbooks(WorkbookThis).Worksheets(7).Cells(I, 1).Value ' SKU
        Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 3).End(xlUp).Offset(1, 0).Row), 3).Value = Workbooks(WorkbookThis).Worksheets(7).Cells(I, 2).Value '
        Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 4).End(xlUp).Offset(1, 0).Row), 4).Value = Workbooks(WorkbookThis).Worksheets(7).Cells(I, 3).Value '
        Workbooks(WorkbookRelatorio).Worksheets(1).Cells(CInt(Workbooks(WorkbookRelatorio).Worksheets(1).Cells(Rows.Count, 5).End(xlUp).Offset(1, 0).Row), 5).Value = Workbooks(WorkbookThis).Worksheets(7).Cells(I, 4).Value '
    Next I


End Function

Function EstatisticasGeral(WorkbookRelatorio As String)
'Workbooks(WorkbookRelatorio).Activate
Dim formul As String
formul = "=CONT.SE(C15:C9999,2)"
'formul = COUNTIF(C15:C9999,2)
'formul = "=CONT"
    Workbooks(WorkbookRelatorio).Worksheets(1).Cells(7, 7).Value = formul

End Function
