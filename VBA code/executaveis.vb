'*****************************************************
'*                                                   *
'*              Planilha de estoque Extra            *
'*                                                   *
'*****************************************************
'*****************************************************
'* Verifica Divergencia de cadtro no DB
'*****************************************************
Sub SkuDivergeDBToExtra()
Dim SkupComp As Long
Dim Limite As Long
Dim I As Long
Dim WorksheetName As String
WorksheetName = "Nao cadastrados Magento"
If WorksheetExiststo(WorksheetName) <> 0 Then
Limite = Sheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    For I = 1 To Limite
        Sheets(WorksheetName).Cells(I, 1).Value = ""
        Sheets(WorksheetName).Cells(I, 2).Value = ""
        Sheets(WorksheetName).Cells(I, 3).Value = ""
    Next I
End If
'exiteSkuPlanxPlan(sheetIndex_I As Integer, sheetIndex_II As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
SkupComp = exiteSkuPlanxPlan(2, 3, 1, 1, 3, WorksheetName) - 1
MsgBox ("Existem " & SkupComp & "  Produtos com cadstro divergente")
Sheets(WorksheetName).Activate
End Sub
'*****************************************************

'*****************************************************
'* Verifica Divergencia de cadtro no DB
'*****************************************************
Sub SkuDivergeExtraToDB()
Dim SkupComp As Long
Dim Limite As Long
Dim I As Long
Dim WorksheetName As String
WorksheetName = "Nao cadastrados Extra"
If WorksheetExiststo(WorksheetName) <> 0 Then
Limite = Sheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    For I = 1 To Limite
        Sheets(WorksheetName).Cells(I, 1).Value = ""
        Sheets(WorksheetName).Cells(I, 2).Value = ""
        Sheets(WorksheetName).Cells(I, 3).Value = ""
    Next I
End If
'exiteSkuPlanxPlan(sheetIndex_I As Integer, sheetIndex_II As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
SkupComp = exiteSkuPlanxPlan(3, 2, 1, 1, 3, WorksheetName)
MsgBox ("Existem " & SkupComp - 1 & "  Produtos com cadstro divergente")
Sheets(WorksheetName).Activate
End Sub
'*****************************************************


'* Cria tabela com  divergencia de estoque dos produtos Habilitados
'*****************************************************
Sub DivergeHabilitadosItoII()
Dim SkupComp As Long
Dim Limite As Long
Dim I As Long
Dim WorksheetName As String
WorksheetName = "Divergente Habilitado"
If WorksheetExiststo(WorksheetName) <> 0 Then
Limite = Sheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    For I = 1 To Limite
        Sheets(WorksheetName).Cells(I, 1).Value = ""
        Sheets(WorksheetName).Cells(I, 2).Value = ""
        Sheets(WorksheetName).Cells(I, 3).Value = ""
    Next I
End If
'ComparaSkuCol(sheetIndex_I As Integer, sheetIndex_II As Integer, colSku As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
SkupComp = ComparaSkuCol(2, 3, 1, 3, 1, 3, WorksheetName)
MsgBox ("Existem " & SkupComp - 1 & " Produtos em estoque divergentes")
Sheets(WorksheetName).Activate
End Sub
'*****************************************************


'*****************************************************
'* Cria tabela com divergencia de Quantidade
'*****************************************************
Sub DivergeQuantidadeItoII()
Dim SkupComp As Long
Dim Limite As Long
Dim I As Long
Dim WorksheetName As String
WorksheetName = "Divergente Quantidade"
If WorksheetExiststo(WorksheetName) <> 0 Then
Limite = Sheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    For I = 1 To Limite
        Sheets(WorksheetName).Cells(I, 1).Value = ""
        Sheets(WorksheetName).Cells(I, 2).Value = ""
        Sheets(WorksheetName).Cells(I, 3).Value = ""
    Next I
End If

'ComparaSkuCol(sheetIndex_I As Integer, sheetIndex_II As Integer, colSku As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
SkupComp = ComparaSkuCol(2, 3, 1, 2, 1, 3, WorksheetName) - 1
MsgBox ("Existem " & SkupComp & "  Produtos com quantidades divergentes")
Sheets(WorksheetName).Activate
End Sub
'*****************************************************

'*****************************************************
'* Ajusta os dados da planilha com banco de dados conforme Partner Extra Estoque
'*****************************************************
Sub PartnerExtraEstoque()
Dim Limite As Long
Dim cont1 As Long
Dim cont2 As Long
Dim I As Long
Dim div As Long
'Function LineExist(Index As Integer, colCompara As Integer, Value, Limite As Long) As Boolean
Limite = Sheets(2).Cells(Rows.Count, 3).End(xlUp).Offset(1, 0).Row
    cont1 = 0
    cont2 = 0
    'div = Sheets(2).Cells(2, 2).Value / 10000
    For I = 2 To Limite
        'valor = IsNumeric(Sheets(2).Cells(2, 2).Value)
        If Sheets(2).Cells(I, 2).Value >= 1000 Then
            div = CLng(Sheets(2).Cells(I, 2).Value) / 10000
            Sheets(2).Cells(I, 2).Value = div
        Else
            div = CLng(Sheets(2).Cells(I, 2).Value) * 1
            Sheets(2).Cells(I, 2).Value = div
        End If
        'Sheets(2).Cells(I, 2).Value = valor
        If Sheets(2).Cells(I, 3).Value = "1" Then
           Sheets(2).Cells(I, 3).Value = "Y"
        Else
            Sheets(2).Cells(I, 3).Value = "N"
        End If
    Next I
    'MsgBox ("Habilitados: " & cont1 & " \nDesabilitados: " & cont2)
    Sheets(2).Activate
    MsgBox ("Dados Atualizados com sucesso")
    
End Sub


'*****************************************************
'*                                                   *
'*              Planilha de precos Extra             *
'*                                                   *
'*****************************************************

'*****************************************************
'* Ajusta os dados da planilha com banco de dados conforme Partner Extra preco
'*****************************************************
Sub PartnerExtraPreco()
Dim Limite As Long
Dim cont1 As Long
Dim cont2 As Long
Dim I As Long
Dim div As Double
'Function LineExist(Index As Integer, colCompara As Integer, Value, Limite As Long) As Boolean
Limite = Sheets(4).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    cont1 = 0
    cont2 = 0
    For I = 2 To Limite
        'coluna special_price
        If Sheets(4).Cells(I, 3).Value = "NULL" Or Sheets(4).Cells(I, 2).Value = "" Then
            Sheets(4).Cells(I, 3).Value = Sheets(4).Cells(I, 2).Value
        End If
        If Sheets(4).Cells(I, 2).Value > 0 Then
        div = CDbl(Sheets(4).Cells(I, 2).Value) / 10000
           Sheets(4).Cells(I, 2).Value = div
        Else
            div = CDbl(Sheets(4).Cells(I, 2).Value) * 1
            Sheets(4).Cells(I, 2).Value = div
        End If
        ' Coluna price
        If Sheets(4).Cells(I, 3).Value > 0 Then
        div = CDbl(Sheets(4).Cells(I, 3).Value) / 10000
           Sheets(4).Cells(I, 3).Value = div
        Else
            div = CDbl(Sheets(4).Cells(I, 3).Value) * 1
            Sheets(4).Cells(I, 3).Value = div
        End If
        
    Next I
    Sheets(4).Activate
    MsgBox ("Dados Atualizados com sucesso")
    
End Sub
'*****************************************************



'*****************************************************
'* Gera relatorio de divergencia em precos de produtos Magento X Extra
'*****************************************************
Sub DivergePrecoIToII()
Dim SkupComp As Long
Dim Limite As Long
Dim I As Long
Dim WorksheetName As String
WorksheetName = "Divergente Precos"
If WorksheetExiststo(WorksheetName) <> 0 Then
Limite = Sheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    For I = 1 To Limite
        Sheets(WorksheetName).Cells(I, 1).Value = ""
        Sheets(WorksheetName).Cells(I, 2).Value = ""
        Sheets(WorksheetName).Cells(I, 3).Value = ""
    Next I
End If

'ComparaSkuCol(sheetIndex_I As Integer, sheetIndex_II As Integer, colSku As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
SkupComp = ComparaSkuCol(4, 5, 1, 3, 1, 3, WorksheetName) - 1
MsgBox ("Existem " & SkupComp & "  Produtos com preços divergentes")
Sheets(WorksheetName).Activate
End Sub
'*****************************************************

'*****************************************************
'* Gera relatorio de divergencia em promocao de produtos Magento X Extra
'*****************************************************
Sub DivergePromocaoIToII()
Dim SkupComp As Long
'ComparaSkuCol(sheetIndex_I As Integer, sheetIndex_II As Integer, colSku As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
Dim Limite As Long
Dim I As Long
Dim WorksheetName As String
WorksheetName = "Divergente Promocao"
If WorksheetExiststo(WorksheetName) <> 0 Then
Limite = Sheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    For I = 1 To Limite
        Sheets(WorksheetName).Cells(I, 1).Value = ""
        Sheets(WorksheetName).Cells(I, 2).Value = ""
        Sheets(WorksheetName).Cells(I, 3).Value = ""
    Next I
End If

SkupComp = ComparaSkuCol(4, 5, 1, 3, 1, 3, WorksheetName) - 1
MsgBox ("Existem " & SkupComp & "  Produtos com promoções divergentes")
Sheets(WorksheetName).Activate
End Sub
'*****************************************************


'*****************************************************
'* Verifica Divergencia de cadastro no DB
'*****************************************************
Sub SkuDivergeDBToExtraPreco()
Dim SkupComp As Long
Dim Limite As Long
Dim I As Long
Dim WorksheetName As String
WorksheetName = "Nao cadastrados Magento P"
If WorksheetExiststo(WorksheetName) <> 0 Then
Limite = Sheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    For I = 1 To Limite
        Sheets(WorksheetName).Cells(I, 1).Value = ""
        Sheets(WorksheetName).Cells(I, 2).Value = ""
        Sheets(WorksheetName).Cells(I, 3).Value = ""
    Next I
End If

'exiteSkuPlanxPlan(sheetIndex_I As Integer, sheetIndex_II As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
SkupComp = exiteSkuPlanxPlan(4, 5, 1, 1, 3, WorksheetName) - 1
MsgBox ("Existem " & SkupComp & "  Produtos com cadstro divergente")
Sheets(WorksheetName).Activate
End Sub
'*****************************************************

'*****************************************************
'* Verifica Divergencia de cadtro no DB
'*****************************************************
Sub SkuDivergeExtraToDBPreco()
Dim SkupComp As Long
Dim Limite As Long
Dim I As Long
Dim WorksheetName As String
WorksheetName = "Nao cadastrados Extra P"
If WorksheetExiststo(WorksheetName) <> 0 Then
Limite = Sheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    For I = 1 To Limite
        Sheets(WorksheetName).Cells(I, 1).Value = ""
        Sheets(WorksheetName).Cells(I, 2).Value = ""
        Sheets(WorksheetName).Cells(I, 3).Value = ""
    Next I
End If
'exiteSkuPlanxPlan(sheetIndex_I As Integer, sheetIndex_II As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
SkupComp = exiteSkuPlanxPlan(5, 4, 1, 1, 3, WorksheetName)
MsgBox ("Existem " & SkupComp - 1 & "  Produtos com cadstro divergente")
Sheets(WorksheetName).Activate
End Sub
'*****************************************************
