'*****************************************************
'* Funcao verifica se planilha existe e rotorna o indice, retorna 0 caso não exista
'* Input: nome da tabela
'* output: Integer Indice da tabela
'*****************************************************
Function WorksheetExiststo(ByValWorksheetName As String) As Integer
Dim WS_Count As Integer
Dim I As Long
WorksheetExiststo = 0
WS_Count = Sheets.Count
For I = 1 To WS_Count
    If ActiveWorkbook.Worksheets(I).Name = ByValWorksheetName Then
        WorksheetExiststo = I
     End If
 Next I
End Function
'*****************************************************

'*****************************************************
'* Funcao percore linhas da tabela e compara os valores da coluna
'* Input': Indice da tabela, indice da coluna, valor para comparar, tamanho da Range
'* output: Booleano
'*****************************************************
Function LineExist(Index As Integer, colCompara As Integer, Value, Limite As Long) As Boolean
Dim I As Long
    LineExist = False
    For I = 1 To Limite
        If Sheets(Index).Cells(I, colCompara).Value = Value Then
            LineExist = True
            Exit For
        End If
    Next I
    
End Function
'*****************************************************


'*****************************************************
'* Funcao adiciona tabela de comparação caso não exista , caso exista atualiza
'* Input': String nome da tabela
'* Output: Integer nome da tabela criada
'*****************************************************
Function AddSheet(WorksheetName As String) As Integer
    Dim pasta As Worksheet
    Dim Linha As Integer
    
    If WorksheetExiststo(WorksheetName) = 0 Then
        ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
        Set pasta = Application.Worksheets(Sheets.Count)
        pasta.Name = WorksheetName
 End If
 AddSheet = WorksheetExiststo(WorksheetName)
End Function

'*****************************************************

'*****************************************************
'* Funcao Verifica se os dados da sheetIndex_I->colCompara existentem na sheetIndex_II->colCompara e escreve na tabela
'* Input': sheetIndex_I,sheetIndex_II,colCompara, colIni, colEnd,sheetTitle
'* Output: Long Numero de linhas da tabela
'*****************************************************
Function exiteSkuPlanxPlan(sheetIndex_I As Integer, sheetIndex_II As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
    Dim valor
    Dim Linha_1 As Long
    Dim Linha_2 As Long
    Dim Linha As Long
    Dim newSheet As Integer
    Dim MaisLinha As Long
    Dim I As Long
    Dim coliniend As Long
    Dim J As Long
    
    Linha_1 = Sheets(sheetIndex_I).Cells(Rows.Count, colCompara).End(xlUp).Offset(1, 0).Row
    Linha_2 = Sheets(sheetIndex_II).Cells(Rows.Count, colCompara).End(xlUp).Offset(1, 0).Row
    'cria tabela
    newSheet = AddSheet(sheetTitle)
    'cria titulos referente a tabela 1
    For coliniend = colIni To colEnd
        Sheets(newSheet).Cells(1, coliniend).Value = Sheets(sheetIndex_I).Cells(1, coliniend).Value
    Next coliniend
        
    For I = 2 To Linha_1
    valor = Sheets(sheetIndex_I).Cells(I, colCompara).Value
        If LineExist(sheetIndex_II, colCompara, valor, Linha_2) = False Then
            For J = 2 To Linha
            Next J
            Linha = J
            For coliniend = colIni To colEnd
                Sheets(newSheet).Cells(Linha, coliniend).Value = Sheets(sheetIndex_I).Cells(I, coliniend).Value
            Next coliniend
        End If
 Next I
 
 exiteSkuPlanxPlan = Sheets(newSheet).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
End Function

'*****************************************************
'* Funcao Verifica se os dados indexados pela col sku da sheetIndex_I->colCompara existentem na sheetIndex_II->colCompara e escreve na tabela
'* Input': sheetIndex_I,sheetIndex_II,colSku colCompara, colIni, colEnd,sheetTitle
'* Output: Long Numero de linhas da tabela
'*****************************************************


'*****************************************************
Function ComparaSkuCol(sheetIndex_I As Integer, sheetIndex_II As Integer, colSku As Integer, colCompara As Integer, colIni As Integer, colEnd As Integer, sheetTitle As String) As Long
    Dim Linha_1 As Long
    Dim Linha_2 As Long
    Dim Linha As Long
    Dim newSheet As Integer
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim TestaSkuCol As Long
    
    Linha_1 = Sheets(sheetIndex_I).Cells(Rows.Count, colCompara).End(xlUp).Offset(1, 0).Row
    Linha_2 = Sheets(sheetIndex_II).Cells(Rows.Count, colCompara).End(xlUp).Offset(1, 0).Row
    'cria tabela
    newSheet = AddSheet(sheetTitle)
    'cria titulos referente a tabela
    Sheets(newSheet).Cells(1, 1).Value = "SKU"
    Sheets(newSheet).Cells(1, 2).Value = "MAGENTO"
    Sheets(newSheet).Cells(1, 3).Value = "EXTRA"
'    For coliniend = colIni To colEnd
 '       Sheets(newSheet).Cells(1, coliniend).Value = Sheets(sheetIndex_I).Cells(1, coliniend).Value
  '  Next coliniend
    
    For I = 2 To Linha_1
        For K = 2 To Linha_2
            If Sheets(sheetIndex_I).Cells(I, colSku).Value = Sheets(sheetIndex_II).Cells(K, colSku).Value And Sheets(sheetIndex_I).Cells(I, colCompara).Value <> Sheets(sheetIndex_II).Cells(K, colCompara).Value Then
                TestaSkuCol = Sheets(newSheet).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
                Sheets(newSheet).Cells(TestaSkuCol, 1).Value = Sheets(sheetIndex_I).Cells(I, colSku).Value
                Sheets(newSheet).Cells(TestaSkuCol, 2).Value = Sheets(sheetIndex_I).Cells(I, colCompara).Value
                'Sheets(newSheet).Cells(TestaSkuCol, 3).Value = Sheets(sheetIndex_II).Cells(K, colSku).Value
                Sheets(newSheet).Cells(TestaSkuCol, 3).Value = Sheets(sheetIndex_II).Cells(K, colCompara).Value
                'Sheets(newSheet).Cells(TestaSkuCol, 5).Value = "OK"
                Exit For
            End If
        Next K
 Next I
 ComparaSkuCol = Sheets(newSheet).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
End Function
