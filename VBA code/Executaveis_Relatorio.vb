Sub GeraRelatorioExtra()

Dim NomeRelatorio As String
 
 'WorkbookThis = "Comparativo de tabelas.xlsm"
 Workbooks("Comparativo de tabelas.xlsm").Activate
 Call PartnerExtraEstoque
 Call PartnerExtraPreco

 NomeRelatorio = CreatWorkbook("Relatorio Extra") ' Recupera nome da planilha criada
 CreatHeaderExtra (NomeRelatorio)
 Workbooks(NomeRelatorio).Activate
 CompDataForRel (NomeRelatorio)
 Workbooks(NomeRelatorio).Save
 
 'Workbooks(NomeRelatorio).Close
  MsgBox ("Relatório criado com sucesso!")
  Workbooks(NomeRelatorio).Activate
End Sub


Sub GeraRelatorioB2W()

Dim NomeRelatorio As String
 
 'WorkbookThis = "Comparativo de tabelas.xlsm"
 Workbooks("Comparativo de tabelas.xlsm").Activate
 Call PartnerExtraEstoque
 Call PartnerExtraPreco

 NomeRelatorio = CreatWorkbook("Relatorio B2W") ' Recupera nome da planilha criada
 CreatHeaderB2W (NomeRelatorio)
 Workbooks(NomeRelatorio).Activate
 CompDataForRelB2W (NomeRelatorio)
 Workbooks(NomeRelatorio).Save
 
 'Workbooks(NomeRelatorio).Close
  MsgBox ("Relatório criado com sucesso!")
  Workbooks(NomeRelatorio).Activate
End Sub


Sub GeraRelatorioGeral()

Dim NomeRelatorio As String
 
 'WorkbookThis = "Comparativo de tabelas.xlsm"
 Workbooks("Comparativo de tabelas.xlsm").Activate
 Call PartnerExtraEstoque
 Call PartnerExtraPreco

 NomeRelatorio = CreatWorkbook("Relatorio Geral") ' Recupera nome da planilha criada
 CreatHeaderRelGeral (NomeRelatorio)
 Workbooks(NomeRelatorio).Activate
 'CompDataForRelB2W (NomeRelatorio)
 DataForRelGeral (NomeRelatorio)
 EstatisticasGeral (NomeRelatorio)
 Workbooks(NomeRelatorio).Save
 
 'Workbooks(NomeRelatorio).Close
  MsgBox ("Relatório criado com sucesso!")
  'Workbooks(NomeRelatorio).Activate
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
     ThisWorkbook.Save
End Sub

Sub testa()
' 'WorkbookThis = "Comparativo de tabelas.xlsm"
Dim st As String
st = "Man_2878"

MsgBox (UCase(st))
End Sub
