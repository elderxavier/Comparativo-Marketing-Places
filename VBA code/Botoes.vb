Sub Elipse1_Clique()
'
' Elipse1_Clique Macro
'

'
    Range("F8").Select
    ActiveWindow.SmallScroll Down:=-6
    Sheets("intrucoes de uso").Select
    Range("D7").Select
    Selection.OnAction = "SkuDivergeIItoI"
    Range("G7").Select
    ActiveWindow.SmallScroll Down:=-9
    Application.Goto Reference:="SkuDivergeIItoI"
    Selection.OnAction = "SkuDivergeIItoI"
    Range("E8").Select
    Application.Goto Reference:="Elipse1_Clique"
    'Call SkuDivergeIItoI
    Selection.OnAction = "SkuDivergeIItoI"
    Range("D8").Select
    ActiveWindow.SmallScroll Down:=-6
    Selection.OnAction = "Elipse1_Clique"
    Range("D10").Select
    ActiveWindow.SmallScroll Down:=-12
    Sheets("Skus Inexistentes na Plan3").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("comp col3 na Plan3").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("comp col na Plan3").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("intrucoes de uso").Select
    Range("G3").Select
    Selection.OnAction = "SkuDivergeIItoI"
    Range("F9").Select
End Sub
    Selection.OnAction = "Elipse1_Clique"
    Range("B10").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("I14").Select
    Selection.ShapeRange.Item(1).Hyperlink.Follow NewWindow:=False, AddHistory _
        :=True
    Range("C9").Select
    ActiveWorkbook.Save
    Range("B11").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("A9").Select
    ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False _
        , DisplayAsIcon:=False, Left:=28.5, Top:=174, Width:=72, Height:=24). _
        Select
    Range("C14").Select
    ActiveSheet.Shapes("CommandButton1").Select
    Range("A11").Select
    ActiveSheet.Shapes("CommandButton1").Select
    Range("B13").Select
    ActiveSheet.Shapes("CommandButton1").Select
    Sheets("intrucoes de uso").Select
    Range("A11:B14").Select
    Range("B11").Activate
    Application.WindowState = xlMinimized
    Range("D15").Select

