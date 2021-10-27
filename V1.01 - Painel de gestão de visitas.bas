Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False

    Call P_Visita
    Call BD_Cons_Tratada
    Call BV_Inicial
    Call Ultimas_Visitas
    Call Visitas_Canceladas
    Call Base_Tratada
    Call Base_de_Visitas
    Call Visitas_D_1
    
    Sheets("MACROS").Select
    Range("B8").Select

    Application.ScreenUpdating = True
    
End Sub

Sub P_Visita()

    Application.ScreenUpdating = False
    
    Sheets("P. VISITA").Select
    Columns("B:D").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("BASE TRATADA").Select
    ActiveSheet.Range("$B$6:$AE$15000").AutoFilter Field:=28, Criteria1:=1, _
        Operator:=11, Criteria2:=0, SubField:=0
    Range("C6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("P. VISITA").Select
    Range("B2").Select
    ActiveSheet.Paste
    Sheets("BASE TRATADA").Select
    Range("AC6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("P. VISITA").Select
    Range("C2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("BASE TRATADA").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B7").Select
    Sheets("BD CADASTRO").Select
    ActiveSheet.Range("$B$5:$E$50000").AutoFilter Field:=4, Criteria1:="<>-", _
        Operator:=xlAnd
    Range("B5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("P. VISITA").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("BD CADASTRO").Select
    Range("E5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("P. VISITA").Select
    Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("BD CADASTRO").Select
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B6").Select
    Sheets("P. VISITA").Select
    Range("I3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("I:J").Select
    Range("I302").Activate
    Selection.ClearContents
    Range("C2").Select
    Selection.Copy
    Range("D2").Select
    ActiveSheet.Paste
    Range("D3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*1"
    Range("D3").Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$B$2:$D$50000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Range("B3").Select
    
    Application.ScreenUpdating = True

End Sub

Sub BD_Cons_Tratada()

    Application.ScreenUpdating = False
    
    Sheets("BD CONS TRATADA").Select
    Columns("B:C").Select
    Selection.ClearContents
    Sheets("BD CONS").Select
    Range("D5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BD CONS TRATADA").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("B2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Sheets("BD CONS").Select
    Range("I5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BD CONS TRATADA").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Columns("D:XFD").Select
    Range("D2").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("B3").Select
    
    Application.ScreenUpdating = True

End Sub

Sub BV_Inicial()

    Application.ScreenUpdating = False
    
    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BV - INICIAL").Select
    Range("B6").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C5").Value > 0 Then
        linhaf = linhai - Range("C5").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C5").Value < 0 Then
        linhaf = linhai + Range("C5").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B7").Select
    Sheets("BD - VISITAS").Select
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BV - INICIAL").Select
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B7").Select
    Application.CutCopyMode = False
     Range("AA7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("AA8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B7").Select
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True

End Sub

Sub Ultimas_Visitas()

    Application.ScreenUpdating = False

    Sheets("BV - INICIAL").Select
    ActiveSheet.Range("$B$6:$AG$30000").AutoFilter Field:=27, Criteria1:= _
        "<>Visita Cancelada", Operator:=xlAnd
    Range("AC6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("ÚLTIMAS VISITAS").Select
    Range("H5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("BV - INICIAL").Select
    Range("AA6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ÚLTIMAS VISITAS").Select
    Range("I5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B5").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:I").Select
    Range("H11049").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("B6:C6").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("ÚLTIMAS VISITAS").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ÚLTIMAS VISITAS").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("C5:C50000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ÚLTIMAS VISITAS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B6").Select
    ActiveSheet.Range("$B$5:$C$50000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Range("B6").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("C6").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.NumberFormat = "m/d/yyyy"
    Range("B6").Select
    Sheets("BV - INICIAL").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B7").Select

    Application.ScreenUpdating = True
    
End Sub

Sub Visitas_Canceladas()

    Application.ScreenUpdating = False
    
    Sheets("VISITAS CANCELADAS").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("BV - INICIAL").Select
    ActiveSheet.Range("$B$6:$AG$50000").AutoFilter Field:=27, Criteria1:= _
        "=Visita Cancelada", Operator:=xlAnd
    Range("AC6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("VISITAS CANCELADAS").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("B5").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B6").Select
    ActiveSheet.Range("$B$5:$D$10000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Sheets("BV - INICIAL").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B7").Select
    
    Application.ScreenUpdating = True

End Sub

Sub Base_Tratada()

Application.ScreenUpdating = False

    Sheets("BASE TRATADA").Select
    Range("R7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("R8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B7").Select

Application.ScreenUpdating = True

End Sub

Sub Base_de_Visitas()

Application.ScreenUpdating = False

    Sheets("BASE TRATADA").Select
    Range("B7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE DE VISITAS").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Sheets("BASE TRATADA").Select
    Range("B7").Select
    Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

Sub Visitas_D_1()

Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Integer
    Dim final As Integer
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("VISITAS D-1").Range("C1").Value)
    final = Abs(Worksheets("VISITAS D-1").Range("B1").Value)
 
    Do While atual > final
        Sheets("VISITAS D-1").Select
        Range("B3").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B3").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("VISITAS D-1").Range("C1").Value)
        final = Abs(Worksheets("VISITAS D-1").Range("B1").Value)
    Loop

    Sheets("VISITAS D-1").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BV - INICIAL").Select
    Range("AB1").Select
    ActiveSheet.Range("$B$6:$AG$50000").AutoFilter Field:=32, Criteria1:="=1", _
        Operator:=xlAnd
    Range("AA6:AF6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("VISITAS D-1").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("BV - INICIAL").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B7").Select
    Sheets("VISITAS D-1").Select
    ActiveWorkbook.Worksheets("VISITAS D-1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("VISITAS D-1").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("C3:C7000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("VISITAS D-1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B4").Select
    ActiveSheet.Range("$B$3:$G$7000").RemoveDuplicates Columns:=Array(3, 5), _
        Header:=xlYes
    ActiveWorkbook.RefreshAll
    Range("B4").Select

Application.ScreenUpdating = True

End Sub

Sub Arquivo_Envio()

    Application.ScreenUpdating = False

    ActiveWorkbook.Save
    ChDir _
        ActiveWorkbook.Path
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("MACROS").Range("C17").Value & " - Gestão de Visitas - Base Ativa RMV - Dados até dia " & Worksheets("MACROS").Range("C18").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Sheets(Array("MACROS", "BD CADASTRO", "BD - PROTOCOLOS D-1", "BD CONS", _
        "BD - VISITAS", "CONSULTOR", "BD - VENDAS", "P. VISITA", "BD CONS TRATADA", _
        "BV - INICIAL")).Select
    Sheets("BV - INICIAL").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets(Array("MACROS", "BD CADASTRO", "BD - PROTOCOLOS D-1", "BD CONS", _
        "BD - VISITAS", "CONSULTOR", "BD - VENDAS", "P. VISITA", "BD CONS TRATADA", _
        "BV - INICIAL", "ÚLTIMAS VISITAS", "VISITAS CANCELADAS", "BASE TRATADA")).Select
    Sheets("BASE TRATADA").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets(Array("MACROS", "BD CADASTRO", "BD - PROTOCOLOS D-1", "BD CONS", _
        "BD - VISITAS", "CONSULTOR", "BD - VENDAS", "P. VISITA", "BD CONS TRATADA", _
        "BV - INICIAL", "ÚLTIMAS VISITAS", "VISITAS CANCELADAS", "BASE TRATADA", _
        "TD GRÁFICOS", "GRÁFICOS")).Select
    Sheets("GRÁFICOS").Activate
    ActiveWindow.SelectedSheets.Delete
    Sheets("TD").Select
    ActiveWindow.SelectedSheets.Visible = False
    Range("B1:C1").Select
    Selection.ClearContents
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE DE VISITAS").Select
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("QUADRO DE PERFORMANCE").Select
    Range("B5").Select
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWorkbook.Save

    Application.ScreenUpdating = True

End Sub
