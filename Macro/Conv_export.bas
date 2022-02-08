Sub Conv_export()
'
' Conv_export Macro
'

'
    
    Sheets("Export").Select
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
    
    Columns("D:D").Select
    ActiveWorkbook.Worksheets("Export").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Export").Sort.SortFields.Add2 Key:=Range("D1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Export").Sort
        .SetRange Range("A1:O142")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:A").Select
    Columns("O:O").Select
    Columns("A:O").Select
    Columns("A:O").EntireColumn.AutoFit
    Columns("G:G").Select
    ActiveWorkbook.Worksheets("Export").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Export").Sort.SortFields.Add2 Key:=Range("G1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Export").Sort
        .SetRange Range("A1:O143")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("D:D").Select
    Selection.Cut
    Columns("A:A").Select
    ActiveSheet.Paste
    Columns("E:E").Select
    Selection.Cut
    Columns("B:B").Select
    ActiveSheet.Paste
    Columns("F:F").Select
    Selection.Cut
    Columns("C:C").Select
    ActiveSheet.Paste
    Columns("G:G").Select
    Selection.Cut
    Columns("D:D").Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Selection.Cut
    Columns("E:E").Select
    ActiveSheet.Paste

    Columns("I:I").Select
    Range("I48").Activate
    Selection.Cut
    Columns("I:N").Select
    Range("I48").Activate
    Application.CutCopyMode = False
    Selection.Cut
    Range("F1").Select
    ActiveSheet.Paste
    
    Sheets("Conv_export").Select
    Range("A2:K500").Select
    Selection.Delete
    Sheets("Export").Select
    Range("A1:K300").Select
    Selection.Copy
    Sheets("Conv_export").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("H2:H" & Range("H600").End(xlUp).Row).Select
    For Each h In Selection
    x = Val(h)
    Range(h.Address) = CDbl(x)
    Next

    Range("J2:J" & Range("J600").End(xlUp).Row).Select
    For Each j In Selection
    x = Val(j)
    Range(j.Address) = CDbl(x)
    Next
    
    Range("K2:K" & Range("K600").End(xlUp).Row).Select
    For Each k In Selection
    x = Val(k)
    Range(k.Address) = CDbl(x)
    Next
    
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"

End Sub