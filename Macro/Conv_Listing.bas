Sub Conv_Listing()
'
' Conv_Listing Macro
'

'
    Sheets("Conv_export").Select
    Columns("D:D").Select
    Selection.Copy
    Sheets("Listing").Select
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$B$1:$B$150").RemoveDuplicates Columns:=1, Header:=xlNo
End Sub