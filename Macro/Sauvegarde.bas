Sub Sauvegarde()
'
' Sauvegarde Macro
'

'
MkDir (ThisWorkbook.Path & "\Clients\" & Sheets("Info").Cells(6, 3).Value & "\")

Dim D As Variant
D = Sheets("Info").Cells(6, 3).Value
Dim T As Integer
T = ThisWorkbook.Sheets.Count
Dim F As Integer
For F = 6 To T

    Sheets(F).Copy
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\Clients\" & D & "\" & Sheets(Sheets.Count).Name & " - " & D & ".xlsx"
    ActiveWorkbook.Close

Next

End Sub