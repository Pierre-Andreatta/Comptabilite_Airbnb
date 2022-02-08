Attribute VB_Name = "Module1"
Sub Fiche_clients()
'
' Fiche_clients Macro
'
Sheets("Listing").Select
Dim N As Integer
N = Range("B1: B500").Cells.Count - Application.WorksheetFunction.CountBlank(Range("B2: B500"))
Dim I As Integer
For I = 2 To N
    Sheets.Add.Move After:=Sheets(Sheets.Count)
'   Sheets(Sheets.Count).Name = Sheets("Listing").Cells(I, 1).Value & " - " & Sheets("Info").Cells(6, 3).Value
    Sheets(Sheets.Count).Name = Sheets("Listing").Cells(I, 1).Value
    
    Sheets("Conv_export").Select
    Range("A" & Sheets("Listing").Cells(I, 3).Value & ":I" & Sheets("Listing").Cells(I, 4).Value).Select
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("A15").Select
    ActiveSheet.Paste
    
    Sheets("Conv_export").Select
    Range("K" & Sheets("Listing").Cells(I, 3).Value & ":K" & Sheets("Listing").Cells(I, 4).Value).Select
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("I15").Select
    ActiveSheet.Paste
    
    Sheets("Conv_export").Select
    Range("A1:I1").Select
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("A14").Select
    ActiveSheet.Paste
    
    Sheets("Conv_export").Select
    Range("K1").Select
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("I14").Select
    ActiveSheet.Paste
    
    'Titre
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "COMPTES " & Sheets("Listing").Cells(I, 1).Value & " " & Sheets("Info").Cells(6, 3).Value & " " & Sheets("Info").Cells(9, 3).Value
    Range("H15:I32").Select
    Selection.NumberFormat = "$#,##0.00_);($#,##0.00)"
    
    'Total
    Range("G29").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("G30").Select
    ActiveCell.FormulaR1C1 = "Honoraires"
    Range("G31").Select
    ActiveCell.FormulaR1C1 = "Rotations"
    Range("G32").Select
    ActiveCell.FormulaR1C1 = "Virement"
    
    'Formule
    Range("H29").Select
    ActiveCell.FormulaLocal = "=SOMME(H15:H28)"
    Range("H30").Select
    ActiveCell.FormulaLocal = "=Info!C12*H29"
    Range("H31").Select
    ActiveCell.FormulaLocal = "=SOMME(I15:I28)"
    Range("H32").Select
    ActiveCell.Formula = "=H29-H30-H31"
Next


'
End Sub
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
Sub Macro_totale()
'
' Macro_totale
'

'
If Sheets("Export").Cells(4, 3).Value = "" Then
    
    If ThisWorkbook.Sheets.Count <= 5 Then
    
        Dim MonDossier As String
        MonDossier = ThisWorkbook.Path & "\Clients\" & Sheets("Info").Cells(6, 3).Value & "\"
        If DossierExiste(MonDossier) = False Then

            Call Conv_export
            Call Conv_Listing
            Call Fiche_clients
            Call Sauvegarde
            
        Else
        MsgBox "Le dossier " & Sheets("Info").Cells(6, 3).Value & " existe déjà"
        End If

    Else
    MsgBox ("Les fiches clients sont déjà éditées, supprimez les pour lancer l'édition")
    End If

Else
MsgBox ("L'export depuis Airbnb n'est pas brut")
End If

End Sub
Sub Bilan()
'
' Bilan Macro
'

'
 
End Sub
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
Public Function DossierExiste(MonDossier As String)


   If Len(Dir(MonDossier, vbDirectory)) > 0 Then
      DossierExiste = True
   Else
      DossierExiste = False
   End If
End Function
Sub TesteSiDossierExiste()
'par Excel-Malin.com ( https://excel-malin.com )

Dim MonDossier As String

MonDossier = ThisWorkbook.Path & "\Clients\" & Sheets("Info").Cells(6, 3).Value & "\"

    If DossierExiste(MonDossier) = True Then
        MsgBox "Le dossier existe..."
    Else
        MsgBox "Le dossier n'existe pas..."
    End If

End Sub
