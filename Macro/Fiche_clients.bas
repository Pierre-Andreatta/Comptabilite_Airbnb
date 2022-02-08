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