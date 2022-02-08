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
