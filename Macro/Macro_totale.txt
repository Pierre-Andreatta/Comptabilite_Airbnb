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