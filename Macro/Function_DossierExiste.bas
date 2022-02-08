Public Function DossierExiste(MonDossier As String)


   If Len(Dir(MonDossier, vbDirectory)) > 0 Then
      DossierExiste = True
   Else
      DossierExiste = False
   End If