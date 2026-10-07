Option Explicit
' Feuille FICHE ÉTUDIANT

Private Sub Worksheet_Activate()
    On Error Resume Next
    RemplirFiche
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo fin
    If Not Intersect(Target, Me.Range("C5")) Is Nothing Then
        RechercherEtudiant
    ElseIf Not Intersect(Target, Me.Range("C7")) Is Nothing Then
        RemplirFiche
    End If
fin:
    Application.EnableEvents = True
End Sub
