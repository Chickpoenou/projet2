Option Explicit
' Feuille FICHE ÉTUDIANT

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo fin
    If Not Intersect(Target, Me.Range("C5")) Is Nothing Then RechercherEtudiant
fin:
    Application.EnableEvents = True
End Sub
