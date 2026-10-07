Option Explicit
' Feuille SUIVI ENSEIGNANTS

Private Sub Worksheet_Activate()
    On Error Resume Next
    RemplirSuivi
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo fin
    If Not Intersect(Target, Me.Range("Q6")) Is Nothing Then RemplirSuivi
fin:
    Application.EnableEvents = True
End Sub
