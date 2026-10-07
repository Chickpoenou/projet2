Option Explicit
' Feuille GROUPES : la composition se met à jour quand la matière ou le tableau change.

Private Sub Worksheet_Activate()
    On Error Resume Next
    RemplirGroupes
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo fin
    NormaliserCodes Me, Target
    If Not Intersect(Target, Me.Range("C5")) Is Nothing Then
        RemplirGroupes
    ElseIf Not Intersect(Target, Me.ListObjects(1).Range) Is Nothing Then
        RemplirGroupes
    End If
fin:
    Application.EnableEvents = True
End Sub
