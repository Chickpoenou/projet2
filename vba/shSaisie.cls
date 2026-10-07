Option Explicit
' Feuille COURS DU JOUR

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo fin
    If Target.CountLarge > 1000 Then Exit Sub
    If Not Intersect(Target, Me.Range(A_ID)) Is Nothing Then
        ChargerSeance Trim$(CStr(Me.Range(A_ID).Value))
    ElseIf Not Intersect(Target, Me.Range(A_CODE)) Is Nothing Then
        MatiereChangee
    ElseIf Not Intersect(Target, Me.Range(A_FILTRE)) Is Nothing Then
        AppliquerFiltre
    ElseIf Not Intersect(Target, ZoneStatuts) Is Nothing Then
        NormaliserStatuts Intersect(Target, ZoneStatuts)
    End If
fin:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.CountLarge > 1 Then Exit Sub
    If Not Intersect(Target, ZoneStatuts) Is Nothing Then
        Cancel = True
        CyclerStatut Target
    End If
End Sub
