Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next
    RemplirSyntheseExposes
End Sub
' Met à jour la colonne « Membres » lorsque la matière, le groupe ou le
' responsable de la remise change.

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject, z As Range, c As Range
    On Error GoTo fin
    NormaliserCodes Me, Target
    Set lo = Me.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set z = Intersect(Target, lo.DataBodyRange)
    If z Is Nothing Then Exit Sub
    If z.Rows.Count > 500 Then Exit Sub
    Application.EnableEvents = False
    Deverrouiller Me
    For Each c In Intersect(z.EntireRow, lo.DataBodyRange.Columns(1)).Cells
        MajMembresLigne lo, lo.ListRows(c.Row - lo.HeaderRowRange.Row).Range
    Next c
    RemplirSyntheseExposes
fin:
    Verrouiller Me
    Application.EnableEvents = True
End Sub
