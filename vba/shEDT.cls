Option Explicit
' Les listes déroulantes affichent « code – matière » : on ne garde que le code.

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo fin
    NormaliserCodes Me, Target
fin:
    Application.EnableEvents = True
End Sub
