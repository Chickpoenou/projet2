Option Explicit
' ==========================================================================
'  modSuivi : rapports, remises, exposés et fiche étudiant
' ==========================================================================

' ----------------------------------------------------------------- RAPPORTS
Public Sub NouveauRapport()
    Dim lo As ListObject, r As Range
    Set lo = Tbl("tblRapports")
    Deverrouiller lo.Parent
    Set r = NouvelleLigne(lo)
    r.Cells(1, Col(lo, "ID")).Value = NouvelID(lo, "R", 3)
    r.Cells(1, Col(lo, "Type")).Value = "Individuel"
    r.Cells(1, Col(lo, "Date consigne")).Value = Date
    r.Cells(1, Col(lo, "Heure limite")).Value = TimeSerial(23, 59, 0)
    Verrouiller lo.Parent
    lo.Parent.Activate
    r.Cells(1, Col(lo, "Code matière")).Select
    MsgBox "Rapport " & r.Cells(1, 1).Value & " créé." & vbLf & vbLf & _
           "1. Choisissez la matière, saisissez l'intitulé, le type (Individuel / Groupe) et la date limite." & vbLf & _
           "2. Cliquez ensuite sur « Générer les remises attendues ».", vbInformation
End Sub

Private Function RapportSelectionne() As Range
    Dim lo As ListObject, z As Range, i As Long, id As String
    Set lo = Tbl("tblRapports")
    If Not lo.DataBodyRange Is Nothing And ActiveSheet Is lo.Parent Then
        Set z = Intersect(ActiveCell, lo.DataBodyRange)
        If Not z Is Nothing Then
            i = z.Row - lo.HeaderRowRange.Row
            Set RapportSelectionne = lo.ListRows(i).Range
            Exit Function
        End If
    End If
    id = InputBox("Numéro du rapport (ex. R001) :", "Rapport")
    If id <> "" Then Set RapportSelectionne = LigneParID(lo, UCase$(Trim$(id)))
End Function

Public Sub GenererRemises()
    Dim loR As ListObject, loM As ListObject, rr As Range, id As String, code As String, typ As String
    Dim existants As Object, v As Variant, i As Long, n As Long, qui As Variant
    Dim liste As Variant, cibles As New Collection, r As Range
    Set loR = Tbl("tblRapports")
    Set loM = Tbl("tblRemises")
    Set rr = RapportSelectionne()
    If rr Is Nothing Then
        MsgBox "Sélectionnez une ligne du tableau des rapports.", vbExclamation
        Exit Sub
    End If
    id = CStr(rr.Cells(1, Col(loR, "ID")).Value)
    code = Trim$(CStr(rr.Cells(1, Col(loR, "Code matière")).Value))
    typ = Trim$(CStr(rr.Cells(1, Col(loR, "Type")).Value))
    If id = "" Or code = "" Or Not MatiereExiste(code) Then
        MsgBox "Le rapport doit avoir un numéro et une matière valides.", vbExclamation
        Exit Sub
    End If
    If Not IsDate(rr.Cells(1, Col(loR, "Date limite")).Value) Then
        If MsgBox("Aucune date limite n'est saisie : le statut restera « En attente »." & vbLf & "Continuer ?", _
                  vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If

    If typ = "Groupe" Then
        For Each qui In GroupesMatiere(code)
            cibles.Add CStr(qui)
        Next qui
    Else
        liste = ListeEtudiantsMatiere(code)
        For i = 1 To NbLignes(liste)
            cibles.Add CStr(liste(i, 1))
        Next i
    End If

    Set existants = CreateObject("Scripting.Dictionary")
    existants.CompareMode = vbTextCompare
    v = Donnees(loM)
    For i = 1 To NbLignes(v)
        If CStr(v(i, 1)) = id Then existants(CStr(v(i, Col(loM, "Remis par")))) = 1
    Next i

    DebutTraitement
    Deverrouiller loM.Parent
    For Each qui In cibles
        If Not existants.Exists(CStr(qui)) Then
            Set r = NouvelleLigne(loM)
            r.Cells(1, Col(loM, "ID rapport")).Value = id
            r.Cells(1, Col(loM, "Remis par")).Value = qui
            MajMembresLigne loM, r
            n = n + 1
        End If
    Next qui
    Verrouiller loM.Parent
    FinTraitement
    MsgBox n & " remise(s) attendue(s) ajoutée(s) pour " & id & " (" & typ & ")." & vbLf & _
           "Sur la feuille REMISES, sélectionnez les lignes et cliquez sur « Marquer remis » à chaque dépôt.", vbInformation
End Sub

Public Sub MarquerRemis()
    Dim lo As ListObject, z As Range, i As Long, n As Long, r As Range
    Set lo = Tbl("tblRemises")
    If lo.DataBodyRange Is Nothing Or Not ActiveSheet Is lo.Parent Then Exit Sub
    Set z = Intersect(Selection, lo.DataBodyRange)
    If z Is Nothing Then
        MsgBox "Sélectionnez la ou les lignes des rapports remis.", vbInformation
        Exit Sub
    End If
    Application.EnableEvents = False
    Deverrouiller lo.Parent
    For i = 1 To lo.ListRows.Count
        Set r = lo.ListRows(i).Range
        If Not Intersect(z.EntireRow, r) Is Nothing And Not r.EntireRow.Hidden Then
            If Trim$(CStr(r.Cells(1, Col(lo, "Date remise")).Value)) = "" Then
                r.Cells(1, Col(lo, "Date remise")).Value = Date
                r.Cells(1, Col(lo, "Heure remise")).Value = TimeSerial(Hour(Now), Minute(Now), 0)
                n = n + 1
            End If
        End If
    Next i
    Verrouiller lo.Parent
    Application.EnableEvents = True
    MsgBox n & " remise(s) enregistrée(s) le " & Format$(Date, "dd/mm/yyyy") & " à " & Format$(Time, "hh:nn") & ".", vbInformation
End Sub

Private Sub FiltrerStatut(lo As ListObject, ByVal critere As String)
    Deverrouiller lo.Parent
    If critere = "" Then
        If lo.AutoFilter Is Nothing Then lo.Range.AutoFilter
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    Else
        lo.Range.AutoFilter Field:=Col(lo, "Statut"), Criteria1:=critere
    End If
    Verrouiller lo.Parent
End Sub

Public Sub AfficherNonRemis(): FiltrerStatut Tbl("tblRemises"), "Non remis": End Sub
Public Sub AfficherEnRetard(): FiltrerStatut Tbl("tblRemises"), "En retard": End Sub
Public Sub ToutAfficherRemises(): FiltrerStatut Tbl("tblRemises"), "": End Sub
Public Sub AfficherExposesRestants()
    Dim lo As ListObject
    Set lo = Tbl("tblExposes")
    Deverrouiller lo.Parent
    lo.Range.AutoFilter Field:=Col(lo, "Statut"), Criteria1:="<>Présenté"
    Verrouiller lo.Parent
End Sub
Public Sub ToutAfficherExposes(): FiltrerStatut Tbl("tblExposes"), "": End Sub

' ----------------------------------------------------------------- EXPOSÉS
Public Sub GenererExposes()
    Dim code As String, lo As ListObject, v As Variant, i As Long, n As Long
    Dim existants As Object, g As Variant, r As Range
    code = CodeDe(CStr(shExposes.Range("C4").Value))
    If code = "" Or Not MatiereExiste(code) Then
        MsgBox "Choisissez d'abord la matière dans la case « Matière » (en haut).", vbExclamation
        Exit Sub
    End If
    Set lo = Tbl("tblExposes")
    Set existants = CreateObject("Scripting.Dictionary")
    existants.CompareMode = vbTextCompare
    v = Donnees(lo)
    For i = 1 To NbLignes(v)
        If StrComp(CStr(v(i, 1)), code, vbTextCompare) = 0 Then existants(CStr(v(i, Col(lo, "Groupe")))) = 1
    Next i
    DebutTraitement
    Deverrouiller lo.Parent
    For Each g In GroupesMatiere(code)
        If Not existants.Exists(CStr(g)) Then
            Set r = NouvelleLigne(lo)
            r.Cells(1, Col(lo, "Code matière")).Value = code
            r.Cells(1, Col(lo, "Groupe")).Value = g
            r.Cells(1, Col(lo, "Statut")).Value = "À venir"
            MajMembresLigne lo, r
            n = n + 1
        End If
    Next g
    Verrouiller lo.Parent
    FinTraitement
    RemplirSyntheseExposes
    MsgBox n & " exposé(s) ajouté(s) pour " & code & "." & vbLf & _
           "Complétez le thème et la date prévue de chaque groupe.", vbInformation
End Sub

Public Sub MarquerPresente()
    Dim lo As ListObject, z As Range, i As Long, n As Long, r As Range
    Set lo = Tbl("tblExposes")
    If lo.DataBodyRange Is Nothing Or Not ActiveSheet Is lo.Parent Then Exit Sub
    Set z = Intersect(Selection, lo.DataBodyRange)
    If z Is Nothing Then
        MsgBox "Sélectionnez la ligne du groupe qui a présenté.", vbInformation
        Exit Sub
    End If
    Application.EnableEvents = False
    Deverrouiller lo.Parent
    For i = 1 To lo.ListRows.Count
        Set r = lo.ListRows(i).Range
        If Not Intersect(z.EntireRow, r) Is Nothing And Not r.EntireRow.Hidden Then
            r.Cells(1, Col(lo, "Statut")).Value = "Présenté"
            If Trim$(CStr(r.Cells(1, Col(lo, "Date de passage")).Value)) = "" Then r.Cells(1, Col(lo, "Date de passage")).Value = Date
            n = n + 1
        End If
    Next i
    Verrouiller lo.Parent
    Application.EnableEvents = True
    RemplirSyntheseExposes
    MsgBox n & " exposé(s) marqué(s) « Présenté ».", vbInformation
End Sub

' ---------------------------------------------------------- FICHE ÉTUDIANT
Public Sub RechercherEtudiant()
    Dim txt As String, v As Variant, lo As ListObject, i As Long, c As Long
    Dim trouves As New Collection, liste As String, choix As String, k As Long
    txt = SansAccents(LCase$(Trim$(CStr(shFiche.Range("C5").Value))))
    If txt = "" Then Exit Sub
    Set lo = Tbl("tblEtudiants")
    c = Col(lo, "Nom et prénoms")
    v = Donnees(lo)
    For i = 1 To NbLignes(v)
        If InStr(1, SansAccents(LCase$(CStr(v(i, c)))), txt, vbBinaryCompare) > 0 Then trouves.Add CStr(v(i, c))
    Next i
    Select Case trouves.Count
        Case 0
            MsgBox "Aucun étudiant ne correspond à « " & shFiche.Range("C5").Value & " ».", vbExclamation
            Exit Sub
        Case 1
            choix = trouves(1)
        Case Else
            For i = 1 To trouves.Count
                If i <= 25 Then liste = liste & i & ". " & trouves(i) & vbLf
            Next i
            If trouves.Count > 25 Then liste = liste & "... (" & trouves.Count - 25 & " autres, précisez la recherche)" & vbLf
            choix = InputBox(trouves.Count & " étudiants correspondent :" & vbLf & vbLf & liste & vbLf & _
                             "Tapez le numéro de l'étudiant :", "Fiche étudiant", "1")
            k = Val(choix)
            If k < 1 Or k > trouves.Count Then Exit Sub
            choix = trouves(k)
    End Select
    Application.EnableEvents = False
    shFiche.Range("C7").Value = choix
    Application.EnableEvents = True
    RemplirFiche
End Sub

Public Sub ExporterFichePDF()
    Dim nom As String, dossier As String, f As String, derL1 As Long, derL2 As Long
    nom = Trim$(CStr(shFiche.Range("C7").Value))
    If nom = "" Then
        MsgBox "Choisissez d'abord un étudiant.", vbExclamation
        Exit Sub
    End If
    On Error GoTo erreur
    RemplirFiche
    ' Zone 1 : identité, bilan, matières, absences et historique (B:Q)
    ' Zone 2 : rapports et exposés (S:AI)
    derL1 = DerniereLigne(shFiche, 2, 17)
    derL2 = DerniereLigne(shFiche, 19, 35)
    Deverrouiller shFiche
    shFiche.PageSetup.PrintArea = "$B$2:$Q$" & derL1 & ",$S$5:$AI$" & derL2
    Verrouiller shFiche
    dossier = DossierExport("Fiches étudiants")
    f = dossier & "\Fiche_" & NomFichier(nom) & "_" & Format$(Date, "yyyy-mm-dd") & ".pdf"
    shFiche.ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Exit Sub
erreur:
    MsgBox "Export impossible : " & Err.Description, vbExclamation
End Sub

Private Function DerniereLigne(ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long) As Long
    Dim c As Long, r As Long
    DerniereLigne = 7
    For c = c1 To c2
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > DerniereLigne Then DerniereLigne = r
    Next c
End Function
