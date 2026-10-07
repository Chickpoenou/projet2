Option Explicit
' ==========================================================================
'  modGroupes : composition des groupes (par défaut ou propre à une matière)
'  - Groupes par défaut : colonne « Groupe » de la feuille ÉTUDIANTS.
'  - Groupes spécifiques : lignes de la feuille GROUPES pour un code matière.
'    Dès qu'une matière y possède au moins une ligne, seuls ces étudiants
'    et ces groupes sont utilisés pour cette matière.
' ==========================================================================

Public Function GroupesSpecifiques(ByVal code As String) As Boolean
    Dim v As Variant, i As Long
    If code = "" Then Exit Function
    v = Donnees(Tbl("tblGroupes"))
    For i = 1 To NbLignes(v)
        If StrComp(CStr(v(i, 1)), code, vbTextCompare) = 0 Then GroupesSpecifiques = True: Exit Function
    Next i
End Function

' Étudiants suivant la matière : tableau (1..n, 1..2) = nom, groupe. Empty si aucun.
Public Function ListeEtudiantsMatiere(ByVal code As String) As Variant
    Dim ve As Variant, vg As Variant, i As Long, n As Long, res() As Variant
    Dim dic As Object, k As Variant
    Dim loE As ListObject, cNom As Long, cGrp As Long
    Set loE = Tbl("tblEtudiants")
    cNom = Col(loE, "Nom et prénoms"): cGrp = Col(loE, "Groupe")
    ve = Donnees(loE)

    If GroupesSpecifiques(code) Then
        Set dic = CreateObject("Scripting.Dictionary")
        dic.CompareMode = vbTextCompare
        vg = Donnees(Tbl("tblGroupes"))
        For i = 1 To NbLignes(vg)
            If StrComp(CStr(vg(i, 1)), code, vbTextCompare) = 0 And Trim$(CStr(vg(i, 2))) <> "" Then
                dic(Trim$(CStr(vg(i, 2)))) = CStr(vg(i, 3))
            End If
        Next i
        If dic.Count = 0 Then Exit Function
        ReDim res(1 To dic.Count, 1 To 2)
        ' Ordre de la liste des étudiants, puis noms absents de cette liste.
        For i = 1 To NbLignes(ve)
            k = Trim$(CStr(ve(i, cNom)))
            If dic.Exists(k) Then
                n = n + 1: res(n, 1) = k: res(n, 2) = dic(k)
                dic.Remove k
            End If
        Next i
        For Each k In dic.Keys
            n = n + 1: res(n, 1) = k: res(n, 2) = dic(k)
        Next k
    Else
        If NbLignes(ve) = 0 Then Exit Function
        ReDim res(1 To NbLignes(ve), 1 To 2)
        For i = 1 To NbLignes(ve)
            If Trim$(CStr(ve(i, cNom))) <> "" Then
                n = n + 1
                res(n, 1) = Trim$(CStr(ve(i, cNom)))
                res(n, 2) = CStr(ve(i, cGrp))
            End If
        Next i
        If n = 0 Then Exit Function
    End If
    ListeEtudiantsMatiere = Tronquer(res, n)
End Function

Private Function Tronquer(res() As Variant, ByVal n As Long) As Variant
    Dim out() As Variant, i As Long
    ReDim out(1 To n, 1 To 2)
    For i = 1 To n
        out(i, 1) = res(i, 1): out(i, 2) = res(i, 2)
    Next i
    Tronquer = out
End Function

Public Function GroupeEtudiant(ByVal nom As String, ByVal code As String) As String
    Dim v As Variant, i As Long
    v = ListeEtudiantsMatiere(code)
    For i = 1 To NbLignes(v)
        If StrComp(v(i, 1), nom, vbTextCompare) = 0 Then GroupeEtudiant = v(i, 2): Exit Function
    Next i
End Function

' Groupes d'une matière, triés (Groupe 2 avant Groupe 10).
Public Function GroupesMatiere(ByVal code As String) As Collection
    Dim v As Variant, i As Long, j As Long, dic As Object, arr As Variant, tmp As Variant
    Set GroupesMatiere = New Collection
    v = ListeEtudiantsMatiere(code)
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare
    For i = 1 To NbLignes(v)
        If Trim$(v(i, 2)) <> "" Then dic(Trim$(v(i, 2))) = 1
    Next i
    If dic.Count = 0 Then Exit Function
    arr = dic.Keys
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CleTri(arr(j)) < CleTri(arr(i)) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
    For i = LBound(arr) To UBound(arr)
        GroupesMatiere.Add arr(i)
    Next i
End Function

Private Function CleTri(ByVal s As String) As String
    Dim i As Long, p As Long
    p = 0
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) Like "#" Then p = i Else Exit For
    Next i
    If p > 0 Then
        CleTri = LCase$(Left$(s, p - 1)) & Format$(Val(Mid$(s, p)), "000000")
    Else
        CleTri = LCase$(s)
    End If
End Function

' "|nom1|nom2|" : membres d'un groupe pour une matière.
Public Function MembresGroupe(ByVal code As String, ByVal groupe As String) As String
    Dim v As Variant, i As Long, s As String
    v = ListeEtudiantsMatiere(code)
    For i = 1 To NbLignes(v)
        If StrComp(Trim$(v(i, 2)), Trim$(groupe), vbTextCompare) = 0 Then s = s & v(i, 1) & "|"
    Next i
    If s <> "" Then MembresGroupe = "|" & s
End Function

' Recalcule la colonne « Membres » des tableaux EXPOSÉS et REMISES.
Public Sub MajTousMembres()
    Dim lo As ListObject, i As Long
    Set lo = Tbl("tblExposes")
    Deverrouiller lo.Parent
    For i = 1 To lo.ListRows.Count
        MajMembresLigne lo, lo.ListRows(i).Range
    Next i
    Verrouiller lo.Parent
    Set lo = Tbl("tblRemises")
    Deverrouiller lo.Parent
    For i = 1 To lo.ListRows.Count
        MajMembresLigne lo, lo.ListRows(i).Range
    Next i
    Verrouiller lo.Parent
End Sub

Public Sub MajMembresLigne(lo As ListObject, r As Range)
    Dim cM As Long, valeur As String, code As String, typ As String, qui As String
    cM = Col(lo, "Membres")
    Select Case lo.Name
        Case "tblExposes"
            code = Trim$(CStr(r.Cells(1, Col(lo, "Code matière")).Value))
            qui = Trim$(CStr(r.Cells(1, Col(lo, "Groupe")).Value))
            If code <> "" And qui <> "" Then valeur = MembresGroupe(code, qui)
        Case "tblRemises"
            Dim rr As Range, loR As ListObject
            Set loR = Tbl("tblRapports")
            Set rr = LigneParID(loR, Trim$(CStr(r.Cells(1, 1).Value)))
            qui = Trim$(CStr(r.Cells(1, Col(lo, "Remis par")).Value))
            If Not rr Is Nothing And qui <> "" Then
                code = CStr(rr.Cells(1, Col(loR, "Code matière")).Value)
                typ = CStr(rr.Cells(1, Col(loR, "Type")).Value)
                If typ = "Groupe" Then valeur = MembresGroupe(code, qui) Else valeur = "|" & qui & "|"
            End If
    End Select
    If CStr(r.Cells(1, cM).Value) <> valeur Then r.Cells(1, cM).Value = valeur
End Sub

' Copie les groupes par défaut dans GROUPES pour la matière choisie.
Public Sub PreparerGroupesMatiere()
    Dim code As String, ve As Variant, loE As ListObject, i As Long, n As Long, data() As Variant
    code = Trim$(CStr(shGroupes.Range("C5").Value))
    If code = "" Or Not MatiereExiste(code) Then
        MsgBox "Choisissez d'abord une matière dans la case « Matière ».", vbExclamation
        Exit Sub
    End If
    If GroupesSpecifiques(code) Then
        MsgBox "La matière " & code & " a déjà des groupes spécifiques." & vbLf & _
               "Modifiez directement la colonne « Groupe » du tableau.", vbInformation
        Exit Sub
    End If
    Set loE = Tbl("tblEtudiants")
    ve = Donnees(loE)
    If NbLignes(ve) = 0 Then Exit Sub
    ReDim data(1 To NbLignes(ve), 1 To 4)
    For i = 1 To NbLignes(ve)
        If Trim$(CStr(ve(i, Col(loE, "Nom et prénoms")))) <> "" Then
            n = n + 1
            data(n, 1) = code
            data(n, 2) = ve(i, Col(loE, "Nom et prénoms"))
            data(n, 3) = ve(i, Col(loE, "Groupe"))
            data(n, 4) = ve(i, Col(loE, "Rôle"))
        End If
    Next i
    DebutTraitement
    Deverrouiller shGroupes
    AjouterBloc Tbl("tblGroupes"), data
    Verrouiller shGroupes
    MajTousMembres
    FinTraitement
    MsgBox n & " étudiants copiés pour " & code & "." & vbLf & _
           "Modifiez maintenant la colonne « Groupe » (ou supprimez les étudiants qui ne suivent pas cette matière).", vbInformation
End Sub

Public Sub SupprimerGroupesMatiere()
    Dim code As String, lo As ListObject, i As Long, n As Long
    code = Trim$(CStr(shGroupes.Range("C5").Value))
    If Not GroupesSpecifiques(code) Then
        MsgBox "Aucun groupe spécifique pour cette matière.", vbInformation
        Exit Sub
    End If
    If MsgBox("Supprimer les groupes spécifiques de " & code & " ?" & vbLf & _
              "La matière utilisera de nouveau les groupes par défaut.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Set lo = Tbl("tblGroupes")
    DebutTraitement
    Deverrouiller shGroupes
    For i = lo.ListRows.Count To 1 Step -1
        If StrComp(CStr(lo.ListRows(i).Range.Cells(1, 1).Value), code, vbTextCompare) = 0 Then
            lo.ListRows(i).Delete
            n = n + 1
        End If
    Next i
    Verrouiller shGroupes
    MajTousMembres
    FinTraitement
    MsgBox n & " ligne(s) supprimée(s).", vbInformation
End Sub
