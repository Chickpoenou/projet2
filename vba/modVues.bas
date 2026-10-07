Option Explicit
' ==========================================================================
'  modVues : listes affichées sur les feuilles de consultation.
'  Remplies par macro (et non par FILTER/UNIQUE) pour fonctionner avec
'  toutes les versions d'Excel (2016, 2019, 2021, 365).
' ==========================================================================

Private Const MAXL As Long = 300   ' lignes maximum par liste

' ---------------------------------------------------------------- outils
Private Function Couper(data() As Variant, ByVal n As Long, ByVal nc As Long) As Variant
    Dim out() As Variant, i As Long, j As Long
    ReDim out(1 To n, 1 To nc)
    For i = 1 To n
        For j = 1 To nc
            out(i, j) = data(i, j)
        Next j
    Next i
    Couper = out
End Function

' Tri décroissant (ou croissant) sur une clé numérique précalculée.
Private Sub Trier(data() As Variant, cles() As Double, ByVal n As Long, ByVal nc As Long, ByVal decroissant As Boolean)
    Dim i As Long, j As Long, k As Long, tmp As Variant, tk As Double, permuter As Boolean
    For i = 2 To n
        j = i
        Do While j > 1
            If decroissant Then permuter = cles(j) > cles(j - 1) Else permuter = cles(j) < cles(j - 1)
            If Not permuter Then Exit Do
            tk = cles(j): cles(j) = cles(j - 1): cles(j - 1) = tk
            For k = 1 To nc
                tmp = data(j, k): data(j, k) = data(j - 1, k): data(j - 1, k) = tmp
            Next k
            j = j - 1
        Loop
    Next i
End Sub

Private Sub Ecrire(ws As Worksheet, ByVal ligne As Long, ByVal cl As Long, ByVal nc As Long, _
                   data() As Variant, ByVal n As Long, ByVal messageVide As String)
    ws.Range(ws.Cells(ligne, cl), ws.Cells(ligne + MAXL - 1, cl + nc - 1)).ClearContents
    If n = 0 Then
        ws.Cells(ligne, cl).Value = messageVide
    Else
        ws.Cells(ligne, cl).Resize(n, nc).Value = Couper(data, n, nc)
    End If
End Sub

' Copie les colonnes c1..c2 de la ligne i de v dans la ligne n de data.
Private Sub CopierLigne(v As Variant, ByVal i As Long, ByVal c1 As Long, ByVal c2 As Long, data() As Variant, ByVal n As Long)
    Dim j As Long
    For j = c1 To c2
        data(n, j - c1 + 1) = v(i, j)
    Next j
End Sub

Public Sub ActualiserVues()
    On Error Resume Next
    RemplirFiche
    RemplirSuivi
    RemplirGroupes
    RemplirSyntheseExposes
End Sub

' --------------------------------------------------------- FICHE ÉTUDIANT
Public Sub RemplirFiche()
    Dim nom As String, lo As ListObject, v As Variant, i As Long, n As Long
    Dim data() As Variant, cles() As Double, c1 As Long, c2 As Long
    Dim cEtu As Long, cStat As Long, cDat As Long, cHeure As Long, cM As Long
    nom = Trim$(CStr(shFiche.Range("C7").Value))
    Application.Calculate
    Application.ScreenUpdating = False
    Deverrouiller shFiche

    ' Présences : absences (B35) et historique (M7)
    Set lo = Tbl("tblPresences")
    v = Donnees(lo)
    cEtu = Col(lo, "Étudiant"): cStat = Col(lo, "Statut")
    cDat = Col(lo, "Date"): cHeure = Col(lo, "Heure")
    c1 = Col(lo, "Code matière")
    ReDim data(1 To MAXL, 1 To 5): ReDim cles(1 To MAXL)
    n = 0
    If nom <> "" Then
        For i = 1 To NbLignes(v)
            If CStr(v(i, cEtu)) = nom And UCase$(CStr(v(i, cStat))) = "A" And n < MAXL Then
                n = n + 1
                CopierLigne v, i, c1, cHeure, data, n
                cles(n) = Nombre(v(i, cDat)) + Nombre(v(i, cHeure))
            End If
        Next i
        Trier data, cles, n, 4, True
    End If
    Ecrire shFiche, 35, 2, 4, data, n, IIf(nom = "", "", "Aucune absence")

    ReDim data(1 To MAXL, 1 To 5): ReDim cles(1 To MAXL)
    n = 0
    If nom <> "" Then
        For i = 1 To NbLignes(v)
            If CStr(v(i, cEtu)) = nom And n < MAXL Then
                n = n + 1
                CopierLigne v, i, c1, cStat, data, n
                cles(n) = Nombre(v(i, cDat)) + Nombre(v(i, cHeure))
            End If
        Next i
        Trier data, cles, n, 5, True
    End If
    Ecrire shFiche, 7, 13, 5, data, n, IIf(nom = "", "", "Aucune séance enregistrée")

    ' Rapports (S7) : individuels ou de ses groupes
    Set lo = Tbl("tblRemises")
    v = Donnees(lo)
    c1 = Col(lo, "Code matière"): c2 = Col(lo, "Statut"): cM = Col(lo, "Membres")
    ReDim data(1 To MAXL, 1 To c2 - c1 + 1)
    n = 0
    If nom <> "" Then
        For i = 1 To NbLignes(v)
            If InStr(1, CStr(v(i, cM)), "|" & nom & "|", vbTextCompare) > 0 And n < MAXL Then
                n = n + 1
                CopierLigne v, i, c1, c2, data, n
            End If
        Next i
    End If
    Ecrire shFiche, 7, 19, c2 - c1 + 1, data, n, IIf(nom = "", "", "Aucun rapport")

    ' Exposés (AB7)
    Set lo = Tbl("tblExposes")
    v = Donnees(lo)
    c1 = Col(lo, "Code matière"): c2 = Col(lo, "Note sur 20"): cM = Col(lo, "Membres")
    ReDim data(1 To MAXL, 1 To c2 - c1 + 1)
    n = 0
    If nom <> "" Then
        For i = 1 To NbLignes(v)
            If InStr(1, CStr(v(i, cM)), "|" & nom & "|", vbTextCompare) > 0 And n < MAXL Then
                n = n + 1
                CopierLigne v, i, c1, c2, data, n
            End If
        Next i
    End If
    Ecrire shFiche, 7, 28, c2 - c1 + 1, data, n, IIf(nom = "", "", "Aucun exposé")

    Verrouiller shFiche
    Application.ScreenUpdating = True
End Sub

' ------------------------------------------------------ SUIVI ENSEIGNANTS
Public Sub RemplirSuivi()
    Dim lo As ListObject, v As Variant, i As Long, n As Long, ens As String
    Dim data() As Variant, cles() As Double, c1 As Long, c2 As Long, cP As Long, cE As Long, cR As Long
    Dim cDat As Long, cDeb As Long, p As String
    Application.Calculate
    Application.ScreenUpdating = False
    Deverrouiller shSuivi
    Set lo = Tbl("tblSeances")
    v = Donnees(lo)
    cP = Col(lo, "Présence prof"): cE = Col(lo, "Enseignant prévu"): cR = Col(lo, "Remplaçant ou motif")
    cDat = Col(lo, "Date"): cDeb = Col(lo, "Début")

    ' Absences, retards, remplacements (B26)
    c1 = Col(lo, "Date"): c2 = cR
    ReDim data(1 To MAXL, 1 To c2 - c1 + 1): ReDim cles(1 To MAXL)
    n = 0
    For i = 1 To NbLignes(v)
        p = CStr(v(i, cP))
        If p <> "" And p <> "Présent" And n < MAXL Then
            n = n + 1
            CopierLigne v, i, c1, c2, data, n
            cles(n) = Nombre(v(i, cDat)) + Nombre(v(i, cDeb))
        End If
    Next i
    Trier data, cles, n, c2 - c1 + 1, True
    Ecrire shSuivi, 26, 2, c2 - c1 + 1, data, n, "Aucune absence ni retard d'enseignant enregistré"

    ' Séances de l'enseignant choisi (P12)
    ens = Trim$(CStr(shSuivi.Range("Q6").Value))
    c2 = Col(lo, "Contenu / chapitres traités")
    ReDim data(1 To MAXL, 1 To c2 - c1 + 1): ReDim cles(1 To MAXL)
    n = 0
    If ens <> "" Then
        For i = 1 To NbLignes(v)
            If (InStr(1, CStr(v(i, cE)), ens, vbTextCompare) > 0 Or StrComp(CStr(v(i, cR)), ens, vbTextCompare) = 0) _
               And n < MAXL Then
                n = n + 1
                CopierLigne v, i, c1, c2, data, n
                cles(n) = Nombre(v(i, cDat)) + Nombre(v(i, cDeb))
            End If
        Next i
        Trier data, cles, n, c2 - c1 + 1, True
    End If
    Ecrire shSuivi, 12, 16, c2 - c1 + 1, data, n, IIf(ens = "", "", "Aucune séance enregistrée pour cet enseignant")
    Verrouiller shSuivi
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------- GROUPES
Public Sub RemplirGroupes()
    Dim code As String, liste As Variant, grp As Collection, g As Variant, i As Long, k As Long, n As Long
    Dim cc As Long, lo As ListObject, membres() As Variant
    Const NBCOL As Long = 16, L0 As Long = 9, C0 As Long = 7
    code = CodeDe(CStr(shGroupes.Range("C5").Value))
    Application.ScreenUpdating = False
    Deverrouiller shGroupes
    shGroupes.Range(shGroupes.Cells(L0, C0), shGroupes.Cells(L0 + 51, C0 + NBCOL - 1)).ClearContents

    ' Filtre du tableau sur la matière choisie
    Set lo = Tbl("tblGroupes")
    On Error Resume Next
    If code <> "" And GroupesSpecifiques(code) Then
        lo.Range.AutoFilter Field:=1, Criteria1:=code
    Else
        If Not lo.AutoFilter Is Nothing Then If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    On Error GoTo 0

    Set grp = GroupesMatiere(code)
    liste = ListeEtudiantsMatiere(code)
    For Each g In grp
        k = k + 1
        If k > NBCOL Then Exit For
        cc = C0 + k - 1
        n = 0
        ReDim membres(1 To 50, 1 To 1)
        For i = 1 To NbLignes(liste)
            If StrComp(Trim$(CStr(liste(i, 2))), CStr(g), vbTextCompare) = 0 And n < 50 Then
                n = n + 1
                membres(n, 1) = liste(i, 1)
            End If
        Next i
        shGroupes.Cells(L0, cc).Value = g & "  (" & n & ")"
        If n > 0 Then shGroupes.Cells(L0 + 1, cc).Resize(n, 1).Value = Couper(membres, n, 1)
    Next g
    If grp.Count > NBCOL Then shGroupes.Cells(L0, C0 + NBCOL - 1).Value = "… " & grp.Count - NBCOL + 1 & " groupes de plus"
    Verrouiller shGroupes
    Application.ScreenUpdating = True
End Sub

' ------------------------------------------------- SYNTHÈSE DES EXPOSÉS
Public Sub RemplirSyntheseExposes()
    Dim r As Long, code As String, lo As ListObject, v As Variant, i As Long
    Dim passes As Object, groupes As Collection, g As Variant, sP As String, sR As String, nR As Long
    Dim cCode As Long, cGrp As Long, cStat As Long, nbExp As Long, arr As Variant
    Application.Calculate
    Set lo = Tbl("tblExposes")
    v = Donnees(lo)
    cCode = Col(lo, "Code matière"): cGrp = Col(lo, "Groupe"): cStat = Col(lo, "Statut")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Deverrouiller shExposes
    For r = 8 To 22
        code = Trim$(CStr(shExposes.Cells(r, 2).Value))
        shExposes.Cells(r, 4).Value = ""
        shExposes.Cells(r, 6).Value = ""
        shExposes.Cells(r, 7).Value = ""
        shExposes.Cells(r, 9).Value = ""
        If code <> "" Then
            Set passes = CreateObject("Scripting.Dictionary")
            passes.CompareMode = vbTextCompare
            nbExp = 0
            For i = 1 To NbLignes(v)
                If StrComp(CStr(v(i, cCode)), code, vbTextCompare) = 0 Then
                    nbExp = nbExp + 1
                    If CStr(v(i, cStat)) = "Présenté" Then passes(Trim$(CStr(v(i, cGrp)))) = 1
                End If
            Next i
            If nbExp = 0 Then
                shExposes.Cells(r, 4).Value = "—"
                shExposes.Cells(r, 6).Value = "—"
                shExposes.Cells(r, 7).Value = "Aucun exposé programmé"
            Else
                Set groupes = GroupesMatiere(code)
                sP = "": sR = "": nR = 0
                For Each g In groupes
                    If passes.Exists(CStr(g)) Then
                        sP = sP & IIf(sP = "", "", ", ") & g
                    Else
                        sR = sR & IIf(sR = "", "", ", ") & g
                        nR = nR + 1
                    End If
                Next g
                ' Groupes présentés hors composition actuelle
                arr = passes.Keys
                For i = LBound(arr) To UBound(arr)
                    If InStr(1, ", " & sP & ",", ", " & arr(i) & ",", vbTextCompare) = 0 Then sP = sP & IIf(sP = "", "", ", ") & arr(i)
                Next i
                shExposes.Cells(r, 4).Value = groupes.Count
                shExposes.Cells(r, 6).Value = nR
                shExposes.Cells(r, 7).Value = IIf(sP = "", "Aucun pour l'instant", sP)
                shExposes.Cells(r, 9).Value = IIf(sR = "", ChrW(10004) & " Tous les groupes ont présenté", sR)
            End If
        End If
    Next r
    Verrouiller shExposes
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
