Option Explicit
' ==========================================================================
'  modExport : export PDF des feuilles de présence et du cahier de texte
'  Les documents sont construits sur la feuille masquée IMPRESSION.
' ==========================================================================

Private Type Seance
    id As String
    jour As Date
    debut As Double
    fin As Double
    prof As String
    typ As String
    contenu As String
    travaux As String
    duree As Double
    taux As Variant
End Type

' ------------------------------------------------------------------ public
Public Sub ExporterPresencePDF()
    Dim code As String, grp As String, mode As String, du As Variant, au As Variant
    Dim codes As Collection, groupes As Collection, c As Variant, g As Variant
    Dim nb As Long, vides As Long, dossier As String, f As String
    LireCriteres code, grp, mode, du, au
    Set codes = CodesChoisis(code)
    If codes.Count = 0 Then Exit Sub
    On Error GoTo erreur
    DebutTraitement
    For Each c In codes
        Set groupes = New Collection
        If grp <> "" Then
            groupes.Add grp
        ElseIf Left$(mode, 8) = "Un seul " Then
            groupes.Add ""
        Else
            For Each g In GroupesMatiere(CStr(c))
                groupes.Add CStr(g)
            Next g
        End If
        For Each g In groupes
            If ConstruirePresence(CStr(c), CStr(g), du, au) Then
                dossier = DossierExport(CStr(c))
                f = dossier & "\Presence_" & NomFichier(CStr(c)) & "_" & _
                    IIf(g = "", "Tous", NomFichier(Replace(CStr(g), " ", ""))) & "_" & Format$(Date, "yyyy-mm-dd") & ".pdf"
                Publier f
                nb = nb + 1
            Else
                vides = vides + 1
            End If
        Next g
    Next c
    FinTraitement
    Bilan nb, vides, IIf(codes.Count = 1, DossierExport(CStr(codes(1))), DossierExport(""))
    Exit Sub
erreur:
    FinTraitement
    shImpression.Visible = xlSheetHidden
    MsgBox "Erreur pendant l'export : " & Err.Description, vbCritical
End Sub

Public Sub ExporterCahierPDF()
    Dim code As String, grp As String, mode As String, du As Variant, au As Variant
    Dim codes As Collection, c As Variant, nb As Long, vides As Long, f As String
    LireCriteres code, grp, mode, du, au
    Set codes = CodesChoisis(code)
    If codes.Count = 0 Then Exit Sub
    On Error GoTo erreur
    DebutTraitement
    For Each c In codes
        If ConstruireCahier(CStr(c), du, au) Then
            f = DossierExport(CStr(c)) & "\CahierDeTexte_" & NomFichier(CStr(c)) & "_" & Format$(Date, "yyyy-mm-dd") & ".pdf"
            Publier f
            nb = nb + 1
        Else
            vides = vides + 1
        End If
    Next c
    FinTraitement
    Bilan nb, vides, IIf(codes.Count = 1, DossierExport(CStr(codes(1))), DossierExport(""))
    Exit Sub
erreur:
    FinTraitement
    shImpression.Visible = xlSheetHidden
    MsgBox "Erreur pendant l'export : " & Err.Description, vbCritical
End Sub

Public Sub OuvrirDossierPDF()
    OuvrirDossier DossierExport("")
End Sub

' ------------------------------------------------------------------ outils
Private Sub LireCriteres(code As String, grp As String, mode As String, du As Variant, au As Variant)
    code = Trim$(CStr(shExport.Range("C6").Value))
    grp = Trim$(CStr(shExport.Range("C7").Value))
    mode = Trim$(CStr(shExport.Range("C8").Value))
    du = shExport.Range("C9").Value
    au = shExport.Range("C10").Value
    If Not IsDate(du) Then du = Empty
    If Not IsDate(au) Then au = Empty
End Sub

Private Function CodesChoisis(ByVal code As String) As Collection
    Set CodesChoisis = New Collection
    If code = "" Then
        Set CodesChoisis = ListeCodes()
    ElseIf MatiereExiste(code) Then
        CodesChoisis.Add code
    Else
        MsgBox "Matière inconnue : " & code, vbExclamation
    End If
End Function

Private Sub Bilan(ByVal nb As Long, ByVal vides As Long, ByVal dossier As String)
    Dim msg As String
    msg = nb & " fichier(s) PDF créé(s) dans :" & vbLf & dossier
    If vides > 0 Then msg = msg & vbLf & vbLf & vides & " document(s) ignoré(s) : aucune séance enregistrée sur la période."
    If nb > 0 Then
        If MsgBox(msg & vbLf & vbLf & "Ouvrir le dossier ?", vbYesNo + vbInformation) = vbYes Then OuvrirDossier dossier
    Else
        MsgBox msg, vbInformation
    End If
End Sub

Private Sub Publier(ByVal f As String)
    shImpression.Visible = xlSheetVisible
    shImpression.ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    shImpression.Visible = xlSheetHidden
End Sub

Private Function DansPeriode(ByVal d As Date, du As Variant, au As Variant) As Boolean
    DansPeriode = True
    If Not IsEmpty(du) Then If d < CDate(du) Then DansPeriode = False
    If Not IsEmpty(au) Then If d > CDate(au) Then DansPeriode = False
End Function

' Séances d'une matière sur la période, triées par date et heure.
Private Function SeancesMatiere(ByVal code As String, du As Variant, au As Variant, s() As Seance) As Long
    Dim lo As ListObject, v As Variant, i As Long, j As Long, n As Long, tmp As Seance
    Set lo = Tbl("tblSeances")
    v = Donnees(lo)
    If NbLignes(v) = 0 Then Exit Function
    ReDim s(1 To NbLignes(v))
    For i = 1 To NbLignes(v)
        If StrComp(CStr(v(i, Col(lo, "Code matière"))), code, vbTextCompare) = 0 And IsDate(v(i, Col(lo, "Date"))) Then
            If DansPeriode(CDate(v(i, Col(lo, "Date"))), du, au) Then
                n = n + 1
                s(n).id = CStr(v(i, Col(lo, "ID")))
                s(n).jour = CDate(v(i, Col(lo, "Date")))
                s(n).debut = Nombre(v(i, Col(lo, "Début")))
                s(n).fin = Nombre(v(i, Col(lo, "Fin")))
                s(n).prof = CStr(v(i, Col(lo, "Présence prof")))
                s(n).typ = CStr(v(i, Col(lo, "Type de séance")))
                s(n).contenu = CStr(v(i, Col(lo, "Contenu / chapitres traités")))
                s(n).travaux = CStr(v(i, Col(lo, "Travaux demandés")))
                s(n).duree = Nombre(v(i, Col(lo, "Durée (h)")))
                s(n).taux = v(i, Col(lo, "Taux de présence"))
            End If
        End If
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If s(j).jour + s(j).debut < s(i).jour + s(i).debut Then tmp = s(i): s(i) = s(j): s(j) = tmp
        Next j
    Next i
    SeancesMatiere = n
End Function

Private Sub PreparerFeuille()
    With shImpression
        .Cells.Clear
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10
        .Cells.RowHeight = 15
        .Cells.ColumnWidth = 8
        .ResetAllPageBreaks
    End With
End Sub

Private Sub EnTete(ByVal titre As String, ByVal derCol As Long, ByVal code As String, ByVal sousTitre As String)
    With shImpression
        .Cells(1, 1).Value = UCase$(CStr(Param("pEcole")))
        .Cells(2, 1).Value = Param("pDepartement") & "  -  " & Param("pNiveau") & " " & Param("pOption") & _
                             "  -  Année académique " & Param("pAnnee") & "  -  " & Param("pSemestre")
        .Cells(4, 1).Value = titre
        .Cells(5, 1).Value = code & "  -  " & InfoMatiere(code, "Intitulé")
        .Cells(6, 1).Value = "Enseignant(s) : " & InfoMatiere(code, "Enseignant(s)") & "     " & sousTitre
        .Range(.Cells(1, 1), .Cells(1, derCol)).Merge
        .Range(.Cells(2, 1), .Cells(2, derCol)).Merge
        .Range(.Cells(4, 1), .Cells(4, derCol)).Merge
        .Range(.Cells(5, 1), .Cells(5, derCol)).Merge
        .Range(.Cells(6, 1), .Cells(6, derCol)).Merge
        .Range(.Cells(1, 1), .Cells(6, 1)).HorizontalAlignment = xlCenter
        .Cells(1, 1).Font.Bold = True: .Cells(1, 1).Font.Size = 13: .Cells(1, 1).Font.Color = CouleurMarine()
        .Cells(2, 1).Font.Italic = True: .Cells(2, 1).Font.Color = RGB(89, 89, 89)
        With .Range(.Cells(4, 1), .Cells(4, derCol))
            .Font.Bold = True: .Font.Size = 15: .Font.Color = RGB(255, 255, 255)
            .Interior.Color = CouleurMarine()
        End With
        .Rows(4).RowHeight = 24
        .Cells(5, 1).Font.Bold = True: .Cells(5, 1).Font.Size = 12: .Cells(5, 1).Font.Color = CouleurBleu()
        .Rows(5).RowHeight = 19
    End With
End Sub

Private Sub MiseEnPage(ByVal derLigne As Long, ByVal derCol As Long, ByVal paysage As Boolean, ByVal lignesTitre As String)
    With shImpression.PageSetup
        .PrintArea = shImpression.Range(shImpression.Cells(1, 1), shImpression.Cells(derLigne, derCol)).Address
        .PrintTitleRows = lignesTitre
        .Orientation = IIf(paysage, xlLandscape, xlPortrait)
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .CenterFooter = "Page &P / &N"
        .LeftFooter = "Édité le " & Format$(Now, "dd/mm/yyyy à hh:nn")
        .RightFooter = Param("pNiveau") & " " & Param("pAnnee")
    End With
End Sub

Private Sub Quadriller(r As Range)
    With r.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = CouleurBordure()
    End With
End Sub

' ------------------------------------------------------- feuille de présence
Private Function ConstruirePresence(ByVal code As String, ByVal grp As String, du As Variant, au As Variant) As Boolean
    Dim s() As Seance, ns As Long, etu As Variant, ne As Long, i As Long, j As Long, n As Long
    Dim dic As Object, v As Variant, lo As ListObject, cle As String, st As String
    Dim noms() As String, groupes() As String, L0 As Long, c0 As Long, derCol As Long, derLigne As Long
    Dim nbP As Long, nbA As Long, nbR As Long, nbE As Long, avecGrp As Boolean
    Dim grille() As Variant, periode As String

    ns = SeancesMatiere(code, du, au, s)
    If ns = 0 Then Exit Function
    etu = ListeEtudiantsMatiere(code)
    For i = 1 To NbLignes(etu)
        If grp = "" Or StrComp(Trim$(etu(i, 2)), grp, vbTextCompare) = 0 Then
            n = n + 1
            ReDim Preserve noms(1 To n): ReDim Preserve groupes(1 To n)
            noms(n) = etu(i, 1): groupes(n) = etu(i, 2)
        End If
    Next i
    If n = 0 Then Exit Function
    ne = n

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare
    Set lo = Tbl("tblPresences")
    v = Donnees(lo)
    For i = 1 To NbLignes(v)
        dic(CStr(v(i, 1)) & "|" & CStr(v(i, Col(lo, "Étudiant")))) = CStr(v(i, Col(lo, "Statut")))
    Next i

    avecGrp = (grp = "")
    c0 = IIf(avecGrp, 4, 3)          ' première colonne de séance
    derCol = c0 + ns + 4             ' + P, A, R, E... puis taux
    L0 = 9                           ' ligne d'en-tête du tableau

    PreparerFeuille
    If IsEmpty(du) And IsEmpty(au) Then
        periode = "Toutes les séances enregistrées"
    Else
        periode = "Période : " & IIf(IsEmpty(du), "début", Format$(du, "dd/mm/yyyy")) & " au " & _
                  IIf(IsEmpty(au), "ce jour", Format$(au, "dd/mm/yyyy"))
    End If
    EnTete "FEUILLE DE PRÉSENCE", derCol, code, "Groupe : " & IIf(grp = "", "tous les groupes", grp) & "     " & periode

    With shImpression
        .Cells(7, 1).Value = "Légende : P = présent   A = absent   R = retard   E = excusé   (en-tête grisé : enseignant absent)"
        .Range(.Cells(7, 1), .Cells(7, derCol)).Merge
        .Cells(7, 1).HorizontalAlignment = xlCenter
        .Cells(7, 1).Font.Size = 8: .Cells(7, 1).Font.Italic = True

        .Cells(L0, 1).Value = "N°"
        .Cells(L0, 2).Value = "Nom et prénoms"
        If avecGrp Then .Cells(L0, 3).Value = "Groupe"
        For j = 1 To ns
            .Cells(L0, c0 + j - 1).Value = Format$(s(j).jour, "dd/mm") & vbLf & Format$(s(j).debut, "hh\hnn")
        Next j
        .Cells(L0, c0 + ns).Value = "P"
        .Cells(L0, c0 + ns + 1).Value = "A"
        .Cells(L0, c0 + ns + 2).Value = "R"
        .Cells(L0, c0 + ns + 3).Value = "E"
        .Cells(L0, c0 + ns + 4).Value = "Taux"

        ReDim grille(1 To ne, 1 To derCol)
        For i = 1 To ne
            nbP = 0: nbA = 0: nbR = 0: nbE = 0
            grille(i, 1) = i
            grille(i, 2) = noms(i)
            If avecGrp Then grille(i, 3) = groupes(i)
            For j = 1 To ns
                cle = s(j).id & "|" & noms(i)
                st = ""
                If dic.Exists(cle) Then st = dic(cle)
                grille(i, c0 + j - 1) = st
                Select Case st
                    Case "P": nbP = nbP + 1
                    Case "A": nbA = nbA + 1
                    Case "R": nbR = nbR + 1
                    Case "E": nbE = nbE + 1
                End Select
            Next j
            grille(i, c0 + ns) = nbP
            grille(i, c0 + ns + 1) = nbA
            grille(i, c0 + ns + 2) = nbR
            grille(i, c0 + ns + 3) = nbE
            If nbP + nbA + nbR + nbE > 0 Then grille(i, c0 + ns + 4) = (nbP + nbR) / (nbP + nbA + nbR + nbE) Else grille(i, c0 + ns + 4) = ""
        Next i
        .Cells(L0 + 1, 1).Resize(ne, derCol).Value = grille
        derLigne = L0 + ne + 1

        ' Ligne de synthèse : présents par séance
        .Cells(derLigne, 2).Value = "Présents (P + R)"
        For j = 1 To ns
            .Cells(derLigne, c0 + j - 1).Formula = "=COUNTIF(" & .Range(.Cells(L0 + 1, c0 + j - 1), .Cells(L0 + ne, c0 + j - 1)).Address(False, False) & _
                ",""P"")+COUNTIF(" & .Range(.Cells(L0 + 1, c0 + j - 1), .Cells(L0 + ne, c0 + j - 1)).Address(False, False) & ",""R"")"
        Next j
        .Rows(derLigne).Font.Bold = True
        .Range(.Cells(derLigne, 1), .Cells(derLigne, derCol)).Interior.Color = CouleurBleuClair()

        ' Mise en forme
        With .Range(.Cells(L0, 1), .Cells(L0, derCol))
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            .Interior.Color = CouleurBleu()
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        .Rows(L0).RowHeight = 30
        For j = 1 To ns
            If s(j).prof = "Absent" Then .Cells(L0, c0 + j - 1).Interior.Color = RGB(128, 128, 128)
        Next j
        Quadriller .Range(.Cells(L0, 1), .Cells(derLigne, derCol))
        .Range(.Cells(L0 + 1, 1), .Cells(derLigne, 1)).HorizontalAlignment = xlCenter
        .Range(.Cells(L0 + 1, 3), .Cells(derLigne, derCol)).HorizontalAlignment = xlCenter
        .Range(.Cells(L0 + 1, c0 + ns + 4), .Cells(L0 + ne, c0 + ns + 4)).NumberFormat = "0%"
        For i = 1 To ne
            If i Mod 2 = 0 Then .Range(.Cells(L0 + i, 1), .Cells(L0 + i, c0 - 1)).Interior.Color = RGB(243, 247, 252)
            For j = 1 To ns
                st = CStr(grille(i, c0 + j - 1))
                If st <> "" Then
                    .Cells(L0 + i, c0 + j - 1).Interior.Color = CouleurStatut(st)
                    .Cells(L0 + i, c0 + j - 1).Font.Color = CouleurTexteStatut(st)
                    .Cells(L0 + i, c0 + j - 1).Font.Bold = True
                End If
            Next j
            If grille(i, c0 + ns + 1) > Nombre(Param("pSeuil")) Then
                .Cells(L0 + i, c0 + ns + 1).Interior.Color = RGB(192, 0, 0)
                .Cells(L0 + i, c0 + ns + 1).Font.Color = RGB(255, 255, 255)
                .Cells(L0 + i, 2).Font.Color = RGB(192, 0, 0)
            End If
        Next i
        .Columns(1).ColumnWidth = 4
        .Columns(2).ColumnWidth = 32
        If avecGrp Then .Columns(3).ColumnWidth = 10
        .Range(.Columns(c0), .Columns(c0 + ns - 1)).ColumnWidth = 6
        .Range(.Columns(c0 + ns), .Columns(c0 + ns + 3)).ColumnWidth = 4
        .Columns(c0 + ns + 4).ColumnWidth = 6

        ' Signatures
        .Cells(derLigne + 2, 2).Value = "Le délégué de classe"
        .Cells(derLigne + 2, c0 + ns).Value = "L'enseignant"
        .Cells(derLigne + 3, 2).Value = CStr(Param("pDelegue"))
        .Range(.Cells(derLigne + 2, 1), .Cells(derLigne + 3, derCol)).Font.Bold = True
        derLigne = derLigne + 6
    End With
    MiseEnPage derLigne, derCol, (ns > 8), "$" & L0 & ":$" & L0
    ConstruirePresence = True
End Function

' ------------------------------------------------------------ cahier de texte
Private Function ConstruireCahier(ByVal code As String, du As Variant, au As Variant) As Boolean
    Dim s() As Seance, ns As Long, j As Long, L As Long, L0 As Long, total As Double, prevu As Double
    Const NC As Long = 8
    ns = SeancesMatiere(code, du, au, s)
    If ns = 0 Then Exit Function
    PreparerFeuille
    prevu = Nombre(InfoMatiere(code, "Volume total (h)"))
    For j = 1 To ns
        If s(j).prof <> "Absent" Then total = total + s(j).duree
    Next j
    EnTete "CAHIER DE TEXTE", NC, code, "Heures effectuées : " & Format$(total, "0.0") & " h / " & _
        Format$(prevu, "0") & " h prévues (" & IIf(prevu > 0, Format$(total / prevu, "0%"), "-") & ")"
    L0 = 8
    With shImpression
        .Cells(L0, 1).Resize(1, NC).Value = Array("N°", "Date", "Horaire", "Durée", "Enseignant", "Type", _
            "Contenu / chapitres traités", "Travaux demandés")
        For j = 1 To ns
            L = L0 + j
            .Cells(L, 1).Value = j
            .Cells(L, 2).Value = s(j).jour
            .Cells(L, 2).NumberFormat = "[$-40C]ddd dd/mm/yy"
            .Cells(L, 3).Value = Format$(s(j).debut, "hh\hnn") & IIf(s(j).fin > 0, " - " & Format$(s(j).fin, "hh\hnn"), "")
            If s(j).duree > 0 Then .Cells(L, 4).Value = Format$(s(j).duree, "0.0") & " h"
            .Cells(L, 5).Value = s(j).prof
            .Cells(L, 6).Value = s(j).typ
            .Cells(L, 7).Value = s(j).contenu
            .Cells(L, 8).Value = s(j).travaux
            If s(j).prof = "Absent" Then
                .Range(.Cells(L, 1), .Cells(L, NC)).Interior.Color = CouleurStatut("A")
            ElseIf j Mod 2 = 0 Then
                .Range(.Cells(L, 1), .Cells(L, NC)).Interior.Color = RGB(243, 247, 252)
            End If
        Next j
        With .Range(.Cells(L0, 1), .Cells(L0, NC))
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            .Interior.Color = CouleurBleu()
            .HorizontalAlignment = xlCenter
        End With
        .Range(.Cells(L0 + 1, 1), .Cells(L0 + ns, NC)).VerticalAlignment = xlTop
        .Range(.Cells(L0 + 1, 7), .Cells(L0 + ns, NC)).WrapText = True
        .Range(.Cells(L0 + 1, 1), .Cells(L0 + ns, 6)).HorizontalAlignment = xlCenter
        Quadriller .Range(.Cells(L0, 1), .Cells(L0 + ns, NC))
        .Columns(1).ColumnWidth = 4
        .Columns(2).ColumnWidth = 12
        .Columns(3).ColumnWidth = 12
        .Columns(4).ColumnWidth = 7
        .Columns(5).ColumnWidth = 10
        .Columns(6).ColumnWidth = 9
        .Columns(7).ColumnWidth = 55
        .Columns(8).ColumnWidth = 30
        .Range(.Rows(L0 + 1), .Rows(L0 + ns)).AutoFit
        .Cells(L0 + ns + 2, 2).Value = "Visa de l'enseignant"
        .Cells(L0 + ns + 2, 7).Value = "Le délégué : " & CStr(Param("pDelegue"))
        .Range(.Cells(L0 + ns + 2, 1), .Cells(L0 + ns + 2, NC)).Font.Bold = True
    End With
    MiseEnPage L0 + ns + 5, NC, True, "$" & L0 & ":$" & L0
    ConstruireCahier = True
End Function
