Option Explicit
' ==========================================================================
'  modOutils : fonctions communes (tableaux, protection, paramètres, couleurs)
' ==========================================================================

' Mot de passe de protection des feuilles (à modifier ici si besoin).
Public Const MDP As String = "GC5BTP"

' Vrai lorsque le mode administrateur est actif (feuilles déverrouillées).
Public ModeAdmin As Boolean

' ---------- Palette (doit rester cohérente avec la feuille PARAMÈTRES) ----------
Public Function CouleurStatut(ByVal statut As String) As Long
    Select Case UCase$(statut)
        Case "P": CouleurStatut = RGB(198, 239, 206)
        Case "A": CouleurStatut = RGB(255, 199, 206)
        Case "R": CouleurStatut = RGB(255, 235, 156)
        Case "E": CouleurStatut = RGB(217, 217, 217)
        Case Else: CouleurStatut = RGB(255, 255, 255)
    End Select
End Function

Public Function CouleurTexteStatut(ByVal statut As String) As Long
    Select Case UCase$(statut)
        Case "P": CouleurTexteStatut = RGB(0, 97, 0)
        Case "A": CouleurTexteStatut = RGB(156, 0, 6)
        Case "R": CouleurTexteStatut = RGB(156, 87, 0)
        Case "E": CouleurTexteStatut = RGB(64, 64, 64)
        Case Else: CouleurTexteStatut = RGB(0, 0, 0)
    End Select
End Function

Public Function CouleurMarine() As Long: CouleurMarine = RGB(31, 56, 100): End Function
Public Function CouleurBleu() As Long: CouleurBleu = RGB(46, 117, 182): End Function
Public Function CouleurBleuClair() As Long: CouleurBleuClair = RGB(221, 235, 247): End Function
Public Function CouleurBordure() As Long: CouleurBordure = RGB(191, 191, 191): End Function

' ---------- Accès aux tableaux ----------
Public Function Tbl(ByVal nom As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set Tbl = ws.ListObjects(nom)
        On Error GoTo 0
        If Not Tbl Is Nothing Then Exit Function
    Next ws
    Err.Raise vbObjectError + 1, , "Tableau introuvable : " & nom
End Function

Public Function Col(lo As ListObject, ByVal nom As String) As Long
    Col = lo.ListColumns(nom).Index
End Function

' Renvoie les données du tableau (tableau 2D 1..n) ou Empty si vide.
Public Function Donnees(lo As ListObject) As Variant
    Dim v As Variant
    If lo.DataBodyRange Is Nothing Then Exit Function
    If lo.ListRows.Count = 1 Then
        If LigneVide(lo.ListRows(1).Range) Then Exit Function
    End If
    v = lo.DataBodyRange.Value
    Donnees = v
End Function

Public Function NbLignes(v As Variant) As Long
    If IsEmpty(v) Then NbLignes = 0 Else NbLignes = UBound(v, 1)
End Function

Public Function LigneVide(r As Range) As Boolean
    LigneVide = (Trim$(CStr(r.Cells(1, 1).Value)) = "")
End Function

' Renvoie une ligne libre du tableau (réutilise la ligne vierge initiale).
Public Function NouvelleLigne(lo As ListObject) As Range
    If lo.ListRows.Count = 0 Then
        Set NouvelleLigne = lo.ListRows.Add.Range
    ElseIf lo.ListRows.Count = 1 And LigneVide(lo.ListRows(1).Range) Then
        Set NouvelleLigne = lo.ListRows(1).Range
    Else
        Set NouvelleLigne = lo.ListRows.Add.Range
    End If
End Function

' Ajoute un bloc de valeurs (tableau 2D 1..n x nb colonnes) en une seule fois.
' Réservé aux tableaux sans colonne calculée.
Public Sub AjouterBloc(lo As ListObject, data As Variant)
    Dim n As Long, nbExist As Long
    n = UBound(data, 1)
    nbExist = lo.ListRows.Count
    If nbExist = 1 Then
        If LigneVide(lo.ListRows(1).Range) Then nbExist = 0
    End If
    lo.Resize lo.HeaderRowRange.Resize(1 + nbExist + n)
    lo.HeaderRowRange.Offset(1 + nbExist, 0).Resize(n).Value = data
End Sub

' Prochain identifiant : préfixe + numéro (ex. S0001, R001).
Public Function NouvelID(lo As ListObject, ByVal prefixe As String, ByVal chiffres As Integer) As String
    Dim v As Variant, i As Long, n As Long, m As Long
    v = Donnees(lo)
    For i = 1 To NbLignes(v)
        If UCase$(Left$(CStr(v(i, 1)), Len(prefixe))) = UCase$(prefixe) Then
            n = Val(Mid$(CStr(v(i, 1)), Len(prefixe) + 1))
            If n > m Then m = n
        End If
    Next i
    NouvelID = prefixe & Format$(m + 1, String$(chiffres, "0"))
End Function

' Ligne du tableau dont la 1re colonne vaut id (Nothing si absente).
Public Function LigneParID(lo As ListObject, ByVal id As String) As Range
    Dim v As Variant, i As Long
    v = Donnees(lo)
    For i = 1 To NbLignes(v)
        If CStr(v(i, 1)) = id Then
            Set LigneParID = lo.ListRows(i).Range
            Exit Function
        End If
    Next i
End Function

' ---------- Paramètres et référentiels ----------
Public Function Param(ByVal nom As String) As Variant
    On Error Resume Next
    Param = ThisWorkbook.Names(nom).RefersToRange.Value
End Function

Public Function InfoMatiere(ByVal code As String, ByVal colonne As String) As String
    Dim lo As ListObject, v As Variant, i As Long, c As Long
    Set lo = Tbl("tblMatieres")
    v = Donnees(lo)
    c = Col(lo, colonne)
    For i = 1 To NbLignes(v)
        If StrComp(CStr(v(i, 1)), code, vbTextCompare) = 0 Then
            InfoMatiere = CStr(v(i, c))
            Exit Function
        End If
    Next i
End Function

Public Function MatiereExiste(ByVal code As String) As Boolean
    Dim v As Variant, i As Long
    v = Donnees(Tbl("tblMatieres"))
    For i = 1 To NbLignes(v)
        If StrComp(CStr(v(i, 1)), code, vbTextCompare) = 0 Then MatiereExiste = True: Exit Function
    Next i
End Function

Public Function ListeCodes() As Collection
    Dim v As Variant, i As Long
    Set ListeCodes = New Collection
    v = Donnees(Tbl("tblMatieres"))
    For i = 1 To NbLignes(v)
        If Trim$(CStr(v(i, 1))) <> "" Then ListeCodes.Add CStr(v(i, 1))
    Next i
End Function

Public Function Nombre(v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then
        If VarType(v) <> vbString Or Len(v) > 0 Then Nombre = CDbl(v)
    End If
End Function

' ---------- Texte ----------
Public Function SansAccents(ByVal s As String) As String
    Const AVEC As String = "àâäáãåçéèêëíìîïñóòôöõúùûüýÿÀÂÄÁÃÅÇÉÈÊËÍÌÎÏÑÓÒÔÖÕÚÙÛÜÝ"
    Const SANS As String = "aaaaaaceeeeiiiinooooouuuuyyAAAAAACEEEEIIIINOOOOOUUUUY"
    Dim i As Long, p As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        p = InStr(1, AVEC, ch, vbBinaryCompare)
        If p > 0 Then Mid$(s, i, 1) = Mid$(SANS, p, 1)
    Next i
    SansAccents = s
End Function

Public Function NomFichier(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "-")
    Next c
    NomFichier = Trim$(s)
End Function

' ---------- Dossiers d'export ----------
Public Function DossierExport(ByVal sousDossier As String) As String
    Dim base As String, chemin As String, parties As Variant, i As Long
    base = ThisWorkbook.Path
    If base = "" Or LCase$(Left$(base, 4)) = "http" Then base = Environ$("USERPROFILE") & "\Documents"
    chemin = base & "\Export_PDF"
    If sousDossier <> "" Then chemin = chemin & "\" & NomFichier(sousDossier)
    parties = Split(Mid$(chemin, Len(base) + 2), "\")
    chemin = base
    For i = LBound(parties) To UBound(parties)
        chemin = chemin & "\" & parties(i)
        If Dir(chemin, vbDirectory) = "" Then MkDir chemin
    Next i
    DossierExport = chemin
End Function

Public Sub OuvrirDossier(ByVal chemin As String)
    On Error Resume Next
    Shell "explorer.exe """ & chemin & """", vbNormalFocus
End Sub

' ---------- Protection ----------
Public Sub Deverrouiller(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=MDP
End Sub

Public Sub Verrouiller(ws As Worksheet)
    If ModeAdmin Then Exit Sub
    If ws.CodeName = "shImpression" Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=MDP, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        UserInterfaceOnly:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Public Sub ProtegerTout()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Deverrouiller ws
        Verrouiller ws
    Next ws
End Sub

Public Sub ModeAdministrateur()
    Dim s As String, ws As Worksheet
    If ModeAdmin Then
        If MsgBox("Le mode administrateur est actif. Reverrouiller les feuilles ?", vbYesNo + vbQuestion) = vbYes Then
            ModeAdmin = False
            ProtegerTout
            MsgBox "Feuilles verrouillées.", vbInformation
        End If
        Exit Sub
    End If
    s = InputBox("Mot de passe administrateur :", "Mode administrateur")
    If s = "" Then Exit Sub
    If s <> MDP Then
        MsgBox "Mot de passe incorrect.", vbExclamation
        Exit Sub
    End If
    ModeAdmin = True
    For Each ws In ThisWorkbook.Worksheets
        Deverrouiller ws
    Next ws
    MsgBox "Mode administrateur actif : toutes les feuilles sont modifiables." & vbLf & _
           "Recliquez sur le bouton (ou fermez le classeur) pour reverrouiller.", vbInformation
End Sub

' ---------- Performances ----------
Public Sub DebutTraitement()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub FinTraitement()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ---------- Interface ----------
' Associe les macros aux zones de texte dont le texte de remplacement
' commence par « macro: » (les boutons sont dessinés à la génération).
Public Sub AttacherMacros()
    Dim ws As Worksheet, shp As Shape, alt As String
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            alt = ""
            On Error Resume Next
            alt = shp.AlternativeText
            On Error GoTo 0
            If LCase$(Left$(alt, 6)) = "macro:" Then
                shp.OnAction = Trim$(Mid$(alt, 7))
            End If
        Next shp
    Next ws
End Sub

Public Sub InitialiserClasseur()
    Dim ws As Worksheet
    On Error Resume Next
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        Deverrouiller ws
    Next ws
    AttacherMacros
    MajTousMembres
    ModeAdmin = False
    ProtegerTout
    Application.Calculate
    shAccueil.Activate
    Application.ScreenUpdating = True
End Sub

Public Sub Actualiser()
    On Error Resume Next
    Application.ScreenUpdating = False
    MajTousMembres
    Application.CalculateFull
    Application.ScreenUpdating = True
End Sub

Public Sub Sauvegarder()
    Dim dossier As String, base As String, nom As String
    On Error GoTo erreur
    ThisWorkbook.Save
    base = ThisWorkbook.Path
    If base = "" Or LCase$(Left$(base, 4)) = "http" Then base = Environ$("USERPROFILE") & "\Documents"
    dossier = base & "\Sauvegardes"
    If Dir(dossier, vbDirectory) = "" Then MkDir dossier
    nom = Replace(ThisWorkbook.Name, ".xlsm", "") & "_" & Format$(Now, "yyyy-mm-dd_hh\hnn") & ".xlsm"
    ThisWorkbook.SaveCopyAs dossier & "\" & nom
    MsgBox "Classeur enregistré." & vbLf & "Copie de sauvegarde : " & dossier & "\" & nom, vbInformation
    Exit Sub
erreur:
    MsgBox "Sauvegarde impossible : " & Err.Description, vbExclamation
End Sub

' ---------- Lignes génériques ----------
Public Sub AjouterLigne()
    Dim lo As ListObject, r As Range, ws As Worksheet
    Set ws = ActiveSheet
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    Deverrouiller ws
    Set r = NouvelleLigne(lo)
    If lo.Name = "tblEtudiants" Then r.Cells(1, 1).Value = Application.WorksheetFunction.Max(lo.ListColumns(1).Range) + 1
    Verrouiller ws
    If lo.Name = "tblEtudiants" Then r.Cells(1, 2).Select Else r.Cells(1, 1).Select
End Sub

Public Sub SupprimerLignes()
    Dim lo As ListObject, ws As Worksheet, zone As Range, i As Long, n As Long
    Dim ids As Collection, id As Variant
    Set ws = ActiveSheet
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set zone = Intersect(Selection, lo.DataBodyRange)
    If zone Is Nothing Then
        MsgBox "Sélectionnez d'abord une ou plusieurs lignes du tableau.", vbInformation
        Exit Sub
    End If
    Set ids = New Collection
    For i = lo.ListRows.Count To 1 Step -1
        If Not Intersect(zone.EntireRow, lo.ListRows(i).Range) Is Nothing Then
            n = n + 1
            ids.Add CStr(lo.ListRows(i).Range.Cells(1, 1).Value)
        End If
    Next i
    If MsgBox("Supprimer définitivement " & n & " ligne(s) ?" & _
              IIf(lo.Name = "tblSeances", vbLf & "Les présences de ces séances seront aussi supprimées.", ""), _
              vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    DebutTraitement
    Deverrouiller ws
    For i = lo.ListRows.Count To 1 Step -1
        If Not Intersect(zone.EntireRow, lo.ListRows(i).Range) Is Nothing Then lo.ListRows(i).Delete
    Next i
    Verrouiller ws
    If lo.Name = "tblSeances" Then
        For Each id In ids
            SupprimerPresences CStr(id)
        Next id
    End If
    FinTraitement
End Sub

Public Sub SupprimerPresences(ByVal id As String)
    Dim lo As ListObject, v As Variant, i As Long, fin As Long
    Set lo = Tbl("tblPresences")
    v = Donnees(lo)
    If NbLignes(v) = 0 Then Exit Sub
    Deverrouiller lo.Parent
    For i = UBound(v, 1) To 1 Step -1
        If CStr(v(i, 1)) = id Then
            fin = i
            Do While i > 1
                If CStr(v(i - 1, 1)) <> id Then Exit Do
                i = i - 1
            Loop
            lo.DataBodyRange.Rows(i & ":" & fin).Delete xlShiftUp
        End If
    Next i
    Verrouiller lo.Parent
End Sub
