Option Explicit
' ==========================================================================
'  modSaisie : feuille COURS DU JOUR (cahier de texte + appel en direct)
' ==========================================================================

Public Const A_ID As String = "C6"
Public Const A_ETAT As String = "D6"
Public Const A_CODE As String = "C7"
Public Const A_DATE As String = "C8"
Public Const A_DEBUT As String = "C9"
Public Const A_FIN As String = "C10"
Public Const A_PROF As String = "C12"
Public Const A_REMPL As String = "C13"
Public Const A_TYPE As String = "C14"
Public Const A_CONTENU As String = "H7"
Public Const A_TRAVAUX As String = "H12"
Public Const A_OBS As String = "I14"
Public Const A_FILTRE As String = "C19"
Public Const L1 As Long = 22        ' première ligne de l'appel
Public Const LMAX As Long = 221     ' dernière ligne de l'appel (200 étudiants)
' Colonnes de l'appel : B N° | C Nom | D Groupe | E Statut | F Arrivée | G Observation
Public Const C_NOM As Long = 3
Public Const C_GRP As Long = 4
Public Const C_STATUT As Long = 5
Public Const C_ARRIVEE As Long = 6
Public Const C_OBSE As Long = 7

Public Function ZoneStatuts() As Range
    Set ZoneStatuts = shSaisie.Range(shSaisie.Cells(L1, C_STATUT), shSaisie.Cells(LMAX, C_STATUT))
End Function

Private Function ZoneAppel() As Range
    Set ZoneAppel = shSaisie.Range(shSaisie.Cells(L1, 2), shSaisie.Cells(LMAX, C_OBSE))
End Function

Private Function NbStatutsSaisis() As Long
    NbStatutsSaisis = Application.WorksheetFunction.CountA(ZoneStatuts)
End Function

' Code matière du cours en cours d'après l'emploi du temps ("" si aucun).
Public Function CoursDuMoment() As String
    Dim lo As ListObject, v As Variant, i As Long, j As Long
    Dim t As Double, marge As Double, deb As Double, fin As Double
    Set lo = Tbl("tblEDT")
    v = Donnees(lo)
    j = Weekday(Date, vbMonday)
    t = CDbl(Time)
    marge = Nombre(Param("pMarge")) / 1440
    For i = 1 To NbLignes(v)
        If Nombre(v(i, Col(lo, "N° jour"))) = j Then
            deb = Nombre(v(i, Col(lo, "Début")))
            fin = Nombre(v(i, Col(lo, "Fin")))
            If t >= deb - marge And t < fin Then
                CoursDuMoment = CStr(v(i, Col(lo, "Code matière")))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ConfirmerAbandon() As Boolean
    ConfirmerAbandon = True
    If Trim$(CStr(shSaisie.Range(A_ID).Value)) = "" And NbStatutsSaisis() > 0 Then
        ConfirmerAbandon = (MsgBox("Un appel non enregistré est en cours de saisie." & vbLf & _
            "L'abandonner ?", vbYesNo + vbExclamation) = vbYes)
    End If
End Function

Private Sub ViderPanneau()
    With shSaisie
        .Range(A_ID).Value = ""
        .Range(A_CODE).Value = ""
        .Range(A_DATE).Value = ""
        .Range(A_DEBUT).Value = ""
        .Range(A_FIN).Value = ""
        .Range(A_PROF).Value = ""
        .Range(A_REMPL).Value = ""
        .Range(A_TYPE).Value = ""
        .Range(A_CONTENU).Value = ""
        .Range(A_TRAVAUX).Value = ""
        .Range(A_OBS).Value = ""
        .Range(A_FILTRE).Value = ""
    End With
End Sub

' ---------------------------------------------------------------- boutons
Public Sub DemarrerCours()
    Dim code As String
    If Not ConfirmerAbandon() Then Exit Sub
    code = CoursDuMoment()
    DebutTraitement
    Deverrouiller shSaisie
    ViderPanneau
    With shSaisie
        .Range(A_CODE).Value = code
        .Range(A_DATE).Value = Date
        .Range(A_DEBUT).Value = TimeSerial(Hour(Now), Minute(Now), 0)
        .Range(A_PROF).Value = "Présent"
        .Range(A_TYPE).Value = "Cours"
    End With
    ChargerListe code, "", False
    Verrouiller shSaisie
    FinTraitement
    shSaisie.Activate
    shSaisie.Range(A_CODE).Select
    If code = "" Then
        MsgBox "Aucun cours n'est prévu à cette heure dans l'emploi du temps." & vbLf & _
               "Choisissez la matière dans la case « Matière », puis cliquez sur ENREGISTRER.", vbInformation
        Exit Sub
    End If
    If Enregistrer(False) Then
        MsgBox "Cours démarré : " & code & " - " & InfoMatiere(code, "Intitulé") & vbLf & _
               "Séance n° " & shSaisie.Range(A_ID).Value & " créée dans le cahier de texte." & vbLf & vbLf & _
               "Indiquez la présence du professeur, faites l'appel puis cliquez sur ENREGISTRER.", vbInformation
    End If
End Sub

Public Sub EnregistrerSeance()
    If Enregistrer(True) Then
        MsgBox "Séance " & shSaisie.Range(A_ID).Value & " enregistrée." & vbLf & ResumeAppel(), vbInformation
    End If
End Sub

Public Sub CloturerSeance()
    If Trim$(CStr(shSaisie.Range(A_CODE).Value)) = "" Then
        MsgBox "Aucune séance en cours.", vbInformation
        Exit Sub
    End If
    Deverrouiller shSaisie
    Application.EnableEvents = False
    shSaisie.Range(A_FIN).Value = TimeSerial(Hour(Now), Minute(Now), 0)
    Application.EnableEvents = True
    Verrouiller shSaisie
    If Enregistrer(True) Then
        MsgBox "Séance " & shSaisie.Range(A_ID).Value & " clôturée à " & Format$(shSaisie.Range(A_FIN).Value, "hh:nn") & "." & _
               vbLf & ResumeAppel(), vbInformation
    End If
End Sub

Public Sub NouvelleSeance()
    If Not ConfirmerAbandon() Then Exit Sub
    DebutTraitement
    Deverrouiller shSaisie
    ViderPanneau
    shSaisie.Range(A_DATE).Value = Date
    shSaisie.Range(A_TYPE).Value = "Cours"
    shSaisie.Range(A_PROF).Value = "Présent"
    ChargerListe "", "", False
    Verrouiller shSaisie
    FinTraitement
    shSaisie.Range(A_CODE).Select
End Sub

Public Sub TousPresents()
    Dim r As Long, n As Long
    Application.EnableEvents = False
    For r = L1 To LMAX
        If Trim$(CStr(shSaisie.Cells(r, C_NOM).Value)) <> "" And Not shSaisie.Rows(r).Hidden Then
            If Trim$(CStr(shSaisie.Cells(r, C_STATUT).Value)) = "" Then
                shSaisie.Cells(r, C_STATUT).Value = "P"
                n = n + 1
            End If
        End If
    Next r
    Application.EnableEvents = True
    MsgBox n & " étudiant(s) marqué(s) présent(s). Corrigez les absents (A), retards (R) et excusés (E).", vbInformation
End Sub

Private Function ResumeAppel() As String
    Dim r As Long, p As Long, a As Long, rt As Long, e As Long, s As String
    For r = L1 To LMAX
        Select Case UCase$(Trim$(CStr(shSaisie.Cells(r, C_STATUT).Value)))
            Case "P": p = p + 1
            Case "A": a = a + 1
            Case "R": rt = rt + 1
            Case "E": e = e + 1
        End Select
    Next r
    s = "Présents : " & p & "   Absents : " & a & "   Retards : " & rt & "   Excusés : " & e
    ResumeAppel = s
End Function

' ---------------------------------------------------------------- liste
' Charge la liste d'appel de la matière. Si id <> "", reprend les présences
' enregistrées ; si conserver = True, garde les statuts déjà saisis.
Public Sub ChargerListe(ByVal code As String, ByVal id As String, ByVal conserver As Boolean)
    Dim liste As Variant, n As Long, i As Long, out() As Variant
    Dim dic As Object, v As Variant, lo As ListObject, k As String, r As Long
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    If conserver Then
        For r = L1 To LMAX
            k = Trim$(CStr(shSaisie.Cells(r, C_NOM).Value))
            If k <> "" Then dic(k) = Array(shSaisie.Cells(r, C_STATUT).Value, _
                shSaisie.Cells(r, C_ARRIVEE).Value, shSaisie.Cells(r, C_OBSE).Value)
        Next r
    ElseIf id <> "" Then
        Set lo = Tbl("tblPresences")
        v = Donnees(lo)
        For i = 1 To NbLignes(v)
            If CStr(v(i, 1)) = id Then
                dic(CStr(v(i, Col(lo, "Étudiant")))) = Array(v(i, Col(lo, "Statut")), _
                    v(i, Col(lo, "Arrivée")), v(i, Col(lo, "Observation")))
            End If
        Next i
    End If

    Deverrouiller shSaisie
    shSaisie.Range(shSaisie.Cells(L1, 1), shSaisie.Cells(LMAX, 1)).EntireRow.Hidden = False
    ZoneAppel.ClearContents
    liste = ListeEtudiantsMatiere(code)
    n = NbLignes(liste)
    If n > LMAX - L1 + 1 Then n = LMAX - L1 + 1
    If n > 0 Then
        ReDim out(1 To n, 1 To 6)
        For i = 1 To n
            out(i, 1) = i
            out(i, 2) = liste(i, 1)
            out(i, 3) = liste(i, 2)
            If dic.Exists(CStr(liste(i, 1))) Then
                out(i, 4) = dic(CStr(liste(i, 1)))(0)
                out(i, 5) = dic(CStr(liste(i, 1)))(1)
                out(i, 6) = dic(CStr(liste(i, 1)))(2)
            End If
        Next i
        shSaisie.Cells(L1, 2).Resize(n, 6).Value = out
    End If
    Verrouiller shSaisie
End Sub

Public Sub AppliquerFiltre()
    Dim g As String, r As Long
    g = Trim$(CStr(shSaisie.Range(A_FILTRE).Value))
    Application.ScreenUpdating = False
    Deverrouiller shSaisie
    For r = L1 To LMAX
        If g = "" Or Trim$(CStr(shSaisie.Cells(r, C_NOM).Value)) = "" Then
            shSaisie.Rows(r).Hidden = False
        Else
            shSaisie.Rows(r).Hidden = (StrComp(Trim$(CStr(shSaisie.Cells(r, C_GRP).Value)), g, vbTextCompare) <> 0)
        End If
    Next r
    Verrouiller shSaisie
    Application.ScreenUpdating = True
End Sub

Public Sub MatiereChangee()
    Dim code As String
    code = Trim$(CStr(shSaisie.Range(A_CODE).Value))
    Application.EnableEvents = False
    ChargerListe code, "", True
    shSaisie.Range(A_FILTRE).Value = ""
    Application.EnableEvents = True
End Sub

' Recharge une séance enregistrée choisie dans la case N° séance.
Public Sub ChargerSeance(ByVal id As String)
    Dim lo As ListObject, r As Range
    If id = "" Then Exit Sub
    Set lo = Tbl("tblSeances")
    Set r = LigneParID(lo, id)
    If r Is Nothing Then
        MsgBox "Séance " & id & " introuvable.", vbExclamation
        Exit Sub
    End If
    DebutTraitement
    Deverrouiller shSaisie
    With shSaisie
        .Range(A_CODE).Value = r.Cells(1, Col(lo, "Code matière")).Value
        .Range(A_DATE).Value = r.Cells(1, Col(lo, "Date")).Value
        .Range(A_DEBUT).Value = r.Cells(1, Col(lo, "Début")).Value
        .Range(A_FIN).Value = r.Cells(1, Col(lo, "Fin")).Value
        .Range(A_PROF).Value = r.Cells(1, Col(lo, "Présence prof")).Value
        .Range(A_REMPL).Value = r.Cells(1, Col(lo, "Remplaçant ou motif")).Value
        .Range(A_TYPE).Value = r.Cells(1, Col(lo, "Type de séance")).Value
        .Range(A_CONTENU).Value = r.Cells(1, Col(lo, "Contenu / chapitres traités")).Value
        .Range(A_TRAVAUX).Value = r.Cells(1, Col(lo, "Travaux demandés")).Value
        .Range(A_OBS).Value = r.Cells(1, Col(lo, "Observations")).Value
        .Range(A_FILTRE).Value = ""
    End With
    ChargerListe CStr(r.Cells(1, Col(lo, "Code matière")).Value), id, False
    Verrouiller shSaisie
    FinTraitement
End Sub

' ---------------------------------------------------------------- écriture
' tblPresences : ID | Étudiant | Groupe | Code matière | Matière | Date | Heure | Statut | Arrivée | Observation | Saisie le
Private Function Enregistrer(ByVal avecAppel As Boolean) As Boolean
    Dim code As String, id As String, lo As ListObject, r As Range
    Dim nbVides As Long, nbSaisis As Long, rep As VbMsgBoxResult, remplir As String
    Dim faireAppel As Boolean, ligne As Long, n As Long, data() As Variant
    Dim intitule As String, statut As String

    code = Trim$(CStr(shSaisie.Range(A_CODE).Value))
    If code = "" Or Not MatiereExiste(code) Then
        MsgBox "Choisissez d'abord une matière valide dans la case « Matière ».", vbExclamation
        Exit Function
    End If
    If Not IsDate(shSaisie.Range(A_DATE).Value) Then
        MsgBox "La date de la séance est invalide.", vbExclamation
        Exit Function
    End If

    If avecAppel Then
        For ligne = L1 To LMAX
            If Trim$(CStr(shSaisie.Cells(ligne, C_NOM).Value)) <> "" Then
                If Trim$(CStr(shSaisie.Cells(ligne, C_STATUT).Value)) = "" Then nbVides = nbVides + 1 Else nbSaisis = nbSaisis + 1
            End If
        Next ligne
        faireAppel = True
        If nbSaisis = 0 And shSaisie.Range(A_PROF).Value = "Absent" Then
            faireAppel = False
        ElseIf nbVides > 0 Then
            rep = MsgBox(nbVides & " étudiant(s) n'ont pas encore de statut." & vbLf & vbLf & _
                "OUI : les marquer ABSENTS (A)" & vbLf & _
                "NON : les marquer PRÉSENTS (P)" & vbLf & _
                "ANNULER : revenir à l'appel", vbYesNoCancel + vbQuestion, "Appel incomplet")
            If rep = vbCancel Then Exit Function
            remplir = IIf(rep = vbYes, "A", "P")
        End If
    End If

    DebutTraitement
    On Error GoTo erreur
    Set lo = Tbl("tblSeances")
    Deverrouiller lo.Parent
    id = Trim$(CStr(shSaisie.Range(A_ID).Value))
    If id <> "" Then Set r = LigneParID(lo, id)
    If r Is Nothing Then
        If id = "" Then id = NouvelID(lo, "S", 4)
        Set r = NouvelleLigne(lo)
    End If
    r.Cells(1, Col(lo, "ID")).Value = id
    r.Cells(1, Col(lo, "Date")).Value = CDate(shSaisie.Range(A_DATE).Value)
    r.Cells(1, Col(lo, "Début")).Value = shSaisie.Range(A_DEBUT).Value
    r.Cells(1, Col(lo, "Fin")).Value = shSaisie.Range(A_FIN).Value
    r.Cells(1, Col(lo, "Code matière")).Value = code
    r.Cells(1, Col(lo, "Présence prof")).Value = shSaisie.Range(A_PROF).Value
    r.Cells(1, Col(lo, "Remplaçant ou motif")).Value = shSaisie.Range(A_REMPL).Value
    r.Cells(1, Col(lo, "Type de séance")).Value = shSaisie.Range(A_TYPE).Value
    r.Cells(1, Col(lo, "Contenu / chapitres traités")).Value = shSaisie.Range(A_CONTENU).Value
    r.Cells(1, Col(lo, "Travaux demandés")).Value = shSaisie.Range(A_TRAVAUX).Value
    r.Cells(1, Col(lo, "Observations")).Value = shSaisie.Range(A_OBS).Value
    Verrouiller lo.Parent

    Deverrouiller shSaisie
    shSaisie.Range(A_ID).Value = id

    If avecAppel And faireAppel Then
        intitule = InfoMatiere(code, "Intitulé")
        ReDim data(1 To LMAX - L1 + 1, 1 To 11)
        For ligne = L1 To LMAX
            If Trim$(CStr(shSaisie.Cells(ligne, C_NOM).Value)) <> "" Then
                statut = UCase$(Trim$(CStr(shSaisie.Cells(ligne, C_STATUT).Value)))
                If statut = "" Then
                    statut = remplir
                    shSaisie.Cells(ligne, C_STATUT).Value = statut
                End If
                n = n + 1
                data(n, 1) = id
                data(n, 2) = shSaisie.Cells(ligne, C_NOM).Value
                data(n, 3) = shSaisie.Cells(ligne, C_GRP).Value
                data(n, 4) = code
                data(n, 5) = intitule
                data(n, 6) = CDate(shSaisie.Range(A_DATE).Value)
                data(n, 7) = shSaisie.Range(A_DEBUT).Value
                data(n, 8) = statut
                data(n, 9) = shSaisie.Cells(ligne, C_ARRIVEE).Value
                data(n, 10) = shSaisie.Cells(ligne, C_OBSE).Value
                data(n, 11) = Now
            End If
        Next ligne
        SupprimerPresences id
        If n > 0 Then
            Set lo = Tbl("tblPresences")
            Deverrouiller lo.Parent
            AjouterBloc lo, Redim2D(data, n, 11)
            Verrouiller lo.Parent
        End If
    End If
    Verrouiller shSaisie
    FinTraitement
    Enregistrer = True
    Exit Function
erreur:
    FinTraitement
    ProtegerTout
    MsgBox "Erreur pendant l'enregistrement : " & Err.Description, vbCritical
End Function

Private Function Redim2D(data() As Variant, ByVal n As Long, ByVal nc As Long) As Variant
    Dim out() As Variant, i As Long, j As Long
    ReDim out(1 To n, 1 To nc)
    For i = 1 To n
        For j = 1 To nc
            out(i, j) = data(i, j)
        Next j
    Next i
    Redim2D = out
End Function

' ---------------------------------------------------------------- statuts
Public Sub CyclerStatut(ByVal c As Range)
    Dim s As String
    If Trim$(CStr(shSaisie.Cells(c.Row, C_NOM).Value)) = "" Then Exit Sub
    Select Case UCase$(Trim$(CStr(c.Value)))
        Case "": s = "P"
        Case "P": s = "A"
        Case "A": s = "R"
        Case "R": s = "E"
        Case Else: s = "P"
    End Select
    Application.EnableEvents = False
    c.Value = s
    MajArrivee c
    Application.EnableEvents = True
End Sub

Public Sub NormaliserStatuts(ByVal zone As Range)
    Dim c As Range, s As String
    Application.EnableEvents = False
    For Each c In zone.Cells
        s = UCase$(Left$(Trim$(CStr(c.Value)), 1))
        If s <> "" And InStr("PARE", s) = 0 Then s = ""
        If CStr(c.Value) <> s Then c.Value = s
        MajArrivee c
    Next c
    Application.EnableEvents = True
End Sub

Private Sub MajArrivee(ByVal c As Range)
    Dim a As Range
    Set a = shSaisie.Cells(c.Row, C_ARRIVEE)
    If c.Value = "R" Then
        If Trim$(CStr(a.Value)) = "" Then a.Value = TimeSerial(Hour(Now), Minute(Now), 0)
    ElseIf c.Value = "P" Or c.Value = "A" Or c.Value = "" Then
        a.Value = ""
    End If
End Sub
