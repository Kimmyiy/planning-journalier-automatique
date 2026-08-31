Option Explicit

'==============================================================================
' MODULE : Module_Groupes
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'
' Description :
'   Gestion des groupes auxiliaires (G1/G2) par weekend et jour férié.
'   Stockage dans la feuille "Groupes" (tableau Dates_Groupes).
'
' Structure de la feuille Groupes :
'   Col A : Date
'   Col B : Groupe        (G1 / G2)
'   Col C : Type          (Weekend / Ferie)
'   Col D : Modif_Manuelle (Oui / Non)
'
' Règles métier :
'   - Fériés en semaine (mar-ven) : groupe indépendant (suit l'alternance)
'   - Lundi férié directement après un WE : même groupe que ce WE (Pentecôte)
'   - WE de 4 jours (ex: Ascension) : le férié et le WE peuvent avoir
'     des groupes différents — séparation via modification manuelle
'   - Tournus annuel : si 2026 commence G1, 2027 commence G2
'   - Les modifications manuelles sont conservées lors d'une régénération
'==============================================================================

' Colonnes de la feuille Groupes
Private Const GRP_COL_DATE   As Long = 1  ' A — Date
Private Const GRP_COL_GROUPE As Long = 2  ' B — Groupe (G1/G2)
Private Const GRP_COL_TYPE   As Long = 3  ' C — Type (Weekend/Ferie)
Private Const GRP_COL_MANUEL As Long = 4  ' D — Modif_Manuelle (Oui/Non)
Private Const LIGNE_DEBUT    As Long = 2  ' Ligne 1 = en-têtes

'==============================================================================
' FUNCTION : GetGroupeJour
' Retourne le groupe (G1/G2) pour une date donnée.
' Lit depuis le tableau Dates_Groupes. Retourne "" si non trouvé.
'==============================================================================
Function GetGroupeJour(ByVal dateJour As Date) As String

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(NOM_FEUILLE_GROUPES)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(NOM_TBL_GROUPES)
    On Error GoTo 0
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then Exit Function

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        Dim d As Date
        On Error Resume Next
        d = tbl.DataBodyRange(i, GRP_COL_DATE).Value
        On Error GoTo 0
        If Int(d) = Int(dateJour) Then
            GetGroupeJour = tbl.DataBodyRange(i, GRP_COL_GROUPE).Value
            Exit Function
        End If
    Next i

    GetGroupeJour = ""

End Function

'==============================================================================
' SUB : SetGroupeJour
' Attribue manuellement un groupe à une date.
' Crée la ligne si elle n'existe pas, met à jour si elle existe.
' Si groupe = "" -> restaure le calcul automatique.
'==============================================================================
Sub SetGroupeJour(ByVal dateJour As Date, ByVal groupe As String)

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_GROUPES)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(NOM_TBL_GROUPES)

    ' Cherche si la date existe déjà
    Dim ligneExiste As Long
    ligneExiste = 0

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        Dim d As Date
        On Error Resume Next
        d = tbl.DataBodyRange(i, GRP_COL_DATE).Value
        On Error GoTo 0
        If Int(d) = Int(dateJour) Then
            ligneExiste = i
            Exit For
        End If
    Next i

    If groupe = "" Then
        ' Suppression : remet le groupe automatique
        If ligneExiste > 0 Then
            tbl.DataBodyRange(ligneExiste, GRP_COL_MANUEL).Value = "Non"
            tbl.DataBodyRange(ligneExiste, GRP_COL_GROUPE).Value = CalculerGroupeAuto(dateJour)
        End If
        Exit Sub
    End If

    If ligneExiste > 0 Then
        ' Met à jour la ligne existante
        tbl.DataBodyRange(ligneExiste, GRP_COL_GROUPE).Value = groupe
        tbl.DataBodyRange(ligneExiste, GRP_COL_MANUEL).Value = "Oui"
    Else
        ' Crée une nouvelle ligne dans le tableau
        Dim nouvRow As ListRow
        Set nouvRow = tbl.ListRows.Add

        Dim jourSem As Long
        jourSem = Weekday(dateJour, vbMonday)

        nouvRow.Range(1, GRP_COL_DATE).Value = dateJour
        nouvRow.Range(1, GRP_COL_DATE).NumberFormat = "dd.mm.yyyy"
        nouvRow.Range(1, GRP_COL_GROUPE).Value = groupe
        nouvRow.Range(1, GRP_COL_TYPE).Value = IIf(jourSem = 6 Or jourSem = 7, "Weekend", "Ferie")
        nouvRow.Range(1, GRP_COL_MANUEL).Value = "Oui"

        Call TrierGroupes(tbl)
    End If

End Sub

'==============================================================================
' FUNCTION : CalculerGroupeAuto (privée)
' Calcule le groupe automatique pour une date basée sur Ref_GroupeG1.
'==============================================================================
Private Function CalculerGroupeAuto(ByVal dateJour As Date) As String

    Dim dateRefG1 As Date
    On Error Resume Next
    dateRefG1 = Sheets(NOM_FEUILLE_PLANNING).Range(CELL_REF_GROUPE_G1).Value
    On Error GoTo 0

    If dateRefG1 = 0 Then
        CalculerGroupeAuto = GROUPE_G1
        Exit Function
    End If

    Dim jourSem As Long
    jourSem = Weekday(dateJour, vbMonday)
    Dim lundiSem As Date
    lundiSem = dateJour - (jourSem - 1)

    Dim nbSem As Long
    nbSem = Int((lundiSem - dateRefG1) / 7)
    CalculerGroupeAuto = IIf(nbSem Mod 2 = 0, GROUPE_G1, GROUPE_G2)

End Function

'==============================================================================
' SUB : GenererGroupesAnnee
' Génère automatiquement tous les weekends et jours fériés d'une année.
' Respecte le tournus annuel et conserve les modifications manuelles.
'==============================================================================
Sub GenererGroupesAnnee(ByVal annee As Long)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(NOM_FEUILLE_GROUPES)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_GROUPES & "' n'existe pas.", vbExclamation
        Exit Sub
    End If

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(NOM_TBL_GROUPES)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Le tableau '" & NOM_TBL_GROUPES & "' n'existe pas.", vbExclamation
        Exit Sub
    End If

    ' Détermine le groupe du premier weekend de l'année
    Dim groupeCourant As String
    groupeCourant = DeterminerGroupePremierWE(annee, tbl)

    ' Collecte les modifications manuelles existantes pour cette année
    Dim datesManuelles() As Date
    Dim groupesManuels()  As String
    Dim nbManuels As Long
    nbManuels = 0
    ReDim datesManuelles(1 To 400)
    ReDim groupesManuels(1 To 400)

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        Dim d As Date
        On Error Resume Next
        d = tbl.DataBodyRange(i, GRP_COL_DATE).Value
        On Error GoTo 0
        If Year(d) = annee And tbl.DataBodyRange(i, GRP_COL_MANUEL).Value = "Oui" Then
            nbManuels = nbManuels + 1
            datesManuelles(nbManuels) = d
            groupesManuels(nbManuels) = tbl.DataBodyRange(i, GRP_COL_GROUPE).Value
        End If
    Next i

    ' Supprime les entrées automatiques de cette année
    i = tbl.ListRows.count
    Do While i >= 1
        On Error Resume Next
        d = tbl.DataBodyRange(i, GRP_COL_DATE).Value
        On Error GoTo 0
        If Year(d) = annee And tbl.DataBodyRange(i, GRP_COL_MANUEL).Value <> "Oui" Then
            tbl.ListRows(i).Delete
        End If
        i = i - 1
    Loop

    ' Tableau temporaire
    Dim tmpDates()   As Date
    Dim tmpGroupes() As String
    Dim tmpTypes()   As String
    Dim tmpManuels() As String
    ReDim tmpDates(1 To 400)
    ReDim tmpGroupes(1 To 400)
    ReDim tmpTypes(1 To 400)
    ReDim tmpManuels(1 To 400)
    Dim nbJours As Long
    nbJours = 0

    Dim groupeSamediCourant As String
    groupeSamediCourant = groupeCourant

    Dim dateDebut As Date: dateDebut = DateSerial(annee, 1, 1)
    Dim dateFin   As Date: dateFin = DateSerial(annee, 12, 31)

    Dim dateCase As Date
    For dateCase = dateDebut To dateFin

        Dim jourSem As Long
        jourSem = Weekday(dateCase, vbMonday)

        Dim estWE    As Boolean
        Dim estFerie As Boolean
        estWE = (jourSem = 6 Or jourSem = 7)
        estFerie = Module_Feries.estFerie(dateCase)

        If Not (estWE Or estFerie) Then GoTo SuivantJour

        ' Vérifie modification manuelle
        Dim groupeJour As String
        Dim estManuel  As Boolean
        estManuel = False
        groupeJour = ""

        Dim j As Long
        For j = 1 To nbManuels
            If Int(datesManuelles(j)) = Int(dateCase) Then
                groupeJour = groupesManuels(j)
                estManuel = True
                Exit For
            End If
        Next j

        If Not estManuel Then
            If jourSem = 6 Then
                groupeJour = groupeCourant
                groupeSamediCourant = groupeCourant
            ElseIf jourSem = 7 Then
                groupeJour = TrouverGroupeDate(dateCase - 1, tmpDates, tmpGroupes, nbJours)
                If groupeJour = "" Then groupeJour = groupeSamediCourant
            ElseIf jourSem = 1 And estFerie Then
                Dim groupeDim As String
                groupeDim = TrouverGroupeDate(dateCase - 1, tmpDates, tmpGroupes, nbJours)
                If groupeDim <> "" Then
                    groupeJour = groupeDim
                Else
                    groupeJour = groupeCourant
                End If
            Else
                If estWE And estFerie Then
                    If jourSem = 6 Then
                        groupeJour = groupeCourant
                        groupeSamediCourant = groupeCourant
                    Else
                        groupeJour = TrouverGroupeDate(dateCase - 1, tmpDates, tmpGroupes, nbJours)
                        If groupeJour = "" Then groupeJour = groupeSamediCourant
                    End If
                Else
                    groupeJour = groupeCourant
                End If
            End If
        End If

        ' Ajoute au tableau temporaire
        nbJours = nbJours + 1
        tmpDates(nbJours) = dateCase
        tmpGroupes(nbJours) = groupeJour
        tmpManuels(nbJours) = IIf(estManuel, "Oui", "Non")

        If estWE And Not estFerie Then
            tmpTypes(nbJours) = "Weekend"
        ElseIf estFerie And Not estWE Then
            tmpTypes(nbJours) = "Ferie"
        Else
            tmpTypes(nbJours) = "Weekend"
        End If

        ' Alternance : change de groupe après le dimanche
        If jourSem = 7 And Not estManuel Then
            groupeCourant = IIf(groupeCourant = GROUPE_G1, GROUPE_G2, GROUPE_G1)
        End If

SuivantJour:
    Next dateCase

    ' Écriture dans le tableau ListObject
    For i = 1 To nbJours
        Dim nouvRow As ListRow
        Set nouvRow = tbl.ListRows.Add

        nouvRow.Range(1, GRP_COL_DATE).Value = tmpDates(i)
        nouvRow.Range(1, GRP_COL_DATE).NumberFormat = "dd.mm.yyyy"
        nouvRow.Range(1, GRP_COL_GROUPE).Value = tmpGroupes(i)
        nouvRow.Range(1, GRP_COL_TYPE).Value = tmpTypes(i)
        nouvRow.Range(1, GRP_COL_MANUEL).Value = tmpManuels(i)
    Next i

    Call TrierGroupes(tbl)

    MsgBox "Groupes " & annee & " générés avec succès !", vbInformation, "Planning v2.0"

End Sub

'==============================================================================
' FUNCTION : TrouverGroupeDate (privée)
' Cherche le groupe d'une date dans le tableau temporaire.
'==============================================================================
Private Function TrouverGroupeDate(ByVal dateJour As Date, _
                                    ByRef dates() As Date, _
                                    ByRef groupes() As String, _
                                    ByVal nb As Long) As String
    Dim i As Long
    For i = 1 To nb
        If Int(dates(i)) = Int(dateJour) Then
            TrouverGroupeDate = groupes(i)
            Exit Function
        End If
    Next i
    TrouverGroupeDate = ""
End Function

'==============================================================================
' FUNCTION : DeterminerGroupePremierWE (privée)
' Détermine le groupe du premier weekend de l'année.
' Inverse le groupe du dernier weekend de l'année précédente.
'==============================================================================
Private Function DeterminerGroupePremierWE(ByVal annee As Long, _
                                             ByVal tbl As ListObject) As String

    ' Cherche le dernier weekend de l'année précédente dans le tableau
    Dim i As Long
    For i = tbl.ListRows.count To 1 Step -1
        Dim d As Date
        On Error Resume Next
        d = tbl.DataBodyRange(i, GRP_COL_DATE).Value
        On Error GoTo 0
        If Year(d) = annee - 1 Then
            Dim jourSem As Long
            jourSem = Weekday(d, vbMonday)
            If jourSem = 6 Or jourSem = 7 Or Module_Feries.estFerie(d) Then
                DeterminerGroupePremierWE = _
                    IIf(tbl.DataBodyRange(i, GRP_COL_GROUPE).Value = GROUPE_G1, GROUPE_G2, GROUPE_G1)
                Exit Function
            End If
        End If
    Next i

    ' Fallback : calcul depuis Ref_GroupeG1
    Dim dateRefG1 As Date
    On Error Resume Next
    dateRefG1 = Sheets(NOM_FEUILLE_PLANNING).Range(CELL_REF_GROUPE_G1).Value
    On Error GoTo 0

    Dim premierSam As Date
    premierSam = DateSerial(annee, 1, 1)
    Do While Weekday(premierSam, vbMonday) <> 6
        premierSam = premierSam + 1
    Loop

    Dim lundiSem As Date
    lundiSem = premierSam - 5
    Dim nbSem As Long
    nbSem = Int((lundiSem - dateRefG1) / 7)
    DeterminerGroupePremierWE = IIf(nbSem Mod 2 = 0, GROUPE_G1, GROUPE_G2)

End Function

'==============================================================================
' SUB : TrierGroupes (privée)
' Trie le tableau Dates_Groupes par date croissante.
'==============================================================================
Private Sub TrierGroupes(ByVal tbl As ListObject)

    If tbl.ListRows.count < 2 Then Exit Sub

    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(GRP_COL_DATE).Range, _
                             Order:=xlAscending
    tbl.Sort.Apply

End Sub

'==============================================================================
' FUNCTION : EstWeekendOuFerie
' Retourne True si la date est un weekend ou un jour férié.
' Utilisée par le calendrier auxiliaires.
'==============================================================================
Function EstWeekendOuFerie(ByVal dateJour As Date) As Boolean
    Dim jourSem As Long
    jourSem = Weekday(dateJour, vbMonday)
    EstWeekendOuFerie = (jourSem = 6 Or jourSem = 7 Or Module_Feries.estFerie(dateJour))
End Function