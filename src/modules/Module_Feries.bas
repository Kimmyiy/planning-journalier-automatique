Option Explicit

'==============================================================================
' MODULE : Module_Feries
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'
' Description :
'   Gestion des jours fériés du canton de Neuchâtel.
'   - Calcul automatique via l'algorithme de Butcher (Pâques)
'   - Fériés fixes et calculés pour Neuchâtel
'   - Ajout manuel de jours fériés exceptionnels
'   - Gestion des fermetures d'entreprise (type "Fermeture")
'   - Génération de la grille mensuelle pour les calendriers
'
' Feuille utilisée : NOM_FEUILLE_FERIES ("Feries")
'   Colonne A : Date
'   Colonne B : Nom du jour férié
'   Colonne C : Type (Fixe / Calcule / Manuel / Fermeture)
'==============================================================================

' Colonnes de la feuille Feries
Private Const CF_DATE As Long = 1  ' A - Date
Private Const CF_NOM  As Long = 2  ' B - Nom
Private Const CF_TYPE As Long = 3  ' C - Type

'==============================================================================
' FUNCTION : DatePaques
' Calcule la date de Pâques pour une année donnée.
' Algorithme de Butcher (méthode analytique).
'==============================================================================
Function DatePaques(ByVal annee As Long) As Date

    Dim a As Long, b As Long, c As Long, d As Long, e As Long
    Dim f As Long, g As Long, h As Long, i As Long, k As Long
    Dim l As Long, m As Long, jour As Long, mois As Long

    a = annee Mod 19
    b = annee \ 100
    c = annee Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    mois = (h + l - 7 * m + 114) \ 31
    jour = ((h + l - 7 * m + 114) Mod 31) + 1

    DatePaques = DateSerial(annee, mois, jour)

End Function

'==============================================================================
' FUNCTION : DateJeuneFederal
' Calcule la date du Jeçne fédéral neuchâtelois :
' 3e dimanche de septembre + 1 jour = lundi.
'==============================================================================
Function DateJeuneFederal(ByVal annee As Long) As Date

    Dim d As Date
    d = DateSerial(annee, 9, 1)

    ' Trouve le 1er dimanche de septembre
    Do While Weekday(d, vbMonday) <> 7
        d = d + 1
    Loop

    ' 3e dimanche = +14, lundi suivant = +15
    DateJeuneFederal = d + 15

End Function

'==============================================================================
' SUB : ChargerFeriesAnnee
' Génère et enregistre les jours fériés de Neuchâtel pour une année.
' Supprime les fériés automatiques existants, conserve les manuels.
'==============================================================================
Sub ChargerFeriesAnnee(ByVal annee As Long)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(NOM_FEUILLE_FERIES)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_FERIES & "' n'existe pas.", vbExclamation
        Exit Sub
    End If

    ' Supprime uniquement les fériés automatiques de cette année.
    ' IMPORTANT : les fermetures d'entreprise (TYPE_FERIE_FERMETURE) sont
    ' préservées au même titre que les fériés manuels — sans ce garde-fou,
    ' toute ré-exécution de ChargerFeriesAnnee effacerait silencieusement
    ' les périodes de fermeture déjà saisies pour l'année.
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(NOM_TBL_FERIES)

    Dim i As Long
    i = ws.Cells(ws.Rows.count, CF_DATE).End(xlUp).Row
    Do While i >= 2
        Dim d As Date
        d = 0
        On Error Resume Next
        d = ws.Cells(i, CF_DATE).Value
        On Error GoTo 0
        If d <> 0 Then
            If Year(d) = annee _
               And ws.Cells(i, CF_TYPE).Value <> TYPE_FERIE_MANUEL _
               And ws.Cells(i, CF_TYPE).Value <> TYPE_FERIE_FERMETURE Then
                ' Suppression via ListObject (convention du projet) au lieu
                ' de ws.Rows(i).Delete, qui efface toute la largeur de la ligne.
                tbl.ListRows(i - tbl.HeaderRowRange.Row).Delete
            End If
        End If
        i = i - 1
    Loop

    ' Calcul des dates mobiles à partir de Pâques
    Dim paques      As Date: paques = DatePaques(annee)
    Dim vendrediSt  As Date: vendrediSt = paques - 2
    Dim lundiPaques As Date: lundiPaques = paques + 1
    Dim ascension   As Date: ascension = paques + 39
    Dim lundiPente  As Date: lundiPente = paques + 50
    Dim jeune       As Date: jeune = DateJeuneFederal(annee)

    ' Tableau des fériés Neuchâtel
    Dim feries(11, 2) As Variant
    feries(0, 0) = DateSerial(annee, 1, 1): feries(0, 1) = "Nouvel An": feries(0, 2) = TYPE_FERIE_FIXE
    feries(1, 0) = DateSerial(annee, 3, 1): feries(1, 1) = "Fete de la Republique": feries(1, 2) = TYPE_FERIE_FIXE
    feries(2, 0) = vendrediSt: feries(2, 1) = "Vendredi Saint": feries(2, 2) = TYPE_FERIE_CALCULE
    feries(3, 0) = lundiPaques: feries(3, 1) = "Lundi de Paques": feries(3, 2) = TYPE_FERIE_CALCULE
    feries(4, 0) = DateSerial(annee, 5, 1): feries(4, 1) = "Fete du Travail": feries(4, 2) = TYPE_FERIE_FIXE
    feries(5, 0) = ascension: feries(5, 1) = "Ascension": feries(5, 2) = TYPE_FERIE_CALCULE
    feries(6, 0) = lundiPente: feries(6, 1) = "Lundi de Pentecote": feries(6, 2) = TYPE_FERIE_CALCULE
    feries(7, 0) = DateSerial(annee, 8, 1): feries(7, 1) = "Fete nationale": feries(7, 2) = TYPE_FERIE_FIXE
    feries(8, 0) = jeune: feries(8, 1) = "Jeune federal": feries(8, 2) = TYPE_FERIE_CALCULE
    feries(9, 0) = DateSerial(annee, 12, 25): feries(9, 1) = "Noel": feries(9, 2) = TYPE_FERIE_FIXE

    ' Reports du 1er janvier si tombe un dimanche
    feries(10, 0) = 0: feries(10, 1) = "": feries(10, 2) = ""
    If Weekday(DateSerial(annee, 1, 1), vbMonday) = 7 Then
        feries(10, 0) = DateSerial(annee, 1, 2)
        feries(10, 1) = "2 Janvier (report Nouvel An)"
        feries(10, 2) = TYPE_FERIE_CALCULE
    End If

    ' Reports de Noël si tombe un dimanche
    feries(11, 0) = 0: feries(11, 1) = "": feries(11, 2) = ""
    If Weekday(DateSerial(annee, 12, 25), vbMonday) = 7 Then
        feries(11, 0) = DateSerial(annee, 12, 26)
        feries(11, 1) = "26 Decembre (report Noel)"
        feries(11, 2) = TYPE_FERIE_CALCULE
    End If

    ' Écriture dans la feuille
    Dim k As Integer
    For k = 0 To 11
        If feries(k, 0) <> 0 And feries(k, 1) <> "" Then
            Dim nouvRow As ListRow
            Set nouvRow = tbl.ListRows.Add
            nouvRow.Range(1, CF_DATE).Value = feries(k, 0)
            nouvRow.Range(1, CF_DATE).NumberFormat = "dd.mm.yyyy"
            nouvRow.Range(1, CF_NOM).Value = feries(k, 1)
            nouvRow.Range(1, CF_TYPE).Value = feries(k, 2)
        End If
    Next k

    Call TrierFeries(ws)

    MsgBox "Fériés " & annee & " chargés avec succès !", vbInformation, "Planning v2.0"

End Sub

'==============================================================================
' SUB : AjouterFerieManuel
' Ajoute un jour férié exceptionnel (type Manuel).
' Vérifie qu'il n'existe pas déjà.
'==============================================================================
Sub AjouterFerieManuel(ByVal dateJour As Date, ByVal nomFerie As String)

    If estFerie(dateJour) Then
        MsgBox "Le " & Format(dateJour, "dd.mm.yyyy") & " est déjà un jour férié.", vbInformation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_FERIES)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(NOM_TBL_FERIES)
    Dim nouvRow As ListRow
    Set nouvRow = tbl.ListRows.Add
    nouvRow.Range(1, CF_DATE).Value = dateJour
    nouvRow.Range(1, CF_DATE).NumberFormat = "dd.mm.yyyy"
    nouvRow.Range(1, CF_NOM).Value = nomFerie
    nouvRow.Range(1, CF_TYPE).Value = TYPE_FERIE_MANUEL
    Call TrierFeries(ws)

End Sub

'==============================================================================
' SUB : SupprimerFerie
' Supprime un jour férié. Demande confirmation si férié automatique.
'==============================================================================
Sub SupprimerFerie(ByVal dateJour As Date)

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_FERIES)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(NOM_TBL_FERIES)

    Dim i As Long
    i = ws.Cells(ws.Rows.count, CF_DATE).End(xlUp).Row

    Do While i >= 2
        Dim d As Date
        d = 0
        On Error Resume Next
        d = ws.Cells(i, CF_DATE).Value
        On Error GoTo 0

        If d <> 0 Then
            If Int(d) = Int(dateJour) Then
                If ws.Cells(i, CF_TYPE).Value <> TYPE_FERIE_MANUEL Then
                    Dim rep As VbMsgBoxResult
                    rep = MsgBox("Ce férié est calculé automatiquement." & vbCrLf & _
                                 "Voulez-vous quand même le supprimer ?", _
                                 vbYesNo + vbQuestion, "Supprimer le férié")
                    If rep = vbNo Then Exit Sub
                End If
                ' Suppression via ListObject (convention du projet) au lieu
                ' de ws.Rows(i).Delete, qui efface toute la largeur de la ligne.
                tbl.ListRows(i - tbl.HeaderRowRange.Row).Delete
                Exit Sub
            End If
        End If

        i = i - 1
    Loop

End Sub

'==============================================================================
' FUNCTION : EstFerie
' Retourne True si la date est un jour férié ou une fermeture.
' Utilisée par GenererPlanning() et les calendriers.
'==============================================================================
Function estFerie(ByVal dateJour As Date) As Boolean

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(NOM_FEUILLE_FERIES)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, CF_DATE).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernLigne
        Dim d As Date
        d = 0
        On Error Resume Next
        d = ws.Cells(i, CF_DATE).Value
        On Error GoTo 0
        If d <> 0 And Int(d) = Int(dateJour) Then
            estFerie = True
            Exit Function
        End If
    Next i

    estFerie = False

End Function

'==============================================================================
' FUNCTION : EstFermeture
' Retourne True si la date est une fermeture d'entreprise.
'==============================================================================
Function EstFermeture(ByVal dateJour As Date) As Boolean

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(NOM_FEUILLE_FERIES)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim ligne As Variant
    ligne = Application.Match(CDbl(dateJour), ws.Columns(CF_DATE), 0)

    If IsError(ligne) Then
        EstFermeture = False
        Exit Function
    End If

    EstFermeture = (Trim(ws.Cells(ligne, CF_TYPE).Value) = TYPE_FERIE_FERMETURE)

End Function

'==============================================================================
' FUNCTION : GetNomFerie
' Retourne le nom du jour férié pour une date donnée.
' Retourne une chaîne vide si la date n'est pas fériée.
'==============================================================================
Function GetNomFerie(ByVal dateJour As Date) As String

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(NOM_FEUILLE_FERIES)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, CF_DATE).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernLigne
        Dim d As Date
        d = 0
        On Error Resume Next
        d = ws.Cells(i, CF_DATE).Value
        On Error GoTo 0
        If d <> 0 And Int(d) = Int(dateJour) Then
            GetNomFerie = ws.Cells(i, CF_NOM).Value
            Exit Function
        End If
    Next i

End Function

'==============================================================================
' SUB : AjouterFermeturePeriode
' Ajoute chaque jour d'une période comme fermeture d'entreprise.
' Si la date existe déjà, met à jour le type en "Fermeture".
'==============================================================================
Sub AjouterFermeturePeriode(ByVal dDebut As Date, _
                              ByVal dFin As Date, _
                              ByVal nom As String)

    If dFin < dDebut Then
        MsgBox "La date de fin doit être postérieure à la date de début.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_FERIES)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(NOM_TBL_FERIES)

    Dim dateCase As Date
    For dateCase = dDebut To dFin

        Dim ligne As Variant
        ligne = Application.Match(CDbl(dateCase), ws.Columns(CF_DATE), 0)

        If IsError(ligne) Then
            ' La date n'existe pas — crée une nouvelle ligne
            Dim nouvRow As ListRow
            Set nouvRow = tbl.ListRows.Add
            nouvRow.Range(1, CF_DATE).Value = dateCase
            nouvRow.Range(1, CF_DATE).NumberFormat = "dd.mm.yyyy"
            nouvRow.Range(1, CF_NOM).Value = nom
            nouvRow.Range(1, CF_TYPE).Value = TYPE_FERIE_FERMETURE

        Else
            ' La date existe — met à jour le type via le ListObject
            ' (convention du projet) au lieu d'écrire directement via
            ' ws.Cells() dans les cellules du tableau.
            Dim ligneTbl As Long
            ligneTbl = CLng(ligne) - tbl.HeaderRowRange.Row
            tbl.ListRows(ligneTbl).Range(1, CF_TYPE).Value = TYPE_FERIE_FERMETURE
            If Trim(tbl.ListRows(ligneTbl).Range(1, CF_NOM).Value) = "" Then
                tbl.ListRows(ligneTbl).Range(1, CF_NOM).Value = nom
            End If
        End If

    Next dateCase

    Call TrierFeries(ws)

End Sub

'==============================================================================
' FUNCTION : GenererMoisCalendrier
' Génère un tableau de 42 cases (6 semaines x 7 jours) pour un mois donné.
' Utilisée par les calendriers du UserForm.
'==============================================================================
Function GenererMoisCalendrier(ByVal annee As Long, _
                                ByVal mois As Long, _
                                ByVal dateRefG1 As Date) As JourCalendrier()

    Dim jours(41) As JourCalendrier

    Dim premierJour As Date
    premierJour = DateSerial(annee, mois, 1)

    ' Décalage pour que la grille commence un lundi
    Dim decalage As Long
    decalage = Weekday(premierJour, vbMonday) - 1

    Dim debutGrille As Date
    debutGrille = premierJour - decalage

    Dim i As Long
    For i = 0 To 41

        Dim dateCase As Date
        dateCase = debutGrille + i

        jours(i).dateJour = dateCase
        jours(i).estDuMois = (Month(dateCase) = mois And Year(dateCase) = annee)

        Dim jourSem As Long
        jourSem = Weekday(dateCase, vbMonday)

        jours(i).estWE = (jourSem = 6 Or jourSem = 7)
        jours(i).estFerie = estFerie(dateCase)
        jours(i).nomFerie = GetNomFerie(dateCase)

        ' Groupe G1/G2 : lit depuis la feuille Groupes en priorité.
        ' Une fermeture d'entreprise (congés collectifs) n'est PAS un jour
        ' nécessitant un roulement auxiliaire : personne ne travaille ce
        ' jour-là, ni fixe ni auxiliaire. Sans cette exclusion, une semaine
        ' de fermeture se voyait attribuer un groupe G1/G2 comme un vrai
        ' jour férié.
        Dim estJourFerieAux As Boolean
        estJourFerieAux = jours(i).estFerie And Not EstFermeture(dateCase)

        If jours(i).estWE Or estJourFerieAux Then
            Dim groupeFeuille As String
            groupeFeuille = Module_Groupes.GetGroupeJour(dateCase)
            If groupeFeuille <> "" Then
                jours(i).groupe = groupeFeuille
            Else
                Dim lundiSem As Date
                lundiSem = dateCase - (jourSem - 1)
                Dim nbSem As Long
                nbSem = Int((lundiSem - dateRefG1) / 7)
                jours(i).groupe = IIf(nbSem Mod 2 = 0, GROUPE_G1, GROUPE_G2)
            End If
        End If

        ' Couleur de fond — priorité : fermeture > hors mois > ferie > WE > normal
        If Not jours(i).estDuMois Then
            jours(i).couleur = RGB(210, 210, 210)
        ElseIf EstFermeture(dateCase) Then
            jours(i).couleur = RGB(255, 0, 0)       ' Rouge — fermeture entreprise
        ElseIf jours(i).estFerie And jours(i).estWE Then
            jours(i).couleur = RGB(255, 140, 0)     ' Orange foncé — ferie + WE
        ElseIf jours(i).estFerie Then
            jours(i).couleur = RGB(255, 200, 100)   ' Orange clair — ferie semaine
        ElseIf jours(i).estWE And jours(i).groupe = GROUPE_G1 Then
            jours(i).couleur = RGB(173, 216, 230)   ' Bleu clair — G1
        ElseIf jours(i).estWE And jours(i).groupe = GROUPE_G2 Then
            jours(i).couleur = RGB(144, 238, 144)   ' Vert clair — G2
        Else
            jours(i).couleur = RGB(255, 255, 255)   ' Blanc — jour normal
        End If

    Next i

    GenererMoisCalendrier = jours

End Function

'==============================================================================
' SUB : TrierFeries (privée)
' Trie la feuille Feries par date croissante.
'==============================================================================
Private Sub TrierFeries(ByVal ws As Worksheet)
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(NOM_TBL_FERIES)
    If tbl.ListRows.count < 2 Then Exit Sub
    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(CF_DATE).Range, _
                             Order:=xlAscending
    tbl.Sort.Apply
End Sub
