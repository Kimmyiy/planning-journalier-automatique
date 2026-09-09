Option Explicit

'==============================================================================
' MODULE : Module_UserForm
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'==============================================================================

'------------------------------------------------------------------------------
' CONSTANTES - Colonnes feuille Personnel
'------------------------------------------------------------------------------
Private Const C_ID        As Long = 1
Private Const C_NOM       As Long = 2
Private Const C_STATUT    As Long = 3
Private Const C_TYPE      As Long = 4
Private Const C_GROUPE    As Long = 5
Private Const C_LUN       As Long = 6
Private Const C_MAR       As Long = 7
Private Const C_MER       As Long = 8
Private Const C_JEU       As Long = 9
Private Const C_VEN       As Long = 10
Private Const C_SAM       As Long = 11
Private Const C_DIM       As Long = 12
Private Const C_P1        As Long = 13
Private Const C_P2        As Long = 14
Private Const C_P3        As Long = 15
Private Const C_D         As Long = 16
Private Const C_VAOX1     As Long = 17
Private Const C_STATS_A   As Long = 18
Private Const C_STATS_B   As Long = 19
Private Const C_STATS_C   As Long = 20
Private Const C_STATS_D   As Long = 21
Private Const C_STATS_V1  As Long = 22
Private Const C_STATS_EXP As Long = 23

'------------------------------------------------------------------------------
' CONSTANTES - Colonnes du tableau Tbl_Parametres
'------------------------------------------------------------------------------
Private Const P_MACHINE    As Long = 1  ' A - Nom machine
Private Const P_SEUIL1     As Long = 2  ' B - Seuil 2 personnes
Private Const P_SEUIL2     As Long = 3  ' C - Seuil 3 personnes
Private Const P_MAX        As Long = 4  ' D - Quantite max
Private Const P_FERMETURE  As Long = 5  ' E - Fermeture (VRAI/FAUX)

' Couleurs groupes auxiliaires (identiques au calendrier)
Private Const COULEUR_G1 As Long = 11393254  ' Bleu clair  RGB(173,216,230)
Private Const COULEUR_G2 As Long = 9498256   ' Vert clair  RGB(144,238,144)
Private Const COULEUR_BLANC As Long = 16777215

'==============================================================================
' Affiche un calendrier mensuel pour la personne selectionnee avec :
' FIXE :
'   Vert       = matin (M)
'   Jaune      = journee complete (J)
'   Gris clair = absent (-)
'   Bleu       = vacances
'   Violet     = maladie
'   Rose pale  = conge
'   Orange     = jour ferie
'   Encadre rouge = aujourd hui
'
' AUXILIAIRE :
'   Vert        = travaille (son groupe actif ce jour)
'   Vert pale   = remplace quelqu un (elle vient mais pas son groupe)
'   Rouge pale  = se fait remplacer (elle ne vient pas)
'   Gris clair  = jour off
'   Orange      = jour ferie
'   Encadre rouge = aujourd hui
'==============================================================================

'------------------------------------------------------------------------------
' CONSTANTES COULEURS - Calendrier personnel (valeurs Long BGR pour VBA)
'------------------------------------------------------------------------------
Private Const CAL_COULEUR_MATIN      As Long = 13434828  ' Vert clair    RGB(204,255,204) -> &H00CCFFCC
Private Const CAL_COULEUR_JOURNEE    As Long = 5025616   ' Vert          RGB(144,238,144) -> &H0090EE90
Private Const CAL_COULEUR_ABSENT     As Long = 14474460  ' Gris clair    RGB(220,220,220) -> &H00DCDCDC
Private Const CAL_COULEUR_VACANCES   As Long = 16764006  ' Bleu clair    RGB(102,204,255) -> &H0066CCFF
Private Const CAL_COULEUR_MALADIE    As Long = 16751052  ' Violet clair  RGB(204,153,255) -> &H00CC99FF
Private Const CAL_COULEUR_CONGE      As Long = 12695295  ' Rose pale     RGB(255,182,193) -> &H00FFB6C1
Private Const CAL_COULEUR_FERIE      As Long = 10079487  ' Orange clair  RGB(255,204,102) -> &H00FFCC66
Private Const CAL_COULEUR_AUX_WORK   As Long = 5025616   ' Vert          RGB(144,238,144) -> &H0090EE90
Private Const CAL_COULEUR_AUX_REMPL  As Long = 13434828  ' Vert pale     RGB(204,255,204) -> &H00CCFFCC
Private Const CAL_COULEUR_AUX_ABS    As Long = 13888254  ' Rouge pale    RGB(255,182,193) -> &H00FFB6C1
Private Const CAL_COULEUR_HORS_MOIS  As Long = 15921906  ' Gris tres clair
Private Const CAL_COULEUR_BLANC      As Long = 16777215  ' Blanc

'==============================================================================

' Couleurs Excel (Long)
Private Const IMP_ROUGE      As Long = 12611584  ' RGB(192,0,0)
Private Const IMP_GRIS_ENT   As Long = 14277081  ' RGB(217,217,217)
Private Const IMP_GRIS_LBL   As Long = 15921906  ' RGB(242,242,242)
Private Const IMP_VERT       As Long = 5287936   ' RGB(0,128,0)

'==============================================================================
' ONGLET PERSONNEL - Recherche et filtrage
' Tri fixe : Type (Fixe d abord) -> Alphabetique -> Groupe (G1 d abord)
' Filtres : texte + archives + Fixe + G1 + G2
'==============================================================================
Sub UF_RecherchePersonnel(ByVal frm As Object, ByVal texte As String)

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    frm.lstPersonnel.Clear

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, C_NOM).End(xlUp).Row

    Dim filtre As String
    filtre = LCase(Trim(texte))

    ' Lecture des filtres actifs
    Dim afficherArchives As Boolean
    Dim filtreFixe       As Boolean
    Dim filtreG1         As Boolean
    Dim filtreG2         As Boolean

    afficherArchives = frm.chkAfficherArchives.Value
    filtreFixe = frm.chkFiltreFixe.Value
    filtreG1 = frm.chkFiltreG1.Value
    filtreG2 = frm.chkFiltreG2.Value

    ' --- Collecte des lignes valides ---
    Dim lignes()  As Long
    Dim noms()    As String
    Dim types()   As String
    Dim groupes() As String
    Dim statuts() As String
    ReDim lignes(1 To dernLigne)
    ReDim noms(1 To dernLigne)
    ReDim types(1 To dernLigne)
    ReDim groupes(1 To dernLigne)
    ReDim statuts(1 To dernLigne)
    Dim count As Long
    count = 0

    Dim i As Long
    For i = 2 To dernLigne

        Dim nom    As String
        Dim statut As String
        Dim typ    As String
        Dim groupe As String

        nom = ws.Cells(i, C_NOM).Value
        statut = ws.Cells(i, C_STATUT).Value
        typ = ws.Cells(i, C_TYPE).Value
        groupe = ws.Cells(i, C_GROUPE).Value

        If statut = "" Then GoTo Suivant

        ' Filtre archives - CORRIGE : affiche archives si coche, masque si non coche
        If Not afficherArchives And statut = "Archive" Then GoTo Suivant

        ' Filtre type Fixe
        If filtreFixe And typ <> "Fixe" Then GoTo Suivant

        ' Filtre groupe G1 (uniquement si G2 n est pas aussi coche)
        If filtreG1 And Not filtreG2 Then
            If groupe <> "G1" Then GoTo Suivant
        End If

        ' Filtre groupe G2 (uniquement si G1 n est pas aussi coche)
        If filtreG2 And Not filtreG1 Then
            If groupe <> "G2" Then GoTo Suivant
        End If

        ' Filtre texte recherche
        If filtre <> "" Then
            If InStr(LCase(nom), filtre) = 0 Then GoTo Suivant
        End If

        count = count + 1
        lignes(count) = i
        noms(count) = nom
        types(count) = typ
        groupes(count) = groupe
        statuts(count) = statut

Suivant:
    Next i

    If count = 0 Then Exit Sub

    ' --- Tri : Type (Fixe=1 / Auxiliaire=2) -> Nom alpha -> Groupe (G1=1 / G2=2) ---
    Dim a As Long, b As Long
    For a = 1 To count - 1
        For b = 1 To count - a

            Dim ordreBType As Integer
            Dim ordreAType As Integer
            ordreBType = IIf(types(b) = "Fixe", 1, 2)
            ordreAType = IIf(types(b + 1) = "Fixe", 1, 2)

            Dim doitEchanger As Boolean
            doitEchanger = False

            If ordreBType > ordreAType Then
                doitEchanger = True
            ElseIf ordreBType = ordreAType Then
                If noms(b) > noms(b + 1) Then
                    doitEchanger = True
                ElseIf noms(b) = noms(b + 1) Then
                    Dim ordreGrpB As Integer
                    Dim ordreGrpA As Integer
                    ordreGrpB = IIf(groupes(b) = "G1", 1, IIf(groupes(b) = "G2", 2, 3))
                    ordreGrpA = IIf(groupes(b + 1) = "G1", 1, IIf(groupes(b + 1) = "G2", 2, 3))
                    If ordreGrpB > ordreGrpA Then doitEchanger = True
                End If
            End If

            If doitEchanger Then
                Dim tmpL As Long:   tmpL = lignes(b):  lignes(b) = lignes(b + 1):   lignes(b + 1) = tmpL
                Dim tmpN As String: tmpN = noms(b):    noms(b) = noms(b + 1):       noms(b + 1) = tmpN
                Dim tmpT As String: tmpT = types(b):   types(b) = types(b + 1):     types(b + 1) = tmpT
                Dim tmpG As String: tmpG = groupes(b): groupes(b) = groupes(b + 1): groupes(b + 1) = tmpG
                Dim tmpS As String: tmpS = statuts(b): statuts(b) = statuts(b + 1): statuts(b + 1) = tmpS
            End If

        Next b
    Next a

    ' --- Ecriture dans la ListBox avec couleurs ---
    For i = 1 To count

        ' Prefixe [G1]/[G2] pour les auxiliaires
        Dim affichageNom As String
        If types(i) = "Auxiliaire" And groupes(i) <> "-" And groupes(i) <> "" Then
            affichageNom = "[" & groupes(i) & "] " & noms(i)
        Else
            affichageNom = noms(i)
        End If

        frm.lstPersonnel.AddItem affichageNom
        frm.lstPersonnel.List(frm.lstPersonnel.ListCount - 1, 1) = statuts(i)
        frm.lstPersonnel.List(frm.lstPersonnel.ListCount - 1, 2) = CStr(lignes(i))
    Next i

End Sub

'==============================================================================
' ONGLET PERSONNEL - Appliquer les couleurs G1/G2 dans la ListBox
' A appeler apres UF_RecherchePersonnel depuis le UserForm
' Note : les ListBox VBA ne supportent pas les couleurs par ligne nativement.
' On utilise une astuce via les Labels de couleur en superposition.
' Solution simple : on code la couleur dans la colonne cachee index 3.
'==============================================================================
Sub UF_AppliquerCouleursListe(ByVal frm As Object)

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim i As Long
    For i = 0 To frm.lstPersonnel.ListCount - 1

        Dim ligneStr As String
        ligneStr = frm.lstPersonnel.List(i, 2)
        If ligneStr = "" Then GoTo SuivantCouleur

        Dim ligne As Long
        ligne = CLng(ligneStr)

        Dim typ    As String
        Dim groupe As String
        typ = ws.Cells(ligne, C_TYPE).Value
        groupe = ws.Cells(ligne, C_GROUPE).Value

        ' Stocke le code couleur en colonne 3 (cachee)
        If typ = "Auxiliaire" Then
            If groupe = "G1" Then
                frm.lstPersonnel.List(i, 3) = "G1"
            ElseIf groupe = "G2" Then
                frm.lstPersonnel.List(i, 3) = "G2"
            End If
        End If

SuivantCouleur:
    Next i

End Sub

'==============================================================================
' ONGLET PERSONNEL - Charger une fiche dans le formulaire
'==============================================================================
Sub UF_ChargerFiche(ByVal frm As Object)

    If frm.lstPersonnel.ListIndex = -1 Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim ligneStr As String
    ligneStr = frm.lstPersonnel.List(frm.lstPersonnel.ListIndex, 2)

    If ligneStr = "" Then Exit Sub

    Dim ligne As Long
    ligne = CLng(ligneStr)

    frm.txtID.Value = ws.Cells(ligne, C_ID).Value
    frm.txtNom.Value = ws.Cells(ligne, C_NOM).Value

    frm.cboStatut.Value = ws.Cells(ligne, C_STATUT).Value
    frm.cboType.Value = ws.Cells(ligne, C_TYPE).Value
    frm.cboGroupe.Value = ws.Cells(ligne, C_GROUPE).Value

    frm.cboLun.Value = ws.Cells(ligne, C_LUN).Value
    frm.cboMar.Value = ws.Cells(ligne, C_MAR).Value
    frm.cboMer.Value = ws.Cells(ligne, C_MER).Value
    frm.cboJeu.Value = ws.Cells(ligne, C_JEU).Value
    frm.cboVen.Value = ws.Cells(ligne, C_VEN).Value
    frm.cboSam.Value = ws.Cells(ligne, C_SAM).Value
    frm.cboDim.Value = ws.Cells(ligne, C_DIM).Value

    frm.chkP1.Value = (ws.Cells(ligne, C_P1).Value = 1)
    frm.chkP2.Value = (ws.Cells(ligne, C_P2).Value = 1)
    frm.chkP3.Value = (ws.Cells(ligne, C_P3).Value = 1)
    frm.chkD.Value = (ws.Cells(ligne, C_D).Value = 1)
    frm.chkVaox1.Value = (ws.Cells(ligne, C_VAOX1).Value = 1)

    frm.lblStatsA.Caption = "A : " & ws.Cells(ligne, C_STATS_A).Value
    frm.lblStatsB.Caption = "B : " & ws.Cells(ligne, C_STATS_B).Value
    frm.lblStatsC.Caption = "C : " & ws.Cells(ligne, C_STATS_C).Value
    frm.lblStatsD.Caption = "D : " & ws.Cells(ligne, C_STATS_D).Value
    frm.lblStatsV1.Caption = "Vaox1 : " & ws.Cells(ligne, C_STATS_V1).Value
    frm.lblStatsEXP.Caption = "EXP : " & ws.Cells(ligne, C_STATS_EXP).Value

    frm.txtLigneActive.Value = CStr(ligne)

End Sub

'==============================================================================
' ONGLET PERSONNEL - Enregistrer une fiche
'==============================================================================
Function UF_EnregistrerFiche(ByVal frm As Object) As Boolean

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_PERSONNEL)

    If Trim(frm.txtNom.Value) = "" Then
        MsgBox "Le nom ne peut pas être vide.", vbExclamation, "Validation"
        UF_EnregistrerFiche = False
        Exit Function
    End If

    Dim ligne      As Long
    Dim estNouvelle As Boolean

    If frm.txtLigneActive.Value = "" Or frm.txtLigneActive.Value = "0" Then
        estNouvelle = True
        ' Ajout via ListObject.ListRows.Add (convention du projet) au lieu
        ' d'écrire directement sous la dernière ligne via ws.Cells() : cela
        ' garantit que la nouvelle fiche est bien intégrée à Tbl_Personnel,
        ' visible pour ChargerPersonnes() lors de la génération du planning.
        Dim tbl As ListObject
        Set tbl = ws.ListObjects(NOM_TBL_PERSONNEL)
        Dim nouvRow As ListRow
        Set nouvRow = tbl.ListRows.Add
        ligne = nouvRow.Range.Row
        ws.Cells(ligne, C_ID).Value = GenererID(ws)
    Else
        ligne = CLng(frm.txtLigneActive.Value)
    End If

    ws.Cells(ligne, C_NOM).Value = Trim(frm.txtNom.Value)
    ws.Cells(ligne, C_STATUT).Value = frm.cboStatut.Value
    ws.Cells(ligne, C_TYPE).Value = frm.cboType.Value
    ws.Cells(ligne, C_GROUPE).Value = frm.cboGroupe.Value
    ws.Cells(ligne, C_LUN).Value = frm.cboLun.Value
    ws.Cells(ligne, C_MAR).Value = frm.cboMar.Value
    ws.Cells(ligne, C_MER).Value = frm.cboMer.Value
    ws.Cells(ligne, C_JEU).Value = frm.cboJeu.Value
    ws.Cells(ligne, C_VEN).Value = frm.cboVen.Value
    ws.Cells(ligne, C_SAM).Value = frm.cboSam.Value
    ws.Cells(ligne, C_DIM).Value = frm.cboDim.Value
    ws.Cells(ligne, C_P1).Value = IIf(frm.chkP1.Value, 1, 0)
    ws.Cells(ligne, C_P2).Value = IIf(frm.chkP2.Value, 1, 0)
    ws.Cells(ligne, C_P3).Value = IIf(frm.chkP3.Value, 1, 0)
    ws.Cells(ligne, C_D).Value = IIf(frm.chkD.Value, 1, 0)
    ws.Cells(ligne, C_VAOX1).Value = IIf(frm.chkVaox1.Value, 1, 0)

    If estNouvelle Then
        ws.Cells(ligne, C_STATS_A).Value = 0
        ws.Cells(ligne, C_STATS_B).Value = 0
        ws.Cells(ligne, C_STATS_C).Value = 0
        ws.Cells(ligne, C_STATS_D).Value = 0
        ws.Cells(ligne, C_STATS_V1).Value = 0
        ws.Cells(ligne, C_STATS_EXP).Value = 0
        MsgBox "Nouvelle fiche créée : " & Trim(frm.txtNom.Value), vbInformation, "Personnel v2.0"
    Else
        MsgBox "Fiche mise à jour : " & Trim(frm.txtNom.Value), vbInformation, "Personnel v2.0"
    End If

    UF_EnregistrerFiche = True

End Function

'==============================================================================
' ONGLET PERSONNEL - Archiver une fiche
'==============================================================================
Sub UF_ArchiverFiche(ByVal frm As Object)

    If frm.txtLigneActive.Value = "" Or frm.txtLigneActive.Value = "0" Then
        MsgBox "Aucune fiche sélectionnée.", vbExclamation, "Archive"
        Exit Sub
    End If

    Dim nom As String
    nom = frm.txtNom.Value

    Dim rep As VbMsgBoxResult
    rep = MsgBox("Archiver la fiche de " & nom & " ?" & vbCrLf & _
                 "La personne ne sera plus incluse dans les plannings.", _
                 vbYesNo + vbQuestion, "Confirmer l'archivage")

    If rep = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim ligne As Long
    ligne = CLng(frm.txtLigneActive.Value)

    ws.Cells(ligne, C_STATUT).Value = "Archive"

    MsgBox nom & " a été archivé(e).", vbInformation, "Personnel v2.0"

End Sub

'==============================================================================
' ONGLET PERSONNEL - Supprimer les statistiques d une fiche
'==============================================================================
Sub UF_SupprimerStats(ByVal frm As Object)

    If frm.txtLigneActive.Value = "" Or frm.txtLigneActive.Value = "0" Then
        MsgBox "Aucune fiche sélectionnée.", vbExclamation, "Stats"
        Exit Sub
    End If

    Dim nom As String
    nom = frm.txtNom.Value

    Dim rep As VbMsgBoxResult
    rep = MsgBox("Remettre à zéro les statistiques de " & nom & " ?" & vbCrLf & _
                 "Cette action est irreversible.", _
                 vbYesNo + vbExclamation, "Confirmer la suppression des stats")

    If rep = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim ligne As Long
    ligne = CLng(frm.txtLigneActive.Value)

    ws.Cells(ligne, C_STATS_A).Value = 0
    ws.Cells(ligne, C_STATS_B).Value = 0
    ws.Cells(ligne, C_STATS_C).Value = 0
    ws.Cells(ligne, C_STATS_D).Value = 0
    ws.Cells(ligne, C_STATS_V1).Value = 0
    ws.Cells(ligne, C_STATS_EXP).Value = 0

    ' Rafraichit l affichage des stats
    frm.lblStatsA.Caption = "A : 0"
    frm.lblStatsB.Caption = "B : 0"
    frm.lblStatsC.Caption = "C : 0"
    frm.lblStatsD.Caption = "D : 0"
    frm.lblStatsV1.Caption = "Vaox1 : 0"
    frm.lblStatsEXP.Caption = "EXP : 0"

    MsgBox "Statistiques remises à zéro pour " & nom & ".", vbInformation, "Personnel v2.0"

End Sub

'==============================================================================
' ONGLET PERSONNEL - Vider le formulaire pour nouvelle fiche
'==============================================================================
Sub UF_NouvellesFiche(ByVal frm As Object)

    frm.txtID.Value = "(auto)"
    frm.txtNom.Value = ""
    frm.txtLigneActive.Value = "0"
    frm.cboStatut.Value = "Actif"
    frm.cboType.Value = "Fixe"
    frm.cboGroupe.Value = "-"

    frm.cboLun.Value = "M"
    frm.cboMar.Value = "M"
    frm.cboMer.Value = "M"
    frm.cboJeu.Value = "M"
    frm.cboVen.Value = "M"
    frm.cboSam.Value = "-"
    frm.cboDim.Value = "-"

    frm.chkP1.Value = False
    frm.chkP2.Value = False
    frm.chkP3.Value = False
    frm.chkD.Value = False
    frm.chkVaox1.Value = False

    frm.lblStatsA.Caption = "A : 0"
    frm.lblStatsB.Caption = "B : 0"
    frm.lblStatsC.Caption = "C : 0"
    frm.lblStatsD.Caption = "D : 0"
    frm.lblStatsV1.Caption = "Vaox1 : 0"
    frm.lblStatsEXP.Caption = "EXP : 0"

    frm.txtNom.SetFocus

End Sub

'==============================================================================
' UTILITAIRE - Genere un ID unique (OP001, OP002...)
'==============================================================================
Private Function GenererID(ByVal ws As Worksheet) As String

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, C_ID).End(xlUp).Row

    Dim maxNum As Long
    maxNum = 0

    Dim i As Long
    For i = 2 To dernLigne
        Dim id As String
        id = CStr(ws.Cells(i, C_ID).Value)
        If Left(id, 2) = "OP" Then
            Dim num As Long
            num = Val(Mid(id, 3))
            If num > maxNum Then maxNum = num
        End If
    Next i

    GenererID = "OP" & Format(maxNum + 1, "000")

End Function

'==============================================================================
' UTILITAIRE - Trouve l ID d une personne par son nom
'==============================================================================
Private Function TrouverID(ByVal ws As Worksheet, ByVal nom As String) As String

    Dim ligne As Variant
    ligne = Application.Match(nom, ws.Columns(C_NOM), 0)

    If IsError(ligne) Then
        TrouverID = ""
    Else
        TrouverID = ws.Cells(ligne, C_ID).Value
    End If

End Function

'==============================================================================
' ONGLET AUXILIAIRES - Charger la liste des weekends
' Affiche les 8 prochains weekends avec groupe
'==============================================================================
Sub UF_ChargerCalendrierWeekend(ByVal frm As Object)

    frm.lstCalendrierWE.Clear
    frm.lstRemplacementsWE.Clear

    Dim wsPl  As Worksheet
    Dim wsRpl As Worksheet
    Set wsPl = Sheets("Planning journalier")
    Set wsRpl = Sheets("Remplacements")

    Dim dateRefG1 As Date
    On Error Resume Next
    dateRefG1 = wsPl.Range("Ref_GroupeG1").Value
    On Error GoTo 0

    If dateRefG1 = 0 Then
        frm.lstCalendrierWE.AddItem "Ref_GroupeG1 non configuree"
        Exit Sub
    End If

    ' Trouve le prochain samedi
    Dim dateDebut As Date
    dateDebut = Date
    Do While Weekday(dateDebut, vbMonday) <> 6
        dateDebut = dateDebut + 1
    Loop

    Dim we As Integer
    For we = 0 To 7  ' 8 prochains weekends

        Dim dateSam As Date
        Dim dateDim As Date
        dateSam = dateDebut + (we * 7)
        dateDim = dateSam + 1

        ' Calcule le groupe actif
        Dim groupe As String
        groupe = Module_Groupes.GetGroupeJour(dateSam)
        If groupe = "" Then
            Dim nbSem As Long
            nbSem = Int((dateSam - dateRefG1) / 7)
            groupe = IIf(nbSem Mod 2 = 0, "G1", "G2")
        End If

        ' Verifie si des remplacements existent ce weekend
        Dim aDesRpl As Boolean
        aDesRpl = False
        Dim dernLigneRpl As Long
        dernLigneRpl = wsRpl.Cells(wsRpl.Rows.count, 1).End(xlUp).Row
        Dim j As Long
        For j = 2 To dernLigneRpl
            Dim dateRpl As Date
            On Error Resume Next
            dateRpl = wsRpl.Cells(j, 1).Value
            On Error GoTo 0
            If Int(dateRpl) = Int(dateSam) Or Int(dateRpl) = Int(dateDim) Then
                aDesRpl = True
                Exit For
            End If
        Next j

        ' Indicateur de remplacement
        Dim indicateur As String
        indicateur = IIf(aDesRpl, " (*)", "")

        Dim ligneCal As String
        ligneCal = Format(dateSam, "dd.mm") & " / " & Format(dateDim, "dd.mm") & _
                   "   [" & groupe & "]" & indicateur

        frm.lstCalendrierWE.AddItem ligneCal
        frm.lstCalendrierWE.List(frm.lstCalendrierWE.ListCount - 1, 1) = CStr(CDbl(dateSam))
        frm.lstCalendrierWE.List(frm.lstCalendrierWE.ListCount - 1, 2) = CStr(CDbl(dateDim))

    Next we

End Sub

'==============================================================================
' ONGLET AUXILIAIRES - Charger les remplacements d un weekend selectionne
' Appelee quand on clique sur un WE dans lstCalendrierWE
'==============================================================================
Sub UF_ChargerRemplacementsWE(ByVal frm As Object)

    frm.lstRemplacementsWE.Clear

    If frm.lstCalendrierWE.ListIndex = -1 Then Exit Sub

    Dim dateSamStr As String
    Dim dateDimStr As String
    dateSamStr = frm.lstCalendrierWE.List(frm.lstCalendrierWE.ListIndex, 1)
    dateDimStr = frm.lstCalendrierWE.List(frm.lstCalendrierWE.ListIndex, 2)

    If dateSamStr = "" Then Exit Sub

    Dim dateSam As Date
    Dim dateDim As Date
    dateSam = CDate(CDbl(dateSamStr))
    dateDim = CDate(CDbl(dateDimStr))

    Dim wsRpl As Worksheet
    Set wsRpl = Sheets("Remplacements")

    Dim dernLigne As Long
    dernLigne = wsRpl.Cells(wsRpl.Rows.count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernLigne
        Dim dateRpl As Date
        On Error Resume Next
        dateRpl = wsRpl.Cells(i, 1).Value
        On Error GoTo 0

        If Int(dateRpl) = Int(dateSam) Or Int(dateRpl) = Int(dateDim) Then
            Dim absent As String
            Dim rempl  As String
            absent = wsRpl.Cells(i, 3).Value
            rempl = wsRpl.Cells(i, 5).Value

            Dim ligneRpl As String
            ligneRpl = Format(dateRpl, "dd.mm") & "  |  " & absent & " -> " & rempl

            frm.lstRemplacementsWE.AddItem ligneRpl
            frm.lstRemplacementsWE.List(frm.lstRemplacementsWE.ListCount - 1, 1) = CStr(i)
        End If
    Next i

End Sub

'==============================================================================
' ONGLET AUXILIAIRES - Charger les auxiliaires disponibles
'==============================================================================
Sub UF_ChargerAuxiliaires(ByVal frm As Object, ByVal groupeActif As String)

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    frm.cboAuxAbsente.Clear
    frm.cboAuxRemplacante.Clear

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, C_NOM).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernLigne

        If ws.Cells(i, C_STATUT).Value <> "Actif" Then GoTo SuivantAux
        If ws.Cells(i, C_TYPE).Value <> "Auxiliaire" Then GoTo SuivantAux

        Dim nom As String
        nom = ws.Cells(i, C_NOM).Value

        If ws.Cells(i, C_GROUPE).Value = groupeActif Then
            frm.cboAuxAbsente.AddItem nom
        End If

        frm.cboAuxRemplacante.AddItem nom

SuivantAux:
    Next i

End Sub

'==============================================================================
' MODIFICATION : UF_EnregistrerRemplacement
' Remplace la version existante dans Module_UserForm.bas
' Ajoute le champ type (Remplacement / Renfort)
' Le UserForm doit avoir une ComboBox cboTypeRpl
'==============================================================================
Sub UF_EnregistrerRemplacement(ByVal frm As Object)

    If frm.dtpDateRemplacement.Value = "" Then
        MsgBox "Veuillez sélectionner une date.", vbExclamation, "Validation"
        Exit Sub
    End If

    If Not IsDate(frm.dtpDateRemplacement.Value) Then
        MsgBox "Date invalide. Format attendu : jj.mm.aaaa", vbExclamation, "Validation"
        Exit Sub
    End If

    ' Pour un renfort, la personne absente est optionnelle
    Dim typeRpl As String
    typeRpl = frm.cboTypeRpl.Value

    If typeRpl = "" Then
        MsgBox "Veuillez sélectionner le type (Remplacement ou Renfort).", vbExclamation, "Validation"
        Exit Sub
    End If

    If typeRpl = "Remplacement" And frm.cboAuxAbsente.Value = "" Then
        MsgBox "Veuillez sélectionner la personne absente.", vbExclamation, "Validation"
        Exit Sub
    End If

    If frm.cboAuxRemplacante.Value = "" Then
        Dim libelle As String
        libelle = IIf(typeRpl = "Renfort", "la personne en renfort", "la personne remplacante")
        MsgBox "Veuillez sélectionner " & libelle & ".", vbExclamation, "Validation"
        Exit Sub
    End If

    If typeRpl = "Remplacement" And _
       frm.cboAuxAbsente.Value = frm.cboAuxRemplacante.Value Then
        MsgBox "La personne absente et la remplacante ne peuvent pas être identiques.", _
               vbExclamation, "Validation"
        Exit Sub
    End If

    Dim dateRpl    As Date
    Dim nomAbsente As String
    Dim nomRempl   As String

    dateRpl = CDate(frm.dtpDateRemplacement.Value)
    nomAbsente = frm.cboAuxAbsente.Value   ' Vide si Renfort
    nomRempl = frm.cboAuxRemplacante.Value

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_PERSONNEL)

    Dim idAbsente As String
    Dim idRempl   As String
    If nomAbsente <> "" Then idAbsente = TrouverID(ws, nomAbsente)
    idRempl = TrouverID(ws, nomRempl)

    ' Enregistre dans Tbl_Remplacements
    Dim wsRpl As Worksheet
    Set wsRpl = Sheets(NOM_FEUILLE_REMPLACEMENTS)

    Dim tbl As ListObject
    Set tbl = wsRpl.ListObjects(NOM_TBL_REMPLACEMENTS)
    Dim nouvRow As ListRow
    Set nouvRow = tbl.ListRows.Add

    nouvRow.Range(1, RPL_COL_DATE).Value = dateRpl
    nouvRow.Range(1, RPL_COL_DATE).NumberFormat = "dd.mm.yyyy"
    nouvRow.Range(1, RPL_COL_ID_ABSENTE).Value = idAbsente
    nouvRow.Range(1, RPL_COL_NOM_ABSENTE).Value = nomAbsente
    nouvRow.Range(1, RPL_COL_ID_REMPLACANT).Value = idRempl
    nouvRow.Range(1, RPL_COL_NOM_REMPLACANT).Value = nomRempl
    nouvRow.Range(1, RPL_COL_TYPE).Value = typeRpl

    ' Message de confirmation
    Dim msgConf As String
    If typeRpl = "Renfort" Then
        msgConf = "Renfort enregistre :" & vbCrLf & _
                  nomRempl & " en renfort" & vbCrLf & _
                  "Le " & Format(dateRpl, "dd.mm.yyyy")
    Else
        msgConf = "Remplacement enregistre :" & vbCrLf & _
                  nomRempl & " remplace " & nomAbsente & vbCrLf & _
                  "Le " & Format(dateRpl, "dd.mm.yyyy")
    End If

    MsgBox msgConf, vbInformation, "Auxiliaires v2.0"

    ' Rafraîchit les listes
    Call UF_ChargerCalendrierWeekend(frm)
    Call UF_ChargerRemplacementsWE(frm)

End Sub


'==============================================================================
' ONGLET AUXILIAIRES - Supprimer un remplacement individuel selectionne
' dans lstRemplacementsWE
'==============================================================================
Sub UF_SupprimerRemplacementIndividuel(ByVal frm As Object)

    If frm.lstRemplacementsWE.ListIndex = -1 Then
        MsgBox "Sélectionnez d'abord un remplacement dans la liste.", vbExclamation, "Suppression"
        Exit Sub
    End If

    Dim ligneStr As String
    ligneStr = frm.lstRemplacementsWE.List(frm.lstRemplacementsWE.ListIndex, 1)

    If ligneStr = "" Then Exit Sub

    Dim rep As VbMsgBoxResult
    rep = MsgBox("Supprimer ce remplacement ?", vbYesNo + vbQuestion, "Confirmer")

    If rep = vbNo Then Exit Sub

    Dim wsRpl As Worksheet
    Set wsRpl = Sheets(NOM_FEUILLE_REMPLACEMENTS)

    ' Suppression via ListObject (convention du projet) au lieu de
    ' wsRpl.Rows(...).Delete.
    Dim tbl As ListObject
    Set tbl = wsRpl.ListObjects(NOM_TBL_REMPLACEMENTS)
    tbl.ListRows(CLng(ligneStr) - tbl.HeaderRowRange.Row).Delete

    MsgBox "Remplacement supprimé.", vbInformation, "Auxiliaires v2.0"

    Call UF_ChargerCalendrierWeekend(frm)
    Call UF_ChargerRemplacementsWE(frm)

End Sub

'==============================================================================
' ONGLET AUXILIAIRES - Ajouter tout un groupe en renfort pour une semaine
' Utilisé quand l'entreprise prend tout un groupe d'auxiliaires (étudiants)
' en semaine plutôt qu'un remplacement individuel. Enregistre un renfort
' (Tbl_Remplacements, Type="Renfort") pour chaque auxiliaire actif du groupe,
' sur les 5 jours ouvrés (lundi-vendredi) de la semaine choisie.
'==============================================================================
Sub UF_AjouterGroupeSemaine(ByVal frm As Object)

    Dim dlg As UserForm_GroupeSemaine
    Set dlg = New UserForm_GroupeSemaine

    ' Propose par défaut le lundi de la semaine en cours
    dlg.DateProposee = Date - (Weekday(Date, vbMonday) - 1)
    dlg.Initialiser
    dlg.Show

    If Not dlg.Confirme Then
        Unload dlg
        Exit Sub
    End If

    Dim groupe   As String
    Dim dLundi   As Date
    groupe = dlg.cboGroupeSemaine.Value
    dLundi = CDate(dlg.txtDebutSemaine.Value)
    Unload dlg

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_PERSONNEL)

    Dim wsRpl As Worksheet
    Set wsRpl = Sheets(NOM_FEUILLE_REMPLACEMENTS)

    Dim tbl As ListObject
    Set tbl = wsRpl.ListObjects(NOM_TBL_REMPLACEMENTS)

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, C_NOM).End(xlUp).Row

    Dim nbAjoutes As Long
    Dim nbDejaPresents As Long
    nbAjoutes = 0
    nbDejaPresents = 0

    Dim i As Long
    For i = 2 To dernLigne

        If ws.Cells(i, C_STATUT).Value <> "Actif" Then GoTo SuivantePersonne
        If ws.Cells(i, C_TYPE).Value <> "Auxiliaire" Then GoTo SuivantePersonne
        If ws.Cells(i, C_GROUPE).Value <> groupe Then GoTo SuivantePersonne

        Dim idPers  As String
        Dim nomPers As String
        idPers = ws.Cells(i, C_ID).Value
        nomPers = ws.Cells(i, C_NOM).Value

        Dim jourOffset As Long
        For jourOffset = 0 To 4   ' Lundi a vendredi
            Dim dateJourSemaine As Date
            dateJourSemaine = dLundi + jourOffset

            If ExisteRenfortPourJour(tbl, nomPers, dateJourSemaine) Then
                nbDejaPresents = nbDejaPresents + 1
            Else
                Dim nouvRow As ListRow
                Set nouvRow = tbl.ListRows.Add
                nouvRow.Range(1, RPL_COL_DATE).Value = dateJourSemaine
                nouvRow.Range(1, RPL_COL_DATE).NumberFormat = "dd.mm.yyyy"
                nouvRow.Range(1, RPL_COL_ID_REMPLACANT).Value = idPers
                nouvRow.Range(1, RPL_COL_NOM_REMPLACANT).Value = nomPers
                nouvRow.Range(1, RPL_COL_TYPE).Value = "Renfort"
                nbAjoutes = nbAjoutes + 1
            End If
        Next jourOffset

SuivantePersonne:
    Next i

    Dim msgResultat As String
    msgResultat = "Groupe " & groupe & " ajouté en renfort du " & _
                  Format(dLundi, "dd.mm.yyyy") & " au " & Format(dLundi + 4, "dd.mm.yyyy") & _
                  " :" & vbCrLf & nbAjoutes & " renfort(s) enregistré(s)."
    If nbDejaPresents > 0 Then
        msgResultat = msgResultat & vbCrLf & nbDejaPresents & " déjà enregistré(s), non dupliqué(s)."
    End If
    MsgBox msgResultat, vbInformation, "Renfort de groupe"

    Call UF_ChargerCalendrierWeekend(frm)

End Sub

'==============================================================================
' FUNCTION : ExisteRenfortPourJour (privée)
' Retourne True si un renfort est déjà enregistré pour cette personne à
' cette date, pour éviter les doublons lors de l'ajout d'un groupe entier.
'==============================================================================
Private Function ExisteRenfortPourJour(ByVal tbl As ListObject, ByVal nom As String, _
                                        ByVal dateJour As Date) As Boolean

    If tbl.ListRows.count = 0 Then Exit Function

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        Dim dRpl As Date
        dRpl = 0
        On Error Resume Next
        dRpl = tbl.DataBodyRange(i, RPL_COL_DATE).Value
        On Error GoTo 0
        If dRpl <> 0 And Int(dRpl) = Int(dateJour) Then
            If Trim(tbl.DataBodyRange(i, RPL_COL_NOM_REMPLACANT).Value) = nom Then
                ExisteRenfortPourJour = True
                Exit Function
            End If
        End If
    Next i

End Function

'==============================================================================
' ONGLET ABSENCES - Charger la liste du personnel fixe
'==============================================================================
Sub UF_ChargerPersonnelAbsences(ByVal frm As Object)

    frm.lstAbsPersonnel.Clear
    frm.lstAbsListe.Clear

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, C_NOM).End(xlUp).Row

    ' Toute personne active (Fixe ou Auxiliaire) peut être déclarée absente :
    ' un auxiliaire absent n'a plus besoin d'être obligatoirement remplacé.
    Dim i As Long
    For i = 2 To dernLigne
        If ws.Cells(i, C_STATUT).Value = "Actif" Then
            frm.lstAbsPersonnel.AddItem ws.Cells(i, C_NOM).Value
        End If
    Next i

End Sub

'==============================================================================
' ONGLET ABSENCES - Charger les absences d une personne selectionnee
'==============================================================================
Sub UF_ChargerAbsencesPersonne(ByVal frm As Object)

    frm.lstAbsListe.Clear

    If frm.lstAbsPersonnel.ListIndex = -1 Then Exit Sub

    Dim nomSelectionne As String
    nomSelectionne = frm.lstAbsPersonnel.Value

    Dim wsV As Worksheet
    Set wsV = Sheets("Vacances")

    Dim dernLigne As Long
    dernLigne = wsV.Cells(wsV.Rows.count, 2).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernLigne

        If Trim(wsV.Cells(i, 2).Value) = nomSelectionne Then

            Dim dDebut     As Date
            Dim dFin       As Date
            Dim typeAbs    As String
            Dim periodeAbs As String

            On Error Resume Next
            dDebut = wsV.Cells(i, 3).Value
            dFin = wsV.Cells(i, 4).Value
            On Error GoTo 0
            typeAbs = wsV.Cells(i, 5).Value
            periodeAbs = Trim(wsV.Cells(i, VAC_COL_PERIODE).Value)
            If periodeAbs = "" Then periodeAbs = PERIODE_JOURNEE

            Dim ligneAbs As String
            ligneAbs = typeAbs & " (" & periodeAbs & ") : " & _
                       Format(dDebut, "dd.mm.yyyy") & " -> " & Format(dFin, "dd.mm.yyyy")

            frm.lstAbsListe.AddItem ligneAbs
            frm.lstAbsListe.List(frm.lstAbsListe.ListCount - 1, 1) = CStr(i)

        End If

    Next i

    ' Met a jour la zone Remarques
    Call UF_MettreAJourRemarques(frm)

End Sub

'==============================================================================
' ONGLET ABSENCES - Enregistrer une absence
'==============================================================================
Sub UF_EnregistrerAbsence(ByVal frm As Object)

    If frm.lstAbsPersonnel.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner une personne.", vbExclamation, "Validation"
        Exit Sub
    End If

    If frm.dtpVacDebut.Value = "" Or frm.dtpVacFin.Value = "" Then
        MsgBox "Veuillez saisir les dates de début et de fin.", vbExclamation, "Validation"
        Exit Sub
    End If

    If Not IsDate(frm.dtpVacDebut.Value) Or Not IsDate(frm.dtpVacFin.Value) Then
        MsgBox "Date invalide. Format attendu : jj.mm.aaaa", vbExclamation, "Validation"
        Exit Sub
    End If

    Dim dDebut As Date
    Dim dFin   As Date
    dDebut = CDate(frm.dtpVacDebut.Value)
    dFin = CDate(frm.dtpVacFin.Value)

    If dFin < dDebut Then
        MsgBox "La date de fin ne peut pas être antérieure à la date de début.", _
               vbExclamation, "Validation"
        Exit Sub
    End If

    Dim nom     As String
    Dim typeAbs As String
    Dim periodeAbs As String
    nom = frm.lstAbsPersonnel.Value
    typeAbs = frm.cboVacType.Value
    periodeAbs = frm.cboVacPeriode.Value
    If periodeAbs = "" Then periodeAbs = PERIODE_JOURNEE

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_PERSONNEL)
    Dim idPerso As String
    idPerso = TrouverID(ws, nom)

    Dim wsV As Worksheet
    Set wsV = Sheets(NOM_FEUILLE_VACANCES)

    If ExisteAbsenceChevauchante(wsV, nom, dDebut, dFin) Then
        Dim repChevauchement As VbMsgBoxResult
        repChevauchement = MsgBox("Une absence existe déjà pour " & nom & _
               " sur une période qui chevauche ces dates." & vbCrLf & _
               "Enregistrer quand même ?", vbYesNo + vbExclamation, "Chevauchement détecté")
        If repChevauchement = vbNo Then Exit Sub
    End If

    ' Ajout via ListObject.ListRows.Add (convention du projet) au lieu
    ' d'écrire directement sous la dernière ligne via ws.Cells().
    Dim tbl As ListObject
    Set tbl = wsV.ListObjects(NOM_TBL_VACANCES)
    Dim nouvRow As ListRow
    Set nouvRow = tbl.ListRows.Add

    nouvRow.Range(1, VAC_COL_ID).Value = idPerso
    nouvRow.Range(1, VAC_COL_NOM).Value = nom
    nouvRow.Range(1, VAC_COL_DEBUT).Value = dDebut
    nouvRow.Range(1, VAC_COL_FIN).Value = dFin
    nouvRow.Range(1, VAC_COL_TYPE).Value = typeAbs
    nouvRow.Range(1, VAC_COL_PERIODE).Value = periodeAbs

    MsgBox "Absence enregistrée pour " & nom & " :" & vbCrLf & _
           typeAbs & " (" & periodeAbs & ") du " & Format(dDebut, "dd.mm.yyyy") & _
           " au " & Format(dFin, "dd.mm.yyyy"), _
           vbInformation, "Absences"

    frm.dtpVacDebut.Value = ""
    frm.dtpVacFin.Value = ""

    Call UF_ChargerAbsencesPersonne(frm)

End Sub

'==============================================================================
' FUNCTION : ExisteAbsenceChevauchante
' Retourne True si une absence existe déjà pour cette personne sur une
' période qui chevauche [dDebut ; dFin]. Évite les doublons/chevauchements
' silencieux dans Tbl_Vacances.
'==============================================================================
Private Function ExisteAbsenceChevauchante(ByVal wsV As Worksheet, ByVal nom As String, _
                                            ByVal dDebut As Date, ByVal dFin As Date) As Boolean

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsV.ListObjects(NOM_TBL_VACANCES)
    On Error GoTo 0
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then Exit Function

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        If Trim(tbl.DataBodyRange(i, VAC_COL_NOM).Value) = nom Then
            Dim dD As Date, dF As Date
            dD = 0: dF = 0
            On Error Resume Next
            dD = tbl.DataBodyRange(i, VAC_COL_DEBUT).Value
            dF = tbl.DataBodyRange(i, VAC_COL_FIN).Value
            On Error GoTo 0
            If dD <> 0 And dF <> 0 Then
                If dDebut <= dF And dFin >= dD Then
                    ExisteAbsenceChevauchante = True
                    Exit Function
                End If
            End If
        End If
    Next i

End Function

'==============================================================================
' ONGLET ABSENCES - Supprimer une absence selectionnee dans lstAbsListe
'==============================================================================
Sub UF_SupprimerAbsence(ByVal frm As Object)

    If frm.lstAbsListe.ListIndex = -1 Then
        MsgBox "Sélectionnez d'abord une absence dans la liste.", vbExclamation, "Suppression"
        Exit Sub
    End If

    Dim ligneStr As String
    ligneStr = frm.lstAbsListe.List(frm.lstAbsListe.ListIndex, 1)

    If ligneStr = "" Then Exit Sub

    Dim rep As VbMsgBoxResult
    rep = MsgBox("Supprimer cette absence ?", vbYesNo + vbQuestion, "Confirmer")

    If rep = vbNo Then Exit Sub

    Dim wsV As Worksheet
    Set wsV = Sheets(NOM_FEUILLE_VACANCES)

    ' Suppression via ListObject (convention du projet) au lieu de
    ' wsV.Rows(...).Delete, qui supprime la ligne sur toute la largeur
    ' de la feuille et pas seulement les colonnes du tableau.
    Dim tbl As ListObject
    Set tbl = wsV.ListObjects(NOM_TBL_VACANCES)
    tbl.ListRows(CLng(ligneStr) - tbl.HeaderRowRange.Row).Delete

    MsgBox "Absence supprimée.", vbInformation, "Absences v2.0"

    Call UF_ChargerAbsencesPersonne(frm)

End Sub

'==============================================================================
' ONGLET ABSENCES - Zone Remarques : prochaines absences (4 max)
'==============================================================================
Sub UF_MettreAJourRemarques(ByVal frm As Object)

    Dim wsV As Worksheet
    Set wsV = Sheets("Vacances")

    Dim dernLigne As Long
    dernLigne = wsV.Cells(wsV.Rows.count, 2).End(xlUp).Row

    Dim absences() As String
    Dim dates()    As Date
    Dim count      As Long
    count = 0

    ReDim absences(1 To 50)
    ReDim dates(1 To 50)

    Dim i As Long
    For i = 2 To dernLigne

        Dim nom    As String
        Dim dDebut As Date
        Dim dFin   As Date
        nom = wsV.Cells(i, 2).Value

        On Error Resume Next
        dDebut = wsV.Cells(i, 3).Value
        dFin = wsV.Cells(i, 4).Value
        On Error GoTo 0

        If nom = "" Then GoTo SuivantRem
        If dFin < Date Then GoTo SuivantRem

        count = count + 1
        dates(count) = dDebut
        absences(count) = nom & " - " & wsV.Cells(i, 5).Value & _
                          " : " & Format(dDebut, "dd.mm") & _
                          " -> " & Format(dFin, "dd.mm.yyyy")

SuivantRem:
    Next i

    If count = 0 Then
        frm.txtRemarques.Value = "Aucune absence prevue."
        Exit Sub
    End If

    ' Tri a bulles par date
    Dim j    As Long
    Dim tmpD As Date
    Dim tmpS As String
    For i = 1 To count - 1
        For j = 1 To count - i
            If dates(j) > dates(j + 1) Then
                tmpD = dates(j):    dates(j) = dates(j + 1):       dates(j + 1) = tmpD
                tmpS = absences(j): absences(j) = absences(j + 1): absences(j + 1) = tmpS
            End If
        Next j
    Next i

    Dim texte  As String
    Dim maxAff As Long
    maxAff = IIf(count < 4, count, 4)

    For i = 1 To maxAff
        texte = texte & absences(i)
        If i < maxAff Then texte = texte & vbCrLf
    Next i

    If count > 4 Then texte = texte & vbCrLf & "(+ " & count - 4 & " autre(s))"

    frm.txtRemarques.Value = texte

End Sub

'==============================================================================
' ONGLET ABSENCES - Charger le personnel dans cboVacNom (compatibilite)
' Conserve pour les appels existants depuis le UserForm
'==============================================================================
Sub UF_ChargerPersonnelVacances(ByVal frm As Object)
    Call UF_ChargerPersonnelAbsences(frm)
End Sub

'==============================================================================
' ONGLET ABSENCES - Charger toutes les absences (compatibilite)
'==============================================================================
Sub UF_ChargerVacances(ByVal frm As Object)
    Call UF_MettreAJourRemarques(frm)
End Sub


'==============================================================================
' ONGLET PARAMETRES - Charger les parametres dans le formulaire
'==============================================================================
Sub UF_ChargerParametres(ByVal frm As Object)

    Dim ws As Worksheet
    Set ws = Sheets("Parametres")

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects("Tbl_Parametres")
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "Tableau Tbl_Parametres introuvable.", vbExclamation
        Exit Sub
    End If

    ' Parcourt chaque ligne du tableau
    Dim i As Long
    For i = 1 To tbl.ListRows.count

        Dim machine As String
        machine = Trim(tbl.DataBodyRange(i, P_MACHINE).Value)

        Dim seuil1   As String
        Dim seuil2   As String
        Dim max_val  As String
        Dim ferme    As Boolean

        seuil1 = tbl.DataBodyRange(i, P_SEUIL1).Value
        seuil2 = tbl.DataBodyRange(i, P_SEUIL2).Value
        max_val = tbl.DataBodyRange(i, P_MAX).Value
        ferme = CBool(tbl.DataBodyRange(i, P_FERMETURE).Value)

        ' Affecte les valeurs aux bons controles selon la machine
        Select Case machine
            Case "Vaox5-A"
                frm.txtS1A.Value = seuil1
                frm.txtS2A.Value = seuil2
                frm.txtMaxA.Value = max_val
                frm.chkFermA.Value = ferme
            Case "Vaox5-B"
                frm.txtS1B.Value = seuil1
                frm.txtS2B.Value = seuil2
                frm.txtMaxB.Value = max_val
                frm.chkFermB.Value = ferme
            Case "Vaox5-C"
                frm.txtS1C.Value = seuil1
                frm.txtS2C.Value = seuil2
                frm.txtMaxC.Value = max_val
                frm.chkFermC.Value = ferme
            Case "Vaox5-D"
                frm.txtS1D.Value = seuil1
                frm.TxtS2D.Value = seuil2
                frm.txtMaxD.Value = max_val
                frm.chkFermD.Value = ferme
            Case "Vaox1-A&B"
                frm.chkFermV1.Value = ferme
        End Select

    Next i
    
    ' Effectif minimum semaine
    On Error Resume Next
    frm.txtEffectifMin.Value = CStr(Sheets(NOM_FEUILLE_PARAMETRES).Range("Effectif_Minimum").Value)
    On Error GoTo 0
    
End Sub

'==============================================================================
' ONGLET PARAMETRES - Enregistrer les parametres depuis le formulaire
'==============================================================================
Sub UF_EnregistrerParametres(ByVal frm As Object)

    ' --- Validation des saisies ---
    If Not ValiderSaisieNumerique(frm.txtS1A.Value, "Seuil 2 pers. Vaox5-A") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtS2A.Value, "Seuil 3 pers. Vaox5-A") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtMaxA.Value, "Max Vaox5-A") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtS1B.Value, "Seuil 2 pers. Vaox5-B") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtS2B.Value, "Seuil 3 pers. Vaox5-B") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtMaxB.Value, "Max Vaox5-B") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtS1C.Value, "Seuil 2 pers. Vaox5-C") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtS2C.Value, "Seuil 3 pers. Vaox5-C") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtMaxC.Value, "Max Vaox5-C") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtS1D.Value, "Seuil 2 pers. Vaox5-D") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.TxtS2D.Value, "Seuil 3 pers. Vaox5-D") Then Exit Sub
    If Not ValiderSaisieNumerique(frm.txtMaxD.Value, "Max Vaox5-D") Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sheets("Parametres")

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Tbl_Parametres")

    ' Ecrit chaque ligne du tableau
    Dim i As Long
    For i = 1 To tbl.ListRows.count

        Dim machine As String
        machine = Trim(tbl.DataBodyRange(i, P_MACHINE).Value)

        Select Case machine
            Case "Vaox5-A"
                tbl.DataBodyRange(i, P_SEUIL1).Value = CLng(frm.txtS1A.Value)
                tbl.DataBodyRange(i, P_SEUIL2).Value = CLng(frm.txtS2A.Value)
                tbl.DataBodyRange(i, P_MAX).Value = CLng(frm.txtMaxA.Value)
                tbl.DataBodyRange(i, P_FERMETURE).Value = frm.chkFermA.Value
            Case "Vaox5-B"
                tbl.DataBodyRange(i, P_SEUIL1).Value = CLng(frm.txtS1B.Value)
                tbl.DataBodyRange(i, P_SEUIL2).Value = CLng(frm.txtS2B.Value)
                tbl.DataBodyRange(i, P_MAX).Value = CLng(frm.txtMaxB.Value)
                tbl.DataBodyRange(i, P_FERMETURE).Value = frm.chkFermB.Value
            Case "Vaox5-C"
                tbl.DataBodyRange(i, P_SEUIL1).Value = CLng(frm.txtS1C.Value)
                tbl.DataBodyRange(i, P_SEUIL2).Value = CLng(frm.txtS2C.Value)
                tbl.DataBodyRange(i, P_MAX).Value = CLng(frm.txtMaxC.Value)
                tbl.DataBodyRange(i, P_FERMETURE).Value = frm.chkFermC.Value
            Case "Vaox5-D"
                tbl.DataBodyRange(i, P_SEUIL1).Value = CLng(frm.txtS1D.Value)
                tbl.DataBodyRange(i, P_SEUIL2).Value = CLng(frm.TxtS2D.Value)
                tbl.DataBodyRange(i, P_MAX).Value = CLng(frm.txtMaxD.Value)
                tbl.DataBodyRange(i, P_FERMETURE).Value = frm.chkFermD.Value
            Case "Vaox1-A&B"
                tbl.DataBodyRange(i, P_FERMETURE).Value = frm.chkFermV1.Value
        End Select

    Next i
    
        ' Sauvegarde effectif minimum
    If IsNumeric(frm.txtEffectifMin.Value) Then
        Sheets(NOM_FEUILLE_PARAMETRES).Range("Effectif_Minimum").Value = CLng(frm.txtEffectifMin.Value)
    End If

    MsgBox "Paramètres enregistrés avec succès !", vbInformation, "Parametres v2.0"

End Sub

'==============================================================================
' UTILITAIRE - Valide qu une saisie est bien numerique et positive
'==============================================================================
Private Function ValiderSaisieNumerique(ByVal valeur As String, _
                                         ByVal nomChamp As String) As Boolean
    If Not IsNumeric(valeur) Then
        MsgBox "Valeur invalide pour : " & nomChamp & vbCrLf & _
               "Veuillez saisir un nombre entier.", vbExclamation, "Validation"
        ValiderSaisieNumerique = False
        Exit Function
    End If
    If CLng(valeur) < 0 Then
        MsgBox nomChamp & " ne peut pas être négatif.", vbExclamation, "Validation"
        ValiderSaisieNumerique = False
        Exit Function
    End If
    ValiderSaisieNumerique = True
End Function

'==============================================================================
' SUB : UF_ChargerCalendrierPersonnel
' Point d entree principal - appele au clic sur lstPersonnel
' et aux clics sur btnCalPrev / btnCalNext
'==============================================================================
Sub UF_ChargerCalendrierPersonnel(ByVal frm As Object)

    ' Recupere la ligne de la personne selectionnee
    If frm.lstPersonnel.ListIndex = -1 Then
        Call EffacerCalendrierPersonnel(frm)
        Exit Sub
    End If

    Dim ligneStr As String
    ligneStr = frm.lstPersonnel.List(frm.lstPersonnel.ListIndex, 2)
    If ligneStr = "" Then Exit Sub

    Dim LignePers As Long
    LignePers = CLng(ligneStr)

    ' Recupere mois/annee affiches
    Dim moisCal  As Long
    Dim anneeCal As Long
    moisCal = CLng(frm.txtCalMois.Value)
    anneeCal = CLng(frm.txtCalAnnee.Value)

    ' Titre du mois
    Dim nomsMois(1 To 12) As String
    nomsMois(1) = "Janvier":    nomsMois(2) = "Fevrier":    nomsMois(3) = "Mars"
    nomsMois(4) = "Avril":      nomsMois(5) = "Mai":        nomsMois(6) = "Juin"
    nomsMois(7) = "Juillet":    nomsMois(8) = "Aout":       nomsMois(9) = "Septembre"
    nomsMois(10) = "Octobre":   nomsMois(11) = "Novembre":  nomsMois(12) = "Decembre"

    frm.lblCalTitre.Caption = nomsMois(moisCal) & " " & anneeCal

    ' Recupere les infos de la personne
    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim typePerso  As String
    Dim groupePerso As String
    typePerso = ws.Cells(LignePers, C_TYPE).Value
    groupePerso = ws.Cells(LignePers, C_GROUPE).Value

    ' Calcule la grille
    Dim premierJour As Date
    premierJour = DateSerial(anneeCal, moisCal, 1)
    Dim decalage As Long
    decalage = Weekday(premierJour, vbMonday) - 1
    Dim debutGrille As Date
    debutGrille = premierJour - decalage

    ' Date de reference G1 pour les auxiliaires
    Dim dateRefG1 As Date
    On Error Resume Next
    dateRefG1 = Sheets("Planning journalier").Range("Ref_GroupeG1").Value
    On Error GoTo 0

    ' Adapte le bouton selon le type
    If typePerso = "Fixe" Then
        frm.btnAjouterAbsence.Caption = "Ajouter une absence"
    Else
        frm.btnAjouterAbsence.Caption = "Enregistrer un remplacement"
    End If

    ' Dessine les 42 cases
    Dim i As Long
    For i = 0 To 41

        Dim infoJour  As String
        Dim dateCase  As Date
        Dim lbl       As MSForms.Label
        Dim couleur   As Long

        dateCase = debutGrille + i
        Set lbl = frm.Controls("lblCal" & i)
        lbl.Tag = CStr(CDbl(dateCase))

        ' --- Determine la couleur ---
        If Month(dateCase) <> moisCal Or Year(dateCase) <> anneeCal Then
            couleur = CAL_COULEUR_HORS_MOIS
            lbl.ForeColor = RGB(180, 180, 180)
        Else
            lbl.ForeColor = RGB(0, 0, 0)
            If typePerso = "Fixe" Then
                couleur = CouleurJourFixe(LignePers, dateCase, ws)
            Else
                couleur = CouleurJourAuxiliaire(ws.Cells(LignePers, C_NOM).Value, _
                                                dateCase, groupePerso, dateRefG1)
            End If
        End If

        lbl.BackColor = couleur

        ' --- Caption : numero du jour + info ---
        infoJour = ""

        If Month(dateCase) = moisCal And Year(dateCase) = anneeCal Then

            If typePerso = "Fixe" Then
                ' Affiche M/J uniquement si la personne travaille ce jour
                ' (pas sur les feries, vacances, maladie, conge, weekend)
                If couleur = CAL_COULEUR_MATIN Or couleur = CAL_COULEUR_JOURNEE Then
                    Dim jourSemF  As Long
                    jourSemF = Weekday(dateCase, vbMonday)
                    If jourSemF <= 5 Then
                        Dim codeHoraire As String
                        codeHoraire = Trim(ws.Cells(LignePers, C_LUN + (jourSemF - 1)).Value)
                        If codeHoraire = "M" Or codeHoraire = "J" Then infoJour = codeHoraire
                    End If
                End If
            Else
                ' Auxiliaire
                If couleur = CAL_COULEUR_AUX_WORK Then infoJour = ws.Cells(LignePers, C_GROUPE).Value
                If couleur = CAL_COULEUR_AUX_REMPL Then infoJour = "Rempl."
                If couleur = CAL_COULEUR_AUX_ABS Then infoJour = "Absent"
                If couleur = CAL_COULEUR_FERIE Then
                    Dim groupeCase As String
                    groupeCase = Module_Groupes.GetGroupeJour(dateCase)
                    If groupeCase <> "" Then infoJour = groupeCase
                End If
            End If

        End If

        If infoJour <> "" Then
            lbl.Caption = Day(dateCase) & vbCrLf & infoJour
        Else
            lbl.Caption = CStr(Day(dateCase))
        End If

        ' Tooltip remplacement pour auxiliaires
        If typePerso = "Auxiliaire" And _
           (couleur = CAL_COULEUR_AUX_ABS Or couleur = CAL_COULEUR_AUX_REMPL) Then
            lbl.ControlTipText = GetTooltipRemplacement( _
                ws.Cells(LignePers, C_NOM).Value, dateCase)
        End If

        ' Encadre rouge pour aujourd hui
        If Int(dateCase) = Int(Date) And Month(dateCase) = moisCal Then
            lbl.BorderStyle = 1
            lbl.BorderColor = RGB(200, 0, 0)
            lbl.Font.Bold = True
        Else
            lbl.BorderStyle = 1
            lbl.BorderColor = RGB(200, 200, 200)
            lbl.Font.Bold = False
        End If

    Next i
    
End Sub

'==============================================================================
' FUNCTION : CouleurJourFixe
' Determine la couleur d un jour pour un employe fixe
' Priorite : Ferie > Absence > Horaire
'==============================================================================
Private Function CouleurJourFixe(ByVal LignePers As Long, _
                                   ByVal dateCase As Date, _
                                   ByVal ws As Worksheet) As Long

    ' 1. Jour ferie ?
    If Module_Feries.estFerie(dateCase) Then
        CouleurJourFixe = CAL_COULEUR_FERIE
        Exit Function
    End If

    ' 2. Absence enregistree ?
    Dim wsV As Worksheet
    Set wsV = Sheets("Vacances")
    Dim nom As String
    nom = ws.Cells(LignePers, C_NOM).Value

    Dim dernLigne As Long
    dernLigne = wsV.Cells(wsV.Rows.count, 2).End(xlUp).Row

    Dim j As Long
    For j = 2 To dernLigne
        If Trim(wsV.Cells(j, 2).Value) = nom Then
            Dim dDebut As Date
            Dim dFin   As Date
            Dim typeAbs As String
            On Error Resume Next
            dDebut = wsV.Cells(j, 3).Value
            dFin = wsV.Cells(j, 4).Value
            On Error GoTo 0
            typeAbs = Trim(wsV.Cells(j, 5).Value)

            If dateCase >= dDebut And dateCase <= dFin Then
                Select Case Trim(typeAbs)
                    Case "Maladie":  CouleurJourFixe = CAL_COULEUR_MALADIE
                    Case "Conge":    CouleurJourFixe = CAL_COULEUR_CONGE
                    Case Else:       CouleurJourFixe = CAL_COULEUR_VACANCES
                End Select
                Exit Function
            End If
        End If
    Next j

    ' 3. Horaire normal
    Dim jourSem As Long
    jourSem = Weekday(dateCase, vbMonday)

    ' Weekend = gris pour les fixes
    If jourSem = 6 Or jourSem = 7 Then
        CouleurJourFixe = CAL_COULEUR_ABSENT
        Exit Function
    End If

    Dim codeH As String
    codeH = Trim(ws.Cells(LignePers, C_LUN + (jourSem - 1)).Value)

    Select Case codeH
        Case "M":    CouleurJourFixe = CAL_COULEUR_MATIN
        Case "J":    CouleurJourFixe = CAL_COULEUR_JOURNEE
        Case Else:   CouleurJourFixe = CAL_COULEUR_ABSENT
    End Select

End Function

'==============================================================================
' FUNCTION : CouleurJourAuxiliaire
' Determine la couleur d un jour pour un auxiliaire
'==============================================================================
Private Function CouleurJourAuxiliaire(ByVal nom As String, _
                                        ByVal dateCase As Date, _
                                        ByVal groupe As String, _
                                        ByVal dateRefG1 As Date) As Long

    ' 1. Jour ferie : traite comme un weekend (les auxiliaires peuvent travailler)
    ' La couleur sera orange mais on continue pour vérifier les remplacements et le groupe
    Dim estFerieJour As Boolean
    estFerieJour = Module_Feries.estFerie(dateCase)

    ' 2. Jour en semaine (pas WE)
    Dim jourSem As Long
    jourSem = Weekday(dateCase, vbMonday)
    If jourSem <> 6 And jourSem <> 7 Then
        ' Si c'est un férié en semaine -> continue vers section 3 pour G1/G2
        If Not estFerieJour Then
            ' Semaine normale : vérifier si remplacement noté
            Dim wsRpl As Worksheet
            Set wsRpl = Sheets(NOM_FEUILLE_REMPLACEMENTS)
            Dim dernLigne As Long
            dernLigne = wsRpl.Cells(wsRpl.Rows.count, 1).End(xlUp).Row
            Dim k As Long
            For k = 2 To dernLigne
                Dim dateRpl As Date
                On Error Resume Next
                dateRpl = wsRpl.Cells(k, 1).Value
                On Error GoTo 0
                If Int(dateRpl) = Int(dateCase) Then
                    If Trim(wsRpl.Cells(k, 5).Value) = nom Then
                        CouleurJourAuxiliaire = CAL_COULEUR_AUX_REMPL
                        Exit Function
                    End If
                End If
            Next k
            CouleurJourAuxiliaire = CAL_COULEUR_ABSENT
            Exit Function
        End If
    End If
    
    ' 3. Weekend : calcule le groupe actif
    Dim lundiSem As Date
    lundiSem = dateCase - (jourSem - 1)
    Dim groupeActif As String
    Dim groupeFeuille As String
    groupeFeuille = Module_Groupes.GetGroupeJour(dateCase)
    If groupeFeuille <> "" Then
        groupeActif = groupeFeuille
    Else
        Dim nbSem As Long
        nbSem = Int((lundiSem - dateRefG1) / 7)
        groupeActif = IIf(nbSem Mod 2 = 0, "G1", "G2")
    End If

    ' 4. Verifie les remplacements
    Dim wsR As Worksheet
    Set wsR = Sheets("Remplacements")
    Dim dernL As Long
    dernL = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernL
        Dim dRpl As Date
        On Error Resume Next
        dRpl = wsR.Cells(i, 1).Value
        On Error GoTo 0
        If Int(dRpl) = Int(dateCase) Then
            ' Elle remplace quelqu un
            If Trim(wsR.Cells(i, 5).Value) = nom Then
                CouleurJourAuxiliaire = CAL_COULEUR_AUX_REMPL
                Exit Function
            End If
            ' Elle se fait remplacer
            If Trim(wsR.Cells(i, 3).Value) = nom Then
                CouleurJourAuxiliaire = CAL_COULEUR_AUX_ABS
                Exit Function
            End If
        End If
    Next i

    ' 5. Son groupe est-il actif ce jour ?
    If groupe = groupeActif Then
        CouleurJourAuxiliaire = IIf(estFerieJour, CAL_COULEUR_FERIE, CAL_COULEUR_AUX_WORK)
    Else
        CouleurJourAuxiliaire = IIf(estFerieJour, CAL_COULEUR_FERIE, CAL_COULEUR_ABSENT)
    End If

End Function

'==============================================================================
' SUB : EffacerCalendrierPersonnel
' Vide le calendrier quand aucune personne n est selectionnee
'==============================================================================
Private Sub EffacerCalendrierPersonnel(ByVal frm As Object)

    frm.lblCalTitre.Caption = ""

    Dim i As Long
    For i = 0 To 41
        Dim lbl As MSForms.Label
        Set lbl = frm.Controls("lblCal" & i)
        lbl.Caption = ""
        lbl.BackColor = CAL_COULEUR_BLANC
        lbl.Font.Bold = False
    Next i

End Sub

'==============================================================================
' SUB : UF_CalendrierMoisPrecedent
' Navigation - mois precedent
'==============================================================================
Sub UF_CalendrierMoisPrecedent(ByVal frm As Object)

    Dim mois  As Long
    Dim annee As Long
    mois = CLng(frm.txtCalMois.Value)
    annee = CLng(frm.txtCalAnnee.Value)

    mois = mois - 1
    If mois < 1 Then
        mois = 12
        annee = annee - 1
    End If

    frm.txtCalMois.Value = CStr(mois)
    frm.txtCalAnnee.Value = CStr(annee)

    Call UF_ChargerCalendrierPersonnel(frm)

End Sub

'==============================================================================
' SUB : UF_CalendrierMoisSuivant
' Navigation - mois suivant
'==============================================================================
Sub UF_CalendrierMoisSuivant(ByVal frm As Object)

    Dim mois  As Long
    Dim annee As Long
    mois = CLng(frm.txtCalMois.Value)
    annee = CLng(frm.txtCalAnnee.Value)

    mois = mois + 1
    If mois > 12 Then
        mois = 1
        annee = annee + 1
    End If

    frm.txtCalMois.Value = CStr(mois)
    frm.txtCalAnnee.Value = CStr(annee)

    Call UF_ChargerCalendrierPersonnel(frm)

End Sub


'==============================================================================
' SUB : UF_AjouterAbsenceDepuisCalendrier
' Ouvre le mini UserForm d absence et enregistre si confirme
'==============================================================================
Private Sub UF_AjouterAbsenceDepuisCalendrier(ByVal frm As Object, _
                                               ByVal nomPerso As String, _
                                               ByVal dateJour As Date)

    ' Ouvre le mini UserForm
    Dim dlg As UserForm_Absence
    Set dlg = New UserForm_Absence

    dlg.NomPersonne = nomPerso
    dlg.DateProposee = dateJour
    dlg.Initialiser

    dlg.Show

    ' Verifie si l utilisateur a confirme
    If Not dlg.Confirme Then
        Unload dlg
        Exit Sub
    End If

    ' Recupere les valeurs saisies
    Dim typeAbs    As String
    Dim periodeAbs As String
    Dim dDebut     As Date
    Dim dFin       As Date

    typeAbs = dlg.cboAbsType.Value
    periodeAbs = dlg.cboAbsPeriode.Value
    If periodeAbs = "" Then periodeAbs = PERIODE_JOURNEE
    dDebut = CDate(dlg.txtAbsDebut.Value)
    dFin = CDate(dlg.txtAbsFin.Value)

    Unload dlg

    ' Recupere l ID de la personne
    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_PERSONNEL)

    Dim idPerso  As String
    Dim lignePer As Variant
    lignePer = Application.Match(nomPerso, ws.Columns(C_NOM), 0)
    If Not IsError(lignePer) Then
        idPerso = ws.Cells(lignePer, C_ID).Value
    End If

    ' Enregistre dans la feuille Vacances via ListObject.ListRows.Add
    ' (convention du projet) au lieu d'écrire directement sous la
    ' dernière ligne via ws.Cells().
    Dim wsV As Worksheet
    Set wsV = Sheets(NOM_FEUILLE_VACANCES)

    If ExisteAbsenceChevauchante(wsV, nomPerso, dDebut, dFin) Then
        Dim repChevauchement As VbMsgBoxResult
        repChevauchement = MsgBox("Une absence existe déjà pour " & nomPerso & _
               " sur une période qui chevauche ces dates." & vbCrLf & _
               "Enregistrer quand même ?", vbYesNo + vbExclamation, "Chevauchement détecté")
        If repChevauchement = vbNo Then Exit Sub
    End If

    Dim tbl As ListObject
    Set tbl = wsV.ListObjects(NOM_TBL_VACANCES)
    Dim nouvRow As ListRow
    Set nouvRow = tbl.ListRows.Add

    nouvRow.Range(1, VAC_COL_ID).Value = idPerso
    nouvRow.Range(1, VAC_COL_NOM).Value = nomPerso
    nouvRow.Range(1, VAC_COL_DEBUT).Value = dDebut
    nouvRow.Range(1, VAC_COL_FIN).Value = dFin
    nouvRow.Range(1, VAC_COL_TYPE).Value = typeAbs
    nouvRow.Range(1, VAC_COL_PERIODE).Value = periodeAbs

    MsgBox typeAbs & " (" & periodeAbs & ") enregistrée pour " & nomPerso & vbCrLf & _
           "Du " & Format(dDebut, "dd.mm.yyyy") & _
           " au " & Format(dFin, "dd.mm.yyyy"), _
           vbInformation, "Absence ajoutée"

    ' Rafraichit le calendrier
    Call UF_ChargerCalendrierPersonnel(frm)

End Sub


'==============================================================================
' SUB : UF_BoutonAjouterAbsence
' Bouton "Ajouter une absence" sous le calendrier
' Bascule vers l onglet Absences avec la personne pre-selectionnee
'==============================================================================
Sub UF_BoutonAjouterAbsence(ByVal frm As Object)

    If frm.lstPersonnel.ListIndex = -1 Then
        MsgBox "Veuillez d'abord sélectionner une personne.", vbExclamation
        Exit Sub
    End If

    ' Recupere le nom de la personne selectionnee
    Dim nomPerso As String
    nomPerso = frm.lstPersonnel.List(frm.lstPersonnel.ListIndex, 0)

    ' Nettoie le prefixe [G1] ou [G2] si present
    If Left(nomPerso, 1) = "[" Then
        nomPerso = Trim(Mid(nomPerso, InStr(nomPerso, "]") + 1))
    End If

    ' Bascule vers l onglet Absences (index 2)
    frm.mpOnglets.Value = 2

    ' Pre-selectionne la personne dans lstAbsPersonnel
    Dim i As Long
    For i = 0 To frm.lstAbsPersonnel.ListCount - 1
        If frm.lstAbsPersonnel.List(i) = nomPerso Then
            frm.lstAbsPersonnel.ListIndex = i
            Call UF_ChargerAbsencesPersonne(frm)
            Exit For
        End If
    Next i

End Sub
'==============================================================================
' SUB : UF_BoutonCalendrierPersonnel
' Remplace btnAjouterAbsence_Click
' Comportement different selon Fixe ou Auxiliaire
'==============================================================================
Sub UF_BoutonCalendrierPersonnel(ByVal frm As Object)

    If frm.lstPersonnel.ListIndex = -1 Then
        MsgBox "Veuillez d'abord sélectionner une personne.", vbExclamation
        Exit Sub
    End If

    Dim ligneStr As String
    ligneStr = frm.lstPersonnel.List(frm.lstPersonnel.ListIndex, 2)
    If ligneStr = "" Then Exit Sub

    Dim LignePers As Long
    LignePers = CLng(ligneStr)

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim typePerso As String
    typePerso = ws.Cells(LignePers, C_TYPE).Value

    If typePerso = "Fixe" Then
        ' Ouvre UserForm_Absence avec la date du jour
        Call UF_AjouterAbsenceDepuisCalendrier(frm, _
             ws.Cells(LignePers, C_NOM).Value, Date)
    Else
        ' Ouvre UserForm_Remplacement avec la date du jour
        Call UF_AjouterRemplacementDepuisCalendrier(frm, _
             ws.Cells(LignePers, C_NOM).Value, Date)
    End If

End Sub

'==============================================================================
' SUB : UF_DoubleclicCalendrierAuxiliaire
' Double-clic sur une case pour un auxiliaire
' Affiche les infos + propose ajout/suppression remplacement
'==============================================================================
Sub UF_DoubleclicCalendrierAuxiliaire(ByVal frm As Object, _
                                       ByVal nomPerso As String, _
                                       ByVal dateJour As Date)

    ' --- Construit le message d infos ---
    Dim msgInfos As String
    msgInfos = Format(dateJour, "dddd dd.mm.yyyy") & vbCrLf
    msgInfos = msgInfos & String(40, "-") & vbCrLf

    ' Jour ferie ?
    If Module_Feries.estFerie(dateJour) Then
        msgInfos = msgInfos & "Jour ferie : " & Module_Feries.GetNomFerie(dateJour) & vbCrLf
    End If

    ' Groupe actif ce jour
    Dim dateRefG1 As Date
    On Error Resume Next
    dateRefG1 = Sheets("Planning journalier").Range("Ref_GroupeG1").Value
    On Error GoTo 0

    Dim jourSem As Long
    jourSem = Weekday(dateJour, vbMonday)

    Dim groupeActif As String
    If jourSem = 6 Or jourSem = 7 Or Module_Feries.estFerie(dateJour) Then
        Dim lundiSem As Date
        lundiSem = dateJour - (jourSem - 1)
        Dim nbSem As Long
        nbSem = Int((lundiSem - dateRefG1) / 7)
        groupeActif = IIf(nbSem Mod 2 = 0, "G1", "G2")
        msgInfos = msgInfos & "Groupe actif : " & groupeActif & vbCrLf
    End If

    ' Verifie remplacements existants
    Dim wsRpl As Worksheet
    Set wsRpl = Sheets("Remplacements")

    Dim dernLigne As Long
    dernLigne = wsRpl.Cells(wsRpl.Rows.count, 1).End(xlUp).Row

    Dim ligneRpl    As Long
    Dim nomRempl    As String
    Dim estAbsente  As Boolean
    Dim estRempl    As Boolean
    ligneRpl = 0
    estAbsente = False
    estRempl = False

    Dim j As Long
    For j = 2 To dernLigne
        Dim dRpl As Date
        On Error Resume Next
        dRpl = wsRpl.Cells(j, 1).Value
        On Error GoTo 0

        If Int(dRpl) = Int(dateJour) Then
            ' Elle se fait remplacer
            If Trim(wsRpl.Cells(j, 3).Value) = nomPerso Then
                estAbsente = True
                ligneRpl = j
                nomRempl = wsRpl.Cells(j, 5).Value
                msgInfos = msgInfos & "Remplacee par : " & nomRempl & vbCrLf
            End If
            ' Elle remplace quelqu un
            If Trim(wsRpl.Cells(j, 5).Value) = nomPerso Then
                estRempl = True
                msgInfos = msgInfos & "Remplace : " & wsRpl.Cells(j, 3).Value & vbCrLf
            End If
        End If
    Next j

    ' --- Propose les actions ---
    Dim choix As VbMsgBoxResult

    If estAbsente Then
        ' Propose suppression du remplacement
        choix = MsgBox(msgInfos & vbCrLf & _
                       "Voulez-vous annuler ce remplacement ?", _
                       vbYesNo + vbQuestion, nomPerso)

        If choix = vbYes Then
            ' Suppression via ListObject (convention du projet).
            Dim tblRplDel As ListObject
            Set tblRplDel = wsRpl.ListObjects(NOM_TBL_REMPLACEMENTS)
            tblRplDel.ListRows(ligneRpl - tblRplDel.HeaderRowRange.Row).Delete
            MsgBox "Remplacement annulé.", vbInformation, "Calendrier"
            Call UF_ChargerCalendrierPersonnel(frm)
        End If

    Else
        ' Propose ajout d un remplacement
        choix = MsgBox(msgInfos & vbCrLf & _
                       "Voulez-vous enregistrer un remplacement pour ce jour ?", _
                       vbYesNo + vbQuestion, nomPerso)

        If choix = vbYes Then
            Call UF_AjouterRemplacementDepuisCalendrier(frm, nomPerso, dateJour)
        End If

    End If

End Sub

'==============================================================================
' SUB : UF_AjouterRemplacementDepuisCalendrier
' Ouvre UserForm_Remplacement et enregistre si confirme
'==============================================================================
Sub UF_AjouterRemplacementDepuisCalendrier(ByVal frm As Object, _
                                            ByVal nomPerso As String, _
                                            ByVal dateJour As Date)

    Dim dlg As UserForm_Remplacement
    Set dlg = New UserForm_Remplacement

    dlg.NomAbsent = nomPerso
    dlg.DateProposee = dateJour
    dlg.Initialiser

    dlg.Show

    If Not dlg.Confirme Then
        Unload dlg
        Exit Sub
    End If

    Dim nomRempl As String
    Dim dateRpl  As Date
    nomRempl = dlg.cboRemplacant.Value
    dateRpl = CDate(dlg.txtRplDate.Value)   ' validee par IsDate dans le dialogue

    Unload dlg

    ' Recupere les IDs
    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim idAbsent As String
    Dim idRempl  As String
    Dim ligneAbs As Variant
    Dim ligneRpl As Variant

    ligneAbs = Application.Match(nomPerso, ws.Columns(C_NOM), 0)
    ligneRpl = Application.Match(nomRempl, ws.Columns(C_NOM), 0)

    If Not IsError(ligneAbs) Then idAbsent = ws.Cells(ligneAbs, C_ID).Value
    If Not IsError(ligneRpl) Then idRempl = ws.Cells(ligneRpl, C_ID).Value

    ' Enregistre dans Remplacements — ce chemin (depuis le calendrier
    ' personnel) enregistre toujours un vrai remplacement d'une personne
    ' nommément absente, jamais un renfort générique.
    Dim wsRpl As Worksheet
    Set wsRpl = Sheets(NOM_FEUILLE_REMPLACEMENTS)

    Dim tbl As ListObject
    Set tbl = wsRpl.ListObjects(NOM_TBL_REMPLACEMENTS)
    Dim nouvRow As ListRow
    Set nouvRow = tbl.ListRows.Add

    nouvRow.Range(1, RPL_COL_DATE).Value = dateRpl
    nouvRow.Range(1, RPL_COL_DATE).NumberFormat = "dd.mm.yyyy"
    nouvRow.Range(1, RPL_COL_ID_ABSENTE).Value = idAbsent
    nouvRow.Range(1, RPL_COL_NOM_ABSENTE).Value = nomPerso
    nouvRow.Range(1, RPL_COL_ID_REMPLACANT).Value = idRempl
    nouvRow.Range(1, RPL_COL_NOM_REMPLACANT).Value = nomRempl
    nouvRow.Range(1, RPL_COL_TYPE).Value = "Remplacement"

    MsgBox "Remplacement enregistré :" & vbCrLf & _
           nomRempl & " remplace " & nomPerso & vbCrLf & _
           "Le " & Format(dateRpl, "dd.mm.yyyy"), _
           vbInformation, "Remplacement"

    Call UF_ChargerCalendrierPersonnel(frm)

End Sub

'==============================================================================
' MISE A JOUR : UF_DoubleclicCalendrier
' Version complete qui gere Fixe ET Auxiliaire
' REMPLACE la version existante dans Module_UserForm
'==============================================================================
Sub UF_DoubleclicCalendrier_V2(ByVal frm As Object, ByVal indexJour As Integer)

    ' Recupere la date de la case
    Dim lbl As MSForms.Label
    Set lbl = frm.Controls("lblCal" & indexJour)

    If lbl.Tag = "" Or lbl.Caption = "" Then Exit Sub

    Dim dateJour As Date
    dateJour = CDate(CDbl(lbl.Tag))

    ' Verifie que le jour est dans le mois affiche
    Dim moisCal  As Long
    Dim anneeCal As Long
    moisCal = CLng(frm.txtCalMois.Value)
    anneeCal = CLng(frm.txtCalAnnee.Value)

    If Month(dateJour) <> moisCal Or Year(dateJour) <> anneeCal Then
        ' Navigation vers ce mois
        frm.txtCalMois.Value = CStr(Month(dateJour))
        frm.txtCalAnnee.Value = CStr(Year(dateJour))
        Call UF_ChargerCalendrierPersonnel(frm)
        Exit Sub
    End If

    ' Recupere la personne selectionnee
    If frm.lstPersonnel.ListIndex = -1 Then Exit Sub

    Dim ligneStr As String
    ligneStr = frm.lstPersonnel.List(frm.lstPersonnel.ListIndex, 2)
    If ligneStr = "" Then Exit Sub

    Dim LignePers As Long
    LignePers = CLng(ligneStr)

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim nomPerso  As String
    Dim typePerso As String
    nomPerso = ws.Cells(LignePers, C_NOM).Value
    typePerso = ws.Cells(LignePers, C_TYPE).Value

    ' Branche selon le type
    If typePerso = "Auxiliaire" Then
        Call UF_DoubleclicCalendrierAuxiliaire(frm, nomPerso, dateJour)
    Else
        Call UF_DoubleclicCalendrier_Fixe(frm, nomPerso, LignePers, dateJour, ws)
    End If

End Sub

'==============================================================================
' SUB : UF_DoubleclicCalendrier_Fixe
' Logique double-clic pour le personnel fixe (extraite de l ancienne version)
'==============================================================================
Private Sub UF_DoubleclicCalendrier_Fixe(ByVal frm As Object, _
                                          ByVal nomPerso As String, _
                                          ByVal LignePers As Long, _
                                          ByVal dateJour As Date, _
                                          ByVal ws As Worksheet)

    Dim msgInfos As String
    msgInfos = Format(dateJour, "dddd dd.mm.yyyy") & vbCrLf
    msgInfos = msgInfos & String(40, "-") & vbCrLf

    If Module_Feries.estFerie(dateJour) Then
        msgInfos = msgInfos & "Jour ferie : " & Module_Feries.GetNomFerie(dateJour) & vbCrLf
    End If

    Dim jourSem As Long
    jourSem = Weekday(dateJour, vbMonday)

    If jourSem <= 5 Then
        Dim codeH As String
        codeH = Trim(ws.Cells(LignePers, C_LUN + (jourSem - 1)).Value)
        Select Case codeH
            Case "M": msgInfos = msgInfos & "Horaire : Matin" & vbCrLf
            Case "J": msgInfos = msgInfos & "Horaire : Journee complete" & vbCrLf
            Case Else: msgInfos = msgInfos & "Horaire : Absent" & vbCrLf
        End Select
    Else
        msgInfos = msgInfos & "Horaire : Weekend" & vbCrLf
    End If

    ' Absence existante ?
    Dim wsV As Worksheet
    Set wsV = Sheets("Vacances")
    Dim dernLigne As Long
    dernLigne = wsV.Cells(wsV.Rows.count, 2).End(xlUp).Row

    Dim ligneAbsence As Long
    Dim typeAbsExist As String
    ligneAbsence = 0

    Dim j As Long
    For j = 2 To dernLigne
        If Trim(wsV.Cells(j, 2).Value) = nomPerso Then
            Dim dDebut As Date
            Dim dFin   As Date
            On Error Resume Next
            dDebut = wsV.Cells(j, 3).Value
            dFin = wsV.Cells(j, 4).Value
            On Error GoTo 0
            If dateJour >= dDebut And dateJour <= dFin Then
                ligneAbsence = j
                typeAbsExist = wsV.Cells(j, 5).Value
                msgInfos = msgInfos & "Absence : " & typeAbsExist & vbCrLf & _
                           "  Du " & Format(dDebut, "dd.mm.yyyy") & _
                           " au " & Format(dFin, "dd.mm.yyyy") & vbCrLf
                Exit For
            End If
        End If
    Next j

    Dim choix As VbMsgBoxResult

    If ligneAbsence > 0 Then
        choix = MsgBox(msgInfos & vbCrLf & "Voulez-vous supprimer cette absence ?", _
                       vbYesNo + vbQuestion, nomPerso & " - " & Format(dateJour, "dd.mm.yyyy"))
        If choix = vbYes Then
            ' Suppression via ListObject (convention du projet).
            Dim tblVacDel As ListObject
            Set tblVacDel = wsV.ListObjects(NOM_TBL_VACANCES)
            tblVacDel.ListRows(ligneAbsence - tblVacDel.HeaderRowRange.Row).Delete
            MsgBox "Absence supprimée.", vbInformation, "Calendrier"
            Call UF_ChargerCalendrierPersonnel(frm)
        End If
    Else
        choix = MsgBox(msgInfos & vbCrLf & "Voulez-vous ajouter une absence ?", _
                       vbYesNo + vbQuestion, nomPerso & " - " & Format(dateJour, "dd.mm.yyyy"))
        If choix = vbYes Then
            Call UF_AjouterAbsenceDepuisCalendrier(frm, nomPerso, dateJour)
        End If
    End If

End Sub

'==============================================================================
' FUNCTION : GetTooltipRemplacement
' Retourne le texte du tooltip pour une case de remplacement
'==============================================================================
Function GetTooltipRemplacement(ByVal nomPerso As String, _
                                  ByVal dateJour As Date) As String

    Dim wsRpl As Worksheet
    Set wsRpl = Sheets("Remplacements")

    Dim dernLigne As Long
    dernLigne = wsRpl.Cells(wsRpl.Rows.count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernLigne
        Dim dRpl As Date
        On Error Resume Next
        dRpl = wsRpl.Cells(i, 1).Value
        On Error GoTo 0
        If Int(dRpl) = Int(dateJour) Then
            ' La personne est absente
            If Trim(wsRpl.Cells(i, 3).Value) = nomPerso Then
                GetTooltipRemplacement = "Remplacee par : " & wsRpl.Cells(i, 5).Value
                Exit Function
            End If
            ' La personne remplace quelqu un
            If Trim(wsRpl.Cells(i, 5).Value) = nomPerso Then
                GetTooltipRemplacement = "Remplace : " & wsRpl.Cells(i, 3).Value
                Exit Function
            End If
        End If
    Next i

End Function


'==============================================================================
' SUB : UF_ImprimerFiche
' Point d entree depuis le bouton "Imprimer la fiche"
'==============================================================================
Sub UF_ImprimerFiche(ByVal frm As Object)

    If frm.lstPersonnel.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner une personne.", vbExclamation
        Exit Sub
    End If

    Dim ligneStr As String
    ligneStr = frm.lstPersonnel.List(frm.lstPersonnel.ListIndex, 2)
    If ligneStr = "" Then Exit Sub

    Dim LignePers As Long
    LignePers = CLng(ligneStr)

    Dim ws As Worksheet
    Set ws = Sheets("Personnel")

    Dim nomPerso As String
    nomPerso = ws.Cells(LignePers, C_NOM).Value

    ' Ouvre le UserForm de choix
    Dim dlg As UserForm_ImpressionFiche
    Set dlg = New UserForm_ImpressionFiche
    dlg.NomPersonne = nomPerso
    dlg.LignePers = LignePers
    dlg.Initialiser
    dlg.Show

    If Not dlg.Confirme Then
        Unload dlg
        Exit Sub
    End If

    Dim avecHoraire   As Boolean
    Dim avecAptitudes As Boolean
    Dim avecStats     As Boolean
    Dim avecAbsences  As Boolean

    avecHoraire = dlg.chkHoraire.Value
    avecAptitudes = dlg.chkAptitudes.Value
    avecStats = dlg.chkStats.Value
    avecAbsences = dlg.chkAbsences.Value

    Unload dlg

    ' Genere la fiche dans une feuille temporaire
    Call GenererFicheExcel(LignePers, ws, nomPerso, _
                           avecHoraire, avecAptitudes, avecStats, avecAbsences)

End Sub

'==============================================================================
' SUB : GenererFicheExcel
' Cree une feuille temporaire avec la fiche mise en forme et imprime
'==============================================================================
Private Sub GenererFicheExcel(ByVal LignePers As Long, _
                               ByVal ws As Worksheet, _
                               ByVal nomPerso As String, _
                               ByVal avecHoraire As Boolean, _
                               ByVal avecAptitudes As Boolean, _
                               ByVal avecStats As Boolean, _
                               ByVal avecAbsences As Boolean)

    ' Gestion d'erreur globale : si une erreur survient n'importe où pendant
    ' la construction de la fiche, ScreenUpdating/DisplayAlerts doivent
    ' impérativement être restaurés, sinon Excel reste visuellement figé
    ' pour l'utilisateur sans aucun message d'erreur.
    On Error GoTo ErrGeneration

    ' Desactive les alertes Excel
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' Supprime l ancienne feuille temporaire si elle existe
    On Error Resume Next
    ThisWorkbook.Sheets("_Fiche_Temp").Delete
    On Error GoTo 0

    ' Cree la feuille temporaire
    Dim wsFiche As Worksheet
    Set wsFiche = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsFiche.Name = "_Fiche_Temp"

    ' Configuration de la page
    With wsFiche.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
    End With

    ' Largeur des colonnes
    wsFiche.Columns("A").ColumnWidth = 20
    wsFiche.Columns("B").ColumnWidth = 22
    wsFiche.Columns("C").ColumnWidth = 20
    wsFiche.Columns("D").ColumnWidth = 22

    Dim ligne As Long
    ligne = 1

    '==========================================================================
    ' EN-TETE
    '==========================================================================
    ' Titre principal
    With wsFiche.Range("A" & ligne & ":D" & ligne)
        .Merge
        .Value = "Fiche du Personnel"
        .Font.Bold = True
        .Font.Size = 18
        .Font.Color = IMP_ROUGE
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    ligne = ligne + 1

    ' Date d impression
    With wsFiche.Range("A" & ligne & ":D" & ligne)
        .Merge
        .Value = "Imprime le : " & Format(Now, "dd.mm.yyyy") & "  -  " & Format(Now, "hh:mm")
        .Font.Size = 9
        .Font.Color = RGB(128, 128, 128)
        .HorizontalAlignment = xlRight
        .RowHeight = 16
    End With
    ligne = ligne + 1

    ' Ligne rouge separatrice
    With wsFiche.Range("A" & ligne & ":D" & ligne)
        .Merge
        .Interior.Color = IMP_ROUGE
        .RowHeight = 3
    End With
    ligne = ligne + 1

    '==========================================================================
    ' INFORMATIONS GENERALES
    '==========================================================================
    ligne = ImpTitreSect(wsFiche, ligne, "INFORMATIONS GENERALES")

    ligne = ImpLigneInfo(wsFiche, ligne, "ID", ws.Cells(LignePers, C_ID).Value, _
                          "Statut", ws.Cells(LignePers, C_STATUT).Value)
    ligne = ImpLigneInfo(wsFiche, ligne, "Nom", ws.Cells(LignePers, C_NOM).Value, _
                          "Type", ws.Cells(LignePers, C_TYPE).Value)

    If ws.Cells(LignePers, C_TYPE).Value = "Auxiliaire" Then
        ligne = ImpLigneInfo(wsFiche, ligne, "Groupe WE", ws.Cells(LignePers, C_GROUPE).Value, "", "")
    End If

    ligne = ligne + 1

    '==========================================================================
    ' HORAIRE
    '==========================================================================
    If avecHoraire Then
        ligne = ImpTitreSect(wsFiche, ligne, "HORAIRE HEBDOMADAIRE")

        Dim joursNoms(6)  As String
        Dim jCols(6)      As Long
        joursNoms(0) = "Lundi":    jCols(0) = C_LUN
        joursNoms(1) = "Mardi":    jCols(1) = C_MAR
        joursNoms(2) = "Mercredi": jCols(2) = C_MER
        joursNoms(3) = "Jeudi":    jCols(3) = C_JEU
        joursNoms(4) = "Vendredi": jCols(4) = C_VEN
        joursNoms(5) = "Samedi":   jCols(5) = C_SAM
        joursNoms(6) = "Dimanche": jCols(6) = C_DIM

        Dim codeLabel(3) As String
        ' On affiche M -> Matin, J -> Journee complete, - -> Absent
        Dim k As Integer
        For k = 0 To 6
            Dim cellHor As Range
            Dim colLet As String
            Select Case k
                Case 0: colLet = "A"
                Case 1: colLet = "B"
                Case 2: colLet = "C"
                Case 3: colLet = "D"
                Case 4: colLet = "A"
                Case 5: colLet = "B"
                Case 6: colLet = "C"
            End Select

            If k = 4 Then ligne = ligne + 0  ' Meme ligne pour Ven/Sam/Dim
            If k >= 4 And k = 4 Then ligne = ligne + 1

            ' Ligne 1 : Lun-Jeu sur ligne, puis Ven-Dim sur ligne suivante
        Next k

        ' Simplifie : affichage en tableau 2 lignes (labels + valeurs) sur 4 colonnes max
        ' Ligne des labels
        Dim rHorLbl As Range
        Set rHorLbl = wsFiche.Range("A" & ligne & ":D" & ligne)
        wsFiche.Cells(ligne, 1).Value = "Lundi"
        wsFiche.Cells(ligne, 2).Value = "Mardi"
        wsFiche.Cells(ligne, 3).Value = "Mercredi"
        wsFiche.Cells(ligne, 4).Value = "Jeudi"
        With rHorLbl
            .Interior.Color = IMP_GRIS_ENT
            .Font.Bold = True
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 18
        End With
        ligne = ligne + 1

        ' Ligne des valeurs
        Dim codeH As String
        Dim i As Integer
        For i = 0 To 3
            codeH = Trim(ws.Cells(LignePers, jCols(i)).Value)
            wsFiche.Cells(ligne, i + 1).Value = IIf(codeH = "M", "Matin", IIf(codeH = "J", "Journee", "Absent"))
        Next i
        With wsFiche.Range("A" & ligne & ":D" & ligne)
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 18
        End With
        ligne = ligne + 1

        ' 2e ligne horaire : Vendredi, Samedi, Dimanche
        wsFiche.Cells(ligne, 1).Value = "Vendredi"
        wsFiche.Cells(ligne, 2).Value = "Samedi"
        wsFiche.Cells(ligne, 3).Value = "Dimanche"
        wsFiche.Cells(ligne, 4).Value = ""
        With wsFiche.Range("A" & ligne & ":D" & ligne)
            .Interior.Color = IMP_GRIS_ENT
            .Font.Bold = True
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 18
        End With
        ligne = ligne + 1

        For i = 4 To 6
            codeH = Trim(ws.Cells(LignePers, jCols(i)).Value)
            wsFiche.Cells(ligne, i - 3).Value = IIf(codeH = "M", "Matin", IIf(codeH = "J", "Journee", "Absent"))
        Next i
        wsFiche.Cells(ligne, 4).Value = ""
        With wsFiche.Range("A" & ligne & ":D" & ligne)
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 18
        End With
        ligne = ligne + 2

    End If

    '==========================================================================
    ' APTITUDES
    '==========================================================================
    If avecAptitudes Then
        ligne = ImpTitreSect(wsFiche, ligne, "APTITUDES ET FORMATIONS")

        Dim aptNoms(4)  As String
        Dim aptCols(4)  As Long
        aptNoms(0) = "Poste 1 (Vaox5 A/B/C)": aptCols(0) = C_P1
        aptNoms(1) = "Poste 2 (Vaox5 A/B/C)": aptCols(1) = C_P2
        aptNoms(2) = "Poste 3 (Vaox5 A/B/C)": aptCols(2) = C_P3
        aptNoms(3) = "Vaox5-D":               aptCols(3) = C_D
        aptNoms(4) = "Vaox1 A&B":             aptCols(4) = C_VAOX1

        Dim a As Integer
        For a = 0 To 4 Step 2
            Dim colA As Integer
            wsFiche.Cells(ligne, 1).Value = aptNoms(a)
            wsFiche.Cells(ligne, 1).Font.Bold = True
            wsFiche.Cells(ligne, 1).Interior.Color = IMP_GRIS_LBL

            Dim valApt As Boolean
            valApt = (ws.Cells(LignePers, aptCols(a)).Value = 1)
            wsFiche.Cells(ligne, 2).Value = IIf(valApt, "Forme", "Non forme")
            wsFiche.Cells(ligne, 2).Font.Color = IIf(valApt, IMP_VERT, IMP_ROUGE)

            If a + 1 <= 4 Then
                wsFiche.Cells(ligne, 3).Value = aptNoms(a + 1)
                wsFiche.Cells(ligne, 3).Font.Bold = True
                wsFiche.Cells(ligne, 3).Interior.Color = IMP_GRIS_LBL
                valApt = (ws.Cells(LignePers, aptCols(a + 1)).Value = 1)
                wsFiche.Cells(ligne, 4).Value = IIf(valApt, "Forme", "Non forme")
                wsFiche.Cells(ligne, 4).Font.Color = IIf(valApt, IMP_VERT, IMP_ROUGE)
            End If

            With wsFiche.Range("A" & ligne & ":D" & ligne)
                .Font.Size = 10
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .RowHeight = 18
            End With
            ligne = ligne + 1
        Next a
        ligne = ligne + 1

    End If

    '==========================================================================
    ' STATISTIQUES
    '==========================================================================
    If avecStats Then
        ligne = ImpTitreSect(wsFiche, ligne, "STATISTIQUES DE PASSAGE PAR MACHINE")

        ' Needs 6 columns — on etale sur A:D en 2 lignes
        Dim statNoms(5) As String
        Dim statCols(5) As Long
        statNoms(0) = "Vaox5-A":    statCols(0) = C_STATS_A
        statNoms(1) = "Vaox5-B":    statCols(1) = C_STATS_B
        statNoms(2) = "Vaox5-C":    statCols(2) = C_STATS_C
        statNoms(3) = "Vaox5-D":    statCols(3) = C_STATS_D
        statNoms(4) = "Vaox1 A&B":  statCols(4) = C_STATS_V1
        statNoms(5) = "Expedition": statCols(5) = C_STATS_EXP

        ' Labels
        Dim s As Integer
        For s = 0 To 3
            wsFiche.Cells(ligne, s + 1).Value = statNoms(s)
            wsFiche.Cells(ligne, s + 1).Font.Bold = True
            wsFiche.Cells(ligne, s + 1).Interior.Color = IMP_GRIS_ENT
        Next s
        With wsFiche.Range("A" & ligne & ":D" & ligne)
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 18
        End With
        ligne = ligne + 1

        ' Valeurs
        For s = 0 To 3
            wsFiche.Cells(ligne, s + 1).Value = ws.Cells(LignePers, statCols(s)).Value
        Next s
        With wsFiche.Range("A" & ligne & ":D" & ligne)
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 18
        End With
        ligne = ligne + 1

        ' Vaox1 + Expedition sur 2e bloc
        wsFiche.Cells(ligne, 1).Value = statNoms(4)
        wsFiche.Cells(ligne, 1).Font.Bold = True
        wsFiche.Cells(ligne, 1).Interior.Color = IMP_GRIS_ENT
        wsFiche.Cells(ligne, 2).Value = statNoms(5)
        wsFiche.Cells(ligne, 2).Font.Bold = True
        wsFiche.Cells(ligne, 2).Interior.Color = IMP_GRIS_ENT
        With wsFiche.Range("A" & ligne & ":D" & ligne)
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 18
        End With
        ligne = ligne + 1

        wsFiche.Cells(ligne, 1).Value = ws.Cells(LignePers, statCols(4)).Value
        wsFiche.Cells(ligne, 2).Value = ws.Cells(LignePers, statCols(5)).Value
        With wsFiche.Range("A" & ligne & ":B" & ligne)
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 18
        End With
        ligne = ligne + 2

    End If

    '==========================================================================
    ' ABSENCES
    '==========================================================================
    If avecAbsences Then
        Dim wsV As Worksheet
        Set wsV = Sheets("Vacances")
        Dim dernLigneV As Long
        dernLigneV = wsV.Cells(wsV.Rows.count, 2).End(xlUp).Row

        ' Compte les absences de cette personne
        Dim nbAbs As Long
        nbAbs = 0
        Dim v As Long
        For v = 2 To dernLigneV
            If Trim(wsV.Cells(v, 2).Value) = nomPerso Then nbAbs = nbAbs + 1
        Next v

        If nbAbs > 0 Then
            ligne = ImpTitreSect(wsFiche, ligne, "ABSENCES ENREGISTREES")

            ' En-tete tableau absences
            wsFiche.Cells(ligne, 1).Value = "Type"
            wsFiche.Cells(ligne, 2).Value = "Date debut"
            wsFiche.Cells(ligne, 3).Value = "Date fin"
            wsFiche.Cells(ligne, 4).Value = "Duree"
            With wsFiche.Range("A" & ligne & ":D" & ligne)
                .Interior.Color = IMP_GRIS_ENT
                .Font.Bold = True
                .Font.Size = 10
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .RowHeight = 18
            End With
            ligne = ligne + 1

            For v = 2 To dernLigneV
                If Trim(wsV.Cells(v, 2).Value) = nomPerso Then
                    Dim dD As Date, dF As Date
                    On Error Resume Next
                    dD = wsV.Cells(v, 3).Value
                    dF = wsV.Cells(v, 4).Value
                    On Error GoTo 0

                    wsFiche.Cells(ligne, 1).Value = wsV.Cells(v, 5).Value
                    wsFiche.Cells(ligne, 2).Value = Format(dD, "dd.mm.yyyy")
                    wsFiche.Cells(ligne, 3).Value = Format(dF, "dd.mm.yyyy")
                    wsFiche.Cells(ligne, 4).Value = DateDiff("d", dD, dF) + 1 & " jour(s)"

                    With wsFiche.Range("A" & ligne & ":D" & ligne)
                        .Font.Size = 10
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                        .RowHeight = 18
                    End With
                    ligne = ligne + 1
                End If
            Next v
        End If
    End If

    '==========================================================================
    ' IMPRESSION
    '==========================================================================
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Active la feuille et lance l impression
    wsFiche.Activate

    If Application.Dialogs(xlDialogPrint).Show Then
        ' Impression lancee
    End If

    ' Supprime la feuille temporaire
    Application.DisplayAlerts = False
    wsFiche.Delete
    Application.DisplayAlerts = True

    ' Retourne sur le Planning journalier
    Sheets(NOM_FEUILLE_PLANNING).Activate

    Exit Sub

ErrGeneration:
    ' Restaure systématiquement l'état d'Excel avant de signaler l'erreur,
    ' pour ne jamais laisser le classeur figé (écran non rafraîchi,
    ' alertes désactivées) suite à un échec de génération de la fiche.
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    ThisWorkbook.Sheets("_Fiche_Temp").Delete
    On Error GoTo 0
    MsgBox "Erreur lors de la génération de la fiche :" & vbCrLf & Err.Description, vbCritical

End Sub

'==============================================================================
' FUNCTION : ImpTitreSect - Affiche un titre de section rouge
'==============================================================================
Private Function ImpTitreSect(ByVal ws As Worksheet, _
                                ByVal ligne As Long, _
                                ByVal titre As String) As Long
    With ws.Range("A" & ligne & ":D" & ligne)
        .Merge
        .Value = titre
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = IMP_ROUGE
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = IMP_ROUGE
        .Borders(xlEdgeBottom).Weight = xlMedium
        .RowHeight = 22
    End With
    ImpTitreSect = ligne + 1
End Function

'==============================================================================
' FUNCTION : ImpLigneInfo - Affiche une ligne label/valeur en 2 colonnes
'==============================================================================
Private Function ImpLigneInfo(ByVal ws As Worksheet, _
                                ByVal ligne As Long, _
                                ByVal lbl1 As String, ByVal val1 As String, _
                                ByVal lbl2 As String, ByVal val2 As String) As Long
    ws.Cells(ligne, 1).Value = lbl1
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Interior.Color = IMP_GRIS_LBL
    ws.Cells(ligne, 2).Value = val1

    If lbl2 <> "" Then
        ws.Cells(ligne, 3).Value = lbl2
        ws.Cells(ligne, 3).Font.Bold = True
        ws.Cells(ligne, 3).Interior.Color = IMP_GRIS_LBL
        ws.Cells(ligne, 4).Value = val2
    End If

    With ws.Range("A" & ligne & ":D" & ligne)
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .RowHeight = 18
    End With
    ImpLigneInfo = ligne + 1
End Function

'==============================================================================
' SUB : UF_InitCalendrierAux
' Initialise le calendrier auxiliaires (appele depuis UserForm_Initialize)
'==============================================================================
Sub UF_InitCalendrierAux(ByVal frm As Object)

    On Error Resume Next

    frm.txtAuxCalMois.Value = CStr(Month(Date))
    If Err.Number <> 0 Then MsgBox "txtAuxCalMois manquant": Err.Clear: Exit Sub

    frm.txtAuxCalAnnee.Value = CStr(Year(Date))
    If Err.Number <> 0 Then MsgBox "txtAuxCalAnnee manquant": Err.Clear: Exit Sub

    Dim dateRef As Date
    dateRef = Sheets("Planning journalier").Range("Ref_GroupeG1").Value

    frm.txtRefGroupeG1.Value = Format(dateRef, "dd.mm.yyyy")
    If Err.Number <> 0 Then MsgBox "txtRefGroupeG1 manquant": Err.Clear: Exit Sub

    frm.lblAuxCalTitre.Caption = ""
    If Err.Number <> 0 Then MsgBox "lblAuxCalTitre manquant": Err.Clear: Exit Sub

    On Error GoTo 0
    Call UF_DessinnerCalendrierAux(frm)

End Sub

'==============================================================================
' SUB : UF_DessinnerCalendrierAux
' Dessine le calendrier mensuel dans l onglet Auxiliaires
'==============================================================================
Sub UF_DessinnerCalendrierAux(ByVal frm As Object)

    Dim moisCal  As Long
    Dim anneeCal As Long
    moisCal = CLng(frm.txtAuxCalMois.Value)
    anneeCal = CLng(frm.txtAuxCalAnnee.Value)

    ' Titre du mois
    Dim noms(1 To 12) As String
    noms(1) = "Janvier":    noms(2) = "Fevrier":    noms(3) = "Mars"
    noms(4) = "Avril":      noms(5) = "Mai":         noms(6) = "Juin"
    noms(7) = "Juillet":    noms(8) = "Aout":        noms(9) = "Septembre"
    noms(10) = "Octobre":   noms(11) = "Novembre":   noms(12) = "Decembre"

    frm.lblAuxCalTitre.Caption = noms(moisCal) & " " & anneeCal

    ' Calcule la grille
    Dim premierJour As Date
    premierJour = DateSerial(anneeCal, moisCal, 1)
    Dim decalage As Long
    decalage = Weekday(premierJour, vbMonday) - 1
    Dim debutGrille As Date
    debutGrille = premierJour - decalage

    ' Dessine les 42 cases
    Dim i As Long
    For i = 0 To 41

        Dim dateCase As Date
        dateCase = debutGrille + i

        Dim lbl As MSForms.Label
        Set lbl = frm.Controls("lblAuxCal" & i)

        Dim estDuMois As Boolean
        estDuMois = (Month(dateCase) = moisCal And Year(dateCase) = anneeCal)

        Dim jourSem As Long
        jourSem = Weekday(dateCase, vbMonday)

        Dim estWE    As Boolean
        Dim estFerie As Boolean
        estWE = (jourSem = 6 Or jourSem = 7)
        estFerie = Module_Feries.estFerie(dateCase) And Not estWE

        ' Recupere le groupe depuis la feuille Groupes
        Dim groupeJour As String
        groupeJour = ""
        If estDuMois And (estWE Or estFerie) Then
            groupeJour = Module_Groupes.GetGroupeJour(dateCase)
        End If

        ' Caption : numero du jour + groupe si WE/ferie
        If estDuMois Then
            If groupeJour <> "" Then
                lbl.Caption = Day(dateCase) & vbCrLf & groupeJour
            Else
                lbl.Caption = CStr(Day(dateCase))
            End If
        Else
            lbl.Caption = CStr(Day(dateCase))
        End If

        lbl.Tag = CStr(CDbl(dateCase))

        ' Couleur
        Dim couleur As Long
        If Not estDuMois Then
            couleur = RGB(210, 210, 210)
            lbl.ForeColor = RGB(170, 170, 170)
        ElseIf estFerie And groupeJour = "G1" Then
            couleur = RGB(255, 200, 100)   ' Orange ferie
            lbl.ForeColor = RGB(0, 0, 128)
        ElseIf estFerie And groupeJour = "G2" Then
            couleur = RGB(255, 200, 100)   ' Orange ferie
            lbl.ForeColor = RGB(0, 100, 0)
        ElseIf estFerie Then
            couleur = RGB(255, 200, 100)
            lbl.ForeColor = RGB(100, 0, 0)
        ElseIf estWE And groupeJour = "G1" Then
            couleur = RGB(173, 216, 230)   ' Bleu clair G1
            lbl.ForeColor = RGB(0, 0, 128)
        ElseIf estWE And groupeJour = "G2" Then
            couleur = RGB(144, 238, 144)   ' Vert clair G2
            lbl.ForeColor = RGB(0, 100, 0)
        ElseIf estWE Then
            couleur = RGB(220, 220, 220)   ' WE sans groupe
            lbl.ForeColor = RGB(0, 0, 0)
        Else
            couleur = RGB(255, 255, 255)   ' Jour normal
            lbl.ForeColor = RGB(0, 0, 0)
        End If

        lbl.BackColor = couleur

        ' Encadre rouge pour aujourd hui
        If Int(dateCase) = Int(Date) And estDuMois Then
            lbl.BorderColor = RGB(200, 0, 0)
            lbl.Font.Bold = True
        Else
            lbl.BorderColor = RGB(200, 200, 200)
            lbl.Font.Bold = (groupeJour <> "")
        End If

        ' Tooltip
        If estFerie And estDuMois Then
            lbl.ControlTipText = Module_Feries.GetNomFerie(dateCase) & _
                                 IIf(groupeJour <> "", " - " & groupeJour, "")
        ElseIf estWE And estDuMois Then
            lbl.ControlTipText = IIf(groupeJour <> "", groupeJour, "Weekend sans attribution")
        Else
            lbl.ControlTipText = Format(dateCase, "dd.mm.yyyy")
        End If

    Next i

End Sub

'==============================================================================
' SUB : UF_AuxCalMoisPrecedent / Suivant
'==============================================================================
Sub UF_AuxCalMoisPrecedent(ByVal frm As Object)
    Dim mois As Long: mois = CLng(frm.txtAuxCalMois.Value)
    Dim annee As Long: annee = CLng(frm.txtAuxCalAnnee.Value)
    mois = mois - 1
    If mois < 1 Then mois = 12: annee = annee - 1
    frm.txtAuxCalMois.Value = CStr(mois)
    frm.txtAuxCalAnnee.Value = CStr(annee)
    Call UF_DessinnerCalendrierAux(frm)
End Sub

Sub UF_AuxCalMoisSuivant(ByVal frm As Object)
    Dim mois As Long: mois = CLng(frm.txtAuxCalMois.Value)
    Dim annee As Long: annee = CLng(frm.txtAuxCalAnnee.Value)
    mois = mois + 1
    If mois > 12 Then mois = 1: annee = annee + 1
    frm.txtAuxCalMois.Value = CStr(mois)
    frm.txtAuxCalAnnee.Value = CStr(annee)
    Call UF_DessinnerCalendrierAux(frm)
End Sub

'==============================================================================
' SUB : UF_ClicCalendrierAux
' Clic sur une case du calendrier auxiliaire
' Ouvre UserForm_GroupeAux pour WE et jours feries uniquement
'==============================================================================
Sub UF_ClicCalendrierAux(ByVal frm As Object, ByVal indexJour As Integer)

    Dim lbl As MSForms.Label
    Set lbl = frm.Controls("lblAuxCal" & indexJour)

    If lbl.Tag = "" Then Exit Sub

    Dim dateJour As Date
    dateJour = CDate(CDbl(lbl.Tag))

    ' Verifie que le jour est dans le mois affiche
    Dim moisCal  As Long
    Dim anneeCal As Long
    moisCal = CLng(frm.txtAuxCalMois.Value)
    anneeCal = CLng(frm.txtAuxCalAnnee.Value)

    If Month(dateJour) <> moisCal Or Year(dateJour) <> anneeCal Then
        frm.txtAuxCalMois.Value = CStr(Month(dateJour))
        frm.txtAuxCalAnnee.Value = CStr(Year(dateJour))
        Call UF_DessinnerCalendrierAux(frm)
        Exit Sub
    End If

    ' Uniquement pour les weekends et jours feries
    Dim jourSem As Long
    jourSem = Weekday(dateJour, vbMonday)
    Dim estWE    As Boolean
    Dim estFerie As Boolean
    estWE = (jourSem = 6 Or jourSem = 7)
    estFerie = Module_Feries.estFerie(dateJour) And Not estWE

    If Not (estWE Or estFerie) Then Exit Sub

    ' Ouvre le UserForm d attribution
    Dim dlg As UserForm_GroupeAux
    Set dlg = New UserForm_GroupeAux
    dlg.dateJour = dateJour
    dlg.Initialiser
    dlg.Show

    If Not dlg.Confirme Then
        Unload dlg
        Exit Sub
    End If

    Dim Action As String
    Action = dlg.Action
    Unload dlg

    Select Case Action
        Case "G1":        Call Module_Groupes.SetGroupeJour(dateJour, "G1")
        Case "G2":        Call Module_Groupes.SetGroupeJour(dateJour, "G2")
        Case "Supprimer": Call Module_Groupes.SetGroupeJour(dateJour, "")
    End Select

    ' Redessine le calendrier
    Call UF_DessinnerCalendrierAux(frm)

End Sub

'==============================================================================
' SUB : UF_EnregistrerRefGroupeG1
' Enregistre la nouvelle date de reference G1 depuis le UserForm
'==============================================================================
Sub UF_EnregistrerRefGroupeG1(ByVal frm As Object)

    Dim valeur As String
    valeur = Trim(frm.txtRefGroupeG1.Value)

    If valeur = "" Then
        MsgBox "Veuillez saisir une date.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(valeur) Then
        MsgBox "Date invalide. Format attendu : jj.mm.aaaa", vbExclamation
        Exit Sub
    End If

    Dim nouvelleDate As Date
    nouvelleDate = CDate(valeur)

    ' Verifie que c est bien un lundi
    If Weekday(nouvelleDate, vbMonday) <> 1 Then
        MsgBox "La date de référence doit être un lundi.", vbExclamation
        Exit Sub
    End If

    ' Met a jour la cellule dans Planning journalier
    Sheets("Planning journalier").Range("Ref_GroupeG1").Value = nouvelleDate

    MsgBox "Date de référence G1 mise a jour : " & Format(nouvelleDate, "dd.mm.yyyy"), _
           vbInformation, "Groupes v2.0"

    ' Redessine le calendrier
    Call UF_DessinnerCalendrierAux(frm)

End Sub

'==============================================================================
' SUB : UF_ChargerGroupesAnnee
' Genere les groupes d une annee depuis le UserForm
'==============================================================================
Sub UF_ChargerGroupesAnnee(ByVal frm As Object)

    Dim annee As Long
    annee = CLng(frm.lblAnneeCharge.Caption)


    Dim rep As VbMsgBoxResult
    rep = MsgBox("Générer les groupes " & annee & " ?" & vbCrLf & _
                 "(Les modifications manuelles seront conservées)", _
                 vbYesNo + vbQuestion, "Générer les groupes")

    If rep = vbNo Then Exit Sub

    Call Module_Groupes.GenererGroupesAnnee(annee)

    ' Navigue vers l annee generee
    frm.txtAuxCalAnnee.Value = CStr(annee)
    frm.txtAuxCalMois.Value = "1"
    Call UF_DessinnerCalendrierAux(frm)

End Sub
'==============================================================================
' SUB : UF_AjouterFermeture
' Ajoute une periode de fermeture entreprise dans la feuille Feries
' Appelee depuis un bouton dans l onglet Calendrier
'==============================================================================
Sub UF_AjouterFermeture(ByVal frm As Object)

    ' Saisie date debut
    Dim strDebut As String
    strDebut = Trim(frm.txtFermDebut.Value)

    If strDebut = "" Then
        MsgBox "Veuillez saisir une date de début.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(strDebut) Then
        MsgBox "Date de début invalide." & vbCrLf & "Format attendu : jj.mm.aaaa", vbExclamation
        Exit Sub
    End If

    ' Saisie date fin
    Dim strFin As String
    strFin = Trim(frm.txtFermFin.Value)

    If strFin = "" Then
        MsgBox "Veuillez saisir une date de fin.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(strFin) Then
        MsgBox "Date de fin invalide." & vbCrLf & "Format attendu : jj.mm.aaaa", vbExclamation
        Exit Sub
    End If

    Dim dDebut As Date
    Dim dFin   As Date
    dDebut = CDate(strDebut)
    dFin = CDate(strFin)

    If dFin < dDebut Then
        MsgBox "La date de fin ne peut pas être antérieure à la date de début.", vbExclamation
        Exit Sub
    End If

    ' Saisie nom optionnel
    Dim nomFermeture As String
    nomFermeture = Trim(frm.txtFermNom.Value)
    If nomFermeture = "" Then nomFermeture = "Fermeture entreprise"

    ' Confirmation
    Dim nbJours As Long
    nbJours = DateDiff("d", dDebut, dFin) + 1

    Dim rep As VbMsgBoxResult
    rep = MsgBox("Ajouter la fermeture suivante ?" & vbCrLf & vbCrLf & _
                 "Nom    : " & nomFermeture & vbCrLf & _
                 "Du     : " & Format(dDebut, "dd.mm.yyyy") & vbCrLf & _
                 "Au     : " & Format(dFin, "dd.mm.yyyy") & vbCrLf & _
                 "Durée  : " & nbJours & " jour(s)", _
                 vbYesNo + vbQuestion, "Confirmer la fermeture")

    If rep = vbNo Then Exit Sub

    ' Enregistre chaque jour dans la feuille Feries
    Call Module_Feries.AjouterFermeturePeriode(dDebut, dFin, nomFermeture)

    MsgBox "Fermeture d'entreprise enregistrée : " & nbJours & " jour(s) ajouté(s).", _
           vbInformation, "Fermeture entreprise"

    ' Vide les champs après enregistrement
    frm.txtFermDebut.Value = ""
    frm.txtFermFin.Value = ""
    frm.txtFermNom.Value = ""

End Sub

'==============================================================================
' SUB : UF_SupprimerFermeture
' Supprime toutes les fermetures d une annee donnee
'==============================================================================
Sub UF_SupprimerFermeturesAnnee(ByVal frm As Object, ByVal annee As Long)

    Dim rep As VbMsgBoxResult
    rep = MsgBox("Supprimer toutes les fermetures de " & annee & " ?", _
                 vbYesNo + vbExclamation, "Supprimer fermetures")

    If rep = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_FERIES)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(NOM_TBL_FERIES)

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    Dim i As Long
    For i = dernLigne To 2 Step -1
        Dim d As Date
        d = 0
        On Error Resume Next
        d = ws.Cells(i, 1).Value
        On Error GoTo 0
        If d <> 0 Then
            If Year(d) = annee And Trim(ws.Cells(i, 3).Value) = TYPE_FERIE_FERMETURE Then
                ' Suppression via ListObject (convention du projet) au lieu
                ' de ws.Rows(i).Delete, qui efface toute la largeur de la ligne.
                tbl.ListRows(i - tbl.HeaderRowRange.Row).Delete
            End If
        End If
    Next i

    MsgBox "Fermetures " & annee & " supprimées.", vbInformation, "Fermetures"

End Sub

