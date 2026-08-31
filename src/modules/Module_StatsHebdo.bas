Option Explicit

'==============================================================================
' MODULE : Module_StatsHebdo
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'
' Description :
'   Génération automatique d'un PDF des statistiques de passages par machine
'   pour tout le personnel actif. Déclenché automatiquement chaque lundi
'   lors de l'impression du planning.
'
'   Le PDF est archivé dans :
'     Archives\Stats\Stats_semaine_aaaa.mm.jj.pdf
'
'   Tri : Personnel Fixe (alpha) -> Auxiliaires G1 (alpha) -> Auxiliaires G2 (alpha)
'   Colonnes : Code opérateur, Nom, Vaox5-A, Vaox5-B, Vaox5-C, Vaox5-D, Vaox1, Expédition
'==============================================================================

' Colonnes de la feuille Personnel : PERS_COL_xxx (Module_Constantes), pas de
' redéfinition locale ici — une redéfinition (ex-S_COL_xxx) désynchronise
' silencieusement ce module si la feuille Personnel est réorganisée un jour.

' Couleurs mise en forme
Private Const ROUGE    As Long = 12611584  ' RGB(192, 0, 0)
Private Const GRIS_ENT As Long = 14277081  ' RGB(217, 217, 217)
Private Const GRIS_FIX As Long = 13434828  ' Vert très clair — Fixe
Private Const GRIS_G1  As Long = 11393254  ' Bleu clair — G1
Private Const GRIS_G2  As Long = 9498256   ' Vert clair — G2

'==============================================================================
' SUB : ImprimerStatsHebdo
' Point d'entrée — appelée depuis Commande_Bouton.ImprimerPlanning()
' uniquement si le jour d'impression est un lundi.
'==============================================================================
Public Sub ImprimerStatsHebdo(ByVal datePlanning As Date)

    ' Vérifie que c'est bien un lundi
    If Weekday(datePlanning, vbMonday) <> 1 Then Exit Sub

    If ThisWorkbook.Path = "" Then
        MsgBox "Le classeur doit être enregistré avant de pouvoir archiver " & _
               "les statistiques hebdomadaires.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrDossier

    ' Vérifie que le dossier Archives\Stats existe
    Dim dossierArchives As String
    dossierArchives = ThisWorkbook.Path & "\" & NOM_DOSSIER_ARCHIVES & "\"

    Dim dossierStats As String
    dossierStats = dossierArchives & NOM_SOUS_DOSSIER_STATS & "\"

    If Dir(dossierArchives, vbDirectory) = "" Then
        MkDir dossierArchives
    End If
    If Dir(dossierStats, vbDirectory) = "" Then
        MkDir dossierStats
    End If

    On Error GoTo 0

    ' Nom du fichier
    Dim nomFichier As String
    nomFichier = "Stats_semaine_" & Format(datePlanning, "yyyy.mm.dd") & ".pdf"

    Dim cheminFinal As String
    cheminFinal = dossierStats & nomFichier

    ' Génère la feuille temporaire et exporte en PDF
    Call GenererFeuilleStats(cheminFinal, datePlanning)

    Exit Sub

ErrDossier:
    MsgBox "Impossible de créer le dossier d'archivage des statistiques :" & _
           vbCrLf & Err.Description, vbExclamation

End Sub

'==============================================================================
' SUB : GenererFeuilleStats (privée)
' Crée la feuille temporaire _Stats_Temp, la met en forme et exporte en PDF.
'==============================================================================
Private Sub GenererFeuilleStats(ByVal cheminFinal As String, _
                                  ByVal datePlanning As Date)

    ' Gestion d'erreur globale : sans elle, un échec (permissions sur
    ' Archives\Stats, PDF déjà ouvert, disque plein) laisse la feuille
    ' temporaire _Stats_Temp visible dans le classeur et l'écran figé
    ' (ScreenUpdating/DisplayAlerts jamais restaurés).
    On Error GoTo ErrGeneration

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' Supprime l'ancienne feuille temporaire si elle existe
    On Error Resume Next
    ThisWorkbook.Sheets("_Stats_Temp").Delete
    On Error GoTo 0

    ' Crée la feuille temporaire
    Dim wsStat As Worksheet
    Set wsStat = ThisWorkbook.Sheets.Add( _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsStat.Name = "_Stats_Temp"

    ' Configuration de la page
    With wsStat.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
    End With

    ' Largeurs des colonnes
    wsStat.Columns("A").ColumnWidth = 12   ' Code
    wsStat.Columns("B").ColumnWidth = 26   ' Nom
    wsStat.Columns("C").ColumnWidth = 10   ' Vaox5-A
    wsStat.Columns("D").ColumnWidth = 10   ' Vaox5-B
    wsStat.Columns("E").ColumnWidth = 10   ' Vaox5-C
    wsStat.Columns("F").ColumnWidth = 10   ' Vaox5-D
    wsStat.Columns("G").ColumnWidth = 10   ' Vaox1
    wsStat.Columns("H").ColumnWidth = 12   ' Expédition

    Dim ligne As Long
    ligne = 1

    ' --- TITRE ---
    With wsStat.Range("A" & ligne & ":H" & ligne)
        .Merge
        .Value = "Statistiques de passages par machine — Semaine du " & _
                              Format(datePlanning, "dd.mm.yyyy")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = ROUGE
        .HorizontalAlignment = xlCenter
        .RowHeight = 26
    End With
    ligne = ligne + 1

    ' Sous-titre date d'édition
    With wsStat.Range("A" & ligne & ":H" & ligne)
        .Merge
        .Value = "Édité le " & Format(Now, "dd.mm.yyyy à hh:mm")
        .Font.Size = 9
        .Font.Color = RGB(128, 128, 128)
        .HorizontalAlignment = xlRight
        .RowHeight = 16
    End With
    ligne = ligne + 1

    ' Ligne rouge séparatrice
    With wsStat.Range("A" & ligne & ":H" & ligne)
        .Merge
        .Interior.Color = ROUGE
        .RowHeight = 3
    End With
    ligne = ligne + 1

    ' --- EN-TÊTE DU TABLEAU ---
    Dim enTetes(1 To 8) As String
    enTetes(1) = "Code":       enTetes(2) = "Nom"
    enTetes(3) = "Vaox5-A":   enTetes(4) = "Vaox5-B"
    enTetes(5) = "Vaox5-C":   enTetes(6) = "Vaox5-D"
    enTetes(7) = "Vaox1 A&B": enTetes(8) = "Expedition"

    Dim col As Integer
    For col = 1 To 8
        With wsStat.Cells(ligne, col)
            .Value = enTetes(col)
            .Font.Bold = True
            .Font.Size = 10
            .Interior.Color = GRIS_ENT
            .HorizontalAlignment = IIf(col <= 2, xlLeft, xlCenter)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next col
    wsStat.Rows(ligne).RowHeight = 18
    ligne = ligne + 1

    ' --- CHARGEMENT DU PERSONNEL ---
    Dim wsPer As Worksheet
    Set wsPer = Sheets(NOM_FEUILLE_PERSONNEL)

    Dim dernLigne As Long
    dernLigne = wsPer.Cells(wsPer.Rows.count, PERS_COL_NOM).End(xlUp).Row

    ' Collecte : 3 groupes — Fixe, G1, G2
    Dim nomsFixe()  As String: Dim lignesFixe()   As Long
    Dim nomsG1()    As String: Dim lignesG1()     As Long
    Dim nomsG2()    As String: Dim lignesG2()     As Long
    ReDim nomsFixe(1 To dernLigne): ReDim lignesFixe(1 To dernLigne)
    ReDim nomsG1(1 To dernLigne): ReDim lignesG1(1 To dernLigne)
    ReDim nomsG2(1 To dernLigne): ReDim lignesG2(1 To dernLigne)

    Dim nbFixe As Long: nbFixe = 0
    Dim nbG1   As Long: nbG1 = 0
    Dim nbG2   As Long: nbG2 = 0

    Dim i As Long
    For i = 2 To dernLigne
        If Trim(wsPer.Cells(i, PERS_COL_STATUT).Value) <> "Actif" Then GoTo SuivantPers

        Dim typePers   As String
        Dim groupePers As String
        typePers = Trim(wsPer.Cells(i, PERS_COL_TYPE).Value)
        groupePers = Trim(wsPer.Cells(i, PERS_COL_GROUPE).Value)

        If typePers = "Fixe" Then
            nbFixe = nbFixe + 1
            nomsFixe(nbFixe) = wsPer.Cells(i, PERS_COL_NOM).Value
            lignesFixe(nbFixe) = i
        ElseIf typePers = "Auxiliaire" And groupePers = GROUPE_G1 Then
            nbG1 = nbG1 + 1
            nomsG1(nbG1) = wsPer.Cells(i, PERS_COL_NOM).Value
            lignesG1(nbG1) = i
        ElseIf typePers = "Auxiliaire" And groupePers = GROUPE_G2 Then
            nbG2 = nbG2 + 1
            nomsG2(nbG2) = wsPer.Cells(i, PERS_COL_NOM).Value
            lignesG2(nbG2) = i
        End If

SuivantPers:
    Next i

    ' Tri alphabétique dans chaque groupe (tri à bulles)
    Call TrierGroupeStats(nomsFixe, lignesFixe, nbFixe)
    Call TrierGroupeStats(nomsG1, lignesG1, nbG1)
    Call TrierGroupeStats(nomsG2, lignesG2, nbG2)

    ' --- ÉCRITURE DES LIGNES ---
    ' Groupe Fixe
    If nbFixe > 0 Then
        ligne = EcrireSousTitre(wsStat, ligne, "Personnel Fixe", RGB(204, 255, 204))
        Dim k As Long
        For k = 1 To nbFixe
            ligne = EcrireLigneStat(wsStat, ligne, wsPer, lignesFixe(k), RGB(240, 255, 240))
        Next k
    End If

    ' Groupe G1
    If nbG1 > 0 Then
        ligne = EcrireSousTitre(wsStat, ligne, "Auxiliaires — Groupe 1", RGB(173, 216, 230))
        For k = 1 To nbG1
            ligne = EcrireLigneStat(wsStat, ligne, wsPer, lignesG1(k), RGB(235, 245, 255))
        Next k
    End If

    ' Groupe G2
    If nbG2 > 0 Then
        ligne = EcrireSousTitre(wsStat, ligne, "Auxiliaires — Groupe 2", RGB(144, 238, 144))
        For k = 1 To nbG2
            ligne = EcrireLigneStat(wsStat, ligne, wsPer, lignesG2(k), RGB(235, 255, 235))
        Next k
    End If

    ' --- EXPORT PDF ---
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    wsStat.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=cheminFinal, _
        Quality:=xlQualityStandard

    ' Supprime la feuille temporaire
    Application.DisplayAlerts = False
    wsStat.Delete
    Application.DisplayAlerts = True

    ' Retourne sur le planning
    Sheets(NOM_FEUILLE_PLANNING).Activate

    Exit Sub

ErrGeneration:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    ThisWorkbook.Sheets("_Stats_Temp").Delete
    On Error GoTo 0
    MsgBox "Erreur lors de l'export des statistiques hebdomadaires :" & _
           vbCrLf & Err.Description, vbExclamation

End Sub

'==============================================================================
' FUNCTION : EcrireSousTitre (privée)
' Écrit une ligne de sous-titre de groupe avec couleur de fond.
'==============================================================================
Private Function EcrireSousTitre(ByVal ws As Worksheet, _
                                   ByVal ligne As Long, _
                                   ByVal titre As String, _
                                   ByVal couleur As Long) As Long
    With ws.Range("A" & ligne & ":H" & ligne)
        .Merge
        .Value = titre
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = couleur
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .RowHeight = 18
    End With
    EcrireSousTitre = ligne + 1
End Function

'==============================================================================
' FUNCTION : EcrireLigneStat (privée)
' Écrit une ligne de statistiques pour une personne.
'==============================================================================
Private Function EcrireLigneStat(ByVal wsStat As Worksheet, _
                                   ByVal ligne As Long, _
                                   ByVal wsPer As Worksheet, _
                                   ByVal lignePer As Long, _
                                   ByVal couleur As Long) As Long

    Dim valeurs(1 To 8) As Variant
    valeurs(1) = wsPer.Cells(lignePer, PERS_COL_ID).Value
    valeurs(2) = wsPer.Cells(lignePer, PERS_COL_NOM).Value
    valeurs(3) = wsPer.Cells(lignePer, PERS_COL_STATS_A).Value
    valeurs(4) = wsPer.Cells(lignePer, PERS_COL_STATS_B).Value
    valeurs(5) = wsPer.Cells(lignePer, PERS_COL_STATS_C).Value
    valeurs(6) = wsPer.Cells(lignePer, PERS_COL_STATS_D).Value
    valeurs(7) = wsPer.Cells(lignePer, PERS_COL_STATS_V1).Value
    valeurs(8) = wsPer.Cells(lignePer, PERS_COL_STATS_EXP).Value

    Dim col As Integer
    For col = 1 To 8
        With wsStat.Cells(ligne, col)
            .Value = valeurs(col)
            .Font.Size = 10
            .Interior.Color = couleur
            .HorizontalAlignment = IIf(col <= 2, xlLeft, xlCenter)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next col

    wsStat.Rows(ligne).RowHeight = 17
    EcrireLigneStat = ligne + 1

End Function

'==============================================================================
' SUB : TrierGroupeStats (privée)
' Tri à bulles alphabétique sur un groupe de noms + lignes associées.
'==============================================================================
Private Sub TrierGroupeStats(ByRef noms() As String, _
                               ByRef lignes() As Long, _
                               ByVal nb As Long)
    Dim a As Long, b As Long
    Dim tmpNom As String: Dim tmpLigne As Long

    For a = 1 To nb - 1
        For b = 1 To nb - a
            If noms(b) > noms(b + 1) Then
                tmpNom = noms(b):      noms(b) = noms(b + 1):     noms(b + 1) = tmpNom
                tmpLigne = lignes(b): lignes(b) = lignes(b + 1): lignes(b + 1) = tmpLigne
            End If
        Next b
    Next a

End Sub