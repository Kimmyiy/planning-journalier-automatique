Option Explicit

'==============================================================================
' MODULE : Module_Constantes
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'
' Description :
'   Centralise toutes les constantes du projet.
'   Noms de feuilles, tableaux, cellules nommées, types et couleurs.
'
'   UTILISATION :
'     Remplacer partout les chaînes littérales par ces constantes.
'     Ex: Sheets("Personnel")           -> Sheets(NOM_FEUILLE_PERSONNEL)
'         ws.ListObjects("Tbl_Param")   -> ws.ListObjects(NOM_TBL_PARAMETRES)
'==============================================================================

'------------------------------------------------------------------------------
' NOMS DES FEUILLES EXCEL
'------------------------------------------------------------------------------
Public Const NOM_FEUILLE_PLANNING       As String = "Planning journalier"
Public Const NOM_FEUILLE_PERSONNEL      As String = "Personnel"
Public Const NOM_FEUILLE_VACANCES       As String = "Vacances"
Public Const NOM_FEUILLE_FERIES         As String = "Feries"
Public Const NOM_FEUILLE_REMPLACEMENTS  As String = "Remplacements"
Public Const NOM_FEUILLE_GROUPES        As String = "Groupes"
Public Const NOM_FEUILLE_PARAMETRES     As String = "Parametres"
Public Const NOM_FEUILLE_LISTES         As String = "Listes"
Public Const NOM_FEUILLE_INDEX          As String = "Index des versions"

'------------------------------------------------------------------------------
' NOMS DES TABLEAUX EXCEL (ListObjects)
'------------------------------------------------------------------------------
Public Const NOM_TBL_PERSONNEL          As String = "Tbl_Personnel"
Public Const NOM_TBL_VACANCES           As String = "Tbl_Vacances"
Public Const NOM_TBL_FERIES             As String = "Tbl_Feries"
Public Const NOM_TBL_REMPLACEMENTS      As String = "Tbl_Remplacements"
Public Const NOM_TBL_GROUPES            As String = "Dates_Groupes"
Public Const NOM_TBL_PARAMETRES         As String = "Tbl_Parametres"

'------------------------------------------------------------------------------
' CELLULES NOMMEES — Feuille Planning journalier
'------------------------------------------------------------------------------
Public Const CELL_DATE_PLANNING         As String = "Date_planning"
Public Const CELL_DATE_PRINT            As String = "Date_print"
Public Const CELL_MAJ_STATS             As String = "MAJ_Stats"
Public Const CELL_REF_GROUPE_G1         As String = "Ref_GroupeG1"
Public Const CELL_QTE_VAOX5A            As String = "EnEpreuve_VaoX5A"
Public Const CELL_QTE_VAOX5B            As String = "EnEpreuve_VaoX5B"
Public Const CELL_QTE_VAOX5C            As String = "EnEpreuve_VaoX5C"
Public Const CELL_QTE_VAOX5D            As String = "EnEpreuve_VaoX5D"

'------------------------------------------------------------------------------
' TYPES D'ABSENCES
'------------------------------------------------------------------------------
Public Const TYPE_ABS_VACANCES          As String = "Vacances"
Public Const TYPE_ABS_MALADIE           As String = "Maladie"
Public Const TYPE_ABS_CONGE             As String = "Conge"

'------------------------------------------------------------------------------
' TYPES DE FERIES
'------------------------------------------------------------------------------
Public Const TYPE_FERIE_FIXE            As String = "Fixe"
Public Const TYPE_FERIE_CALCULE         As String = "Calcule"
Public Const TYPE_FERIE_MANUEL          As String = "Manuel"
Public Const TYPE_FERIE_FERMETURE       As String = "Fermeture"

'------------------------------------------------------------------------------
' GROUPES AUXILIAIRES
'------------------------------------------------------------------------------
Public Const GROUPE_G1                  As String = "G1"
Public Const GROUPE_G2                  As String = "G2"

'------------------------------------------------------------------------------
' NOMS DES MACHINES
' MACHINE_xxx : libellé tel qu'affiché dans le planning journalier
' PARAM_VAOX1 : clé de recherche dans Tbl_Parametres (colonne A) pour Vaox1
' Les deux graphies historiques de Vaox1 A&B ("Vaox1-A&B" affichage machine
' vs "Vaox1 A&B" planning) sont centralisées ici pour éviter toute divergence.
'------------------------------------------------------------------------------
Public Const MACHINE_VAOX1              As String = "Vaox1 A&B"
Public Const PARAM_VAOX1                As String = "Vaox1-A&B"
Public Const MACHINE_VAOX5A             As String = "Vaox5-A"
Public Const MACHINE_VAOX5B             As String = "Vaox5-B"
Public Const MACHINE_VAOX5C             As String = "Vaox5-C"
Public Const MACHINE_VAOX5D             As String = "Vaox5-D"

'------------------------------------------------------------------------------
' DOSSIERS D'ARCHIVAGE (PDF planning, sauvegardes, stats hebdo)
'------------------------------------------------------------------------------
Public Const NOM_DOSSIER_ARCHIVES       As String = "Archives"
Public Const NOM_SOUS_DOSSIER_STATS     As String = "Stats"

'------------------------------------------------------------------------------
' COLONNES DE LA FEUILLE Personnel (tableau Tbl_Personnel)
' Utilisées par Planning_Auto, Module_UserForm et Module_StatsHebdo :
' source UNIQUE pour éviter toute désynchronisation entre modules.
'------------------------------------------------------------------------------
Public Const PERS_COL_ID         As Long = 1   ' A — Identifiant
Public Const PERS_COL_NOM        As Long = 2   ' B — Nom
Public Const PERS_COL_STATUT     As Long = 3   ' C — Statut (Actif / Archive)
Public Const PERS_COL_TYPE       As Long = 4   ' D — Type (Fixe / Auxiliaire)
Public Const PERS_COL_GROUPE     As Long = 5   ' E — Groupe WE (G1 / G2)
Public Const PERS_COL_LUN        As Long = 6   ' F — Horaire lundi
Public Const PERS_COL_MAR        As Long = 7   ' G — Horaire mardi
Public Const PERS_COL_MER        As Long = 8   ' H — Horaire mercredi
Public Const PERS_COL_JEU        As Long = 9   ' I — Horaire jeudi
Public Const PERS_COL_VEN        As Long = 10  ' J — Horaire vendredi
Public Const PERS_COL_SAM        As Long = 11  ' K — Horaire samedi
Public Const PERS_COL_DIM        As Long = 12  ' L — Horaire dimanche
Public Const PERS_COL_P1         As Long = 13  ' M — Aptitude Poste 1
Public Const PERS_COL_P2         As Long = 14  ' N — Aptitude Poste 2
Public Const PERS_COL_P3         As Long = 15  ' O — Aptitude Poste 3
Public Const PERS_COL_D          As Long = 16  ' P — Aptitude Vaox5-D
Public Const PERS_COL_VAOX1      As Long = 17  ' Q — Aptitude Vaox1 A&B
Public Const PERS_COL_STATS_A    As Long = 18  ' R — Passages Vaox5-A
Public Const PERS_COL_STATS_B    As Long = 19  ' S — Passages Vaox5-B
Public Const PERS_COL_STATS_C    As Long = 20  ' T — Passages Vaox5-C
Public Const PERS_COL_STATS_D    As Long = 21  ' U — Passages Vaox5-D
Public Const PERS_COL_STATS_V1   As Long = 22  ' V — Passages Vaox1
Public Const PERS_COL_STATS_EXP  As Long = 23  ' W — Passages Expédition

'------------------------------------------------------------------------------
' COLONNES DU TABLEAU Tbl_Vacances (feuille Vacances)
'------------------------------------------------------------------------------
Public Const VAC_COL_ID          As Long = 1
Public Const VAC_COL_NOM         As Long = 2
Public Const VAC_COL_DEBUT       As Long = 3
Public Const VAC_COL_FIN         As Long = 4
Public Const VAC_COL_TYPE        As Long = 5

'------------------------------------------------------------------------------
' COLONNES DU TABLEAU Tbl_Remplacements (feuille Remplacements)
'------------------------------------------------------------------------------
Public Const RPL_COL_DATE           As Long = 1
Public Const RPL_COL_ID_ABSENTE     As Long = 2
Public Const RPL_COL_NOM_ABSENTE    As Long = 3
Public Const RPL_COL_ID_REMPLACANT  As Long = 4
Public Const RPL_COL_NOM_REMPLACANT As Long = 5
Public Const RPL_COL_TYPE           As Long = 6

'------------------------------------------------------------------------------
' COLONNES DU TABLEAU Tbl_Parametres (feuille Parametres)
'------------------------------------------------------------------------------
Public Const PARAM_MACHINE       As Long = 1  ' A — Nom de la machine
Public Const PARAM_SEUIL1        As Long = 2  ' B — Seuil 2 personnes
Public Const PARAM_SEUIL2        As Long = 3  ' C — Seuil 3 personnes
Public Const PARAM_MAX           As Long = 4  ' D — Quantité max
Public Const PARAM_FERMETURE     As Long = 5  ' E — Fermeture (VRAI/FAUX)

'------------------------------------------------------------------------------
' MISE EN PAGE — Feuille Planning journalier (tableaux Planning_Matin/AM)
'------------------------------------------------------------------------------
Public Const PLAN_COL_NOMS       As Long = 6   ' F — Colonne des noms affectés
Public Const PLAN_COL_POSTES     As Long = 7   ' G — Colonne des postes
Public Const LIGNE_DEBUT_MATIN   As Long = 6
Public Const LIGNE_FIN_MATIN     As Long = 20
Public Const LIGNE_DEBUT_AM      As Long = 24
Public Const LIGNE_FIN_AM        As Long = 27
Public Const NB_LIGNES_MATIN     As Long = 15
Public Const NB_LIGNES_AM        As Long = 3

'------------------------------------------------------------------------------
' COULEURS CALENDRIER (valeurs Long VBA)
' Note : VBA BackColor utilise le format RGB(r, g, b)
'------------------------------------------------------------------------------
Public Const CAL_COULEUR_G1             As Long = 11393254  ' Bleu clair    RGB(173, 216, 230)
Public Const CAL_COULEUR_G2             As Long = 9498256   ' Vert clair    RGB(144, 238, 144)
Public Const CAL_COULEUR_FERIE          As Long = 10079487  ' Orange clair  RGB(255, 204, 102)
Public Const CAL_COULEUR_FERMETURE      As Long = 255       ' Rouge vif     RGB(255,   0,   0)
Public Const CAL_COULEUR_WE             As Long = 14277081  ' Gris clair    RGB(217, 217, 217)
Public Const CAL_COULEUR_NORMAL         As Long = 16777215  ' Blanc

'==============================================================================
' FUNCTION : GetFeuille
' Retourne une feuille par son nom avec un message d'erreur explicite.
' Utiliser pour eviter les erreurs silencieuses.
'==============================================================================
Function GetFeuille(ByVal nomFeuille As String) As Worksheet

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Sheets(nomFeuille)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Feuille introuvable : """ & nomFeuille & """" & vbCrLf & _
               "Vérifiez le nom de la feuille dans le classeur.", _
               vbCritical, "Erreur de configuration"
    End If

    Set GetFeuille = ws

End Function