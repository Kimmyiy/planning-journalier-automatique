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