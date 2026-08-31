Option Explicit

'==============================================================================
' MODULE : Commande_Bouton
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'
' Description :
'   Commandes liées aux boutons de la feuille "Planning journalier".
'   - Navigation dans les dates
'   - Impression du planning avec archivage PDF
'   - Anti-doublon d'incrémentation des statistiques
'   - Ouverture du UserForm de gestion du personnel
'==============================================================================

' Lignes/colonnes du tableau Planning_Matin et dossier d'archivage :
' LIGNE_DEBUT_MATIN, LIGNE_FIN_MATIN, PLAN_COL_NOMS, PLAN_COL_POSTES et
' NOM_DOSSIER_ARCHIVES viennent de Module_Constantes (source unique,
' partagée avec Planning_Auto).

'==============================================================================
' SUB : DatePlusUn
' Ajoute 1 jour à la date du planning.
'==============================================================================
Sub DatePlusUn()
    On Error GoTo ErrDate
    Call Mem
    wsP.Range(CELL_DATE_PLANNING).Value = wsP.Range(CELL_DATE_PLANNING).Value + 1
    Exit Sub
ErrDate:
    MsgBox "Impossible de modifier la date du planning :" & vbCrLf & Err.Description, vbExclamation
End Sub

'==============================================================================
' SUB : DateMoinsUn
' Retire 1 jour à la date du planning.
'==============================================================================
Sub DateMoinsUn()
    On Error GoTo ErrDate
    Call Mem
    wsP.Range(CELL_DATE_PLANNING).Value = wsP.Range(CELL_DATE_PLANNING).Value - 1
    Exit Sub
ErrDate:
    MsgBox "Impossible de modifier la date du planning :" & vbCrLf & Err.Description, vbExclamation
End Sub

'==============================================================================
' SUB : DateAujourdhui
' Remplace la date du planning par la date du jour.
'==============================================================================
Sub DateAujourdhui()
    On Error GoTo ErrDate
    Call Mem
    wsP.Range(CELL_DATE_PLANNING).Value = Date
    Exit Sub
ErrDate:
    MsgBox "Impossible de modifier la date du planning :" & vbCrLf & Err.Description, vbExclamation
End Sub

'==============================================================================
' SUB : ImprimerPlanning
' Gère l'impression du planning avec :
'   - Vérification anti-doublon (évite d'incrémenter les stats deux fois)
'   - Boîte de dialogue impression standard Excel
'   - Archivage automatique en PDF (Archives\annee\mois\date\hh.mm.ss.pdf)
'   - Incrémentation des statistiques dans la feuille Personnel
'==============================================================================
Sub ImprimerPlanning()

    On Error GoTo ErrImpression
    Call Mem

    Dim datePlanning As Date
    datePlanning = wsP.Range(CELL_DATE_PLANNING).Value

    ' --- Anti-doublon : vérifie si ce planning a déjà été imprimé ---
    Dim incrementerStats As Boolean
    incrementerStats = True

    If wsP.Range(CELL_MAJ_STATS).Value = datePlanning Then

        Dim rep As VbMsgBoxResult
        rep = MsgBox("Ce planning à déjà été imprimé." & vbCrLf & _
                     "Voulez-vous incrémenter les statistiques ?", _
                     vbYesNo + vbExclamation, "Attention")

        If rep = vbNo Then incrementerStats = False

    End If

    ' --- Boîte de dialogue impression ---
    ' Date_print n'est enregistrée qu'après confirmation de l'impression :
    ' l'écrire avant laisserait un état incohérent (impression "datée" alors
    ' qu'elle a été annulée) si l'utilisateur clique sur Annuler.
    If Application.Dialogs(xlDialogPrint).Show = False Then Exit Sub

    ' --- Enregistre la date et l'heure de la dernière impression ---
    wsP.Range(CELL_DATE_PRINT).Value = Now

    ' --- Construction du chemin d'archivage PDF ---
    ' Organisation : Archives \ annee \ mois \ date_planning \ hh.nn.ss.pdf
    Dim dossierRacine As String
    Dim dossierAnnee  As String
    Dim dossierMois   As String
    Dim dossierJour   As String
    Dim nomFichier    As String
    Dim cheminFinal   As String

    dossierRacine = ThisWorkbook.Path & "\" & NOM_DOSSIER_ARCHIVES & "\"
    dossierAnnee = dossierRacine & Format(datePlanning, "yyyy") & "\"
    dossierMois = dossierAnnee & Format(datePlanning, "mm") & "\"
    dossierJour = dossierMois & Format(datePlanning, "yyyy.mm.dd") & "\"

    ' Crée les dossiers manquants
    If Dir(dossierRacine, vbDirectory) = "" Then MkDir dossierRacine
    If Dir(dossierAnnee, vbDirectory) = "" Then MkDir dossierAnnee
    If Dir(dossierMois, vbDirectory) = "" Then MkDir dossierMois
    If Dir(dossierJour, vbDirectory) = "" Then MkDir dossierJour

    ' Nom du fichier = horodatage de l'impression (unicité garantie)
    nomFichier = Format(Now, "hh.nn.ss") & ".pdf"
    cheminFinal = dossierJour & nomFichier

    ' --- Export PDF ---
    wsP.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=cheminFinal, _
        Quality:=xlQualityStandard

    ' --- Incrémentation des statistiques ---
    If incrementerStats Then

        Dim ligne As Long
        For ligne = LIGNE_DEBUT_MATIN To LIGNE_FIN_MATIN

            Dim nom   As String
            Dim poste As String
            nom = Trim(wsP.Cells(ligne, PLAN_COL_NOMS).Value)
            poste = Trim(wsP.Cells(ligne, PLAN_COL_POSTES).Value)

            If nom <> "" And poste <> "" Then
                Call Planning_Auto.IncrementStats(nom, poste)
            End If

        Next ligne

        ' Enregistre la date du dernier planning imprimé avec stats
        wsP.Range(CELL_MAJ_STATS).Value = datePlanning
        
        ' Sauvegarde automatique après mise à jour des statistiques
        ThisWorkbook.Save
        
        ' Export stats hebdomadaires si c'est un lundi
        Call Module_StatsHebdo.ImprimerStatsHebdo(datePlanning)
        
        MsgBox "Impression réussie." & vbCrLf & _
               "Statistiques mises à jour.", _
               vbInformation, "Planning v2.0"

    Else

        MsgBox "Planning imprimé sans mise à jour des statistiques.", _
               vbInformation, "Planning v2.0"

    End If

    Exit Sub

ErrImpression:
    MsgBox "Erreur lors de l'impression." & vbCrLf & Err.Description, vbCritical

End Sub

'==============================================================================
' SUB : OuvrirGestionPersonnel
' Lance le UserForm de gestion du personnel.
'==============================================================================
Sub OuvrirGestionPersonnel()
    Call Mem
    UserForm_GestionPersonnel.Show
End Sub