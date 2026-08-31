Option Explicit

'==============================================================================
' USERFORM : UserForm_ImpressionFiche
' Projet   : Planning Journalier Automatique
' Version  : 2.0.0
' Auteurs  : JNF
'            Avec la collaboration de Claude (Anthropic)
' Date     : 2026
'
' Description :
'   Formulaire de sélection des sections à inclure dans la fiche personnel
'   avant impression. Appelé par Module_UserForm.UF_ImprimerFiche().
'
' Propriétés du UserForm :
'   Name    : UserForm_ImpressionFiche
'   Caption : Imprimer la fiche
'   Width   : 280
'   Height  : 260
'
' Contrôles requis :
'   lblNomFiche    — Label    — Nom de la personne (affiché en titre)
'   chkHoraire     — CheckBox — Inclure l'horaire hebdomadaire
'   chkAptitudes   — CheckBox — Inclure les aptitudes et formations
'   chkStats       — CheckBox — Inclure les statistiques de passages
'   chkAbsences    — CheckBox — Inclure les absences enregistrées
'   chkTout        — CheckBox — Sélectionner / désélectionner tout
'   btnImprimerFiche — CommandButton — Lancer l'impression
'   btnAnnulerFiche  — CommandButton — Annuler
'==============================================================================

' Propriétés publiques transmises par l'appelant
Public NomPersonne As String
Public LignePers   As Long
Public Confirme    As Boolean

'==============================================================================
' INITIALISATION
' Toutes les sections cochées par défaut.
'==============================================================================
Private Sub UserForm_Initialize()
    chkHoraire.Value = True
    chkAptitudes.Value = True
    chkStats.Value = True
    chkAbsences.Value = True
    Confirme = False
End Sub

'==============================================================================
' SUB : Initialiser
' Affiche le nom de la personne dans le label de titre.
'==============================================================================
Public Sub Initialiser()
    lblNomFiche.Caption = NomPersonne
End Sub

'==============================================================================
' CheckBox : Tout sélectionner / désélectionner
'==============================================================================
Private Sub chkTout_Click()
    Dim etat As Boolean
    etat = chkTout.Value
    chkHoraire.Value = etat
    chkAptitudes.Value = etat
    chkStats.Value = etat
    chkAbsences.Value = etat
End Sub

'==============================================================================
' BOUTON : Imprimer
'==============================================================================
Private Sub btnImprimerFiche_Click()
    Confirme = True
    Me.Hide
End Sub

'==============================================================================
' BOUTON : Annuler
'==============================================================================
Private Sub btnAnnulerFiche_Click()
    Confirme = False
    Me.Hide
End Sub

'==============================================================================
' FERMETURE par la croix
'==============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Confirme = False
End Sub
