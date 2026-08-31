Option Explicit

'==============================================================================
' USERFORM : UserForm_GroupeAux
' Projet   : Planning Journalier Automatique
' Version  : 2.0.0
' Auteurs  : JNF
'            Avec la collaboration de Claude (Anthropic)
' Date     : 2026
'
' Description :
'   Mini formulaire d'attribution du groupe auxiliaire (G1/G2) pour un jour
'   sélectionné dans le calendrier de l'onglet Auxiliaires.
'   Appelé par Module_UserForm.UF_ClicCalendrierAux().
'
' Propriétés du UserForm :
'   Name    : UserForm_GroupeAux
'   Caption : Attribution du groupe
'   Width   : 260
'   Height  : 180
'
' Contrôles requis :
'   lblDateGroupe   — Label  — Affiche la date sélectionnée
'   lblGroupeActuel — Label  — Affiche le groupe actuel
'   btnG1           — CommandButton — Attribuer Groupe 1
'   btnG2           — CommandButton — Attribuer Groupe 2
'   btnSupprimer    — CommandButton — Supprimer l'attribution manuelle
'   btnAnnuler      — CommandButton — Annuler
'==============================================================================

' Propriétés publiques transmises par l'appelant
Public dateJour As Date
Public Confirme As Boolean
Public Action   As String  ' "G1", "G2" ou "Supprimer"

'==============================================================================
' INITIALISATION
'==============================================================================
Private Sub UserForm_Initialize()
    Confirme = False
    Action = ""
End Sub

'==============================================================================
' SUB : Initialiser
' Affiche la date et le groupe actuel.
'==============================================================================
Public Sub Initialiser()

    lblDateGroupe.Caption = "Date : " & Format(dateJour, "dddd dd.mm.yyyy")

    Dim groupeActuel As String
    groupeActuel = Module_Groupes.GetGroupeJour(dateJour)

    If groupeActuel <> "" Then
        lblGroupeActuel.Caption = "Groupe actuel : " & groupeActuel
    Else
        lblGroupeActuel.Caption = "Aucune attribution"
    End If

End Sub

'==============================================================================
' BOUTONS d'attribution
'==============================================================================
Private Sub btnG1_Click()
    Action = GROUPE_G1
    Confirme = True
    Me.Hide
End Sub

Private Sub btnG2_Click()
    Action = GROUPE_G2
    Confirme = True
    Me.Hide
End Sub

Private Sub btnSupprimer_Click()
    Action = "Supprimer"
    Confirme = True
    Me.Hide
End Sub

Private Sub btnAnnuler_Click()
    Confirme = False
    Me.Hide
End Sub

'==============================================================================
' FERMETURE par la croix
'==============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Confirme = False
End Sub
