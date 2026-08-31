Option Explicit

'==============================================================================
' USERFORM : UserForm_GroupeSemaine
' Projet   : Planning Journalier Automatique
' Version  : 2.0.2
' Auteurs  : JNF
'            Avec la collaboration de Claude (Anthropic)
' Date     : 2026
'
' Description :
'   Mini formulaire pour enregistrer tout un groupe d'auxiliaires (G1 ou G2)
'   en renfort sur les 5 jours ouvrés (lundi-vendredi) d'une semaine donnée.
'   Utilisé quand l'entreprise fait appel à tout un groupe d'auxiliaires
'   (ex. étudiants) en semaine, plutôt qu'un remplacement individuel.
'   Appelé par Module_UserForm.UF_AjouterGroupeSemaine().
'
' Propriétés du UserForm :
'   Name    : UserForm_GroupeSemaine
'   Caption : Ajouter un groupe pour la semaine
'   Width   : 320
'   Height  : 190
'
' Contrôles requis (NOUVEAU FORMULAIRE - à créer dans l'éditeur VBA) :
'   lblGroupeSemaineInfo — Label     — texte d'aide, ex. "Enregistre tout le
'                          groupe en renfort du lundi au vendredi"
'   cboGroupeSemaine     — ComboBox  — Groupe (G1 / G2)
'   txtDebutSemaine      — TextBox   — Lundi de la semaine (jj.mm.aaaa)
'   btnGroupeSemaineConfirmer — CommandButton — "Enregistrer"
'   btnGroupeSemaineAnnuler   — CommandButton — "Annuler"
'==============================================================================

' Propriétés publiques transmises par l'appelant
Public DateProposee As Date   ' Lundi de la semaine en cours, pré-rempli
Public Confirme     As Boolean

'==============================================================================
' INITIALISATION
'==============================================================================
Private Sub UserForm_Initialize()

    With cboGroupeSemaine
        .AddItem GROUPE_G1
        .AddItem GROUPE_G2
        .Value = GROUPE_G1
    End With

    lblGroupeSemaineInfo.Caption = "Enregistre tout le groupe en renfort, " & _
        "du lundi au vendredi de la semaine choisie."

    Confirme = False

End Sub

'==============================================================================
' SUB : Initialiser
' Pré-remplit le lundi de la semaine avec la date transmise par l'appelant.
'==============================================================================
Public Sub Initialiser()
    txtDebutSemaine.Value = Format(DateProposee, "dd.mm.yyyy")
End Sub

'==============================================================================
' BOUTON : Confirmer
'==============================================================================
Private Sub btnGroupeSemaineConfirmer_Click()

    If cboGroupeSemaine.Value = "" Then
        MsgBox "Veuillez sélectionner un groupe.", vbExclamation
        cboGroupeSemaine.SetFocus
        Exit Sub
    End If

    If Not IsDate(txtDebutSemaine.Value) Then
        MsgBox "Date invalide. Format attendu : jj.mm.aaaa", vbExclamation
        txtDebutSemaine.SetFocus
        Exit Sub
    End If

    If Weekday(CDate(txtDebutSemaine.Value), vbMonday) <> 1 Then
        MsgBox "Veuillez indiquer le lundi de la semaine concernée.", vbExclamation
        txtDebutSemaine.SetFocus
        Exit Sub
    End If

    Confirme = True
    Me.Hide

End Sub

'==============================================================================
' BOUTON : Annuler
'==============================================================================
Private Sub btnGroupeSemaineAnnuler_Click()
    Confirme = False
    Me.Hide
End Sub

'==============================================================================
' FERMETURE par la croix
'==============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Confirme = False
End Sub
