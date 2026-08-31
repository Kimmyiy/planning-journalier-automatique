Option Explicit

'==============================================================================
' USERFORM : UserForm_Absence
' Projet   : Planning Journalier Automatique
' Version  : 2.0.0
' Auteurs  : JNF
'            Avec la collaboration de Claude (Anthropic)
' Date     : 2026
'
' Description :
'   Mini formulaire d'ajout d'une absence depuis le calendrier personnel.
'   Appelé par Module_UserForm.UF_AjouterAbsenceDepuisCalendrier().
'
' Propriétés du UserForm :
'   Name    : UserForm_Absence
'   Caption : Ajouter une absence
'   Width   : 300
'   Height  : 220
'
' Contrôles requis :
'   txtAbsNom    — TextBox  — Nom de la personne (lecture seule)
'   cboAbsType   — ComboBox — Type d'absence (Vacances / Maladie / Congé)
'   cboAbsPeriode — ComboBox — Période (Journée / Matin / Après-midi) — NOUVEAU
'   txtAbsDebut  — TextBox  — Date de début (jj.mm.aaaa)
'   txtAbsFin    — TextBox  — Date de fin   (jj.mm.aaaa)
'   btnAbsConfirmer — CommandButton — Confirmer
'   btnAbsAnnuler   — CommandButton — Annuler
'==============================================================================

' Propriétés publiques transmises par l'appelant
Public NomPersonne  As String
Public DateProposee As Date
Public Confirme     As Boolean

'==============================================================================
' INITIALISATION
' Remplit uniquement la ComboBox. Les champs sont remplis par Initialiser().
'==============================================================================
Private Sub UserForm_Initialize()

    With cboAbsType
        .AddItem TYPE_ABS_VACANCES
        .AddItem TYPE_ABS_MALADIE
        .AddItem TYPE_ABS_CONGE
        .Value = TYPE_ABS_VACANCES
    End With

    With cboAbsPeriode
        .AddItem PERIODE_JOURNEE
        .AddItem PERIODE_MATIN
        .AddItem PERIODE_APRESMIDI
        .Value = PERIODE_JOURNEE
    End With

    Confirme = False

End Sub

'==============================================================================
' SUB : Initialiser
' Appelée après l'assignation de NomPersonne et DateProposee.
' Pré-remplit les champs avec les valeurs transmises par l'appelant.
'==============================================================================
Public Sub Initialiser()
    txtAbsNom.Value = NomPersonne
    txtAbsNom.Enabled = False
    txtAbsDebut.Value = Format(DateProposee, "dd.mm.yyyy")
    txtAbsFin.Value = Format(DateProposee, "dd.mm.yyyy")
End Sub

'==============================================================================
' BOUTON : Confirmer
'==============================================================================
Private Sub btnAbsConfirmer_Click()

    ' Validation du type
    If cboAbsType.Value = "" Then
        MsgBox "Veuillez sélectionner un type d'absence.", vbExclamation
        Exit Sub
    End If

    ' Validation des dates
    If Not IsDate(txtAbsDebut.Value) Then
        MsgBox "Date de début invalide." & vbCrLf & "Format attendu : jj.mm.aaaa", _
               vbExclamation
        txtAbsDebut.SetFocus
        Exit Sub
    End If

    If Not IsDate(txtAbsFin.Value) Then
        MsgBox "Date de fin invalide." & vbCrLf & "Format attendu : jj.mm.aaaa", _
               vbExclamation
        txtAbsFin.SetFocus
        Exit Sub
    End If

    Dim dDebut As Date
    Dim dFin   As Date
    dDebut = CDate(txtAbsDebut.Value)
    dFin = CDate(txtAbsFin.Value)

    If dFin < dDebut Then
        MsgBox "La date de fin ne peut pas être antérieure à la date de début.", _
               vbExclamation
        txtAbsFin.SetFocus
        Exit Sub
    End If

    Confirme = True
    Me.Hide

End Sub

'==============================================================================
' BOUTON : Annuler
'==============================================================================
Private Sub btnAbsAnnuler_Click()
    Confirme = False
    Me.Hide
End Sub

'==============================================================================
' FERMETURE par la croix
'==============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Confirme = False
End Sub

