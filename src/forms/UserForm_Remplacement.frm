Option Explicit

'==============================================================================
' USERFORM : UserForm_Remplacement
' Projet   : Planning Journalier Automatique
' Version  : 2.0.0
' Auteurs  : JNF
'            Avec la collaboration de Claude (Anthropic)
' Date     : 2026
'
' Description :
'   Mini formulaire d'enregistrement d'un remplacement depuis le calendrier
'   auxiliaires. Appelé par Module_UserForm.UF_AjouterRemplacementDepuisCalendrier().
'
' Propriétés du UserForm :
'   Name    : UserForm_Remplacement
'   Caption : Enregistrer un remplacement
'   Width   : 320
'   Height  : 200
'
' Contrôles requis :
'   txtRplAbsent   — TextBox  — Personne absente (lecture seule)
'   txtRplDate     — TextBox  — Date du remplacement (lecture seule)
'   cboRemplacant  — ComboBox — Liste des auxiliaires disponibles
'   btnRplConfirmer — CommandButton — Confirmer
'   btnRplAnnuler   — CommandButton — Annuler
'==============================================================================

' Propriétés publiques transmises par l'appelant
Public NomAbsent    As String
Public DateProposee As Date
Public Confirme     As Boolean

'==============================================================================
' INITIALISATION
' Charge la liste de tous les auxiliaires actifs dans la ComboBox.
'==============================================================================
Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Set ws = Sheets(NOM_FEUILLE_PERSONNEL)

    Dim dernLigne As Long
    dernLigne = ws.Cells(ws.Rows.count, 2).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernLigne
        If ws.Cells(i, 3).Value = "Actif" And _
           ws.Cells(i, 4).Value = "Auxiliaire" Then
            cboRemplacant.AddItem ws.Cells(i, 2).Value
        End If
    Next i

    Confirme = False

End Sub

'==============================================================================
' SUB : Initialiser
' Pré-remplit les champs et retire la personne absente de la liste.
'==============================================================================
Public Sub Initialiser()

    txtRplAbsent.Value = NomAbsent
    txtRplAbsent.Enabled = False
    txtRplDate.Value = Format(DateProposee, "dd.mm.yyyy")
    txtRplDate.Enabled = False

    ' Retire la personne absente de la liste des remplaçants
    Dim i As Long
    For i = 0 To cboRemplacant.ListCount - 1
        If cboRemplacant.List(i) = NomAbsent Then
            cboRemplacant.RemoveItem i
            Exit For
        End If
    Next i

End Sub

'==============================================================================
' BOUTON : Confirmer
'==============================================================================
Private Sub btnRplConfirmer_Click()

    If cboRemplacant.Value = "" Then
        MsgBox "Veuillez sélectionner un remplacant.", vbExclamation
        cboRemplacant.SetFocus
        Exit Sub
    End If

    Confirme = True
    Me.Hide

End Sub

'==============================================================================
' BOUTON : Annuler
'==============================================================================
Private Sub btnRplAnnuler_Click()
    Confirme = False
    Me.Hide
End Sub

'==============================================================================
' FERMETURE par la croix
'==============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Confirme = False
End Sub

