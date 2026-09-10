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
'   txtRplDate     — TextBox  — Date de début du remplacement (jj.mm.aaaa,
'                               modifiable — pré-remplie avec le jour cliqué
'                               dans le calendrier)
'   txtRplDateFin  — TextBox  — Date de fin du remplacement (jj.mm.aaaa,
'                               NOUVEAU — pré-remplie avec la même date que
'                               le début ; permet d'enregistrer un remplacement
'                               sur plusieurs jours, ex. un week-end entier,
'                               en une seule fois)
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
    ' La date est pré-remplie avec le jour cliqué dans le calendrier mais
    ' reste modifiable : un remplacement peut concerner une autre date que
    ' celle sur laquelle on a double-cliqué (ex. correction, jour voisin).
    txtRplDate.Value = Format(DateProposee, "dd.mm.yyyy")

    ' Date de fin pré-remplie identique à la date de début (remplacement
    ' d'un seul jour par défaut) — à modifier pour couvrir plusieurs jours
    ' (ex. un week-end entier).
    txtRplDateFin.Value = Format(DateProposee, "dd.mm.yyyy")

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

    If Not IsDate(txtRplDate.Value) Then
        MsgBox "Date de début invalide. Format attendu : jj.mm.aaaa", vbExclamation
        txtRplDate.SetFocus
        Exit Sub
    End If

    If Not IsDate(txtRplDateFin.Value) Then
        MsgBox "Date de fin invalide. Format attendu : jj.mm.aaaa", vbExclamation
        txtRplDateFin.SetFocus
        Exit Sub
    End If

    If CDate(txtRplDateFin.Value) < CDate(txtRplDate.Value) Then
        MsgBox "La date de fin ne peut pas être antérieure à la date de début.", vbExclamation
        txtRplDateFin.SetFocus
        Exit Sub
    End If

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

