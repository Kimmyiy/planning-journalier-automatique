Option Explicit

'==============================================================================
' MODULE  : UserForm_GestionPersonnel
' Projet  : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'==============================================================================

' Variables du calendrier feriés (onglet Calendrier)
Private m_Mois  As Long
Private m_Annee As Long

Private Sub btnAjouterFermeture_Click()
    Call Module_UserForm.UF_AjouterFermeture(Me)
    Call DessinnerCalendrier
End Sub

Private Sub btnAnnulerFermeture_Click()
    txtFermDebut.Value = ""
    txtFermFin.Value = ""
    txtFermNom.Value = ""
End Sub

Private Sub cboAuxAbsente_Change()

End Sub
Private Sub cboTypeRpl_Change()
    cboAuxAbsente.Enabled = (cboTypeRpl.Value = "Remplacement")
    If cboTypeRpl.Value = "Renfort" Then cboAuxAbsente.Value = ""
End Sub

'==============================================================================
' INITIALISATION
'==============================================================================
Private Sub UserForm_Initialize()

    ' --- ComboBox Statut ---
    With cboStatut
        .AddItem "Actif"
        .AddItem "Archive"
    End With

    ' --- ComboBox Type ---
    With cboType
        .AddItem "Fixe"
        .AddItem "Auxiliaire"
    End With

    ' --- ComboBox Groupe ---
    With cboGroupe
        .AddItem "-"
        .AddItem "G1"
        .AddItem "G2"
    End With

    ' --- ComboBox Horaires ---
    Dim cboH As Variant
    For Each cboH In Array(cboLun, cboMar, cboMer, cboJeu, cboVen, cboSam, cboDim)
        cboH.AddItem "M"
        cboH.AddItem "J"
        cboH.AddItem "-"
    Next cboH

    ' --- ComboBox Type absence ---
    With cboVacType
        .AddItem "Vacances"
        .AddItem "Conge"
        .AddItem "Maladie"
    End With

    ' --- ComboBox Periode absence (nouveau) ---
    With cboVacPeriode
        .AddItem PERIODE_JOURNEE
        .AddItem PERIODE_MATIN
        .AddItem PERIODE_APRESMIDI
        .Value = PERIODE_JOURNEE
    End With
    
        ' --- ComboBox Type remplacement ---
    With cboTypeRpl
        .AddItem "Remplacement"
        .AddItem "Renfort"
        .Value = "Remplacement"
    End With

    ' --- Etat initial ---
    txtLigneActive.Value = "0"

    ' --- Chargements onglet Personnel ---
    Call Module_UserForm.UF_RecherchePersonnel(Me, "")

    ' --- Chargements onglet Auxiliaires ---
    Call Module_UserForm.UF_ChargerCalendrierWeekend(Me)
    Call Module_UserForm.UF_ChargerAuxiliaires(Me, "G1")

    ' --- Chargements onglet Absences ---
    Call Module_UserForm.UF_ChargerPersonnelAbsences(Me)
    Call Module_UserForm.UF_MettreAJourRemarques(Me)

    ' --- Calendrier ferias (onglet Calendrier) ---
    Call InitCalendrier

    ' --- Calendrier personnel (onglet Personnel) ---
    txtCalMois.Value = CStr(Month(Date))
    txtCalAnnee.Value = CStr(Year(Date))

    ' --- Parametres machines ---
    Call Module_UserForm.UF_ChargerParametres(Me)

    ' --- Calendrier auxiliaires (onglet Auxiliaires) ---
    ' IMPORTANT : appele en dernier car necessite que tous les controles
    ' de l onglet Auxiliaires soient initialises
    Call Module_UserForm.UF_InitCalendrierAux(Me)
    
    ' Ouvre sur l'onglet Personnel au démarrage
    mpOnglets.Value = 0

End Sub

'==============================================================================
' ONGLET PERSONNEL
'==============================================================================
Private Sub txtRecherche_Change()
    Call Module_UserForm.UF_RecherchePersonnel(Me, txtRecherche.Value)
End Sub

Private Sub chkAfficherArchives_Click()
    Call Module_UserForm.UF_RecherchePersonnel(Me, txtRecherche.Value)
End Sub

Private Sub chkFiltreFixe_Click()
    Call Module_UserForm.UF_RecherchePersonnel(Me, txtRecherche.Value)
End Sub

Private Sub chkFiltreG1_Click()
    Call Module_UserForm.UF_RecherchePersonnel(Me, txtRecherche.Value)
End Sub

Private Sub chkFiltreG2_Click()
    Call Module_UserForm.UF_RecherchePersonnel(Me, txtRecherche.Value)
End Sub

Private Sub lstPersonnel_Click()
    Call Module_UserForm.UF_ChargerFiche(Me)
    Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me)
End Sub

Private Sub btnNouvelleFiche_Click()
    Call Module_UserForm.UF_NouvellesFiche(Me)
End Sub

Private Sub btnEnregistrer_Click()
    Dim ok As Boolean
    ok = Module_UserForm.UF_EnregistrerFiche(Me)
    If ok Then Call Module_UserForm.UF_RecherchePersonnel(Me, txtRecherche.Value)
End Sub

Private Sub btnArchiverFiche_Click()
    Call Module_UserForm.UF_ArchiverFiche(Me)
    Call Module_UserForm.UF_RecherchePersonnel(Me, txtRecherche.Value)
End Sub

Private Sub btnSupprimerStats_Click()
    Call Module_UserForm.UF_SupprimerStats(Me)
End Sub

Private Sub btnImprimerFiche_Click()
    Call Module_UserForm.UF_ImprimerFiche(Me)
End Sub

Private Sub cboType_Change()
    Dim estAux As Boolean
    estAux = (cboType.Value = "Auxiliaire")
    cboGroupe.Enabled = estAux
    cboSam.Enabled = estAux
    cboDim.Enabled = estAux
    If Not estAux Then
        cboGroupe.Value = "-"
        cboSam.Value = "-"
        cboDim.Value = "-"
    End If
End Sub

' Calendrier personnel - navigation
Private Sub btnCalPrev_Click()
    Call Module_UserForm.UF_CalendrierMoisPrecedent(Me)
End Sub

Private Sub btnCalNext_Click()
    Call Module_UserForm.UF_CalendrierMoisSuivant(Me)
End Sub

Private Sub btnAjouterAbsence_Click()
    Call Module_UserForm.UF_BoutonCalendrierPersonnel(Me)
End Sub

' Calendrier personnel - 42 clics
Private Sub lblCal0_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal1_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal2_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal3_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal4_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal5_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal6_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal7_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal8_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal9_Click():  Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal10_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal11_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal12_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal13_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal14_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal15_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal16_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal17_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal18_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal19_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal20_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal21_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal22_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal23_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal24_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal25_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal26_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal27_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal28_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal29_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal30_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal31_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal32_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal33_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal34_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal35_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal36_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal37_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal38_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal39_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal40_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub
Private Sub lblCal41_Click(): Call Module_UserForm.UF_ChargerCalendrierPersonnel(Me): End Sub

' Calendrier personnel - 42 double-clics
Private Sub lblCal0_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 0):  End Sub
Private Sub lblCal1_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 1):  End Sub
Private Sub lblCal2_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 2):  End Sub
Private Sub lblCal3_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 3):  End Sub
Private Sub lblCal4_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 4):  End Sub
Private Sub lblCal5_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 5):  End Sub
Private Sub lblCal6_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 6):  End Sub
Private Sub lblCal7_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 7):  End Sub
Private Sub lblCal8_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 8):  End Sub
Private Sub lblCal9_DblClick(ByVal Cancel As MSForms.ReturnBoolean):  Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 9):  End Sub
Private Sub lblCal10_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 10): End Sub
Private Sub lblCal11_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 11): End Sub
Private Sub lblCal12_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 12): End Sub
Private Sub lblCal13_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 13): End Sub
Private Sub lblCal14_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 14): End Sub
Private Sub lblCal15_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 15): End Sub
Private Sub lblCal16_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 16): End Sub
Private Sub lblCal17_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 17): End Sub
Private Sub lblCal18_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 18): End Sub
Private Sub lblCal19_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 19): End Sub
Private Sub lblCal20_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 20): End Sub
Private Sub lblCal21_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 21): End Sub
Private Sub lblCal22_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 22): End Sub
Private Sub lblCal23_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 23): End Sub
Private Sub lblCal24_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 24): End Sub
Private Sub lblCal25_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 25): End Sub
Private Sub lblCal26_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 26): End Sub
Private Sub lblCal27_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 27): End Sub
Private Sub lblCal28_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 28): End Sub
Private Sub lblCal29_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 29): End Sub
Private Sub lblCal30_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 30): End Sub
Private Sub lblCal31_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 31): End Sub
Private Sub lblCal32_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 32): End Sub
Private Sub lblCal33_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 33): End Sub
Private Sub lblCal34_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 34): End Sub
Private Sub lblCal35_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 35): End Sub
Private Sub lblCal36_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 36): End Sub
Private Sub lblCal37_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 37): End Sub
Private Sub lblCal38_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 38): End Sub
Private Sub lblCal39_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 39): End Sub
Private Sub lblCal40_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 40): End Sub
Private Sub lblCal41_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Module_UserForm.UF_DoubleclicCalendrier_V2(Me, 41): End Sub

'==============================================================================
' ONGLET AUXILIAIRES
'==============================================================================
Private Sub lstCalendrierWE_Click()

    If lstCalendrierWE.ListIndex = -1 Then Exit Sub

    Dim ligneCal As String
    ligneCal = lstCalendrierWE.List(lstCalendrierWE.ListIndex, 0)

    Dim groupe As String
    groupe = "G1"
    If InStr(ligneCal, "[G2]") > 0 Then groupe = "G2"

    Call Module_UserForm.UF_ChargerAuxiliaires(Me, groupe)

    Dim dateSamStr As String
    dateSamStr = lstCalendrierWE.List(lstCalendrierWE.ListIndex, 1)

    If dateSamStr <> "" Then
        Dim dateSam As Date
        dateSam = CDate(CDbl(dateSamStr))
        dtpDateRemplacement.Value = Format(dateSam, "dd.mm.yyyy")
    End If

    Call Module_UserForm.UF_ChargerRemplacementsWE(Me)

End Sub

Private Sub btnEnregistrerRpl_Click()
    Call Module_UserForm.UF_EnregistrerRemplacement(Me)
End Sub

Private Sub btnSupprimerRpl_Click()
    Call Module_UserForm.UF_SupprimerRemplacementIndividuel(Me)
End Sub

Private Sub btnSupprimerRplIndividuel_Click()
    Call Module_UserForm.UF_SupprimerRemplacementIndividuel(Me)
End Sub

' NOUVEAU CONTROLE A CREER : btnAjouterGroupeSemaine (CommandButton),
' caption "Ajouter un groupe pour la semaine…", onglet Auxiliaires.
Private Sub btnAjouterGroupeSemaine_Click()
    Call Module_UserForm.UF_AjouterGroupeSemaine(Me)
End Sub

' Calendrier auxiliaires - navigation
Private Sub btnAuxCalPrev_Click()
    Call Module_UserForm.UF_AuxCalMoisPrecedent(Me)
End Sub

Private Sub btnAuxCalNext_Click()
    Call Module_UserForm.UF_AuxCalMoisSuivant(Me)
End Sub

' Reference G1
Private Sub btnSaveRefG1_Click()
    Call Module_UserForm.UF_EnregistrerRefGroupeG1(Me)
End Sub

' Generer groupes
Private Sub btnGenererGroupes_Click()
    Call Module_UserForm.UF_ChargerGroupesAnnee(Me)
End Sub

' Calendrier auxiliaires - 42 clics
Private Sub lblAuxCal0_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 0):  End Sub
Private Sub lblAuxCal1_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 1):  End Sub
Private Sub lblAuxCal2_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 2):  End Sub
Private Sub lblAuxCal3_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 3):  End Sub
Private Sub lblAuxCal4_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 4):  End Sub
Private Sub lblAuxCal5_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 5):  End Sub
Private Sub lblAuxCal6_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 6):  End Sub
Private Sub lblAuxCal7_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 7):  End Sub
Private Sub lblAuxCal8_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 8):  End Sub
Private Sub lblAuxCal9_Click():  Call Module_UserForm.UF_ClicCalendrierAux(Me, 9):  End Sub
Private Sub lblAuxCal10_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 10): End Sub
Private Sub lblAuxCal11_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 11): End Sub
Private Sub lblAuxCal12_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 12): End Sub
Private Sub lblAuxCal13_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 13): End Sub
Private Sub lblAuxCal14_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 14): End Sub
Private Sub lblAuxCal15_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 15): End Sub
Private Sub lblAuxCal16_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 16): End Sub
Private Sub lblAuxCal17_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 17): End Sub
Private Sub lblAuxCal18_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 18): End Sub
Private Sub lblAuxCal19_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 19): End Sub
Private Sub lblAuxCal20_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 20): End Sub
Private Sub lblAuxCal21_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 21): End Sub
Private Sub lblAuxCal22_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 22): End Sub
Private Sub lblAuxCal23_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 23): End Sub
Private Sub lblAuxCal24_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 24): End Sub
Private Sub lblAuxCal25_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 25): End Sub
Private Sub lblAuxCal26_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 26): End Sub
Private Sub lblAuxCal27_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 27): End Sub
Private Sub lblAuxCal28_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 28): End Sub
Private Sub lblAuxCal29_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 29): End Sub
Private Sub lblAuxCal30_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 30): End Sub
Private Sub lblAuxCal31_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 31): End Sub
Private Sub lblAuxCal32_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 32): End Sub
Private Sub lblAuxCal33_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 33): End Sub
Private Sub lblAuxCal34_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 34): End Sub
Private Sub lblAuxCal35_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 35): End Sub
Private Sub lblAuxCal36_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 36): End Sub
Private Sub lblAuxCal37_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 37): End Sub
Private Sub lblAuxCal38_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 38): End Sub
Private Sub lblAuxCal39_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 39): End Sub
Private Sub lblAuxCal40_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 40): End Sub
Private Sub lblAuxCal41_Click(): Call Module_UserForm.UF_ClicCalendrierAux(Me, 41): End Sub

'==============================================================================
' ONGLET ABSENCES
'==============================================================================
Private Sub lstAbsPersonnel_Click()
    Call Module_UserForm.UF_ChargerAbsencesPersonne(Me)
End Sub

Private Sub btnEnregistrerAbs_Click()
    Call Module_UserForm.UF_EnregistrerAbsence(Me)
End Sub

Private Sub btnSupprimerAbs_Click()
    Call Module_UserForm.UF_SupprimerAbsence(Me)
End Sub

'==============================================================================
' ONGLET CALENDRIER FERIAS
'==============================================================================
Private Sub InitCalendrier()
    m_Mois = Month(Date)
    m_Annee = Year(Date)
    spnAnnee.Min = 2020
    spnAnnee.Max = 2040
    spnAnnee.Value = m_Annee
    lblAnneeCharge.Caption = CStr(m_Annee)
    Call DessinnerCalendrier
End Sub

Private Sub DessinnerCalendrier()

    Dim dateRefG1 As Date
    On Error Resume Next
    dateRefG1 = Sheets("Planning journalier").Range("Ref_GroupeG1").Value
    On Error GoTo 0

    If dateRefG1 = 0 Then
        dateRefG1 = DateSerial(2026, 1, 5)
        lblInfoFerie.Caption = "Attention : Ref_GroupeG1 non configuree"
    End If

    Dim noms(1 To 12) As String
    noms(1) = "Janvier":    noms(2) = "Fevrier":    noms(3) = "Mars"
    noms(4) = "Avril":      noms(5) = "Mai":         noms(6) = "Juin"
    noms(7) = "Juillet":    noms(8) = "Aout":        noms(9) = "Septembre"
    noms(10) = "Octobre":   noms(11) = "Novembre":   noms(12) = "Decembre"

    lblMoisAnnee.Caption = noms(m_Mois) & " " & m_Annee

    Dim jours() As JourCalendrier
    jours = Module_Feries.GenererMoisCalendrier(m_Annee, m_Mois, dateRefG1)

    Dim i As Long
    For i = 0 To 41
        Dim lbl As MSForms.Label
        Set lbl = Me.Controls("lblJ" & i)
        lbl.Caption = Day(jours(i).dateJour)
        lbl.BackColor = jours(i).couleur
        lbl.Tag = CStr(CDbl(jours(i).dateJour))
        If Not jours(i).estDuMois Then
            lbl.ForeColor = RGB(170, 170, 170)
            lbl.Font.Bold = False
        ElseIf jours(i).estFerie Then
            lbl.ForeColor = RGB(150, 0, 0)
            lbl.Font.Bold = True
        Else
            lbl.ForeColor = RGB(0, 0, 0)
            lbl.Font.Bold = False
        End If
        If jours(i).estFerie And jours(i).estDuMois Then
            lbl.ControlTipText = jours(i).nomFerie
        ElseIf jours(i).estWE And jours(i).estDuMois Then
            lbl.ControlTipText = "Weekend - " & jours(i).groupe
        Else
            lbl.ControlTipText = Format(jours(i).dateJour, "dd.mm.yyyy")
        End If
    Next i

    lblInfoFerie.Caption = "Cliquez sur un jour pour ajouter/supprimer un ferie"

End Sub

Private Sub btnMoisPrec_Click()
    m_Mois = m_Mois - 1
    If m_Mois < 1 Then: m_Mois = 12: m_Annee = m_Annee - 1
    Call DessinnerCalendrier
End Sub

Private Sub btnMoisSuiv_Click()
    m_Mois = m_Mois + 1
    If m_Mois > 12 Then: m_Mois = 1: m_Annee = m_Annee + 1
    Call DessinnerCalendrier
End Sub

Private Sub spnAnnee_Change()
    lblAnneeCharge.Caption = CStr(spnAnnee.Value)
End Sub

Private Sub btnChargerAnnee_Click()
    Dim annee As Long
    annee = CLng(spnAnnee.Value)
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Charger les fériés " & annee & " de Neuchâtel ?", _
                 vbYesNo + vbQuestion, "Charger les fériés")
    If rep = vbNo Then Exit Sub
    Call Module_Feries.ChargerFeriesAnnee(annee)
    m_Annee = annee
    Call DessinnerCalendrier
End Sub

Sub CalendrierJourClic(ByVal indexJour As Integer)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls("lblJ" & indexJour)
    If lbl.Tag = "" Then Exit Sub
    Dim dateJour As Date
    dateJour = CDate(CDbl(lbl.Tag))
    If Month(dateJour) <> m_Mois Or Year(dateJour) <> m_Annee Then
        m_Mois = Month(dateJour)
        m_Annee = Year(dateJour)
        Call DessinnerCalendrier
        Exit Sub
    End If
    If Module_Feries.estFerie(dateJour) Then
        Dim nomF As String
        nomF = Module_Feries.GetNomFerie(dateJour)
        Dim rep As VbMsgBoxResult
        rep = MsgBox("Le " & Format(dateJour, "dd.mm.yyyy") & " est un férié : " & nomF & vbCrLf & _
                     "Voulez-vous le supprimer ?", vbYesNo + vbQuestion, "Supprimer le férié")
        If rep = vbYes Then
            Call Module_Feries.SupprimerFerie(dateJour)
            lblInfoFerie.Caption = "Ferie supprime : " & Format(dateJour, "dd.mm.yyyy")
        End If
    Else
        Dim nomManuel As String
        nomManuel = InputBox("Nom du jour ferie :" & vbCrLf & _
                             Format(dateJour, "dddd dd.mm.yyyy"), _
                             "Ajouter un ferie manuel", "Ferie exceptionnel")
        If nomManuel <> "" Then
            Call Module_Feries.AjouterFerieManuel(dateJour, nomManuel)
            lblInfoFerie.Caption = "Ferie ajoute : " & nomManuel
        End If
    End If
    Call DessinnerCalendrier
End Sub

' Calendrier ferias - 42 clics
Private Sub lblJ0_Click():  Call CalendrierJourClic(0):  End Sub
Private Sub lblJ1_Click():  Call CalendrierJourClic(1):  End Sub
Private Sub lblJ2_Click():  Call CalendrierJourClic(2):  End Sub
Private Sub lblJ3_Click():  Call CalendrierJourClic(3):  End Sub
Private Sub lblJ4_Click():  Call CalendrierJourClic(4):  End Sub
Private Sub lblJ5_Click():  Call CalendrierJourClic(5):  End Sub
Private Sub lblJ6_Click():  Call CalendrierJourClic(6):  End Sub
Private Sub lblJ7_Click():  Call CalendrierJourClic(7):  End Sub
Private Sub lblJ8_Click():  Call CalendrierJourClic(8):  End Sub
Private Sub lblJ9_Click():  Call CalendrierJourClic(9):  End Sub
Private Sub lblJ10_Click(): Call CalendrierJourClic(10): End Sub
Private Sub lblJ11_Click(): Call CalendrierJourClic(11): End Sub
Private Sub lblJ12_Click(): Call CalendrierJourClic(12): End Sub
Private Sub lblJ13_Click(): Call CalendrierJourClic(13): End Sub
Private Sub lblJ14_Click(): Call CalendrierJourClic(14): End Sub
Private Sub lblJ15_Click(): Call CalendrierJourClic(15): End Sub
Private Sub lblJ16_Click(): Call CalendrierJourClic(16): End Sub
Private Sub lblJ17_Click(): Call CalendrierJourClic(17): End Sub
Private Sub lblJ18_Click(): Call CalendrierJourClic(18): End Sub
Private Sub lblJ19_Click(): Call CalendrierJourClic(19): End Sub
Private Sub lblJ20_Click(): Call CalendrierJourClic(20): End Sub
Private Sub lblJ21_Click(): Call CalendrierJourClic(21): End Sub
Private Sub lblJ22_Click(): Call CalendrierJourClic(22): End Sub
Private Sub lblJ23_Click(): Call CalendrierJourClic(23): End Sub
Private Sub lblJ24_Click(): Call CalendrierJourClic(24): End Sub
Private Sub lblJ25_Click(): Call CalendrierJourClic(25): End Sub
Private Sub lblJ26_Click(): Call CalendrierJourClic(26): End Sub
Private Sub lblJ27_Click(): Call CalendrierJourClic(27): End Sub
Private Sub lblJ28_Click(): Call CalendrierJourClic(28): End Sub
Private Sub lblJ29_Click(): Call CalendrierJourClic(29): End Sub
Private Sub lblJ30_Click(): Call CalendrierJourClic(30): End Sub
Private Sub lblJ31_Click(): Call CalendrierJourClic(31): End Sub
Private Sub lblJ32_Click(): Call CalendrierJourClic(32): End Sub
Private Sub lblJ33_Click(): Call CalendrierJourClic(33): End Sub
Private Sub lblJ34_Click(): Call CalendrierJourClic(34): End Sub
Private Sub lblJ35_Click(): Call CalendrierJourClic(35): End Sub
Private Sub lblJ36_Click(): Call CalendrierJourClic(36): End Sub
Private Sub lblJ37_Click(): Call CalendrierJourClic(37): End Sub
Private Sub lblJ38_Click(): Call CalendrierJourClic(38): End Sub
Private Sub lblJ39_Click(): Call CalendrierJourClic(39): End Sub
Private Sub lblJ40_Click(): Call CalendrierJourClic(40): End Sub
Private Sub lblJ41_Click(): Call CalendrierJourClic(41): End Sub

'==============================================================================
' ONGLET PARAMETRES
'==============================================================================
Private Sub btnEnregistrerParams_Click()
    Call Module_UserForm.UF_EnregistrerParametres(Me)
End Sub

'==============================================================================
' FERMETURE
'==============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Save
End Sub

