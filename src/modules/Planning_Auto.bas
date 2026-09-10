Option Explicit

'==============================================================================
' MODULE : Planning_Auto
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'
' Description :
'   Module principal de génération du planning journalier.
'   Affecte les opératrices aux machines selon leur présence,
'   leurs aptitudes et leurs statistiques de passages.
'
' Historique :
'   v1.2.3 — Tri dynamique des machines Vaox5 par nombre de personnes
'             requises (croissant) — machine la moins dotée en premier.
'             Vaox1 A&B reste toujours prioritaire.
'   v2.0.0 — Fusion Horaire/Restrictions/Statistiques -> feuille Personnel.
'             Gestion auxiliaires (weekends et jours fériés).
'             Gestion vacances, congés, remplacements.
'             Fériés Neuchâtel via Module_Feries.
'             Double seuil et fermeture machine conservés.
'             Tirage aléatoire ex-aequo conservé.
'             Structure Personne remplace Operatrice.
'==============================================================================

'------------------------------------------------------------------------------
' CONSTANTES — Colonnes de la feuille Planning journalier
' (colonnes/lignes de mise en page, colonnes Personnel/Parametres :
'  centralisées dans Module_Constantes pour éviter toute désynchronisation
'  avec Commande_Bouton, Module_UserForm et Module_StatsHebdo)
'------------------------------------------------------------------------------

' Vaox1 A&B : toujours 1 personne fixe
Private Const NB_VAOX1 As Long = 1

'------------------------------------------------------------------------------
' RÉFÉRENCES GLOBALES AUX FEUILLES
' Initialisées par Mem(), utilisées dans tout le module
'------------------------------------------------------------------------------
Public wsP   As Worksheet  ' Planning journalier
Public wsPer As Worksheet  ' Personnel
Public wsV   As Worksheet  ' Vacances
Public wsRpl As Worksheet  ' Remplacements
Public wsS   As Worksheet  ' Paramètres machines

'------------------------------------------------------------------------------
' STRUCTURE : Personne
' Représente une opératrice disponible pour la journée.
'------------------------------------------------------------------------------
Public Type Personne
    nom          As String   ' Nom complet
    LignePers    As Long     ' Ligne dans la feuille Personnel
    codePresence As String   ' "M" = matin, "J" = journée complète, "A" = après-midi uniquement
End Type

'==============================================================================
' SUB : Mem
' Initialise les références aux feuilles.
' À appeler en début de chaque procédure publique.
'==============================================================================
Sub Mem()
    Set wsP = ThisWorkbook.Sheets(NOM_FEUILLE_PLANNING)
    Set wsPer = ThisWorkbook.Sheets(NOM_FEUILLE_PERSONNEL)
    Set wsV = ThisWorkbook.Sheets(NOM_FEUILLE_VACANCES)
    Set wsRpl = ThisWorkbook.Sheets(NOM_FEUILLE_REMPLACEMENTS)
    Set wsS = ThisWorkbook.Sheets(NOM_FEUILLE_PARAMETRES)
End Sub

'==============================================================================
' SUB : GenererPlanning
' Procédure principale — appelée par le bouton "Générer le planning".
'
' Flux :
'   1. Lit la date du planning
'   2. Détermine le type de jour (semaine / WE / férié)
'   3. Charge les personnes disponibles (fixes ou auxiliaires)
'   4. Construit la liste des postes (tri dynamique v1.2.3)
'   5. Affecte les personnes aux postes (min passages + tirage aléatoire)
'   6. Écrit les affectations dans la feuille Planning journalier
'==============================================================================
Public Sub GenererPlanning()

    On Error GoTo ErrFeuille

    ' Un seul amorçage du générateur aléatoire pour toute la génération
    ' (voir ChoisirPersonne : ne pas appeler Randomize à chaque poste).
    Randomize

    Call Mem

    ' --- Lecture de la date du planning ---
    ' Gestion d'erreur étendue jusqu'à la fin de la procédure : couvre aussi
    ' l'accès aux feuilles et aux plages nommées (Date_planning, EnEpreuve_*),
    ' au cas où l'une d'elles serait renommée ou supprimée.
    Dim planningDate As Date
    planningDate = wsP.Range(CELL_DATE_PLANNING).Value

    ' --- Détermination du type de jour ---
    Dim jourSem           As Long
    Dim estWeekend        As Boolean
    Dim estFerieEnSemaine As Boolean
    Dim estJourAux        As Boolean

    jourSem = Weekday(planningDate, vbMonday)
    estWeekend = (jourSem = 6 Or jourSem = 7)
    estFerieEnSemaine = Module_Feries.estFerie(planningDate) And Not estWeekend
    estJourAux = estWeekend Or estFerieEnSemaine

    ' Avertissement si jour férié en semaine
    If estFerieEnSemaine Then
        MsgBox "Attention : le " & Format(planningDate, "dd.mm.yyyy") & _
               " est un jour férié (" & Module_Feries.GetNomFerie(planningDate) & ")." & vbCrLf & _
               "Le planning sera généré avec les auxiliaires.", _
               vbInformation, "Jour férié"
    End If

    ' --- Détermination du groupe auxiliaire actif ---
    Dim groupeActif As String
    If estJourAux Then groupeActif = GroupeWeekend(planningDate)

    ' --- Chargement des personnes disponibles ---
    Dim personnes()  As Personne
    Dim nbPersonnes  As Long
    nbPersonnes = ChargerPersonnes(planningDate, jourSem, estJourAux, groupeActif, personnes)

    If nbPersonnes = 0 Then
        MsgBox "Aucune personne disponible pour ce jour.", vbExclamation
        Exit Sub
    End If

    ' --- Effacer les affectations précédentes ---
    Call EffacerAffectations

    ' Vérification de l'effectif des 14 prochains jours
    Call VerifierEffectifSemaine(planningDate)

    '==========================================================================
    ' PLANNING MATIN
    '==========================================================================

    ' Construction de la liste des postes (tri dynamique v1.2.3)
    Dim postesMatin() As String
    Dim nbPostesMatin As Long
    nbPostesMatin = ConstruireListePostes(postesMatin, personnes, nbPersonnes)

    ' Tableau de disponibilité — une personne absente le matin ("A" = présente
    ' l'après-midi uniquement) n'est pas candidate aux postes machines du matin.
    Dim disponibles() As Boolean
    ReDim disponibles(1 To nbPersonnes)
    Dim i As Long
    For i = 1 To nbPersonnes
        disponibles(i) = (personnes(i).codePresence <> "A")
    Next i

    ' Affectations aux postes machines
    Dim affectationsMatin() As String
    ReDim affectationsMatin(1 To nbPostesMatin)

    Dim p As Long
    For p = 1 To nbPostesMatin

        Dim nomAffecte As String
        nomAffecte = ChoisirPersonne(postesMatin(p), personnes, nbPersonnes, disponibles)
        affectationsMatin(p) = nomAffecte

        ' Marque la personne comme non disponible
        If nomAffecte <> "" Then
            Dim j As Long
            For j = 1 To nbPersonnes
                If personnes(j).nom = nomAffecte Then
                    disponibles(j) = False
                    Exit For
                End If
            Next j
        End If

    Next p

    ' Personnes restantes -> Expédition
    Dim nbExp As Long
    nbExp = 0
    ReDim Preserve affectationsMatin(1 To nbPostesMatin + nbPersonnes)
    ReDim Preserve postesMatin(1 To nbPostesMatin + nbPersonnes)

    For i = 1 To nbPersonnes
        If disponibles(i) Then
            nbPostesMatin = nbPostesMatin + 1
            postesMatin(nbPostesMatin) = "Expedition"
            affectationsMatin(nbPostesMatin) = personnes(i).nom
            nbExp = nbExp + 1
        End If
    Next i

    ' Écriture dans la feuille (ordre d'affichage fixe)
    Dim nbOmisMatin As Long
    nbOmisMatin = EcrireAffectationsMatin(postesMatin, affectationsMatin, nbPostesMatin)

    '==========================================================================
    ' PLANNING APRÈS-MIDI
    '==========================================================================
    wsP.Range("Planning_ApresMidi").ClearContents

    Dim ligneAM As Long
    ligneAM = LIGNE_DEBUT_AM

    Dim nbOmisAM As Long
    nbOmisAM = 0

    For i = 1 To nbPersonnes
        If personnes(i).codePresence = "J" Or personnes(i).codePresence = "A" Then
            If ligneAM > LIGNE_FIN_AM Then
                nbOmisAM = nbOmisAM + 1
            Else
                wsP.Cells(ligneAM, PLAN_COL_NOMS).Value = personnes(i).nom
                ligneAM = ligneAM + 1
            End If
        End If
    Next i

    If nbOmisMatin > 0 Or nbOmisAM > 0 Then
        MsgBox "Planning généré, mais la capacité d'affichage est dépassée :" & vbCrLf & _
               IIf(nbOmisMatin > 0, nbOmisMatin & " personne(s) non affichée(s) le matin." & vbCrLf, "") & _
               IIf(nbOmisAM > 0, nbOmisAM & " personne(s) non affichée(s) l'après-midi." & vbCrLf, "") & _
               "Vérifiez le nombre de personnes présentes ce jour.", _
               vbExclamation, "Planning v2.0 — capacité dépassée"
    Else
        MsgBox "Planning généré avec succès !", vbInformation, "Planning v2.0"
    End If

    Exit Sub

ErrFeuille:
    MsgBox "Erreur lors de la génération du planning :" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Vérifiez que les feuilles et les plages nommées " & _
           "(Date_planning, EnEpreuve_VaoX5A/B/C/D...) existent " & _
           "et sont correctement nommées.", _
           vbCritical

End Sub

'==============================================================================
' FUNCTION : GroupeWeekend
' Retourne le groupe actif (G1/G2) pour un jour donné.
' Lit depuis la feuille Groupes via Module_Groupes.
'==============================================================================
Public Function GroupeWeekend(ByVal dateJour As Date) As String

    Dim groupe As String
    groupe = Module_Groupes.GetGroupeJour(dateJour)

    ' Fallback sur calcul automatique si pas dans la feuille Groupes
    If groupe = "" Then
        Dim dateRefG1 As Date
        On Error Resume Next
        dateRefG1 = wsP.Range(CELL_REF_GROUPE_G1).Value
        On Error GoTo 0
        If dateRefG1 = 0 Then
            GroupeWeekend = GROUPE_G1
            Exit Function
        End If
        Dim jourSem As Long
        jourSem = Weekday(dateJour, vbMonday)
        Dim lundiSem As Date
        lundiSem = dateJour - (jourSem - 1)
        Dim nbSem As Long
        nbSem = Int((lundiSem - dateRefG1) / 7)
        groupe = IIf(nbSem Mod 2 = 0, GROUPE_G1, GROUPE_G2)
    End If

    GroupeWeekend = groupe

End Function

'==============================================================================
' FUNCTION : ChargerPersonnes
' Charge les personnes disponibles pour la journée.
' Fixe (semaine) ou Auxiliaires (weekend/férié) selon le type de jour.
' Gère les absences, remplacements et congés.
'==============================================================================
Public Function ChargerPersonnes(ByVal planningDate As Date, _
                                   ByVal jourSem As Long, _
                                   ByVal estJourAux As Boolean, _
                                   ByVal groupeActif As String, _
                                   ByRef personnes() As Personne) As Long

    Dim dernLigne As Long
    dernLigne = wsPer.Cells(wsPer.Rows.count, PERS_COL_NOM).End(xlUp).Row

    ReDim personnes(1 To dernLigne)
    Dim count As Long
    count = 0

    Dim i As Long
    For i = 2 To dernLigne

        ' Filtre statut Actif
        If Trim(wsPer.Cells(i, PERS_COL_STATUT).Value) <> "Actif" Then GoTo SuivantPers

        Dim typePers  As String
        Dim groupePers As String
        typePers = Trim(wsPer.Cells(i, PERS_COL_TYPE).Value)
        groupePers = Trim(wsPer.Cells(i, PERS_COL_GROUPE).Value)

        Dim nomPers As String
        nomPers = wsPer.Cells(i, PERS_COL_NOM).Value

        If estJourAux Then
            ' Jour WE ou férié : uniquement les auxiliaires du groupe actif
            ' + les remplaçants enregistrés (l'absence éventuelle est déjà
            ' vérifiée dans EstDisponibleAuxiliaire)
            If typePers <> "Auxiliaire" Then GoTo SuivantPers
            If Not EstDisponibleAuxiliaire(nomPers, planningDate, groupePers, groupeActif) Then
                GoTo SuivantPers
            End If

            count = count + 1
            personnes(count).nom = nomPers
            personnes(count).LignePers = i
            personnes(count).codePresence = "M"

        ElseIf typePers = "Fixe" Then
            ' Jour de semaine : les fixes présents selon leur horaire, en
            ' tenant compte d'une éventuelle absence par demi-journée.
            Dim colHoraire As Long
            colHoraire = PERS_COL_LUN + (jourSem - 1)
            Dim codeH As String
            codeH = Trim(wsPer.Cells(i, colHoraire).Value)

            If codeH <> "M" And codeH <> "J" Then GoTo SuivantPers

            Dim codeFinalFixe As String
            codeFinalFixe = CodePresenceApresAbsence(codeH, ObtenirPeriodeAbsence(nomPers, planningDate))
            If codeFinalFixe = "" Then GoTo SuivantPers

            count = count + 1
            personnes(count).nom = nomPers
            personnes(count).LignePers = i
            personnes(count).codePresence = codeFinalFixe

        ElseIf typePers = "Auxiliaire" Then
            ' Renfort auxiliaire en semaine (enregistré dans Tbl_Remplacements,
            ' individuellement ou via "Ajouter un groupe pour la semaine").
            If Not EstEnRenfort(nomPers, planningDate) Then GoTo SuivantPers

            Dim codeFinalRenfort As String
            codeFinalRenfort = CodePresenceApresAbsence("J", ObtenirPeriodeAbsence(nomPers, planningDate))
            If codeFinalRenfort = "" Then GoTo SuivantPers

            count = count + 1
            personnes(count).nom = nomPers
            personnes(count).LignePers = i
            personnes(count).codePresence = codeFinalRenfort
        End If

SuivantPers:
    Next i

    ChargerPersonnes = count

End Function

'==============================================================================
' FUNCTION : ObtenirPeriodeAbsence
' Retourne la période d'absence (PERIODE_JOURNEE/MATIN/APRESMIDI) enregistrée
' pour cette personne à cette date, ou "" si elle n'est pas absente.
' Une ligne sans valeur de période (absences saisies avant l'ajout de cette
' colonne) est traitée comme une absence de la journée complète.
'==============================================================================
Private Function ObtenirPeriodeAbsence(ByVal nom As String, ByVal dateJour As Date) As String

    Dim dernLigne As Long
    dernLigne = wsV.Cells(wsV.Rows.count, VAC_COL_NOM).End(xlUp).Row

    Dim i As Long
    For i = 2 To dernLigne
        If Trim(wsV.Cells(i, VAC_COL_NOM).Value) = nom Then
            ' dDebut/dFin réinitialisées à 0 avant chaque lecture : si la
            ' cellule est vide/invalide, on ignore la ligne au lieu de
            ' réutiliser silencieusement la date de l'itération précédente.
            Dim dDebut As Date, dFin As Date
            dDebut = 0: dFin = 0
            On Error Resume Next
            dDebut = wsV.Cells(i, VAC_COL_DEBUT).Value
            dFin = wsV.Cells(i, VAC_COL_FIN).Value
            On Error GoTo 0
            If dDebut <> 0 And dFin <> 0 Then
                If dateJour >= dDebut And dateJour <= dFin Then
                    Dim periode As String
                    periode = Trim(wsV.Cells(i, VAC_COL_PERIODE).Value)
                    If periode = "" Then periode = PERIODE_JOURNEE
                    ObtenirPeriodeAbsence = periode
                    Exit Function
                End If
            End If
        End If
    Next i

    ObtenirPeriodeAbsence = ""

End Function

'==============================================================================
' FUNCTION : EstAbsent
' Retourne True si la personne a une absence enregistrée pour cette date,
' quelle que soit la période (journée complète ou demi-journée).
'==============================================================================
Private Function EstAbsent(ByVal nom As String, ByVal dateJour As Date) As Boolean
    EstAbsent = (ObtenirPeriodeAbsence(nom, dateJour) <> "")
End Function

'==============================================================================
' FUNCTION : CodePresenceApresAbsence
' Combine l'horaire de base ("M" ou "J") avec une période d'absence
' éventuelle pour déterminer la présence réelle du jour :
'   - absente toute la journée, ou absente sur le seul moment où elle
'     travaille -> retourne "" (la personne est exclue du planning)
'   - absente le matin seulement (horaire "J") -> "A" (présente l'après-midi)
'   - absente l'après-midi seulement -> "M" (présente le matin uniquement)
'   - pas d'absence -> l'horaire de base inchangé
'==============================================================================
Private Function CodePresenceApresAbsence(ByVal codeBase As String, ByVal periode As String) As String
    Select Case periode
        Case PERIODE_JOURNEE
            CodePresenceApresAbsence = ""
        Case PERIODE_MATIN
            CodePresenceApresAbsence = IIf(codeBase = "M", "", "A")
        Case PERIODE_APRESMIDI
            CodePresenceApresAbsence = "M"
        Case Else
            CodePresenceApresAbsence = codeBase
    End Select
End Function

'==============================================================================
' FUNCTION : ConstruireListePostes
' Détermine les postes machines à pourvoir (sans Expédition).
'
' Tri dynamique v1.2.3 :
'   Les machines Vaox5 sont remplies par ordre croissant du nombre
'   de personnes requises. La machine qui nécessite le moins de monde
'   est remplie en premier afin de garantir les affectations sur les
'   petits postes avant de saturer les candidats sur les gros.
'   Vaox1 A&B reste toujours en première position (priorité absolue).
'
' Calcul du nombre de personnes par machine :
'   Fermée           -> 0 personne
'   0 < qté < Seuil1 -> 1 personne
'   Seuil1 =< qté < Seuil2 -> 2 personnes
'   qté => Seuil2    -> 3 personnes
'==============================================================================
Private Function ConstruireListePostes(ByRef postes() As String, _
                                        ByRef personnes() As Personne, _
                                        ByVal nbPersonnes As Long) As Long

    ' --- Lecture des quantités en épreuve ---
    Dim qteA As Long: qteA = Val(wsP.Range(CELL_QTE_VAOX5A).Value)
    Dim qteB As Long: qteB = Val(wsP.Range(CELL_QTE_VAOX5B).Value)
    Dim qteC As Long: qteC = Val(wsP.Range(CELL_QTE_VAOX5C).Value)
    Dim qteD As Long: qteD = Val(wsP.Range(CELL_QTE_VAOX5D).Value)

    ' --- Lecture des paramètres depuis Tbl_Parametres ---
    Dim s1A As Long, s2A As Long, maxA As Long, fermA As Boolean
    Dim s1B As Long, s2B As Long, maxB As Long, fermB As Boolean
    Dim s1C As Long, s2C As Long, maxC As Long, fermC As Boolean
    Dim s1D As Long, s2D As Long, maxD As Long, fermD As Boolean
    Dim fermV1 As Boolean, maxV1 As Long

    If Not LireParametresMachine(MACHINE_VAOX5A, s1A, s2A, maxA, fermA) Then Exit Function
    If Not LireParametresMachine(MACHINE_VAOX5B, s1B, s2B, maxB, fermB) Then Exit Function
    If Not LireParametresMachine(MACHINE_VAOX5C, s1C, s2C, maxC, fermC) Then Exit Function
    If Not LireParametresMachine(MACHINE_VAOX5D, s1D, s2D, maxD, fermD) Then Exit Function
    If Not LireParametresMachine(PARAM_VAOX1, 0, 0, maxV1, fermV1) Then Exit Function

    ' --- Calcul du nombre de personnes par machine ---
    Dim nbVaox1 As Long: nbVaox1 = IIf(fermV1, 0, NB_VAOX1)
    Dim nbA As Long: nbA = IIf(fermA, 0, CalculerNbPersonnes(qteA, s1A, s2A))
    Dim nbB As Long: nbB = IIf(fermB, 0, CalculerNbPersonnes(qteB, s1B, s2B))
    Dim nbC As Long: nbC = IIf(fermC, 0, CalculerNbPersonnes(qteC, s1C, s2C))
    Dim nbD As Long: nbD = IIf(fermD, 0, CalculerNbPersonnes(qteD, s1D, s2D))

    ' --- Construction de la liste ---
    Dim maxPostes As Long
    maxPostes = nbVaox1 + nbA + nbB + nbC + nbD
    If maxPostes = 0 Then
        ' Base 1-based, cohérente avec le ReDim Preserve fait ensuite dans
        ' GenererPlanning : ReDim postes(0) créerait un tableau 0-based et
        ' ReDim Preserve postes(1 To ...) plante alors (erreur 9, la borne
        ' inférieure d'un tableau ne peut pas être changée avec Preserve).
        ReDim postes(1 To 1)
        ConstruireListePostes = 0
        Exit Function
    End If

    ReDim postes(1 To maxPostes)
    Dim idx As Long
    idx = 1

    ' Vaox1 A&B : priorité absolue, toujours en premier
    Dim k As Long
    For k = 1 To nbVaox1
        postes(idx) = MACHINE_VAOX1
        idx = idx + 1
    Next k

    ' --- Tri dynamique Vaox5 : ordre croissant par nb personnes requises ---
    Dim codesMachines(1 To 4)      As String
    Dim nbPersonnesMachine(1 To 4) As Long

    codesMachines(1) = "A": nbPersonnesMachine(1) = nbA
    codesMachines(2) = "B": nbPersonnesMachine(2) = nbB
    codesMachines(3) = "C": nbPersonnesMachine(3) = nbC
    codesMachines(4) = "D": nbPersonnesMachine(4) = nbD

    ' Tri à bulles (croissant)
    Dim ti As Long, tj As Long
    Dim tmpCode As String, tmpNb As Long

    For ti = 1 To 3
        For tj = 1 To 4 - ti
            If nbPersonnesMachine(tj) > nbPersonnesMachine(tj + 1) Then
                tmpCode = codesMachines(tj)
                codesMachines(tj) = codesMachines(tj + 1)
                codesMachines(tj + 1) = tmpCode
                tmpNb = nbPersonnesMachine(tj)
                nbPersonnesMachine(tj) = nbPersonnesMachine(tj + 1)
                nbPersonnesMachine(tj + 1) = tmpNb
            End If
        Next tj
    Next ti

    ' Ajout des postes dans l'ordre d'affectation
    Dim m As Long
    For m = 1 To 4
        For k = 1 To nbPersonnesMachine(m)
            postes(idx) = "Vaox5-" & codesMachines(m) & " " & k
            idx = idx + 1
        Next k
    Next m

    ConstruireListePostes = idx - 1

End Function

'==============================================================================
' FUNCTION : LireParametresMachine
' Lit les paramètres d'une machine depuis le tableau Tbl_Parametres.
' Retourne False si la machine n'est pas trouvée.
'==============================================================================
Private Function LireParametresMachine(ByVal nomMachine As String, _
                                        ByRef seuil1 As Long, _
                                        ByRef seuil2 As Long, _
                                        ByRef maxVal As Long, _
                                        ByRef ferme As Boolean) As Boolean

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsS.ListObjects(NOM_TBL_PARAMETRES)
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "Tableau '" & NOM_TBL_PARAMETRES & "' introuvable.", vbCritical
        LireParametresMachine = False
        Exit Function
    End If

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        If Trim(tbl.DataBodyRange(i, PARAM_MACHINE).Value) = nomMachine Then
            seuil1 = Val(tbl.DataBodyRange(i, PARAM_SEUIL1).Value)
            seuil2 = Val(tbl.DataBodyRange(i, PARAM_SEUIL2).Value)
            maxVal = Val(tbl.DataBodyRange(i, PARAM_MAX).Value)
            ferme = CBool(tbl.DataBodyRange(i, PARAM_FERMETURE).Value)
            LireParametresMachine = True
            Exit Function
        End If
    Next i

    LireParametresMachine = False

End Function

'==============================================================================
' FUNCTION : CalculerNbPersonnes
' Calcule le nombre de personnes requises selon la quantité et les seuils.
'==============================================================================
Private Function CalculerNbPersonnes(ByVal qte As Long, _
                                      ByVal seuil1 As Long, _
                                      ByVal seuil2 As Long) As Long
    If qte = 0 Then
        CalculerNbPersonnes = 0
    ElseIf qte < seuil1 Then
        CalculerNbPersonnes = 1
    ElseIf qte < seuil2 Then
        CalculerNbPersonnes = 2
    Else
        CalculerNbPersonnes = 3
    End If
End Function

'==============================================================================
' FUNCTION : SommePassagesMachine
' Calcule la somme des passages sur une machine pour les personnes présentes.
' Utilisée pour le tri dynamique de l'ordre d'affectation.
'==============================================================================
Private Function SommePassagesMachine(ByVal colStats As Long, _
                                       ByRef personnes() As Personne, _
                                       ByVal nbPersonnes As Long) As Long
    Dim total As Long
    Dim i As Long
    For i = 1 To nbPersonnes
        If personnes(i).LignePers > 0 Then
            total = total + Val(wsPer.Cells(personnes(i).LignePers, colStats).Value)
        End If
    Next i
    SommePassagesMachine = total
End Function

'==============================================================================
' FUNCTION : ChoisirPersonne
' Choisit la meilleure personne disponible pour un poste.
' Critère : minimum de passages sur la machine.
' Départage : tirage aléatoire entre ex-aequo.
'==============================================================================
Private Function ChoisirPersonne(ByVal poste As String, _
                                   ByRef personnes() As Personne, _
                                   ByVal nbPersonnes As Long, _
                                   ByRef disponibles() As Boolean) As String

    Dim colApt   As Long: colApt = GetColAptitudePourPoste(poste)
    Dim colStats As Long: colStats = GetColStatsPourPoste(poste)

    Dim candidates()    As String
    Dim statsCandidat() As Long
    ReDim candidates(1 To nbPersonnes)
    ReDim statsCandidat(1 To nbPersonnes)
    Dim nbCandidates As Long
    nbCandidates = 0

    Dim i As Long
    For i = 1 To nbPersonnes

        If Not disponibles(i) Then GoTo SuiteCandidats

        ' Vérification de l'aptitude
        If colApt > 0 Then
            If wsPer.Cells(personnes(i).LignePers, colApt).Value <> 1 Then
                GoTo SuiteCandidats
            End If
        End If

        ' Lecture des passages
        Dim passages As Long
        If colStats > 0 Then
            passages = Val(wsPer.Cells(personnes(i).LignePers, colStats).Value)
        Else
            passages = 0
        End If

        nbCandidates = nbCandidates + 1
        candidates(nbCandidates) = personnes(i).nom
        statsCandidat(nbCandidates) = passages

SuiteCandidats:
    Next i

    If nbCandidates = 0 Then
        ChoisirPersonne = ""
        Exit Function
    End If

    ' Minimum de passages
    Dim minPassages As Long
    minPassages = statsCandidat(1)
    Dim c As Long
    For c = 2 To nbCandidates
        If statsCandidat(c) < minPassages Then minPassages = statsCandidat(c)
    Next c

    ' Collecte des ex-aequo
    Dim exAequo()  As String
    ReDim exAequo(1 To nbCandidates)
    Dim nbExAequo As Long
    nbExAequo = 0

    For c = 1 To nbCandidates
        If statsCandidat(c) = minPassages Then
            nbExAequo = nbExAequo + 1
            exAequo(nbExAequo) = candidates(c)
        End If
    Next c

    ' Tirage aléatoire parmi les ex-aequo
    ' Note : Randomize n'est PAS appelé ici mais une seule fois en tête de
    ' GenererPlanning. L'appeler à chaque poste en ex-aequo réamorce le
    ' générateur depuis l'horloge système en quelques millisecondes,
    ' ce qui peut produire des tirages corrélés au lieu de tirages
    ' réellement indépendants.
    If nbExAequo = 1 Then
        ChoisirPersonne = exAequo(1)
    Else
        ChoisirPersonne = exAequo(Int(Rnd() * nbExAequo) + 1)
    End If

End Function

'==============================================================================
' FUNCTION : GetColStatsPourPoste
' Retourne la colonne statistiques dans Personnel pour un poste donné.
'==============================================================================
Private Function GetColStatsPourPoste(ByVal poste As String) As Long
    Dim p As String
    p = UCase(Trim(poste))
    If InStr(p, "VAOX5-A") > 0 Then GetColStatsPourPoste = PERS_COL_STATS_A:     Exit Function
    If InStr(p, "VAOX5-B") > 0 Then GetColStatsPourPoste = PERS_COL_STATS_B:     Exit Function
    If InStr(p, "VAOX5-C") > 0 Then GetColStatsPourPoste = PERS_COL_STATS_C:     Exit Function
    If InStr(p, "VAOX5-D") > 0 Then GetColStatsPourPoste = PERS_COL_STATS_D:     Exit Function
    If InStr(p, "VAOX1") > 0 Then GetColStatsPourPoste = PERS_COL_STATS_V1:      Exit Function
    If InStr(p, "EXPEDITION") > 0 Then GetColStatsPourPoste = PERS_COL_STATS_EXP: Exit Function
    GetColStatsPourPoste = 0
End Function

'==============================================================================
' FUNCTION : GetColAptitudePourPoste
' Retourne la colonne aptitude dans Personnel pour un poste donné.
' Retourne 0 pour Expédition (toujours autorisé).
'==============================================================================
Private Function GetColAptitudePourPoste(ByVal poste As String) As Long
    Dim p As String
    p = UCase(Trim(poste))

    ' Vaox5-A/B/C : aptitude selon le numéro de poste (1, 2 ou 3)
    If InStr(p, "VAOX5-A") > 0 Or _
       InStr(p, "VAOX5-B") > 0 Or _
       InStr(p, "VAOX5-C") > 0 Then
        Dim num As String
        num = Right(Trim(poste), 1)
        Select Case num
            Case "1": GetColAptitudePourPoste = PERS_COL_P1
            Case "2": GetColAptitudePourPoste = PERS_COL_P2
            Case "3": GetColAptitudePourPoste = PERS_COL_P3
            Case Else: GetColAptitudePourPoste = 0
        End Select
        Exit Function
    End If

    If InStr(p, "VAOX5-D") > 0 Then GetColAptitudePourPoste = PERS_COL_D:     Exit Function
    If InStr(p, "VAOX1") > 0 Then GetColAptitudePourPoste = PERS_COL_VAOX1: Exit Function
    GetColAptitudePourPoste = 0  ' Expédition = toujours autorisé

End Function

'==============================================================================
' SUB : EffacerAffectations
' Vide les tableaux Planning_Matin et Planning_ApresMidi.
'==============================================================================
Private Sub EffacerAffectations()
    wsP.Range("Planning_Matin").ClearContents
    wsP.Range("Planning_ApresMidi").ClearContents
End Sub

'==============================================================================
' FUNCTION : EcrireAffectationsMatin
' Écrit les affectations dans la feuille Planning journalier.
' Ordre d'affichage toujours fixe : Vaox1 > A > B > C > D > Expédition,
' indépendamment de l'ordre d'affectation (tri dynamique v1.2.3).
'
' Retourne le nombre d'affectations qui n'ont pas pu être écrites faute de
' place (le tableau Planning_Matin ne compte que NB_LIGNES_MATIN lignes) :
' l'appelant doit avertir l'utilisateur si ce nombre est supérieur à 0,
' pour ne jamais perdre silencieusement une personne présente du planning.
'==============================================================================
Private Function EcrireAffectationsMatin(ByRef postes() As String, _
                                          ByRef affectations() As String, _
                                          ByVal nbPostes As Long) As Long

    Dim ordreAffichage(1 To 6) As String
    ordreAffichage(1) = "VAOX1"
    ordreAffichage(2) = "VAOX5-A"
    ordreAffichage(3) = "VAOX5-B"
    ordreAffichage(4) = "VAOX5-C"
    ordreAffichage(5) = "VAOX5-D"
    ordreAffichage(6) = "EXP"

    Dim ligneCourante As Long
    ligneCourante = LIGNE_DEBUT_MATIN

    Dim nbOmis As Long
    nbOmis = 0

    Dim g As Long
    For g = 1 To 6

        Dim i As Long
        For i = 1 To nbPostes

            Dim posteUpper As String
            posteUpper = UCase(Trim(postes(i)))

            Dim correspond As Boolean
            correspond = False

            Select Case ordreAffichage(g)
                Case "VAOX1":   correspond = InStr(posteUpper, "VAOX1") > 0
                Case "VAOX5-A": correspond = InStr(posteUpper, "VAOX5-A") > 0
                Case "VAOX5-B": correspond = InStr(posteUpper, "VAOX5-B") > 0
                Case "VAOX5-C": correspond = InStr(posteUpper, "VAOX5-C") > 0
                Case "VAOX5-D": correspond = InStr(posteUpper, "VAOX5-D") > 0
                Case "EXP":     correspond = InStr(posteUpper, "EXP") > 0
            End Select

            If correspond Then
                If ligneCourante > LIGNE_DEBUT_MATIN + NB_LIGNES_MATIN - 1 Then
                    nbOmis = nbOmis + 1
                Else
                    wsP.Cells(ligneCourante, PLAN_COL_NOMS).Value = affectations(i)
                    wsP.Cells(ligneCourante, PLAN_COL_POSTES).Value = postes(i)
                    ligneCourante = ligneCourante + 1
                End If
            End If

        Next i

    Next g

    EcrireAffectationsMatin = nbOmis

End Function

'==============================================================================
' SUB : VerifierEffectifSemaine
' Scanne les 14 prochains jours ouvrables depuis la date du planning.
' Affiche un rapport si le nombre de fixes présents < effectif minimum.
' N'empêche pas la génération.
'==============================================================================
Private Sub VerifierEffectifSemaine(ByVal dateDepart As Date)

    ' Lit l'effectif minimum depuis la cellule nommée
    Dim effMin As Long
    On Error Resume Next
    effMin = CLng(wsS.Range("Effectif_Minimum").Value)
    On Error GoTo 0

    If effMin <= 0 Then Exit Sub  ' Paramètre non configuré -> pas de vérification

    Dim rapport    As String
    Dim nbAlertes  As Long
    rapport = ""
    nbAlertes = 0

    Dim nbJoursScanner As Long
    nbJoursScanner = 14

    Dim joursScannés As Long
    joursScannés = 0

    Dim dateCase As Date
    dateCase = dateDepart + 1  ' Commence au lendemain

    Do While joursScannés < nbJoursScanner

        Dim jourSem As Long
        jourSem = Weekday(dateCase, vbMonday)

        ' Uniquement les jours de semaine (lun-ven), pas fériés
        If jourSem >= 1 And jourSem <= 5 And Not Module_Feries.estFerie(dateCase) Then

            joursScannés = joursScannés + 1

            ' Compte les fixes présents ce jour
            Dim nbPresents As Long
            nbPresents = CompterFixesPresents(dateCase, jourSem)

            If nbPresents < effMin Then
                nbAlertes = nbAlertes + 1
                Dim nomJour As String
                Select Case jourSem
                    Case 1: nomJour = "Lundi"
                    Case 2: nomJour = "Mardi"
                    Case 3: nomJour = "Mercredi"
                    Case 4: nomJour = "Jeudi"
                    Case 5: nomJour = "Vendredi"
                End Select

                rapport = rapport & "  " & nomJour & " " & Format(dateCase, "dd.mm.yyyy") & _
                          " : " & nbPresents & " personne(s) / " & effMin & " requises" & vbCrLf
            End If

        End If

        dateCase = dateCase + 1

    Loop

    ' Affiche le rapport si des alertes existent
    If nbAlertes > 0 Then
        MsgBox "Attention — Manque de personnel prévu :" & vbCrLf & vbCrLf & _
               rapport & vbCrLf & _
               "Pensez à prévoir des renforts auxiliaires.", _
               vbExclamation, "Effectif insuffisant"
    End If

End Sub

'==============================================================================
' FUNCTION : CompterFixesPresents
' Compte les fixes présents un jour donné (horaire + absences).
'==============================================================================
Private Function CompterFixesPresents(ByVal dateJour As Date, _
                                       ByVal jourSem As Long) As Long

    Dim dernLigne As Long
    dernLigne = wsPer.Cells(wsPer.Rows.count, PERS_COL_NOM).End(xlUp).Row

    Dim count As Long
    count = 0

    Dim i As Long
    For i = 2 To dernLigne

        If Trim(wsPer.Cells(i, PERS_COL_STATUT).Value) <> "Actif" Then GoTo SuivantFix
        If Trim(wsPer.Cells(i, PERS_COL_TYPE).Value) <> "Fixe" Then GoTo SuivantFix

        ' Vérifie l'horaire du jour
        Dim colH As Long
        colH = PERS_COL_LUN + (jourSem - 1)
        Dim codeH As String
        codeH = Trim(wsPer.Cells(i, colH).Value)
        If codeH <> "M" And codeH <> "J" Then GoTo SuivantFix

        ' Vérifie absence
        If EstAbsent(wsPer.Cells(i, PERS_COL_NOM).Value, dateJour) Then GoTo SuivantFix

        count = count + 1

SuivantFix:
    Next i

    ' Ajoute les renforts auxiliaires déjà planifiés ce jour
    count = count + CompterRenforts(dateJour)

    CompterFixesPresents = count

End Function

'==============================================================================
' FUNCTION : CompterRenforts
' Compte les auxiliaires enregistrés en renfort pour une date donnée.
'==============================================================================
Private Function CompterRenforts(ByVal dateJour As Date) As Long

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsRpl.ListObjects(NOM_TBL_REMPLACEMENTS)
    On Error GoTo 0
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then Exit Function

    Dim count As Long
    count = 0

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        Dim dRpl As Date
        dRpl = 0
        On Error Resume Next
        dRpl = tbl.DataBodyRange(i, RPL_COL_DATE).Value
        On Error GoTo 0
        If dRpl <> 0 And Int(dRpl) = Int(dateJour) Then
            If Trim(tbl.DataBodyRange(i, RPL_COL_TYPE).Value) = "Renfort" Then
                count = count + 1
            End If
        End If
    Next i

    CompterRenforts = count

End Function

'==============================================================================
' FUNCTION : EstEnRenfort
' Retourne True si cette personne est enregistrée comme renfort (ou
' remplaçante) pour cette date dans Tbl_Remplacements. Utilisée pour intégrer
' les auxiliaires en renfort de semaine (individuel ou groupe entier) dans le
' planning, ce que CompterRenforts (qui ne fait que compter) ne fait pas.
'==============================================================================
Private Function EstEnRenfort(ByVal nom As String, ByVal dateJour As Date) As Boolean

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsRpl.ListObjects(NOM_TBL_REMPLACEMENTS)
    On Error GoTo 0
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then Exit Function

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        Dim dRpl As Date
        dRpl = 0
        On Error Resume Next
        dRpl = tbl.DataBodyRange(i, RPL_COL_DATE).Value
        On Error GoTo 0
        If dRpl <> 0 And Int(dRpl) = Int(dateJour) Then
            If Trim(tbl.DataBodyRange(i, RPL_COL_NOM_REMPLACANT).Value) = nom Then
                EstEnRenfort = True
                Exit Function
            End If
        End If
    Next i

End Function

'==============================================================================
' MODIFICATION : EstDisponibleAuxiliaire
' Remplace la version existante dans Planning_Auto.bas
' Ajoute la gestion des renforts en semaine.
'==============================================================================
Private Function EstDisponibleAuxiliaire(ByVal nom As String, _
                                          ByVal dateJour As Date, _
                                          ByVal groupePers As String, _
                                          ByVal groupeActif As String) As Boolean

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsRpl.ListObjects(NOM_TBL_REMPLACEMENTS)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        Dim i As Long
        For i = 1 To tbl.ListRows.count
            Dim dRpl As Date
            dRpl = 0
            On Error Resume Next
            dRpl = tbl.DataBodyRange(i, RPL_COL_DATE).Value
            On Error GoTo 0

            If dRpl <> 0 And Int(dRpl) = Int(dateJour) Then
                Dim typeRpl As String
                typeRpl = Trim(tbl.DataBodyRange(i, RPL_COL_TYPE).Value)

                ' Se fait remplacer -> absent
                If Trim(tbl.DataBodyRange(i, RPL_COL_NOM_ABSENTE).Value) = nom Then
                    EstDisponibleAuxiliaire = False
                    Exit Function
                End If

                ' Est remplaçant ou renfort -> présent
                If Trim(tbl.DataBodyRange(i, RPL_COL_NOM_REMPLACANT).Value) = nom Then
                    EstDisponibleAuxiliaire = True
                    Exit Function
                End If
            End If
        Next i
    End If

    ' Pas de remplacement/renfort -> présent uniquement si son groupe est actif
    ' Pour les jours de semaine normaux, groupeActif = "" donc toujours absent.
    ' Une absence déclarée (onglet Absences) prime sur la présence de groupe :
    ' un auxiliaire peut désormais être simplement absent sans qu'on lui ait
    ' cherché de remplaçante.
    If groupeActif = "" Then
        EstDisponibleAuxiliaire = False
    ElseIf EstAbsent(nom, dateJour) Then
        EstDisponibleAuxiliaire = False
    Else
        EstDisponibleAuxiliaire = (groupePers = groupeActif)
    End If

End Function


'==============================================================================
' SUB : IncrementStats (Public)
' Incrémente le compteur de passages dans Personnel.
' Appelée par Commande_Bouton lors de l'impression.
'==============================================================================
Public Sub IncrementStats(ByVal nom As String, ByVal poste As String)

    Dim colStats As Long
    colStats = GetColStatsPourPoste(poste)
    If colStats = 0 Then Exit Sub

    Dim lignePer As Variant
    lignePer = Application.Match(nom, wsPer.Columns(PERS_COL_NOM), 0)
    If IsError(lignePer) Then Exit Sub

    wsPer.Cells(lignePer, colStats).Value = _
        wsPer.Cells(lignePer, colStats).Value + 1

End Sub