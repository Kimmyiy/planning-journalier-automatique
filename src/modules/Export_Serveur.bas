Option Explicit

'==============================================================================
' MODULE : Export_Serveur
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'
' Description :
'   Gestion de l'export du fichier Planning vers le serveur d'entreprise.
'   Une confirmation par saisie de code évite tout export déclenché par un
'   clic accidentel.
'   IMPORTANT : ce code est affiché en clair dans la boîte de dialogue —
'   ce n'est PAS un contrôle de sécurité ni une protection d'accès au
'   serveur, seulement un garde-fou anti-clic accidentel. Si un contrôle
'   d'accès réel au partage réseau est nécessaire, il doit être géré par
'   les droits Windows/AD sur CHEMIN_SERVEUR, pas par ce code VBA.
'   Génère automatiquement README.txt et Prompt_IA.txt lors de chaque export.
'
'   NE PAS MODIFIER ce module sauf pour :
'     - Changer le chemin serveur (CHEMIN_SERVEUR)
'     - Changer le nom du fichier (NOM_FICHIER_SERVEUR)
'     - Changer le code de confirmation (CODE_CONFIRMATION_EXPORT)
'     - Mettre à jour la documentation (ContenuReadme / ContenuPromptIA)
'==============================================================================

'------------------------------------------------------------------------------
' CONSTANTES — Adapter si le chemin ou le code de confirmation change
' ATTENTION : CHEMIN_SERVEUR contient "Test Planning 2.0.0" — à vérifier
' avant un passage en production (chemin de test vs chemin définitif).
'------------------------------------------------------------------------------
Private Const CHEMIN_SERVEUR              As String = "\\scmaas01\Partage\BOLO_Files\01_Administration\03_Commun\Test Planning 2.0.0\"
Private Const NOM_FICHIER_SERVEUR         As String = "Planning.xlsm"
Private Const CODE_CONFIRMATION_EXPORT    As String = "UpDate"

'==============================================================================
' SUB : ExporterVersServeur
'==============================================================================
Public Sub ExporterVersServeur()

    On Error GoTo ErrExport

    Dim saisie As String
    saisie = InputBox("Pour valider l'action, veuillez ecrire """ & CODE_CONFIRMATION_EXPORT & """" & vbCrLf & vbCrLf & _
                      "Cette action exportera le fichier vers le serveur" & vbCrLf & _
                      "et ecrasera la version existante.", _
                      "Export vers le serveur")

    If saisie = "" Then
        MsgBox "Export annulé.", vbInformation, "Annulé"
        Exit Sub
    End If

    If Trim(saisie) <> CODE_CONFIRMATION_EXPORT Then
        MsgBox "Code de confirmation incorrect. Export annulé.", vbExclamation, "Export annulé"
        Exit Sub
    End If

    Application.StatusBar = "Sauvegarde en cours..."
    ThisWorkbook.Save

    Application.StatusBar = "Vérification du serveur en cours..."

    If Dir(CHEMIN_SERVEUR, vbDirectory) = "" Then
        MsgBox "Impossible d'accéder au serveur :" & vbCrLf & CHEMIN_SERVEUR & vbCrLf & vbCrLf & _
               "Vérifiez que vous êtes connecté au réseau d'entreprise.", _
               vbCritical, "Serveur inaccessible"
        Application.StatusBar = False
        Exit Sub
    End If

    Dim dossierArchives As String
    dossierArchives = CHEMIN_SERVEUR & "Archives"
    If Dir(dossierArchives, vbDirectory) = "" Then MkDir dossierArchives

    ' Nom d'archive horodaté (date ET heure) pour ne jamais écraser une
    ' archive précédente en cas de double export le même jour — sans quoi
    ' un deuxième export efface silencieusement la trace du premier.
    Dim nomArchive As String
    nomArchive = "Planning_Sauvegarde_" & Format(Now, "YYYY-MM-DD_hh.nn.ss") & ".xlsm"
    Dim cheminArchive As String
    cheminArchive = dossierArchives & "\" & nomArchive

    On Error Resume Next
    ThisWorkbook.SaveCopyAs cheminArchive
    If Err.Number <> 0 Then
        MsgBox "Attention : impossible de créer l'archive sur le serveur." & vbCrLf & _
               Err.Description, vbExclamation, "Avertissement"
        Err.Clear
    End If
    On Error GoTo 0

    Application.StatusBar = "Export vers le serveur en cours..."

    Dim cheminServeur As String
    cheminServeur = CHEMIN_SERVEUR & NOM_FICHIER_SERVEUR

    On Error Resume Next
    ThisWorkbook.SaveCopyAs cheminServeur
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'export vers le serveur :" & vbCrLf & _
               Err.Description, vbCritical, "Erreur export"
        Err.Clear
        Application.StatusBar = False
        Exit Sub
    End If
    On Error GoTo 0

    Application.StatusBar = "Generation de la documentation..."
    Call ExporterDocumentation

    ' Copie de la documentation vers le serveur — les échecs sont
    ' maintenant vérifiés individuellement pour ne jamais annoncer un
    ' export "réussi" alors qu'un des deux fichiers n'a pas pu être copié.
    Application.StatusBar = "Copie de la documentation vers le serveur..."
    Dim erreursDoc As String
    erreursDoc = ""

    On Error Resume Next
    Err.Clear
    FileCopy ThisWorkbook.Path & "\README.txt", CHEMIN_SERVEUR & "README.txt"
    If Err.Number <> 0 Then erreursDoc = erreursDoc & "README.txt" & vbCrLf: Err.Clear
    FileCopy ThisWorkbook.Path & "\Prompt_IA.txt", CHEMIN_SERVEUR & "Prompt_IA.txt"
    If Err.Number <> 0 Then erreursDoc = erreursDoc & "Prompt_IA.txt" & vbCrLf: Err.Clear
    On Error GoTo 0

    Application.StatusBar = False

    If erreursDoc <> "" Then
        MsgBox "Export réussi, mais la copie de ces fichiers a échoué :" & vbCrLf & _
               erreursDoc & vbCrLf & _
               "Serveur  : " & cheminServeur & vbCrLf & _
               "Archive  : " & cheminArchive, _
               vbExclamation, "Export terminé avec avertissement"
    Else
        MsgBox "Export réussi !" & vbCrLf & vbCrLf & _
               "Serveur  : " & cheminServeur & vbCrLf & vbCrLf & _
               "Archive  : " & cheminArchive & vbCrLf & vbCrLf & _
               "Documentation mise à jour (README.txt et Prompt_IA.txt)", _
               vbInformation, "Export terminé"
    End If

    Exit Sub

ErrExport:
    ' Gestion d'erreur globale : sans elle, un échec de ThisWorkbook.Save,
    ' de MkDir ou de l'écriture de la documentation laisse la barre de
    ' statut Excel bloquée sur un message intermédiaire indéfiniment.
    Application.StatusBar = False
    MsgBox "Erreur lors de l'export vers le serveur :" & vbCrLf & Err.Description, vbCritical

End Sub

'==============================================================================
' SUB : ExporterDocumentation
'==============================================================================
Private Sub ExporterDocumentation()

    Dim dossierLocal As String
    dossierLocal = ThisWorkbook.Path & "\"
    Dim dateMaj As String
    dateMaj = Format(Now, "dd.mm.yyyy hh:mm")

    Dim ff As Integer
    ff = FreeFile
    Open dossierLocal & "README.txt" For Output As #ff
    Print #ff, ContenuReadme(dateMaj)
    Close #ff

    ff = FreeFile
    Open dossierLocal & "Prompt_IA.txt" For Output As #ff
    Print #ff, ContenuPromptIA(dateMaj)
    Close #ff

End Sub

'==============================================================================
' FUNCTION : ContenuReadme
'==============================================================================
Private Function ContenuReadme(ByVal dateMaj As String) As String

    Dim t As String

    t = "================================================================================" & vbCrLf
    t = t & " README - PLANNING JOURNALIER AUTOMATIQUE v2.0.0" & vbCrLf
    t = t & " Derniere mise a jour : " & dateMaj & vbCrLf
    t = t & "================================================================================" & vbCrLf
    t = t & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & " 1. DESCRIPTION DU PROJET" & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & "Ce fichier Excel automatise la génération du planning journalier des opératrices" & vbCrLf
    t = t & "sur les machines de production. Il affecte chaque opératrice presente à un poste" & vbCrLf
    t = t & "en respectant leurs aptitudes et en équilibrant les passages grace aux" & vbCrLf
    t = t & "statistiques historiques. Gère aussi les auxiliaires de weekend." & vbCrLf
    t = t & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & " 2. STRUCTURE DU FICHIER" & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & "  Feuilles :" & vbCrLf
    t = t & "  * Planning journalier  - Feuille principale. Planning du jour." & vbCrLf
    t = t & "  * Personnel            - Source unique : horaires, aptitudes, statistiques" & vbCrLf
    t = t & "  * Vacances             - Absences, congés et maladies du personnel fixe" & vbCrLf
    t = t & "  * Feries               - Jours fériés Neuchatel + fermetures entreprise" & vbCrLf
    t = t & "  * Remplacements        - Remplacements des auxiliaires de weekend" & vbCrLf
    t = t & "  * Groupes              - Calendrier G1/G2 par weekend et jour férié" & vbCrLf
    t = t & "  * Parametres           - Seuils quantité et fermeture machines (Tbl_Parametres)" & vbCrLf
    t = t & "  * Listes               - Listes de référence" & vbCrLf
    t = t & "  * Index des versions   - Historique de toutes les modifications" & vbCrLf
    t = t & vbCrLf
    t = t & "  Tableaux Excel nommes :" & vbCrLf
    t = t & "  * Tbl_Personnel        - Feuille Personnel" & vbCrLf
    t = t & "  * Tbl_Vacances         - Feuille Vacances" & vbCrLf
    t = t & "  * Tbl_Feries           - Feuille Feries" & vbCrLf
    t = t & "  * Tbl_Remplacements    - Feuille Remplacements" & vbCrLf
    t = t & "  * Dates_Groupes        - Feuille Groupes" & vbCrLf
    t = t & "  * Tbl_Parametres       - Feuille Parametres" & vbCrLf
    t = t & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & " 3. PLAGES NOMMEES IMPORTANTES" & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & "  Date_planning          - Date du planning a générer" & vbCrLf
    t = t & "  Date_print             - Horodatage dernière impression validée" & vbCrLf
    t = t & "  MAJ_Stats              - Date derniere mise à jour statistiques" & vbCrLf
    t = t & "  Ref_GroupeG1           - Lundi de référence pour le groupe G1" & vbCrLf
    t = t & "  Planning_Matin         - F6:G20" & vbCrLf
    t = t & "  Planning_ApresMidi     - F24:G27" & vbCrLf
    t = t & "  EnEpreuve_VaoX5A/B/C/D - Quantités en épreuve par machine" & vbCrLf
    t = t & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & " 4. BOUTONS ET LEURS ACTIONS" & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & "  [Generer le planning]  - Calcule et affiche les affectations du jour" & vbCrLf
    t = t & "  [Print]                - Imprime, archive PDF et met à jour les statistiques" & vbCrLf
    t = t & "                           + export PDF stats hebdo automatique chaque lundi" & vbCrLf
    t = t & "  [< / >]                - Navigation dans les dates" & vbCrLf
    t = t & "  [Aujourd'hui]          - Remet la date du jour" & vbCrLf
    t = t & "  [Gestion Personnel]    - Ouvre le UserForm de gestion (5 onglets)" & vbCrLf
    t = t & "  [Exporter vers serveur]- Envoie le fichier sur le serveur (protege MDP)" & vbCrLf
    t = t & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & " 5. MODULES VBA" & vbCrLf
    t = t & "--------------------------------------------------------------------------------" & vbCrLf
    t = t & "  Module_Constantes.bas  - Toutes les constantes du projet (feuilles, tableaux)" & vbCrLf
    t = t & "  Module_Types.bas       - Types publics partagés (JourCalendrier)" & vbCrLf
    t = t & "  Planning_Auto.bas      - Génération du planning, affectations, statistiques" & vbCrLf
    t = t & "  Commande_Bouton.bas    - Boutons interface (dates, impression, archivage PDF)" & vbCrLf
    t = t & "  Module_Feries.bas      - Fériés Neuchâtel, fermetures, grille calendrier" & vbCrLf
    t = t & "  Module_Groupes.bas     - Groupes G1/G2 par weekend et férié (ListObject)" & vbCrLf
    t = t & "  Module_UserForm.bas    - Logique métier de tous les onglets du UserForm" & vbCrLf
    t = t & "  Module_StatsHebdo.bas  - Export PDF stats hebdomadaires automatique (lundi)" & vbCrLf
    t = t & "  Module_Renfort         - Gestion renforts auxiliaires en semaine" & vbCrLf
    t = t & "  Export_Serveur.bas     - Export serveur + génération documentation" & vbCrLf
    t = t & vbCrLf
    t = t & "  UserForms :" & vbCrLf
    t = t & "  UserForm_GestionPersonnel  - Formulaire principal (5 onglets)" & vbCrLf
    t = t & "  UserForm_Absence           - Ajout d'une absence depuis le calendrier" & vbCrLf
    t = t & "  UserForm_Remplacement      - Enregistrement d'un remplacement auxiliaire" & vbCrLf
    t = t & "  UserForm_GroupeAux         - Attribution G1/G2 sur le calendrier auxiliaires" & vbCrLf
    t = t & "  UserForm_ImpressionFiche   - Choix des sections avant impression fiche" & vbCrLf
    t = t & vbCrLf
    t = t & "  Archives générées automatiquement :" & vbCrLf
    t = t & "  Archives\aaaa\mm\aaaa.mm.jj\hh.nn.ss.pdf  - PDF du planning imprimé" & vbCrLf
    t = t & "  Archives\Stats\Stats_semaine_aaaa.mm.jj.pdf - Stats hebdo (chaque lundi)" & vbCrLf
    t = t & vbCrLf
    t = t & "================================================================================" & vbCrLf

    ContenuReadme = t

End Function

'==============================================================================
' FUNCTION : ContenuPromptIA
'==============================================================================
Private Function ContenuPromptIA(ByVal dateMaj As String) As String

    Dim t As String

    t = "================================================================================" & vbCrLf
    t = t & " PROMPT DE DEMARRAGE - PLANNING JOURNALIER AUTOMATIQUE v2.0.0" & vbCrLf
    t = t & " A utiliser tel quel pour reprendre le projet avec Claude (Anthropic)" & vbCrLf
    t = t & " Dernière mise à jour : " & dateMaj & vbCrLf
    t = t & "================================================================================" & vbCrLf
    t = t & vbCrLf
    t = t & "Bonjour," & vbCrLf
    t = t & vbCrLf
    t = t & "Je travaille sur un projet VBA Excel appelé ""Planning Journalier Automatique""" & vbCrLf
    t = t & "développe avec l'aide de Claude (Anthropic). Je suis le seul développeur." & vbCrLf
    t = t & "Tous les échanges sont en francais." & vbCrLf
    t = t & vbCrLf
    t = t & "--- CONTEXTE METIER ---" & vbCrLf
    t = t & "Affectation journalière d'opératrices sur des machines de production." & vbCrLf
    t = t & "  - Personnel fixe (semaine) et auxiliaires (weekends et jours feries)" & vbCrLf
    t = t & "  - Aptitudes par machine (Vaox5-A/B/C postes 1/2/3, Vaox5-D, Vaox1 A&B)" & vbCrLf
    t = t & "  - Equilibrage des passages via statistiques historiques" & vbCrLf
    t = t & "  - Double seuil de quantité pour 1/2/3 personnes par machine" & vbCrLf
    t = t & "  - Tri dynamique v1.2.3 : machine la moins dotée remplie en premier" & vbCrLf
    t = t & "  - Tirage aléatoire pour les ex-aequo de statistiques" & vbCrLf
    t = t & "  - Groupes auxiliaires G1/G2 en alternance par weekend (tournus annuel)" & vbCrLf
    t = t & "  - Export PDF stats hebdomadaires automatique chaque lundi" & vbCrLf
    t = t & "  - Renforts auxiliaires en semaine enregistrés dans Tbl_Remplacements" & vbCrLf
    t = t & "  - Notification manque de personnel sur 14 jours à chaque génération" & vbCrLf
    t = t & vbCrLf
    t = t & "--- ARCHITECTURE v2.0.0 ---" & vbCrLf
    t = t & "  Feuilles Excel :" & vbCrLf
    t = t & "    Planning journalier, Personnel, Vacances, Fériés," & vbCrLf
    t = t & "    Remplacements, Groupes, Paramètres, Listes, Index des versions" & vbCrLf
    t = t & vbCrLf
    t = t & "  Tableaux nommés :" & vbCrLf
    t = t & "    Tbl_Personnel, Tbl_Vacances, Tbl_Feries, Tbl_Remplacements," & vbCrLf
    t = t & "    Dates_Groupes, Tbl_Parametres" & vbCrLf
    t = t & "    IMPORTANT : toutes les écritures dans ces tableaux se font via" & vbCrLf
    t = t & "    ListObject.ListRows.Add (jamais via ws.Cells directement)" & vbCrLf
    t = t & vbCrLf
    t = t & "  Modules VBA :" & vbCrLf
    t = t & "    Module_Constantes  - Toutes les constantes (feuilles, tableaux, types)" & vbCrLf
    t = t & "    Module_Types       - Type public JourCalendrier" & vbCrLf
    t = t & "    Planning_Auto      - Génération planning, affectation, stats" & vbCrLf
    t = t & "                         GenererPlanning(), ChargerPersonnes()," & vbCrLf
    t = t & "                         ConstruireListePostes(), ChoisirPersonne()," & vbCrLf
    t = t & "                         IncrementStats(), EcrireAffectationsMatin()" & vbCrLf
    t = t & "    Commande_Bouton    - Boutons dates, impression + auto-save + stats hebdo" & vbCrLf
    t = t & "                         ImprimerPlanning(), DatePlusUn/MoinsUn()" & vbCrLf
    t = t & "    Module_Feries      - Fériés Neuchâtel, fermetures entreprise" & vbCrLf
    t = t & "                         ChargerFeriesAnnee(), EstFerie(), EstFermeture()," & vbCrLf
    t = t & "                         GenererMoisCalendrier(), AjouterFermeturePeriode()" & vbCrLf
    t = t & "                         Ecriture via Tbl_Feries (ListObject)" & vbCrLf
    t = t & "    Module_Groupes     - Groupes G1/G2 par weekend et férié" & vbCrLf
    t = t & "                         GetGroupeJour(), SetGroupeJour()," & vbCrLf
    t = t & "                         GenererGroupesAnnee()" & vbCrLf
    t = t & "                         Ecriture via Dates_Groupes (ListObject)" & vbCrLf
    t = t & "    Module_UserForm    - Logique métier de tous les onglets du UserForm" & vbCrLf
    t = t & "                         Calendriers personnel et auxiliaire," & vbCrLf
    t = t & "                         absences, remplacements, impression fiche" & vbCrLf
    t = t & "                         Ecriture via Tbl_Remplacements (ListObject)" & vbCrLf
    t = t & "    Module_StatsHebdo  - PDF stats hebdomadaires (chaque lundi à l'impression)" & vbCrLf
    t = t & "                         ImprimerStatsHebdo() - généré _Stats_Temp puis PDF" & vbCrLf
    t = t & "                         Archive dans Archives\Stats\" & vbCrLf
    t = t & "    Module intégré dans Planning_Auto et Module_UserForm :" & vbCrLf
    t = t & "    Renforts auxiliaires en semaine (type Renfort dans Tbl_Remplacements)" & vbCrLf
    t = t & "    Notification effectif insuffisant sur 14 jours (cellule Effectif_Minimum)" & vbCrLf
    t = t & "    VerifierEffectifSemaine(), CompterFixesPresents(), CompterRenforts()" & vbCrLf
    t = t & "    Export_Serveur     - Export serveur protege MDP + documentation" & vbCrLf
    t = t & vbCrLf
    t = t & "  UserForms :" & vbCrLf
    t = t & "    UserForm_GestionPersonnel  - 5 onglets : Personnel / Auxiliaires /" & vbCrLf
    t = t & "                                Absences / Calendrier / Parametres" & vbCrLf
    t = t & "    UserForm_Absence           - Ajout absence depuis calendrier personnel" & vbCrLf
    t = t & "    UserForm_Remplacement      - Enregistrement remplacement auxiliaire" & vbCrLf
    t = t & "    UserForm_GroupeAux         - Attribution G1/G2 sur calendrier auxiliaires" & vbCrLf
    t = t & "    UserForm_ImpressionFiche   - Choix sections avant impression fiche" & vbCrLf
    t = t & vbCrLf
    t = t & "--- CONVENTIONS DE CODE ---" & vbCrLf
    t = t & "  - Option Explicit dans tous les modules" & vbCrLf
    t = t & "  - Toutes les constantes centralisées dans Module_Constantes" & vbCrLf
    t = t & "  - Colonnes référencees par constantes (PERS_COL_xxx, PARAM_xxx)" & vbCrLf
    t = t & "  - Ecriture dans les tableaux TOUJOURS via ListObject.ListRows.Add" & vbCrLf
    t = t & "  - Jamais via ws.Cells() pour écrire dans un ListObject" & vbCrLf
    t = t & "  - Commentaires en francais avec accents" & vbCrLf
    t = t & "  - Gestion d'erreurs avec On Error / MsgBox explicites" & vbCrLf
    t = t & "  - Logique métier dans Module_UserForm, évènements dans le UserForm" & vbCrLf
    t = t & "  - Préférences de livraison : fichiers .bas complets, prêts à coller" & vbCrLf
    t = t & "  - Discuter et valider avant de coder" & vbCrLf
    t = t & vbCrLf
    t = t & "--- VERSION ---" & vbCrLf
    t = t & "  Version courante : v2.0.0 (" & dateMaj & ")" & vbCrLf
    t = t & "  L'index des versions dans Excel détaille toutes les modifications." & vbCrLf
    t = t & vbCrLf
    t = t & "--- POUR CONTINUER LE PROJET ---" & vbCrLf
    t = t & "Je vais te partager les fichiers .bas concernés par la modification souhaitée." & vbCrLf
    t = t & "Merci de :" & vbCrLf
    t = t & "  1. Lire et comprendre le code existant avant de proposer des modifications" & vbCrLf
    t = t & "  2. Discuter et valider la compréhension avant d'écrire du code" & vbCrLf
    t = t & "  3. Conserver les conventions de code listées ci-dessus" & vbCrLf
    t = t & "  4. Me proposer des fichiers .bas complets prêts à copier-coller" & vbCrLf
    t = t & "  5. Me lister les changements pour l'Index des versions" & vbCrLf
    t = t & vbCrLf
    t = t & "================================================================================" & vbCrLf

    ContenuPromptIA = t

End Function