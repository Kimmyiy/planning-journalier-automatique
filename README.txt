================================================================================
 README - PLANNING JOURNALIER AUTOMATIQUE v2.0.0
 Derniere mise a jour : 11.06.2026 14:46
================================================================================

--------------------------------------------------------------------------------
 1. DESCRIPTION DU PROJET
--------------------------------------------------------------------------------
Ce fichier Excel automatise la génération du planning journalier des opératrices
sur les machines de production. Il affecte chaque opératrice presente à un poste
en respectant leurs aptitudes et en équilibrant les passages grace aux
statistiques historiques. Gère aussi les auxiliaires de weekend.

--------------------------------------------------------------------------------
 2. STRUCTURE DU FICHIER
--------------------------------------------------------------------------------
  Feuilles :
  * Planning journalier  - Feuille principale. Planning du jour.
  * Personnel            - Source unique : horaires, aptitudes, statistiques
  * Vacances             - Absences, congés et maladies du personnel fixe
  * Feries               - Jours fériés Neuchatel + fermetures entreprise
  * Remplacements        - Remplacements des auxiliaires de weekend
  * Groupes              - Calendrier G1/G2 par weekend et jour férié
  * Parametres           - Seuils quantité et fermeture machines (Tbl_Parametres)
  * Listes               - Listes de référence
  * Index des versions   - Historique de toutes les modifications

  Tableaux Excel nommes :
  * Tbl_Personnel        - Feuille Personnel
  * Tbl_Vacances         - Feuille Vacances
  * Tbl_Feries           - Feuille Feries
  * Tbl_Remplacements    - Feuille Remplacements
  * Dates_Groupes        - Feuille Groupes
  * Tbl_Parametres       - Feuille Parametres

--------------------------------------------------------------------------------
 3. PLAGES NOMMEES IMPORTANTES
--------------------------------------------------------------------------------
  Date_planning          - Date du planning a générer
  Date_print             - Horodatage dernière impression validée
  MAJ_Stats              - Date derniere mise à jour statistiques
  Ref_GroupeG1           - Lundi de référence pour le groupe G1
  Planning_Matin         - F6:G20
  Planning_ApresMidi     - F24:G27
  EnEpreuve_VaoX5A/B/C/D - Quantités en épreuve par machine

--------------------------------------------------------------------------------
 4. BOUTONS ET LEURS ACTIONS
--------------------------------------------------------------------------------
  [Generer le planning]  - Calcule et affiche les affectations du jour
  [Print]                - Imprime, archive PDF et met à jour les statistiques
                           + export PDF stats hebdo automatique chaque lundi
  [< / >]                - Navigation dans les dates
  [Aujourd'hui]          - Remet la date du jour
  [Gestion Personnel]    - Ouvre le UserForm de gestion (5 onglets)
  [Exporter vers serveur]- Envoie le fichier sur le serveur (protege MDP)

--------------------------------------------------------------------------------
 5. MODULES VBA
--------------------------------------------------------------------------------
  Module_Constantes.bas  - Toutes les constantes du projet (feuilles, tableaux)
  Module_Types.bas       - Types publics partagés (JourCalendrier)
  Planning_Auto.bas      - Génération du planning, affectations, statistiques
  Commande_Bouton.bas    - Boutons interface (dates, impression, archivage PDF)
  Module_Feries.bas      - Fériés Neuchâtel, fermetures, grille calendrier
  Module_Groupes.bas     - Groupes G1/G2 par weekend et férié (ListObject)
  Module_UserForm.bas    - Logique métier de tous les onglets du UserForm
  Module_StatsHebdo.bas  - Export PDF stats hebdomadaires automatique (lundi)
  Module_Renfort         - Gestion renforts auxiliaires en semaine
  Export_Serveur.bas     - Export serveur + génération documentation

  UserForms :
  UserForm_GestionPersonnel  - Formulaire principal (5 onglets)
  UserForm_Absence           - Ajout d'une absence depuis le calendrier
  UserForm_Remplacement      - Enregistrement d'un remplacement auxiliaire
  UserForm_GroupeAux         - Attribution G1/G2 sur le calendrier auxiliaires
  UserForm_ImpressionFiche   - Choix des sections avant impression fiche

  Archives générées automatiquement :
  Archives\aaaa\mm\aaaa.mm.jj\hh.nn.ss.pdf  - PDF du planning imprimé
  Archives\Stats\Stats_semaine_aaaa.mm.jj.pdf - Stats hebdo (chaque lundi)

================================================================================

