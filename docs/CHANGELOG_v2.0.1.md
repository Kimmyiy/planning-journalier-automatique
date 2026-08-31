# Planning Journalier Automatique — v2.0.1 (revue de code)

Revue de code complète du VBA v2.0.0 (tous les modules, UserForms et classes
de feuilles), à copier dans l'onglet **Index des versions**. Aucun fichier
`.xlsm` n'a été modifié directement — seul le code source `.bas`/`.frm` a été
corrigé, à réimporter dans l'éditeur VBA (Alt+F11 → clic droit sur chaque
module → *Supprimer* puis *Fichier → Importer un fichier...*).

## Corrections critiques (bugs bloquants ou perte de données)

- **Planning_Auto.bas** — `GenererPlanning` pouvait planter (erreur 9,
  "indice hors limites") si aucune machine n'était active un jour donné
  (`ReDim Preserve` changeait la borne inférieure d'un tableau, ce qui est
  interdit en VBA).
- **Module_Feries.bas** — `ChargerFeriesAnnee` effaçait silencieusement
  toutes les fermetures d'entreprise (vacances collectives) déjà saisies
  pour l'année à chaque régénération des fériés.
- **Module_Feries.bas / Module_Groupes.bas** — une fermeture d'entreprise
  était traitée comme un jour férié nécessitant un roulement auxiliaire :
  une semaine de fermeture se voyait attribuer un groupe G1/G2 comme un
  vrai jour férié, alors que personne ne travaille ce jour-là.
- **Module_UserForm.bas** — l'ajout d'une nouvelle fiche Personnel ou d'une
  nouvelle absence écrivait directement via `ws.Cells()` sous le dernier
  tableau au lieu de `ListObject.ListRows.Add` : la nouvelle ligne pouvait
  se retrouver hors de `Tbl_Personnel`/`Tbl_Vacances`, invisible pour la
  génération du planning. Corrigé pour `UF_EnregistrerFiche`,
  `UF_EnregistrerAbsence`, `UF_AjouterAbsenceDepuisCalendrier`.
- **Export_Serveur.bas** — le "mot de passe" de confirmation d'export était
  affiché en clair dans la boîte de dialogue elle-même : ce n'était donc
  pas un contrôle de sécurité, seulement une confirmation anti-clic. Le
  module a été renommé et redocumenté en conséquence (aucune fonctionnalité
  perdue, juste plus honnête sur ce que ça protège réellement).

## Corrections importantes

- Tirage aléatoire des ex-aequo (`ChoisirPersonne`) : `Randomize` n'est plus
  appelé à chaque poste (risque de tirages corrélés) mais une seule fois en
  tête de `GenererPlanning`.
- `EcrireAffectationsMatin` : une personne présente ne peut plus disparaître
  silencieusement du planning si l'effectif dépasse la capacité d'affichage
  (15 lignes) — un message d'avertissement est maintenant affiché.
- Gestion d'erreur étendue à toute la génération du planning (plages
  nommées `Date_planning`, `EnEpreuve_VaoX5A/B/C/D`...).
- Détection des chevauchements d'absences pour une même personne dans
  `Tbl_Vacances` (avec confirmation avant doublon).
- Validation `IsDate()` ajoutée avant conversion de date dans
  `UF_EnregistrerAbsence` et `UF_EnregistrerRemplacement` (évite un
  plantage sur saisie invalide).
- `GenererFicheExcel` : ajout d'une gestion d'erreur globale — un échec en
  cours de génération ne laisse plus Excel visuellement figé
  (`ScreenUpdating`/`DisplayAlerts` non restaurés).
- Suppressions de lignes dans `Tbl_Vacances`/`Tbl_Remplacements`/`Tbl_Feries`
  effectuées via `ListObject.ListRows.Delete` au lieu de `ws.Rows().Delete`
  (qui effaçait toute la largeur de la ligne, au-delà du tableau).
- `AjouterFermeturePeriode` : la mise à jour d'une date existante se fait
  désormais via le `ListObject` au lieu d'une écriture directe `ws.Cells()`.
- Variable de date non réinitialisée avant lecture protégée
  (`On Error Resume Next`) dans plusieurs boucles de Module_Feries et
  Module_Groupes : une cellule vide/invalide pouvait réutiliser la date de
  l'itération précédente.
- `Module_StatsHebdo` : gestion d'erreur ajoutée (plus de feuille
  `_Stats_Temp` orpheline en cas d'échec d'export PDF), chemin
  `Archives\Stats` désormais centralisé dans `Module_Constantes`.
- `Export_Serveur` : gestion d'erreur globale (la barre de statut Excel ne
  reste plus bloquée en cas d'échec), vérification des erreurs `FileCopy`
  de la documentation (n'annonce plus un export "réussi" à tort), nom
  d'archive horodaté à la seconde près (une deuxième sauvegarde le même
  jour n'écrase plus la première).
- `Commande_Bouton` : `Date_print` n'est plus enregistrée avant confirmation
  de l'impression (évite un état incohérent si l'utilisateur annule) ;
  gestion d'erreur ajoutée sur la navigation de dates.

## Nettoyage / conventions

- Centralisation dans `Module_Constantes` de toutes les colonnes
  auparavant dupliquées dans plusieurs modules avec des noms différents
  (`PERS_COL_xxx`, `VAC_COL_xxx`, `RPL_COL_xxx`, `PARAM_xxx`,
  `PLAN_COL_xxx`, lignes du planning, noms des machines Vaox1/Vaox5,
  dossiers d'archivage) : une seule source de vérité, comme demandé dans
  les conventions du projet.
- Correction de la double graphie de Vaox1 A&B ("Vaox1-A&B" vs
  "Vaox1 A&B") via deux constantes dédiées (`MACHINE_VAOX1`/`PARAM_VAOX1`).
- Suppression de code mort : ancienne version `UF_DoubleclicCalendrier`
  (v1, ~150 lignes, plus appelée depuis le formulaire), bloc commenté de
  l'ancienne version de `UF_InitCalendrierAux`, branche toujours fausse
  dans `Module_Groupes.GenererGroupesAnnee`.
- Corrections de fautes/encodage dans quelques commentaires et messages
  utilisateur (accents, caractères mal réencodés).

## Non traité dans cette passe (recommandations pour la suite)

- `Module_UserForm.bas` (3200 lignes) reste le plus gros module du projet
  et concentre toute la logique métier des 5 onglets ; un découpage par
  onglet faciliterait la maintenance future si le fichier continue de
  grossir.
- De nombreuses colonnes de `Tbl_Vacances`/`Tbl_Remplacements` restent
  référencées par index littéral (`2`, `3`, `5`...) dans les portions de
  `Module_UserForm.bas` qui n'ont pas été touchées par cette revue ; les
  constantes correspondantes existent maintenant dans `Module_Constantes`
  et peuvent être substituées progressivement.
- Le chemin serveur dans `Export_Serveur.bas` contient encore le segment
  `"Test Planning 2.0.0"` — à vérifier/mettre à jour avant un passage en
  production réelle.
- Aucun test automatisé (normal en VBA) : une checklist de vérification
  manuelle après import est recommandée avant de remettre le fichier en
  production.
