# Planning Journalier Automatique — v2.0.3

Quatre demandes fonctionnelles + une correction de bug.

## 1. Ajouter tout un groupe d'auxiliaires en renfort pour une semaine

Nouveau mini-formulaire **UserForm_GroupeSemaine**, ouvert depuis un nouveau
bouton **"Ajouter un groupe pour la semaine…"** dans l'onglet Auxiliaires.
On choisit le groupe (G1/G2) et le lundi de la semaine concernée : chaque
auxiliaire actif du groupe est enregistré en renfort (Tbl_Remplacements,
Type="Renfort") du lundi au vendredi. Les doublons (personne déjà
enregistrée ce jour-là) sont détectés et non recréés.

**Changement de fond nécessaire** : jusqu'ici, un renfort enregistré pour un
jour de semaine n'était compté que dans les statistiques d'alerte "effectif
insuffisant" — il n'apparaissait **jamais réellement dans le planning
généré** (`ChargerPersonnes` ne chargeait que le personnel Fixe en semaine).
C'est corrigé : un auxiliaire en renfort pour une date donnée est maintenant
inclus dans la génération du planning de ce jour, journée complète par
défaut.

## 2. Date bloquée sur "Ajouter un remplacement" depuis la fiche

Dans `UserForm_Remplacement` (le mini-formulaire ouvert en double-cliquant un
jour sur le calendrier personnel d'une auxiliaire), le champ date était
`Enabled = False` : il affichait le jour cliqué sans jamais permettre de le
changer. Il est maintenant modifiable (avec validation de format), pour
corriger une date sans devoir fermer et recommencer sur le bon jour.

En chemin, une incohérence a été corrigée : ce même formulaire écrivait
"Weekend"/"Semaine" dans la colonne Type de Tbl_Remplacements, un vocabulaire
différent de "Remplacement"/"Renfort" utilisé partout ailleurs (et qui
faussait le comptage des renforts). Il écrit maintenant "Remplacement",
cohérent avec le reste.

## 3. Absences pour les auxiliaires

Jusqu'ici, un auxiliaire absent devait obligatoirement être remplacé
(onglet Auxiliaires). L'onglet **Absences** liste maintenant tout le
personnel actif, Fixe et Auxiliaire — un auxiliaire peut être déclaré
simplement absent, sans qu'on lui cherche de remplaçante.

Pour que ça compte réellement dans le planning : `EstDisponibleAuxiliaire`
(Planning_Auto.bas) vérifie maintenant les absences enregistrées et exclut
l'auxiliaire du planning du weekend/férié concerné, sauf si un remplacement
explicite existe déjà pour cette date (qui reste prioritaire).

## 4. Absences du personnel fixe par demi-journée

Nouvelle colonne **Période** dans Tbl_Vacances (Journée complète / Matin /
Après-midi), déjà ajoutée dans le classeur avec la valeur "Journée" pour les
28 absences existantes (comportement inchangé pour l'historique). Nouveau
champ **cboVacPeriode** dans l'onglet Absences et **cboAbsPeriode** dans le
mini-formulaire du calendrier personnel.

Le calcul de présence (`ChargerPersonnes`) tient maintenant compte de la
période : une absence "Matin" sur une personne à l'horaire "Journée
complète" la rend disponible l'après-midi seulement (et inversement) ; une
absence qui couvre le seul moment où la personne travaille l'exclut
complètement, comme avant.

## Bug signalé : message à la fermeture de Gestion du personnel

Vous avez mentionné un message indiquant qu'un "hook" empêche la
sauvegarde à la fermeture du formulaire. Aucune trace de terme "hook" n'a
été trouvée dans le code VBA du projet (recherché dans tous les modules),
et `UserForm_QueryClose` se contente d'un `ThisWorkbook.Save` simple, sans
gestionnaire `Workbook_BeforeSave`/`BeforeClose` particulier. Impossible de
corriger sans le texte exact du message — merci de le copier-coller (ou une
capture) la prochaine fois qu'il apparaît, pour que je puisse localiser la
cause réelle plutôt que deviner.

## Contrôles à créer manuellement dans l'éditeur VBA (Alt+F11, mode création)

Le code ne peut pas ajouter de nouveaux contrôles visuels sur un formulaire
existant (seul le code peut être livré en texte) — à ajouter vous-même
avant d'importer les fichiers :

- **UserForm_GestionPersonnel**, onglet Auxiliaires :
  `btnAjouterGroupeSemaine` (CommandButton, caption "Ajouter un groupe pour
  la semaine…")
- **UserForm_GestionPersonnel**, onglet Absences : `cboVacPeriode`
  (ComboBox, à côté de `cboVacType`)
- **UserForm_Absence** (mini-formulaire) : `cboAbsPeriode` (ComboBox, à côté
  de `cboAbsType`)
- **Nouveau UserForm** `UserForm_GroupeSemaine` (à créer entièrement) :
  `lblGroupeSemaineInfo` (Label), `cboGroupeSemaine` (ComboBox),
  `txtDebutSemaine` (TextBox), `btnGroupeSemaineConfirmer` et
  `btnGroupeSemaineAnnuler` (CommandButton) — détail complet en en-tête du
  fichier `UserForm_GroupeSemaine.frm`.

## Checklist de vérification après import

- [ ] Générer un planning un jour de semaine normal → rien ne doit changer.
- [ ] Enregistrer un groupe entier en renfort pour une semaine, puis générer
      le planning d'un des jours concernés → les auxiliaires du groupe
      doivent apparaître dans les affectations.
- [ ] Ouvrir la fiche d'une auxiliaire, double-cliquer un jour, vérifier que
      la date proposée dans "Ajouter un remplacement" est modifiable.
- [ ] Déclarer une absence pour une auxiliaire depuis l'onglet Absences,
      vérifier qu'elle n'apparaît plus dans le planning du weekend concerné
      (sans avoir enregistré de remplacement).
- [ ] Déclarer une absence "Matin" pour une personne fixe à l'horaire
      "Journée complète", générer le planning de ce jour → elle ne doit pas
      apparaître dans les postes machines du matin, mais doit apparaître
      dans la liste de l'après-midi.
