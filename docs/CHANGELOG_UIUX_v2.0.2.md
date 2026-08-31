# Planning Journalier Automatique — v2.0.2 (audit et corrections UI/UX)

Audit d'ergonomie mené par 3 agents en parallèle (feuille "Planning journalier"
+ boutons, UserForms de gestion, feuilles de données secondaires), puis
corrections appliquées directement dans le classeur `.xlsm` (édition XML
chirurgicale, sans passer par un ré-enregistrement complet du fichier —
VBA, boutons, graphique restent strictement identiques byte à byte).

## Corrigé dans le classeur (workbook/2.0.0_Planning.xlsm)

- **Tbl_Vacances tronqué** : le tableau structuré ne couvrait que 2 lignes
  (A1:E2) alors que 28 lignes d'absences réelles existaient en dessous, hors
  tableau — invisibles pour la détection de chevauchement et l'affichage
  dans le UserForm. Étendu à A1:E30.
- **Dates au format américain** dans Vacances (colonnes Date_Debut/Date_Fin) :
  remises au format jj.mm.aaaa, cohérent avec les autres feuilles.
- **Couleur par machine jamais affichée** : la mise en forme conditionnelle
  cherchait le nom de la machine dans la colonne Nom au lieu de la colonne
  Poste — elle ne s'est donc jamais déclenchée depuis la création du fichier.
  Corrigée et étendue au bloc après-midi (qui n'en avait pas du tout).
- **Texte invisible (blanc sur blanc)** sur les cellules "MàJ statistique le"
  et "lundi du 1er WE du G1" (F32:G33) — rendu lisible.
- **Listes déroulantes de saisie cassées** (#REF! partout, plus un résidu
  pointant vers un classeur externe fermé et éparpillé sur des milliers de
  cellules parasites) : reconstruites proprement pour la colonne Poste et le
  Nom de l'après-midi, avec message d'aide et message d'erreur explicite.
- **Planning_ApresMidi trop court** (F23:G26 vs 4 personnes possibles dans le
  code) : étendu à F23:G27.
- Libellés incohérents : "Post" → "Poste" (en-tête après-midi), "Jour
  précédant" → "Jour précédent" (faute).
- **29 plages nommées cassées ou orphelines** (`#REF!`, ou pointant vers un
  classeur externe fermé) supprimées du Gestionnaire de noms.

## Corrigé dans le code VBA (à réimporter, cf. session précédente)

- `Module_Constantes.bas` : `NOM_FEUILLE_LISTES`/`NOM_FEUILLE_INDEX` ne
  correspondaient pas aux noms réels des onglets ("Liste" vs "Listes",
  "Index des Versions" vs "index des versions"). Corrigé (constantes non
  utilisées actuellement, donc sans risque).

## Recommandé mais non appliqué automatiquement (à faire dans Excel, sans risque)

Ce sont des réglages purement visuels, plus sûrs à faire directement dans
Excel (quelques clics) qu'en édition XML :
- Couleur d'onglet : orange pour "Planning journalier" (feuille quotidienne),
  gris/bleu clair pour les feuilles de référence.
- Figer la ligne d'en-tête sur Vacances, Feries, Groupes, Remplacements
  (Affichage > Figer les volets > Figer la ligne supérieure).
- Mise en forme conditionnelle G1/G2 sur la colonne Groupe de la feuille
  Groupes (mêmes couleurs que le calendrier : bleu clair G1, vert clair G2).
- Élargir les colonnes Nom_Absent/Nom_Remplaçant de la feuille Remplacements
  (actuellement en largeur par défaut, texte tronqué).
- Séparateur de milliers sur les grandes quantités (colonnes B:D du bloc
  "Quantités", ex. 11682 → 11'682).
- Distinguer visuellement (fond gris clair) les colonnes de statistiques
  auto-générées de Personnel (R:W) des colonnes de saisie manuelle.
- Ajouter un bouton "Exporter vers serveur" sur la feuille — la fonction VBA
  existe (`Export_Serveur.bas`) mais n'est reliée à aucun bouton visible.

## Constaté, à trancher avec le métier (pas une correction technique)

- La colonne "Poste" de l'après-midi n'a aucun effet : le code n'y écrit
  jamais et ne compte pas ces affectations dans les statistiques. À décider :
  la supprimer (garder juste "Nom") ou compléter le code pour qu'elle serve.
- Deux boutons de suppression de remplacement identiques dans l'onglet
  Auxiliaires du UserForm (`btnSupprimerRpl` et `btnSupprimerRplIndividuel`)
  appellent la même action — à vérifier dans l'éditeur VBA/le formulaire
  lequel est réellement affiché, et retirer l'autre.
- Onglet Auxiliaires : mélange saisie quotidienne (remplacement/renfort) et
  configuration annuelle rare (référence G1, génération des groupes) — à
  séparer visuellement ou déplacer vers l'onglet Paramètres.
- Plusieurs confirmations modales (MsgBox "OK") pour des actions fréquentes
  (absence, remplacement) — pourraient devenir un simple message non
  bloquant pour fluidifier la saisie répétée.
