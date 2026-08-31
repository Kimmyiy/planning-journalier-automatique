Option Explicit

'==============================================================================
' MODULE : Module_Types
' Projet : Planning Journalier Automatique
' Version : 2.0.0
' Auteurs : JNF
'           Avec la collaboration de Claude (Anthropic)
' Date    : 2026
'
' Description :
'   Déclarations des types publics partagés entre tous les modules.
'   Ce module doit être chargé avant tout autre module qui utilise ces types.
'==============================================================================

'------------------------------------------------------------------------------
' TYPE : JourCalendrier
' Représente un jour dans la grille du calendrier mensuel (42 cases).
' Utilisé par Module_Feries.GenererMoisCalendrier().
'------------------------------------------------------------------------------
Public Type JourCalendrier
    dateJour  As Date     ' Date du jour
    estDuMois As Boolean  ' True si le jour appartient au mois affiché
    estWE     As Boolean  ' True si samedi ou dimanche
    estFerie  As Boolean  ' True si jour férié ou fermeture
    nomFerie  As String   ' Nom du jour férié (vide si non férié)
    groupe    As String   ' Groupe auxiliaire actif (G1 / G2 / vide)
    couleur   As Long     ' Couleur de fond pour l'affichage
End Type