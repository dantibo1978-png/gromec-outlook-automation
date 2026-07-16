# HANDOFF

## Objectif
Automatiser le traitement des confirmations de commande fournisseurs chez Gromec :
classification des courriels Outlook, extraction PDF via Claude, comparaison
prix/quantités vs commande envoyée, sync SAP via DTW.

## État actuel
- Fait : 6 correctifs de stabilité (crash-loop SyncDTW, PUT→PATCH Firebase,
  retry écritures, timeout DTW.exe, fuite COM Outlook, verrou d'instance).
- Fait : mode `-Audit` (trace locale, jamais Firebase) sur VerifierConfirmation.ps1.
- Fait : classifieur corrigé — EstConfirmation dépend uniquement de QTYPE.
- Fait : ACTION_REQUISE contourne le garde-fou anti-doublon (BC déjà OK).
- Fait : report des lignes déjà confirmées OK entre plusieurs courriels d'un
  même BC (ex. Masco : confirmation puis rapport d'écarts partiel).
- Fait : fallback CORPS→PDF si extraction CORPS vide et PJ PDF présente.
- Fait : rejet des PDF renvoyés identiques (hash SHA256) au PO envoyé.
- Fait : migration Firebase "actionRequise" rendue vraiment unique (gatée) —
  cause probable du dépassement de quota (15,79 Go/15j sur 10 Go/mois).
- Reste à faire : surveiller Masco en production réelle sur quelques jours.
- Abandonné : cas DM Valve ambigu, écarts multi-BC dans un seul courriel.
- Bloqué : SAP UI API pas installé — bloque l'automatisation Mailer SAP.

## Fichiers touchés
- `VerifierConfirmation.ps1` — classification, recherche courriel envoyé,
  extraction PDF, comparaison, écriture Firebase (le plus modifié).
- `SyncDTW.ps1` — boucle 30s Firebase → SAP DTW.
- `index.html` — dashboard, ajout message `PDF_IDENTIQUE_AU_PO_ENVOYE`.
- `firebase-rules.json` — règles Firebase de référence (à coller manuellement).

## Décisions prises
- Fallback CORPS→PDF seulement si CORPS vide ET PJ PDF existe (zéro impact
  sur le cas CORPS qui fonctionne, majorité des fournisseurs).
- Détection PDF-identique par hash, pas nom/taille — un PDF annoté même
  légèrement reste traité normalement.
- Pas de rotation de projet Firebase pour le quota : corriger la cause
  (migration retéléchargée) plutôt qu'une rotation manuelle fragile.

## Prochaine étape
Après reset du quota Firebase (1er août), vérifier dans Firebase Console →
Usage que la consommation redescend sous 10 Go/mois. Si Masco reste stable
en production, considérer le sujet clos.

## Pièges
- `Update-ScriptSiNecessaire` doit propager TOUS les paramètres à la relance
  (`$argList`) — `-Audit` et `-Nombre` oubliés une fois chacun.
- `Write-FirebaseHistorique` : PATCH jamais PUT sur une entrée existante
  (PUT efface `syncSAP`/`confirme` gérés par SyncDTW.ps1).
- Champ `CONFIRMATION: OUI/NON` du classifieur = reliquat legacy, ne pas s'y
  fier — seul `QTYPE` décide.
- Un bloc "migration unique" sans drapeau persistant se réexécute à CHAQUE
  démarrage, pas juste une fois.
