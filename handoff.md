# Gromec Outlook Automation — Handoff

## Objectif

Automatiser le traitement des confirmations de commande fournisseurs chez Gromec Inc. (distributeur industriel, Québec). Le système surveille la boîte Outlook, classe chaque courriel reçu d'un fournisseur, et compare les prix/quantités confirmés avec la commande originale envoyée.

**Flux principal :**
1. Un courriel fournisseur arrive dans Outlook
2. VBA Outlook déclenche `VerifierConfirmation.ps1` avec l'EntryID du courriel
3. Claude Haiku classifie le courriel (confirmation? avis d'expédition? facturation? …)
4. Si c'est une confirmation : le script retrouve le courriel de commande envoyé, extrait les items des deux PDF (commande vs confirmation), compare prix/quantités
5. Le résultat est écrit dans Firebase Realtime Database
6. Le dashboard web (`index.html`) affiche les résultats, permet de copier les prix/quantités vers SAP via DTW
7. `SyncDTW.ps1` tourne en arrière-plan, poll Firebase, génère les fichiers DTW et les importe dans SAP

## Architecture / fichiers

| Fichier | Lignes | Rôle |
|---|---|---|
| `VerifierConfirmation.ps1` | ~2600 | Script principal : classification (Claude Haiku), recherche du courriel envoyé, extraction PDF (Claude Sonnet), comparaison prix/quantités, écriture Firebase |
| `SyncDTW.ps1` | ~1100 | Boucle 30s : poll Firebase pour les actions utilisateur (copier prix, copier quantités, reclassification, relances), génère fichiers DTW, lance l'import SAP |
| `index.html` | ~2400 | Dashboard web (Firebase Realtime) : affiche l'historique, boutons d'action, soumission manuelle de BC |
| `DTW/*.xml` | — | Templates DTW pour l'import SAP (UpdatePO, UpdatePriceList, UpdateLeadTime) |
| `outils/*.html` | — | Outils web auxiliaires (analyseur SAP, recherche commandes, délais produits, relance fournisseur) |

## Déploiement

- Les scripts se mettent à jour automatiquement depuis `main` de ce repo GitHub au lancement
- **Toute modification doit être mergée dans `main`** pour être déployée
- Branche de développement : `claude/project-context-continuity-kyr9xh`
- Firebase : Realtime Database (URL dans `config.json` sur le poste Windows, pas dans le repo)

## Décisions prises

### Classifieur (refonte récente)

**Ancien modèle** : 8 questions booléennes (Q1-Q8) avec règle `(Q1|Q5) ET (Q4|Q2|Q7) ET !Q6 ET !Q8`. Problème : les questions se contredisaient (un avis d'expédition mentionne des quantités → Q2=OUI, donc classé comme confirmation).

**Nouveau modèle** (actuel) : 3 questions factuelles + 1 catégorisation mutuellement exclusive.
```
Q1_NUMERO_BC: OUI/NON        — numéro de commande Gromec (9XXXXXX) présent?
Q3_DATE_LIVRAISON: OUI/NON   — date de livraison confirmée?
Q5_DOCUMENT_COMMANDE: OUI/NON — PJ contient un vrai document de confirmation fournisseur?
QTYPE: CONFIRMATION / ACTION_REQUISE / AVIS_EXPEDITION / SUIVI_STATUT / FACTURATION / AUTRE
```

Règle : c'est une confirmation si `QTYPE ∈ {CONFIRMATION, ACTION_REQUISE}`.  
`ACTION_REQUISE` → `poReviseRequis = true` (badge "Confirmation requise" dans le dashboard).

### Recherche du courriel envoyé (`Find-CourrielEnvoyeCorrespondant`)

3 stratégies en cascade :
1. **Numéro BC dans Subject ou Body** — itération manuelle des items envoyés (pas de Restrict, évite le problème Jet/DASL)
2. **Multi-BC fallback** — si le premier numéro BC extrait ne trouve rien, essaie tous les 90xxxxx du courriel via `Get-TousNumerosBC`
3. **Domaine fournisseur** — si aucun BC ne matche, cherche par domaine de l'expéditeur (avec support des alias via Firebase `domaines_alias`)

### Alias de domaines

Chargés au démarrage depuis Firebase (`gromec_vba/parametres/domaines_alias`). Mapping bidirectionnel : envoyé à `@nimatec.com`, répondu par `@boshart.com` → les deux sont reconnus comme le même fournisseur.

### Optimisation Firebase

Le poll toutes les 30s téléchargeait ~1 MB (tout l'historique) = ~2.8 GB/jour. Corrigé avec :
- Champ `actionRequise` sur chaque entrée (true = action en attente, false = terminé)
- `SyncDTW.ps1` poll uniquement `orderBy="actionRequise"&equalTo=true` (quelques KB)
- Le téléchargement complet ne se fait plus qu'une fois par heure (pour le nettoyage)
- Index Firebase requis : `.indexOn: ["entryID", "numeroCommande", "actionRequise"]`

### Reclassification manuelle

Pipeline VBA → Firebase `reclassifications/` → SyncDTW → `VerifierConfirmation.ps1 -Force`.  
Le flag `-Force` (paramètre `ForcerTraitement`) :
- Ignore le verdict du classifieur (traite toujours comme confirmation)
- Bypass `Test-BCDejaTraitee` (permet de retraiter un BC déjà vu)
- Bypass `Test-ConversationTraitee` et la vérification des catégories Outlook

### Extraction PDF (prompt Claude Sonnet)

- Format attendu : `ARTICLE|ligne|code|description|qté|unité|prix_unitaire|prix_total`
- Instructions conditionnelles pour les cas spéciaux (numéros de ligne non séquentiels, double colonne de prix Masco)
- Si le document a deux colonnes de prix (ex: "Masco Unit Price" et "Customer Sent PO Unit Price"), utiliser le prix du fournisseur

## Cas connus résolus

| Cas | Problème | Solution |
|---|---|---|
| Boshart/Nimatec | Envoyé à @nimatec.com, répondu par @boshart.com → NON TROUVÉ | Alias de domaines dans Firebase |
| Boshart "REPLY NEEDED TO RELEASE ORDER" | Classé comme non-confirmation | QTYPE=ACTION_REQUISE |
| Trueline 9007257 | NON TROUVÉ malgré BC dans les deux sujets | Fix du filtre Restrict (Jet/DASL mixé), puis remplacement par itération manuelle |
| Masco 9007237 | NON TROUVÉ, premier BC extrait incorrect | Multi-BC fallback (`Get-TousNumerosBC`) |
| Masco (2e email même BC) | Bloqué par Test-BCDejaTraitee | `ForcerTraitement` propagé dans la chaîne |
| Intermetalink 9003331 | Discussion crédit classée comme confirmation | QTYPE=FACTURATION |
| Westlake Pipe | Question logistique classée comme confirmation | QTYPE=SUIVI_STATUT |
| CCTF nipples | Avis d'expédition partiel classé comme confirmation | QTYPE=AVIS_EXPEDITION |

## Problèmes connus non résolus

- **Masco** : l'extraction PDF et la recherche de courriel envoyé fonctionnent mal avec leurs Order Acknowledgements. Laissé de côté pour l'instant.
- **Firebase Rules** : l'utilisateur doit publier les règles avec `.read: true`, `.write: true` et les indexes.

## Prochaines étapes

1. **SAP UI API** — Installer le composant UI API via le setup SAP B1 (requis pour l'automatisation Mailer)
2. **SAP Mailer automation** — Une fois l'UI API installée, automatiser l'envoi des commandes par email depuis SAP
3. **Surveiller le classifieur** — Le nouveau modèle QTYPE est en production. Valider sur quelques jours qu'il n'y a pas de régressions.
4. **Masco** — Revisiter quand les autres problèmes seront stabilisés
