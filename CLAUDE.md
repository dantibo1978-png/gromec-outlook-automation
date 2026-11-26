# Notes pour Claude

## Requetes SAP donnees a l'utilisateur
- Ne JAMAIS commencer une requete SQL a executer dans SAP (gestionnaire de requetes) par `SELECT` :
  SAP l'ajoute automatiquement. Commencer directement par la liste des colonnes, ex :
  `"ItemCode", "ItemName" FROM "OITM" WHERE ...`
- (Le SQL envoye par le code VBA via DI API / Recordset garde bien son `SELECT`.)
