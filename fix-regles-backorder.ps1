# Script one-shot pour corriger la regle de quantite backorder dans Firebase
# Executer une seule fois puis supprimer

$config = Get-Content "U:\GromecOutlook\config.json" | ConvertFrom-Json
$FirebaseUrl = $config.firebaseUrl

$nouvellesRegles = "- qty = quantite COMMANDEE (colonne Comm./Ord., Qty Ordered, Quantity, etc.) -- ne PAS additionner la colonne AV/BO, Back Order, ou Balance; le back-order represente la portion non disponible de la meme quantite, pas une quantite supplementaire`n- netUnit = prix UNITAIRE net final`n- code = code article du FOURNISSEUR (pas le code SAP Gromec)`n- Si devise USD detectee: ecrire USD, sinon CAD"

$jsonBody = $nouvellesRegles | ConvertTo-Json -Compress
Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/regles_generales.json" -Method Put -Body $jsonBody -ContentType "application/json; charset=utf-8" -TimeoutSec 15
Write-Host "Regles generales mises a jour dans Firebase."
