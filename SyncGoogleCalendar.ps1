$CHEMIN_VERS_indexjs = ""
$CHEMIN_VERS_DOSSIER_EXPORTS_XLS = ""

$cmd = "node '$CHEMIN_VERS_indexjs' '$CHEMIN_VERS_DOSSIER_EXPORTS_XLS'"
Write-Host $cmd
Write-Host ""
Invoke-Expression $cmd