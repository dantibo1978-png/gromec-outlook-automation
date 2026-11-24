# MonitorSync.ps1
# Surveille SyncDTW.ps1 et le relance automatiquement s'il s'arrete
# Lancer: Start-Process powershell -ArgumentList "-File U:\GromecOutlook\MonitorSync.ps1" -WindowStyle Hidden
# Ou en tache planifiee Windows (au demarrage)

$syncScript = "U:\GromecOutlook\SyncDTW.ps1"
$logFile = "U:\GromecOutlook\MonitorSync.log"
$checkInterval = 60  # secondes

function Write-MonitorLog($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Write-Host $line
    try { Add-Content -Path $logFile -Value $line -ErrorAction SilentlyContinue } catch {}
}

function Is-SyncRunning {
    $procs = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" 2>$null
    foreach ($p in $procs) {
        if ($p.CommandLine -and $p.CommandLine -like "*SyncDTW*" -and $p.CommandLine -notlike "*MonitorSync*") {
            return $true
        }
    }
    return $false
}

function Start-Sync {
    Write-MonitorLog "RELANCE SyncDTW.ps1"
    Start-Process powershell -ArgumentList "-File `"$syncScript`"" -WindowStyle Hidden
    Start-Sleep -Seconds 5
    if (Is-SyncRunning) {
        Write-MonitorLog "OK SyncDTW demarre avec succes"
    } else {
        Write-MonitorLog "ERREUR SyncDTW n'a pas demarre"
    }
}

Write-MonitorLog "MonitorSync demarre (intervalle: ${checkInterval}s)"

if (-not (Is-SyncRunning)) {
    Write-MonitorLog "SyncDTW pas en cours - lancement initial"
    Start-Sync
}

while ($true) {
    Start-Sleep -Seconds $checkInterval
    if (-not (Is-SyncRunning)) {
        Write-MonitorLog "SyncDTW arrete detecte!"
        Start-Sync
    }
}
