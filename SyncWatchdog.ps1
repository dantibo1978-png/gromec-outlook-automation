<#
.SYNOPSIS
    Watchdog pour SyncDTW - Verifie que le sync tourne et le relance si necessaire.
.DESCRIPTION
    Lance par le Planificateur de taches Windows toutes les 5 minutes.
    Si SyncDTW.ps1 ne tourne pas, le relance automatiquement.
    Pas besoin de droits admin.
.NOTES
    Pour installer la tache planifiee, executer dans PowerShell:

    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-ExecutionPolicy Bypass -WindowStyle Hidden -File "U:\GromecOutlook\SyncWatchdog.ps1"'
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days 365)
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
    Register-ScheduledTask -TaskName "GromecSyncWatchdog" -Action $action -Trigger $trigger -Settings $settings -Description "Surveille et relance SyncDTW si necessaire"
#>

$scriptDir = "U:\GromecOutlook"
$syncScript = Join-Path $scriptDir "SyncDTW.ps1"
$logFile = Join-Path $scriptDir "SyncWatchdog.log"

function Write-WatchdogLog {
    param([string]$msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts  $msg"
    try {
        Add-Content -Path $logFile -Value $line -ErrorAction SilentlyContinue
        # Garder le log petit (max 500 lignes)
        if (Test-Path $logFile) {
            $lines = Get-Content $logFile -ErrorAction SilentlyContinue
            if ($lines.Count -gt 500) {
                $lines[-200..-1] | Set-Content $logFile -ErrorAction SilentlyContinue
            }
        }
    } catch {}
}

# Verifier si SyncDTW tourne
$syncRunning = $false
try {
    $procs = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
        if ($p.CommandLine -and $p.CommandLine -like "*SyncDTW*") {
            $syncRunning = $true
            break
        }
    }
} catch {}

if ($syncRunning) {
    # Tout va bien, ne rien faire
    exit 0
}

# SyncDTW ne tourne pas - le relancer
Write-WatchdogLog "SyncDTW ne tourne pas - relance en cours..."

try {
    Start-Process powershell -ArgumentList @(
        "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", $syncScript
    ) -WindowStyle Hidden
    Start-Sleep -Seconds 5

    # Verifier que ca a demarre
    $ok = $false
    $procs = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
        if ($p.CommandLine -and $p.CommandLine -like "*SyncDTW*") {
            $ok = $true
            break
        }
    }
    if ($ok) {
        Write-WatchdogLog "SyncDTW relance avec succes."
    } else {
        Write-WatchdogLog "ERREUR: SyncDTW n'a pas demarre apres relance."
    }
} catch {
    Write-WatchdogLog "ERREUR relance: $($_.Exception.Message)"
}
