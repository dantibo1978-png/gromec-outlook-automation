# MonitorSync.ps1
# Surveille SyncDTW.ps1 et affiche une alerte clignotante si arrete
# Lancer: Start-Process powershell -ArgumentList "-File U:\GromecOutlook\MonitorSync.ps1" -WindowStyle Hidden

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$syncScript = "U:\GromecOutlook\SyncDTW.ps1"
$checkInterval = 30  # secondes

function Is-SyncRunning {
    $procs = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" 2>$null
    foreach ($p in $procs) {
        if ($p.CommandLine -and $p.CommandLine -like "*SyncDTW*") {
            return $true
        }
    }
    return $false
}

function Show-Alert {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SYNC DTW ARRETE!"
    $form.Size = New-Object System.Drawing.Size(420, 200)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::DarkRed

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "SyncDTW est arrete!"
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = [System.Drawing.Color]::White
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(400, 50)
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.TextAlign = "MiddleCenter"
    $form.Controls.Add($label)

    $btnRestart = New-Object System.Windows.Forms.Button
    $btnRestart.Text = "Relancer SyncDTW"
    $btnRestart.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $btnRestart.Size = New-Object System.Drawing.Size(250, 50)
    $btnRestart.Location = New-Object System.Drawing.Point(80, 85)
    $btnRestart.BackColor = [System.Drawing.Color]::LimeGreen
    $btnRestart.ForeColor = [System.Drawing.Color]::Black
    $btnRestart.FlatStyle = "Flat"
    $btnRestart.Add_Click({
        Start-Process powershell -ArgumentList "-File `"$syncScript`"" -WindowStyle Hidden
        $form.Tag = "restarted"
        $form.Close()
    })
    $form.Controls.Add($btnRestart)

    $btnIgnore = New-Object System.Windows.Forms.Button
    $btnIgnore.Text = "Ignorer"
    $btnIgnore.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnIgnore.Size = New-Object System.Drawing.Size(80, 30)
    $btnIgnore.Location = New-Object System.Drawing.Point(320, 145)
    $btnIgnore.BackColor = [System.Drawing.Color]::Gray
    $btnIgnore.ForeColor = [System.Drawing.Color]::White
    $btnIgnore.FlatStyle = "Flat"
    $btnIgnore.Add_Click({ $form.Close() })
    $form.Controls.Add($btnIgnore)

    # Flash effect via timer
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $flash = @{ on = $true }
    $timer.Add_Tick({
        if ($flash.on) {
            $form.BackColor = [System.Drawing.Color]::Red
            $form.Opacity = 1.0
        } else {
            $form.BackColor = [System.Drawing.Color]::DarkRed
            $form.Opacity = 0.85
        }
        $flash.on = -not $flash.on
    })
    $timer.Start()

    # Flash taskbar
    $form.WindowState = "Minimized"
    $form.Show()
    $form.WindowState = "Normal"
    $form.Activate()

    [System.Windows.Forms.Application]::Run($form)
    $timer.Stop()
    $timer.Dispose()

    return $form.Tag
}

# Boucle principale
while ($true) {
    if (-not (Is-SyncRunning)) {
        $result = Show-Alert
        if ($result -eq "restarted") {
            Start-Sleep -Seconds 5
        }
    }
    Start-Sleep -Seconds $checkInterval
}
