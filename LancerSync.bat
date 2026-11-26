@echo off
title GromecSync - NE PAS FERMER
:loop
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "U:\GromecOutlook\SyncDTW.ps1"
echo %date% %time% - SyncDTW arrete, redemarrage dans 10 secondes...
timeout /t 10 /nobreak >nul
goto loop
