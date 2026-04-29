@echo off

:: Deblocage de TOUS les fichiers du projet (telechargement ZIP depuis GitHub)
:: Sans ca, Windows bloque les scripts et les modules PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Get-ChildItem -Path '%~dp0' -Recurse -File | Unblock-File -ErrorAction SilentlyContinue"

:: Lancement
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Bootstrap.ps1"

if %errorlevel% neq 0 ( echo. && echo ERREUR code %errorlevel% && pause )
