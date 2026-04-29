@echo off
chcp 65001 > nul

:: Deblocage des scripts (telechargement ZIP depuis GitHub)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Get-ChildItem -Path '%~dp0' -Recurse -Include '*.ps1','*.psm1','*.psd1' | Unblock-File -ErrorAction SilentlyContinue"

:: Lancement
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Bootstrap.ps1"

if %errorlevel% neq 0 ( echo. && echo ERREUR code %errorlevel% && pause )
