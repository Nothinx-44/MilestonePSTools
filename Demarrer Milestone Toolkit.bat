@echo off

:: Si le dossier Dependencies/ n'est pas present mais qu'une archive Dependencies.zip existe,
:: la decompresser automatiquement une seule fois avant de lancer l'application.
if not exist "%~dp0Dependencies\" (
    if exist "%~dp0Dependencies.zip" (
        echo Extraction de Dependencies.zip...
        powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; $base = '%~dp0'; $zip = Join-Path $base 'Dependencies.zip'; $dest = Join-Path $base 'Dependencies'; if (-not (Test-Path $dest) -and (Test-Path $zip)) { [System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $dest) }"
    )
)

:: Si des archives .nupkg sont presentes dans Dependencies/, extraire chaque module manquant.
if exist "%~dp0Dependencies\" (
    echo Verification des archives Dependencies\*.nupkg...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; $base = '%~dp0'; $deps = Join-Path $base 'Dependencies'; Get-ChildItem -Path $deps -Filter '*.nupkg' -File | ForEach-Object { $archive = $_.FullName; $name = $_.BaseName; $dest = Join-Path $deps $name; if (-not (Test-Path $dest)) { Write-Host 'Extraction de' $_.Name 'dans' $dest; [System.IO.Compression.ZipFile]::ExtractToDirectory($archive, $dest) } }"
)

:: Deblocage de TOUS les fichiers du projet (telechargement ZIP depuis GitHub)
:: Sans ca, Windows bloque les scripts et les modules PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Get-ChildItem -Path '%~dp0' -Recurse -File | Unblock-File -ErrorAction SilentlyContinue"

:: Lancement
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Bootstrap.ps1"

if %errorlevel% neq 0 ( echo. && echo ERREUR code %errorlevel% && pause )
