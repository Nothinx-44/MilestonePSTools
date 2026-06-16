<#
.SYNOPSIS
    Applique une mise a jour de Milestone Toolkit apres que le processus principal a quitte.
    Appele par le bouton de mise a jour dans Show-StartupCheck — ne pas executer directement.
.PARAMETER Target
    Dossier racine de l'installation a mettre a jour.
.PARAMETER Source
    Dossier contenant les nouveaux fichiers extraits depuis l'archive GitHub.
.PARAMETER Launcher
    Chemin du fichier .bat a relancer apres la mise a jour.
#>
param(
    [Parameter(Mandatory)] [string]$Target,
    [Parameter(Mandatory)] [string]$Source,
    [Parameter(Mandatory)] [string]$Launcher
)

# Attend que le processus parent libere les fichiers (max 5 s)
Start-Sleep -Seconds 2
for ($i = 0; $i -lt 20; $i++) {
    try   { Get-ChildItem -Path $Target -ErrorAction Stop | Out-Null; break }
    catch { Start-Sleep -Milliseconds 250 }
}

try {
    Copy-Item -Path (Join-Path $Source '*') -Destination $Target -Recurse -Force -ErrorAction Stop
} catch {
    $errLog = Join-Path $env:TEMP 'MilestoneToolkitUpdate_error.txt'
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Echec de la copie des fichiers de mise a jour : $_" |
        Out-File -FilePath $errLog -Encoding UTF8 -Append
}

if (Test-Path $Launcher) {
    Start-Process -FilePath $Launcher
}
