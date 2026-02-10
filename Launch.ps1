<#
.SYNOPSIS
    Lance l'application Milestone Toolkit.
.DESCRIPTION
    Point d'entree utilisateur. Double-cliquer ou executer depuis PowerShell.
    Verifie la version de PowerShell et lance le bootstrap applicatif.
.NOTES
    Prerequis : Windows PowerShell 5.1+ ou PowerShell 7+ sur Windows.
    Les modules MilestonePSTools et ImportExcel seront installes automatiquement si absents.
#>

#Requires -Version 5.1

# Determiner le repertoire racine du projet
$AppRoot = $PSScriptRoot

# Verifier qu'on est sur Windows (Milestone est Windows only)
if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    Write-Error "Milestone Toolkit necessite Windows."
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}

# Lancer l'application
try {
    & (Join-Path $AppRoot 'src/App.ps1') -RootPath $AppRoot
}
catch {
    Write-Error "Erreur fatale : $_"
    Write-Error $_.ScriptStackTrace
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}
