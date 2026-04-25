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
# $PSScriptRoot peut etre vide dans un exe compile par PS2EXE : fallback sur le chemin du processus
$AppRoot = if ($PSScriptRoot) {
    $PSScriptRoot
} else {
    Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}

# Cacher la fenetre console immediatement (fonctionne en .ps1 et en exe compile)
# La session reste interactive ce qui est requis par Connect-ManagementServer -ShowDialog
Add-Type -Name ConsoleHider -Namespace '' -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr h, int n);
'@ -ErrorAction SilentlyContinue
try { [ConsoleHider]::ShowWindow([ConsoleHider]::GetConsoleWindow(), 0) | Out-Null } catch {}

# Verifier qu'on est sur Windows (Milestone est Windows only)
if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    Write-Error "Milestone Toolkit necessite Windows."
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}

# Charger et afficher la fenetre de verification des dependances
try {
    . (Join-Path $AppRoot 'src/Core/Show-StartupCheck.ps1')
    $shouldContinue = Show-StartupCheck -AppRoot $AppRoot
}
catch {
    Write-Error "Erreur lors du demarrage : $_"
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}

if (-not $shouldContinue) {
    exit 0
}

# Lancer l'application principale
try {
    & (Join-Path $AppRoot 'src/App.ps1') -RootPath $AppRoot
}
catch {
    Write-Error "Erreur fatale : $_"
    Write-Error $_.ScriptStackTrace
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}
