<#
.SYNOPSIS
    Bootstrap interne de Milestone Toolkit. Appele par Launch.bat.
    Ne pas executer directement — utiliser Launch.bat a la racine du projet.
#>

#Requires -Version 5.1

# Version centrale — modifier ici uniquement
$script:AppVersion = '4.9.7'

# Applique TLS 1.2 des le debut du processus — requis par PowerShell Gallery.
# PowerShell 5.1 utilise TLS 1.0 par defaut, ce qui bloque Install-Module / Save-Module.
[Net.ServicePointManager]::SecurityProtocol =
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# Force Bypass au niveau du processus (complementaire au flag de la ligne de commande)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue

# $PSScriptRoot = .../src/  =>  AppRoot = parent = racine du projet
$AppRoot = if ($PSScriptRoot) {
    Split-Path -Parent $PSScriptRoot
} else {
    Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}


Add-Type -Name ConsoleHider -Namespace '' -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr h, int n);
'@ -ErrorAction SilentlyContinue
try { [ConsoleHider]::ShowWindow([ConsoleHider]::GetConsoleWindow(), 0) | Out-Null } catch {}

if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    Write-Error "Milestone Toolkit requires Windows."
    Read-Host "Press Enter to quit"
    exit 1
}

try {
    . (Join-Path $AppRoot 'src/Core/Show-LanguagePicker.ps1')
    . (Join-Path $AppRoot 'src/Core/Show-StartupCheck.ps1')

    # Lire la langue sauvegardee dans config.json (evite de montrer le selecteur a chaque demarrage)
    $_configPath = Join-Path $AppRoot 'config.json'
    $script:Lang = $null
    if (Test-Path $_configPath) {
        try {
            $_cfg = Get-Content $_configPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($_cfg.language -in @('fr', 'en')) { $script:Lang = $_cfg.language }
        } catch {}
    }

    if (-not $script:Lang) {
        $script:Lang = Show-LanguagePicker
        # Sauvegarder la preference pour les prochains demarrages
        try {
            if (Test-Path $_configPath) {
                $_cfg = Get-Content $_configPath -Raw -Encoding UTF8 | ConvertFrom-Json
            } else {
                $_cfg = [PSCustomObject]@{ outputDirectory='./Output'; snapshotQuality=95; csvDelimiter=';'; csvEncoding='UTF8' }
            }
            Add-Member -InputObject $_cfg -NotePropertyName 'language' -NotePropertyValue $script:Lang -Force
            ConvertTo-Json $_cfg -Depth 10 | Set-Content $_configPath -Encoding UTF8
        } catch {}
    }

    . (Join-Path $AppRoot "src/Lang/$script:Lang.ps1")

    $shouldContinue = Show-StartupCheck -AppRoot $AppRoot
}
catch {
    Write-Error "Startup error: $_"
    Read-Host "Press Enter to quit"
    exit 1
}

if (-not $shouldContinue) { exit 0 }

try {
    & (Join-Path $AppRoot 'src/App.ps1') -RootPath $AppRoot -Lang $script:Lang
}
catch {
    Write-Error "Fatal error: $_"
    Write-Error $_.ScriptStackTrace
    Read-Host "Press Enter to quit"
    exit 1
}
