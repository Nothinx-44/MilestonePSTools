<#
.SYNOPSIS
    Lance l'application Milestone Toolkit.
.DESCRIPTION
    Point d'entree utilisateur. Double-cliquer ou executer depuis PowerShell.
    Verifie la version de PowerShell et lance le bootstrap applicatif.
.NOTES
    Prerequis : Windows PowerShell 5.1+ ou PowerShell 7+ sur Windows.
#>

#Requires -Version 5.1

$AppRoot = if ($PSScriptRoot) {
    $PSScriptRoot
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

    $script:Lang = Show-LanguagePicker
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
