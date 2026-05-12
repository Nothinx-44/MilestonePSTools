<#
.SYNOPSIS
    Bootstrap interne de Milestone Toolkit. Appele par Launch.bat.
    Ne pas executer directement — utiliser Launch.bat a la racine du projet.
#>

#Requires -Version 5.1

# Version centrale — modifier ici uniquement
$script:AppVersion = '4.9.0'

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

function Get-GitHubLatestRelease {
    param(
        [Parameter(Mandatory)] [string]$Repository
    )

    try {
        $headers = @{ 'User-Agent' = 'MilestoneToolkitUpdater' }
        return Invoke-RestMethod -Uri "https://api.github.com/repos/$Repository/releases/latest" -Headers $headers -ErrorAction Stop
    }
    catch {
        return $null
    }
}

function Get-ComparableVersion {
    param([string]$VersionString)

    if (-not $VersionString) { return $null }
    $clean = $VersionString.Trim()
    if ($clean.StartsWith('v')) { $clean = $clean.Substring(1) }
    try {
        return [version]$clean
    }
    catch {
        return $null
    }
}

function Invoke-AutoUpdate {
    param(
        [Parameter(Mandatory)] [string]$Repository,
        [Parameter(Mandatory)] [string]$AppRoot
    )

    Add-Type -AssemblyName System.Windows.Forms

    $release = Get-GitHubLatestRelease -Repository $Repository
    if (-not $release) { return }

    $remoteVersion = Get-ComparableVersion -VersionString $release.tag_name
    $currentVersion = Get-ComparableVersion -VersionString $script:AppVersion
    if (-not $remoteVersion -or -not $currentVersion) { return }
    if ($remoteVersion -le $currentVersion) { return }

    $a = [char]0x00E0  # a grave (a) — encodage independant du systeme
    $message = "Une nouvelle version est disponible : v$($release.tag_name). Voulez-vous mettre $a jour maintenant ?"
    $caption = "Milestone Toolkit - Mise $a jour disponible"
    $result = [System.Windows.Forms.MessageBox]::Show($message, $caption, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $tempDir = Join-Path $env:TEMP ("MilestoneToolkitUpdate_{0}" -f ([guid]::NewGuid()))
    $zipPath  = Join-Path $tempDir 'release.zip'
    $extract  = Join-Path $tempDir 'extract'
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
    New-Item -Path $extract -ItemType Directory -Force | Out-Null

    try {
        $headers = @{ 'User-Agent' = 'MilestoneToolkitUpdater' }
        Invoke-WebRequest -Uri $release.zipball_url -OutFile $zipPath -Headers $headers -UseBasicParsing -ErrorAction Stop
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extract)

        $srcRoot = Get-ChildItem -Path $extract | Where-Object { $_.PSIsContainer } | Select-Object -First 1
        if (-not $srcRoot) { return }

        $updaterScript = Join-Path $tempDir 'Update-Tool.ps1'
        $batPath = Join-Path $AppRoot 'Demarrer Milestone Toolkit.bat'
        $scriptContent = @"
param(
    [string]
    [string]
    [string]
)
Start-Sleep -Seconds 2
for ($i = 0; $i -lt 20; $i++) {
    try {
        Get-ChildItem -Path $args[0] -ErrorAction Stop | Out-Null
        break
    }
    catch {
        Start-Sleep -Milliseconds 250
    }
}
try {
    Copy-Item -Path (Join-Path $args[1] '*') -Destination $args[0] -Recurse -Force -ErrorAction Stop
}
catch {
}
Start-Process -FilePath $args[2]
"@
        Set-Content -Path $updaterScript -Value $scriptContent -Encoding UTF8

        Start-Process -FilePath (Get-Command powershell).Source -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$updaterScript`"","`"$AppRoot`"","`"$($srcRoot.FullName)`"","`"$batPath`"" -WindowStyle Hidden
        exit 0
    }
    catch {
        return
    }
}

# La verification de mise a jour est geree dans Show-StartupCheck (fenetre des dependances)

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
