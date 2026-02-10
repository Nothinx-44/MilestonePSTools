<#
.SYNOPSIS
    Point d'entree principal de l'application Milestone Toolkit.
.DESCRIPTION
    Charge la configuration, initialise les modules, etablit la connexion au serveur
    Milestone, charge l'interface WPF et connecte les evenements aux actions.
#>

param(
    [Parameter()]
    [string]$RootPath
)

# ============================================================
# 1. CHEMINS ET CONFIGURATION
# ============================================================

$script:AppRoot = $RootPath
$script:SrcPath = Join-Path $AppRoot 'src'

# Charger la configuration
$configPath = Join-Path $AppRoot 'config.json'
if (Test-Path $configPath) {
    $configRaw = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
}
else {
    Write-Warning "Fichier config.json introuvable. Utilisation des valeurs par defaut."
    $configRaw = [PSCustomObject]@{
        outputDirectory = './Output'
        snapshotQuality = 95
        csvDelimiter    = ';'
        csvEncoding     = 'UTF8'
    }
}

# Convertir les chemins relatifs en absolus
$outputDir = $configRaw.outputDirectory
if (-not [System.IO.Path]::IsPathRooted($outputDir)) {
    $outputDir = Join-Path $AppRoot $outputDir
}

$script:Config = @{
    outputDirectory = $outputDir
    snapshotQuality = [int]$configRaw.snapshotQuality
    csvDelimiter    = $configRaw.csvDelimiter
    csvEncoding     = $configRaw.csvEncoding
    logDirectory    = Join-Path $AppRoot 'Logs'
}

# ============================================================
# 2. CHARGEMENT DES SCRIPTS
# ============================================================

# Core
. (Join-Path $SrcPath 'Core/Initialize-Modules.ps1')
. (Join-Path $SrcPath 'Core/Write-ActivityLog.ps1')
. (Join-Path $SrcPath 'Core/Invoke-PtzPreset.ps1')

# Actions
. (Join-Path $SrcPath 'Actions/Get-SnapshotSelected.ps1')
. (Join-Path $SrcPath 'Actions/Get-SnapshotAll.ps1')
. (Join-Path $SrcPath 'Actions/Export-HardwareReport.ps1')
. (Join-Path $SrcPath 'Actions/Set-CameraGroupByModel.ps1')
. (Join-Path $SrcPath 'Actions/Get-PtzPresetSnapshot.ps1')

# ============================================================
# 3. INITIALISATION DES MODULES ET CONNEXION
# ============================================================

$initLog = { param($Message) Write-Host $Message }
Initialize-RequiredModules -Log $initLog

Write-Host 'Connexion au serveur Milestone...' -ForegroundColor Cyan
Connect-ManagementServer -ShowDialog -AcceptEula -Force
Write-Host 'Connecte.' -ForegroundColor Green

# ============================================================
# 4. MASQUER LA CONSOLE
# ============================================================

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class ConsoleHelper {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$consoleHandle = [ConsoleHelper]::GetConsoleWindow()
[ConsoleHelper]::ShowWindow($consoleHandle, 0) | Out-Null

# ============================================================
# 5. CHARGEMENT DE L'INTERFACE WPF
# ============================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$xamlPath = Join-Path $SrcPath 'UI/MainWindow.xaml'
$xamlContent = Get-Content $xamlPath -Raw -Encoding UTF8

$xamlReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlContent))
$script:Window = [System.Windows.Markup.XamlReader]::Load($xamlReader)

# ============================================================
# 6. REFERENCES AUX ELEMENTS UI
# ============================================================

$script:LogOutput        = $Window.FindName('LogOutput')
$script:ActionStatus     = $Window.FindName('ActionStatus')
$script:ProgressBar      = $Window.FindName('ProgressBar')
$script:StatusIndicator  = $Window.FindName('StatusIndicator')
$script:StatusText       = $Window.FindName('StatusText')

$script:BtnSnapshotSelected = $Window.FindName('BtnSnapshotSelected')
$script:BtnSnapshotAll      = $Window.FindName('BtnSnapshotAll')
$script:BtnPtzSnapshot      = $Window.FindName('BtnPtzSnapshot')
$script:BtnExportHardware   = $Window.FindName('BtnExportHardware')
$script:BtnGroupByModel     = $Window.FindName('BtnGroupByModel')
$script:BtnClearLog         = $Window.FindName('BtnClearLog')

# Mettre a jour le statut de connexion
$StatusIndicator.Fill = [System.Windows.Media.Brushes]::LightGreen
$StatusText.Text = 'Connecte'

# ============================================================
# 7. FONCTIONS UTILITAIRES UI
# ============================================================

function Write-UILog {
    param([string]$Message)
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $logLine = "[$timestamp] $Message`r`n"
    $LogOutput.AppendText($logLine)
    $LogOutput.ScrollToEnd()

    # Forcer le rafraichissement de l'UI
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [Action]{}
    )

    # Ecrire aussi dans le fichier de log
    Write-ActivityLog -Message $Message -Level 'INFO' -LogDirectory $Config.logDirectory
}

# Liste de tous les boutons d'action pour activer/desactiver en bloc
$script:ActionButtons = @(
    $BtnSnapshotSelected, $BtnSnapshotAll, $BtnPtzSnapshot,
    $BtnExportHardware, $BtnGroupByModel
)

function Set-UIBusy {
    param([string]$ActionName)
    foreach ($btn in $ActionButtons) { $btn.IsEnabled = $false }
    $ActionStatus.Text = $ActionName
    $ProgressBar.IsIndeterminate = $true
    $ProgressBar.Visibility = [System.Windows.Visibility]::Visible
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
}

function Set-UIReady {
    foreach ($btn in $ActionButtons) { $btn.IsEnabled = $true }
    $ActionStatus.Text = 'Pret'
    $ProgressBar.IsIndeterminate = $false
    $ProgressBar.Visibility = [System.Windows.Visibility]::Collapsed
    $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
}

function Invoke-Action {
    <#
    .SYNOPSIS
        Execute une action en gerant l'etat de l'UI (busy/ready, logging, erreurs).
    #>
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Set-UIBusy -ActionName $Name
    Write-UILog "--- $Name ---"

    try {
        & $Action
        Write-UILog "Action terminee avec succes."
    }
    catch {
        Write-UILog "ERREUR: $_"
        Write-ActivityLog -Message $_.Exception.Message -Level 'ERROR' -LogDirectory $Config.logDirectory
    }
    finally {
        Set-UIReady
    }
}

# ============================================================
# 8. BRANCHEMENT DES EVENEMENTS
# ============================================================

$logCallback = { param($Message) Write-UILog $Message }

$BtnSnapshotSelected.Add_Click({
    Invoke-Action -Name 'Snapshot - Selection' -Action {
        Get-SnapshotSelected -Config $Config -Log $logCallback
    }
})

$BtnSnapshotAll.Add_Click({
    Invoke-Action -Name 'Snapshot - Toutes les cameras' -Action {
        Get-SnapshotAll -Config $Config -Log $logCallback
    }
})

$BtnPtzSnapshot.Add_Click({
    Invoke-Action -Name 'Snapshot - Presets PTZ' -Action {
        Get-PtzPresetSnapshot -Config $Config -Log $logCallback
    }
})

$BtnExportHardware.Add_Click({
    Invoke-Action -Name 'Export Hardware' -Action {
        Export-HardwareReport -Config $Config -Log $logCallback
    }
})

$BtnGroupByModel.Add_Click({
    Invoke-Action -Name 'Grouper par Modele' -Action {
        Set-CameraGroupByModel -Config $Config -Log $logCallback
    }
})

$BtnClearLog.Add_Click({
    $LogOutput.Clear()
})

# Gestion de la fermeture
$Window.Add_Closing({
    Write-ActivityLog -Message 'Fermeture de l application' -Level 'INFO' -LogDirectory $Config.logDirectory
    try { Disconnect-ManagementServer } catch {}

    # Reafficher la console
    [ConsoleHelper]::ShowWindow($consoleHandle, 5) | Out-Null
})

# ============================================================
# 9. AFFICHAGE
# ============================================================

Write-UILog "Application demarree. Connecte au serveur Milestone."
Write-UILog "Repertoire de sortie : $($Config.outputDirectory)"

$Window.Add_Loaded({ $Window.Activate() })
[void]$Window.ShowDialog()
