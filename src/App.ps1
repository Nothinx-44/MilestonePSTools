<#
.SYNOPSIS
    Point d'entree principal de l'application Milestone Toolkit.
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

$configPath = Join-Path $AppRoot 'config.json'
if (Test-Path $configPath) {
    $configRaw = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
}
else {
    Write-Warning "Fichier config.json introuvable. Utilisation des valeurs par defaut."
    $configRaw = [PSCustomObject]@{
        installMode     = 'Auto'
        outputDirectory = './Output'
        snapshotQuality = 95
        csvDelimiter    = ';'
        csvEncoding     = 'UTF8'
    }
}

$outputDir = $configRaw.outputDirectory
if (-not [System.IO.Path]::IsPathRooted($outputDir)) {
    $outputDir = Join-Path $AppRoot $outputDir
}

$script:DependenciesPath = Join-Path $AppRoot 'Dependencies'

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

. (Join-Path $SrcPath 'Core/Initialize-Modules.ps1')
. (Join-Path $SrcPath 'Core/Write-ActivityLog.ps1')
. (Join-Path $SrcPath 'Core/Invoke-PtzPreset.ps1')

. (Join-Path $SrcPath 'Actions/Get-SnapshotSelected.ps1')
. (Join-Path $SrcPath 'Actions/Get-SnapshotAll.ps1')
. (Join-Path $SrcPath 'Actions/Export-HardwareReport.ps1')
. (Join-Path $SrcPath 'Actions/Set-CameraGroupByModel.ps1')
. (Join-Path $SrcPath 'Actions/Get-PtzPresetSnapshot.ps1')

# ============================================================
# 3. INITIALISATION DES MODULES ET CONNEXION
# ============================================================

$installMode = $configRaw.installMode
if (-not $installMode -or $installMode -eq 'Auto') {
    if (Test-Path $DependenciesPath) {
        $installMode = 'Offline'
        Write-Host 'Mode Offline detecte (dossier Dependencies/ present).' -ForegroundColor Yellow
    }
    else {
        $installMode = 'Online'
    }
}

$initLog = { param($Message) Write-Host $Message }
Initialize-RequiredModules -InstallMode $installMode -DependenciesPath $DependenciesPath -Log $initLog

Write-Host 'Connexion au serveur Milestone...' -ForegroundColor Cyan
Connect-ManagementServer -ShowDialog -AcceptEula -Force
Write-Host 'Connecte.' -ForegroundColor Green

# ============================================================
# 4. MASQUER LA CONSOLE
# ============================================================

Add-Type -Name ConsoleHelper -Namespace '' -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@ -ErrorAction SilentlyContinue

$consoleHandle = [ConsoleHelper]::GetConsoleWindow()
[ConsoleHelper]::ShowWindow($consoleHandle, 0) | Out-Null

# ============================================================
# 5. CHARGEMENT DE L'INTERFACE WPF
# ============================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$xamlPath    = Join-Path $SrcPath 'UI/MainWindow.xaml'
$xamlContent = Get-Content $xamlPath -Raw -Encoding UTF8
$xamlReader  = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlContent))
$script:Window = [System.Windows.Markup.XamlReader]::Load($xamlReader)

# ============================================================
# 6. REFERENCES AUX ELEMENTS UI
# ============================================================

$script:LogOutput        = $Window.FindName('LogOutput')
$script:ActionStatus     = $Window.FindName('ActionStatus')
$script:ProgressBar      = $Window.FindName('ProgressBar')
$script:StatusIndicator  = $Window.FindName('StatusIndicator')
$script:StatusText       = $Window.FindName('StatusText')
$script:OutputDirText    = $Window.FindName('OutputDirText')

$script:BtnSnapshotSelected = $Window.FindName('BtnSnapshotSelected')
$script:BtnSnapshotAll      = $Window.FindName('BtnSnapshotAll')
$script:BtnPtzSnapshot      = $Window.FindName('BtnPtzSnapshot')
$script:BtnExportHardware   = $Window.FindName('BtnExportHardware')
$script:BtnGroupByModel     = $Window.FindName('BtnGroupByModel')
$script:BtnClearLog         = $Window.FindName('BtnClearLog')
$script:BtnCancel           = $Window.FindName('BtnCancel')
$script:BtnOutputDir        = $Window.FindName('BtnOutputDir')

# Initialiser le document RichTextBox (supprimer le paragraphe vide par defaut)
$script:LogOutput.Document.Blocks.Clear()
$script:LogOutput.Document.PagePadding = [System.Windows.Thickness]::new(16, 12, 16, 12)

# Statut de connexion
$script:StatusIndicator.Fill = [System.Windows.Media.Brushes]::LightGreen
$script:StatusText.Text      = 'Connecte'

# Dossier de sortie dans la sidebar
$script:OutputDirText.Text   = $script:Config.outputDirectory

# ============================================================
# 7. ETAT PARTAGE POUR CANCEL ET PROGRESS
# ============================================================

$script:CancelRequested = $false

# Callbacks passes aux actions
$script:IsCancelled = { $script:CancelRequested }

$script:ReportProgress = {
    param([int]$Current, [int]$Total)
    if ($Total -gt 0) {
        $script:ProgressBar.IsIndeterminate = $false
        $script:ProgressBar.Maximum         = [double]$Total
        $script:ProgressBar.Value           = [double]$Current
    }
    # Pomper le dispatcher : traite les evenements en attente (clic Annuler inclus)
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background, [Action]{}
    )
}

# ============================================================
# 8. FONCTIONS UTILITAIRES UI
# ============================================================

function Write-UILog {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $fullMsg   = "[$timestamp] $Message"

    $brush = switch ($Level) {
        'ERROR'   { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(243,139,168)) }
        'WARN'    { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(249,168, 37)) }
        'SUCCESS' { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(166,227,161)) }
        'ACTION'  { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(137,180,250)) }
        default   { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(205,214,244)) }
    }

    $run  = [System.Windows.Documents.Run]::new($fullMsg)
    $run.Foreground = $brush
    $para = [System.Windows.Documents.Paragraph]::new($run)
    $para.Margin     = [System.Windows.Thickness]::new(0)
    $para.LineHeight = 20
    $script:LogOutput.Document.Blocks.Add($para)
    $script:LogOutput.ScrollToEnd()

    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Render, [Action]{}
    )

    $fileLevel = if ($Level -in 'ACTION','INFO') { 'INFO' } else { $Level }
    Write-ActivityLog -Message $Message -Level $fileLevel -LogDirectory $script:Config.logDirectory
}

$script:ActionButtons = @(
    $BtnSnapshotSelected, $BtnSnapshotAll, $BtnPtzSnapshot,
    $BtnExportHardware, $BtnGroupByModel
)

function Set-UIBusy {
    param([string]$ActionName)
    foreach ($btn in $script:ActionButtons) { $btn.IsEnabled = $false }
    $script:BtnCancel.Visibility        = [System.Windows.Visibility]::Visible
    $script:ActionStatus.Text           = $ActionName
    $script:ProgressBar.IsIndeterminate = $true
    $script:ProgressBar.Visibility      = [System.Windows.Visibility]::Visible
    $Window.Cursor                      = [System.Windows.Input.Cursors]::Wait
}

function Set-UIReady {
    foreach ($btn in $script:ActionButtons) { $btn.IsEnabled = $true }
    $script:BtnCancel.Visibility        = [System.Windows.Visibility]::Collapsed
    $script:ActionStatus.Text           = 'Pret'
    $script:ProgressBar.IsIndeterminate = $false
    $script:ProgressBar.Value           = 0
    $script:ProgressBar.Visibility      = [System.Windows.Visibility]::Collapsed
    $Window.Cursor                      = [System.Windows.Input.Cursors]::Arrow
}

function Invoke-Action {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    $script:CancelRequested = $false
    Set-UIBusy -ActionName $Name
    Write-UILog "--- $Name ---" 'ACTION'

    try {
        & $Action

        if ($script:CancelRequested) {
            Write-UILog "Operation annulee." 'WARN'
        }
        else {
            Write-UILog "Action terminee avec succes." 'SUCCESS'
        }
    }
    catch {
        Write-UILog "ERREUR: $_" 'ERROR'
        Write-ActivityLog -Message $_.Exception.Message -Level 'ERROR' -LogDirectory $script:Config.logDirectory
    }
    finally {
        Set-UIReady
    }
}

# ============================================================
# 9. BRANCHEMENT DES EVENEMENTS
# ============================================================

$logCallback = {
    param([string]$Message)
    $level = if ($Message -match '^ERREUR|^ERROR')           { 'ERROR'   }
             elseif ($Message -match '^AVERTISSEMENT|^WARN') { 'WARN'    }
             else                                            { 'INFO'    }
    Write-UILog -Message $Message -Level $level
}

$BtnSnapshotSelected.Add_Click({
    Invoke-Action -Name 'Snapshot - Selection' -Action {
        Get-SnapshotSelected -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled
    }
})

$BtnSnapshotAll.Add_Click({
    Invoke-Action -Name 'Snapshot - Toutes les cameras' -Action {
        Get-SnapshotAll -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnPtzSnapshot.Add_Click({
    Invoke-Action -Name 'Snapshot - Presets PTZ' -Action {
        Get-PtzPresetSnapshot -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnExportHardware.Add_Click({
    Invoke-Action -Name 'Export Hardware' -Action {
        Export-HardwareReport -Config $script:Config -Log $logCallback
    }
})

$BtnGroupByModel.Add_Click({
    Invoke-Action -Name 'Grouper par Modele' -Action {
        Set-CameraGroupByModel -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnClearLog.Add_Click({
    $script:LogOutput.Document.Blocks.Clear()
})

$BtnCancel.Add_Click({
    $script:CancelRequested = $true
    $script:BtnCancel.IsEnabled = $false
    $script:ActionStatus.Text = 'Annulation en cours...'
})

$BtnOutputDir.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description  = 'Choisir le dossier de sortie'
    $dialog.SelectedPath = $script:Config.outputDirectory
    $dialog.ShowNewFolderButton = $true

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:Config.outputDirectory = $dialog.SelectedPath
        $script:OutputDirText.Text     = $dialog.SelectedPath
        Write-UILog "Dossier de sortie change : $($dialog.SelectedPath)" 'INFO'
    }
})

$Window.Add_Closing({
    Write-ActivityLog -Message 'Fermeture de l application' -Level 'INFO' -LogDirectory $script:Config.logDirectory
    try { Disconnect-ManagementServer } catch {}
    [ConsoleHelper]::ShowWindow($consoleHandle, 5) | Out-Null
})

# ============================================================
# 10. AFFICHAGE
# ============================================================

Write-UILog "Application demarree. Connecte au serveur Milestone." 'SUCCESS'
Write-UILog "Repertoire de sortie : $($script:Config.outputDirectory)"

$Window.Add_Loaded({ $Window.Activate() })
[void]$Window.ShowDialog()
