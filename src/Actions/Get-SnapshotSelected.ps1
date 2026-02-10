<#
.SYNOPSIS
    Capture un snapshot d'une camera selectionnee via le dialogue Milestone.
.PARAMETER Config
    Hashtable de configuration (outputDirectory, snapshotQuality).
.PARAMETER Log
    Scriptblock callback pour logger vers l'UI. Usage : & $Log "message"
#>

function Get-SnapshotSelected {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log
    )

    & $Log "Ouverture du selecteur de camera..."

    $camera = Select-Camera
    if (-not $camera) {
        & $Log "Aucune camera selectionnee. Operation annulee."
        return
    }

    $snapshotDir = Join-Path $Config.outputDirectory 'Snapshots'
    if (-not (Test-Path $snapshotDir)) {
        New-Item -Path $snapshotDir -ItemType Directory -Force | Out-Null
    }

    & $Log "Capture du snapshot de '$($camera.Name)'..."

    $camera | Get-Snapshot `
        -UseFriendlyName `
        -Behavior GetEnd `
        -Quality $Config.snapshotQuality `
        -Save `
        -Path $snapshotDir

    & $Log "Snapshot enregistre dans : $snapshotDir"
}
