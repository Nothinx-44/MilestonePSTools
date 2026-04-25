<#
.SYNOPSIS
    Capture un snapshot d'une camera selectionnee via le dialogue Milestone.
#>

function Get-SnapshotSelected {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log,

        [Parameter()]
        [scriptblock]$Cancel = { $false },

        [Parameter()]
        [nullable[datetime]]$SnapshotTime = $null
    )

    & $Log "Ouverture du selecteur de camera..."

    $camera = Select-Camera
    if (-not $camera) {
        & $Log "Aucune camera selectionnee. Operation annulee."
        return
    }

    if (& $Cancel) { return }

    $snapshotDir = Join-Path $Config.outputDirectory 'Snapshots'
    if (-not (Test-Path $snapshotDir)) {
        New-Item -Path $snapshotDir -ItemType Directory -Force | Out-Null
    }

    if ($SnapshotTime) {
        & $Log "Mode historique : $($SnapshotTime.ToString('dd/MM/yyyy HH:mm'))"
    }
    & $Log "Capture du snapshot de '$($camera.Name)'..."

    try {
        if ($SnapshotTime) {
            $camera | Get-Snapshot `
                -UseFriendlyName `
                -Behavior GetNearest `
                -Time $SnapshotTime `
                -Quality $Config.snapshotQuality `
                -Save `
                -Path $snapshotDir
        }
        else {
            $camera | Get-Snapshot `
                -UseFriendlyName `
                -Behavior GetEnd `
                -Quality $Config.snapshotQuality `
                -Save `
                -Path $snapshotDir
        }
        & $Log "Snapshot enregistre dans : $snapshotDir"
    }
    catch {
        & $Log "ERREUR: Echec du snapshot de '$($camera.Name)' : $_"
    }
}
