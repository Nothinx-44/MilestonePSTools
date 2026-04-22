<#
.SYNOPSIS
    Capture un snapshot de toutes les cameras du VMS Milestone.
#>

function Get-SnapshotAll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log,

        [Parameter()]
        [scriptblock]$Cancel = { $false },

        [Parameter()]
        [scriptblock]$ReportProgress = {},

        [Parameter()]
        [nullable[datetime]]$SnapshotTime = $null
    )

    $snapshotDir = Join-Path $Config.outputDirectory 'Snapshots'
    if (-not (Test-Path $snapshotDir)) {
        New-Item -Path $snapshotDir -ItemType Directory -Force | Out-Null
    }

    & $Log "Recuperation de la liste des cameras..."
    $cameras = Get-VmsCamera
    $total   = @($cameras).Count
    & $Log "$total cameras trouvees. Capture en cours..."

    if ($SnapshotTime) {
        & $Log "Mode historique : $($SnapshotTime.ToString('dd/MM/yyyy HH:mm'))"
    }

    $count = 0
    foreach ($cam in $cameras) {
        if (& $Cancel) {
            & $Log "AVERTISSEMENT: Operation annulee apres $count / $total cameras."
            break
        }

        $count++
        & $ReportProgress $count $total
        & $Log "[$count/$total] Snapshot de '$($cam.Name)'..."

        try {
            if ($SnapshotTime) {
                $cam | Get-Snapshot `
                    -UseFriendlyName `
                    -Behavior GetNearest `
                    -Time $SnapshotTime `
                    -Quality $Config.snapshotQuality `
                    -Save `
                    -Path $snapshotDir
            }
            else {
                $cam | Get-Snapshot `
                    -UseFriendlyName `
                    -Behavior GetEnd `
                    -Quality $Config.snapshotQuality `
                    -Save `
                    -Path $snapshotDir
            }
        }
        catch {
            & $Log "ERREUR sur '$($cam.Name)': $_"
        }
    }

    if (-not (& $Cancel)) {
        & $Log "$count snapshots traites dans : $snapshotDir"
    }
}
