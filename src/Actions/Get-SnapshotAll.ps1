<#
.SYNOPSIS
    Capture un snapshot de toutes les cameras du VMS Milestone.
.PARAMETER Config
    Hashtable de configuration (outputDirectory, snapshotQuality).
.PARAMETER Log
    Scriptblock callback pour logger vers l'UI.
#>

function Get-SnapshotAll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log
    )

    $snapshotDir = Join-Path $Config.outputDirectory 'Snapshots'
    if (-not (Test-Path $snapshotDir)) {
        New-Item -Path $snapshotDir -ItemType Directory -Force | Out-Null
    }

    & $Log "Recuperation de la liste des cameras..."
    $cameras = Get-VmsCamera
    & $Log "$($cameras.Count) cameras trouvees. Capture en cours..."

    $count = 0
    foreach ($cam in $cameras) {
        $count++
        & $Log "[$count/$($cameras.Count)] Snapshot de '$($cam.Name)'..."

        try {
            $cam | Get-Snapshot `
                -UseFriendlyName `
                -Behavior GetEnd `
                -Quality $Config.snapshotQuality `
                -Save `
                -Path $snapshotDir
        }
        catch {
            & $Log "ERREUR sur '$($cam.Name)': $_"
        }
    }

    & $Log "Terminee. $count snapshots traites dans : $snapshotDir"
}
