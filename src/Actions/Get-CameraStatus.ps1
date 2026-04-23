<#
.SYNOPSIS
    Etat temps reel de toutes les cameras via l'Event Server Milestone (Get-ItemState).
#>

function Get-CameraStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log,

        [Parameter()]
        [scriptblock]$Cancel = { $false },

        [Parameter()]
        [scriptblock]$ReportProgress = {}
    )

    & $Log "Interrogation de l'Event Server..."

    try {
        $states = @(Get-ItemState -CamerasOnly -ErrorAction Stop)
    }
    catch {
        & $Log "ERREUR: Get-ItemState : $_"
        return
    }

    $total = $states.Count
    & $Log "$total cameras interrogees."

    $rows  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $ok    = 0
    $ko    = 0
    $count = 0

    $disabled = 0

    foreach ($state in $states) {
        if (& $Cancel) { & $Log "AVERTISSEMENT: Operation annulee apres $count / $total." ; break }

        $count++
        & $ReportProgress $count $total

        $name = if ($state.Name -and $state.Name -ne 'Not available') { $state.Name }
                else { $state.FQID.ObjectId.ToString() }

        switch ($state.State) {
            'Responding' {
                $ok++
                & $Log "  [OK] $name"
            }
            'Disabled' {
                $disabled++
                # Camera desactivee volontairement, pas une erreur
            }
            default {
                $ko++
                & $Log "  ERREUR: $name : $($state.State)"
            }
        }

        if ($state.State -ne 'Disabled') {
            $rows.Add([PSCustomObject]@{
                Nom  = $name
                Etat = if ($state.State -eq 'Responding') { 'OK' } else { $state.State }
                Type = $state.ItemType
            })
        }
    }

    if ($disabled -gt 0) {
        & $Log "  ($disabled camera(s) desactivee(s) ignoree(s))"
    }

    if ($ko -gt 0) {
        & $Log "AVERTISSEMENT: $ko camera(s) hors ligne ou en erreur sur $($total - $disabled) actives."
    }
    else {
        & $Log "Toutes les cameras actives ($ok) sont operationnelles."
    }

    if ($rows.Count -gt 0 -and -not (& $Cancel)) {
        if (-not (Test-Path $Config.outputDirectory)) {
            New-Item -Path $Config.outputDirectory -ItemType Directory -Force | Out-Null
        }
        $csvPath = Join-Path $Config.outputDirectory 'Etat_Cameras.csv'
        $rows | Export-Csv -Path $csvPath -NoTypeInformation `
            -Encoding $Config.csvEncoding -Delimiter $Config.csvDelimiter
        & $Log "Rapport exporte : $csvPath"
    }
}
