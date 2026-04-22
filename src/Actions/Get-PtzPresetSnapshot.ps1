<#
.SYNOPSIS
    Capture un snapshot a chaque position preset PTZ des cameras selectionnees.
#>

function Get-PtzPresetSnapshot {
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

    $outputDir = Join-Path $Config.outputDirectory 'PTZ_Snapshots'
    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    }

    & $Log "Selection des cameras PTZ..."

    $cameras = Select-Camera |
        Where-Object { $_.Enabled } |
        Get-Camera |
        Where-Object { $_.Enabled -and $_.PtzPresetFolder.PtzPresets.Count -gt 0 }

    if (-not $cameras) {
        & $Log "Aucune camera PTZ avec presets selectionnee."
        return
    }

    $cameraList  = @($cameras)
    $totalCams   = $cameraList.Count
    & $Log "$totalCams camera(s) PTZ avec presets trouvee(s)."

    # Compter le total de presets pour la barre de progression
    $totalPresets = ($cameraList | ForEach-Object { $_.PtzPresetFolder.PtzPresets.Count } | Measure-Object -Sum).Sum
    $donePresets  = 0

    foreach ($camera in $cameraList) {
        if (& $Cancel) {
            & $Log "AVERTISSEMENT: Operation annulee."
            break
        }

        $presets = $camera.PtzPresetFolder.PtzPresets
        & $Log "Camera '$($camera.Name)' : $($presets.Count) preset(s)."

        foreach ($ptzPreset in $presets) {
            if (& $Cancel) {
                & $Log "AVERTISSEMENT: Operation annulee."
                break
            }

            $donePresets++
            & $ReportProgress $donePresets $totalPresets

            & $Log "  Deplacement vers preset '$($ptzPreset.Name)'..."

            try {
                Invoke-PtzPreset -PtzPreset $ptzPreset -VerifyCoordinates
            }
            catch {
                & $Log "  AVERTISSEMENT: Verification de position echouee: $_"
            }

            & $Log "  Capture du snapshot..."
            try {
                $camera | Get-Snapshot `
                    -Quality  $Config.snapshotQuality `
                    -Save     $true `
                    -Path     $outputDir `
                    -FileName "$($camera.Name) -- $($ptzPreset.Name).jpg" `
                    -Behavior GetEnd
                & $Log "  Snapshot '$($ptzPreset.Name)' enregistre."
            }
            catch {
                & $Log "ERREUR: Snapshot '$($ptzPreset.Name)' echoue : $_"
            }
        }
    }

    if (-not (& $Cancel)) {
        & $Log "Capture PTZ terminee. Fichiers dans : $outputDir"
    }
}
