<#
.SYNOPSIS
    Capture un snapshot a chaque position preset PTZ des cameras selectionnees.
.DESCRIPTION
    Ouvre le selecteur de camera, filtre les cameras PTZ avec presets,
    parcourt chaque preset, deplace la camera et capture un snapshot.
.PARAMETER Config
    Hashtable de configuration (outputDirectory, snapshotQuality).
.PARAMETER Log
    Scriptblock callback pour logger vers l'UI.
#>

function Get-PtzPresetSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log
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

    $cameraCount = @($cameras).Count
    & $Log "$cameraCount camera(s) PTZ avec presets trouvee(s)."

    foreach ($camera in $cameras) {
        $presets = $camera.PtzPresetFolder.PtzPresets
        & $Log "Camera '$($camera.Name)' : $($presets.Count) preset(s)."

        foreach ($ptzPreset in $presets) {
            & $Log "  Deplacement vers preset '$($ptzPreset.Name)'..."

            try {
                Invoke-PtzPreset -PtzPreset $ptzPreset -VerifyCoordinates
            }
            catch {
                & $Log "  AVERTISSEMENT: Verification de position echouee: $_"
            }

            & $Log "  Capture du snapshot..."
            $snapshotParams = @{
                Quality  = $Config.snapshotQuality
                Save     = $true
                Path     = $outputDir
                FileName = "$($camera.Name) -- $($ptzPreset.Name).jpg"
            }
            $camera | Get-Snapshot @snapshotParams -Behavior GetEnd

            & $Log "  Snapshot '$($ptzPreset.Name)' enregistre."
        }
    }

    & $Log "Capture PTZ terminee. Fichiers dans : $outputDir"
}
