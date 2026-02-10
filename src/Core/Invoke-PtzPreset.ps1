<#
.SYNOPSIS
    Deplace une camera PTZ vers un preset et verifie optionnellement la position.
.DESCRIPTION
    Envoie une commande MIP pour declencher un preset PTZ sur une camera Milestone.
    Peut verifier que la camera a atteint la position cible (coordonnees absolues).
#>

function Invoke-PtzPreset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [VideoOS.Platform.ConfigurationItems.PtzPreset]
        $PtzPreset,

        [Parameter()]
        [switch]$VerifyCoordinates,

        [Parameter()]
        [double]$Tolerance = 0.001,

        [Parameter()]
        [int]$Timeout = 5
    )

    process {
        # Extraire l'ID camera depuis le chemin du preset
        $cameraId = if ($PtzPreset.ParentItemPath -match 'Camera\[(.{36})\]') {
            $Matches[1]
        }
        else {
            Write-Error "Impossible de parser l'ID camera depuis ParentItemPath '$($PtzPreset.ParentItemPath)'"
            return
        }

        $camera     = Get-Camera -Id $cameraId
        $cameraItem = $camera | Get-PlatformItem
        $presetItem = [VideoOS.Platform.Configuration]::Instance.GetItem(
            [guid]::new($PtzPreset.Id),
            [VideoOS.Platform.Kind]::Preset
        )

        # Declencher le preset
        $params = @{
            MessageId          = 'Control.TriggerCommand'
            DestinationEndpoint = $presetItem.FQID
            UseEnvironmentManager = $true
        }
        Send-MipMessage @params

        if (-not $VerifyCoordinates) {
            return
        }

        # Verification des coordonnees absolues
        $hasPan  = $cameraItem.Properties['pan']  -eq 'Absolute'
        $hasTilt = $cameraItem.Properties['tilt'] -eq 'Absolute'
        $hasZoom = $cameraItem.Properties['zoom'] -eq 'Absolute'

        if (-not ($hasPan -and $hasTilt -and $hasZoom)) {
            Write-Warning "La camera n'utilise pas le positionnement PTZ absolu. Verification des coordonnees ignoree."
            return
        }

        $positionReached = $false
        $stopwatch = [Diagnostics.Stopwatch]::StartNew()

        while ($stopwatch.ElapsedMilliseconds -lt ($Timeout * 1000)) {
            $position = Send-MipMessage -MessageId Control.PTZGetAbsoluteRequest `
                -DestinationEndpoint $cameraItem.FQID -UseEnvironmentManager

            $xDiff = [Math]::Abs([Math]::Abs($position.Pan)  - [Math]::Abs($PtzPreset.Pan))
            $yDiff = [Math]::Abs([Math]::Abs($position.Tilt) - [Math]::Abs($PtzPreset.Tilt))
            $zDiff = [Math]::Abs([Math]::Abs($position.Zoom) - [Math]::Abs($PtzPreset.Zoom))

            if ($xDiff -le $Tolerance -and $yDiff -le $Tolerance -and $zDiff -le $Tolerance) {
                $positionReached = $true
                Start-Sleep -Milliseconds 2500
                break
            }

            Start-Sleep -Milliseconds 100
        }

        if (-not $positionReached) {
            Write-Error "La camera n'a pas atteint la position du preset dans le delai imparti ($Timeout s)."
        }
    }
}
