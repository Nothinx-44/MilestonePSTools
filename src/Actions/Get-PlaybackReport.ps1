<#
.SYNOPSIS
    Premier et dernier enregistrement disponible par camera (Get-PlaybackInfo).
#>

function Get-PlaybackReport {
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

    & $Log "Recuperation des cameras..."
    $cameras = @(Get-VmsCamera)
    $total   = $cameras.Count
    & $Log "$total cameras trouvees. Recuperation des plages d'enregistrement..."

    # Table de correspondance Path -> Nom
    $camByPath = @{}
    foreach ($cam in $cameras) { $camByPath[$cam.Path] = $cam.Name }

    try {
        # -Parallel active automatiquement les runspaces internes pour >= 60 cameras
        $playbackData = @($cameras | Get-PlaybackInfo -Parallel -ErrorAction Stop)
    }
    catch {
        & $Log "ERREUR: Get-PlaybackInfo : $_"
        return
    }

    $rows  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $count = 0

    foreach ($info in $playbackData) {
        if (& $Cancel) { & $Log "AVERTISSEMENT: Operation annulee apres $count / $total." ; break }

        $count++
        & $ReportProgress $count $total

        $name = if ($camByPath.ContainsKey($info.Path)) { $camByPath[$info.Path] } else { $info.Path }

        if ($info.Begin -and $info.End) {
            $begin    = $info.Begin.ToLocalTime()
            $end      = $info.End.ToLocalTime()
            $beginStr = $begin.ToString('dd/MM/yyyy HH:mm')
            $endStr   = $end.ToString('dd/MM/yyyy HH:mm')
            $ts       = $end - $begin
            $durStr   = if ([int]$ts.TotalDays -gt 0) { "$([int]$ts.TotalDays)j $($ts.Hours)h" }
                        else                           { "$($ts.Hours)h $($ts.Minutes)m" }

            & $Log "  $name : $beginStr -> $endStr ($durStr)"
        }
        else {
            $beginStr = 'N/A'
            $endStr   = 'N/A'
            $durStr   = 'N/A'
            & $Log "  AVERTISSEMENT: $name : aucun enregistrement trouve"
        }

        $rows.Add([PSCustomObject]@{
            Nom             = $name
            PremierEnreg    = $beginStr
            DernierEnreg    = $endStr
            DureeDisponible = $durStr
        })
    }

    if ($rows.Count -gt 0 -and -not (& $Cancel)) {
        if (-not (Test-Path $Config.outputDirectory)) {
            New-Item -Path $Config.outputDirectory -ItemType Directory -Force | Out-Null
        }
        $csvPath = Join-Path $Config.outputDirectory 'Dates_Enregistrement.csv'
        $rows | Export-Csv -Path $csvPath -NoTypeInformation `
            -Encoding $Config.csvEncoding -Delimiter $Config.csvDelimiter
        & $Log "Rapport exporte : $csvPath"
    }
}
