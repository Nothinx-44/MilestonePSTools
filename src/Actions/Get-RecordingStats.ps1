<#
.SYNOPSIS
    Statistiques d'enregistrement et de flux video par camera Milestone (7 derniers jours).
#>

function Get-RecordingStats {
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

    $endTime   = (Get-Date).Date
    $startTime = $endTime.AddDays(-7)
    & $Log "Periode analysee : $($startTime.ToString('dd/MM/yyyy')) -> $($endTime.ToString('dd/MM/yyyy'))"

    & $Log "Recuperation des cameras..."
    $cameras = @(Get-VmsCamera)
    $total   = $cameras.Count
    & $Log "$total cameras trouvees."

    $rows  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $count = 0

    foreach ($cam in $cameras) {
        if (& $Cancel) { & $Log "AVERTISSEMENT: Operation annulee apres $count / $total cameras." ; break }

        $count++
        & $ReportProgress $count $total
        & $Log "[$count/$total] $($cam.Name)"

        $row = [PSCustomObject]@{
            Nom              = $cam.Name
            Active           = $cam.Enabled
            Sequences_Rec    = 'N/A'
            PctTemps_Rec     = 'N/A'
            TempsEnregistre  = 'N/A'
            Sequences_Motion = 'N/A'
            PctTemps_Motion  = 'N/A'
            FPS_Live         = 'N/A'
            Bitrate_Live_kbs = 'N/A'
            Resolution       = 'N/A'
        }

        # Stats enregistrement (RecordingSequence)
        try {
            $recResult = $cam | Get-CameraRecordingStats `
                -StartTime $startTime -EndTime $endTime `
                -SequenceType RecordingSequence -ErrorAction Stop

            if ($recResult -and $recResult.RecordingStats) {
                $rs  = $recResult.RecordingStats
                $pct = if ($null -ne $rs.PercentRecorded) { "$([math]::Round($rs.PercentRecorded, 1)) %" } else { 'N/A' }
                $dur = if ($null -ne $rs.TimeRecorded) {
                    $ts = [TimeSpan]$rs.TimeRecorded
                    if ([int]$ts.TotalDays -gt 0) { "$([int]$ts.TotalDays)j $($ts.Hours)h $($ts.Minutes)m" }
                    else                          { "$($ts.Hours)h $($ts.Minutes)m" }
                } else { 'N/A' }

                $row.Sequences_Rec  = $rs.SequenceCount
                $row.PctTemps_Rec   = $pct
                $row.TempsEnregistre = $dur
                & $Log "  Enreg. : $($rs.SequenceCount) sequences | $pct | $dur"
            }
        }
        catch {
            & $Log "  AVERTISSEMENT: Stats enregistrement : $_"
        }

        # Stats mouvement (MotionSequence)
        try {
            $motResult = $cam | Get-CameraRecordingStats `
                -StartTime $startTime -EndTime $endTime `
                -SequenceType MotionSequence -ErrorAction Stop

            if ($motResult -and $motResult.RecordingStats) {
                $ms     = $motResult.RecordingStats
                $motPct = if ($null -ne $ms.PercentRecorded) { "$([math]::Round($ms.PercentRecorded, 1)) %" } else { 'N/A' }

                $row.Sequences_Motion = $ms.SequenceCount
                $row.PctTemps_Motion  = $motPct
                & $Log "  Motion : $($ms.SequenceCount) sequences | $motPct"
            }
        }
        catch {
            & $Log "  AVERTISSEMENT: Stats mouvement : $_"
        }

        # Statistiques live du flux video
        try {
            $vidStats = $cam | Get-VideoDeviceStatistics -ErrorAction Stop
            if ($vidStats) {
                $fps = if ($null -ne $vidStats.FPS)             { [math]::Round($vidStats.FPS, 1)        } else { 'N/A' }
                $bps = if ($null -ne $vidStats.BPS)             { [math]::Round($vidStats.BPS / 1000, 0) } else { 'N/A' }
                $res = if ($null -ne $vidStats.ImageResolution) { $vidStats.ImageResolution              } else { 'N/A' }

                $row.FPS_Live         = $fps
                $row.Bitrate_Live_kbs = $bps
                $row.Resolution       = $res
                & $Log "  Live   : $fps FPS | $bps kbps | $res"
            }
        }
        catch {}

        $rows.Add($row)
    }

    if ($rows.Count -gt 0 -and -not (& $Cancel)) {
        if (-not (Test-Path $Config.outputDirectory)) {
            New-Item -Path $Config.outputDirectory -ItemType Directory -Force | Out-Null
        }
        $csvPath = Join-Path $Config.outputDirectory 'Stats_Enregistrement.csv'
        $rows | Export-Csv -Path $csvPath -NoTypeInformation `
            -Encoding $Config.csvEncoding -Delimiter $Config.csvDelimiter
        & $Log "Rapport exporte : $csvPath"
    }
}
