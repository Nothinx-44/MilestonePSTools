<#
.SYNOPSIS
    Capture un snapshot de toutes les cameras du VMS Milestone (en parallele).
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
    $cameras = @(Get-VmsCamera)
    $total   = $cameras.Count
    & $Log "$total cameras trouvees. Capture en parallele..."

    if ($SnapshotTime) {
        & $Log "Mode historique : $($SnapshotTime.ToString('dd/MM/yyyy HH:mm'))"
    }

    $behavior    = if ($SnapshotTime) { 'GetNearest' } else { 'GetEnd' }
    $quality     = $Config.snapshotQuality
    $useTime     = [bool]$SnapshotTime
    $snapTime    = $SnapshotTime   # valeur concrete (ou $null)

    $snapScript = {
        param($camera, $behavior, $useTime, $snapTime, $quality, $dir)
        try {
            if ($useTime) {
                $camera | Get-Snapshot `
                    -UseFriendlyName -Behavior $behavior `
                    -Time $snapTime -Quality $quality `
                    -Save -Path $dir
            } else {
                $camera | Get-Snapshot `
                    -UseFriendlyName -Behavior $behavior `
                    -Quality $quality `
                    -Save -Path $dir
            }
            return $camera.Name
        } catch {
            return $null
        }
    }

    $maxThreads = [Math]::Min($total, 12)
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreads)
    $pool.ApartmentState = 'MTA'
    $pool.Open()

    $jobs = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($cam in $cameras) {
        if (& $Cancel) {
            & $Log "AVERTISSEMENT: Operation annulee avant lancement."
            break
        }

        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($snapScript).AddArgument($cam).AddArgument($behavior).AddArgument($useTime).AddArgument($snapTime).AddArgument($quality).AddArgument($snapshotDir)

        $jobs.Add(@{ PS = $ps; Handle = $ps.BeginInvoke(); Name = $cam.Name })
    }

    # Collecte des resultats des qu'ils sont prets
    $pending  = [System.Collections.Generic.List[hashtable]]::new($jobs)
    $received = 0
    $errors   = 0

    while ($pending.Count -gt 0) {
        $completed = @($pending | Where-Object { $_.Handle.IsCompleted })

        foreach ($job in $completed) {
            [void]$pending.Remove($job)
            try {
                $result = $job.PS.EndInvoke($job.Handle)
                if ($result) {
                    $received++
                    & $Log "  [OK $received/$total] $($job.Name)"
                } else {
                    $errors++
                    & $Log "  AVERTISSEMENT: Echec snapshot '$($job.Name)'"
                }
            } catch {
                $errors++
                & $Log "  ERREUR '$($job.Name)': $_"
            } finally {
                $job.PS.Dispose()
            }
            & $ReportProgress ($total - $pending.Count) $total
        }

        if ($pending.Count -gt 0) { Start-Sleep -Milliseconds 150 }
    }

    $pool.Close()
    $pool.Dispose()

    & $Log "$received snapshots enregistres$(if ($errors -gt 0) { ", $errors echecs" }) dans : $snapshotDir"
}
