function Get-SnapshotAll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$Config,
        [Parameter(Mandatory)] [scriptblock]$Log,
        [Parameter()] [scriptblock]$Cancel = { $false },
        [Parameter()] [scriptblock]$ReportProgress = {},
        [Parameter()] [nullable[datetime]]$SnapshotTime = $null
    )

    $snapshotDir = Join-Path $Config.outputDirectory 'Snapshots'
    if (-not (Test-Path $snapshotDir)) {
        New-Item -Path $snapshotDir -ItemType Directory -Force | Out-Null
    }

    & $Log $script:T.SA_LogCams
    $cameras = @(Get-VmsCamera)
    $total   = $cameras.Count
    & $Log ($script:T.SA_LogFound -f $total)

    if ($SnapshotTime) { & $Log ($script:T.SA_LogHistorique -f $SnapshotTime.ToString('dd/MM/yyyy HH:mm')) }

    $behavior = if ($SnapshotTime) { 'GetNearest' } else { 'GetEnd' }
    $quality  = $Config.snapshotQuality
    $useTime  = [bool]$SnapshotTime
    $snapTime = $SnapshotTime

    $snapScript = {
        param($camera, $behavior, $useTime, $snapTime, $quality, $dir)
        try {
            if ($useTime) {
                $camera | Get-Snapshot -UseFriendlyName -Behavior $behavior `
                    -Time $snapTime -Quality $quality -Save -Path $dir
            } else {
                $camera | Get-Snapshot -UseFriendlyName -Behavior $behavior `
                    -Quality $quality -Save -Path $dir
            }
            return $camera.Name
        } catch { return $null }
    }

    $maxThreads = [Math]::Min($total, 12)
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreads)
    $pool.ApartmentState = 'MTA'
    $pool.Open()

    try {
        $jobs = [System.Collections.Generic.List[hashtable]]::new()

        foreach ($cam in $cameras) {
            if (& $Cancel) { & $Log $script:T.SA_LogCancelled ; break }
            $ps = [PowerShell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($snapScript).AddArgument($cam).AddArgument($behavior).AddArgument($useTime).AddArgument($snapTime).AddArgument($quality).AddArgument($snapshotDir)
            $jobs.Add(@{ PS = $ps; Handle = $ps.BeginInvoke(); Name = $cam.Name })
        }

        $pending  = [System.Collections.Generic.List[hashtable]]::new($jobs)
        $received = 0
        $errors   = 0
        $timeout  = [datetime]::UtcNow.AddMinutes(10)

        while ($pending.Count -gt 0) {
            # Timeout de securite : evite un blocage infini si une camera ne repond jamais
            if ([datetime]::UtcNow -gt $timeout) {
                foreach ($job in @($pending)) {
                    try { $job.PS.Stop() } catch {}
                    $job.PS.Dispose()
                    $errors++
                }
                $pending.Clear()
                & $Log $script:T.SA_LogTimeout
                break
            }

            $completed = @($pending | Where-Object { $_.Handle.IsCompleted })
            foreach ($job in $completed) {
                [void]$pending.Remove($job)
                try {
                    $result = $job.PS.EndInvoke($job.Handle)
                    if ($result) {
                        $received++
                        & $Log ($script:T.SA_LogOk -f $received, $total, $job.Name)
                    } else {
                        $errors++
                        & $Log ($script:T.SA_LogFailed -f $job.Name)
                    }
                } catch {
                    $errors++
                    & $Log ($script:T.SA_LogError -f $job.Name, $_)
                } finally { $job.PS.Dispose() }
                & $ReportProgress ($total - $pending.Count) $total
            }
            if ($pending.Count -gt 0) { Start-Sleep -Milliseconds 150 }
        }
    }
    finally {
        $pool.Close()
        $pool.Dispose()
    }

    $msg = if ($errors -gt 0) { $script:T.SA_LogDoneErr -f $received, $errors, $snapshotDir }
           else                { $script:T.SA_LogDone    -f $received, $snapshotDir }
    & $Log $msg
}
