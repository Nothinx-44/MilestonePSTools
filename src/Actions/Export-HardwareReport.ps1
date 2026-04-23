<#
.SYNOPSIS
    Exporte un rapport Excel de tous les equipements Milestone.
    Inclut : infos hardware, flux video (codec/resolution/FPS par stream), retention disponible, snapshot optionnel.
#>

function Export-HardwareReport {
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

    # Helper : lit une propriete dans Settings d'un stream (retourne N/A si absente)
    function Get-StreamSetting {
        param($stream, [string]$key)
        if ($stream -and $stream.Settings -and $stream.Settings[$key]) {
            return $stream.Settings[$key]
        }
        return 'N/A'
    }

    # ----------------------------------------------------------------
    # Dialogues de confirmation
    # ----------------------------------------------------------------
    $confirm = [System.Windows.MessageBox]::Show(
        "Ce rapport inclut les mots de passe des cameras en clair dans le fichier Excel.`n`nAssurez-vous de stocker le fichier dans un emplacement securise.`n`nContinuer ?",
        'Avertissement — Donnees sensibles',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
        & $Log "Export annule par l'utilisateur."
        return
    }

    $snapAnswer = [System.Windows.MessageBox]::Show(
        "Voulez-vous inclure un snapshot de chaque camera ?`n`nLes snapshots sont recuperes en parallele.",
        'Options — Snapshots',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    $includeSnapshots = ($snapAnswer -eq [System.Windows.MessageBoxResult]::Yes)

    # ----------------------------------------------------------------
    # Recuperation des donnees de base
    # ----------------------------------------------------------------
    & $Log "Generation du rapport hardware..."
    $camReport = @(Get-VmsCameraReport -IncludePlainTextPassword)
    $total     = $camReport.Count
    & $Log "$total equipements trouves."

    & $Log "Chargement des objets camera..."
    $vmsCameras  = @(Get-VmsCamera)
    $vmsCamByName = @{}
    $vmsCamByPath = @{}
    foreach ($c in $vmsCameras) {
        $vmsCamByName[$c.Name] = $c
        $vmsCamByPath[$c.Path] = $c.Name
    }

    # ----------------------------------------------------------------
    # PHASE 0a : Flux video (codec / resolution / FPS)
    # ----------------------------------------------------------------
    & $Log "Recuperation des configurations de flux video..."
    $streamLookup = @{}   # camName -> @{ Rec=$stream; Live=$stream }

    try {
        $allStreams = @($vmsCameras | Get-VmsCameraStream -Enabled -ErrorAction Stop)
        foreach ($s in $allStreams) {
            $name = $s.Camera.Name
            if (-not $streamLookup.ContainsKey($name)) {
                $streamLookup[$name] = @{ Rec = $null ; Live = $null ; Extra = 0 }
            }
            if ($s.Recorded -and -not $streamLookup[$name].Rec) {
                $streamLookup[$name].Rec = $s
            }
            elseif ($s.LiveDefault -and -not $streamLookup[$name].Live) {
                $streamLookup[$name].Live = $s
            }
            else {
                $streamLookup[$name].Extra++
            }
        }
        & $Log "$($allStreams.Count) flux trouves pour $($streamLookup.Count) cameras."
    }
    catch {
        & $Log "AVERTISSEMENT: Impossible de recuperer les flux video : $_"
    }

    # ----------------------------------------------------------------
    # PHASE 0b : Retention disponible (Get-PlaybackInfo)
    # ----------------------------------------------------------------
    & $Log "Recuperation des dates d'enregistrement..."
    $retentionLookup = @{}   # camName -> string

    try {
        $playbackData = @($vmsCameras | Get-PlaybackInfo -Parallel -ErrorAction Stop)
        foreach ($pb in $playbackData) {
            $name = $vmsCamByPath[$pb.Path]
            if ($name) {
                if ($pb.Begin -and $pb.End) {
                    $ts   = $pb.End - $pb.Begin
                    $days = [int]$ts.TotalDays
                    $retentionLookup[$name] = "$days jours"
                }
                else {
                    $retentionLookup[$name] = 'Aucun'
                }
            }
        }
        & $Log "Dates recuperees pour $($retentionLookup.Count) cameras."
    }
    catch {
        & $Log "AVERTISSEMENT: Impossible de recuperer les dates d'enregistrement : $_"
    }

    # ----------------------------------------------------------------
    # PHASE 1 : Snapshots en parallele (optionnel)
    # ----------------------------------------------------------------
    $snapPaths = @{}

    if ($includeSnapshots) {
        & $Log "Recuperation des snapshots en parallele..."
        $quality = $Config.snapshotQuality
        $tempDir = Join-Path $env:TEMP "MilestoneHW_$(Get-Random)"
        New-Item $tempDir -ItemType Directory -Force | Out-Null

        $snapScript = {
            param($camera, $quality, $filePath)
            try {
                $snap = $camera | Get-Snapshot -Behavior GetEnd -Quality $quality -ErrorAction Stop
                if ($snap -and $snap.Bytes -and $snap.Bytes.Length -gt 0) {
                    [System.IO.File]::WriteAllBytes($filePath, $snap.Bytes)
                    return $filePath
                }
            } catch { }
            return $null
        }

        $maxThreads = [Math]::Min($total, 12)
        $pool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreads)
        $pool.ApartmentState = 'MTA'
        $pool.Open()

        $jobs = [System.Collections.Generic.List[hashtable]]::new()

        foreach ($cam in $camReport) {
            $vmsCamera = $vmsCamByName[$cam.Name]
            if (-not $vmsCamera) { continue }
            $safeName = $cam.Name -replace '[\\/:*?"<>|]', '_'
            $filePath = Join-Path $tempDir "$safeName.jpg"
            $ps = [PowerShell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($snapScript).AddArgument($vmsCamera).AddArgument($quality).AddArgument($filePath)
            $jobs.Add(@{ PS = $ps; Handle = $ps.BeginInvoke(); Name = $cam.Name })
        }

        $pending  = [System.Collections.Generic.List[hashtable]]::new($jobs)
        $received = 0

        while ($pending.Count -gt 0) {
            $completed = @($pending | Where-Object { $_.Handle.IsCompleted })
            foreach ($job in $completed) {
                [void]$pending.Remove($job)
                try {
                    $result = $job.PS.EndInvoke($job.Handle)
                    if ($result) {
                        $snapPaths[$job.Name] = $result
                        $received++
                        & $Log "  [OK $received/$($jobs.Count)] $($job.Name)"
                    }
                    else { & $Log "  AVERTISSEMENT: Snapshot vide '$($job.Name)'" }
                }
                catch { & $Log "  AVERTISSEMENT: '$($job.Name)' : $_" }
                finally { $job.PS.Dispose() }
                & $ReportProgress ($jobs.Count - $pending.Count) $jobs.Count
            }
            if ($pending.Count -gt 0) { Start-Sleep -Milliseconds 150 }
        }

        $pool.Close()
        $pool.Dispose()
        & $Log "$($snapPaths.Count) / $($jobs.Count) snapshots recuperes."
    }

    # ----------------------------------------------------------------
    # PHASE 2 : Construction du fichier Excel
    # ----------------------------------------------------------------
    if (-not (Test-Path $Config.outputDirectory)) {
        New-Item -Path $Config.outputDirectory -ItemType Directory -Force | Out-Null
    }
    $xlsxPath = Join-Path $Config.outputDirectory 'Liste_des_Cameras.xlsx'

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    }
    catch {
        & $Log "ERREUR: Excel n'est pas installe sur ce poste."
        if ($includeSnapshots) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
        return
    }

    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Add()
        $sheet    = $workbook.Sheets.Item(1)
        $sheet.Name = 'Cameras'

        # Colonnes fixes + flux + retention + snapshot optionnel
        $baseHeaders  = @('Nom','Fabricant','Modele','IP','MAC','Firmware','ServeurRec','Utilisateur','MotDePasse')
        $streamHeaders = @('Codec (Enreg.)','Resolution (Enreg.)','FPS (Enreg.)','Codec (Live)','Resolution (Live)','FPS (Live)','Flux supplementaires')
        $retHeader    = @('Retention disponible')
        $snapHeader   = if ($includeSnapshots) { @('Snapshot') } else { @() }

        $headers = $baseHeaders + $streamHeaders + $retHeader + $snapHeader
        $snapCol = $headers.Count   # derniere colonne si snapshots

        # En-tetes avec couleurs par groupe
        $groupColors = @{
            'base'    = @{ Bg = 0x44413D ; Fg = 0xF4D6CD }
            'stream'  = @{ Bg = 0x1D3557 ; Fg = 0xA8DADC }
            'ret'     = @{ Bg = 0x1B4332 ; Fg = 0xA6E3A1 }
            'snap'    = @{ Bg = 0x2D2B55 ; Fg = 0xCBA6F7 }
        }

        for ($c = 0; $c -lt $headers.Count; $c++) {
            $cell = $sheet.Cells.Item(1, $c + 1)
            $cell.Value2    = $headers[$c]
            $cell.Font.Bold = $true
            $cell.Font.Size = 11
            $cell.HorizontalAlignment = -4108

            $grp = if ($c -lt $baseHeaders.Count)                             { 'base'   }
                   elseif ($c -lt $baseHeaders.Count + $streamHeaders.Count)  { 'stream' }
                   elseif ($c -lt $baseHeaders.Count + $streamHeaders.Count + $retHeader.Count) { 'ret' }
                   else                                                         { 'snap'   }

            $cell.Interior.Color = $groupColors[$grp].Bg
            $cell.Font.Color     = $groupColors[$grp].Fg
        }

        $sheet.Rows.Item(2).Select() | Out-Null
        $sheet.Application.ActiveWindow.FreezePanes = $true

        $rowH  = 90
        $row   = 2
        $count = 0

        & $Log "Construction du fichier Excel..."

        foreach ($cam in $camReport) {
            if (& $Cancel) {
                & $Log "AVERTISSEMENT: Operation annulee apres $count / $total cameras."
                break
            }

            $count++
            & $ReportProgress $count $total
            & $Log "[$count/$total] $($cam.Name)"

            # Donnees hardware
            $sheet.Cells.Item($row, 1) = $cam.Name
            $sheet.Cells.Item($row, 2) = $cam.DriverFamily
            $sheet.Cells.Item($row, 3) = $cam.Model
            $ip = if ($cam.Address -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})') { $Matches[1] } else { $cam.Address }
            $sheet.Cells.Item($row, 4) = $ip
            $sheet.Cells.Item($row, 5) = $cam.MAC
            $sheet.Cells.Item($row, 6) = $cam.Firmware
            $sheet.Cells.Item($row, 7) = $cam.RecorderName
            $sheet.Cells.Item($row, 8) = $cam.Username
            $sheet.Cells.Item($row, 9) = $cam.Password

            # Flux video
            $si = $streamLookup[$cam.Name]
            $recStream  = if ($si) { $si.Rec  } else { $null }
            $liveStream = if ($si) { $si.Live } else { $null }
            $extraCount = if ($si) { $si.Extra } else { 0 }

            # Si meme stream pour live et enregistrement, pas de colonne Live separee
            $sameStream = $recStream -and $liveStream -and ($recStream.Name -eq $liveStream.Name)

            $sheet.Cells.Item($row, 10) = Get-StreamSetting $recStream 'Codec'
            $sheet.Cells.Item($row, 11) = Get-StreamSetting $recStream 'Resolution'
            $sheet.Cells.Item($row, 12) = Get-StreamSetting $recStream 'FPS'

            if ($sameStream) {
                $sheet.Cells.Item($row, 13) = ''
                $sheet.Cells.Item($row, 14) = ''
                $sheet.Cells.Item($row, 15) = ''
            }
            else {
                $sheet.Cells.Item($row, 13) = Get-StreamSetting $liveStream 'Codec'
                $sheet.Cells.Item($row, 14) = Get-StreamSetting $liveStream 'Resolution'
                $sheet.Cells.Item($row, 15) = Get-StreamSetting $liveStream 'FPS'
            }

            $sheet.Cells.Item($row, 16) = if ($extraCount -gt 0) { "$extraCount flux supp." } else { '' }

            # Retention
            $ret = if ($retentionLookup.ContainsKey($cam.Name)) { $retentionLookup[$cam.Name] } else { 'N/A' }
            $sheet.Cells.Item($row, 17) = $ret

            # Snapshot
            if ($includeSnapshots) {
                $sheet.Rows.Item($row).RowHeight = $rowH
                $snapFile = $snapPaths[$cam.Name]
                if ($snapFile -and (Test-Path $snapFile)) {
                    try {
                        $cell   = $sheet.Cells.Item($row, $snapCol)
                        $left   = [double]$cell.Left
                        $top    = [double]$cell.Top
                        $width  = [double]$cell.Width
                        $height = [double]$cell.Height
                        $shape  = $sheet.Shapes.AddPicture(
                            $snapFile,
                            [Microsoft.Office.Core.MsoTriState]::msoFalse,
                            [Microsoft.Office.Core.MsoTriState]::msoCTrue,
                            $left, $top, $width, $height
                        )
                        $shape.Placement = 1
                    }
                    catch { & $Log "  AVERTISSEMENT: Image '$($cam.Name)' : $_" }
                }
            }

            $row++
        }

        # Mise en forme finale
        if ($includeSnapshots) { $sheet.Columns.Item($snapCol).ColumnWidth = 28 }
        for ($c = 1; $c -lt $snapCol; $c++) { $sheet.Columns.Item($c).AutoFit() | Out-Null }

        $range = $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item($row-1, $snapCol))
        $range.Borders.LineStyle = 1
        $range.Borders.Weight    = 2

        $workbook.SaveAs($xlsxPath, 51)
        & $Log "Rapport exporte : $xlsxPath"
    }
    finally {
        try { $workbook.Close($false) } catch {}
        try { $excel.Quit() }           catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        if ($includeSnapshots -and (Test-Path $tempDir)) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
