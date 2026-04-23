<#
.SYNOPSIS
    Exporte un rapport Excel de tous les equipements (hardware) du VMS Milestone,
    avec un snapshot optionnel ancre dans la cellule de chaque camera.
    Les snapshots sont recuperes en parallele pour minimiser le temps d'attente.
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

    # --- Avertissement mots de passe ---
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

    # --- Option snapshots ---
    $snapAnswer = [System.Windows.MessageBox]::Show(
        "Voulez-vous inclure un snapshot de chaque camera ?`n`nLes snapshots sont recuperes en parallele pour aller plus vite.",
        'Options — Snapshots',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    $includeSnapshots = ($snapAnswer -eq [System.Windows.MessageBoxResult]::Yes)

    & $Log "Generation du rapport hardware$(if ($includeSnapshots) { ' avec snapshots' })..."
    $camReport = @(Get-VmsCameraReport -IncludePlainTextPassword)
    $total = $camReport.Count
    & $Log "$total equipements trouves."

    $vmsCamLookup = @{}
    if ($includeSnapshots) {
        & $Log "Chargement des objets camera..."
        Get-VmsCamera | ForEach-Object { $vmsCamLookup[$_.Name] = $_ }
    }

    if (-not (Test-Path $Config.outputDirectory)) {
        New-Item -Path $Config.outputDirectory -ItemType Directory -Force | Out-Null
    }

    $xlsxPath = Join-Path $Config.outputDirectory 'Liste_des_Cameras.xlsx'
    $tempDir  = Join-Path $env:TEMP "MilestoneHW_$(Get-Random)"
    if ($includeSnapshots) { New-Item $tempDir -ItemType Directory -Force | Out-Null }

    # ----------------------------------------------------------------
    # PHASE 1 : Recuperation parallele des snapshots
    # Chaque runspace recoit l'objet camera et appelle Get-Snapshot.
    # Les fichiers JPEG sont ecrits dans $tempDir ; le nom = cle du dict.
    # ----------------------------------------------------------------
    $snapPaths = @{}   # name -> chemin fichier JPEG (ou absent si echec)

    if ($includeSnapshots) {
        & $Log "Recuperation des snapshots en parallele (jusqu'a 6 simultanes)..."

        $quality = $Config.snapshotQuality

        # Script execute dans chaque runspace
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

        # Parallelisme adaptatif : 1 thread par camera, max 12
        $maxThreads = [Math]::Min($camReport.Count, 12)
        $pool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreads)
        $pool.ApartmentState = 'MTA'
        $pool.Open()

        $jobs = [System.Collections.Generic.List[hashtable]]::new()

        foreach ($cam in $camReport) {
            $vmsCamera = $vmsCamLookup[$cam.Name]
            if (-not $vmsCamera) { continue }

            $safeName = $cam.Name -replace '[\\/:*?"<>|]', '_'
            $filePath = Join-Path $tempDir "$safeName.jpg"

            $ps = [PowerShell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($snapScript).AddArgument($vmsCamera).AddArgument($quality).AddArgument($filePath)

            $jobs.Add(@{
                PS       = $ps
                Handle   = $ps.BeginInvoke()
                Name     = $cam.Name
            })
        }

        # Collecte des resultats des qu'ils sont prets (polling non-bloquant)
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
                    } else {
                        & $Log "  AVERTISSEMENT: Snapshot vide '$($job.Name)'"
                    }
                } catch {
                    & $Log "  AVERTISSEMENT: '$($job.Name)' : $_"
                } finally {
                    $job.PS.Dispose()
                }
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
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    }
    catch {
        & $Log "ERREUR: Excel n'est pas installe sur ce poste."
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return
    }

    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Add()
        $sheet    = $workbook.Sheets.Item(1)
        $sheet.Name = 'Cameras'

        $headers = if ($includeSnapshots) {
            @('Nom','Fabricant','Modele','IP','MAC','Firmware','ServeurRec','Utilisateur','MotDePasse','Snapshot')
        } else {
            @('Nom','Fabricant','Modele','IP','MAC','Firmware','ServeurRec','Utilisateur','MotDePasse')
        }
        $snapCol = $headers.Count

        for ($c = 0; $c -lt $headers.Count; $c++) {
            $cell = $sheet.Cells.Item(1, $c + 1)
            $cell.Value2              = $headers[$c]
            $cell.Font.Bold           = $true
            $cell.Font.Size           = 11
            $cell.Interior.Color      = 0x44413D
            $cell.Font.Color          = 0xF4D6CD
            $cell.HorizontalAlignment = -4108
        }

        $sheet.Rows.Item(2).Select() | Out-Null
        $sheet.Application.ActiveWindow.FreezePanes = $true

        $rowH = 90
        $row  = 2
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

            $sheet.Cells.Item($row, 1) = $cam.Name
            $sheet.Cells.Item($row, 2) = $cam.DriverFamily
            $sheet.Cells.Item($row, 3) = $cam.Model
            $sheet.Cells.Item($row, 4) = $cam.Address
            $sheet.Cells.Item($row, 5) = $cam.MAC
            $sheet.Cells.Item($row, 6) = $cam.Firmware
            $sheet.Cells.Item($row, 7) = $cam.RecorderName
            $sheet.Cells.Item($row, 8) = $cam.Username
            $sheet.Cells.Item($row, 9) = $cam.Password

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

                        $shape = $sheet.Shapes.AddPicture(
                            $snapFile,
                            [Microsoft.Office.Core.MsoTriState]::msoFalse,
                            [Microsoft.Office.Core.MsoTriState]::msoCTrue,
                            $left, $top, $width, $height
                        )
                        $shape.Placement = 1   # xlMoveAndSize
                    } catch {
                        & $Log "  AVERTISSEMENT: Insertion image '$($cam.Name)' : $_"
                    }
                }
            }

            $row++
        }

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
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
