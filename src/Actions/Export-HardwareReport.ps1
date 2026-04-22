<#
.SYNOPSIS
    Exporte un rapport Excel de tous les equipements (hardware) du VMS Milestone,
    avec un snapshot optionnel ancre dans la cellule de chaque camera.
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
        "Voulez-vous inclure un snapshot de chaque camera ?`n`nCela peut prendre plusieurs minutes selon le nombre de cameras.",
        'Options — Snapshots',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    $includeSnapshots = ($snapAnswer -eq [System.Windows.MessageBoxResult]::Yes)

    & $Log "Generation du rapport hardware$(if ($includeSnapshots) { ' avec snapshots' })..."
    $camReport = @(Get-VmsCameraReport -IncludePlainTextPassword)
    $total = $camReport.Count
    & $Log "$total equipements trouves."

    # Lookup VmsCamera par nom (uniquement si snapshots demandes)
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
    if ($includeSnapshots) {
        New-Item $tempDir -ItemType Directory -Force | Out-Null
    }

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    }
    catch {
        & $Log "ERREUR: Excel n'est pas installe sur ce poste."
        return
    }

    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Add()
        $sheet    = $workbook.Sheets.Item(1)
        $sheet.Name = 'Cameras'

        # ---- En-tetes ----
        $headers = if ($includeSnapshots) {
            @('Nom','Fabricant','Modele','IP','MAC','Firmware','ServeurRec','Utilisateur','MotDePasse','Snapshot')
        } else {
            @('Nom','Fabricant','Modele','IP','MAC','Firmware','ServeurRec','Utilisateur','MotDePasse')
        }
        $snapCol = $headers.Count  # derniere colonne

        for ($c = 0; $c -lt $headers.Count; $c++) {
            $cell = $sheet.Cells.Item(1, $c + 1)
            $cell.Value2              = $headers[$c]
            $cell.Font.Bold           = $true
            $cell.Font.Size           = 11
            $cell.Interior.Color      = 0x44413D
            $cell.Font.Color          = 0xF4D6CD
            $cell.HorizontalAlignment = -4108   # xlCenter
        }

        # Figer la ligne d'en-tete
        $sheet.Rows.Item(2).Select() | Out-Null
        $sheet.Application.ActiveWindow.FreezePanes = $true

        $rowH = 90   # hauteur de ligne en points pour les images
        $row  = 2
        $count = 0

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
                # Fixer la hauteur de ligne avant de lire les dimensions de la cellule
                $sheet.Rows.Item($row).RowHeight = $rowH

                try {
                    $vmsCamera = $vmsCamLookup[$cam.Name]
                    if ($vmsCamera) {
                        $snap = $vmsCamera | Get-Snapshot -Behavior GetEnd -Quality $Config.snapshotQuality -ErrorAction Stop
                        if ($snap -and $snap.Bytes -and $snap.Bytes.Length -gt 0) {
                            $safeName = $cam.Name -replace '[\\/:*?"<>|]', '_'
                            $snapPath = Join-Path $tempDir "$safeName.jpg"
                            [System.IO.File]::WriteAllBytes($snapPath, $snap.Bytes)

                            $cell   = $sheet.Cells.Item($row, $snapCol)
                            $left   = [double]$cell.Left
                            $top    = [double]$cell.Top
                            $width  = [double]$cell.Width
                            $height = [double]$cell.Height

                            $shape = $sheet.Shapes.AddPicture(
                                $snapPath,
                                [Microsoft.Office.Core.MsoTriState]::msoFalse,
                                [Microsoft.Office.Core.MsoTriState]::msoCTrue,
                                $left, $top, $width, $height
                            )
                            $shape.Placement = 1  # xlMoveAndSize — ancre dans la cellule
                        }
                    }
                }
                catch {
                    & $Log "  AVERTISSEMENT: Snapshot '$($cam.Name)' : $_"
                }
            }

            $row++
        }

        # ---- Mise en forme finale ----
        if ($includeSnapshots) {
            $sheet.Columns.Item($snapCol).ColumnWidth = 28
        }

        for ($c = 1; $c -lt $snapCol; $c++) {
            $sheet.Columns.Item($c).AutoFit() | Out-Null
        }

        $range = $sheet.Range(
            $sheet.Cells.Item(1, 1),
            $sheet.Cells.Item($row - 1, $snapCol)
        )
        $range.Borders.LineStyle = 1
        $range.Borders.Weight    = 2

        $workbook.SaveAs($xlsxPath, 51)
        & $Log "Rapport exporte : $xlsxPath"
    }
    finally {
        try { $workbook.Close($false) } catch {}
        try { $excel.Quit() }           catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        if ($includeSnapshots) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
