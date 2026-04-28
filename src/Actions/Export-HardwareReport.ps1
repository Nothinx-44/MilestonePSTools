<#
.SYNOPSIS
    Exporte un rapport Excel de tous les equipements Milestone.
    Inclut : infos hardware, flux video (codec/resolution/FPS par stream), retention disponible, snapshot optionnel.
    L'utilisateur selectionne les colonnes a inclure via une fenetre de choix.
    Les mots de passe n'apparaissent dans l'export que si la colonne est explicitement cochee.
#>

function Show-ExportColumnSelector {
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Options d'export"
        Width="530" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#1E1E2E" FontFamily="Segoe UI">
    <StackPanel Margin="20,16,20,20">

        <!-- Titre + boutons de selection rapide -->
        <DockPanel Margin="0,0,0,10">
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                <Button x:Name="BtnSelectAll" Content="Tout cocher"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="11" Padding="10,4" Cursor="Hand" Margin="0,0,6,0"/>
                <Button x:Name="BtnDeselectAll" Content="Tout decocher"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="11" Padding="10,4" Cursor="Hand"/>
            </StackPanel>
            <TextBlock Text="Selectionnez les colonnes a inclure dans l'export :"
                       Foreground="#CDD6F4" FontSize="13" VerticalAlignment="Center"/>
        </DockPanel>

        <!-- Informations hardware -->
        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="Informations hardware"
                           Foreground="#F4D6CD" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <UniformGrid Columns="3">
                    <CheckBox x:Name="ChkNom"         Content="Nom"              IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFabricant"   Content="Fabricant"        IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkModele"      Content="Modele"           IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkIP"          Content="IP"               IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkMAC"         Content="MAC"              IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFirmware"    Content="Firmware"         IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkServeurRec"  Content="Serveur Enreg."   IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkUtilisateur" Content="Utilisateur"      IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkMotDePasse"  Content="Mot de passe (!)" IsChecked="False"
                              Foreground="#FAB387" FontSize="12" Margin="4,5,4,5"/>
                </UniformGrid>
            </StackPanel>
        </Border>

        <!-- Flux video -->
        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="Flux video"
                           Foreground="#A8DADC" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <UniformGrid Columns="3">
                    <CheckBox x:Name="ChkCodecEnreg"      Content="Codec (Enreg.)"      IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkResolutionEnreg" Content="Resolution (Enreg.)" IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFPSEnreg"        Content="FPS (Enreg.)"        IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkCodecLive"       Content="Codec (Live)"         IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkResolutionLive"  Content="Resolution (Live)"    IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFPSLive"         Content="FPS (Live)"           IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFluxSupp"        Content="Flux supplementaires" IsChecked="True"
                              Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                </UniformGrid>
            </StackPanel>
        </Border>

        <!-- Retention -->
        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="Retention"
                           Foreground="#A6E3A1" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <CheckBox x:Name="ChkRetention" Content="Retention disponible" IsChecked="True"
                          Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
            </StackPanel>
        </Border>

        <!-- Options -->
        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="Options"
                           Foreground="#CBA6F7" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <CheckBox x:Name="ChkSnapshot" Content="Snapshot (lent - 1 image par camera)" IsChecked="False"
                          Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
            </StackPanel>
        </Border>

        <!-- Boutons -->
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="BtnCancel" Content="Annuler" IsCancel="True"
                    Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                    FontSize="13" Padding="22,8" Cursor="Hand" Margin="0,0,10,0"/>
            <Button x:Name="BtnExport" Content="Exporter" IsDefault="True"
                    Background="#A6E3A1" Foreground="#1E1E2E" BorderThickness="0"
                    FontWeight="Bold" FontSize="13" Padding="22,8" Cursor="Hand"/>
        </StackPanel>

    </StackPanel>
</Window>
'@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    # Capture des references checkbox AVANT l'enregistrement des handlers (evite les problemes de scope)
    $checkboxes = [ordered]@{
        'Nom'             = $window.FindName('ChkNom')
        'Fabricant'       = $window.FindName('ChkFabricant')
        'Modele'          = $window.FindName('ChkModele')
        'IP'              = $window.FindName('ChkIP')
        'MAC'             = $window.FindName('ChkMAC')
        'Firmware'        = $window.FindName('ChkFirmware')
        'ServeurRec'      = $window.FindName('ChkServeurRec')
        'Utilisateur'     = $window.FindName('ChkUtilisateur')
        'MotDePasse'      = $window.FindName('ChkMotDePasse')
        'CodecEnreg'      = $window.FindName('ChkCodecEnreg')
        'ResolutionEnreg' = $window.FindName('ChkResolutionEnreg')
        'FPSEnreg'        = $window.FindName('ChkFPSEnreg')
        'CodecLive'       = $window.FindName('ChkCodecLive')
        'ResolutionLive'  = $window.FindName('ChkResolutionLive')
        'FPSLive'         = $window.FindName('ChkFPSLive')
        'FluxSupp'        = $window.FindName('ChkFluxSupp')
        'Retention'       = $window.FindName('ChkRetention')
        'Snapshot'        = $window.FindName('ChkSnapshot')
    }

    $window.FindName('BtnSelectAll').Add_Click({
        foreach ($chk in $checkboxes.Values) { $chk.IsChecked = $true }
    })

    $window.FindName('BtnDeselectAll').Add_Click({
        foreach ($chk in $checkboxes.Values) { $chk.IsChecked = $false }
    })

    $window.FindName('BtnExport').Add_Click({
        $sel = [System.Collections.Generic.List[string]]::new()
        foreach ($name in $checkboxes.Keys) {
            if ($checkboxes[$name].IsChecked -eq $true) { $sel.Add($name) }
        }
        if ($sel.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                'Veuillez selectionner au moins une colonne.',
                'Aucune colonne selectionnee',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
            return
        }
        $window.Tag = [string[]]$sel
        $window.DialogResult = $true
    })

    if ($window.ShowDialog() -eq $true) { return [string[]]$window.Tag }
    return $null
}


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

    function Get-StreamSetting {
        param($stream, [string]$key)
        if ($stream -and $stream.Settings -and $stream.Settings[$key]) { return $stream.Settings[$key] }
        return 'N/A'
    }

    # ----------------------------------------------------------------
    # Fenetre de selection des colonnes
    # ----------------------------------------------------------------
    $selectedColumns = Show-ExportColumnSelector
    if ($null -eq $selectedColumns) {
        & $Log "Export annule par l'utilisateur."
        return
    }

    $includePassword  = $selectedColumns -contains 'MotDePasse'
    $includeSnapshots = $selectedColumns -contains 'Snapshot'
    $needStreams       = ($selectedColumns | Where-Object {
        $_ -in 'CodecEnreg','ResolutionEnreg','FPSEnreg','CodecLive','ResolutionLive','FPSLive','FluxSupp'
    }).Count -gt 0
    $needRetention    = $selectedColumns -contains 'Retention'

    # ----------------------------------------------------------------
    # Definition de toutes les colonnes possibles (ordre canonique)
    # ----------------------------------------------------------------
    $allColumnDefs = @(
        @{ Name = 'Nom';             Group = 'base';   Header = 'Nom' }
        @{ Name = 'Fabricant';       Group = 'base';   Header = 'Fabricant' }
        @{ Name = 'Modele';          Group = 'base';   Header = 'Modele' }
        @{ Name = 'IP';              Group = 'base';   Header = 'IP' }
        @{ Name = 'MAC';             Group = 'base';   Header = 'MAC' }
        @{ Name = 'Firmware';        Group = 'base';   Header = 'Firmware' }
        @{ Name = 'ServeurRec';      Group = 'base';   Header = 'ServeurRec' }
        @{ Name = 'Utilisateur';     Group = 'base';   Header = 'Utilisateur' }
        @{ Name = 'MotDePasse';      Group = 'base';   Header = 'MotDePasse' }
        @{ Name = 'CodecEnreg';      Group = 'stream'; Header = 'Codec (Enreg.)' }
        @{ Name = 'ResolutionEnreg'; Group = 'stream'; Header = 'Resolution (Enreg.)' }
        @{ Name = 'FPSEnreg';        Group = 'stream'; Header = 'FPS (Enreg.)' }
        @{ Name = 'CodecLive';       Group = 'stream'; Header = 'Codec (Live)' }
        @{ Name = 'ResolutionLive';  Group = 'stream'; Header = 'Resolution (Live)' }
        @{ Name = 'FPSLive';         Group = 'stream'; Header = 'FPS (Live)' }
        @{ Name = 'FluxSupp';        Group = 'stream'; Header = 'Flux supplementaires' }
        @{ Name = 'Retention';       Group = 'ret';    Header = 'Retention disponible' }
        @{ Name = 'Snapshot';        Group = 'snap';   Header = 'Snapshot' }
    )

    # Filtrer en preservant l'ordre canonique
    $activeColumns = @($allColumnDefs | Where-Object { $selectedColumns -contains $_.Name })

    # Index (1-based) de la colonne Snapshot, 0 si absente
    $snapColIndex = 0
    for ($i = 0; $i -lt $activeColumns.Count; $i++) {
        if ($activeColumns[$i].Name -eq 'Snapshot') { $snapColIndex = $i + 1; break }
    }

    # ----------------------------------------------------------------
    # Recuperation des donnees de base
    # ----------------------------------------------------------------
    & $Log "Generation du rapport hardware..."
    if ($includePassword) {
        $camReport = @(Get-VmsCameraReport -IncludePlainTextPassword)
    } else {
        $camReport = @(Get-VmsCameraReport)
    }
    $total = $camReport.Count
    & $Log "$total equipements trouves."

    & $Log "Chargement des objets camera..."
    $vmsCameras   = @(Get-VmsCamera)
    $vmsCamByName = @{}
    $vmsCamByPath = @{}
    foreach ($c in $vmsCameras) {
        $vmsCamByName[$c.Name] = $c
        $vmsCamByPath[$c.Path] = $c.Name
    }

    # ----------------------------------------------------------------
    # PHASE 0a : Flux video (uniquement si colonnes selectionnees)
    # ----------------------------------------------------------------
    $streamLookup = @{}
    if ($needStreams) {
        & $Log "Recuperation des configurations de flux video..."
        try {
            $allStreams = @($vmsCameras | Get-VmsCameraStream -Enabled -ErrorAction Stop)
            foreach ($s in $allStreams) {
                $name = $s.Camera.Name
                if (-not $streamLookup.ContainsKey($name)) {
                    $streamLookup[$name] = @{ Rec = $null; Live = $null; Extra = 0 }
                }
                if ($s.Recorded -and -not $streamLookup[$name].Rec)        { $streamLookup[$name].Rec = $s }
                elseif ($s.LiveDefault -and -not $streamLookup[$name].Live) { $streamLookup[$name].Live = $s }
                else                                                          { $streamLookup[$name].Extra++ }
            }
            & $Log "$($allStreams.Count) flux trouves pour $($streamLookup.Count) cameras."
        }
        catch { & $Log "AVERTISSEMENT: Impossible de recuperer les flux video : $_" }
    }

    # ----------------------------------------------------------------
    # PHASE 0b : Retention disponible (uniquement si colonne selectionnee)
    # ----------------------------------------------------------------
    $retentionLookup = @{}
    if ($needRetention) {
        & $Log "Recuperation des dates d'enregistrement..."
        try {
            $playbackData = @($vmsCameras | Get-PlaybackInfo -Parallel -ErrorAction Stop)
            foreach ($pb in $playbackData) {
                $name = $vmsCamByPath[$pb.Path]
                if ($name) {
                    if ($pb.Begin -and $pb.End) {
                        $days = [int]($pb.End - $pb.Begin).TotalDays
                        $retentionLookup[$name] = "$days jours"
                    }
                    else { $retentionLookup[$name] = 'Aucun' }
                }
            }
            & $Log "Dates recuperees pour $($retentionLookup.Count) cameras."
        }
        catch { & $Log "AVERTISSEMENT: Impossible de recuperer les dates d'enregistrement : $_" }
    }

    # ----------------------------------------------------------------
    # PHASE 1 : Snapshots en parallele (uniquement si colonne selectionnee)
    # ----------------------------------------------------------------
    $tempDir   = $null
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
        if ($tempDir -and (Test-Path $tempDir)) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
        return
    }

    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    $groupColors = @{
        'base'   = @{ Bg = 0x44413D; Fg = 0xF4D6CD }
        'stream' = @{ Bg = 0x1D3557; Fg = 0xA8DADC }
        'ret'    = @{ Bg = 0x1B4332; Fg = 0xA6E3A1 }
        'snap'   = @{ Bg = 0x2D2B55; Fg = 0xCBA6F7 }
    }

    try {
        $workbook = $excel.Workbooks.Add()
        $sheet    = $workbook.Sheets.Item(1)
        $sheet.Name = 'Cameras'

        # En-tetes dynamiques selon les colonnes actives
        for ($c = 0; $c -lt $activeColumns.Count; $c++) {
            $col  = $activeColumns[$c]
            $cell = $sheet.Cells.Item(1, $c + 1)
            $cell.Value2              = $col.Header
            $cell.Font.Bold           = $true
            $cell.Font.Size           = 11
            $cell.HorizontalAlignment = -4108
            $cell.Interior.Color      = $groupColors[$col.Group].Bg
            $cell.Font.Color          = $groupColors[$col.Group].Fg
        }

        $sheet.Rows.Item(2).Select() | Out-Null
        $sheet.Application.ActiveWindow.FreezePanes = $true

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

            $ip = if ($cam.Address -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})') { $Matches[1] } else { $cam.Address }

            $si         = $streamLookup[$cam.Name]
            $recStream  = if ($si) { $si.Rec  } else { $null }
            $liveStream = if ($si) { $si.Live } else { $null }
            $extraCount = if ($si) { $si.Extra } else { 0 }
            $sameStream = $recStream -and $liveStream -and ($recStream.Name -eq $liveStream.Name)

            $ret = if ($retentionLookup.ContainsKey($cam.Name)) { $retentionLookup[$cam.Name] } else { 'N/A' }

            # Table de toutes les valeurs possibles, indexees par nom de colonne
            $values = @{
                'Nom'             = $cam.Name
                'Fabricant'       = $cam.DriverFamily
                'Modele'          = $cam.Model
                'IP'              = $ip
                'MAC'             = $cam.MAC
                'Firmware'        = $cam.Firmware
                'ServeurRec'      = $cam.RecorderName
                'Utilisateur'     = $cam.Username
                'MotDePasse'      = if ($includePassword) { $cam.Password } else { '' }
                'CodecEnreg'      = Get-StreamSetting $recStream 'Codec'
                'ResolutionEnreg' = Get-StreamSetting $recStream 'Resolution'
                'FPSEnreg'        = Get-StreamSetting $recStream 'FPS'
                'CodecLive'       = if ($sameStream) { '' } else { Get-StreamSetting $liveStream 'Codec' }
                'ResolutionLive'  = if ($sameStream) { '' } else { Get-StreamSetting $liveStream 'Resolution' }
                'FPSLive'         = if ($sameStream) { '' } else { Get-StreamSetting $liveStream 'FPS' }
                'FluxSupp'        = if ($extraCount -gt 0) { "$extraCount flux supp." } else { '' }
                'Retention'       = $ret
            }

            # Ecriture des colonnes selectionnees dans l'ordre canonique
            for ($c = 0; $c -lt $activeColumns.Count; $c++) {
                $colName = $activeColumns[$c].Name
                if ($colName -eq 'Snapshot') { continue }
                $sheet.Cells.Item($row, $c + 1) = $values[$colName]
            }

            # Snapshot (embedding image dans la cellule)
            if ($snapColIndex -gt 0) {
                $sheet.Rows.Item($row).RowHeight = 90
                $snapFile = $snapPaths[$cam.Name]
                if ($snapFile -and (Test-Path $snapFile)) {
                    try {
                        $cell  = $sheet.Cells.Item($row, $snapColIndex)
                        $shape = $sheet.Shapes.AddPicture(
                            $snapFile,
                            [Microsoft.Office.Core.MsoTriState]::msoFalse,
                            [Microsoft.Office.Core.MsoTriState]::msoCTrue,
                            [double]$cell.Left, [double]$cell.Top,
                            [double]$cell.Width, [double]$cell.Height
                        )
                        $shape.Placement = 1
                    }
                    catch { & $Log "  AVERTISSEMENT: Image '$($cam.Name)' : $_" }
                }
            }

            $row++
        }

        # Mise en forme finale
        if ($snapColIndex -gt 0) { $sheet.Columns.Item($snapColIndex).ColumnWidth = 28 }
        for ($c = 1; $c -le $activeColumns.Count; $c++) {
            if ($c -ne $snapColIndex) { $sheet.Columns.Item($c).AutoFit() | Out-Null }
        }

        $lastCol = $activeColumns.Count
        $range   = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($row - 1, $lastCol))
        $range.Borders.LineStyle = 1
        $range.Borders.Weight    = 2

        $workbook.SaveAs($xlsxPath, 51)
        & $Log "Rapport exporte : $xlsxPath"
    }
    finally {
        try { $workbook.Close($false) } catch {}
        try { $excel.Quit() }           catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        if ($tempDir -and (Test-Path $tempDir)) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
