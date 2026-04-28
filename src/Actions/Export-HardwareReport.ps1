<#
.SYNOPSIS
    Exporte un rapport Excel de tous les equipements Milestone.
    Colonnes selectionnables via une fenetre de choix. Mots de passe exclus par defaut.
#>

function Show-ExportColumnSelector {
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($script:T.EH_DialogTitle)"
        Width="530" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#1E1E2E" FontFamily="Segoe UI">
    <StackPanel Margin="20,16,20,20">

        <DockPanel Margin="0,0,0,10">
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                <Button x:Name="BtnSelectAll" Content="$($script:T.EH_BtnSelectAll)"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="11" Padding="10,4" Cursor="Hand" Margin="0,0,6,0"/>
                <Button x:Name="BtnDeselectAll" Content="$($script:T.EH_BtnDeselAll)"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="11" Padding="10,4" Cursor="Hand"/>
            </StackPanel>
            <TextBlock Text="$($script:T.EH_SelectCols)"
                       Foreground="#CDD6F4" FontSize="13" VerticalAlignment="Center"/>
        </DockPanel>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpHardware)"
                           Foreground="#F4D6CD" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <UniformGrid Columns="3">
                    <CheckBox x:Name="ChkNom"         Content="$($script:T.EH_ChkNom)"        IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFabricant"   Content="$($script:T.EH_ChkFabricant)"  IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkModele"      Content="$($script:T.EH_ChkModele)"     IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkIP"          Content="$($script:T.EH_ChkIP)"         IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkMAC"         Content="$($script:T.EH_ChkMAC)"        IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFirmware"    Content="$($script:T.EH_ChkFirmware)"   IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkServeurRec"  Content="$($script:T.EH_ChkServeurRec)" IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkUtilisateur" Content="$($script:T.EH_ChkUser)"       IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkMotDePasse"  Content="$($script:T.EH_ChkPassword)"   IsChecked="False" Foreground="#FAB387" FontSize="12" Margin="4,5,4,5"/>
                </UniformGrid>
            </StackPanel>
        </Border>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpFlux)"
                           Foreground="#A8DADC" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <UniformGrid Columns="3">
                    <CheckBox x:Name="ChkCodecEnreg"      Content="$($script:T.EH_ChkCodecRec)"  IsChecked="True" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkResolutionEnreg" Content="$($script:T.EH_ChkResRec)"    IsChecked="True" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFPSEnreg"        Content="$($script:T.EH_ChkFpsRec)"   IsChecked="True" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkCodecLive"       Content="$($script:T.EH_ChkCodecLive)" IsChecked="True" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkResolutionLive"  Content="$($script:T.EH_ChkResLive)"   IsChecked="True" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFPSLive"         Content="$($script:T.EH_ChkFpsLive)"  IsChecked="True" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFluxSupp"        Content="$($script:T.EH_ChkFluxSupp)" IsChecked="True" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                </UniformGrid>
            </StackPanel>
        </Border>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpRetention)"
                           Foreground="#A6E3A1" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <CheckBox x:Name="ChkRetention" Content="$($script:T.EH_ChkRetention)" IsChecked="True"
                          Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
            </StackPanel>
        </Border>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpOptions)"
                           Foreground="#CBA6F7" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <CheckBox x:Name="ChkSnapshot" Content="$($script:T.EH_ChkSnapshot)" IsChecked="False"
                          Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
            </StackPanel>
        </Border>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="BtnCancel" Content="$($script:T.EH_BtnCancel)" IsCancel="True"
                    Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                    FontSize="13" Padding="22,8" Cursor="Hand" Margin="0,0,10,0"/>
            <Button x:Name="BtnExport" Content="$($script:T.EH_BtnExport)" IsDefault="True"
                    Background="#A6E3A1" Foreground="#1E1E2E" BorderThickness="0"
                    FontWeight="Bold" FontSize="13" Padding="22,8" Cursor="Hand"/>
        </StackPanel>

    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

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
                $script:T.EH_NoColumn, $script:T.EH_NoColumnTitle,
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
        [Parameter(Mandatory)] [hashtable]$Config,
        [Parameter(Mandatory)] [scriptblock]$Log,
        [Parameter()] [scriptblock]$Cancel = { $false },
        [Parameter()] [scriptblock]$ReportProgress = {}
    )

    function Get-StreamSetting {
        param($stream, [string]$key)
        if ($stream -and $stream.Settings -and $stream.Settings[$key]) { return $stream.Settings[$key] }
        return 'N/A'
    }

    $selectedColumns = Show-ExportColumnSelector
    if ($null -eq $selectedColumns) {
        & $Log $script:T.EH_Cancelled
        return
    }

    $includePassword  = $selectedColumns -contains 'MotDePasse'
    $includeSnapshots = $selectedColumns -contains 'Snapshot'
    $needStreams       = ($selectedColumns | Where-Object {
        $_ -in 'CodecEnreg','ResolutionEnreg','FPSEnreg','CodecLive','ResolutionLive','FPSLive','FluxSupp'
    }).Count -gt 0
    $needRetention    = $selectedColumns -contains 'Retention'

    # Definition des colonnes dans l'ordre canonique
    $allColumnDefs = @(
        @{ Name = 'Nom';             Group = 'base';   Header = $script:T.XL_Nom }
        @{ Name = 'Fabricant';       Group = 'base';   Header = $script:T.XL_Fabricant }
        @{ Name = 'Modele';          Group = 'base';   Header = $script:T.XL_Modele }
        @{ Name = 'IP';              Group = 'base';   Header = $script:T.XL_IP }
        @{ Name = 'MAC';             Group = 'base';   Header = $script:T.XL_MAC }
        @{ Name = 'Firmware';        Group = 'base';   Header = $script:T.XL_Firmware }
        @{ Name = 'ServeurRec';      Group = 'base';   Header = $script:T.XL_ServeurRec }
        @{ Name = 'Utilisateur';     Group = 'base';   Header = $script:T.XL_Utilisateur }
        @{ Name = 'MotDePasse';      Group = 'base';   Header = $script:T.XL_MotDePasse }
        @{ Name = 'CodecEnreg';      Group = 'stream'; Header = $script:T.XL_CodecEnreg }
        @{ Name = 'ResolutionEnreg'; Group = 'stream'; Header = $script:T.XL_ResEnreg }
        @{ Name = 'FPSEnreg';        Group = 'stream'; Header = $script:T.XL_FpsEnreg }
        @{ Name = 'CodecLive';       Group = 'stream'; Header = $script:T.XL_CodecLive }
        @{ Name = 'ResolutionLive';  Group = 'stream'; Header = $script:T.XL_ResLive }
        @{ Name = 'FPSLive';         Group = 'stream'; Header = $script:T.XL_FpsLive }
        @{ Name = 'FluxSupp';        Group = 'stream'; Header = $script:T.XL_FluxSupp }
        @{ Name = 'Retention';       Group = 'ret';    Header = $script:T.XL_Retention }
        @{ Name = 'Snapshot';        Group = 'snap';   Header = $script:T.XL_Snapshot }
    )

    $activeColumns = @($allColumnDefs | Where-Object { $selectedColumns -contains $_.Name })

    $snapColIndex = 0
    for ($i = 0; $i -lt $activeColumns.Count; $i++) {
        if ($activeColumns[$i].Name -eq 'Snapshot') { $snapColIndex = $i + 1; break }
    }

    & $Log $script:T.EH_LogGenerating
    if ($includePassword) {
        $camReport = @(Get-VmsCameraReport -IncludePlainTextPassword)
    } else {
        $camReport = @(Get-VmsCameraReport)
    }
    $total = $camReport.Count
    & $Log ($script:T.EH_LogFound -f $total)

    & $Log $script:T.EH_LogLoadCams
    $vmsCameras   = @(Get-VmsCamera)
    $vmsCamByName = @{}
    $vmsCamByPath = @{}
    foreach ($c in $vmsCameras) {
        $vmsCamByName[$c.Name] = $c
        $vmsCamByPath[$c.Path] = $c.Name
    }

    $streamLookup = @{}
    if ($needStreams) {
        & $Log $script:T.EH_LogStreams
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
            & $Log ($script:T.EH_LogStreamsOk -f $allStreams.Count, $streamLookup.Count)
        }
        catch { & $Log ($script:T.EH_LogStreamsErr -f $_) }
    }

    $retentionLookup = @{}
    if ($needRetention) {
        & $Log $script:T.EH_LogPlayback
        try {
            $playbackData = @($vmsCameras | Get-PlaybackInfo -Parallel -ErrorAction Stop)
            foreach ($pb in $playbackData) {
                $name = $vmsCamByPath[$pb.Path]
                if ($name) {
                    if ($pb.Begin -and $pb.End) {
                        $days = [int]($pb.End - $pb.Begin).TotalDays
                        $retentionLookup[$name] = "$days j"
                    }
                    else { $retentionLookup[$name] = $script:T.XL_Aucun }
                }
            }
            & $Log ($script:T.EH_LogPlaybackOk -f $retentionLookup.Count)
        }
        catch { & $Log ($script:T.EH_LogPlaybackErr -f $_) }
    }

    $tempDir   = $null
    $snapPaths = @{}

    if ($includeSnapshots) {
        & $Log $script:T.EH_LogSnaps
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
                        & $Log ($script:T.EH_LogSnapOk -f $received, $jobs.Count, $job.Name)
                    }
                    else { & $Log ($script:T.EH_LogSnapEmpty -f $job.Name) }
                }
                catch { & $Log ($script:T.EH_LogSnapErr -f $job.Name, $_) }
                finally { $job.PS.Dispose() }
                & $ReportProgress ($jobs.Count - $pending.Count) $jobs.Count
            }
            if ($pending.Count -gt 0) { Start-Sleep -Milliseconds 150 }
        }

        $pool.Close()
        $pool.Dispose()
        & $Log ($script:T.EH_LogSnapsDone -f $snapPaths.Count, $jobs.Count)
    }

    if (-not (Test-Path $Config.outputDirectory)) {
        New-Item -Path $Config.outputDirectory -ItemType Directory -Force | Out-Null
    }
    $xlsxPath = Join-Path $Config.outputDirectory $script:T.XL_FileName

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    }
    catch {
        & $Log $script:T.EH_LogNoExcel
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
        $sheet.Name = $script:T.XL_SheetName

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

        & $Log $script:T.EH_LogBuilding

        foreach ($cam in $camReport) {
            if (& $Cancel) {
                & $Log ($script:T.EH_LogCancelled -f $count, $total)
                break
            }

            $count++
            & $ReportProgress $count $total
            & $Log ($script:T.EH_LogCamRow -f $count, $total, $cam.Name)

            $ip = if ($cam.Address -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})') { $Matches[1] } else { $cam.Address }

            $si         = $streamLookup[$cam.Name]
            $recStream  = if ($si) { $si.Rec  } else { $null }
            $liveStream = if ($si) { $si.Live } else { $null }
            $extraCount = if ($si) { $si.Extra } else { 0 }
            $sameStream = $recStream -and $liveStream -and ($recStream.Name -eq $liveStream.Name)

            $ret = if ($retentionLookup.ContainsKey($cam.Name)) { $retentionLookup[$cam.Name] } else { 'N/A' }

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
                'FluxSupp'        = if ($extraCount -gt 0) { $script:T.XL_ExtraFlux -f $extraCount } else { '' }
                'Retention'       = $ret
            }

            for ($c = 0; $c -lt $activeColumns.Count; $c++) {
                $colName = $activeColumns[$c].Name
                if ($colName -eq 'Snapshot') { continue }
                $sheet.Cells.Item($row, $c + 1) = $values[$colName]
            }

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
                    catch { & $Log ($script:T.EH_LogImgErr -f $cam.Name, $_) }
                }
            }

            $row++
        }

        if ($snapColIndex -gt 0) { $sheet.Columns.Item($snapColIndex).ColumnWidth = 28 }
        for ($c = 1; $c -le $activeColumns.Count; $c++) {
            if ($c -ne $snapColIndex) { $sheet.Columns.Item($c).AutoFit() | Out-Null }
        }

        $lastCol = $activeColumns.Count
        $range   = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($row - 1, $lastCol))
        $range.Borders.LineStyle = 1
        $range.Borders.Weight    = 2

        $workbook.SaveAs($xlsxPath, 51)
        & $Log ($script:T.EH_LogSaved -f $xlsxPath)
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
