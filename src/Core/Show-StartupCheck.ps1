<#
.SYNOPSIS
    Fenetre GUI de verification des dependances au demarrage.
#>

function Show-StartupCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AppRoot
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    $script:_SC_Result      = $false
    $script:_SC_AppRoot     = $AppRoot
    $script:_SC_DepsPath    = Join-Path $AppRoot 'Dependencies'
    $script:_SC_IsOffline   = Test-Path $script:_SC_DepsPath
    $script:_SC_DepRows     = @{}
    $script:_SC_Modules     = @(
        @{ Name = 'MilestonePSTools'; Description = $script:T.SC_ModuleDesc }
    )

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Milestone Toolkit"
        Width="580" Height="530"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E2E"
        FontFamily="Segoe UI">

    <Window.Resources>
        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="#89B4FA"/>
            <Setter Property="Foreground"      Value="#1E1E2E"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="20,10"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#B4D0FF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#7AA2F7"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SecondaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="Transparent"/>
            <Setter Property="Foreground"      Value="#A6ADC8"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="Padding"         Value="20,10"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush"     Value="#45475A"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#313244"/>
                                <Setter Property="Foreground" Value="#CDD6F4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="WarnBtn" TargetType="Button">
            <Setter Property="Background"      Value="#F9A825"/>
            <Setter Property="Foreground"      Value="#1E1E2E"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="20,10"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#FFCA28"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#F57F17"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="GreenBtn" TargetType="Button">
            <Setter Property="Background"      Value="#A6E3A1"/>
            <Setter Property="Foreground"      Value="#1E1E2E"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="14,8"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#C3F0BE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#89D4A1"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#181825" Padding="32,24,32,20">
            <StackPanel Orientation="Horizontal">
                <Border Width="38" Height="38" Background="#89B4FA" CornerRadius="8"
                        Margin="0,0,14,0" VerticalAlignment="Center">
                    <TextBlock Text="M" FontSize="22" FontWeight="Bold"
                               Foreground="#1E1E2E"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Milestone Toolkit"
                               FontSize="20" FontWeight="Bold" Foreground="#CDD6F4"/>
                    <TextBlock x:Name="HeaderSubtitle"
                               Text="Verification des dependances au demarrage"
                               FontSize="12" Foreground="#6C7086" Margin="0,3,0,0"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="24,20,24,0">
            <StackPanel x:Name="DepsPanel"/>
        </ScrollViewer>

        <Border x:Name="OfflineBanner" Grid.Row="2"
                Background="#2A2A1A" BorderBrush="#F9A825" BorderThickness="0,0,0,0"
                CornerRadius="6" Margin="24,12,24,0" Padding="16,12"
                Visibility="Collapsed">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock x:Name="OfflineBannerTitle"
                               Text="" FontSize="12" FontWeight="SemiBold" Foreground="#F9A825"/>
                    <TextBlock x:Name="OfflineBannerText"
                               Text="" FontSize="11" Foreground="#A6ADC8"
                               TextWrapping="Wrap" Margin="0,3,0,0"/>
                </StackPanel>
                <Button x:Name="BtnSaveDeps" Grid.Column="1"
                        Content="" Margin="12,0,0,0"
                        Style="{StaticResource GreenBtn}" Visibility="Collapsed"/>
            </Grid>
        </Border>

        <StackPanel Grid.Row="3" Margin="24,12,24,8">
            <ProgressBar x:Name="ProgressBar"
                         Height="3" IsIndeterminate="True"
                         Background="#313244" Foreground="#89B4FA"
                         BorderThickness="0" Visibility="Collapsed" Margin="0,0,0,10"/>
            <TextBlock x:Name="StatusText"
                       Text="" FontSize="12" Foreground="#A6ADC8" TextWrapping="Wrap"/>
        </StackPanel>

        <Border Grid.Row="4" Background="#181825" Padding="24,14">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="BtnQuit"    Content=""
                        Style="{StaticResource SecondaryBtn}" Margin="0,0,10,0"/>
                <Button x:Name="BtnInstall" Content=""
                        Style="{StaticResource WarnBtn}"
                        Visibility="Collapsed" Margin="0,0,10,0"/>
                <Button x:Name="BtnLaunch"  Content=""
                        Style="{StaticResource PrimaryBtn}" IsEnabled="False"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
'@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $script:_SC_Win         = [System.Windows.Markup.XamlReader]::Load($reader)

    $script:_SC_DepsPanel          = $script:_SC_Win.FindName('DepsPanel')
    $script:_SC_Progress           = $script:_SC_Win.FindName('ProgressBar')
    $script:_SC_Status             = $script:_SC_Win.FindName('StatusText')
    $script:_SC_BtnInstall         = $script:_SC_Win.FindName('BtnInstall')
    $script:_SC_BtnLaunch          = $script:_SC_Win.FindName('BtnLaunch')
    $script:_SC_BtnQuit            = $script:_SC_Win.FindName('BtnQuit')
    $script:_SC_BtnSaveDeps        = $script:_SC_Win.FindName('BtnSaveDeps')
    $script:_SC_OfflineBanner      = $script:_SC_Win.FindName('OfflineBanner')
    $script:_SC_OfflineBannerTitle = $script:_SC_Win.FindName('OfflineBannerTitle')
    $script:_SC_OfflineBannerText  = $script:_SC_Win.FindName('OfflineBannerText')

    # Textes traduits
    $script:_SC_Win.Title                           = $script:T.SC_WindowTitle
    $script:_SC_Win.FindName('HeaderSubtitle').Text = $script:T.SC_Subtitle
    $script:_SC_BtnQuit.Content                     = $script:T.SC_BtnQuit
    $script:_SC_BtnInstall.Content                  = $script:T.SC_BtnInstall
    $script:_SC_BtnLaunch.Content                   = $script:T.SC_BtnLaunch
    $script:_SC_BtnSaveDeps.Content                 = $script:T.SC_BtnSaveDeps
    $script:_SC_Status.Text                         = $script:T.SC_StatusInit

    foreach ($mod in $script:_SC_Modules) {
        $name = $mod.Name
        $desc = $mod.Description

        $indicator = [System.Windows.Shapes.Ellipse]::new()
        $indicator.Width  = 12
        $indicator.Height = 12
        $indicator.Fill   = [System.Windows.Media.Brushes]::Gray
        $indicator.VerticalAlignment = 'Center'
        $indicator.Margin = [System.Windows.Thickness]::new(0,0,14,0)

        $lblName = [System.Windows.Controls.TextBlock]::new()
        $lblName.Text       = $name
        $lblName.FontSize   = 13
        $lblName.FontWeight = [System.Windows.FontWeights]::SemiBold
        $lblName.Foreground = [System.Windows.Media.Brushes]::White

        $lblDesc = [System.Windows.Controls.TextBlock]::new()
        $lblDesc.Text       = $desc
        $lblDesc.FontSize   = 11
        $lblDesc.Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(108,112,134))
        $lblDesc.Margin     = [System.Windows.Thickness]::new(0,3,0,0)

        $nameStack = [System.Windows.Controls.StackPanel]::new()
        $nameStack.VerticalAlignment = 'Center'
        [void]$nameStack.Children.Add($lblName)
        [void]$nameStack.Children.Add($lblDesc)

        $lblStatus = [System.Windows.Controls.TextBlock]::new()
        $lblStatus.Text              = $script:T.SC_StatusWaiting
        $lblStatus.FontSize          = 12
        $lblStatus.Foreground        = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(166,173,200))
        $lblStatus.VerticalAlignment = 'Center'

        $rowGrid = [System.Windows.Controls.Grid]::new()
        $c0 = [System.Windows.Controls.ColumnDefinition]::new(); $c0.Width = [System.Windows.GridLength]::Auto
        $c1 = [System.Windows.Controls.ColumnDefinition]::new(); $c1.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
        $c2 = [System.Windows.Controls.ColumnDefinition]::new(); $c2.Width = [System.Windows.GridLength]::Auto
        [void]$rowGrid.ColumnDefinitions.Add($c0)
        [void]$rowGrid.ColumnDefinitions.Add($c1)
        [void]$rowGrid.ColumnDefinitions.Add($c2)
        [System.Windows.Controls.Grid]::SetColumn($indicator, 0)
        [System.Windows.Controls.Grid]::SetColumn($nameStack,  1)
        [System.Windows.Controls.Grid]::SetColumn($lblStatus,  2)
        [void]$rowGrid.Children.Add($indicator)
        [void]$rowGrid.Children.Add($nameStack)
        [void]$rowGrid.Children.Add($lblStatus)

        $card = [System.Windows.Controls.Border]::new()
        $card.Background   = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(24,24,37))
        $card.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $card.Padding      = [System.Windows.Thickness]::new(16,14,16,14)
        $card.Margin       = [System.Windows.Thickness]::new(0,0,0,10)
        $card.Child        = $rowGrid
        [void]$script:_SC_DepsPanel.Children.Add($card)

        $script:_SC_DepRows[$name] = @{ Indicator = $indicator; Status = $lblStatus; Available = $false }
    }

    $script:_SC_Refresh = {
        $script:_SC_Win.Dispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Render, [Action]{}
        )
    }

    $script:_SC_SetStatus = {
        param([string]$Name, [string]$State, [string]$Label)
        $row = $script:_SC_DepRows[$Name]
        $row.Status.Text = $Label
        $color = switch ($State) {
            'checking'   { [System.Windows.Media.Color]::FromRgb(249,168, 37) }
            'ok'         { [System.Windows.Media.Color]::FromRgb(166,227,161) }
            'missing'    { [System.Windows.Media.Color]::FromRgb(243,139,168) }
            'installing' { [System.Windows.Media.Color]::FromRgb(137,180,250) }
            'error'      { [System.Windows.Media.Color]::FromRgb(243,139,168) }
        }
        $row.Indicator.Fill = [System.Windows.Media.SolidColorBrush]::new($color)
        if ($State -eq 'ok')                 { $row.Available = $true  }
        if ($State -in @('missing','error')) { $row.Available = $false }
        & $script:_SC_Refresh
    }

    $script:_SC_Check = {
        $allOk = $true

        foreach ($mod in $script:_SC_Modules) {
            $name = $mod.Name
            & $script:_SC_SetStatus $name 'checking' $script:T.SC_Checking
            $found = $false

            if ($script:_SC_IsOffline) {
                $localPath = Join-Path $script:_SC_DepsPath $name
                if (Test-Path $localPath) {
                    & $script:_SC_SetStatus $name 'ok' $script:T.SC_LocalCache
                    $found = $true
                }
            }

            if (-not $found) {
                $installed = Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue
                if ($installed) {
                    $ver = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version
                    & $script:_SC_SetStatus $name 'ok' ($script:T.SC_Installed -f $ver)
                    $found = $true
                }
            }

            if (-not $found) {
                & $script:_SC_SetStatus $name 'missing' $script:T.SC_Missing
                $allOk = $false
            }
        }

        if ($script:_SC_IsOffline) {
            $script:_SC_OfflineBanner.Visibility    = 'Visible'
            $script:_SC_OfflineBannerTitle.Text     = $script:T.SC_OfflineCacheTitle

            $missingLocally = $script:_SC_Modules | Where-Object {
                -not (Test-Path (Join-Path $script:_SC_DepsPath $_.Name))
            }

            if ($missingLocally) {
                $names = ($missingLocally | ForEach-Object { $_.Name }) -join ', '
                $script:_SC_OfflineBannerText.Text  = $script:T.SC_OfflineCacheMissing -f $names
                $script:_SC_BtnSaveDeps.Visibility  = 'Visible'
            }
            else {
                $script:_SC_OfflineBannerText.Text  = $script:T.SC_OfflineCacheOk
                $script:_SC_BtnSaveDeps.Visibility  = 'Collapsed'
            }
        }
        else {
            $script:_SC_OfflineBanner.Visibility    = 'Visible'
            $script:_SC_OfflineBannerTitle.Text     = $script:T.SC_OnlineTitle
            $script:_SC_OfflineBannerText.Text      = $script:T.SC_OnlineText
            $script:_SC_BtnSaveDeps.Visibility      = 'Visible'
        }

        if ($allOk) {
            $script:_SC_BtnLaunch.IsEnabled   = $true
            $script:_SC_BtnInstall.Visibility = 'Collapsed'
            $script:_SC_Status.Text = $script:T.SC_AllOk
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(166,227,161))
        }
        elseif ($script:_SC_IsOffline) {
            $script:_SC_BtnInstall.Visibility = 'Collapsed'
            $script:_SC_Status.Text = $script:T.SC_OfflineMissing
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(243,139,168))
        }
        else {
            $script:_SC_BtnInstall.Visibility = 'Visible'
            $script:_SC_Status.Text = $script:T.SC_NeedInstall
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(249,168,37))
        }

        & $script:_SC_Refresh
    }

    # Preparation pour Install-Module (bouton "Installer les dependances")
    $script:_SC_PrepareGallery = {
        # TLS 1.2 requis par PSGallery (PS 5.1 utilise TLS 1.0 par defaut)
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        # NuGet requis pour Install-Module
        $null = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 `
            -Force -Scope CurrentUser -ErrorAction SilentlyContinue
    }

    # Telechargement direct depuis l'API NuGet PSGallery
    # N'utilise que Invoke-WebRequest + ZipFile .NET — aucune dependance PowerShellGet/NuGet
    $script:_SC_DownloadModuleNuGet = {
        param([string]$ModuleName, [string]$DestFolder)

        $ProgressPreference = 'SilentlyContinue'

        # TLS 1.2 obligatoire
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

        $nupkgUrl    = "https://www.powershellgallery.com/api/v2/package/$ModuleName"
        $tempNupkg   = Join-Path $env:TEMP "$ModuleName.nupkg"
        $tempExtract = Join-Path $env:TEMP "$ModuleName.extracted"

        # Etape 1 — telechargement
        try {
            Invoke-WebRequest -Uri $nupkgUrl -OutFile $tempNupkg -UseBasicParsing -ErrorAction Stop
        }
        catch {
            throw "[Etape 1 - Telechargement] $($_.Exception.GetType().Name): $($_.Exception.Message)"
        }

        if (-not (Test-Path $tempNupkg) -or (Get-Item $tempNupkg).Length -eq 0) {
            throw "[Etape 1] Fichier telecharge vide ou absent : $tempNupkg"
        }

        # Etape 2 — extraction ZIP
        try {
            if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            [System.IO.Compression.ZipFile]::ExtractToDirectory($tempNupkg, $tempExtract)
        }
        catch {
            throw "[Etape 2 - Extraction] $($_.Exception.GetType().Name): $($_.Exception.Message)"
        }

        # Etape 3 — copie des fichiers du module (sans les metadonnees NuGet)
        try {
            $excludeNames = @('[Content_Types].xml')
            $excludeExts  = @('.nuspec', '.psmdcp')
            $excludeDirs  = @('_rels', 'package')
            Get-ChildItem $tempExtract | Where-Object {
                $_.Name      -notin $excludeNames -and
                $_.Extension -notin $excludeExts  -and
                $_.Name      -notin $excludeDirs
            } | Copy-Item -Destination $DestFolder -Recurse -Force -ErrorAction Stop
        }
        catch {
            throw "[Etape 3 - Copie] $($_.Exception.GetType().Name): $($_.Exception.Message)"
        }

        # Nettoyage
        Remove-Item $tempNupkg, $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
    }

    $script:_SC_Install = {
        $script:_SC_BtnInstall.IsEnabled = $false
        $script:_SC_BtnQuit.IsEnabled    = $false
        $script:_SC_Progress.Visibility  = 'Visible'
        $script:_SC_Status.Foreground    = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(166,173,200))

        $anyError = $false
        $oldProgress    = $ProgressPreference
        $oldInformation = $InformationPreference
        $ProgressPreference    = 'SilentlyContinue'
        $InformationPreference = 'SilentlyContinue'

        $script:_SC_Status.Text = $script:T.SC_NuGet
        & $script:_SC_Refresh
        & $script:_SC_PrepareGallery

        foreach ($mod in $script:_SC_Modules) {
            $name = $mod.Name
            if ($script:_SC_DepRows[$name].Available) { continue }

            & $script:_SC_SetStatus $name 'installing' $script:T.SC_Installing
            $script:_SC_Status.Text = $script:T.SC_InstallingMod -f $name
            & $script:_SC_Refresh

            try {
                $null = Install-Module -Name $name -Repository PSGallery `
                    -Force -Scope CurrentUser `
                    -ErrorAction Stop -WarningAction SilentlyContinue
                $installed = Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue
                $ver = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version
                & $script:_SC_SetStatus $name 'ok' ($script:T.SC_Installed -f $ver)
            }
            catch {
                & $script:_SC_SetStatus $name 'error' ($script:T.SC_ErrGeneric -f $_.Exception.Message)
                $anyError = $true
            }
        }

        $ProgressPreference    = $oldProgress
        $InformationPreference = $oldInformation
        $script:_SC_Progress.Visibility = 'Collapsed'
        $script:_SC_BtnQuit.IsEnabled   = $true

        if ($anyError) {
            $script:_SC_BtnInstall.IsEnabled = $true
            $script:_SC_Status.Text = $script:T.SC_InstallError
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(243,139,168))
        }
        else {
            $script:_SC_BtnLaunch.IsEnabled   = $true
            $script:_SC_BtnInstall.Visibility = 'Collapsed'
            $script:_SC_Status.Text = $script:T.SC_InstallDone
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(166,227,161))
            $script:_SC_BtnSaveDeps.Visibility = 'Visible'
        }
        & $script:_SC_Refresh
    }

    $script:_SC_SaveDeps = {
        $confirm = [System.Windows.MessageBox]::Show(
            $script:T.SC_SaveConfirm, $script:T.SC_SaveTitle, 'YesNo', 'Question'
        )
        if ($confirm -ne 'Yes') { return }

        $script:_SC_BtnSaveDeps.IsEnabled = $false
        $script:_SC_BtnQuit.IsEnabled     = $false
        $script:_SC_Progress.Visibility   = 'Visible'
        $script:_SC_Status.Foreground     = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(166,173,200))

        $anyError = $false
        $oldProgress    = $ProgressPreference
        $oldInformation = $InformationPreference
        $ProgressPreference    = 'SilentlyContinue'
        $InformationPreference = 'SilentlyContinue'

        if (-not (Test-Path $script:_SC_DepsPath)) {
            New-Item -Path $script:_SC_DepsPath -ItemType Directory -Force | Out-Null
        }

        foreach ($mod in $script:_SC_Modules) {
            $name = $mod.Name
            & $script:_SC_SetStatus $name 'installing' $script:T.SC_Saving
            $script:_SC_Status.Text = $script:T.SC_DownloadingMod -f $name
            & $script:_SC_Refresh

            try {
                # TLS 1.2 obligatoire pour PSGallery
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                # Noms uniques pour eviter les conflits avec des runs precedents
                $runId       = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
                $localPath   = Join-Path $script:_SC_DepsPath $name
                $tempNupkg   = Join-Path $env:TEMP "$name.$runId.zip"
                $tempExtract = Join-Path $env:TEMP "$name.$runId.extract"

                # Nettoyage du dossier de destination
                if (Test-Path $localPath) { Remove-Item $localPath -Recurse -Force -ErrorAction SilentlyContinue }
                New-Item $localPath -ItemType Directory -Force | Out-Null

                # Telechargement via WebClient (plus simple et fiable qu'Invoke-WebRequest)
                $wc = [System.Net.WebClient]::new()
                $wc.DownloadFile("https://www.powershellgallery.com/api/v2/package/$name", $tempNupkg)
                $wc.Dispose()

                # Extraction (nupkg = ZIP) — Expand-Archive -Force gere les dossiers existants
                Expand-Archive -Path $tempNupkg -DestinationPath $tempExtract -Force -ErrorAction Stop

                # Copie des fichiers du module (sans les metadonnees NuGet)
                $excludeNames = @('[Content_Types].xml')
                $excludeExts  = @('.nuspec', '.psmdcp')
                $excludeDirs  = @('_rels', 'package')
                Get-ChildItem $tempExtract | Where-Object {
                    $_.Name -notin $excludeNames -and
                    $_.Extension -notin $excludeExts -and
                    $_.Name -notin $excludeDirs
                } | Copy-Item -Destination $localPath -Recurse -Force

                # Nettoyage
                Remove-Item $tempNupkg, $tempExtract -Recurse -Force -ErrorAction SilentlyContinue

                & $script:_SC_SetStatus $name 'ok' $script:T.SC_CacheOk
            }
            catch {
                $errMsg = "$($_.Exception.GetType().Name): $($_.Exception.Message)"
                if ($_.Exception.InnerException) {
                    $errMsg += " | $($_.Exception.InnerException.Message)"
                }
                # MessageBox pour voir l'erreur exacte
                [System.Windows.MessageBox]::Show(
                    $errMsg, 'Erreur SaveDeps',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
                & $script:_SC_SetStatus $name 'error' ($script:T.SC_ErrGeneric -f $errMsg)
                $anyError = $true
            }
        }

        $ProgressPreference    = $oldProgress
        $InformationPreference = $oldInformation
        $script:_SC_Progress.Visibility   = 'Collapsed'
        $script:_SC_BtnSaveDeps.IsEnabled = $true
        $script:_SC_BtnQuit.IsEnabled     = $true

        if (-not $anyError) {
            $script:_SC_IsOffline = $true
            $script:_SC_BtnSaveDeps.Content = $script:T.SC_BtnUpdateCache
            & $script:_SC_Check
        }
        else {
            $script:_SC_Status.Text = $script:T.SC_SaveError
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(243,139,168))
            & $script:_SC_Refresh
        }
    }

    $script:_SC_Win.Add_Loaded({
        try   { & $script:_SC_Check }
        catch {
            [System.Windows.MessageBox]::Show(
                ($script:T.SC_ErrCheck -f $_), $script:T.SC_ErrTitle, 'OK', 'Error'
            ) | Out-Null
        }
    })

    $script:_SC_BtnInstall.Add_Click({
        try   { & $script:_SC_Install }
        catch {
            [System.Windows.MessageBox]::Show(
                ($script:T.SC_ErrInstall -f $_), $script:T.SC_ErrTitle, 'OK', 'Error'
            ) | Out-Null
            $script:_SC_BtnInstall.IsEnabled = $true
            $script:_SC_BtnQuit.IsEnabled    = $true
            $script:_SC_Progress.Visibility  = 'Collapsed'
        }
    })

    $script:_SC_BtnSaveDeps.Add_Click({
        try   { & $script:_SC_SaveDeps }
        catch {
            [System.Windows.MessageBox]::Show(
                ($script:T.SC_ErrGeneric -f $_), $script:T.SC_ErrTitle, 'OK', 'Error'
            ) | Out-Null
        }
    })

    $script:_SC_BtnLaunch.Add_Click({
        $script:_SC_Result = $true
        $script:_SC_Win.Close()
    })

    $script:_SC_BtnQuit.Add_Click({
        $script:_SC_Result = $false
        $script:_SC_Win.Close()
    })

    [void]$script:_SC_Win.ShowDialog()
    return $script:_SC_Result
}
