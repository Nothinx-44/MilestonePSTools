<#
.SYNOPSIS
    Exporte un rapport CSV de tous les equipements (hardware) du VMS Milestone.
#>

function Export-HardwareReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log
    )

    # Avertissement : le rapport contient les mots de passe en clair
    $confirm = [System.Windows.MessageBox]::Show(
        "Ce rapport inclut les mots de passe des cameras en clair dans le fichier CSV.`n`nAssurez-vous de stocker le fichier dans un emplacement securise.`n`nContinuer ?",
        'Avertissement — Donnees sensibles',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
        & $Log "Export annule par l'utilisateur."
        return
    }

    & $Log "Generation du rapport hardware..."

    $cameras = Get-VmsCameraReport -IncludePlainTextPassword | ForEach-Object {
        [PSCustomObject]@{
            Nom        = $_.Name
            Fabricant  = $_.DriverFamily
            Modele     = $_.Model
            IP         = $_.Address
            MAC        = $_.MAC
            Firmware   = $_.Firmware
            Activation = $_.Enabled
            ServeurRec = $_.RecorderName
            User       = $_.Username
            Pass       = $_.Password
            Gps        = $_.GpsCoordinates
        }
    }

    & $Log "$($cameras.Count) equipements trouves."

    if (-not (Test-Path $Config.outputDirectory)) {
        New-Item -Path $Config.outputDirectory -ItemType Directory -Force | Out-Null
    }

    $csvPath = Join-Path $Config.outputDirectory 'Liste_des_Cameras.csv'

    $cameras | Export-Csv `
        -Path $csvPath `
        -NoTypeInformation `
        -Encoding $Config.csvEncoding `
        -Delimiter $Config.csvDelimiter

    & $Log "Rapport exporte : $csvPath"
}
