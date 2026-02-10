<#
.SYNOPSIS
    Exporte un rapport CSV de tous les equipements (hardware) du VMS Milestone.
.PARAMETER Config
    Hashtable de configuration (outputDirectory, csvDelimiter, csvEncoding).
.PARAMETER Log
    Scriptblock callback pour logger vers l'UI.
#>

function Export-HardwareReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log
    )

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
