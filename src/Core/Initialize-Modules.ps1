<#
.SYNOPSIS
    Verifie et importe les modules requis pour l'application.
.DESCRIPTION
    Installe automatiquement les modules manquants (MilestonePSTools, ImportExcel)
    puis les importe dans la session courante.
#>

function Initialize-RequiredModules {
    [CmdletBinding()]
    param(
        [Parameter()]
        [scriptblock]$Log = { param($Message) Write-Host $Message }
    )

    $modules = @(
        @{ Name = 'MilestonePSTools'; Required = $true }
        @{ Name = 'ImportExcel';      Required = $true }
    )

    foreach ($mod in $modules) {
        $name = $mod.Name

        if (-not (Get-Module -ListAvailable -Name $name)) {
            & $Log "Installation du module $name..."
            try {
                Install-Module -Name $name -Force -Scope CurrentUser -ErrorAction Stop
                & $Log "Module $name installe avec succes."
            }
            catch {
                if ($mod.Required) {
                    throw "Impossible d'installer le module requis '$name': $_"
                }
                & $Log "AVERTISSEMENT: Echec de l'installation de $name : $_"
            }
        }
        else {
            & $Log "Module $name deja disponible."
        }

        try {
            Import-Module -Name $name -Force -ErrorAction Stop
            & $Log "Module $name importe."
        }
        catch {
            throw "Impossible d'importer le module '$name': $_"
        }
    }
}
