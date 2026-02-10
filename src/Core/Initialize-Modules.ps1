<#
.SYNOPSIS
    Verifie et importe les modules requis pour l'application.
.DESCRIPTION
    Supporte deux modes d'installation :
      - Online  : telecharge depuis PowerShell Gallery (Install-Module)
      - Offline : charge depuis le dossier local Dependencies/
    Le mode est determine par le parametre InstallMode.
    Le dossier offline est prioritaire : si le module y est present, il est utilise
    meme en mode online.
#>

function Initialize-RequiredModules {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('Online', 'Offline')]
        [string]$InstallMode = 'Online',

        [Parameter()]
        [string]$DependenciesPath,

        [Parameter()]
        [scriptblock]$Log = { param($Message) Write-Host $Message }
    )

    $modules = @(
        @{ Name = 'MilestonePSTools' }
    )

    foreach ($mod in $modules) {
        $name = $mod.Name
        $importedFromLocal = $false

        # --- Tentative de chargement depuis le dossier Dependencies/ ---
        if ($DependenciesPath -and (Test-Path $DependenciesPath)) {
            $localModulePath = Join-Path $DependenciesPath $name
            if (Test-Path $localModulePath) {
                & $Log "Chargement de $name depuis Dependencies/..."
                try {
                    Import-Module $localModulePath -Force -ErrorAction Stop
                    & $Log "Module $name charge (offline)."
                    $importedFromLocal = $true
                }
                catch {
                    & $Log "AVERTISSEMENT: Echec du chargement local de $name : $_"
                }
            }
        }

        if ($importedFromLocal) { continue }

        # --- Mode Offline strict : le module doit etre dans Dependencies/ ---
        if ($InstallMode -eq 'Offline') {
            # Verifier si deja installe sur le systeme
            if (Get-Module -ListAvailable -Name $name) {
                & $Log "Module $name trouve dans l'environnement systeme."
                Import-Module -Name $name -Force -ErrorAction Stop
                & $Log "Module $name importe."
                continue
            }

            throw ("Module '$name' introuvable. En mode Offline, placez le module dans " +
                   "le dossier Dependencies/ avec : .\Save-Dependencies.ps1")
        }

        # --- Mode Online : installer depuis PSGallery si absent ---
        if (-not (Get-Module -ListAvailable -Name $name)) {
            & $Log "Installation de $name depuis PowerShell Gallery..."
            try {
                Install-Module -Name $name -Force -Scope CurrentUser -ErrorAction Stop
                & $Log "Module $name installe."
            }
            catch {
                throw "Impossible d'installer le module '$name': $_"
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
