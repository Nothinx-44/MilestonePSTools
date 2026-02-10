<#
.SYNOPSIS
    Cree des Device Groups dans Milestone organises par modele de camera.
.DESCRIPTION
    Cree un dossier parent "Modele" dans les Device Groups, puis un sous-dossier
    par modele de camera, et y ajoute les cameras correspondantes.
.PARAMETER Config
    Hashtable de configuration.
.PARAMETER Log
    Scriptblock callback pour logger vers l'UI.
#>

function Set-CameraGroupByModel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [scriptblock]$Log
    )

    $parentFolderName = 'Modele'

    & $Log "Recuperation des informations cameras..."
    $cameras = Get-VmsCameraReport
    & $Log "$($cameras.Count) cameras trouvees."

    # Creer ou recuperer le dossier parent
    $parentFolder = Get-VmsDeviceGroup -Name $parentFolderName -ErrorAction SilentlyContinue
    if (-not $parentFolder) {
        $parentFolder = New-VmsDeviceGroup -Name $parentFolderName
        & $Log "Dossier parent '$parentFolderName' cree."
    }

    # Grouper par modele
    $camerasByModel = $cameras | Group-Object -Property Model
    & $Log "$($camerasByModel.Count) modeles differents detectes."

    foreach ($group in $camerasByModel) {
        $model = $group.Name
        if ([string]::IsNullOrWhiteSpace($model)) {
            $model = 'Inconnu'
        }

        # Creer ou recuperer le sous-dossier du modele
        $deviceGroup = Get-VmsDeviceGroup -ParentGroup $parentFolder -Name $model -ErrorAction SilentlyContinue
        if (-not $deviceGroup) {
            $deviceGroup = New-VmsDeviceGroup -ParentGroup $parentFolder -Name $model
        }

        foreach ($camera in $group.Group) {
            Add-VmsDeviceGroupMember -Group $deviceGroup -DeviceId $camera.Id
        }

        & $Log "Modele '$model' : $($group.Count) camera(s) ajoutee(s)."
    }

    & $Log "Organisation par modele terminee."
}
