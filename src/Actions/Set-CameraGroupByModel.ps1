<#
.SYNOPSIS
    Cree des Device Groups dans Milestone organises par modele de camera.
#>

function Set-CameraGroupByModel {
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

    $parentFolderName = 'Modele'

    & $Log "Recuperation des informations cameras..."
    $cameras = Get-VmsCameraReport
    & $Log "$($cameras.Count) cameras trouvees."

    $parentFolder = Get-VmsDeviceGroup -Name $parentFolderName -ErrorAction SilentlyContinue
    if (-not $parentFolder) {
        $parentFolder = New-VmsDeviceGroup -Name $parentFolderName
        & $Log "Dossier parent '$parentFolderName' cree."
    }

    $camerasByModel = $cameras | Group-Object -Property Model
    $total          = $camerasByModel.Count
    & $Log "$total modeles differents detectes."

    $done = 0
    foreach ($group in $camerasByModel) {
        if (& $Cancel) {
            & $Log "AVERTISSEMENT: Operation annulee apres $done / $total modeles."
            break
        }

        $done++
        & $ReportProgress $done $total

        $model = if ([string]::IsNullOrWhiteSpace($group.Name)) { 'Inconnu' } else { $group.Name }

        $deviceGroup = Get-VmsDeviceGroup -ParentGroup $parentFolder -Name $model -ErrorAction SilentlyContinue
        if (-not $deviceGroup) {
            $deviceGroup = New-VmsDeviceGroup -ParentGroup $parentFolder -Name $model
        }

        foreach ($camera in $group.Group) {
            # Verifier si la camera est deja membre du groupe pour eviter les doublons
            $alreadyMember = Get-VmsDeviceGroupMember -Group $deviceGroup -ErrorAction SilentlyContinue |
                Where-Object { $_.Id -eq $camera.Id }
            if (-not $alreadyMember) {
                Add-VmsDeviceGroupMember -Group $deviceGroup -DeviceId $camera.Id
            }
        }

        & $Log "Modele '$model' : $($group.Count) camera(s) ajoutee(s)."
    }

    if (-not (& $Cancel)) {
        & $Log "Organisation par modele terminee."
    }
}
