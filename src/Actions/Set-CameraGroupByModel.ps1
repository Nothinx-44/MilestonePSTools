function Set-CameraGroupByModel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$Config,
        [Parameter(Mandatory)] [scriptblock]$Log,
        [Parameter()] [scriptblock]$Cancel = { $false },
        [Parameter()] [scriptblock]$ReportProgress = {}
    )

    $parentFolderName = $script:T.GM_ParentFolder

    & $Log $script:T.GM_LogRetrieving
    $cameras = Get-VmsCameraReport
    & $Log ($script:T.GM_LogFound -f $cameras.Count)

    $parentFolder = Get-VmsDeviceGroup -Name $parentFolderName -ErrorAction SilentlyContinue
    if (-not $parentFolder) {
        $parentFolder = New-VmsDeviceGroup -Name $parentFolderName
        & $Log ($script:T.GM_LogParentCreated -f $parentFolderName)
    }

    $camerasByModel = $cameras | Group-Object -Property Model
    $total          = $camerasByModel.Count
    & $Log ($script:T.GM_LogModels -f $total)

    $done = 0
    foreach ($group in $camerasByModel) {
        if (& $Cancel) { & $Log ($script:T.GM_LogCancelled -f $done, $total) ; break }

        $done++
        & $ReportProgress $done $total

        $model = if ([string]::IsNullOrWhiteSpace($group.Name)) { $script:T.GM_Unknown } else { $group.Name }

        $deviceGroup = Get-VmsDeviceGroup -ParentGroup $parentFolder -Name $model -ErrorAction SilentlyContinue
        if (-not $deviceGroup) {
            $deviceGroup = New-VmsDeviceGroup -ParentGroup $parentFolder -Name $model
        }

        foreach ($camera in $group.Group) {
            $alreadyMember = Get-VmsDeviceGroupMember -Group $deviceGroup -ErrorAction SilentlyContinue |
                Where-Object { $_.Id -eq $camera.Id }
            if (-not $alreadyMember) {
                Add-VmsDeviceGroupMember -Group $deviceGroup -DeviceId $camera.Id
            }
        }

        & $Log ($script:T.GM_LogModel -f $model, $group.Count)
    }

    if (-not (& $Cancel)) { & $Log $script:T.GM_LogDone }
}
