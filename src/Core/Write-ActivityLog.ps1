<#
.SYNOPSIS
    Systeme de logging centralise pour l'application.
.DESCRIPTION
    Ecrit les messages dans un fichier de log et/ou vers un callback UI.
    Format : [HH:mm:ss] [LEVEL] Message
#>

function Write-ActivityLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO',

        [Parameter()]
        [string]$LogDirectory,

        [Parameter()]
        [scriptblock]$UICallback
    )

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $logLine  = "[$timestamp] [$Level] $Message"

    # Ecriture vers le fichier de log
    if ($LogDirectory) {
        if (-not (Test-Path $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }
        $logFile = Join-Path $LogDirectory ("MilestoneToolkit_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd'))
        Add-Content -Path $logFile -Value $logLine -Encoding UTF8
    }

    # Callback vers l'UI
    if ($UICallback) {
        & $UICallback $logLine
    }
}
