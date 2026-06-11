param(
    [switch]$Schedule,
    [string]$TaskName = "Log Archival Job",
    [string]$ScheduleType = "DAILY",   # DAILY, MINUTE, HOURLY
    [string]$StartTime = "01:00",
    [int]$Interval = 30,
    [string]$SourceFolder = "C:\Logs",
    [string]$TargetFolder = "D:\Archive",
    [string]$Extension = "*.log",
    [int]$RetentionDays = 7,
    [string]$LogPath = "C:\Logs\ArchiveJob.log"
)

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $currentIdentity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if ($Schedule)
{
    $scriptPath = $MyInvocation.MyCommand.Definition
    $psExe = Join-Path $PSHOME "powershell.exe"

    # Build argument list (preserve all params + ensure re-entry with -Schedule)
    $argList = @(
        "-NoProfile"
        "-ExecutionPolicy Bypass"
        "-File `"$scriptPath`""
        "-Schedule"
        "-TaskName `"$TaskName`""
        "-ScheduleType `"$ScheduleType`""
        "-StartTime `"$StartTime`""
        "-Interval $Interval"
        "-SourceFolder `"$SourceFolder`""
        "-TargetFolder `"$TargetFolder`""
        "-Extension `"$Extension`""
        "-RetentionDays $RetentionDays"
        "-LogPath `"$LogPath`""
    ) -join ' '

    # --- AUTO-ELEVATION ---
    if (-not (Test-IsAdministrator))
    {
        Write-Host "Not running as Administrator. Relaunching elevated..."

        Start-Process -FilePath $psExe `
            -ArgumentList $argList `
            -Verb RunAs

        exit
    }

    Write-Host "Running in elevated context. Creating/updating scheduled task: $TaskName"

    # --- BUILD SCHEDULE ---
    switch ($ScheduleType.ToUpper()) {
        "MINUTE" {
            $scheduleArgs = "/sc MINUTE /MO $Interval /st $StartTime"
        }
        "HOURLY" {
            $scheduleArgs = "/sc HOURLY /MO $Interval /st $StartTime"
        }
        default {
            $scheduleArgs = "/sc DAILY /st $StartTime"
        }
    }

    # --- CLEAN RECREATE ---
    schtasks.exe /delete /tn "$TaskName" /f 2>$null | Out-Null

    $taskCommand = "`"$psExe`" $argList"

    $createCmd = @"
schtasks.exe /create `
    /tn "$TaskName" `
    /tr "$taskCommand" `
    /RL HIGHEST `
    /F `
    $scheduleArgs
"@

    Write-Host "Registering scheduled task..."
    cmd.exe /c $createCmd

    Write-Host "Scheduled task '$TaskName' created/updated successfully."
    exit
}

# Fail fast + consistent error behavior
$ErrorActionPreference = 'Stop'

# Optional: enforce TLS (useful on hardened servers)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

try {
    Write-Host "Starting Archive Job at $(Get-Date)"

    # OPTIONAL: If functions are stored separately
    # . "C:\Scripts\ArchiveFunctions.ps1"

    $result = Invoke-LogArchivalJob `
        -SourceFolder $SourceFolder `
        -TargetFolder $TargetFolder `
        -Extension $Extension `
        -RetentionDays $RetentionDays `
        -LogPath $LogPath `
        -Verbose

    # Output structured result to console/log capture
    $result | Format-List | Out-String | Write-Host

    if (-not $result.OverallSuccess) {
        Write-Error "Archive job reported failure."
        exit 1
    }

    Write-Host "Archive Job completed successfully."
    exit 0
}
catch {
    Write-Error "Archive Job hard failure: $($_.Exception.Message)"
    exit 2
}