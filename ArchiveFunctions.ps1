function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$LogPath
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $entry = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message

    switch ($Level) {
        'ERROR' { Write-Error $entry }
        'WARN'  { Write-Warning $entry }
        'DEBUG' { Write-Verbose $entry }
        default { Write-Verbose $entry }
    }

    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        try {
            $logFolder = Split-Path -Path $LogPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($logFolder) -and -not (Test-Path -LiteralPath $logFolder)) {
                New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
            }

            Add-Content -Path $LogPath -Value $entry -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file '$LogPath'. $($_.Exception.Message)"
        }
    }
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 2,

        [Parameter(Mandatory = $false)]
        [string]$OperationName = 'Operation',

        [Parameter(Mandatory = $false)]
        [string]$LogPath
    )

    $attempt = 0
    do {
        $attempt++
        try {
            Write-LogMessage -Level DEBUG -Message "$OperationName attempt $attempt of $RetryCount." -LogPath $LogPath
            return & $ScriptBlock
        }
        catch {
            $msg = "$OperationName failed on attempt $attempt of $RetryCount. $($_.Exception.Message)"
            if ($attempt -ge $RetryCount) {
                Write-LogMessage -Level ERROR -Message $msg -LogPath $LogPath
                throw
            }

            Write-LogMessage -Level WARN -Message "$msg Retrying in $RetryDelaySeconds second(s)." -LogPath $LogPath
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    } while ($attempt -lt $RetryCount)
}

Function Zip-Yesterday {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceFolder,

        [Parameter(Mandatory = $True, ValueFromPipeline = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetFolder,

        [Parameter(Mandatory = $False, ValueFromPipeline = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$Extension = '*.log',

        [Parameter(Mandatory = $False)]
        [int]$RetryCount = 3,

        [Parameter(Mandatory = $False)]
        [int]$RetryDelaySeconds = 2,

        [Parameter(Mandatory = $False)]
        [string]$LogPath,

        [Parameter(Mandatory = $False)]
        [switch]$Recurse
    )

    begin {
        try {
            Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction Stop
        }
        catch {
            throw "Failed to load System.IO.Compression.FileSystem. $($_.Exception.Message)"
        }
    }

    process {
        $tempFolderName = $null
        $zipCreated = $false
        $files = @()

        try {
            Write-LogMessage -Level INFO -Message "Starting Zip-Yesterday. SourceFolder='$SourceFolder', TargetFolder='$TargetFolder', Extension='$Extension'." -LogPath $LogPath

            if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
                throw "SourceFolder does not exist or is not a folder: $SourceFolder"
            }

            if (-not (Test-Path -LiteralPath $TargetFolder -PathType Container)) {
                throw "TargetFolder does not exist or is not a folder: $TargetFolder"
            }

            [datetime]$yesterdayStart = (Get-Date).Date.AddDays(-1)
            [datetime]$todayStart     = (Get-Date).Date

            Write-LogMessage -Level DEBUG -Message "Selecting files with LastWriteTime >= '$yesterdayStart' and < '$todayStart'." -LogPath $LogPath

            $gciParams = @{
                Path        = $SourceFolder
                Filter      = $Extension
                File        = $true
                ErrorAction = 'Stop'
            }

            if ($Recurse.IsPresent) {
                $gciParams['Recurse'] = $true
            }

            $files = @(Get-ChildItem @gciParams | Where-Object {
                $_.LastWriteTime -ge $yesterdayStart -and $_.LastWriteTime -lt $todayStart
            })

            if (-not $files -or $files.Count -eq 0) {
                Write-LogMessage -Level INFO -Message "No files found in '$SourceFolder' matching '$Extension' written yesterday." -LogPath $LogPath

                [pscustomobject]@{
                    FunctionName = 'Zip-Yesterday'
                    SourceFolder = $SourceFolder
                    TargetFolder = $TargetFolder
                    Extension    = $Extension
                    FileCount    = 0
                    ZipPath      = $null
                    Success      = $true
                    Action       = 'NoFilesFound'
                    StartRange   = $yesterdayStart
                    EndRange     = $todayStart
                }
                return
            }

            Write-LogMessage -Level INFO -Message "File(s) selected for ZIP: $($files.Count)" -LogPath $LogPath
            Write-LogMessage -Level DEBUG -Message ($files | Select-Object FullName, Length, LastWriteTime | Out-String -Width 200) -LogPath $LogPath

            $tempFolderName = Join-Path -Path $env:TEMP -ChildPath ("Zip-Yesterday_{0}_{1}" -f (Get-Date -Format 'yyyyMMddHHmmssfff'), [guid]::NewGuid().ToString('N'))
            $destination = Join-Path -Path $TargetFolder -ChildPath ("LogFiles-{0}.zip" -f $yesterdayStart.ToString('yyyyMMdd'))

            Write-LogMessage -Level DEBUG -Message "Temp folder: $tempFolderName" -LogPath $LogPath
            Write-LogMessage -Level DEBUG -Message "Destination ZIP: $destination" -LogPath $LogPath

            if (Test-Path -LiteralPath $destination) {
                Write-LogMessage -Level WARN -Message "ZIP '$destination' already exists. Assuming function already ran for this day. No action taken." -LogPath $LogPath

                [pscustomobject]@{
                    FunctionName = 'Zip-Yesterday'
                    SourceFolder = $SourceFolder
                    TargetFolder = $TargetFolder
                    Extension    = $Extension
                    FileCount    = $files.Count
                    ZipPath      = $destination
                    Success      = $true
                    Action       = 'ZipAlreadyExists'
                    StartRange   = $yesterdayStart
                    EndRange     = $todayStart
                }
                return
            }

            if ($PSCmdlet.ShouldProcess($tempFolderName, 'Create temporary folder for ZIP staging')) {
                New-Item -Path $tempFolderName -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }

            Write-LogMessage -Level INFO -Message "Copying files to temp folder before compression." -LogPath $LogPath

            foreach ($file in $files) {
                $destinationFile = Join-Path -Path $tempFolderName -ChildPath $file.Name

                if ($PSCmdlet.ShouldProcess($file.FullName, "Copy to temp folder '$tempFolderName'")) {
                    Invoke-WithRetry -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName "Copy '$($file.FullName)'" -LogPath $LogPath -ScriptBlock {
                        Copy-Item -LiteralPath $file.FullName -Destination $destinationFile -Force -ErrorAction Stop
                    }
                }
            }

            if ($PSCmdlet.ShouldProcess($destination, 'Create ZIP archive from temp folder')) {
                Invoke-WithRetry -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName "Create ZIP '$destination'" -LogPath $LogPath -ScriptBlock {
                    if (Test-Path -LiteralPath $destination) {
                        Remove-Item -LiteralPath $destination -Force -ErrorAction Stop
                    }

                    [System.IO.Compression.ZipFile]::CreateFromDirectory(
                        $tempFolderName,
                        $destination,
                        [System.IO.Compression.CompressionLevel]::Optimal,
                        $false
                    )
                }
            }

            if (-not (Test-Path -LiteralPath $destination -PathType Leaf)) {
                throw "ZIP creation did not produce the expected file: $destination"
            }

            $zipItem = Get-Item -LiteralPath $destination -ErrorAction Stop
            if ($zipItem.Length -le 0) {
                throw "ZIP file exists but is empty: $destination"
            }

            $zipCreated = $true
            Write-LogMessage -Level INFO -Message "ZIP created successfully: '$destination' ($($zipItem.Length) bytes)." -LogPath $LogPath

            Write-LogMessage -Level INFO -Message "Removing original files after successful ZIP creation." -LogPath $LogPath

            foreach ($file in $files) {
                if ($PSCmdlet.ShouldProcess($file.FullName, 'Delete original file after successful ZIP verification')) {
                    Invoke-WithRetry -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName "Delete original '$($file.FullName)'" -LogPath $LogPath -ScriptBlock {
                        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    }
                }
            }

            [pscustomobject]@{
                FunctionName = 'Zip-Yesterday'
                SourceFolder = $SourceFolder
                TargetFolder = $TargetFolder
                Extension    = $Extension
                FileCount    = $files.Count
                ZipPath      = $destination
                ZipSizeBytes = $zipItem.Length
                Success      = $true
                Action       = 'ZipCreated'
                StartRange   = $yesterdayStart
                EndRange     = $todayStart
            }
        }
        catch {
            Write-LogMessage -Level ERROR -Message "Zip-Yesterday failed. $($_.Exception.Message)" -LogPath $LogPath
            throw
        }
        finally {
            if (-not [string]::IsNullOrWhiteSpace($tempFolderName) -and (Test-Path -LiteralPath $tempFolderName)) {
                try {
                    if ($PSCmdlet.ShouldProcess($tempFolderName, 'Remove temporary folder')) {
                        Remove-Item -LiteralPath $tempFolderName -Force -Recurse -ErrorAction Stop
                        Write-LogMessage -Level DEBUG -Message "Removed temp folder '$tempFolderName'." -LogPath $LogPath
                    }
                }
                catch {
                    Write-LogMessage -Level WARN -Message "Failed to remove temp folder '$tempFolderName'. $($_.Exception.Message)" -LogPath $LogPath
                }
            }
        }
    }
} # END ZIP YESTERDAY FUNCTION

Function Purge-OldZips {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetFolder,

        [Parameter(Mandatory = $False, ValueFromPipeline = $true, Position = 1)]
        [ValidateRange(0, 36500)]
        [int]$OlderThan = 5,

        [Parameter(Mandatory = $False)]
        [int]$RetryCount = 3,

        [Parameter(Mandatory = $False)]
        [int]$RetryDelaySeconds = 2,

        [Parameter(Mandatory = $False)]
        [string]$LogPath
    )

    process {
        try {
            Write-LogMessage -Level INFO -Message "Starting Purge-OldZips. TargetFolder='$TargetFolder', OlderThan='$OlderThan'." -LogPath $LogPath

            if (-not (Test-Path -LiteralPath $TargetFolder -PathType Container)) {
                throw "TargetFolder does not exist or is not a folder: $TargetFolder"
            }

            [datetime]$cutoffDate = (Get-Date).Date.AddDays(-$OlderThan)
            Write-LogMessage -Level DEBUG -Message "Selecting ZIP files with LastWriteTime < '$cutoffDate'." -LogPath $LogPath

            $files = @(Get-ChildItem -Path $TargetFolder -Filter '*.zip' -File -ErrorAction Stop | Where-Object {
                $_.LastWriteTime -lt $cutoffDate
            })

            if (-not $files -or $files.Count -eq 0) {
                Write-LogMessage -Level INFO -Message "No ZIP files found in '$TargetFolder' older than $OlderThan day(s)." -LogPath $LogPath

                [pscustomobject]@{
                    FunctionName = 'Purge-OldZips'
                    TargetFolder = $TargetFolder
                    OlderThan    = $OlderThan
                    PurgedCount  = 0
                    Success      = $true
                    Action       = 'NoFilesFound'
                    CutoffDate   = $cutoffDate
                }
                return
            }

            Write-LogMessage -Level INFO -Message "ZIP file(s) selected for purge: $($files.Count)" -LogPath $LogPath
            Write-LogMessage -Level DEBUG -Message ($files | Select-Object FullName, Length, LastWriteTime | Out-String -Width 200) -LogPath $LogPath

            $purgedCount = 0

            foreach ($file in $files) {
                if ($PSCmdlet.ShouldProcess($file.FullName, 'Delete old ZIP file')) {
                    Invoke-WithRetry -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName "Delete ZIP '$($file.FullName)'" -LogPath $LogPath -ScriptBlock {
                        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    }

                    $purgedCount++
                    Write-LogMessage -Level INFO -Message "Deleted ZIP file '$($file.FullName)'." -LogPath $LogPath
                }
            }

            [pscustomobject]@{
                FunctionName = 'Purge-OldZips'
                TargetFolder = $TargetFolder
                OlderThan    = $OlderThan
                PurgedCount  = $purgedCount
                Success      = $true
                Action       = 'FilesPurged'
                CutoffDate   = $cutoffDate
            }
        }
        catch {
            Write-LogMessage -Level ERROR -Message "Purge-OldZips failed. $($_.Exception.Message)" -LogPath $LogPath
            throw
        }
    }
} # END PURGE OLD ZIPS FUNCTION

Function Invoke-LogArchivalJob {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceFolder,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetFolder,

        [Parameter(Mandatory = $false)]
        [string]$Extension = '*.log',

        [Parameter(Mandatory = $false)]
        [int]$RetentionDays = 5,

        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 2,

        [Parameter(Mandatory = $false)]
        [string]$LogPath,

        [Parameter(Mandatory = $false)]
        [switch]$Recurse
    )

    $jobStart = Get-Date

    Write-LogMessage -Level INFO -Message "===== Starting Log Archival Job =====" -LogPath $LogPath
    Write-LogMessage -Level INFO -Message "SourceFolder='$SourceFolder' TargetFolder='$TargetFolder' RetentionDays='$RetentionDays'" -LogPath $LogPath

    $zipResult = $null
    $purgeResult = $null

    try {

        # --- STEP 1: ZIP YESTERDAY ---
        if ($PSCmdlet.ShouldProcess("Zip-Yesterday", "Archive yesterday's logs")) {
            Write-LogMessage -Level INFO -Message "Step 1: Running Zip-Yesterday" -LogPath $LogPath

            $zipParams = @{
                SourceFolder         = $SourceFolder
                TargetFolder         = $TargetFolder
                Extension            = $Extension
                RetryCount           = $RetryCount
                RetryDelaySeconds    = $RetryDelaySeconds
                LogPath              = $LogPath
                ErrorAction          = 'Stop'
            }

            if ($Recurse.IsPresent) {
                $zipParams['Recurse'] = $true
            }

            $zipResult = Zip-Yesterday @zipParams -Verbose:$VerbosePreference
        }

        # --- STEP 2: PURGE OLD ZIPS ---
        if ($PSCmdlet.ShouldProcess("Purge-OldZips", "Clean up old archives")) {
            Write-LogMessage -Level INFO -Message "Step 2: Running Purge-OldZips" -LogPath $LogPath

            $purgeResult = Purge-OldZips `
                -TargetFolder $TargetFolder `
                -OlderThan $RetentionDays `
                -RetryCount $RetryCount `
                -RetryDelaySeconds $RetryDelaySeconds `
                -LogPath $LogPath `
                -ErrorAction Stop `
                -Verbose:$VerbosePreference
        }

        $jobEnd = Get-Date
        $duration = [math]::Round(($jobEnd - $jobStart).TotalSeconds, 2)

        Write-LogMessage -Level INFO -Message "===== Log Archival Job Completed Successfully in $duration second(s) =====" -LogPath $LogPath

        # --- CONSOLIDATED OUTPUT ---
        return [pscustomobject]@{
            JobName            = 'LogArchival'
            StartTime          = $jobStart
            EndTime            = $jobEnd
            DurationSeconds    = $duration

            SourceFolder       = $SourceFolder
            TargetFolder       = $TargetFolder
            Extension          = $Extension
            RetentionDays      = $RetentionDays

            ZipFileCount       = $zipResult.FileCount
            ZipPath            = $zipResult.ZipPath
            ZipAction          = $zipResult.Action

            PurgedFileCount    = $purgeResult.PurgedCount
            PurgeAction        = $purgeResult.Action

            OverallSuccess     = $true
        }
    }
    catch {
        $jobEnd = Get-Date
        $duration = [math]::Round(($jobEnd - $jobStart).TotalSeconds, 2)

        Write-LogMessage -Level ERROR -Message "Log Archival Job FAILED. $($_.Exception.Message)" -LogPath $LogPath

        return [pscustomobject]@{
            JobName            = 'LogArchival'
            StartTime          = $jobStart
            EndTime            = $jobEnd
            DurationSeconds    = $duration

            SourceFolder       = $SourceFolder
            TargetFolder       = $TargetFolder
            Extension          = $Extension
            RetentionDays      = $RetentionDays

            ZipFileCount       = if ($zipResult) { $zipResult.FileCount } else { $null }
            ZipPath            = if ($zipResult) { $zipResult.ZipPath } else { $null }
            ZipAction          = if ($zipResult) { $zipResult.Action } else { 'Failed' }

            PurgedFileCount    = if ($purgeResult) { $purgeResult.PurgedCount } else { $null }
            PurgeAction        = if ($purgeResult) { $purgeResult.Action } else { 'NotExecuted' }

            OverallSuccess     = $false
            Error              = $_.Exception.Message
        }

        throw
    }
}