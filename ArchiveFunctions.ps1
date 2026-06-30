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
            Add-Type -AssemblyName 'System.IO.Compression' -ErrorAction Stop
            Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction Stop
        }
        catch {
            throw "Failed to load ZIP compression assemblies. $($_.Exception.Message)"
        }
    }

    process {
        try {
            Write-LogMessage -Level INFO -Message "Starting Zip-Yesterday. SourceFolder='$SourceFolder', TargetFolder='$TargetFolder', Extension='$Extension'." -LogPath $LogPath

            if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
                throw "SourceFolder does not exist or is not a folder: $SourceFolder"
            }

            if (-not (Test-Path -LiteralPath $TargetFolder -PathType Container)) {
                throw "TargetFolder does not exist or is not a folder: $TargetFolder"
            }

            [datetime]$todayStart = (Get-Date).Date

            Write-LogMessage -Level DEBUG -Message "Selecting files with LastWriteTime earlier than '$todayStart'. Today's files will not be archived." -LogPath $LogPath

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
                $_.LastWriteTime -lt $todayStart
            })

            if (-not $files -or $files.Count -eq 0) {
                Write-LogMessage -Level INFO -Message "No files found in '$SourceFolder' matching '$Extension' older than today." -LogPath $LogPath

                return [pscustomobject]@{
                    FunctionName    = 'Zip-Yesterday'
                    SourceFolder    = $SourceFolder
                    TargetFolder    = $TargetFolder
                    Extension       = $Extension
                    FileCount       = 0
                    ZipPath         = $null
                    ZipSizeBytes    = 0
                    Success         = $true
                    Action          = 'NoFilesFound'
                    ArchivedDates   = @()
                    SkippedDates    = @()
                    SkippedFileCount = 0
                    ArchiveResult   = @()
                    SelectionCutoff = $todayStart
                }
            }

            Write-LogMessage -Level INFO -Message "File(s) selected for ZIP: $($files.Count)" -LogPath $LogPath
            Write-LogMessage -Level DEBUG -Message ($files | Select-Object FullName, Length, LastWriteTime | Out-String -Width 200) -LogPath $LogPath

            $sourceRootItem = Get-Item -LiteralPath $SourceFolder -ErrorAction Stop
            $sourceRoot = $sourceRootItem.FullName.TrimEnd('\')

            $groups = @(
                $files |
                    Group-Object { $_.LastWriteTime.Date } |
                    Sort-Object { [datetime]$_.Name }
            )

            $archiveResults = New-Object System.Collections.Generic.List[object]
            $skippedResults = New-Object System.Collections.Generic.List[object]

            $totalArchivedFiles = 0
            $totalZipSizeBytes = 0
            $totalSkippedFiles = 0

            foreach ($group in $groups) {
                [datetime]$archiveDate = [datetime]$group.Name
                $dateText = $archiveDate.ToString('yyyyMMdd')
                $destination = Join-Path -Path $TargetFolder -ChildPath ("LogFiles-{0}.zip" -f $dateText)

                if (Test-Path -LiteralPath $destination) {
                    Write-LogMessage `
                        -Level WARN `
                        -Message ("ZIP '{0}' already exists for {1}. Skipping {2} log file(s)." -f `
                            $destination,
                            $archiveDate.ToString('yyyy-MM-dd'),
                            $group.Group.Count) `
                        -LogPath $LogPath

                    $totalSkippedFiles += $group.Group.Count

                    $skippedResults.Add([pscustomobject]@{
                        Date         = $archiveDate
                        ZipPath      = $destination
                        FilesSkipped = $group.Group.Count
                        Success      = $true
                        Action       = 'ZipAlreadyExists'
                    }) | Out-Null

                    continue
                }

                Write-LogMessage -Level INFO -Message "Processing archive date '$($archiveDate.ToString('yyyy-MM-dd'))'. File count: $($group.Group.Count). Destination ZIP: '$destination'." -LogPath $LogPath

                $zipArchive = $null
                $filesArchivedForDate = 0
                $filesDeletedForDate = 0

                try {
                    if ($PSCmdlet.ShouldProcess($destination, "Create ZIP archive for $dateText")) {
                        $zipArchive = Invoke-WithRetry `
                            -RetryCount $RetryCount `
                            -RetryDelaySeconds $RetryDelaySeconds `
                            -OperationName "Open ZIP '$destination'" `
                            -LogPath $LogPath `
                            -ScriptBlock ({
                                [System.IO.Compression.ZipFile]::Open(
                                    $destination,
                                    [System.IO.Compression.ZipArchiveMode]::Update
                                )
                            }.GetNewClosure())

                        foreach ($file in $group.Group) {
                            $entryName = $file.FullName

                            if ($entryName.StartsWith($sourceRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                                $entryName = $entryName.Substring($sourceRoot.Length).TrimStart('\')
                            }
                            else {
                                $entryName = $file.Name
                            }

                            $entryName = $entryName -replace '\\', '/'

                            Write-LogMessage -Level DEBUG -Message "Adding ZIP entry '$entryName' from '$($file.FullName)'." -LogPath $LogPath

                            Invoke-WithRetry `
                                -RetryCount $RetryCount `
                                -RetryDelaySeconds $RetryDelaySeconds `
                                -OperationName "Add '$($file.FullName)' to ZIP '$destination'" `
                                -LogPath $LogPath `
                                -ScriptBlock ({
                                    $existingEntry = $zipArchive.GetEntry($entryName)

                                    if ($null -ne $existingEntry) {
                                        $existingEntry.Delete()
                                    }

                                    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                                        $zipArchive,
                                        $file.FullName,
                                        $entryName,
                                        [System.IO.Compression.CompressionLevel]::Optimal
                                    ) | Out-Null
                                }.GetNewClosure())

                            $filesArchivedForDate++
                        }
                    }
                }
                finally {
                    if ($null -ne $zipArchive) {
                        $zipArchive.Dispose()
                        $zipArchive = $null
                    }
                }

                if (-not (Test-Path -LiteralPath $destination -PathType Leaf)) {
                    throw "ZIP creation did not produce the expected file: $destination"
                }

                $zipItem = Get-Item -LiteralPath $destination -ErrorAction Stop

                if ($zipItem.Length -le 0) {
                    throw "ZIP file exists but is empty: $destination"
                }

                Write-LogMessage -Level INFO -Message "ZIP verified successfully: '$destination' ($($zipItem.Length) bytes)." -LogPath $LogPath
                Write-LogMessage -Level INFO -Message "Removing original files for archive date '$($archiveDate.ToString('yyyy-MM-dd'))' after successful ZIP verification." -LogPath $LogPath

                foreach ($file in $group.Group) {
                    if ($PSCmdlet.ShouldProcess($file.FullName, 'Delete original file after successful ZIP verification')) {
                        Invoke-WithRetry `
                            -RetryCount $RetryCount `
                            -RetryDelaySeconds $RetryDelaySeconds `
                            -OperationName "Delete original '$($file.FullName)'" `
                            -LogPath $LogPath `
                            -ScriptBlock ({
                                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                            }.GetNewClosure())

                        $filesDeletedForDate++
                    }
                }

                $totalArchivedFiles += $filesArchivedForDate
                $totalZipSizeBytes += $zipItem.Length

                $archiveResults.Add([pscustomobject]@{
                    Date          = $archiveDate
                    ZipPath       = $destination
                    ZipSizeBytes  = $zipItem.Length
                    FilesArchived = $filesArchivedForDate
                    FilesDeleted  = $filesDeletedForDate
                    Success       = $true
                    Action        = 'ZipCreated'
                }) | Out-Null
            }

            Write-LogMessage -Level INFO -Message "Zip-Yesterday completed. Archived $totalArchivedFiles file(s) across $($archiveResults.Count) ZIP file(s). Skipped $totalSkippedFiles file(s) because matching ZIP file(s) already existed." -LogPath $LogPath

            $finalAction = if ($archiveResults.Count -gt 0 -and $skippedResults.Count -gt 0) {
                'ZipCreatedOrSkipped'
            }
            elseif ($archiveResults.Count -gt 0) {
                'ZipCreated'
            }
            elseif ($skippedResults.Count -gt 0) {
                'AllMatchingZipsAlreadyExist'
            }
            else {
                'NoActionTaken'
            }

            #Write-Host "archiveResults count = $($archiveResults.Count)"
            #Write-Host "skippedResults count = $($skippedResults.Count)"
            #Write-Host "finalAction = $finalAction"

            return [pscustomobject]@{
                FunctionName     = 'Zip-Yesterday'
                FileCount        = $totalArchivedFiles
                ZipPath          = @()
                ZipSizeBytes     = $totalZipSizeBytes
                Success          = $true
                Action           = $finalAction
                SkippedFileCount = $totalSkippedFiles
                SkippedDates     = @()
            }

            <#
            return [pscustomobject]@{
                FunctionName     = 'Zip-Yesterday'
                SourceFolder     = $SourceFolder
                TargetFolder     = $TargetFolder
                Extension        = $Extension
                FileCount        = $totalArchivedFiles
                ZipPath          = @($archiveResults | ForEach-Object { $_.ZipPath })
                ZipSizeBytes     = $totalZipSizeBytes
                Success          = $true
                Action           = $finalAction
                ArchivedDates    = @($archiveResults | ForEach-Object { $_.Date })
                SkippedDates     = @($skippedResults | ForEach-Object { $_.Date })
                SkippedFileCount = $totalSkippedFiles
                ArchiveResult    = @($archiveResults)
                SkippedResult    = @($skippedResults)
                SelectionCutoff  = $todayStart
            } #>
        }
        catch {
            Write-LogMessage -Level ERROR -Message "Zip-Yesterday failed. $($_.Exception.Message)" -LogPath $LogPath
            throw
        }
    }
}

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

            Write-LogMessage -Level DEBUG -Message "Selecting ZIP files with LastWriteTime older than '$cutoffDate'." -LogPath $LogPath

            $files = @(Get-ChildItem -Path $TargetFolder -Filter '*.zip' -File -ErrorAction Stop | Where-Object {
                $_.LastWriteTime -lt $cutoffDate
            })

            if (-not $files -or $files.Count -eq 0) {
                Write-LogMessage -Level INFO -Message "No ZIP files found in '$TargetFolder' older than $OlderThan day(s)." -LogPath $LogPath

                return [pscustomobject]@{
                    FunctionName = 'Purge-OldZips'
                    TargetFolder = $TargetFolder
                    OlderThan    = $OlderThan
                    PurgedCount  = 0
                    Success      = $true
                    Action       = 'NoFilesFound'
                    CutoffDate   = $cutoffDate
                }
            }

            Write-LogMessage -Level INFO -Message "ZIP file(s) selected for purge: $($files.Count)" -LogPath $LogPath
            Write-LogMessage -Level DEBUG -Message ($files | Select-Object FullName, Length, LastWriteTime | Out-String -Width 200) -LogPath $LogPath

            $purgedCount = 0

            foreach ($file in $files) {
                if ($PSCmdlet.ShouldProcess($file.FullName, 'Delete old ZIP file')) {
                    Invoke-WithRetry `
                        -RetryCount $RetryCount `
                        -RetryDelaySeconds $RetryDelaySeconds `
                        -OperationName "Delete ZIP '$($file.FullName)'" `
                        -LogPath $LogPath `
                        -ScriptBlock ({
                            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                        }.GetNewClosure())

                    $purgedCount++
                    Write-LogMessage -Level INFO -Message "Deleted ZIP file '$($file.FullName)'." -LogPath $LogPath
                }
            }

            return [pscustomobject]@{
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
}

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
        if ($PSCmdlet.ShouldProcess("Zip-Yesterday", "Archive log files older than today")) {
            Write-LogMessage -Level INFO -Message "Step 1: Running Zip-Yesterday to archive log files older than today" -LogPath $LogPath

            $zipParams = @{
                SourceFolder      = $SourceFolder
                TargetFolder      = $TargetFolder
                Extension         = $Extension
                RetryCount        = $RetryCount
                RetryDelaySeconds = $RetryDelaySeconds
                LogPath           = $LogPath
                ErrorAction       = 'Stop'
            }

            if ($Recurse.IsPresent) {
                $zipParams['Recurse'] = $true
            }

            $zipResult = Zip-Yesterday @zipParams -Verbose:$VerbosePreference
        }

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

        return [pscustomobject]@{
            JobName          = 'LogArchival'
            StartTime        = $jobStart
            EndTime          = $jobEnd
            DurationSeconds  = $duration

            SourceFolder     = $SourceFolder
            TargetFolder     = $TargetFolder
            Extension        = $Extension
            RetentionDays    = $RetentionDays

            ZipFileCount     = $zipResult.FileCount
            ZipPath          = $zipResult.ZipPath
            ZipAction        = $zipResult.Action
            SkippedFileCount = $zipResult.SkippedFileCount
            SkippedDates     = $zipResult.SkippedDates

            PurgedFileCount  = $purgeResult.PurgedCount
            PurgeAction      = $purgeResult.Action

            OverallSuccess   = $true
        }
    }
    catch {
        $jobEnd = Get-Date
        $duration = [math]::Round(($jobEnd - $jobStart).TotalSeconds, 2)

        Write-LogMessage -Level ERROR -Message "Log Archival Job FAILED. $($_.Exception.Message)" -LogPath $LogPath

        return [pscustomobject]@{
            JobName          = 'LogArchival'
            StartTime        = $jobStart
            EndTime          = $jobEnd
            DurationSeconds  = $duration

            SourceFolder     = $SourceFolder
            TargetFolder     = $TargetFolder
            Extension        = $Extension
            RetentionDays    = $RetentionDays

            ZipFileCount     = if ($zipResult) { $zipResult.FileCount } else { $null }
            ZipPath          = if ($zipResult) { $zipResult.ZipPath } else { $null }
            ZipAction        = if ($zipResult) { $zipResult.Action } else { 'Failed' }
            SkippedFileCount = if ($zipResult) { $zipResult.SkippedFileCount } else { $null }
            SkippedDates     = if ($zipResult) { $zipResult.SkippedDates } else { $null }

            PurgedFileCount  = if ($purgeResult) { $purgeResult.PurgedCount } else { $null }
            PurgeAction      = if ($purgeResult) { $purgeResult.Action } else { 'NotExecuted' }

            OverallSuccess   = $false
            Error            = $_.Exception.Message
        }

        throw
    }
}