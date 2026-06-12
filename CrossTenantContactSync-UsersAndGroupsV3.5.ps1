<#
.SYNOPSIS
    Synchronizes users and groups from a source Entra ID tenant to mail contacts in a target Exchange Online tenant using Microsoft Graph delta queries.

.DESCRIPTION
    This script performs incremental (delta-based) synchronization of users and groups from a source tenant into Exchange Online mail contacts 
    in a target tenant. It is designed for cross-tenant collaboration scenarios such as GAL synchronization or external directory visibility.

    The script uses Microsoft Graph delta queries with explicit property selection to ensure:
        - Efficient incremental processing
        - No per-object Graph hydration calls
        - AttributeFilters can be evaluated directly from delta payloads
        - First-run correctness without requiring reconciliation

    Core architecture principles:

        - Fully delta-driven main processing pipeline
        - No per-object Graph API calls during sync loops
        - Dynamic Graph $select expansion based on AttributeFilters
        - Safe handling of missing properties with debug visibility
        - Batch-based reconciliation for long-term consistency

.PARAMETER ConfigXmlPath
    Optional path to XML configuration file. Values from XML are merged with parameters.
    Parameters take precedence.

.PARAMETER SourceTenantId
    Source tenant ID (GUID).

.PARAMETER SourceTenantName
    Friendly name used for display/logging.

.PARAMETER SourceClientId
    Graph App Registration Client ID.

.PARAMETER SourceClientSecret
    Graph App client secret.

.PARAMETER TargetOrganization
    Target Exchange Online org (usually *.onmicrosoft.com).

.PARAMETER TargetTenantId
    Target tenant ID (GUID).

.PARAMETER TargetTenantName
    Friendly name of target tenant.

.PARAMETER TargetExoAppId
    App ID for Exchange Online app-only auth.

.PARAMETER TargetExoCertThumbprint
    Certificate thumbprint for EXO auth.

.PARAMETER StateRoot
    Path for state (deltaLink) files.

.PARAMETER LogRoot
    Path for log files.

.PARAMETER IncludeDomains
    Allow-list of email domains.

.PARAMETER ExcludeDomains
    Block-list of email domains.

.PARAMETER AppendSourceTenantToDisplayName
    Appends source tenant name to displayName.

.PARAMETER AllowUpnFallback
    Allows UPN to be used when mail is missing.

.PARAMETER DisableDeletes
    Prevents deletion of contacts.

.PARAMETER TopUsers
    Limits number of processed objects (testing only).

.PARAMETER LogLevel
    Logging verbosity (0–3). Default = 3.
    
        0 = INFO only  
        1 = INFO + ERROR  
        2 = INFO + WARN + ERROR  
        3 = ALL (DEBUG included)

.PARAMETER ReconciliationIntervalHours
    Interval for batch reconciliation.

.PARAMETER ForceReconciliation
    Forces reconciliation immediately.

.PARAMETER SourceObjectType
    User | Group | Both.

.PARAMETER ForceFullSync
    Forces delta reset (User / Group / Both).

.PARAMETER SeedTargetFromSource
    Controls seeding/adoption behavior.

.PARAMETER ForceDisplayNameRefresh
    Forces displayName updates.

.BEHAVIOR

    First Run:
        - No state file required
        - Full dataset returned by delta query
        - AttributeFilters evaluated immediately
        - No reconciliation executed
        - Reconciliation timestamp initialized

    Subsequent Runs:
        - Only delta changes processed
        - Reconciliation enforces long-term consistency

.FILTERING MODEL

    AttributeFilters:
        - Evaluated directly from delta payload
        - No per-object Graph calls required

    Missing Properties:
        - Logged at DEBUG level:
            "FILTER PROPERTY MISSING: ..."

.PERFORMANCE CHARACTERISTICS

        - No per-object Graph API calls
        - Single-pass delta processing
        - Optional EXO bulk preload
        - Batch reconciliation (Graph $batch)
        - Minimal EXO writes

.EXAMPLES

    Standard run:
        .\Sync.ps1 -ConfigXmlPath "C:\Temp\Config.xml"

    Debug run:
        .\Sync.ps1 -ConfigXmlPath "Config.xml" -LogLevel 3

    Force full sync:
        .\Sync.ps1 -ConfigXmlPath "Config.xml" -ForceFullSync Both

    Force reconciliation:
        .\Sync.ps1 -ConfigXmlPath "Config.xml" -ForceReconciliation

.NOTES

    Requirements:
        - ExchangeOnlineManagement module
        - Graph API permissions:
            User.Read.All
            Group.Read.All
            Directory.Read.All (recommended)

    Design:
        - Fully delta-driven processing
        - No Graph hydration
        - First-run correctness without reconciliation
        - Batch reconciliation for convergence

    Limitations:
        - Graph delta may omit selected properties in rare cases
        - Missing properties are logged and skipped (no fallback fetch)

.AUTHOR
    Darryl Kegg

.VERSION
    3.6.0

.CHANGELOG

    2026-06-10 – v3.6.0
        * Eliminated all per-object Graph hydration calls from main sync loops
        * Converted script to fully delta-driven processing model
        * Ensured AttributeFilters apply on first run without reconciliation
        * Added required filter properties (onPremisesExtensionAttributes) to delta $select
        * Fixed group delta query to include dynamic $select expansion
        * Added debug logging for missing Graph properties during filter evaluation
        * Simplified first-run behavior (no hydration, no reconciliation)
        * Improved reconciliation timestamp initialization

    2026-06-10 – v2.6.0
        * Replaced run-count reconciliation with time-based model
        * Added -ForceReconciliation switch
        * Fixed early-exit logic for reconciliation scenarios
        * Improved execution flow consistency

    2026-06-10 – v2.5.0
        * Added LogLevel parameter
        * Implemented unified batch reconciliation engine
        * Added proxyAddresses fallback for groups
        * Improved EXO lookup models and delete handling

    2026-06-10 – v2.4.0
        * Added XML-driven AttributeFilters
        * Implemented nested property evaluation
        * Added dynamic Graph $select expansion

    2026-06-09 – v2.3.0
        * Added write optimization
        * Improved logging and update detection
        * Fixed group delta StrictMode issues

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # Optional external config file. If supplied, values here override/augment parameters.
    [string]$ConfigXmlPath="C:\temp\TenantContactSync\TenantAtoTenantBConfig.xml",

    # ---------- Source tenant / Graph ----------
    # Source tenant name is a FRIENDLY name that could be appended to the end of the target DisplayName
    # when using the switch $AppendSourceTenantToDisplayName = True
    [string]$SourceTenantId,
    [string]$SourceTenantName,
    [string]$SourceClientId,
    [string]$SourceClientSecret,

    # ---------- Target tenant / Exchange Online ----------
    # For Connect-ExchangeOnline -Organization, use the tenant's primary onmicrosoft.com domain
    # or accepted org domain expected by EXO app-only auth.
    [string]$TargetOrganization,
    [string]$TargetTenantId,
    [string]$TargetTenantName,

    # App-only Exchange Online auth (certificate-based)
    [string]$TargetExoAppId,
    [string]$TargetExoCertThumbprint,

    # ---------- Runtime ----------
    [string]$StateRoot = "C:\Temp\TenantContactSync\State",
    [string]$LogRoot   = "C:\Temp\TenantContactSync\Logs",

    # Optional filter/include logic
    [string[]]$IncludeDomains,
    [string[]]$ExcludeDomains,
    [string]$DefaultTargetOU, # placeholder if you later choose hybrid/on-prem target logic

    # Contact display behavior
    [bool]$AppendSourceTenantToDisplayName = $true,

    # If true, users without mail will still be synced using UPN as ExternalEmailAddress
    [bool]$AllowUpnFallback = $true,

    # Safety
    [bool]$DisableDeletes = $false,

    #scoping
    [int]$TopUsers = 0,
    [int]$LogLevel = 3,
    [int]$ReconciliationIntervalHours = 8,
    [switch]$ForceReconciliation,
    [int]$MaxUserResults  = 0,
    [int]$MaxGroupResults = 0,

    # Source object type
    [ValidateSet('User','Group','Both')]
    [string]$SourceObjectType = 'Both',
    [ValidateSet('None','User','Group','Both')]
    [string]$ForceFullSync = 'None',
    [ValidateSet('None','User','Group','Both')]
    [string]$SeedTargetFromSource = 'None',
    [switch]$ForceDisplayNameRefresh

)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# $ForceReconciliation = $true #for testing

# -------------------- Utility / logging --------------------

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

$script:LogLevelMap = @{ 
    ERROR = 1 
    WARN = 2 
    INFO = 0 
    DEBUG = 3 
    }

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )

    # Map log levels to numeric values
    $levelMap = @{
        INFO  = 0
        ERROR = 1
        WARN  = 2
        DEBUG = 3
    }

    # Determine if this message should be emitted
    $emit = $false

    switch ($LogLevel) {
        0 { if ($Level -eq 'INFO') { $emit = $true } }
        1 { if ($Level -in @('INFO','ERROR')) { $emit = $true } }
        2 { if ($Level -in @('INFO','WARN','ERROR')) { $emit = $true } }
        3 { $emit = $true }
    }

    if (-not $emit) {
        return
    }

    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[{0}] [{1}] {2}" -f $ts, $Level, $Message

    Write-Host $line
    Add-Content -LiteralPath $script:LogFile -Value $line
}

function Import-RequiredModule {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed."
    }
    Import-Module $Name -ErrorAction Stop
}

function Get-ConfigFromXml {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config XML not found: $Path"
    }

    [xml]$xml = Get-Content -LiteralPath $Path -Raw

    $attributeFilters = @()

    if ($xml.Configuration.Filters.AttributeFilters.Filter) {
        foreach ($f in @($xml.Configuration.Filters.AttributeFilters.Filter)) {
            if (-not $f.PropertyPath) { continue }
            if (-not $f.Value) { continue }

            $attributeFilters += [pscustomobject]@{
                ObjectType   = if ($f.ObjectType) { ([string]$f.ObjectType).Trim() } else { 'Both' }
                PropertyPath = ([string]$f.PropertyPath).Trim()
                Operator     = if ($f.Operator) { ([string]$f.Operator).Trim() } else { 'Equals' }
                Value        = ([string]$f.Value).Trim()
            }

        }
    }

    return [pscustomobject]@{
        SourceTenantId              = $xml.Configuration.Source.TenantId
        SourceTenantName            = $xml.Configuration.Source.TenantName
        SourceClientId              = $xml.Configuration.Source.ClientId
        SourceClientSecret          = $xml.Configuration.Source.ClientSecret

        TargetOrganization          = $xml.Configuration.Target.Organization
        TargetTenantId              = $xml.Configuration.Target.TenantId
        TargetTenantName            = $xml.Configuration.Target.TenantName
        TargetExoAppId              = $xml.Configuration.Target.ExoAppId
        TargetExoCertThumbprint     = $xml.Configuration.Target.ExoCertThumbprint

        StateRoot                   = $xml.Configuration.Runtime.StateRoot
        LogRoot                     = $xml.Configuration.Runtime.LogRoot

        IncludeDomains              = @($xml.Configuration.Filters.IncludeDomains.Domain | Where-Object { $_ })
        ExcludeDomains              = @($xml.Configuration.Filters.ExcludeDomains.Domain | Where-Object { $_ })
        ServiceAccountPatterns      = @($xml.Configuration.Filters.ServiceAccountPatterns.Pattern | Where-Object { $_ })
        AttributeFilters            = $attributeFilters

        AppendSourceTenantToDisplayName = [System.Convert]::ToBoolean(($xml.Configuration.Runtime.AppendSourceTenantToDisplayName | ForEach-Object { if ($_ -eq $null -or $_ -eq '') { 'true' } else { $_ } }))
        AllowUpnFallback                = [System.Convert]::ToBoolean(($xml.Configuration.Runtime.AllowUpnFallback | ForEach-Object { if ($_ -eq $null -or $_ -eq '') { 'true' } else { $_ } }))
        DisableDeletes                  = [System.Convert]::ToBoolean(($xml.Configuration.Runtime.DisableDeletes | ForEach-Object { if ($_ -eq $null -or $_ -eq '') { 'false' } else { $_ } }))

        SourceObjectType = if ($xml.Configuration.Runtime.SourceObjectType) { $xml.Configuration.Runtime.SourceObjectType } else { 'User' }
        TopUsers = if ($xml.Configuration.Runtime.TopUsers) { [int]$xml.Configuration.Runtime.TopUsers } else { 0 }
    }
}

function Merge-Config {
    param(
        [Parameter(Mandatory)][object]$ParamConfig,
        [Parameter()][object]$XmlConfig
    )

    $merged = [ordered]@{}
    $allKeys = @(
        'SourceTenantId','SourceTenantName','SourceClientId','SourceClientSecret',
        'TargetOrganization','TargetTenantId','TargetTenantName','TargetExoAppId','TargetExoCertThumbprint',
        'StateRoot','LogRoot','IncludeDomains','ExcludeDomains','ServiceAccountPatterns','AttributeFilters',
        'AppendSourceTenantToDisplayName','AllowUpnFallback','DisableDeletes','TopUsers','SourceObjectType'
    )

foreach ($k in $allKeys) {

    $paramValue = $ParamConfig.$k
    $xmlValue   = if ($null -ne $XmlConfig) { $XmlConfig.$k } else { $null }

    $useValue = $null

    # ---------- ARRAYS ----------
    if ($paramValue -is [System.Array]) {

        if ($paramValue -and $paramValue.Count -gt 0) {
            $useValue = $paramValue
        }
        elseif ($xmlValue -and $xmlValue.Count -gt 0) {
            $useValue = $xmlValue
        }
        else {
            $useValue = @()
        }
    }

    # ---------- BOOLEANS (FIXED PRECEDENCE) ----------
    elseif ($paramValue -is [bool]) {

        if ($PSBoundParameters.ContainsKey($k)) {
            # explicitly passed parameter wins
            $useValue = $paramValue
        }
        elseif ($null -ne $paramValue) {
            # parameter default wins over XML (fixes your issue)
            $useValue = $paramValue
        }
        elseif ($null -ne $xmlValue -and ($xmlValue.ToString().Trim()) -ne '') {
            $useValue = [System.Convert]::ToBoolean($xmlValue)
        }
        else {
            $useValue = $false
        }
    }

    # ---------- STRINGS / OTHER ----------
    else {

        if ($null -ne $paramValue -and ($paramValue.ToString().Trim()) -ne '') {
            $useValue = $paramValue
        }
        elseif ($null -ne $xmlValue -and ($xmlValue.ToString().Trim()) -ne '') {
            $useValue = $xmlValue
        }
        else {
            $useValue = $paramValue
        }
    }

    $merged[$k] = $useValue
}

    return [pscustomobject]$merged
}

function Assert-Config {
    param([Parameter(Mandatory)][object]$Config)

    $required = @(
        'SourceTenantId','SourceTenantName','SourceClientId','SourceClientSecret',
        'TargetOrganization','TargetTenantId','TargetTenantName','TargetExoAppId','TargetExoCertThumbprint'
    )

    foreach ($r in $required) {
        if (-not $Config.$r) {
            throw "Missing required configuration value: $r"
        }
    }
}

function Assert-AttributeFilterDesign {
    param([Parameter(Mandatory)][object[]]$AttributeFilters)

    foreach ($filter in @($AttributeFilters)) {

        if (-not $filter) { continue }

        $objectType = if ($filter.ObjectType) { [string]$filter.ObjectType } else { 'Both' }
        $propertyPath = if ($filter.PropertyPath) { [string]$filter.PropertyPath } else { '' }

        if ($objectType -in @('User','Both')) {

            if ($propertyPath -notmatch '^onPremisesExtensionAttributes\.extensionAttribute([1-9]|1[0-5])$') {
                throw "Invalid user AttributeFilter PropertyPath '$propertyPath'. User filters must use onPremisesExtensionAttributes.extensionAttributeX"
            }
        }
    }
}

# -------------------- State --------------------

function Get-StateFilePath {
    param([Parameter(Mandatory)][object]$Config)

    $safeName = "{0}_TO_{1}.json" -f `
        ($Config.SourceTenantId -replace '[^a-zA-Z0-9\-]','_'), `
        ($Config.TargetTenantId -replace '[^a-zA-Z0-9\-]','_')

    return (Join-Path $Config.StateRoot $safeName)
}

function Load-State {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{
            Version     = 2
            UserDeltaLink   = $null
            GroupDeltaLink  = $null
            LastRunUtc  = $null
            LastReconciliationUtc = $null
        }
    }

    $state = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json

    if (-not ($state.PSObject.Properties['UserDeltaLink'])) { $state | Add-Member -NotePropertyName UserDeltaLink -NotePropertyValue $null }

    if (-not ($state.PSObject.Properties['GroupDeltaLink'])) { $state | Add-Member -NotePropertyName GroupDeltaLink -NotePropertyValue $null }

    return $state
}

function Save-State {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$State
    )

    $State.LastRunUtc = [DateTime]::UtcNow.ToString("o")
    $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
}

# -------------------- Graph source side --------------------

function Get-GraphToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    $resp = Invoke-RestMethod `
        -Method Post `
        -Uri $tokenEndpoint `
        -Body $body `
        -ContentType 'application/x-www-form-urlencoded'

    return $resp.access_token
}

function Invoke-GraphJson {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$AccessToken
    )

    $headers = @{
        Authorization    = "Bearer $AccessToken"
        ConsistencyLevel = "eventual"
    }

    try {
        return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers -ErrorAction Stop
    }
    catch {
        $responseBody = $null
        $exceptionMessage = $null

        if ($_.Exception -and $_.Exception.Message) {
            $exceptionMessage = $_.Exception.Message
        }
        else {
            $exceptionMessage = ($_ | Out-String).Trim()
        }

        # Safely extract ErrorDetails
        if ($_.PSObject.Properties['ErrorDetails'] -and $null -ne $_.ErrorDetails) {

            if ($_.ErrorDetails -is [string]) {
                $responseBody = $_.ErrorDetails
            }
            elseif ($_.ErrorDetails.PSObject.Properties['Message']) {
                $responseBody = $_.ErrorDetails.Message
            }
            else {
                $responseBody = ($_.ErrorDetails | Out-String).Trim()
            }
        }

        # Fallback: try to read raw web response body
        if (-not $responseBody -and $_.Exception -and $_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                if ($stream) {
                    $reader = New-Object System.IO.StreamReader($stream)
                    try {
                        $responseBody = $reader.ReadToEnd()
                    }
                    finally {
                        $reader.Dispose()
                        $stream.Dispose()
                    }
                }
            }
            catch {
                # swallow fallback parsing failure
            }
        }

        if (-not $responseBody) {
            $responseBody = "<no response body available>"
        }

        Write-Log "GRAPH GET FAILED: $Uri" "ERROR"
        Write-Log "GRAPH ERROR: $exceptionMessage" "ERROR"
        Write-Log "GRAPH RAW RESPONSE: $responseBody" "ERROR"

        throw
    }
}

function Get-UserDeltaChanges {
    param(
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter()][string]$DeltaLink,
        [Parameter()][object[]]$AttributeFilters,
        [Parameter()][int]$MaxResults = 0
    )

    $baseProps = @(
        'id','displayName','givenName','surname','mail','userPrincipalName',
        'companyName','department','jobTitle','businessPhones','mobilePhone',
        'officeLocation','onPremisesExtensionAttributes'
    )

    $extraProps = Get-TopLevelSelectPropertiesFromAttributeFilters -AttributeFilters $AttributeFilters -ObjectType 'User'
    $allProps   = @($baseProps + $extraProps | Select-Object -Unique)

    $select = [System.Web.HttpUtility]::UrlEncode(($allProps -join ','))

    $uri = if ($DeltaLink) {
        $DeltaLink
    }
    else {
        "https://graph.microsoft.com/v1.0/users/delta?`$select=$select"
    }

    $allChanges = New-Object System.Collections.Generic.List[object]
    $finalDeltaLink = $null

    <# do {
        $page = Invoke-GraphJson -Uri $uri -AccessToken $AccessToken

        $valueProp = $page.PSObject.Properties['value']
        if ($valueProp -and $valueProp.Value) {
            foreach ($item in $valueProp.Value) {
                $allChanges.Add($item)
            }
        }

        $nextLinkProp = $page.PSObject.Properties['@odata.nextLink']
        $deltaLinkProp = $page.PSObject.Properties['@odata.deltaLink']

        if ($nextLinkProp) {
            $uri = $nextLinkProp.Value
        }
        elseif ($deltaLinkProp) {
            $uri = $null
            $finalDeltaLink = $deltaLinkProp.Value
        }
        else {
            Write-Log "WARNING: No nextLink or deltaLink returned by Graph" "WARN"
            $uri = $null
        }
    } while ($uri) #>

    do {
    $page = Invoke-GraphJson -Uri $uri -AccessToken $AccessToken

    $valueProp = $page.PSObject.Properties['value']
    if ($valueProp -and $valueProp.Value) {

        foreach ($item in $valueProp.Value) {

            if ($MaxResults -gt 0 -and $allChanges.Count -ge $MaxResults) {
                Write-Log "MaxResults limit reached ($MaxResults) - stopping delta paging early" "WARN"
                $uri = $null
                break
            }

            $allChanges.Add($item)
        }
    }

    if (-not $uri) { break }

    $nextLinkProp = $page.PSObject.Properties['@odata.nextLink']
    $deltaLinkProp = $page.PSObject.Properties['@odata.deltaLink']

    if ($nextLinkProp -and ($MaxResults -eq 0 -or $allChanges.Count -lt $MaxResults)) {
        $uri = $nextLinkProp.Value
    }
    elseif ($deltaLinkProp) {
        $uri = $null
        $finalDeltaLink = $deltaLinkProp.Value
    }
    else {
        $uri = $null
    }

} while ($uri)

    return [pscustomobject]@{
        Changes   = $allChanges
        DeltaLink = $finalDeltaLink
    }
}

function Get-GroupDeltaChanges {
    param(
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter()][string]$DeltaLink,
        [Parameter()][object[]]$AttributeFilters,
        [Parameter()][int]$MaxResults = 0
    )

    # DO NOT include filter-derived or complex properties in delta select
    $baseProps = @(
        'id','displayName','mail','securityEnabled','mailEnabled','groupTypes', 'proxyAddresses'

    )

    $extraProps = Get-TopLevelSelectPropertiesFromAttributeFilters -AttributeFilters $AttributeFilters -ObjectType 'Group'

    #Strip unsupported properties for GROUP DELTA ONLY
    $extraProps = $extraProps | Where-Object { $_ -ne 'onPremisesExtensionAttributes' }

    $allProps = @(
        $baseProps +
        $extraProps
    ) | Select-Object -Unique

    if ($allProps -contains 'onPremisesExtensionAttributes') {
        Write-Log "WARNING: removing unsupported property 'onPremisesExtensionAttributes' from group delta select" "WARN"
        $allProps = $allProps | Where-Object { $_ -ne 'onPremisesExtensionAttributes' }
    }


    $select = [System.Web.HttpUtility]::UrlEncode(($allProps -join ','))

    $uri = if ($DeltaLink) {
        $DeltaLink
    }
    else {
        "https://graph.microsoft.com/v1.0/groups/delta?`$select=$select"
    }

    $allChanges = New-Object System.Collections.Generic.List[object]
    $finalDeltaLink = $null

    <# do {
        $page = Invoke-GraphJson -Uri $uri -AccessToken $AccessToken

        if ($page.value) {
            foreach ($item in $page.value) {
                $allChanges.Add($item)
            }
        }

        $nextLinkProp = $page.PSObject.Properties['@odata.nextLink'] 
        $deltaLinkProp = $page.PSObject.Properties['@odata.deltaLink'] 
        if ($nextLinkProp) { $uri = $nextLinkProp.Value } 
        elseif ($deltaLinkProp) { $finalDeltaLink = $deltaLinkProp.Value; $uri = $null }
        else { Write-Log "WARNING: No nextLink or deltaLink returned by Graph (group delta)" "WARN"; $uri = $null }

    } while ($uri) #>

    do {
        $page = Invoke-GraphJson -Uri $uri -AccessToken $AccessToken

        if ($page.value) {

            foreach ($item in $page.value) {

                if ($MaxResults -gt 0 -and $allChanges.Count -ge $MaxResults) {
                    Write-Log "Group MaxResults limit reached ($MaxResults) — stopping delta paging early" "WARN"
                    $uri = $null
                    break
                }

                $allChanges.Add($item)
            }
        }

        if (-not $uri) { break }

        $nextLinkProp  = $page.PSObject.Properties['@odata.nextLink']
        $deltaLinkProp = $page.PSObject.Properties['@odata.deltaLink']

        if ($nextLinkProp -and ($MaxResults -eq 0 -or $allChanges.Count -lt $MaxResults)) {
            $uri = $nextLinkProp.Value
        }
        elseif ($deltaLinkProp) {
            $finalDeltaLink = $deltaLinkProp.Value
            $uri = $null
        }
        else {
            Write-Log "WARNING: No nextLink or deltaLink returned by Graph (group delta)" "WARN"
            $uri = $null
        }

    } while ($uri)


    return [pscustomobject]@{
        Changes   = $allChanges
        DeltaLink = $finalDeltaLink
    }
}

function Test-GroupInScope {
    param(
        [Parameter(Mandatory)][object]$Group
    )

    # -------------------- SAFE PROPERTY ACCESS --------------------
    $groupTypesProp = $Group.PSObject.Properties['groupTypes']
    $securityProp   = $Group.PSObject.Properties['securityEnabled']
    $mailProp       = $Group.PSObject.Properties['mailEnabled']

    # Extract values safely
    $groupTypes = @()
    if ($groupTypesProp -and $groupTypesProp.Value) {
        $groupTypes = $groupTypesProp.Value
    }

    $securityEnabled = $true
    if ($securityProp -and $null -ne $securityProp.Value) {
        $securityEnabled = [bool]$securityProp.Value
    }

    $mailEnabled = $false
    if ($mailProp -and $null -ne $mailProp.Value) {
        $mailEnabled = [bool]$mailProp.Value
    }

    # -------------------- LOG FOR DEBUG --------------------
    Write-Log "GROUP SCOPE CHECK: id=$($Group.id) mailEnabled=$mailEnabled securityEnabled=$securityEnabled groupTypes=$groupTypes" "DEBUG"

    # -------------------- SCOPE DECISION --------------------

    # Allow M365 (Unified) groups
    if ($groupTypes -and ($groupTypes -contains "Unified")) {
        return $true
    }

    # Allow mail-enabled security groups
    if ($mailEnabled) {
        return $true
    }

    # Skip everything else
    return $false
}

function Resolve-RecipientConflictByEmail {
    param(
        [Parameter(Mandatory)][string]$Email
    )

    $allRecipients = @(Find-ExistingRecipientByEmail -Email $Email)

    if ($allRecipients.Count -eq 0) {
        return [pscustomobject]@{
            ConflictFound = $false
            RecipientType = $null
            Recipient     = $null
        }
    }

    if ($allRecipients.Count -gt 1) {
        Write-Log "Multiple recipient conflicts found for $Email. Count=$($allRecipients.Count)" "WARN"

        return [pscustomobject]@{
            ConflictFound = $true
            RecipientType = 'Multiple'
            Recipient     = $allRecipients
        }
    }

    $recipient = $allRecipients[0]

    Write-Log "Recipient conflict found for $Email :: Type=$($recipient.RecipientType) Identity=$($recipient.Identity) PrimarySmtp=$($recipient.PrimarySmtpAddress)" "WARN"

    return [pscustomobject]@{
        ConflictFound = $true
        RecipientType = $recipient.RecipientType
        Recipient     = $recipient
    }
}

function Find-TargetContactByEmail {
    param([Parameter(Mandatory)][string]$Email)

    try {
        $normalized = $Email.Trim().ToLowerInvariant()

        $results = Get-MailContact -ResultSize Unlimited | Where-Object {
            $_.ExternalEmailAddress -and
            $_.ExternalEmailAddress.ToString().ToLowerInvariant().Replace("smtp:","") -eq $normalized
        }

        return @($results)
    }
    catch {
        Write-Log "MailContact lookup failed for $Email :: $($_.Exception.Message)" "WARN"
        return @()
    }
}

function Get-PrimaryGroupExternalEmail {

    param([Parameter(Mandatory)][object]$Group)

    # -------------------- MAIL PROPERTY --------------------
    $mailProp = $Group.PSObject.Properties['mail']

    if ($mailProp -and $mailProp.Value -and $mailProp.Value.ToString().Trim()) {
        return $mailProp.Value.ToString().Trim().ToLowerInvariant()
    }

    # -------------------- PROXY ADDRESSES (FALLBACK) --------------------
    $proxyProp = $Group.PSObject.Properties['proxyAddresses']

    if ($proxyProp -and $proxyProp.Value) {

        foreach ($addr in $proxyProp.Value) {

            # Primary SMTP is uppercase "SMTP:"
            if ($addr -cmatch '^SMTP:') {
                return ($addr -replace '^SMTP:', '').ToLowerInvariant()
            }
        }

        # If no uppercase primary, take first lowercase smtp
        foreach ($addr in $proxyProp.Value) {
            if ($addr -match '^smtp:') {
                return ($addr -replace '^smtp:', '').ToLowerInvariant()
            }
        }
    }

    return $null
}

# -------------------- Filtering / normalization --------------------

function Get-TopLevelSelectPropertiesFromAttributeFilters {
    param(
        [object[]]$AttributeFilters,
        [ValidateSet('User','Group')][string]$ObjectType
    )

    $props = New-Object System.Collections.Generic.List[string]

    foreach ($filter in @($AttributeFilters)) {

        if (-not $filter) { continue }

        $filterObjectType = if ($filter.ObjectType) { [string]$filter.ObjectType } else { 'Both' }

        if ($filterObjectType -notin @('Both', $ObjectType)) {
            continue
        }

        if (-not $filter.PropertyPath) { continue }

        $topLevel = ([string]$filter.PropertyPath -split '\.')[0]

        if ($topLevel -and -not $props.Contains($topLevel)) {
            $props.Add($topLevel)
        }
    }

    return @($props)
}

function Get-NestedPropertyValue {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][string]$PropertyPath
    )

    $current = $InputObject

    foreach ($segment in ($PropertyPath -split '\.')) {

        if ($null -eq $current) { return $null }

        # Handle Graph SDK objects explicitly
        if ($current -isnot [psobject]) {
            try {
                $current = $current.$segment
                continue
            }
            catch {
                return $null
            }
        }

        $prop = $current.PSObject.Properties[$segment]

        if ($prop) {
            $current = $prop.Value
        }
        else {
            try {
                $current = $current.$segment
            }
            catch {
                return $null
            }
        }
    }

    return $current
}

function Test-AttributeFilterMatch {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [object[]]$AttributeFilters,
        [ValidateSet('User','Group')][string]$ObjectType
    )

    foreach ($filter in @($AttributeFilters)) {

        if (-not $filter) { continue }

        $filterObjectType = if ($filter.ObjectType) { [string]$filter.ObjectType } else { 'Both' }

        if ($filterObjectType -notin @('Both', $ObjectType)) {
            continue
        }

        $propertyPath = [string]$filter.PropertyPath
        $operator     = if ($filter.Operator) { [string]$filter.Operator } else { 'Equals' }
        $expected     = [string]$filter.Value

        if (-not $propertyPath) { continue }

        # Use corrected nested property resolver
        $actual = Get-NestedPropertyValue -InputObject $InputObject -PropertyPath $propertyPath

        # Debug logging (keep this while validating)
        Write-Log "DEBUG FILTER RAW: Property='$propertyPath' Actual='$actual' Expected='$expected'" "DEBUG"

        #if ($null -eq $actual) { continue }
        if ($null -eq $actual) {
            Write-Log "FILTER PROPERTY MISSING: Property='$propertyPath' ObjectType='$ObjectType'" "DEBUG"
            continue
        }
        
        # Normalize values (this is critical)
        $actualString   = ([string]$actual).Trim().ToLowerInvariant()
        $expectedString = ([string]$expected).Trim().ToLowerInvariant()

        switch -Regex ($operator.ToLowerInvariant()) {

            '^equals$' {
                if ($actualString -eq $expectedString) {
                    Write-Log "FILTER MATCHED: $propertyPath = '$actualString'" "WARN"
                    return $true
                }
            }

            '^notequals$' {
                if ($actualString -ne $expectedString) {
                    return $true
                }
            }

            '^contains$' {
                if ($actualString -like "*$expectedString*") {
                    return $true
                }
            }

            '^match$' {
                if ($actualString -match $expectedString) {
                    return $true
                }
            }

            default {
                Write-Log "Unknown AttributeFilter operator '$operator' for PropertyPath '$propertyPath'" 'WARN'
            }
        }
    }

    return $false
}

function Get-PrimaryExternalEmail {
    param(
        [Parameter(Mandatory)][object]$User,
        [Parameter(Mandatory)][bool]$AllowUpnFallback
    )

    $mailProp = $User.PSObject.Properties['mail']
    $upnProp  = $User.PSObject.Properties['userPrincipalName']

    if ($mailProp -and $mailProp.Value -and $mailProp.Value.ToString().Trim()) {
        return $mailProp.Value.ToString().Trim().ToLowerInvariant()
    }

    if ($AllowUpnFallback -and $upnProp -and $upnProp.Value -and $upnProp.Value.ToString().Trim()) {
        return $upnProp.Value.ToString().Trim().ToLowerInvariant()
    }

    return $null
}

function Test-UserInScope {
    param(
        [Parameter(Mandatory)][object]$User,
        [string[]]$IncludeDomains,
        [string[]]$ExcludeDomains,
        [bool]$AllowUpnFallback
    )

    $email = Get-PrimaryExternalEmail -User $User -AllowUpnFallback:$AllowUpnFallback
    if (-not $email) { return $false }

    $domain = ($email -split '@')[-1].ToLowerInvariant()

    if ($IncludeDomains -and $IncludeDomains.Count -gt 0) {
        if ($IncludeDomains.ForEach({ $_.ToLowerInvariant() }) -notcontains $domain) {
            return $false
        }
    }

    if ($ExcludeDomains -and $ExcludeDomains.Count -gt 0) {
        if ($ExcludeDomains.ForEach({ $_.ToLowerInvariant() }) -contains $domain) {
            return $false
        }
    }

    return $true
}

function Get-SyncKey { 
    param( 
        [Parameter(Mandatory)][string]$SourceTenantId,
        [Parameter(Mandatory)][string]$SourceObjectId,
        [Parameter(Mandatory)][string]$SourceObjectType
    ) 
    
    return "XTSYNC|$SourceTenantId|$SourceObjectType|$SourceObjectId" 
}

function Get-LegacySyncKey { 
    param( 
        [Parameter(Mandatory)][string]$SourceTenantId, 
        [Parameter(Mandatory)][string]$SourceObjectId 
    ) 
    
    return "XTSYNC|$SourceTenantId|$SourceObjectId" 
}

function New-SafeAlias {
    param(
        [Parameter(Mandatory)][string]$Email,
        [Parameter(Mandatory)][string]$SourceObjectId
    )

    $local = ($Email -split '@')[0]
    $local = ($local -replace '[^a-zA-Z0-9._-]', '')
    if (-not $local) { $local = "contact" }

    $suffix = ($SourceObjectId -replace '-','').Substring(0,8)
    $alias = "{0}_{1}" -f $local, $suffix

    if ($alias.Length -gt 64) {
        $alias = $alias.Substring(0,64)
    }

    return $alias
}

function Get-DisplayNameForTarget {
    param(
        [Parameter(Mandatory)][object]$User,
        [Parameter(Mandatory)][bool]$AppendSourceTenantToDisplayName,
        [Parameter(Mandatory)][string]$SourceTenantName
    )

    $name = if ($User.displayName) { $User.displayName.Trim() } else { $User.userPrincipalName }
    if ($AppendSourceTenantToDisplayName) {
        return "$name ($SourceTenantName)"
    }
    return $name
}

# -------------------- Exchange target side --------------------

function Find-ExistingRecipientByEmail {
    param(
        [Parameter(Mandatory)][string]$Email
    )

    $normalized = $Email.Trim().ToLowerInvariant()
    try {
        try {
            $matches = Get-Recipient -ResultSize Unlimited -Filter "EmailAddresses -eq 'smtp:$normalized'"
        }
        catch {
            Write-Log "Recipient lookup failed for $Email :: $($_.Exception.Message)" "WARN"
            return @()
        }

        return @($matches)

        }
    catch {
        Write-Log "Recipient lookup failed for $Email :: $($_.Exception.Message)" "WARN"
        return @()
    }
}

function Connect-TargetExchange {
    param([Parameter(Mandatory)][object]$Config)

    Import-RequiredModule -Name ExchangeOnlineManagement

    Connect-ExchangeOnline `
        -AppId $Config.TargetExoAppId `
        -CertificateThumbprint $Config.TargetExoCertThumbprint `
        -Organization $Config.TargetOrganization `
        -ShowBanner:$false | Out-Null
}

function Disconnect-TargetExchangeSafe {
    try {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    } catch { }
}

function Get-ExistingTargetContacts {
    param([Parameter(Mandatory)][string]$SourceTenantId)

    $prefix = "XTSYNC|$SourceTenantId|*"
    $contacts = @()

    try {
        # Prefer server-side filter where possible
        $contacts = Get-MailContact -ResultSize Unlimited -Filter "CustomAttribute15 -like '$prefix'"
    }
    catch {
        Write-Log "Get-MailContact server-side filter failed, falling back to client-side filtering. $_" 'WARN'
        $contacts = Get-MailContact -ResultSize Unlimited | Where-Object { $_.CustomAttribute15 -like $prefix }
    }

    $idx = @{}
    foreach ($c in $contacts) {
        $storedKey = $c.CustomAttribute15

        if (-not $storedKey) { continue }

        # Always store the exact key that exists on the object 
        if (-not $idx.ContainsKey($storedKey)) { $idx[$storedKey] = $c }

        # Backward compatibility: # old format = XTSYNC|TenantId|ObjectId # map it to new User key so lookups still work 
        $parts = $storedKey -split '\|'

        if ($parts.Count -eq 3 -and $parts[0] -eq 'XTSYNC' -and $parts[1] -eq $SourceTenantId) {
            $legacyObjectId = $parts[2]
            $newUserKey = Get-SyncKey -SourceTenantId $SourceTenantId -SourceObjectId $legacyObjectId -SourceObjectType 'User'

            if (-not $idx.ContainsKey($newUserKey)) {
            $idx[$newUserKey] = $c
            }
        }
    }
 
    return $idx
}

function Set-TargetMailContact {
    param(
        [Parameter(Mandatory)][object]$User,
        [Parameter(Mandatory)][object]$Config,
        [Parameter()][hashtable]$ExistingContacts,
        [Parameter()][object]$ExistingContact
    )

    $externalEmail = Get-PrimaryExternalEmail -User $User -AllowUpnFallback:$Config.AllowUpnFallback
    if (-not $externalEmail) {
        Write-Log "Skipping user $($User.id) because no mail/UPN could be used." 'WARN'
        return "Skipped"
    }

    $syncKey = Get-SyncKey -SourceTenantId $Config.SourceTenantId -SourceObjectId $User.id -SourceObjectType 'User'
    $displayName = Get-DisplayNameForTarget -User $User -AppendSourceTenantToDisplayName:$Config.AppendSourceTenantToDisplayName -SourceTenantName $Config.SourceTenantName
    $alias = New-SafeAlias -Email $externalEmail -SourceObjectId $User.id

    # -------------------- HYBRID LOOKUP --------------------
    $existing = $ExistingContact

    # ======================================================
    # ADOPTION (SYNC-KEY MISSING BUT SMTP MATCH EXISTS)
    # ======================================================

    if (-not $existing) {

        $recipientConflict = Resolve-RecipientConflictByEmail -Email $externalEmail

        if ($recipientConflict.ConflictFound) {

            $recipient     = $recipientConflict.Recipient
            $recipientType = $recipientConflict.RecipientType

            # -------------------- MAILCONTACT → ADOPT --------------------
            if ($recipientType -eq 'MailContact') {

                Write-Log "Adopting existing MailContact (SMTP match): $externalEmail" "WARN"

                $identity = $recipient.Identity

                if ($SeedTargetFromSource -ne 'None') {
                    Write-Log "Seed mode active — adoption skipped: $externalEmail" "WARN"
                    return "Skipped"
                }

                if ($PSCmdlet.ShouldProcess($identity, "Adopt existing MailContact")) {

                    Invoke-ExoWithRetry {
                        Set-MailContact `
                            -Identity $identity `
                            -DisplayName $displayName `
                            -ExternalEmailAddress $externalEmail `
                            -CustomAttribute15 $syncKey
                    }

                    Write-Log "Adopted existing contact: $displayName <$externalEmail> [$syncKey]" "INFO"

                    return "Updated"
                }
            }

            # -------------------- MULTIPLE MATCHES --------------------
            elseif ($recipientType -eq 'Multiple') {

                Write-Log "Multiple recipient conflict — skipping adoption: $externalEmail" "WARN"
                return "Skipped"
            }

            # -------------------- NON-CONTACT OBJECT --------------------
            else {

                Write-Log ("Cannot adopt — SMTP owned by {0}: {1}" -f $recipientType, $externalEmail) "WARN"
                return "Skipped"
            }
        }
    }

    # ======================================================
    # NORMAL PROCESSING (ONLY IF NOT FILTERED)
    # ======================================================

    if ($existing) {

        $identity = $existing.Identity
        $existingKey = $existing.CustomAttribute15

        $needsUpdate = $false

        if ($ForceDisplayNameRefresh -and $Config.AppendSourceTenantToDisplayName) { 
            $needsUpdate = $true 
        }
        elseif ($existing.DisplayName -ne $displayName) { 
            $needsUpdate = $true 
        }

        $existingEmail = $existing.ExternalEmailAddress.ToString().ToLower().Replace("smtp:","")
        if ($existingEmail -ne $externalEmail.ToLower()) { $needsUpdate = $true }

        if ($existingKey -ne $syncKey) { $needsUpdate = $true }

        if (-not $needsUpdate) {
            Write-Log "No changes required: $displayName <$externalEmail>" "DEBUG"
            return "Unchanged"
        }

        if ($PSCmdlet.ShouldProcess($identity, "Update target mail contact")) {

            if ($ForceDisplayNameRefresh) { 
                Write-Log "Forcing displayName refresh: $displayName" "WARN" 
            } 
            else { 
                Write-Log "Updating contact: Identity=$identity Email=$externalEmail SyncKey=$syncKey" "DEBUG"
            }

            Invoke-ExoWithRetry {
                Set-MailContact -Identity $identity -DisplayName $displayName -ExternalEmailAddress $externalEmail -CustomAttribute15 $syncKey
            }

            Write-Log "Updated contact: $displayName <$externalEmail> [$syncKey]" "INFO"
            return "Updated"
        }

    }
else {

        if ($PSCmdlet.ShouldProcess($displayName, "Create target mail contact")) {

            $shortName = if ($displayName.Length -gt 40) {
                $displayName.Substring(0,40)
            } else {
                $displayName
            }

            $uniqueSuffix = ($User.id -replace '-', '').Substring(0,16)
            $Name = "$shortName-$uniqueSuffix"

            if ($Name.Length -gt 64) {
                $Name = $Name.Substring(0,64)
            }

            # ------------------------------------------------------
            # PRE-CREATE RECIPIENT CONFLICT CHECK
            # ------------------------------------------------------
            $recipientConflict = Resolve-RecipientConflictByEmail -Email $externalEmail

            if ($recipientConflict.ConflictFound) {

                $recipient = $recipientConflict.Recipient
                $recipientType = $recipientConflict.RecipientType

                if ($recipientType -eq 'MailContact') {

                    Write-Log "Found existing MailContact with same SMTP — adopting instead of creating: $externalEmail" "WARN"

                    #$existingContact = Find-TargetContactByEmail -Email $externalEmail | Select-Object -First 1
                    $key = $externalEmail.Trim().ToLowerInvariant()
                    $existingContact = $null

                    if ($existingContactsByEmail -and $existingContactsByEmail.ContainsKey($key)) {
                        $existingContact = $existingContactsByEmail[$key] | Select-Object -First 1
                    }


                    if ($existingContact) {

                        Invoke-ExoWithRetry {
                            Set-MailContact `
                                -Identity $existingContact.Identity `
                                -DisplayName $displayName `
                                -ExternalEmailAddress $externalEmail `
                                -CustomAttribute15 $syncKey
                        }

                        Write-Log "Adopted existing contact: $displayName <$externalEmail> [$syncKey]" "INFO"
                        return "Updated"
                    }
                    else {
                        Write-Log "Recipient conflict reported MailContact but Get-MailContact could not resolve it: $externalEmail" "WARN"
                        return "Skipped"
                    }
                }
                else {
                    Write-Log ("Cannot create MailContact — SMTP already owned by {0}: {1}" -f $recipientType, $externalEmail) "WARN"
                    return "Skipped"
                }
            }

            Write-Log "Creating contact: Name=$Name DisplayName=$displayName Alias=$alias Email=$externalEmail SyncKey=$syncKey"

            $new = $null

            try {
                $new = Invoke-ExoWithRetry {
                    New-MailContact -Name $Name -DisplayName $displayName -Alias $alias -ExternalEmailAddress $externalEmail
                }
            }
            catch {
                Write-Log "Create failed for $displayName <$externalEmail> :: $($_.Exception.Message)" "ERROR"
                return "Skipped"
            }

            if (-not $new -or -not $new.Identity) {
                Write-Log "ERROR: Failed to create contact for $displayName" "ERROR"
                return "Skipped"
            }

            Invoke-ExoWithRetry {
                Set-MailContact -Identity $new.Identity -CustomAttribute15 $syncKey
            }

            Write-Log "Created contact: $displayName <$externalEmail> [$syncKey]" "INFO"
            return "Created"
        }
    }
}

function Set-TargetMailContactFromGroup {
    param(
        [Parameter(Mandatory)][object]$Group,
        [Parameter(Mandatory)][object]$Config,
        [Parameter()][hashtable]$ExistingContacts,
        [Parameter()][object]$ExistingContact
    )

    $externalEmail = Get-PrimaryGroupExternalEmail -Group $Group
    if (-not $externalEmail) {
        Write-Log "Skipping group $($Group.id) because no mail value is present." "WARN"
        return "Skipped"
    }

    $syncKey = Get-SyncKey -SourceTenantId $Config.SourceTenantId -SourceObjectId $Group.id -SourceObjectType 'Group'

    $displayName = Get-DisplayNameForTarget `
        -User $Group `
        -AppendSourceTenantToDisplayName $Config.AppendSourceTenantToDisplayName `
        -SourceTenantName $Config.SourceTenantName

    $alias = New-SafeAlias -Email $externalEmail -SourceObjectId $Group.id

    # -------------------- HYBRID LOOKUP --------------------
    $existing = $ExistingContact

    # ======================================================
    # ADOPTION (SYNC-KEY MISSING BUT SMTP MATCH EXISTS)
    # ======================================================

    if (-not $existing) {

        $recipientConflict = Resolve-RecipientConflictByEmail -Email $externalEmail

        if ($recipientConflict.ConflictFound) {

            $recipient     = $recipientConflict.Recipient
            $recipientType = $recipientConflict.RecipientType

            # -------------------- MAILCONTACT → ADOPT --------------------
            if ($recipientType -eq 'MailContact') {

                Write-Log "Adopting existing MailContact (SMTP match): $externalEmail" "WARN"

                $identity = $recipient.Identity

                if ($SeedTargetFromSource -ne 'None') {
                    Write-Log "Seed mode active — adoption skipped: $externalEmail" "WARN"
                    return "Skipped"
                }

                if ($PSCmdlet.ShouldProcess($identity, "Adopt existing MailContact")) {

                    Invoke-ExoWithRetry {
                        Set-MailContact `
                            -Identity $identity `
                            -DisplayName $displayName `
                            -ExternalEmailAddress $externalEmail `
                            -CustomAttribute15 $syncKey
                    }

                    Write-Log "Adopted existing contact: $displayName <$externalEmail> [$syncKey]" "INFO"

                    return "Updated"
                }
            }

            # -------------------- MULTIPLE MATCHES --------------------
            elseif ($recipientType -eq 'Multiple') {

                Write-Log "Multiple recipient conflict — skipping adoption: $externalEmail" "WARN"
                return "Skipped"
            }

            # -------------------- NON-CONTACT OBJECT --------------------
            else {

                Write-Log ("Cannot adopt — SMTP owned by {0}: {1}" -f $recipientType, $externalEmail) "WARN"
                return "Skipped"
            }
        }
    }

    # ======================================================
    # UPDATE EXISTING
    # ======================================================

    if ($existing) {

        $identity    = $existing.Identity
        $existingKey = $existing.CustomAttribute15

        $needsUpdate = $false

        if ($ForceDisplayNameRefresh -and $Config.AppendSourceTenantToDisplayName) {
            $needsUpdate = $true
        }
        elseif ($existing.DisplayName -ne $displayName) {
            $needsUpdate = $true
        }

        $existingEmail = $existing.ExternalEmailAddress.ToString().ToLower().Replace("smtp:", "")
        if ($existingEmail -ne $externalEmail.ToLower()) {
            $needsUpdate = $true
        }

        if ($existingKey -ne $syncKey) {
            $needsUpdate = $true
        }

        if (-not $needsUpdate) {
            Write-Log "No changes required (group): $displayName <$externalEmail>" "DEBUG"
            return "Unchanged"
        }

        if ($PSCmdlet.ShouldProcess($identity, "Update group contact")) {

            Write-Log "Updating group contact: Identity=$identity Email=$externalEmail SyncKey=$syncKey"

            Invoke-ExoWithRetry {
                Set-MailContact `
                    -Identity $identity `
                    -DisplayName $displayName `
                    -ExternalEmailAddress $externalEmail `
                    -CustomAttribute15 $syncKey
            }

            Write-Log "Updated group contact: $displayName <$externalEmail> [$syncKey]"
            return "Updated"
        }
    }

    # ======================================================
    # CREATE NEW (WITH CONFLICT HANDLING)
    # ======================================================

    else {

        if ($PSCmdlet.ShouldProcess($displayName, "Create group contact")) {

            $shortName = if ($displayName.Length -gt 40) {
                $displayName.Substring(0,40)
            } else {
                $displayName
            }

            $uniqueSuffix = ($Group.id -replace '-', '').Substring(0,16)
            $Name = "$shortName-$uniqueSuffix"

            if ($Name.Length -gt 64) {
                $Name = $Name.Substring(0,64)
            }

            # ------------------------------------------------------
            # PRE-CREATE RECIPIENT CONFLICT CHECK (CRITICAL)
            # ------------------------------------------------------

            $recipientConflict = Resolve-RecipientConflictByEmail -Email $externalEmail

            if ($recipientConflict.ConflictFound) {

                $recipient     = $recipientConflict.Recipient
                $recipientType = $recipientConflict.RecipientType

                # ------------------------------------------------------
                # CASE 1: EXISTING MAILCONTACT → ADOPT
                # ------------------------------------------------------
                if ($recipientType -eq 'MailContact') {

                    Write-Log "Adopting existing GROUP MailContact due to SMTP match: $externalEmail" "WARN"

                    #$existingContact = Find-TargetContactByEmail -Email $externalEmail | Select-Object -First 1
                    $key = $externalEmail.Trim().ToLowerInvariant()
                    $existingContact = $null

                    if ($existingContactsByEmail -and $existingContactsByEmail.ContainsKey($key)) {
                        $existingContact = $existingContactsByEmail[$key] | Select-Object -First 1
                    }


                    if ($existingContact) {

                        Invoke-ExoWithRetry {
                            Set-MailContact `
                                -Identity $existingContact.Identity `
                                -DisplayName $displayName `
                                -ExternalEmailAddress $externalEmail `
                                -CustomAttribute15 $syncKey
                        }

                        Write-Log "Adopted existing group contact: $displayName <$externalEmail> [$syncKey]" "INFO"
                        return "Updated"
                    }
                    else {
                        Write-Log "Conflict reported MailContact but lookup failed for $externalEmail" "WARN"
                        return "Skipped"
                    }
                }

                # ------------------------------------------------------
                # CASE 2: MULTIPLE MATCHES → SKIP
                # ------------------------------------------------------
                elseif ($recipientType -eq 'Multiple') {

                    Write-Log "Multiple recipient conflict detected — skipping group creation: $externalEmail" "WARN"
                    return "Skipped"
                }

                # ------------------------------------------------------
                # CASE 3: NON-CONTACT RECIPIENT → SKIP (CRITICAL)
                # ------------------------------------------------------
                else {

                    Write-Log ("Cannot create MailContact — SMTP already owned by {0}: {1}" -f $recipientType, $externalEmail) "WARN"
                    return "Skipped"
                }
            }

            # ------------------------------------------------------
            # SAFE TO CREATE
            # ------------------------------------------------------

            Write-Log "Creating group contact: Name=$Name DisplayName=$displayName Email=$externalEmail SyncKey=$syncKey"

            $new = $null

            try {
                $new = Invoke-ExoWithRetry {
                    New-MailContact `
                        -Name $Name `
                        -DisplayName $displayName `
                        -Alias $alias `
                        -ExternalEmailAddress $externalEmail
                }
            }
            catch {
                Write-Log "ERROR: Failed to create group contact: $($_.Exception.Message)" "ERROR"
                return "Skipped"
            }

            if (-not $new -or -not $new.Identity) {
                Write-Log "ERROR: Creation returned null for group contact $displayName" "ERROR"
                return "Skipped"
            }

            Invoke-ExoWithRetry {
                Set-MailContact -Identity $new.Identity -CustomAttribute15 $syncKey
            }

            Write-Log "Created group contact: $displayName <$externalEmail> [$syncKey]"
            return "Created"
        }
    }
}

function Remove-TargetMailContactBySyncKey {
    param(
        [Parameter(Mandatory)][string]$SyncKey,
        [Parameter()][hashtable]$ExistingContacts,
        [Parameter(Mandatory)][bool]$DisableDeletes
    )

    if ($DisableDeletes) {
        Write-Log "Delete suppressed by DisableDeletes for sync key: $SyncKey" 'WARN'
        return
    }

    # -------------------- HYBRID LOOKUP --------------------
    $existing = $null

    if ($ExistingContacts) {
        if ($ExistingContacts.ContainsKey($SyncKey)) {
            $existing = $ExistingContacts[$SyncKey]
        }
    }
    else {
        try {
            $existing = Invoke-ExoWithRetry {
                Get-MailContact -ResultSize 1 -Filter "CustomAttribute15 -eq '$SyncKey'"
            }
        }
        catch {
            Write-Log "Lookup failed for delete (syncKey=$SyncKey): $($_.Exception.Message)" "WARN"
            return
        }
    }

    # -------------------- DELETE --------------------
    if ($existing) {
        if ($PSCmdlet.ShouldProcess($existing.Identity, "Remove target mail contact")) {

            $displayName = $existing.DisplayName
            $email = $existing.ExternalEmailAddress

            Invoke-ExoWithRetry {
                Remove-MailContact -Identity $existing.Identity -Confirm:$false
            }

            $contact = $null
            
            if ($ExistingContacts -and $ExistingContacts.ContainsKey($SyncKey)) { $contact = $ExistingContacts[$SyncKey] }

            if ($displayName -or $email) { 
                Write-Log "Removed contact: $displayName <$email>" "INFO" 
            } else { 
                Write-Log "Removed contact with sync key: $SyncKey" "INFO" 
            }
            
            # Write-Log "Removed contact with sync key: $SyncKey"
            Return "Deleted"
        }
    }
    else {
        Write-Log "Delete requested but no existing target contact found for sync key: $SyncKey" 'WARN'
        return "NotFound"
    }
}


function Get-SafeName { 
    param([string]$Value) 
    if (-not $Value) { return "Contact" } 
    $Value = $Value.Trim() 
    if ($Value.Length -gt 64) { return $Value.Substring(0,64) } 
    return $Value }

function Get-SafeAlias { 
    param([string]$Email,[string]$Id)
    $base = ($Email -split '@')[0] 
    $base = $base -replace '[^a-zA-Z0-9._-]', '' 
    $suffix = ($Id -replace '-', '').Substring(0,8) 
    $alias = "$base`_$suffix" 
    if ($alias.Length -gt 64) { $alias = $alias.Substring(0,64) } 
    return $alias }

function Invoke-ExoWithRetry {
    param( 
        [scriptblock]$ScriptBlock, 
        [int]$MaxRetries = 3, 
        [int]$InitialDelaySeconds = 2 
    ) 
    
    for ($i = 1; $i -le $MaxRetries; $i++) { 
        try { 
            return & $ScriptBlock 
        } 
        catch {
            $msg = $_.Exception.Message

            # Non-transient / deterministic failures: do not retry
            $nonTransientPatterns = @(
                'proxy address.*already being used',
                'ProxyAddressExistsException',
                'is already being used by the proxy addresses',
                'cannot find object',
                'doesn''t exist',
                'is not unique',
                'invalid',
                'not authorized',
                'insufficient privileges'
            )

            foreach ($pattern in $nonTransientPatterns) {
                if ($msg -match $pattern) {
                    Write-Log "EXO non-transient error (no retry): $msg" "ERROR"
                    throw
                }
            }
            
            if ($i -ge $MaxRetries) { 
                Write-Log "EXO operation failed after $MaxRetries attempts: $msg" "ERROR" 
                throw 
            } 

            $delay = [int]($InitialDelaySeconds * [math]::Pow(2, ($i - 1)))

            Write-Log "EXO transient error (attempt $i/$MaxRetries): $msg — retrying in $delay seconds" "WARN"
            Start-Sleep -Seconds $delay 
        } 
    } 
}

function Get-SafeUpn {
    param([object]$User)

    $upnProp = $User.PSObject.Properties['userPrincipalName']

    if ($upnProp -and $upnProp.Value -and $upnProp.Value.ToString().Trim()) {
        return $upnProp.Value.ToString().Trim().ToLowerInvariant()
    }

    return $null
}

function Get-TargetContactBySyncKey {
    param(
        [Parameter(Mandatory)][string]$SyncKey
    )

    try {
        return Invoke-ExoWithRetry {
            Get-MailContact -ResultSize 1 -Filter "CustomAttribute15 -eq '$SyncKey'"
        }
    }
    catch {
        Write-Log "Lookup failed for syncKey $SyncKey :: $($_.Exception.Message)" "WARN"
        return $null
    }
}

function Invoke-GraphBatch {
    param(
        [Parameter(Mandatory)][array]$Requests,
        [Parameter(Mandatory)][string]$AccessToken
    )

    $body = @{
        requests = $Requests
    } | ConvertTo-Json -Depth 5

    try {
        return Invoke-RestMethod `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Headers @{
                Authorization = "Bearer $AccessToken"
                "Content-Type" = "application/json"
            } `
            -Body $body `
            -ErrorAction Stop
    }
    catch {
        $responseBody = $null
        $exceptionMessage = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { ($_ | Out-String).Trim() }

        if ($_.PSObject.Properties['ErrorDetails'] -and $null -ne $_.ErrorDetails) {
            if ($_.ErrorDetails -is [string]) {
                $responseBody = $_.ErrorDetails
            }
            elseif ($_.ErrorDetails.PSObject.Properties['Message']) {
                $responseBody = $_.ErrorDetails.Message
            }
            else {
                $responseBody = ($_.ErrorDetails | Out-String).Trim()
            }
        }

        if (-not $responseBody) {
            $responseBody = "<no response body available>"
        }

        Write-Log "GRAPH BATCH FAILED" "ERROR"
        Write-Log "GRAPH BATCH ERROR: $exceptionMessage" "ERROR"
        Write-Log "GRAPH BATCH RAW RESPONSE: $responseBody" "ERROR"
        throw
    }
}

function Invoke-BatchReconciliation {
    param(
        [Parameter(Mandatory)][hashtable]$ExistingContacts,
        [Parameter(Mandatory)][object]$Config,
        [Parameter(Mandatory)][string]$GraphToken,
        [ref]$DeletedCounter,
        [ref]$SkippedCounter
    )

    Write-Log "Starting unified batch reconciliation pass..." "WARN"

    if (-not $ExistingContacts -or $ExistingContacts.Count -eq 0) {
        Write-Log "No existing contacts to reconcile." "DEBUG"
        return
    }

    # Group keys by object type
    $keysByType = @{
        User  = @()
        Group = @()
    }

    foreach ($key in $ExistingContacts.Keys) {

        if (-not $key.StartsWith("XTSYNC|")) { continue }

        $parts = $key -split '\|'
        if ($parts.Count -lt 4) { continue }

        $type = $parts[2]

        if ($keysByType.ContainsKey($type)) {
            $keysByType[$type] += $key
        }
    }

    $batchSize = 20

    foreach ($type in $keysByType.Keys) {

        $keys = $keysByType[$type]

        if ($keys.Count -eq 0) { continue }

        Write-Log "Reconciling $type objects: count=$($keys.Count)" "WARN"

        for ($i = 0; $i -lt $keys.Count; $i += $batchSize) {

            # Safe slice
            $endIndex = [Math]::Min($i + $batchSize - 1, $keys.Count - 1)
            $batchKeys = $keys[$i..$endIndex]

            $requests = @()
            $idMap    = @{}
            $counter  = 1

            foreach ($key in $batchKeys) {

                $parts = $key -split '\|'
                if ($parts.Count -lt 4) { continue }

                $objectId = $parts[3]
                $reqId    = "$counter"

                # Type-based endpoint routing

                if ($type -eq 'User') {

                    # USER filtering is constrained to onPremisesExtensionAttributes only
                    $url = "/users/$($objectId)?`$select=id,userPrincipalName,displayName,onPremisesExtensionAttributes"
                }
                elseif ($type -eq 'Group') {

                    # GROUP filtering remains fully dynamic based on XML
                    $extraProps = Get-TopLevelSelectPropertiesFromAttributeFilters `
                        -AttributeFilters $Config.AttributeFilters `
                        -ObjectType 'Group'

                    $selectProps = @(
                        'id',
                        'displayName',
                        'mail',
                        'mailEnabled',
                        'securityEnabled',
                        'groupTypes',
                        'proxyAddresses'
                    ) + $extraProps | Select-Object -Unique

                    $url = "/groups/$($objectId)?`$select=$($selectProps -join ',')"
                }

                $requests += @{
                    id     = $reqId
                    method = "GET"
                    url    = $url
                }

                $idMap[$reqId] = $key
                $counter++
            }

            if ($requests.Count -eq 0) { continue }

            $response = Invoke-GraphBatch -Requests $requests -AccessToken $GraphToken

            foreach ($res in $response.responses) {

                if ($res.status -ne 200) {
                    Write-Log "Batch call failed: id=$($res.id) status=$($res.status)" "WARN"
                    continue
                }

                $body    = $res.body
                $syncKey = $idMap[$res.id]

            if ($type -eq 'User') {

                $name = $body.userPrincipalName

                foreach ($filter in @($Config.AttributeFilters)) {

                    if (-not $filter) { continue }

                    $filterObjectType = if ($filter.ObjectType) { [string]$filter.ObjectType } else { 'Both' }

                    if ($filterObjectType -notin @('User','Both')) {
                        continue
                    }

                    $propertyPath = [string]$filter.PropertyPath
                    if (-not $propertyPath) { continue }

                    $value = Get-NestedPropertyValue `
                        -InputObject $body `
                        -PropertyPath $propertyPath

                    Write-Log "RECON DEBUG (User): $name Property='$propertyPath' Value='$value'" "DEBUG"
                }
            }
            else {

                $name = $body.displayName

                foreach ($filter in @($Config.AttributeFilters)) {

                    if (-not $filter) { continue }

                    $filterObjectType = if ($filter.ObjectType) { [string]$filter.ObjectType } else { 'Both' }

                    if ($filterObjectType -notin @('Group','Both')) {
                        continue
                    }

                    $propertyPath = [string]$filter.PropertyPath
                    if (-not $propertyPath) { continue }

                    $value = Get-NestedPropertyValue `
                        -InputObject $body `
                        -PropertyPath $propertyPath

                    Write-Log "RECON DEBUG (Group): $name Property='$propertyPath' Value='$value'" "DEBUG"
                }
            }


                # Reuse same filter logic
                $match = Test-AttributeFilterMatch `
                    -InputObject $body `
                    -AttributeFilters $Config.AttributeFilters `
                    -ObjectType $type

                if ($match) {

                    Write-Log "RECON DELETE ($type): $name [$syncKey]" "WARN"

                    $deleteResult = Remove-TargetMailContactBySyncKey `
                        -SyncKey $syncKey `
                        -ExistingContacts $ExistingContacts `
                        -DisableDeletes:$Config.DisableDeletes

                    if ($deleteResult -eq "Deleted") {
                        $DeletedCounter.Value++
                    }
                    else {
                        $SkippedCounter.Value++
                    }
                }
            }
        }
    }
}


# -------------------- Main --------------------

$paramConfig = [pscustomobject]@{
    SourceTenantId                  = $SourceTenantId
    SourceTenantName                = $SourceTenantName
    SourceClientId                  = $SourceClientId
    SourceClientSecret              = $SourceClientSecret

    TargetOrganization              = $TargetOrganization
    TargetTenantId                  = $TargetTenantId
    TargetTenantName                = $TargetTenantName
    TargetExoAppId                  = $TargetExoAppId
    TargetExoCertThumbprint         = $TargetExoCertThumbprint

    StateRoot                       = $StateRoot
    LogRoot                         = $LogRoot
    IncludeDomains                  = $IncludeDomains
    ExcludeDomains                  = $ExcludeDomains
    ServiceAccountPatterns          = @($servicePatterns)
    AttributeFilters                = @()

    AppendSourceTenantToDisplayName = [bool]$AppendSourceTenantToDisplayName
    AllowUpnFallback                = [bool]$AllowUpnFallback
    DisableDeletes                  = [bool]$DisableDeletes
    TopUsers                        = $TopUsers 
    SourceObjectType                = $SourceObjectType
}

$xmlConfig = $null
if ($ConfigXmlPath) {
    $xmlConfig = Get-ConfigFromXml -Path $ConfigXmlPath
}

$config = Merge-Config -ParamConfig $paramConfig -XmlConfig $xmlConfig

$config.AppendSourceTenantToDisplayName = [bool]$config.AppendSourceTenantToDisplayName
$config.AllowUpnFallback = [bool]$config.AllowUpnFallback
$config.DisableDeletes = [bool]$config.DisableDeletes
$config.TopUsers = [int]$config.TopUsers

Assert-Config -Config $config

Assert-AttributeFilterDesign -AttributeFilters $config.AttributeFilters

Ensure-Folder -Path $config.StateRoot
Ensure-Folder -Path $config.LogRoot

$script:LogFile = Join-Path $config.LogRoot ("{0:yyyyMMdd_HHmmss}_{1}_TO_{2}.log" -f (Get-Date), $config.SourceTenantName, $config.TargetTenantName)
New-Item -ItemType File -Path $script:LogFile -Force | Out-Null

$stateFile = Get-StateFilePath -Config $config
$state = Load-State -Path $stateFile

$IsFirstRun = -not $state.UserDeltaLink -and -not $state.GroupDeltaLink

if ($IsFirstRun) {
    Write-Log "First run detected" "WARN"
}

Write-Log "Loaded LastReconciliationUtc: $($state.LastReconciliationUtc)" "DEBUG"

$servicePatterns = @()

if ($config.ServiceAccountPatterns -and $config.ServiceAccountPatterns.Count -gt 0) {
    $servicePatterns = $config.ServiceAccountPatterns
    Write-Log "ServiceAccountPatterns defined in config, using those settings." "WARN"
}
else {
    # Safe fallback if XML is missing or empty
    $servicePatterns = @(
        '^sync_',
        '^aadconnect',
        '^msol_',
        '^svc[-_]',
        '^service',
        'azureadconnect'
    )

    Write-Log "ServiceAccountPatterns not defined in config. Using built-in defaults." "WARN"
}

if ($config.AttributeFilters -and $config.AttributeFilters.Count -gt 0) {
    foreach ($af in $config.AttributeFilters) {
        Write-Log "AttributeFilter loaded: ObjectType=$($af.ObjectType) PropertyPath=$($af.PropertyPath) Operator=$($af.Operator) Value=$($af.Value)" "INFO"
    }
}
else {
    Write-Log "No AttributeFilters defined in config." "INFO"
}

# ---------------------------------------------------------
# Time-based reconciliation trigger
# ---------------------------------------------------------

if (-not ($state.PSObject.Properties.Name -contains 'LastReconciliationUtc')) {
    $state | Add-Member -MemberType NoteProperty -Name LastReconciliationUtc -Value $null
}

$RunReconciliation = $false
$nowUtc = [DateTimeOffset]::UtcNow

if ($ForceReconciliation) {
    Write-Log "ForceReconciliation switch detected — forcing reconciliation run" "WARN" 
    $RunReconciliation = $true 
}

elseif ($IsFirstRun) {

    Write-Log "First run detected — skipping reconciliation" "WARN"
    $RunReconciliation = $false
}

elseif (-not $state.LastReconciliationUtc) {

    Write-Log "No previous reconciliation timestamp found — reconciliation will run" "WARN"
    $RunReconciliation = $true
}
else {

    try {
        #$lastRecon = [DateTime]::Parse($state.LastReconciliationUtc)
        $lastRecon = [DateTimeOffset]::Parse($state.LastReconciliationUtc)

        $elapsedHours = ($nowUtc - $lastRecon).TotalHours

        Write-Log "Reconciliation timing: last=$($lastRecon.UtcDateTime) now=$($nowUtc.UtcDateTime) elapsedHours=$([math]::Round($elapsedHours,2))" "DEBUG"

        # ---------------------------------------------------------
        # Max age override (force reconciliation if too old)
        # ---------------------------------------------------------

        $MaxReconciliationAgeHours = 24

        if ($elapsedHours -ge $MaxReconciliationAgeHours) {
            Write-Log "Max age override triggered: $([math]::Round($elapsedHours,2)) hours since last reconciliation (threshold: $MaxReconciliationAgeHours)" "WARN"
            $RunReconciliation = $true
        }

        elseif ($elapsedHours -ge $ReconciliationIntervalHours) {
            Write-Log "Reconciliation interval met: $([math]::Round($elapsedHours,2)) hours (threshold: $ReconciliationIntervalHours)" "WARN"
            $RunReconciliation = $true
        }
        else {
            Write-Log "Skipping reconciliation: only $([math]::Round($elapsedHours,2)) hours elapsed (threshold: $ReconciliationIntervalHours)" "DEBUG"
        }
    }
    catch {
        Write-Log "Invalid LastReconciliationUtc value — forcing reconciliation" "WARN"
        $RunReconciliation = $true
    }
}

Write-Log "Starting sync: $($config.SourceTenantName) -> $($config.TargetTenantName)"
Write-Log "State file: $stateFile"
Write-Log "Log file  : $script:LogFile"

if ($SeedTargetFromSource -ne 'None') {
    Write-Log "SEED MODE ENABLED — target lookup and preload are disabled" "WARN"
}

try {
    $graphToken = Get-GraphToken `
        -TenantId $config.SourceTenantId `
        -ClientId $config.SourceClientId `
        -ClientSecret $config.SourceClientSecret

    Write-Log "Acquired Graph token for source tenant."

    # ---------------- FORCE FULL SYNC OVERRIDE ----------------

    if ($ForceFullSync -ne 'None') {

        Write-Log "ForceFullSync override enabled: $ForceFullSync" "WARN"

        if ($ForceFullSync -in @('User','Both')) {
            $state.UserDeltaLink = $null
            Write-Log "User deltaLink reset (forced full sync)"
        }

        if ($ForceFullSync -in @('Group','Both')) {
            $state.GroupDeltaLink = $null
            Write-Log "Group deltaLink reset (forced full sync)"
        }
    }

    # ---------------- DELTA PRE-CHECK (EARLY EXIT OPTIMIZATION) ----------------

    $userDeltaResult  = $null
    $groupDeltaResult = $null

    $userChangeCount  = 0
    $groupChangeCount = 0

    # -------- USER DELTA --------
    if ($config.SourceObjectType -in @('User','Both')) {

        #$userDeltaResult = Get-UserDeltaChanges -AccessToken $graphToken -DeltaLink $state.UserDeltaLink
        # $userDeltaResult = Get-UserDeltaChanges -AccessToken $graphToken -DeltaLink $state.UserDeltaLink -AttributeFilters $config.AttributeFilters
        $userDeltaResult = Get-UserDeltaChanges -AccessToken $graphToken -DeltaLink $state.UserDeltaLink -AttributeFilters $config.AttributeFilters -maxresults $MaxUserResults

        if ($userDeltaResult -and $userDeltaResult.Changes) {
            $userChangeCount = $userDeltaResult.Changes.Count
        }

        Write-Log "User delta changes returned: $userChangeCount"
    }

    # -------- GROUP DELTA --------
    if ($config.SourceObjectType -in @('Group','Both')) {

        #$groupDeltaResult = Get-GroupDeltaChanges -AccessToken $graphToken -DeltaLink $state.GroupDeltaLink
        #$groupDeltaResult = Get-GroupDeltaChanges -AccessToken $graphToken -DeltaLink $state.GroupDeltaLink -AttributeFilters $config.AttributeFilters
        $groupDeltaResult = Get-GroupDeltaChanges -AccessToken $graphToken -DeltaLink $state.GroupDeltaLink -AttributeFilters $config.AttributeFilters -MaxResults $MaxGroupResults

        if ($groupDeltaResult -and $groupDeltaResult.Changes) {
            $groupChangeCount = $groupDeltaResult.Changes.Count
        }

        Write-Log "Group delta changes returned: $groupChangeCount"
    }

    # -------- EARLY EXIT --------
    if ($userChangeCount -eq 0 -and $groupChangeCount -eq 0) {

        if (-not $RunReconciliation) {

            Write-Log "No changes detected (users+groups). Exiting before Exchange Online connection."

            # Update delta links if present
            if ($userDeltaResult -and $userDeltaResult.PSObject.Properties['DeltaLink'] -and $userDeltaResult.DeltaLink) {
                $state.UserDeltaLink = $userDeltaResult.DeltaLink
            }

            if ($groupDeltaResult -and $groupDeltaResult.PSObject.Properties['DeltaLink'] -and $groupDeltaResult.DeltaLink) {
                $state.GroupDeltaLink = $groupDeltaResult.DeltaLink
            }

            Save-State -Path $stateFile -State $state
            Write-Log "State file updated. No processing required."

            return
        }
        else {
            Write-Log "No delta changes, but reconciliation is required — continuing execution" "WARN"
        }
}

    # ---------------- HYBRID LOOKUP MODE ----------------

    $TotalChangeCount = $userChangeCount + $groupChangeCount
    $UseBulkLookup = $false

    # Threshold — tune this (10–50 recommended)
    $BulkThreshold = 25
    
    if ($MaxUserResults -gt 0 -or $MaxGroupResults -gt 0) {
        Write-Log "Test mode detected — forcing BULK lookup mode" "WARN"
        $UseBulkLookup = $true
    }
    elseif ($TotalChangeCount -ge $BulkThreshold) {
        $UseBulkLookup = $true
        Write-Log "Using BULK contact preload mode (changes: $TotalChangeCount)" "INFO"
    }
    else {
        Write-Log "Using ON-DEMAND lookup mode (changes: $TotalChangeCount)" "INFO"
    }
    
    #STOP

    Connect-TargetExchange -Config $config
    Write-Log "Connected to Exchange Online target tenant."

    $existingContacts = $null
    $existingContactsByEmail = $null

    if ($SeedTargetFromSource -ne 'None') { 
        Write-Log "Seed mode active — skipping target contact preload" "WARN"
        $existingContacts = $null 
    } 
    elseif ($UseBulkLookup) { 
        $existingContacts = Get-ExistingTargetContacts -SourceTenantId $config.SourceTenantId 
        Write-Log "Loaded $($existingContacts.Count) contacts (bulk mode)"
        
        $existingContactsByEmail = @{}
        $seenContactIds = @{}

        if ($existingContacts) {
            foreach ($c in $existingContacts.Values) {

                if (-not $c.Identity) { continue }

                $identityKey = $c.Identity.ToString()
                if ($seenContactIds.ContainsKey($identityKey)) { continue }
                $seenContactIds[$identityKey] = $true

                if ($c.ExternalEmailAddress) {
                    $email = $c.ExternalEmailAddress.ToString().ToLower().Replace("smtp:","")

                    if (-not $existingContactsByEmail.ContainsKey($email)) {
                        $existingContactsByEmail[$email] = @()
                    }

                    $existingContactsByEmail[$email] += $c
                }
            }

            Write-Log "Built email index: $($existingContactsByEmail.Count) unique addresses" "DEBUG"
        }
         
    } else { 
        Write-Log "Using ON-DEMAND lookup mode — no preload required" "INFO" 
        $existingContacts = $null 
    }
    
    $createdOrUpdated = 0
    $deleted = 0
    $skipped = 0

    $ProcessedCount = 0

    # USER MAIN PROCESSING LOOP
    if ($config.SourceObjectType -in @('User','Both')) {

    foreach ($item in $userDeltaResult.Changes) {

        if ($TopUsers -gt 0 -and $ProcessedCount -ge $TopUsers) {
            Write-Log "TopUsers limit reached ($TopUsers). Stopping processing." "WARN"
            break
        }

        $syncKey = Get-SyncKey -SourceTenantId $config.SourceTenantId -SourceObjectId $item.id -SourceObjectType 'User'

        # -------------------- RESOLVE EXISTING --------------------
        $existing = $null

        if ($SeedTargetFromSource -in @('User','Both')) {
            $existing = $null
        }
        elseif ($UseBulkLookup) {
            if ($existingContacts -and $existingContacts.ContainsKey($syncKey)) {
                $existing = $existingContacts[$syncKey]
            }
        }
        else {
            $existing = Get-TargetContactBySyncKey -SyncKey $syncKey
        }

        # -------------------- DELTA DELETE --------------------
        if ($item.PSObject.Properties.Name -contains '@removed') {

            if ($SeedTargetFromSource -ne 'None') {
                Write-Log "Skipping delete (seed mode): $syncKey" "WARN"
                continue
            }

            $deleteResult = Remove-TargetMailContactBySyncKey `
                -SyncKey $syncKey `
                -ExistingContacts $existingContacts `
                -DisableDeletes:$config.DisableDeletes

            if ($deleteResult -eq "Deleted") { $deleted++ }
            continue
        }

        # -------------------- BASIC OBJECT INFO --------------------
        $displayName = '<no displayName>'
        if ($item.displayName) { $displayName = $item.displayName }

        $upn = Get-SafeUpn -User $item

        if (-not $upn) {
            Write-Log "Skipping object with missing UPN: [$($item.id)] displayName='$displayName'" 'WARN'
            $skipped++
            continue
        }

        if ($upn -like "*#EXT#*") {
            Write-Log "Skipping guest user: $upn" "WARN"
            $skipped++
            continue
        }

        # -------------------- SERVICE ACCOUNT FILTER --------------------
        $isServiceAccount = $false
        foreach ($pattern in $servicePatterns) {
            if ($upn -match $pattern) {
                Write-Log "Skipping service/sync account: $upn" 'DEBUG'
                $isServiceAccount = $true
                break
            }
        }
        if ($isServiceAccount) {
            $skipped++
            continue
        }

        # =========================================================
        # FIX: SAFE PROPERTY CHECK + HYDRATION
        # =========================================================

        # Generic USER filter debug driven by XML (restricted to onPremisesExtensionAttributes by validation)
        foreach ($filter in @($config.AttributeFilters)) {

            if (-not $filter) { continue }

            $filterObjectType = if ($filter.ObjectType) { [string]$filter.ObjectType } else { 'Both' }

            if ($filterObjectType -notin @('User','Both')) {
                continue
            }

            $propertyPath = [string]$filter.PropertyPath
            if (-not $propertyPath) { continue }

            $value = Get-NestedPropertyValue `
                -InputObject $item `
                -PropertyPath $propertyPath

            Write-Log "DEBUG USER FILTER: Property='$propertyPath' Value='$value'" "DEBUG"
        }

        # =========================================================
        # ATTRIBUTE FILTER
        # =========================================================

        $matchedUserAttributeFilter = Test-AttributeFilterMatch `
            -InputObject $item `
            -AttributeFilters $config.AttributeFilters `
            -ObjectType 'User'

        if ($matchedUserAttributeFilter) {

            Write-Log "FILTER MATCHED: user excluded by AttributeFilters for $upn" "WARN"

            if ($existing -and $SeedTargetFromSource -eq 'None') {

                Write-Log "Deleting contact due to AttributeFilter: $syncKey" "WARN"

                $deleteResult = Remove-TargetMailContactBySyncKey `
                    -SyncKey $syncKey `
                    -ExistingContacts $existingContacts `
                    -DisableDeletes:$config.DisableDeletes

                if ($deleteResult -eq "Deleted") {
                    $deleted++
                }
                else {
                    $skipped++
                }
            }
            else {
                Write-Log "Filter matched but no existing contact found: $syncKey" "DEBUG"
                $skipped++
            }

            continue
        }

        # -------------------- DOMAIN SCOPE CHECK --------------------

        if (-not (Test-UserInScope `
            -User $item `
            -IncludeDomains $config.IncludeDomains `
            -ExcludeDomains $config.ExcludeDomains `
            -AllowUpnFallback:$config.AllowUpnFallback)) {

            $skipped++
            continue
        }

        # -------------------- NORMAL PROCESSING --------------------

        $ProcessedCount++

        $result = Set-TargetMailContact `
            -User $item `
            -Config $config `
            -ExistingContacts $existingContacts `
            -ExistingContact $existing

        switch ($result) {
            "Created"   { $createdOrUpdated++ }
            "Updated"   { $createdOrUpdated++ }
            "Unchanged" { }
            "Skipped"   { $skipped++ }
        }
    }
}

    # GROUP MAIN PROCESSING LOOP
    if ($config.SourceObjectType -in @('Group','Both')) {

    foreach ($group in $groupDeltaResult.Changes) {

        if ($TopUsers -gt 0 -and $ProcessedCount -ge $TopUsers) {
            Write-Log "TopUsers limit reached ($TopUsers). Stopping processing." "WARN"
            break
        }

        $debugMailEnabled = $null 
        $debugSecurityEnabled = $null 
        $debugGroupTypes = $null 
        if ($group.PSObject.Properties.Name -contains 'mailEnabled') {
            $debugMailEnabled = $group.mailEnabled 
        } 
        
        if ($group.PSObject.Properties.Name -contains 'securityEnabled') {
            $debugSecurityEnabled = $group.securityEnabled 
        } 
        
        if ($group.PSObject.Properties.Name -contains 'groupTypes') { 
            $debugGroupTypes = $group.groupTypes 
        } 
        
        Write-Log "GROUP DEBUG: id=$($group.id) mailEnabled=$debugMailEnabled securityEnabled=$debugSecurityEnabled groupTypes=$debugGroupTypes" "DEBUG"

        $syncKey = Get-SyncKey `
            -SourceTenantId $config.SourceTenantId `
            -SourceObjectId $group.id `
            -SourceObjectType 'Group'

        # -------------------- RESOLVE EXISTING --------------------
        $existing = $null

        if ($SeedTargetFromSource -in @('Group','Both')) {
            $existing = $null
        }
        elseif ($UseBulkLookup) {
            if ($existingContacts -and $existingContacts.ContainsKey($syncKey)) {
                $existing = $existingContacts[$syncKey]
            }
        }
        else {
            $existing = Get-TargetContactBySyncKey -SyncKey $syncKey
        }

        # -------------------- DELTA DELETE --------------------
        if ($group.PSObject.Properties.Name -contains '@removed') {

            Write-Log "DELTA DELETE DETECTED: id=$($group.id) syncKey=$syncKey" "DEBUG"

            if ($SeedTargetFromSource -ne 'None') {
                Write-Log "Skipping delete (seed mode): $syncKey" "WARN"
                $skipped++
                continue
            }

            Write-Log "DELETE REQUEST: type=Group syncKey=$syncKey usingOnDemand=$(-not $existingContacts)" "DEBUG"

            $deleteResult = Remove-TargetMailContactBySyncKey `
                -SyncKey $syncKey `
                -ExistingContacts $existingContacts `
                -DisableDeletes:$config.DisableDeletes

            if ($deleteResult -eq "Deleted") { 
                $deleted++
                Write-Log "DELETE RESULT: syncKey=$syncKey result=$deleteResult" "DEBUG"    
            }
            continue
        }

        Write-Log "Processing group: $($group.id)" "DEBUG"

        # -------------------- SCOPE CHECK --------------------
        if (-not (Test-GroupInScope -Group $group)) {
            Write-Log "Group skipped (out of scope): $($group.id)" "DEBUG"
            $skipped++
            continue
        }

        Write-Log "DEBUG PROXY: $($group.proxyAddresses)" "WARN"

        # Generic GROUP-ONLY filter debug driven by XML
        foreach ($filter in @($config.AttributeFilters)) {

            if (-not $filter) { continue }

            $filterObjectType = if ($filter.ObjectType) { [string]$filter.ObjectType } else { 'Both' }

            # STRICTLY GROUP-ONLY
            if ($filterObjectType -notin @('Group','Both')) {
                continue
            }

            $propertyPath = [string]$filter.PropertyPath
            if (-not $propertyPath) { continue }

            $value = Get-NestedPropertyValue `
                -InputObject $group `
                -PropertyPath $propertyPath

            Write-Log "DEBUG GROUP FILTER: Property='$propertyPath' Value='$value'" "DEBUG"
        }

        # =========================================================
        # ATTRIBUTE FILTER (WITH GENERIC GROUP HYDRATION)
        # =========================================================

        $requiredGroupFilterProps = Get-TopLevelSelectPropertiesFromAttributeFilters `
            -AttributeFilters $config.AttributeFilters `
            -ObjectType 'Group'

        $missingGroupFilterProps = @(
            $requiredGroupFilterProps | Where-Object {
                $group.PSObject.Properties.Name -notcontains $_
            }
        )

        if ($missingGroupFilterProps.Count -gt 0) {

            Write-Log "Group filter properties missing from delta payload for $($group.id): $($missingGroupFilterProps -join ', ')" "DEBUG"

            $groupSelectProps = @(
                'id',
                'displayName',
                'mail',
                'mailEnabled',
                'securityEnabled',
                'groupTypes',
                'proxyAddresses'
            ) + $requiredGroupFilterProps | Select-Object -Unique

            $groupSelect = ($groupSelectProps -join ',')

            try {
                $group = Invoke-GraphJson `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)?`$select=$groupSelect" `
                    -AccessToken $graphToken
            }
            catch {
                Write-Log "Failed to hydrate group for filter check: $($_.Exception.Message)" "WARN"
            }
        }

        $matchedGroupAttributeFilter = Test-AttributeFilterMatch `
            -InputObject $group `
            -AttributeFilters $config.AttributeFilters `
            -ObjectType 'Group'

        if ($matchedGroupAttributeFilter) {

            Write-Log "FILTER MATCHED (GROUP): $($group.displayName) [$syncKey]" "WARN"

            if ($existing -and $SeedTargetFromSource -eq 'None') {

                Write-Log "Deleting group contact due to AttributeFilter: $syncKey" "WARN"

                $deleteResult = Remove-TargetMailContactBySyncKey `
                    -SyncKey $syncKey `
                    -ExistingContacts $existingContacts `
                    -DisableDeletes:$config.DisableDeletes

                if ($deleteResult -eq "Deleted") {
                    $deleted++
                }
                else {
                    $skipped++
                }
            }
            else {
                Write-Log "Filter matched but no existing group contact found: $syncKey" "DEBUG"
                $skipped++
            }

            continue
        }

        # -------------------- NORMAL PROCESSING --------------------

        $ProcessedCount++

        $result = Set-TargetMailContactFromGroup `
            -Group $group `
            -Config $config `
            -ExistingContacts $existingContacts `
            -ExistingContact $existing

        switch ($result) {
            "Created"   { $createdOrUpdated++ }
            "Updated"   { $createdOrUpdated++ }
            "Unchanged" { }
            "Skipped"   { $skipped++ }
        }
    }
}

    # =========================================================
    # BATCH RECONCILIATION (RUN PERIODICALLY)
    # =========================================================
    if ($RunReconciliation) {

        Write-Log "Reconciliation execution block entered" "WARN"

        # Force preload when running in ON-DEMAND mode
        if (-not $existingContacts) {

            Write-Log "Forcing contact preload for reconciliation *** this may take a long time for larger tenants" "WARN"

            $existingContacts = Get-ExistingTargetContacts `
                -SourceTenantId $config.SourceTenantId
        }

        if ($existingContacts -and $existingContacts.Count -gt 0) {

            Invoke-BatchReconciliation `
                -ExistingContacts $existingContacts `
                -Config $config `
                -GraphToken $graphToken `
                -DeletedCounter ([ref]$deleted) `
                -SkippedCounter ([ref]$skipped)
        }
        else {
            Write-Log "Skipping reconciliation: no existing contacts loaded" "WARN"
        }

        # Persist timestamp AFTER successful reconciliation
        $state.LastReconciliationUtc = [DateTime]::UtcNow.ToString("o")
        Write-Log "Reconciliation timestamp updated: $($state.LastReconciliationUtc)" "DEBUG"

        Save-State -Path $stateFile -State $state
        Write-Log "State persisted after reconciliation update" "DEBUG"

    }

    # =========================================================

    if ($TopUsers -eq 0) {

        # ---------------- USER DELTA ----------------
        if ($userDeltaResult -and $userDeltaResult.PSObject.Properties['DeltaLink'] -and $userDeltaResult.DeltaLink) {
            $state.UserDeltaLink = $userDeltaResult.DeltaLink
            #write-Log "User delta link: $($userDeltaResult.DeltaLink)"
        }
        else {
            Write-Log "User delta link not available." "WARN"
        }

        # ---------------- GROUP DELTA (ONLY IF GROUPS ARE IN SCOPE) ----------------
        if ($config.SourceObjectType -in @('Group','Both')) {

            if ($groupDeltaResult -and $groupDeltaResult.PSObject.Properties['DeltaLink'] -and $groupDeltaResult.DeltaLink) {
                $state.GroupDeltaLink = $groupDeltaResult.DeltaLink
                #Write-Log "Group delta link: $($groupDeltaResult.DeltaLink)"
            }
            else {
                Write-Log "Group delta link not available." "WARN"
            }
        }

        Write-Log "Saving state file to: $stateFile"

        if ($IsFirstRun -and -not $state.LastReconciliationUtc) {
            $state.LastReconciliationUtc = [DateTime]::UtcNow.ToString("o")
            Write-Log "Initialized reconciliation timestamp on first run" "DEBUG"
        }

        # ALWAYS persist state unless in test mode
        if ($TopUsers -eq 0) { 
            Save-State -Path $stateFile -State $state 
            Write-Log "Final state persisted" "DEBUG" 
        } else { 
            Write-Log "TopUsers test mode active — skipping state save" "WARN" 
        }

    }
    else {
        Write-Log "TopUsers test mode active OR no deltaLink returned; state not updated." "WARN"
    }


    Write-Log "Sync complete."
    $totalProcessed = $ProcessedCount + $deleted + $skipped
    Write-Log "Summary:" 
    Write-Log " Created/Updated : $createdOrUpdated" 
    Write-Log " Deleted : $deleted" 
    Write-Log " Skipped : $skipped" 
    Write-Log " TotalProcessed : $totalProcessed"
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    throw
}
finally {
    Disconnect-TargetExchangeSafe
}