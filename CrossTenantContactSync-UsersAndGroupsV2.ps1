<#
.SYNOPSIS
    Synchronizes users and groups from a source Entra ID tenant to mail contacts in a target Exchange Online tenant using Microsoft Graph delta queries.

.DESCRIPTION
    This script performs incremental (delta-based) synchronization of users and groups from a source tenant into Exchange Online mail contacts 
    in a target tenant. It is designed for cross-tenant collaboration scenarios such as GAL synchronization or external directory visibility.

    The script uses Microsoft Graph delta queries to retrieve only changed objects since the last run. A state file is used to persist deltaLinks, 
    enabling efficient incremental processing.

    Key workflow:

        1. Load configuration (parameters and/or XML file)
        2. Authenticate to Microsoft Graph (source tenant)
        3. Run delta queries for users and/or groups
        4. If no changes detected → exit early (no EXO connection)
        5. Determine processing mode (bulk vs. on-demand lookup)
        6. Connect to Exchange Online (target tenant)
        7. Retrieve existing contacts (bulk mode) OR lookup per object (on-demand)
        8. Process changes:
            - Create new contacts
            - Update existing contacts (only when changes detected)
            - Remove contacts for deleted objects (if enabled)
        9. Persist updated deltaLinks to state file

    The script is optimized for large environments, minimizing Exchange Online calls and avoiding unnecessary updates.

.AUTHOR
    Darryl Kegg

.VERSION
    2.3.0

.CHANGELOG

    2026-06-09 – v2.3.0
        * Added write optimization using Test-MailContactNeedsUpdate (skip no-op updates)
        * Implemented status return model (Created / Updated / Unchanged / Skipped)
        * Enhanced logging to include change reasons for updates
        * Fixed group delta StrictMode crash (@odata.nextLink / deltaLink handling)
        * Made Test-GroupInScope null-safe (handles missing Graph properties)
        * Fixed delete logic to support both bulk and on-demand lookup modes

    2026-06-09 – v2.2.0
        * Implemented hybrid contact lookup model (bulk vs. per-object lookup)
        * Introduced BulkThreshold auto-switching logic
        * Refactored Set-TargetMailContact and Set-TargetMailContactFromGroup to use ExistingContact parameter
        * Removed mandatory dependency on preloaded contact hashtable
        * Eliminated full contact preload requirement for small delta runs
        * Improved performance for low-change scenarios 

    2026-06-09 – v2.1.0
        * Added delta pre-check logic (early exit when no changes detected)
        * Reordered execution flow to avoid unnecessary EXO connection
        * Added safe property access for deltaLink handling (StrictMode-safe)
        * Fixed service account filtering logic (proper loop exit behavior)
        * Improved logging for delta usage and state updates

    2026-06-08 – v2.0.0
        * Unified User + Group processing model
        * Added support for group synchronization (mail-enabled security groups)
        * Introduced syncKey format with object type (User/Group)
        * Implemented retry logic for Exchange Online operations
        * Added name/alias normalization and truncation handling
        * Implemented convergence logic for SMTP conflicts
        * Improved filtering for external users (#EXT#) and service accounts

.PARAMETER ConfigXmlPath
    Optional path to XML configuration file. Values from the file are merged with parameters.

.PARAMETER SourceTenantId
    Azure AD / Entra ID tenant ID of the source environment.

.PARAMETER SourceTenantName
    Friendly name of the source tenant, optionally appended to display names in target.

.PARAMETER SourceClientId
    App registration client ID used to authenticate to Microsoft Graph.

.PARAMETER SourceClientSecret
    Client secret for the Graph app registration.

.PARAMETER TargetOrganization
    Exchange Online organization domain used for connection.

.PARAMETER TargetTenantId
    Tenant ID of the target environment.

.PARAMETER TargetTenantName
    Friendly name of the target tenant.

.PARAMETER TargetExoAppId
    App registration ID used for Exchange Online app-only authentication.

.PARAMETER TargetExoCertThumbprint
    Certificate thumbprint used for Exchange Online authentication.

.PARAMETER StateRoot
    Directory used to store delta state files.

.PARAMETER LogRoot
    Directory used to store execution logs.

.PARAMETER IncludeDomains
    Optional list of email domains to include during processing. If defined, only objects with matching domains are processed.

.PARAMETER ExcludeDomains
    Optional list of email domains to exclude during processing.

.PARAMETER AppendSourceTenantToDisplayName
    If true, appends the source tenant name to the DisplayName of created contacts in the target tenant.

.PARAMETER AllowUpnFallback
    If true, uses UserPrincipalName as the ExternalEmailAddress when the mail attribute is not populated.

.PARAMETER DisableDeletes
    If true, prevents deletion of contacts in the target tenant even when source objects are deleted.

.PARAMETER TopUsers
    Limits number of processed objects for testing scenarios. When greater than 0, state file is not updated.

.PARAMETER SourceObjectType
    Specifies which object types to process:
        User  - Process users only
        Group - Process groups only
        Both  - Process both users and groups

.PARAMETER ForceFullSync
    Forces a full synchronization by ignoring stored deltaLinks for the specified object type(s).
    This causes Microsoft Graph to return the full dataset instead of only incremental changes.

    Values:
        None  - Use normal delta behavior (default)
        User  - Reset user deltaLink only
        Group - Reset group deltaLink only
        Both  - Reset both user and group deltaLinks

    Notes:
        - Does NOT delete the state file on disk
        - Can be combined with SeedTargetFromSource for first-run scenarios

.PARAMETER SeedTargetFromSource
    Enables seed mode (initial population mode) for the specified object type(s).
    In this mode, ALL target Exchange Online reads are skipped, and only create operations are performed.

    Behavior:
        - Skips Get-MailContact preload (bulk mode)
        - Skips per-object lookups
        - Disables delete operations
        - Processes objects as "create-only"
        - Relies on Exchange Online to enforce uniqueness (SMTP conflict handling still applies)

    Values:
        None  - Normal sync behavior (default)
        User  - Seed users only
        Group - Seed groups only
        Both  - Seed both users and groups

    Intended use:
        - First-run initialization
        - Rebuilding target tenant contacts
        - Performance-optimized bulk creation scenarios

    Warning:
        - Does not detect existing contacts prior to creation attempts
        - May trigger SMTP conflict handling logic if duplicates already exist
.EXAMPLE
    .\CrossTenantContactSync.ps1 -ConfigXmlPath "C:\Config\SyncConfig.xml"

.EXAMPLE
    .\CrossTenantContactSync.ps1 -SourceObjectType User -TopUsers 10

.EXAMPLE Full rebuild of groups (no delta, no lookup, fastest mode) 
    .\Script.ps1 -ForceFullSync Group -SeedTargetFromSource Group 
    
.EXAMPLE Normal incremental sync 
    .\Script.ps1 
    
.EXAMPLE Force full resync using existing lookup logic 
    .\Script.ps1 -ForceFullSync Both

.NOTES
    Requirements:
        - ExchangeOnlineManagement module
        - Graph API permissions (User.Read.All, Group.Read.All)
        - Certificate-based EXO authentication

    Behavior:
        - Uses delta state tracking
        - Early exit when no changes
        - Hybrid lookup for performance optimization
        - StrictMode enforced for reliability

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

    # Source object type
    [ValidateSet('User','Group','Both')]
    [string]$SourceObjectType = 'Both',
    [ValidateSet('None','User','Group','Both')]
    [string]$ForceFullSync = 'Group',
    [ValidateSet('None','User','Group','Both')]
    [string]$SeedTargetFromSource = 'None'

)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -------------------- Utility / logging --------------------

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
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
        'StateRoot','LogRoot','IncludeDomains','ExcludeDomains','ServiceAccountPatterns',
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

    $resp = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType 'application/x-www-form-urlencoded'
    return $resp.access_token
}

function Invoke-GraphJson {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$AccessToken
    )

    $headers = @{
        Authorization     = "Bearer $AccessToken"
        ConsistencyLevel  = "eventual"
    }

    return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers
}

function Get-UserDeltaChanges {
    param(
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter()][string]$DeltaLink
    )

    # Track the fields we actually use
    $select = [System.Web.HttpUtility]::UrlEncode("id,displayName,givenName,surname,mail,userPrincipalName,companyName,department,jobTitle,businessPhones,mobilePhone,officeLocation")
    $uri = if ($DeltaLink) {
        $DeltaLink
    }
    else {
        "https://graph.microsoft.com/v1.0/users/delta?`$select=$select"
    }

    $allChanges = New-Object System.Collections.Generic.List[object]
    $finalDeltaLink = $null

    do {
        #this makes the log REALLY chatty
        #Write-Log "Calling Graph delta page: $uri" 'DEBUG'
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
    } while ($uri)

    return [pscustomobject]@{
        Changes   = $allChanges
        DeltaLink = $finalDeltaLink
    }
}

function Get-GroupDeltaChanges {

    param(
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter()][string]$DeltaLink
    )

    $select = [System.Web.HttpUtility]::UrlEncode("id,displayName,mail,securityEnabled,mailEnabled,groupTypes")

    $uri = if ($DeltaLink) {
        $DeltaLink
    }
    else {
        "https://graph.microsoft.com/v1.0/groups/delta?`$select=$select"
    }

    $allChanges = New-Object System.Collections.Generic.List[object]
    $finalDeltaLink = $null

    do {

        $page = Invoke-GraphJson -Uri $uri -AccessToken $AccessToken

        # ---------------- VALUE ----------------
        $valueProp = $page.PSObject.Properties['value']
        if ($valueProp -and $valueProp.Value) {
            foreach ($item in $valueProp.Value) {
                $allChanges.Add($item)
            }
        }

        # ---------------- NEXT LINK ----------------
        $nextLinkProp  = $page.PSObject.Properties['@odata.nextLink']
        $deltaLinkProp = $page.PSObject.Properties['@odata.deltaLink']

        if ($nextLinkProp) {
            $uri = $nextLinkProp.Value
        }
        elseif ($deltaLinkProp) {
            $uri = $null
            $finalDeltaLink = $deltaLinkProp.Value
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

    # SAFE property access
    $groupTypesProp = $Group.PSObject.Properties['groupTypes']
    $securityProp   = $Group.PSObject.Properties['securityEnabled']
    $mailProp       = $Group.PSObject.Properties['mailEnabled']

    # Extract values safely
    if ($groupTypesProp) {
        $groupTypes = $groupTypesProp.Value
    }
    else {
        $groupTypes = @()
    }

    if ($securityProp) {
        $securityEnabled = $securityProp.Value
    }
    else {
        $securityEnabled = $true
    }

    if ($mailProp) {
        $mailEnabled = $mailProp.Value
    }
    else {
        $mailEnabled = $true
    }

    # Skip Microsoft 365 groups only if explicitly marked
    if ($groupTypes -and ($groupTypes -contains "Unified")) {
        return $false
    }

    # Skip only if explicitly invalid
    if (-not ($securityEnabled -and $mailEnabled)) {
        return $false
    }

    return $true
}

function Get-PrimaryGroupExternalEmail { 
    param([Parameter(Mandatory)][object]$Group)

    if ($Group.mail -and $Group.mail.Trim()) {
         return $Group.mail.ToLower()
    }

    return $null

}

# -------------------- Filtering / normalization --------------------

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
    $legacySyncKey = Get-LegacySyncKey -SourceTenantId $Config.SourceTenantId -SourceObjectId $User.id

    $displayName = Get-DisplayNameForTarget -User $User -AppendSourceTenantToDisplayName:$Config.AppendSourceTenantToDisplayName -SourceTenantName $Config.SourceTenantName
    $alias = New-SafeAlias -Email $externalEmail -SourceObjectId $User.id

    # -------------------- HYBRID LOOKUP --------------------
    $existing = $ExistingContact

    if ($existing) {

        $identity = $existing.Identity
        $existingKey = $existing.CustomAttribute15

        $needsUpdate = $false

        if ($existing.DisplayName -ne $displayName) { $needsUpdate = $true }

        $existingEmail = $existing.ExternalEmailAddress.ToString().ToLower().Replace("smtp:","")
        if ($existingEmail -ne $externalEmail.ToLower()) { $needsUpdate = $true }

        if ($existingKey -ne $syncKey) { $needsUpdate = $true }

        if (-not $needsUpdate) {
            Write-Log "No changes required: $displayName <$externalEmail>" "DEBUG"
            return "Unchanged"
        }

        if ($PSCmdlet.ShouldProcess($identity, "Update target mail contact")) {

            Write-Log "Updating contact: Identity=$identity Email=$externalEmail SyncKey=$syncKey"

            Invoke-ExoWithRetry {
                Set-MailContact -Identity $identity -DisplayName $displayName -ExternalEmailAddress $externalEmail -CustomAttribute15 $syncKey
            }

            Write-Log "Updated contact: $displayName <$externalEmail> [$syncKey]"
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

            Write-Log "Creating contact: Name=$Name DisplayName=$displayName Alias=$alias Email=$externalEmail SyncKey=$syncKey"

            $new = $null

            try {
                $new = Invoke-ExoWithRetry {
                    New-MailContact -Name $Name -DisplayName $displayName -Alias $alias -ExternalEmailAddress $externalEmail
                }
            }
            catch {

                $msg = $_.Exception.Message

                # -------------------- SMTP CONVERGENCE --------------------
                if ($msg -like "*proxy address*already being used*") {

                    Write-Log "SMTP conflict detected. Attempting convergence: $externalEmail" "WARN"

                    $existingByEmail = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$externalEmail'" -ErrorAction SilentlyContinue

                    if (-not $existingByEmail) {
                        Write-Log "ERROR: Could not locate existing object for $externalEmail" "ERROR"
                        throw
                    }

                    if ($existingByEmail.RecipientType -ne "MailContact") {
                        Write-Log "ERROR: SMTP belongs to non-MailContact object: $($existingByEmail.RecipientType)" "ERROR"
                        return
                    }

                    if ($existingByEmail.CustomAttribute15 -and $existingByEmail.CustomAttribute15 -ne $syncKey) {
                        Write-Log "WARNING: Existing contact already has different syncKey" "WARN"
                        throw
                    }

                    $identity = $existingByEmail.Identity

                    Invoke-ExoWithRetry {
                        Set-MailContact -Identity $identity -DisplayName $displayName -CustomAttribute15 $syncKey
                    }

                    Write-Log "Convergence complete: $externalEmail -> [$syncKey]"
                    return "Updated"
                }

                throw
            }

            if (-not $new -or -not $new.Identity) {
                Write-Log "ERROR: Failed to create contact for $displayName" "ERROR"
                return "Skipped"
            }

            Invoke-ExoWithRetry {
                Set-MailContact -Identity $new.Identity -CustomAttribute15 $syncKey
            }

            Write-Log "Created contact: $displayName <$externalEmail> [$syncKey]"
            Return "Created"
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
    # UPDATE EXISTING
    # ======================================================
    if ($existing) {

        $identity    = $existing.Identity
        $existingKey = $existing.CustomAttribute15

        $needsUpdate = $false

        if ($existing.DisplayName -ne $displayName) { $needsUpdate = $true }

        $existingEmail = $existing.ExternalEmailAddress.ToString().ToLower().Replace("smtp:","")
        if ($existingEmail -ne $externalEmail.ToLower()) { $needsUpdate = $true }

        if ($existingKey -ne $syncKey) { $needsUpdate = $true }

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
            Return "Updated"
        }
    }

    # ======================================================
    # CREATE NEW
    # ======================================================
    else {

        if ($PSCmdlet.ShouldProcess($displayName, "Create group contact")) {

            # -------------------- NAME SAFETY --------------------
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
            Return "Created"
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

            Invoke-ExoWithRetry {
                Remove-MailContact -Identity $existing.Identity -Confirm:$false
            }

            Write-Log "Removed contact with sync key: $SyncKey"
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
            
            if ($i -ge $MaxRetries) { 
                Write-Log "EXO operation failed after $MaxRetries attempts: $msg" "ERROR" 
                throw 
            } 

            $delay = $InitialDelaySeconds * [math]::Pow(2, ($i - 1))

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

Assert-Config -Config $config

Ensure-Folder -Path $config.StateRoot
Ensure-Folder -Path $config.LogRoot

$script:LogFile = Join-Path $config.LogRoot ("{0:yyyyMMdd_HHmmss}_{1}_TO_{2}.log" -f (Get-Date), $config.SourceTenantName, $config.TargetTenantName)
New-Item -ItemType File -Path $script:LogFile -Force | Out-Null

$stateFile = Get-StateFilePath -Config $config
$state = Load-State -Path $stateFile

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

        $userDeltaResult = Get-UserDeltaChanges -AccessToken $graphToken -DeltaLink $state.UserDeltaLink

        if ($userDeltaResult -and $userDeltaResult.Changes) {
            $userChangeCount = $userDeltaResult.Changes.Count
        }

        Write-Log "User delta changes returned: $userChangeCount"
    }

    # -------- GROUP DELTA --------
    if ($config.SourceObjectType -in @('Group','Both')) {

        $groupDeltaResult = Get-GroupDeltaChanges -AccessToken $graphToken -DeltaLink $state.GroupDeltaLink

        if ($groupDeltaResult -and $groupDeltaResult.Changes) {
            $groupChangeCount = $groupDeltaResult.Changes.Count
        }

        Write-Log "Group delta changes returned: $groupChangeCount"
    }

    # -------- EARLY EXIT --------
    if ($userChangeCount -eq 0 -and $groupChangeCount -eq 0) {

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

    # ---------------- HYBRID LOOKUP MODE ----------------

    $TotalChangeCount = $userChangeCount + $groupChangeCount
    $UseBulkLookup = $false

    # Threshold — tune this (10–50 recommended)
    $BulkThreshold = 25

    if ($TotalChangeCount -ge $BulkThreshold) {
        $UseBulkLookup = $true
        Write-Log "Using BULK contact preload mode (changes: $TotalChangeCount)" "INFO"
    }
    else {
        Write-Log "Using ON-DEMAND lookup mode (changes: $TotalChangeCount)" "INFO"
    }
    
    Connect-TargetExchange -Config $config
    Write-Log "Connected to Exchange Online target tenant."

    $existingContacts = $null

    if ($UseBulkLookup -and $SeedTargetFromSource -eq 'None') {
        $existingContacts = Get-ExistingTargetContacts -SourceTenantId $config.SourceTenantId
        Write-Log "Loaded $($existingContacts.Count) contacts (bulk mode)"
    }
    else {
        Write-Log "Seed mode active — skipping target contact preload" "WARN"
        $existingContacts = $null
    }


    $createdOrUpdated = 0
    $deleted = 0
    $skipped = 0

    $ProcessedCount = 0

    if ($config.SourceObjectType -in @('User','Both')) { 
        # DID THIS ABOVE $userDeltaResult = Get-UserDeltaChanges -AccessToken $graphToken -DeltaLink $state.UserDeltaLink 
        
        foreach ($item in $userDeltaResult.Changes) {

            if ($TopUsers -gt 0 -and $ProcessedCount -ge $TopUsers) { 
                Write-Log "TopUsers limit reached ($TopUsers). Stopping processing." "WARN" 
                break 
            } 
 
            $syncKey = Get-SyncKey -SourceTenantId $config.SourceTenantId -SourceObjectId $item.id -SourceObjectType 'User' 

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

            $isRemoved = $false
            if ($item.PSObject.Properties.Name -contains '@removed') {
                $isRemoved = $true
            }
            if ($isRemoved) {

                if ($SeedTargetFromSource -ne 'None') {
                    Write-Log "Skipping delete (seed mode): $syncKey" "WARN"
                    continue
                }

                $deleteResult = Remove-TargetMailContactBySyncKey -SyncKey $syncKey -ExistingContacts $existingContacts -DisableDeletes:$config.DisableDeletes
                if ($deleteResult -eq "Deleted") { $deleted++ }
                continue
            }
            
            $displayName = '<no displayName>' 
            $dnProp = $item.PSObject.Properties['displayName'] 
            if ($dnProp -and $dnProp.Value -and $dnProp.Value.ToString().Trim()) { $displayName = $dnProp.Value }

            $upn = Get-SafeUpn -User $item

            if (-not $upn) {
                Write-Log "Skipping object with missing UPN: [$($item.id)] displayName='$displayName'" 'WARN'
                $skipped++
                continue
            }

            if ($upn -like "*#EXT#*") { 
            
                Write-Log "Skipping user $upn because it's a Guest \ External Member user in the source tenant" "WARN"
                $skipped++
                continue 
            }

            $isServiceAccount = $false

            foreach ($pattern in $servicePatterns) {
                if ($upn -match $pattern) {
                    Write-Log "Skipping service/sync account (pattern: $pattern): $upn" 'DEBUG'
                    $isServiceAccount = $true
                    break
                }
            }

            if ($isServiceAccount) { $skipped++; continue }
            
           
            if (-not (Test-UserInScope -User $item -IncludeDomains $config.IncludeDomains -ExcludeDomains $config.ExcludeDomains -AllowUpnFallback:$config.AllowUpnFallback)) { 
                $skipped++
                continue 
                } 
                
                $ProcessedCount++

                $result = Set-TargetMailContact -User $item -Config $config -ExistingContacts $existingContacts -ExistingContact $existing

                switch ($result) { 
                    "Created" { $createdOrUpdated++ } 
                    "Updated" { $createdOrUpdated++ } 
                    "Unchanged" { } 
                    "Skipped" { $skipped++ } 
                }
            } 
    }

    if ($config.SourceObjectType -in @('Group','Both')) { 
        # DID THIS ABOVE  $groupDeltaResult = Get-GroupDeltaChanges -AccessToken $graphToken -DeltaLink $state.GroupDeltaLink 
        
        foreach ($group in $groupDeltaResult.Changes) { 

            if ($TopUsers -gt 0 -and $ProcessedCount -ge $TopUsers) { 
                Write-Log "TopUsers limit reached ($TopUsers). Stopping processing." "WARN" 
                break 
            } 
            
            #$ProcessedCount++
            
            $syncKey = Get-SyncKey -SourceTenantId $config.SourceTenantId -SourceObjectId $group.id -SourceObjectType 'Group' 

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

            $isRemoved = $false
            if ($group.PSObject.Properties.Name -contains '@removed') {
                $isRemoved = $true
            }

            if ($isRemoved) {

                if ($SeedTargetFromSource -ne 'None') {
                    Write-Log "Skipping delete (seed mode): $syncKey" "WARN"
                    $skipped++
                    continue
                }

                $deleteResult = Remove-TargetMailContactBySyncKey -SyncKey $syncKey -ExistingContacts $existingContacts -DisableDeletes:$config.DisableDeletes
                if ($deleteResult -eq "Deleted") { $deleted++ }
                continue
            }

            Write-Log "Processing group: $($group.id)" "DEBUG"
                            
            if (-not (Test-GroupInScope -Group $group)) { 
                Write-Log "Group skipped (out of scope): $($group.id)" "DEBUG"
                $skipped++
                continue 
                } 

                $ProcessedCount++
                
                $result = Set-TargetMailContactFromGroup -Group $group -Config $config -ExistingContacts $existingContacts -ExistingContact $existing

                switch ($result) { 
                    "Created" { $createdOrUpdated++ } 
                    "Updated" { $createdOrUpdated++ } 
                    "Unchanged" { } 
                    "Skipped" { $skipped++ } 
                }
            } 
    }

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

        if ($ForceFullSync -eq 'None') {
            Save-State -Path $stateFile -State $state
        }
        else {
            Write-Log "ForceFullSync active — skipping state save" "WARN"
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