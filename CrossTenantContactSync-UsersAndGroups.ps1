[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # Optional external config file. If supplied, values here override/augment parameters.
    [string]$ConfigXmlPath="C:\temp\TenantContactSync\TenantAtoTenantBConfig.xml",

    # ---------- Source tenant / Graph ----------
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
    [string]$SourceObjectType = 'User'

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
        'StateRoot','LogRoot','IncludeDomains','ExcludeDomains',
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
        
        <#
        if ($page.value) {
            foreach ($item in $page.value) {
                $allChanges.Add($item)
            }
        }
        #>
        
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

        <#
        if ($page.'@odata.nextLink') {
            $uri = $page.'@odata.nextLink'
        }
        else {
            $uri = $null
            $finalDeltaLink = $page.'@odata.deltaLink'
        }
        #>


    } while ($uri)

    return [pscustomobject]@{
        Changes   = $allChanges
        DeltaLink = $finalDeltaLink
    }
}

function Get-GroupDeltaChanges { 
    param( [Parameter(Mandatory)][string]$AccessToken,
    [Parameter()][string]$DeltaLink )
     
    $select = [System.Web.HttpUtility]::UrlEncode("id,displayName,mail,securityEnabled,mailEnabled,groupTypes") 
    $uri = if ($DeltaLink) { $DeltaLink } else { "https://graph.microsoft.com/v1.0/groups/delta?`$select=$select" } 
    
    $allChanges = @() 
    $finalDeltaLink = $null 
    
    do { 
        $page = Invoke-GraphJson -Uri $uri -AccessToken $AccessToken 
        
        if ($page.value) { $allChanges += $page.value } 
        if ($page.'@odata.nextLink') { $uri = $page.'@odata.nextLink' } 
        else { $uri = $null; $finalDeltaLink = $page.'@odata.deltaLink' } 
        
        } 
        
        while ($uri) 
        
        return [pscustomobject]@{ Changes = $allChanges ; DeltaLink = $finalDeltaLink } 
}

function Test-GroupInScope { param([Parameter(Mandatory)][object]$Group)

    if ($Group.groupTypes -contains "Unified") { return $false }

    return ($Group.securityEnabled -and $Group.mailEnabled)
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
        [Parameter(Mandatory)][hashtable]$ExistingContacts
    )

    $externalEmail = Get-PrimaryExternalEmail -User $User -AllowUpnFallback:$Config.AllowUpnFallback
    if (-not $externalEmail) {
        Write-Log "Skipping user $($User.id) because no mail/UPN could be used." 'WARN'
        return
    }

    $syncKey = Get-SyncKey -SourceTenantId $Config.SourceTenantId -SourceObjectId $User.id -SourceObjectType 'User'
    $legacySyncKey = Get-LegacySyncKey -SourceTenantId $Config.SourceTenantId -SourceObjectId $User.id

    $displayName = Get-DisplayNameForTarget -User $User -AppendSourceTenantToDisplayName:$Config.AppendSourceTenantToDisplayName -SourceTenantName $Config.SourceTenantName
    $alias = New-SafeAlias -Email $externalEmail -SourceObjectId $User.id

    $phone = $null

    $bpProp = $User.PSObject.Properties['businessPhones']
    $mobileProp = $User.PSObject.Properties['mobilePhone']

    if ($bpProp -and $bpProp.Value -and $bpProp.Value.Count -gt 0) {
        $phone = $bpProp.Value[0]
    }
    elseif ($mobileProp -and $mobileProp.Value) {
        $phone = $mobileProp.Value
    }

    if ($ExistingContacts.ContainsKey($syncKey)) {
        $existing = $ExistingContacts[$syncKey]
        $identity = $existing.Identity
        $existingKey = $existing.CustomAttribute15

        if ($PSCmdlet.ShouldProcess($identity, "Update target mail contact")) {
            Set-MailContact -Identity $identity `
                -DisplayName $displayName `
                -ExternalEmailAddress $externalEmail `
                -CustomAttribute15 $syncKey

            if ($existingKey -ne $syncKey) {
                Write-Log "Upgraded legacy sync key on user contact: $existingKey -> $syncKey"
                if ($ExistingContacts.ContainsKey($existingKey)) { $ExistingContacts.Remove($existingKey) | Out-Null }
                if ($ExistingContacts.ContainsKey($legacySyncKey)) { $ExistingContacts.Remove($legacySyncKey) | Out-Null }
                $ExistingContacts[$syncKey] = Get-MailContact -Identity $identity
            }

            Write-Log "Updated contact: $displayName <$externalEmail> [$syncKey]"
        }
    }
    else {
        if ($PSCmdlet.ShouldProcess($displayName, "Create target mail contact")) {

            # need this here to avoid errors on setting names greater than 64 characters when Displayname is long

            
            $shortName = if ($displayName.Length -gt 40) {
                $displayName.Substring(0,40)
            } else {
                $displayName
            }

            $uniqueSuffix = $User.id.Substring(0,8)

            $Name = "$shortName-$uniqueSuffix"

            if ($Name.Length -gt 64) {
                $Name = $Name.Substring(0,64)
            }
            
            $new = New-MailContact `
                -Name $Name `
                -DisplayName $displayName `
                -Alias $alias `
                -ExternalEmailAddress $externalEmail

            
            if (-not $new -or -not $new.Identity) {
                Write-Log "ERROR: Failed to create contact for $displayName" "ERROR"
                return
            }

            Set-MailContact -Identity $new.Identity -CustomAttribute15 $syncKey

            $ExistingContacts[$syncKey] = Get-MailContact -Identity $new.Identity
            Write-Log "Created contact: $displayName <$externalEmail> [$syncKey]"
        }
    }
}

function Set-TargetMailContactFromGroup { 
    param( 
        [Parameter(Mandatory)][object]$Group, 
        [Parameter(Mandatory)][object]$Config, 
        [Parameter(Mandatory)][hashtable]$ExistingContacts 
    ) 
    
    $externalEmail = Get-PrimaryGroupExternalEmail -Group $Group 
    if (-not $externalEmail) { return } 
    
    $syncKey = Get-SyncKey -SourceTenantId $Config.SourceTenantId -SourceObjectId $Group.id -SourceObjectType 'Group' 
    $displayName = Get-DisplayNameForTarget -User $Group -AppendSourceTenantToDisplayName $Config.AppendSourceTenantToDisplayName -SourceTenantName $Config.SourceTenantName
    $alias = New-SafeAlias -Email $externalEmail -SourceObjectId $Group.id 
    
    if ($ExistingContacts.ContainsKey($syncKey)) { 
        $existing = $ExistingContacts[$syncKey] 
        $identity = $existing.Identity 
        $existingKey = $existing.CustomAttribute15 
        
        if ($PSCmdlet.ShouldProcess($identity, "Update group contact")) { 
            Set-MailContact ` 
                -Identity $identity ` 
                -DisplayName $displayName ` 
                -ExternalEmailAddress $externalEmail ` 
                -CustomAttribute15 $syncKey 
                
            if ($existingKey -ne $syncKey) { 
                Write-Log "Upgraded sync key on group contact: $existingKey -> $syncKey" 
                
                if ($ExistingContacts.ContainsKey($existingKey)) { 
                    $ExistingContacts.Remove($existingKey) | Out-Null } 
                    $ExistingContacts[$syncKey] = Get-MailContact -Identity $identity 
                } 
            } 
            
            } else { 
            
                $shortName = if ($displayName.Length -gt 40) { 
                    $displayName.Substring(0,40) 
                    } else { 
                        $displayName 
                    } 
                    
                    $uniqueSuffix = $Group.id.Substring(0,8) 
                    $Name = "$shortName-$uniqueSuffix" 
                    
                    if ($Name.Length -gt 64) { 
                        $Name = $Name.Substring(0,64) 
                    } 
                    
                    $new = New-MailContact ` 
                        -Name $Name ` 
                        -DisplayName $displayName ` 
                        -Alias $alias ` 
                        -ExternalEmailAddress $externalEmail 
                        
                    if (-not $new -or -not $new.Identity) { 
                        Write-Log "ERROR: Failed to create group contact for $displayName" "ERROR" 
                        return 
                    } 
                    
                    Set-MailContact -Identity $new.Identity -CustomAttribute15 $syncKey 
                    $ExistingContacts[$syncKey] = Get-MailContact -Identity $new.Identity 
                } 
            }

function Remove-TargetMailContactBySyncKey {
    param(
        [Parameter(Mandatory)][string]$SyncKey,
        [Parameter(Mandatory)][hashtable]$ExistingContacts,
        [Parameter(Mandatory)][bool]$DisableDeletes
    )

    if ($DisableDeletes) {
        Write-Log "Delete suppressed by DisableDeletes for sync key: $SyncKey" 'WARN'
        return
    }

    if ($ExistingContacts.ContainsKey($SyncKey)) {
        $existing = $ExistingContacts[$SyncKey]
        if ($PSCmdlet.ShouldProcess($existing.Identity, "Remove target mail contact")) {
            Remove-MailContact -Identity $existing.Identity -Confirm:$false
            $ExistingContacts.Remove($SyncKey) | Out-Null
            Write-Log "Removed contact with sync key: $SyncKey"
        }
    }
    else {
        Write-Log "Delete requested but no existing target contact found for sync key: $SyncKey" 'WARN'
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

Ensure-Folder -Path $config.StateRoot
Ensure-Folder -Path $config.LogRoot

$script:LogFile = Join-Path $config.LogRoot ("{0:yyyyMMdd_HHmmss}_{1}_TO_{2}.log" -f (Get-Date), $config.SourceTenantName, $config.TargetTenantName)
New-Item -ItemType File -Path $script:LogFile -Force | Out-Null

$stateFile = Get-StateFilePath -Config $config
$state = Load-State -Path $stateFile

Write-Log "Starting sync: $($config.SourceTenantName) -> $($config.TargetTenantName)"
Write-Log "State file: $stateFile"
Write-Log "Log file  : $script:LogFile"

try {
    $graphToken = Get-GraphToken `
        -TenantId $config.SourceTenantId `
        -ClientId $config.SourceClientId `
        -ClientSecret $config.SourceClientSecret

    Write-Log "Acquired Graph token for source tenant."

    Connect-TargetExchange -Config $config
    Write-Log "Connected to Exchange Online target tenant."

    $existingContacts = Get-ExistingTargetContacts -SourceTenantId $config.SourceTenantId
    Write-Log "Loaded $($existingContacts.Count) existing synced target contacts for source tenant."

    $createdOrUpdated = 0
    $deleted = 0
    $skipped = 0

    $userDeltaResult = $null 
    $groupDeltaResult = $null

    $ProcessedCount = 0

    if ($config.SourceObjectType -in @('User','Both')) { 
        $userDeltaResult = Get-UserDeltaChanges -AccessToken $graphToken -DeltaLink $state.UserDeltaLink 
        
        foreach ($item in $userDeltaResult.Changes) {

            if ($TopUsers -gt 0 -and $ProcessedCount -ge $TopUsers) { 
                Write-Log "TopUsers limit reached ($TopUsers). Stopping processing." "WARN" 
                break 
            } 
            
            $ProcessedCount++
            
            $syncKey = Get-SyncKey -SourceTenantId $config.SourceTenantId -SourceObjectId $item.id -SourceObjectType 'User' 

            $isRemoved = $false
            if ($item.PSObject.Properties.Name -contains '@removed') {
                $isRemoved = $true
            }
            if ($isRemoved) { 
                Remove-TargetMailContactBySyncKey -SyncKey $syncKey -ExistingContacts $existingContacts -DisableDeletes:$config.DisableDeletes 
                continue } 
            
            if (-not (Test-UserInScope -User $item -IncludeDomains $config.IncludeDomains -ExcludeDomains $config.ExcludeDomains -AllowUpnFallback:$config.AllowUpnFallback)) { 
                continue 
                } 
                
                Set-TargetMailContact -User $item -Config $config -ExistingContacts $existingContacts 
            } 
    }

    if ($config.SourceObjectType -in @('Group','Both')) { 
        $groupDeltaResult = Get-GroupDeltaChanges -AccessToken $graphToken -DeltaLink $state.GroupDeltaLink 
        
        foreach ($group in $groupDeltaResult.Changes) { 

            if ($TopUsers -gt 0 -and $ProcessedCount -ge $TopUsers) { 
                Write-Log "TopUsers limit reached ($TopUsers). Stopping processing." "WARN" 
                break 
            } 
            
            $ProcessedCount++
            
            $syncKey = Get-SyncKey -SourceTenantId $config.SourceTenantId -SourceObjectId $group.id -SourceObjectType 'Group' 

            $isRemoved = $false
            if ($group.PSObject.Properties.Name -contains '@removed') {
                $isRemoved = $true
            }

            if ($isRemoved) { 
                Remove-TargetMailContactBySyncKey -SyncKey $syncKey -ExistingContacts $existingContacts -DisableDeletes:$config.DisableDeletes 
                continue } 
                
            if (-not (Test-GroupInScope -Group $group)) { 
                continue 
                } 
                
                Set-TargetMailContactFromGroup -Group $group -Config $config -ExistingContacts $existingContacts 
            } 
    }

    <#
    Write-Log "Graph returned $($deltaResult.Changes.Count) change records."



    # so we can limit number of users
    $ProcessedCount = 0

    foreach ($item in $deltaResult.Changes) {

        # if we hit our limit, stop

        if ($TopUsers -gt 0 -and $processedCount -ge $TopUsers) 
        { 
            Write-Log "TopUsers limit reached ($TopUsers). Stopping processing." "WARN" 
            break
        } 
        
        $processedCount++

        $syncKey = Get-SyncKey -SourceTenantId $config.SourceTenantId -SourceObjectId $item.id

        # Deletion from Graph delta is represented via @removed
        $isRemoved = $false
        if ($item.PSObject.Properties.Name -contains '@removed') {
            $isRemoved = $true
        }

        if ($isRemoved) {
            Remove-TargetMailContactBySyncKey -SyncKey $syncKey -ExistingContacts $existingContacts -DisableDeletes:$config.DisableDeletes
            $deleted++
            continue
        }

        if (-not (Test-UserInScope -User $item -IncludeDomains $config.IncludeDomains -ExcludeDomains $config.ExcludeDomains -AllowUpnFallback:$config.AllowUpnFallback)) {
            Write-Log "Skipping user out of scope or with no usable external address: $($item.userPrincipalName) [$($item.id)]" 'WARN'
            $skipped++
            continue
        }

        Set-TargetMailContact -User $item -Config $config -ExistingContacts $existingContacts
        $createdOrUpdated++
    }
    #>

    if ($TopUsers -eq 0) { 
        if ($userDeltaResult -and $userDeltaResult.DeltaLink) { 
            $state.UserDeltaLink = $userDeltaResult.DeltaLink 
        } 
        
        if ($groupDeltaResult -and $groupDeltaResult.DeltaLink) { 
            $state.GroupDeltaLink = $groupDeltaResult.DeltaLink 
        } 
        
        Save-State -Path $stateFile -State $state 
        
        Write-Log "User delta link: $($userDeltaResult.DeltaLink)" 
        Write-Log "Group delta link: $($groupDeltaResult.DeltaLink)" 
        Write-Log "Saving state file to: $stateFile"

    }
    else { Write-Log "TopUsers test mode active OR no deltaLink returned; state not updated." 'WARN' }

    Write-Log "Sync complete."
    Write-Log "Summary: Created/Updated=$createdOrUpdated Deleted=$deleted Skipped=$skipped"
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    throw
}
finally {
    Disconnect-TargetExchangeSafe
}