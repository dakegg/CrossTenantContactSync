# Tenant Contact Synchronization Script

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-green)
![Exchange Online](https://img.shields.io/badge/Exchange%20Online-AppOnlyAuth-orange)
![License](https://img.shields.io/badge/License-MIT-lightgrey)
![Status](https://img.shields.io/badge/Status-Active-success)

---

## Cross-Tenant Contact Sync using Microsoft Graph + Exchange Online

---

## Overview

This PowerShell solution provides a lightweight, automation-friendly way to synchronize users from a **source Microsoft Entra tenant** into a **target tenant as mail contacts**.

Unlike **Cross-Tenant Synchronization (CTS)**, which provisions users as B2B identities, this script is focused on **Exchange-based visibility and routing**, using contacts instead of guest users.

### Key Capabilities

- Create Mail Contacts in the target tenant
- Handle Adds, Updates, Deletes
- Uses Microsoft Graph (App + Secret) for source
- Uses Exchange Online (App + Certificate) for target
- Maintains state with Graph delta queries
- Multi-tenant support via XML configuration

---

## Architecture

```
Source Tenant (Graph) --> PowerShell Script --> Target Tenant (Exchange)
```

---

## Prerequisites

- PowerShell 5.1+
- ExchangeOnlineManagement module
- App registrations in both tenants
- Certificate installed on execution host

---

## Setup Guide

### Source Tenant App Registration

1. Entra ID -> App registrations -> New
2. Add permissions:
   - User.Read.All
   - Directory.Read.All
   - Group.Read.Aall
3. Create client secret

---

### Target Tenant App Registration

1. Create certificate:

```
New-SelfSignedCertificate -CertStoreLocation Cert:\LocalMachine\My
```

2. Upload .cer file to app registration

3. Add permission:
   - Exchange.ManageAsApp

---

## Execution

### Manual

```
.\TenantContactSync.ps1 -ConfigXmlPath ".\config.xml" -SourceObjectType User
```

```
.\TenantContactSync.ps1 -ConfigXmlPath ".\config.xml" -SourceObjectType Group
```

```
.\TenantContactSync.ps1 -ConfigXmlPath ".\config.xml" -SourceObjectType Both
```

**Parameter Switches ALWAYS take priority over XML settings**


### Test Mode

```
.\TenantContactSync.ps1 -ConfigXmlPath ".\config.xml" -TopUsers 10
```

---

## Troubleshooting

- Auth failures: verify cert access
- Graph failures: verify API permissions
- Duplicate contacts: clean target tenant

---

## License

MIT
