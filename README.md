<div align="center">

# 🌍 M365 Multi-Geo Manager

**Automate Microsoft 365 Multi-Geo `UsageLocation` assignments and license provisioning at scale.**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-0078D4?logo=microsoft&logoColor=white)](https://learn.microsoft.com/en-us/graph/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.0.0-orange)](https://github.com/SouhaielMorhag/M365-MultiGeo-Manager/releases)
[![Blog](https://img.shields.io/badge/Blog-msendpoint.com-blueviolet)](https://msendpoint.com)

</div>

---

## 📋 Overview

In a Microsoft 365 Multi-Geo environment, every user must have a `UsageLocation` set **before** data-residency-sensitive services (Exchange Online, SharePoint, Teams) are provisioned in the correct geography. Managing this across hundreds of accounts is tedious and error-prone.

**M365 Multi-Geo Manager** solves this with a single PowerShell script that:

- 🔍 **Analyses** every user in the tenant and detects their correct region from `OfficeLocation` and `City` attributes
- 📍 **Updates** `UsageLocation` only where it differs from the target
- 🔑 **Assigns** the Multi-Geo license (any SKU — picked interactively) to users who are missing it
- 🔒 **Handles disabled accounts** (Teams Phone resource accounts) — location fix only, no license wasted
- 🧪 **DryRun mode** — generates a full HTML compliance report before touching anything
- 📊 **Reports** — dark-themed HTML dashboard + CSV + timestamped log after every run

---

## ✨ Features

| Feature | Description |
|---|---|
| **Interactive SKU picker** | Lists all tenant SKUs with seat counts after login — no hardcoded values |
| **Dynamic region engine** | 14-country keyword mapping, fully customisable |
| **DryRun / Preview mode** | Full HTML report showing planned changes — nothing applied |
| **Partial success tracking** | Location and license operations tracked independently — one failure never hides the other |
| **Disabled account support** | Resource accounts (Auto Attendants, Call Queues) get location fixed, never licensed |
| **HTML compliance dashboard** | Dark-themed cards with action badges, compliance notes, KPI bar |
| **CSV + log export** | Timestamped files after every run |
| **CI / automation ready** | Pass `-SkuPartNumber` to skip the interactive picker |
| **Zero hardcoded config** | Domain, SKU, region, output path — all parameters |

---

## 🚀 Quick Start

### Prerequisites

```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force
```

> The script installs missing modules automatically if run interactively.

### 1 — Preview first (always recommended)

```powershell
.\Set-MultiGeoUsageLocation.ps1 -Domain "contoso.com" -DryRun
```

Opens a browser for Global Admin authentication, then generates a full HTML compliance report showing exactly what would change. **Nothing is applied.**

### 2 — Live execution

```powershell
.\Set-MultiGeoUsageLocation.ps1 -Domain "contoso.com"
```

### 3 — CI / automation (skip interactive prompts)

```powershell
.\Set-MultiGeoUsageLocation.ps1 -Domain "contoso.com" -SkuPartNumber "OFFICE365_MULTIGEO"
```

### 4 — Location only (no license assignment)

```powershell
.\Set-MultiGeoUsageLocation.ps1 -Domain "contoso.com" -SkipLicenseAssignment
```

---

## 📦 Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-Domain` | ✅ | — | Tenant domain to filter users (`contoso.com`) |
| `-SkuPartNumber` | ❌ | — | SKU to assign. If omitted, interactive picker is shown |
| `-DefaultLocation` | ❌ | `FR` | Fallback ISO code for users with undetectable region |
| `-ExcludeUsers` | ❌ | `@()` | Array of UPNs to skip entirely |
| `-DryRun` | ❌ | — | Preview mode — full HTML report, no changes applied |
| `-OutputPath` | ❌ | Script dir | Folder for HTML / CSV / log output |
| `-SkipLicenseAssignment` | ❌ | — | Update UsageLocation only, skip license |

---

## 🗺️ Region Detection Engine

The script detects the target country from each user's `OfficeLocation` and `City` attributes using a configurable keyword map:

| ISO Code | Country | Keywords matched |
|---|---|---|
| `FR` | France | paris, lyon, bordeaux, france… |
| `CA` | Canada | montréal, toronto, vancouver, canada… |
| `US` | United States | memphis, new york, chicago, usa… |
| `GB` | United Kingdom | london, manchester, england… |
| `DE` | Germany | frankfurt, berlin, germany… |
| `SG` | Singapore | singapore |
| `CN` | China | shanghai, beijing, china… |
| `JP` | Japan | tokyo, osaka, japan… |
| `AU` | Australia | sydney, melbourne, australia… |
| `IN` | India | mumbai, bangalore, india… |
| `BR` | Brazil | são paulo, rio, brazil… |
| `ES` | Spain | madrid, barcelona, spain… |
| `IT` | Italy | rome, milan, italy… |
| `NL` | Netherlands | amsterdam, rotterdam, netherlands… |

**Add your own regions** by editing the `$RegionRules` hashtable at the top of the script:

```powershell
$RegionRules = [ordered]@{
    "FR" = "paris|france"
    "ZA" = "johannesburg|cape town|south africa"   # ← add your own
    # ...
}
```

---

## 🔐 Disabled Accounts — Teams Phone Resource Accounts

Auto Attendants and Call Queues create resource accounts in a **disabled state**. These accounts:
- ✅ **Need `UsageLocation`** set for Teams Phone routing to work correctly
- ❌ **Do not need** the Multi-Geo license (no Exchange mailbox, no OneDrive, no SharePoint MySite)

The script handles this automatically:

```
Disabled resource account
  ├── UsageLocation missing/wrong → ✅ Updated to correct region
  └── Multi-Geo license missing  → ✋ Skipped (no seat consumed)
```

Disabled accounts appear in the report with a 🔒 **Disabled account** badge.

---

## 📊 Output Files

Each run generates 3 timestamped files in `$OutputPath`:

| File | Description |
|---|---|
| `MultiGeo_Report_YYYYMMDD_HHmmss.html` | Dark-themed compliance dashboard |
| `MultiGeo_Report_YYYYMMDD_HHmmss.csv` | Full tabular export (Excel / SIEM) |
| `MultiGeo_YYYYMMDD_HHmmss.log` | Timestamped action log |

### DryRun report

- Orange **"Needs action"** cards with **"Location to fix"** and **"License missing"** badges
- **Compliance notes** per card explaining *why* each change is required
- KPI bar: accounts to remediate / location fixes / licenses to assign / disabled accounts
- Full amber **DRY RUN** banner — impossible to confuse with a live report

### Live execution report

- Green **"Success"**, amber **"Partial"**, red **"Error"** cards
- Badges reflect **what actually happened**: `✓ Location updated` / `✗ License failed`
- KPI bar counts actual operations performed, not what was planned
- Auto-opens after run completes

---

## 🔐 Required Permissions

Authentication uses Microsoft Graph delegated permissions:

| Permission | Purpose |
|---|---|
| `User.ReadWrite.All` | Read users, update UsageLocation, assign licenses |
| `Organization.Read.All` | Enumerate subscribed SKUs |

First-time consent requires a **Global Administrator** or **Privileged Role Administrator**.

---

## 🧪 Running in CI / Unattended Mode

The interactive confirmation prompt is skipped automatically when `[Environment]::UserInteractive` is `$false` (Azure Automation, GitHub Actions, Scheduled Tasks).

Authenticate via Service Principal before calling the script:

```powershell
Connect-MgGraph -ClientId $env:CLIENT_ID -TenantId $env:TENANT_ID -CertificateThumbprint $env:CERT_THUMB
.\Set-MultiGeoUsageLocation.ps1 -Domain "contoso.com" -SkuPartNumber "OFFICE365_MULTIGEO" -OutputPath "/tmp/reports"
```

---

## 📄 License

MIT © [Souhaiel Morhag](https://msendpoint.com)

---

## 🤝 Contributing

Issues and PRs are welcome. Please open an issue before submitting large changes.

---

<div align="center">

*Built with ❤️ by [msendpoint.com](https://msendpoint.com) — Microsoft 365 & Endpoint Management*

</div>
