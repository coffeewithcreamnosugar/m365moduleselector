# M365ModuleSelector

M365ModuleSelector is a PowerShell helper module for installing, updating, importing, and connecting to common Microsoft 365 administration modules from one text menu.

## Supported Services

- Microsoft Graph
- Exchange Online
- SharePoint Online
- Microsoft Teams
- PnP PowerShell
- Microsoft Entra
- Azure
- Power BI
- Power Platform
- Microsoft Graph Beta
- Purview Compliance

Microsoft Graph sign-in uses `Microsoft.Graph.Authentication` so the selector does not need to import the full `Microsoft.Graph` rollup module just to connect.

## Requirements

- PowerShell 7.2 or newer
- Internet access to install or update modules from the PowerShell Gallery
- The Microsoft 365, Azure, or Power Platform roles required by the service you connect to

## Usage

Clone or download this repository, then import the module manifest:

```powershell
Import-Module .\M365ModuleSelector.psd1
Connect-M365Module
```

After you select a service, the menu checks whether the module is installed, offers to update it when a newer version is available, imports it, and then starts the service connection prompts.

Submenus and prompts support `b` or `back` to return to the main menu.

To prefill supported authentication prompts with a username, pass a UPN when opening the menu:

```powershell
Connect-M365Module -UserPrincipalName admin@contoso.com
```

You can also set or clear the default UPN for the current PowerShell session:

```powershell
Set-M365ModuleUserPrincipalName -UserPrincipalName admin@contoso.com
Set-M365ModuleUserPrincipalName
```

## Direct Commands

You can also call a specific connector directly:

```powershell
Connect-ModuleMicrosoftGraph
Connect-ModuleExchangeOnline
Connect-ModuleSharePointOnline
Connect-ModuleMicrosoftTeams
Connect-ModulePnP
Connect-ModuleEntra
Connect-ModuleAzure
Connect-ModulePowerBI
Connect-ModulePowerPlatform
Connect-ModuleMicrosoftGraphBeta
Connect-ModulePurviewCompliance
Set-M365ModuleUserPrincipalName
```

## Authentication Modes

Several services support a browser or device-code authentication choice:

```powershell
Connect-ModuleMicrosoftGraph -AuthenticationMode Browser
Connect-ModuleMicrosoftGraph -AuthenticationMode DeviceCode
```

The same `-AuthenticationMode` pattern is available for Graph Beta, Exchange Online, Microsoft Teams, PnP PowerShell, Microsoft Entra, Azure, and Purview Compliance.

Exchange Online, Microsoft Teams, Azure, and Purview Compliance can use the cached or supplied `-UserPrincipalName` value to reduce repeated username entry during authentication.

## Microsoft Graph Scopes

When connecting to Microsoft Graph or Microsoft Graph Beta without specifying scopes, the module prompts you to select common delegated scopes such as users, groups, directory, applications, devices, audit logs, reports, mail, calendars, and sites.

Select multiple scopes with comma-separated menu numbers:

```text
1,3,5
```

Or enter custom scopes when prompted.

You can also pass scopes directly:

```powershell
Connect-ModuleMicrosoftGraph -Scopes User.Read.All,Group.Read.All
```

Some tenant-specific admin scopes, such as `Policy.Read.All`, may require a tenant hint:

```powershell
Connect-ModuleMicrosoftGraph -TenantId contoso.onmicrosoft.com -Scopes Policy.Read.All,Policy.ReadWrite.ApplicationConfiguration
```

## Notes

Microsoft Graph Beta uses preview APIs that can change. Prefer Microsoft Graph v1.0 for production scripts unless you specifically need a beta-only capability.

Some Power Platform PowerShell behavior may vary by module version and PowerShell runtime.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE).
