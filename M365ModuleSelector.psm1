$script:Banner = @"
  ____ ___  _   _ _   _ _____ ____ _____     __  __ _____  __  ____  __  __  ___  ____  _   _ _     _____
 / ___/ _ \| \ | | \ | | ____/ ___|_   _|   |  \/  |___ / / /_| ___||  \/  |/ _ \|  _ \| | | | |   | ____|
| |  | | | |  \| |  \| |  _|| |     | |_____| |\/| | |_ \| '_ \___ \| |\/| | | | | | | | | | | |   |  _|
| |__| |_| | |\  | |\  | |__| |___  | |_____| |  | |___) | (_) |__) | |  | | |_| | |_| | |_| | |___| |___
 \____\___/|_| \_|_| \_|_____\____| |_|     |_|  |_|____/ \___/____/|_|  |_|\___/|____/ \___/|_____|_____|
"@

$script:ModuleRequirements = @{
    "Microsoft.Graph.Authentication"         = [version]"7.0"
    "ExchangeOnlineManagement"               = [version]"5.1"
    "Microsoft.Online.SharePoint.PowerShell" = [version]"5.1"
    "MicrosoftTeams"                         = [version]"5.1"
    "PnP.PowerShell"                         = [version]"7.2"
    "Microsoft.Entra"                        = [version]"7.0"
    "Az.Accounts"                            = [version]"7.2"
    "MicrosoftPowerBIMgmt"                   = [version]"7.2"
    "Microsoft.PowerApps.Administration.PowerShell" = [version]"5.1"
    "Microsoft.Graph.Beta"                   = [version]"7.0"
}

$script:CommonGraphScopes = [ordered]@{
    "1"  = @{ Name = "User.Read.All"; Description = "Read all users" }
    "2"  = @{ Name = "User.ReadWrite.All"; Description = "Read and write all users" }
    "3"  = @{ Name = "Group.Read.All"; Description = "Read all groups" }
    "4"  = @{ Name = "Group.ReadWrite.All"; Description = "Read and write all groups" }
    "5"  = @{ Name = "Directory.Read.All"; Description = "Read directory data" }
    "6"  = @{ Name = "Directory.ReadWrite.All"; Description = "Read and write directory data" }
    "7"  = @{ Name = "Application.Read.All"; Description = "Read applications" }
    "8"  = @{ Name = "Application.ReadWrite.All"; Description = "Read and write applications" }
    "9"  = @{ Name = "Device.Read.All"; Description = "Read devices" }
    "10" = @{ Name = "Device.ReadWrite.All"; Description = "Read and write devices" }
    "11" = @{ Name = "Organization.Read.All"; Description = "Read organization info" }
    "12" = @{ Name = "RoleManagement.Read.Directory"; Description = "Read directory roles" }
    "13" = @{ Name = "AuditLog.Read.All"; Description = "Read audit logs" }
    "14" = @{ Name = "Reports.Read.All"; Description = "Read usage reports" }
    "15" = @{ Name = "Mail.Read"; Description = "Read mail" }
    "16" = @{ Name = "Mail.ReadWrite"; Description = "Read and write mail" }
    "17" = @{ Name = "Calendars.Read"; Description = "Read calendars" }
    "18" = @{ Name = "Sites.Read.All"; Description = "Read SharePoint sites" }
    "19" = @{ Name = "Sites.ReadWrite.All"; Description = "Read and write SharePoint sites" }
}

$script:DefaultUserPrincipalName = $null
$script:DefaultUserPrincipalNamePrompted = $false
$script:BackSelection = "__M365MODULESELECTOR_BACK__"
$script:AvailableModuleVersionCache = @{}

function Test-ModuleInstalled {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    [bool](Get-Module -ListAvailable -Name $ModuleName)
}

function Get-InstalledModuleVersion {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    $module = Get-Module -ListAvailable -Name $ModuleName |
        Sort-Object -Property Version -Descending |
        Select-Object -First 1

    if ($module) {
        return [version]$module.Version
    }

    return [version]"0.0.0.0"
}

function Get-AvailableModuleVersion {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    if ($script:AvailableModuleVersionCache.ContainsKey($ModuleName)) {
        Write-Host "Using cached PowerShell Gallery version for $ModuleName."
        return $script:AvailableModuleVersionCache[$ModuleName]
    }

    try {
        $module = Find-Module -Name $ModuleName -ErrorAction Stop
        $version = [version]$module.Version
        $script:AvailableModuleVersionCache[$ModuleName] = $version
        return $version
    }
    catch {
        Write-Warning "Could not check the PowerShell Gallery for '$ModuleName': $($_.Exception.Message)"
        $script:AvailableModuleVersionCache[$ModuleName] = $null
        return $null
    }
}

function Test-PowerShellVersionForModule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    if (-not $script:ModuleRequirements.ContainsKey($ModuleName)) {
        return $true
    }

    $minimumVersion = $script:ModuleRequirements[$ModuleName]
    if ($PSVersionTable.PSVersion -ge $minimumVersion) {
        return $true
    }

    Write-Warning "'$ModuleName' requires PowerShell $minimumVersion or newer. Current version: $($PSVersionTable.PSVersion)."
    return $false
}

function Test-M365ModuleCloudProviderError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    return ($Message -match '0x8007016A' -or $Message -match 'cloud file provider is not running')
}

function Invoke-M365TimedStep {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Activity,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        & $ScriptBlock
    }
    finally {
        $stopwatch.Stop()
        Write-Host ("[timing] {0}: {1:N2}s" -f $Activity, $stopwatch.Elapsed.TotalSeconds) -ForegroundColor DarkGray
    }
}

function Install-OrUpdateModule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,

        [ValidateSet("CurrentUser", "AllUsers")]
        [string]$Scope = "CurrentUser"
    )

    if (-not (Invoke-M365TimedStep -Activity "Check PowerShell version for $ModuleName" -ScriptBlock {
        Test-PowerShellVersionForModule -ModuleName $ModuleName
    })) {
        return $false
    }

    $installedVersion = Invoke-M365TimedStep -Activity "Find installed $ModuleName version" -ScriptBlock {
        Get-InstalledModuleVersion -ModuleName $ModuleName
    }
    $availableVersion = Invoke-M365TimedStep -Activity "Check PowerShell Gallery for $ModuleName" -ScriptBlock {
        Get-AvailableModuleVersion -ModuleName $ModuleName
    }

    try {
        if ($installedVersion -eq [version]"0.0.0.0") {
            Write-Host "$ModuleName is not installed. Installing to $Scope..."
            Invoke-M365TimedStep -Activity "Install $ModuleName" -ScriptBlock {
                Install-Module -Name $ModuleName -Scope $Scope -Force -AllowClobber -ErrorAction Stop
            }
        }
        elseif ($availableVersion -and $availableVersion -gt $installedVersion) {
            Write-Host "A newer version of $ModuleName is available: $installedVersion -> $availableVersion."
            $update = Read-Host "Update $ModuleName now? (Y/N)"

            if ($update -match '^(Y|y)') {
                try {
                    Invoke-M365TimedStep -Activity "Update $ModuleName" -ScriptBlock {
                        Update-Module -Name $ModuleName -Force -ErrorAction Stop
                    }
                }
                catch {
                    if ($_.Exception.Message -match 'was not installed by using Install-Module') {
                        Write-Warning "$ModuleName cannot be updated with Update-Module because it was not installed from PowerShell Gallery. Installing the latest version to $Scope instead..."
                        Invoke-M365TimedStep -Activity "Install latest $ModuleName" -ScriptBlock {
                            Install-Module -Name $ModuleName -Scope $Scope -Force -AllowClobber -ErrorAction Stop
                        }
                    }
                    else {
                        throw
                    }
                }
            }
        }
        else {
            Write-Host "$ModuleName is installed (version $installedVersion)."
        }

        $loadedModule = Get-Module -Name $ModuleName
        if ($loadedModule) {
            Write-Host "$ModuleName is already loaded (version $($loadedModule.Version))."
        }
        else {
            Invoke-M365TimedStep -Activity "Import $ModuleName" -ScriptBlock {
                Import-Module -Name $ModuleName -ErrorAction Stop
            }
        }
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message

        if (Test-M365ModuleCloudProviderError -Message $errorMessage) {
            Write-Error "Could not prepare module '$ModuleName': PowerShell found the module in a cloud-backed location, but the files are not available locally. Start OneDrive, mark the module folder as 'Always keep on this device', or reinstall the module into a local PowerShell module path. Original error: $errorMessage"
            return $false
        }

        Write-Error "Could not prepare module '$ModuleName': $errorMessage"
        return $false
    }
}

function Connect-ModuleService {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ConnectCommand,

        [switch]$SkipModulePreparation
    )

    if (-not $SkipModulePreparation -and -not (Install-OrUpdateModule -ModuleName $ModuleName)) {
        return
    }

    try {
        Invoke-M365TimedStep -Activity "Run $ModuleName connection command" -ScriptBlock {
            & $ConnectCommand
        }
        Write-Host "Connected using $ModuleName."
    }
    catch {
        Write-Error "Failed to connect using ${ModuleName}: $($_.Exception.Message)"
    }
}

function Test-M365ModuleBackSelection {
    [CmdletBinding()]
    param (
        [AllowNull()]
        [string]$Selection
    )

    if ([string]::IsNullOrWhiteSpace($Selection)) {
        return $false
    }

    $normalizedSelection = $Selection.Trim().ToLowerInvariant()
    return ($normalizedSelection -eq "b" -or $normalizedSelection -eq "back")
}

function Set-M365ModuleUserPrincipalName {
    [CmdletBinding()]
    param (
        [AllowEmptyString()]
        [string]$UserPrincipalName
    )

    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        $script:DefaultUserPrincipalName = $null
        Write-Host "Default username cleared."
        return
    }

    $script:DefaultUserPrincipalName = $UserPrincipalName.Trim()
    Write-Host "Default username set to $script:DefaultUserPrincipalName."
}

function Read-M365ModuleUserPrincipalName {
    [CmdletBinding()]
    param()

    if ($script:DefaultUserPrincipalName -or $script:DefaultUserPrincipalNamePrompted) {
        return $script:DefaultUserPrincipalName
    }

    $script:DefaultUserPrincipalNamePrompted = $true
    $userPrincipalName = Read-Host "Enter username/UPN to prefill supported sign-in prompts (leave blank to skip)"

    if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
        return $null
    }

    $script:DefaultUserPrincipalName = $userPrincipalName.Trim()
    return $script:DefaultUserPrincipalName
}

function Read-ModuleAuthenticationMode {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServiceName
    )

    Write-Host "Select $ServiceName authentication mode:"
    Write-Host "1. Browser"
    Write-Host "2. Device code"
    Write-Host "b. Back"

    $selection = Read-Host "Enter your choice (1-2 or b)"
    switch ($selection.ToLowerInvariant()) {
        "1" { return "Browser" }
        "browser" { return "Browser" }
        "2" { return "DeviceCode" }
        "devicecode" { return "DeviceCode" }
        "device code" { return "DeviceCode" }
        "b" { return $script:BackSelection }
        "back" { return $script:BackSelection }
        default {
            Write-Host "Invalid selection. Defaulting to browser authentication."
            return "Browser"
        }
    }
}

function Read-GraphScopes {
    [CmdletBinding()]
    param()

    Write-Host "Select Microsoft Graph scopes. Use comma-separated numbers for more than one scope."
    foreach ($scope in $script:CommonGraphScopes.GetEnumerator()) {
        Write-Host "$($scope.Key). $($scope.Value.Name) " -NoNewline
        Write-Host "- $($scope.Value.Description)" -ForegroundColor DarkGray
    }
    Write-Host "c. Custom scopes"
    Write-Host "b. Back"

    $selection = Read-Host "Enter scope choice(s), c for custom, b to go back, or press Enter for User.Read.All"
    if (-not $selection) {
        return @("User.Read.All")
    }

    if (Test-M365ModuleBackSelection -Selection $selection) {
        return $script:BackSelection
    }

    if ($selection.ToLowerInvariant() -eq "c") {
        $customScopes = Read-Host "Enter comma-separated Graph scopes"
        if (Test-M365ModuleBackSelection -Selection $customScopes) {
            return $script:BackSelection
        }

        return @($customScopes -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $selectedScopes = foreach ($choice in ($selection -split ',')) {
        $key = $choice.Trim()
        if ($script:CommonGraphScopes.Contains($key)) {
            $script:CommonGraphScopes[$key].Name
        }
        else {
            Write-Warning "Skipping invalid scope choice '$key'."
        }
    }

    if ($selectedScopes) {
        return @($selectedScopes)
    }

    Write-Host "No valid scope choices selected. Defaulting to User.Read.All."
    return @("User.Read.All")
}

function Connect-ModuleMicrosoftGraph {
    [CmdletBinding()]
    param (
        [string[]]$Scopes,
        [string]$TenantId,

        [ValidateSet("Browser", "DeviceCode")]
        [string]$AuthenticationMode
    )

    if (-not (Install-OrUpdateModule -ModuleName "Microsoft.Graph.Authentication")) {
        return
    }

    if (-not $Scopes) {
        $Scopes = Read-GraphScopes
        if ($Scopes -eq $script:BackSelection) {
            return
        }
    }

    if (-not $TenantId -and ($Scopes | Where-Object { $_ -like "Policy.*" })) {
        $TenantId = Read-Host -Prompt "Enter tenant ID or domain for policy scopes (leave blank to use default, or b to go back)"
        if (Test-M365ModuleBackSelection -Selection $TenantId) {
            return
        }
    }

    if (-not $AuthenticationMode) {
        $selectedAuthenticationMode = Read-ModuleAuthenticationMode -ServiceName "Microsoft Graph"
        if ($selectedAuthenticationMode -eq $script:BackSelection) {
            return
        }
        $AuthenticationMode = $selectedAuthenticationMode
    }

    Connect-ModuleService -ModuleName "Microsoft.Graph.Authentication" -SkipModulePreparation -ConnectCommand {
        $connectParams = @{
            Scopes = $Scopes
            NoWelcome = $true
        }

        if ($TenantId) {
            $connectParams.TenantId = $TenantId
        }

        if ($AuthenticationMode -eq "DeviceCode") {
            $connectParams.UseDeviceCode = $true
        }

        Connect-MgGraph @connectParams
    }
}

function Connect-ModuleMicrosoftGraphBeta {
    [CmdletBinding()]
    param (
        [string[]]$Scopes,
        [string]$TenantId,

        [ValidateSet("Browser", "DeviceCode")]
        [string]$AuthenticationMode
    )

    if (-not (Install-OrUpdateModule -ModuleName "Microsoft.Graph.Authentication")) {
        return
    }

    if (-not $Scopes) {
        $Scopes = Read-GraphScopes
        if ($Scopes -eq $script:BackSelection) {
            return
        }
    }

    if (-not $TenantId -and ($Scopes | Where-Object { $_ -like "Policy.*" })) {
        $TenantId = Read-Host -Prompt "Enter tenant ID or domain for policy scopes (leave blank to use default, or b to go back)"
        if (Test-M365ModuleBackSelection -Selection $TenantId) {
            return
        }
    }

    if (-not $AuthenticationMode) {
        $selectedAuthenticationMode = Read-ModuleAuthenticationMode -ServiceName "Microsoft Graph Beta"
        if ($selectedAuthenticationMode -eq $script:BackSelection) {
            return
        }
        $AuthenticationMode = $selectedAuthenticationMode
    }

    Connect-ModuleService -ModuleName "Microsoft.Graph.Authentication" -SkipModulePreparation -ConnectCommand {
        $connectParams = @{
            Scopes = $Scopes
            NoWelcome = $true
        }

        if ($TenantId) {
            $connectParams.TenantId = $TenantId
        }

        if ($AuthenticationMode -eq "DeviceCode") {
            $connectParams.UseDeviceCode = $true
        }

        Connect-MgGraph @connectParams
    }
}

function Connect-ModuleExchangeOnline {
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,

        [ValidateSet("Browser", "DeviceCode")]
        [string]$AuthenticationMode
    )

    if (-not $UserPrincipalName) {
        $UserPrincipalName = $script:DefaultUserPrincipalName
    }

    if (-not (Install-OrUpdateModule -ModuleName "ExchangeOnlineManagement")) {
        return
    }

    if (-not $AuthenticationMode) {
        $selectedAuthenticationMode = Read-ModuleAuthenticationMode -ServiceName "Exchange Online"
        if ($selectedAuthenticationMode -eq $script:BackSelection) {
            return
        }
        $AuthenticationMode = $selectedAuthenticationMode
    }

    Connect-ModuleService -ModuleName "ExchangeOnlineManagement" -SkipModulePreparation -ConnectCommand {
        $connectParams = @{
            ShowBanner = $false
        }

        if ($UserPrincipalName) {
            $connectParams.UserPrincipalName = $UserPrincipalName
        }

        if ($AuthenticationMode -eq "DeviceCode") {
            $connectParams.Device = $true
        }

        Connect-ExchangeOnline @connectParams
    }
}

function Connect-ModuleSharePointOnline {
    [CmdletBinding()]
    param (
        [string]$AdminUrl
    )

    if (-not (Install-OrUpdateModule -ModuleName "Microsoft.Online.SharePoint.PowerShell")) {
        return
    }

    if (-not $AdminUrl) {
        $AdminUrl = Read-Host -Prompt "Enter SharePoint admin URL (https://tenant-admin.sharepoint.com) or b to go back"
        if (Test-M365ModuleBackSelection -Selection $AdminUrl) {
            return
        }
    }

    Connect-ModuleService -ModuleName "Microsoft.Online.SharePoint.PowerShell" -SkipModulePreparation -ConnectCommand {
        Connect-SPOService -Url $AdminUrl
    }
}

function Connect-ModuleMicrosoftTeams {
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,

        [ValidateSet("Browser", "DeviceCode")]
        [string]$AuthenticationMode
    )

    if (-not $UserPrincipalName) {
        $UserPrincipalName = $script:DefaultUserPrincipalName
    }

    if (-not (Install-OrUpdateModule -ModuleName "MicrosoftTeams")) {
        return
    }

    if (-not $AuthenticationMode) {
        $selectedAuthenticationMode = Read-ModuleAuthenticationMode -ServiceName "Microsoft Teams"
        if ($selectedAuthenticationMode -eq $script:BackSelection) {
            return
        }
        $AuthenticationMode = $selectedAuthenticationMode
    }

    Connect-ModuleService -ModuleName "MicrosoftTeams" -SkipModulePreparation -ConnectCommand {
        $connectParams = @{}

        if ($UserPrincipalName) {
            $connectParams.AccountId = $UserPrincipalName
        }

        if ($AuthenticationMode -eq "DeviceCode") {
            $connectParams.UseDeviceAuthentication = $true
        }

        Connect-MicrosoftTeams @connectParams
    }
}

function Connect-ModulePnP {
    [CmdletBinding()]
    param (
        [string]$SiteUrl,
        [string]$ClientId,
        [string]$Tenant,

        [ValidateSet("Browser", "DeviceCode")]
        [string]$AuthenticationMode
    )

    if (-not (Install-OrUpdateModule -ModuleName "PnP.PowerShell")) {
        return
    }

    if (-not $SiteUrl) {
        $SiteUrl = Read-Host -Prompt "Enter SharePoint site URL (https://tenant.sharepoint.com/sites/site) or b to go back"
        if (Test-M365ModuleBackSelection -Selection $SiteUrl) {
            return
        }
    }

    if (-not $AuthenticationMode) {
        $selectedAuthenticationMode = Read-ModuleAuthenticationMode -ServiceName "PnP PowerShell"
        if ($selectedAuthenticationMode -eq $script:BackSelection) {
            return
        }
        $AuthenticationMode = $selectedAuthenticationMode
    }

    if (-not $ClientId) {
        $ClientId = $env:ENTRAID_CLIENT_ID
        if (-not $ClientId) {
            $ClientId = $env:ENTRAID_APP_ID
        }
    }

    if (-not $ClientId) {
        Write-Warning "PnP.PowerShell now requires an Entra ID App client id. Set ENTRAID_CLIENT_ID or ENTRAID_APP_ID, or pass -ClientId to Connect-ModulePnP."
        $ClientId = Read-Host -Prompt 'Enter Entra ID App ClientId (leave blank to try without it, or b to go back)'
        if (Test-M365ModuleBackSelection -Selection $ClientId) {
            return
        }
    }

    $connectParams = @{ Url = $SiteUrl }
    if ($AuthenticationMode -eq "DeviceCode") {
        $connectParams.DeviceLogin = $true
    }
    else {
        $connectParams.Interactive = $true
    }

    if ($ClientId) { $connectParams.ClientId = $ClientId }
    if ($Tenant) { $connectParams.Tenant = $Tenant }

    Connect-ModuleService -ModuleName "PnP.PowerShell" -SkipModulePreparation -ConnectCommand {
        Connect-PnPOnline @connectParams
    }
}

function Connect-ModuleEntra {
    [CmdletBinding()]
    param (
        [string[]]$Scopes = @("User.Read.All"),

        [ValidateSet("Browser", "DeviceCode")]
        [string]$AuthenticationMode
    )

    if (-not (Install-OrUpdateModule -ModuleName "Microsoft.Entra")) {
        return
    }

    if (-not $AuthenticationMode) {
        $selectedAuthenticationMode = Read-ModuleAuthenticationMode -ServiceName "Microsoft Entra"
        if ($selectedAuthenticationMode -eq $script:BackSelection) {
            return
        }
        $AuthenticationMode = $selectedAuthenticationMode
    }

    Connect-ModuleService -ModuleName "Microsoft.Entra" -SkipModulePreparation -ConnectCommand {
        $connectParams = @{
            Scopes = $Scopes
            NoWelcome = $true
        }

        if ($AuthenticationMode -eq "DeviceCode") {
            $connectParams.UseDeviceCode = $true
        }

        Connect-Entra @connectParams
    }
}

function Connect-ModuleAzure {
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,
        [string]$Tenant,
        [string]$Subscription,

        [ValidateSet("Browser", "DeviceCode")]
        [string]$AuthenticationMode
    )

    if (-not $UserPrincipalName) {
        $UserPrincipalName = $script:DefaultUserPrincipalName
    }

    if (-not (Install-OrUpdateModule -ModuleName "Az.Accounts")) {
        return
    }

    if (-not $AuthenticationMode) {
        $selectedAuthenticationMode = Read-ModuleAuthenticationMode -ServiceName "Azure"
        if ($selectedAuthenticationMode -eq $script:BackSelection) {
            return
        }
        $AuthenticationMode = $selectedAuthenticationMode
    }

    Connect-ModuleService -ModuleName "Az.Accounts" -SkipModulePreparation -ConnectCommand {
        $connectParams = @{}
        if ($UserPrincipalName) { $connectParams.AccountId = $UserPrincipalName }
        if ($Tenant) { $connectParams.Tenant = $Tenant }
        if ($Subscription) { $connectParams.Subscription = $Subscription }
        if ($AuthenticationMode -eq "DeviceCode") {
            $connectParams.UseDeviceAuthentication = $true
        }

        Connect-AzAccount @connectParams
    }
}

function Connect-ModulePowerBI {
    [CmdletBinding()]
    param(
        [ValidateSet("Public", "Germany", "China", "USGov", "USGovHigh", "USGovMil")]
        [string]$Environment = "Public"
    )

    if (-not (Install-OrUpdateModule -ModuleName "MicrosoftPowerBIMgmt")) {
        return
    }

    Connect-ModuleService -ModuleName "MicrosoftPowerBIMgmt" -SkipModulePreparation -ConnectCommand {
        Connect-PowerBIServiceAccount -Environment $Environment
    }
}

function Connect-ModulePowerPlatform {
    [CmdletBinding()]
    param(
        [ValidateSet("prod", "preview", "tip1", "tip2", "usgov", "usgovhigh", "dod", "china")]
        [string]$Endpoint = "prod"
    )

    if (-not (Install-OrUpdateModule -ModuleName "Microsoft.PowerApps.Administration.PowerShell")) {
        return
    }

    Connect-ModuleService -ModuleName "Microsoft.PowerApps.Administration.PowerShell" -SkipModulePreparation -ConnectCommand {
        Add-PowerAppsAccount -Endpoint $Endpoint
    }
}

function Connect-ModulePurviewCompliance {
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,

        [ValidateSet("Browser", "DeviceCode")]
        [string]$AuthenticationMode
    )

    if (-not $UserPrincipalName) {
        $UserPrincipalName = $script:DefaultUserPrincipalName
    }

    if (-not (Install-OrUpdateModule -ModuleName "ExchangeOnlineManagement")) {
        return
    }

    if (-not $AuthenticationMode) {
        $selectedAuthenticationMode = Read-ModuleAuthenticationMode -ServiceName "Purview Compliance"
        if ($selectedAuthenticationMode -eq $script:BackSelection) {
            return
        }
        $AuthenticationMode = $selectedAuthenticationMode
    }

    Connect-ModuleService -ModuleName "ExchangeOnlineManagement" -SkipModulePreparation -ConnectCommand {
        $connectParams = @{}

        if ($UserPrincipalName) {
            $connectParams.UserPrincipalName = $UserPrincipalName
        }

        if ($AuthenticationMode -eq "DeviceCode") {
            $connectParams.Device = $true
        }

        Connect-IPPSSession @connectParams
    }
}

function Connect-M365Module {
    [CmdletBinding()]
    param (
        [string]$Title = 'Select a service to connect to',

        [string]$UserPrincipalName
    )

    if ($UserPrincipalName) {
        Set-M365ModuleUserPrincipalName -UserPrincipalName $UserPrincipalName
    }
    else {
        $UserPrincipalName = Read-M365ModuleUserPrincipalName
    }

    while ($true) {
        Write-Host $script:Banner -ForegroundColor Cyan
        Write-Host $Title

        $options = [ordered]@{
            "1" = @{ Name = "Microsoft Graph"; Module = "Microsoft.Graph.Authentication" }
            "2" = @{ Name = "Exchange Online"; Module = "ExchangeOnlineManagement" }
            "3" = @{ Name = "SharePoint Online"; Module = "Microsoft.Online.SharePoint.PowerShell" }
            "4" = @{ Name = "Microsoft Teams"; Module = "MicrosoftTeams" }
            "5" = @{ Name = "PnP PowerShell"; Module = "PnP.PowerShell" }
            "6" = @{ Name = "Microsoft Entra"; Module = "Microsoft.Entra" }
            "7" = @{ Name = "Azure"; Module = "Az.Accounts" }
            "8" = @{ Name = "Power BI"; Module = "MicrosoftPowerBIMgmt" }
            "9" = @{ Name = "Power Platform"; Module = "Microsoft.PowerApps.Administration.PowerShell" }
            "10" = @{ Name = "Microsoft Graph Beta"; Module = "Microsoft.Graph.Beta" }
            "11" = @{ Name = "Purview Compliance"; Module = "ExchangeOnlineManagement" }
            "q" = "Return to shell"
        }

        foreach ($option in $options.GetEnumerator()) {
            if ($option.Value -is [hashtable]) {
                Write-Host "$($option.Key). $($option.Value.Name) " -NoNewline
                Write-Host "($($option.Value.Module))" -ForegroundColor DarkYellow
            }
            else {
                Write-Host "$($option.Key). $($option.Value)"
            }
        }

        $selection = Read-Host "Enter your choice (1-11 or q to return to shell)"

        switch ($selection.ToLowerInvariant()) {
            "1" { Connect-ModuleMicrosoftGraph }
            "2" { Connect-ModuleExchangeOnline -UserPrincipalName $UserPrincipalName }
            "3" { Connect-ModuleSharePointOnline }
            "4" { Connect-ModuleMicrosoftTeams -UserPrincipalName $UserPrincipalName }
            "5" { Connect-ModulePnP }
            "6" { Connect-ModuleEntra }
            "7" { Connect-ModuleAzure -UserPrincipalName $UserPrincipalName }
            "8" { Connect-ModulePowerBI }
            "9" { Connect-ModulePowerPlatform }
            "10" { Connect-ModuleMicrosoftGraphBeta }
            "11" { Connect-ModulePurviewCompliance -UserPrincipalName $UserPrincipalName }
            "q" {
                Write-Host "Returning to shell..."
                return
            }
            default {
                Write-Host "Invalid selection. Please try again."
            }
        }
    }
}

Export-ModuleMember -Function Connect-M365Module, Connect-ModuleMicrosoftGraph, Connect-ModuleExchangeOnline, `
    Connect-ModuleSharePointOnline, Connect-ModuleMicrosoftTeams, Connect-ModulePnP, Connect-ModuleEntra, `
    Connect-ModuleAzure, Connect-ModulePowerBI, Connect-ModulePowerPlatform, `
    Connect-ModuleMicrosoftGraphBeta, Connect-ModulePurviewCompliance, Set-M365ModuleUserPrincipalName
