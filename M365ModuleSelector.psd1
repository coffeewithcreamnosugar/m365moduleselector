@{
    RootModule = 'M365ModuleSelector.psm1'
    ModuleVersion = '1.0.0'
    GUID = '458a394d-0252-4f9d-98e5-faf1ae39e233'
    Author = 'Jason Beckett'
    CompanyName = 'ASAP IT LLC'
    Copyright = '(c) 2026 Jason Beckett'
    Description = 'Helpers to install and connect to Microsoft 365 PowerShell modules.'
    PowerShellVersion = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport = @(
        'Connect-M365Module',
        'Connect-ModuleMicrosoftGraph',
        'Connect-ModuleExchangeOnline',
        'Connect-ModuleSharePointOnline',
        'Connect-ModuleMicrosoftTeams',
        'Connect-ModulePnP',
        'Connect-ModuleEntra',
        'Connect-ModuleAzure',
        'Connect-ModulePowerBI',
        'Connect-ModulePowerPlatform',
        'Connect-ModuleMicrosoftGraphBeta',
        'Connect-ModulePurviewCompliance',
        'Set-M365ModuleUserPrincipalName'
    )
    AliasesToExport = @('connect-m365module')
    CmdletsToExport = @()
    VariablesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('Microsoft365', 'Graph', 'Exchange', 'SharePoint', 'Teams', 'Entra')
            LicenseUri = ''
            ProjectUri = ''
            ReleaseNotes = 'Initial module manifest.'
        }
    }
}
