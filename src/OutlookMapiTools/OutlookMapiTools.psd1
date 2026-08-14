@{
    RootModule        = 'OutlookMapiTools.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '9f2c1d64-5b3a-4c88-a1e7-6d0f2b7c4e19'
    Author            = 'Shyam Desigan'
    Copyright         = '(c) Shyam Desigan. All rights reserved.'
    Description       = 'Read-only helpers for extracting records from Outlook desktop over MAPI/COM, for environments where browser-based Graph access is blocked by conditional access.'

    PowerShellVersion = '5.1'
    # COM interop against a locally installed Outlook - Windows only.
    CompatiblePSEditions = @('Desktop', 'Core')

    FunctionsToExport = @(
        'Connect-OutlookMapi',
        'Get-OutlookMailFolder',
        'Resolve-SmtpAddress',
        'Test-SmtpAddress',
        'Request-FolderSync',
        'Get-OutlookCacheWindow'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags       = @('Outlook', 'MAPI', 'COM', 'Exchange', 'ConditionalAccess', 'Windows')
            LicenseUri = 'https://opensource.org/licenses/MIT'
        }
    }
}
