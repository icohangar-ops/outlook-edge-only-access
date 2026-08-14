<#
    Pester tests for OutlookMapiTools.

    Pure-logic tests run anywhere. Tests that need a live Outlook COM server are
    tagged 'RequiresOutlook' and skip automatically when it is unavailable, so
    the suite stays green in CI.

    Run:  Invoke-Pester -Path .\tests
    Skip Outlook-dependent tests:  Invoke-Pester -Path .\tests -ExcludeTagFilter RequiresOutlook
#>

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\src\OutlookMapiTools\OutlookMapiTools.psd1'
    Import-Module $modulePath -Force

    $script:OutlookAvailable = $false
    try {
        $probe = New-Object -ComObject Outlook.Application
        $null  = $probe.GetNamespace('MAPI')
        $script:OutlookAvailable = $true
    } catch {
        $script:OutlookAvailable = $false
    }
}

Describe 'Module surface' {
    It 'exports the documented functions' {
        $exported = (Get-Module OutlookMapiTools).ExportedFunctions.Keys
        $exported | Should -Contain 'Connect-OutlookMapi'
        $exported | Should -Contain 'Get-OutlookMailFolder'
        $exported | Should -Contain 'Resolve-SmtpAddress'
        $exported | Should -Contain 'Test-SmtpAddress'
        $exported | Should -Contain 'Request-FolderSync'
        $exported | Should -Contain 'Get-OutlookCacheWindow'
    }

    It 'has a manifest that parses' {
        $manifest = Import-PowerShellDataFile (Join-Path $PSScriptRoot '..\src\OutlookMapiTools\OutlookMapiTools.psd1')
        $manifest.ModuleVersion | Should -Not -BeNullOrEmpty
        $manifest.RootModule    | Should -Be 'OutlookMapiTools.psm1'
    }
}

Describe 'Test-SmtpAddress' {
    It 'accepts a normal address' {
        Test-SmtpAddress -Address 'someone@example.com' | Should -BeTrue
    }

    It 'accepts subdomains and plus addressing' {
        Test-SmtpAddress -Address 'first.last+tag@mail.example.co.uk' | Should -BeTrue
    }

    It 'rejects an Exchange legacy DN' {
        # This is what .Address returns for internal recipients.
        $dn = '/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=abc123'
        Test-SmtpAddress -Address $dn | Should -BeFalse
    }

    It 'rejects empty, whitespace and malformed input' {
        Test-SmtpAddress -Address ''            | Should -BeFalse
        Test-SmtpAddress -Address '   '         | Should -BeFalse
        Test-SmtpAddress -Address 'not-an-addr' | Should -BeFalse
        Test-SmtpAddress -Address 'a@b'         | Should -BeFalse
        Test-SmtpAddress -Address 'a b@c.com'   | Should -BeFalse
    }
}

Describe 'Get-OutlookCacheWindow' {
    It 'reports a month count and the caveat' {
        $result = Get-OutlookCacheWindow
        $result.EffectiveMonths | Should -BeOfType [int]
        $result.Note            | Should -Match 'NEW accounts'
    }

    It 'defaults to 12 months for a version with no keys present' {
        # An implausible Office version guarantees both keys are absent.
        $result = Get-OutlookCacheWindow -OfficeVersion '99.9'
        $result.EffectiveMonths | Should -Be 12
        $result.UserKeyValue    | Should -BeNullOrEmpty
    }
}

Describe 'Live mailbox access' -Tag 'RequiresOutlook' {
    BeforeEach {
        if (-not $script:OutlookAvailable) { Set-ItResult -Skipped -Because 'Outlook COM is not available' }
    }

    It 'connects to MAPI' {
        $ns = Connect-OutlookMapi
        $ns | Should -Not -BeNullOrEmpty
    }

    It 'enumerates mail folders in the default store' {
        $ns = Connect-OutlookMapi
        $folders = Get-OutlookMailFolder -Namespace $ns
        $folders.Count | Should -BeGreaterThan 0
    }

    It 'throws a helpful error for an unknown store' {
        $ns = Connect-OutlookMapi
        { Get-OutlookMailFolder -Namespace $ns -StoreName 'no-such-store-xyz' } |
            Should -Throw -ExpectedMessage '*No store named*'
    }
}
