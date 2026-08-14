<#
    OutlookMapiTools - read-only helpers for Outlook desktop via MAPI/COM.

    Every function here reads. Nothing sends, modifies or deletes mail.
    Requires classic Outlook (Office 16) installed, running, and signed in
    to a mail profile.
#>

Set-StrictMode -Version Latest

# Outlook enum values used below. Declared explicitly so callers do not need
# the interop assembly loaded.
$script:OlFolderSentMail = 5
$script:OlFolderInbox    = 6
$script:OlMailItem       = 43
$script:OlMailItemType   = 0          # Folder.DefaultItemType for mail folders
$script:PrSmtpAddress    = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'

# Recipient.Type
$script:OlBcc = 3

function Connect-OutlookMapi {
    <#
    .SYNOPSIS
        Returns the MAPI namespace for the running Outlook instance.
    .DESCRIPTION
        Outlook must already be running and signed in. If the COM server is
        unavailable this throws with a message pointing at the usual cause.
    .EXAMPLE
        $ns = Connect-OutlookMapi
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param()

    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        throw "Could not attach to Outlook COM. Is classic Outlook running and signed in? Original error: $($_.Exception.Message)"
    }
    return $outlook.GetNamespace('MAPI')
}

function Get-OutlookMailFolder {
    <#
    .SYNOPSIS
        Enumerates mail folders in a store, depth-first.
    .PARAMETER Namespace
        MAPI namespace from Connect-OutlookMapi.
    .PARAMETER StoreName
        Display name of the store to walk. Defaults to the primary mailbox.
        Delegate and shared mailboxes are ignored unless named explicitly.
    .PARAMETER MaxDepth
        Guards against pathological folder trees.
    .EXAMPLE
        Get-OutlookMailFolder -Namespace $ns | Select-Object Name, @{n='N';e={$_.Items.Count}}
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory)] $Namespace,
        [string] $StoreName,
        [int]    $MaxDepth = 6
    )

    $root = if ($StoreName) {
        $store = $Namespace.Stores | Where-Object { $_.DisplayName -eq $StoreName }
        if (-not $store) { throw "No store named '$StoreName'. Available: $(($Namespace.Stores | ForEach-Object DisplayName) -join ', ')" }
        $store.GetRootFolder()
    } else {
        $Namespace.DefaultStore.GetRootFolder()
    }

    $found = [System.Collections.Generic.List[object]]::new()

    function Walk {
        param($Folder, [int]$Depth)
        if ($Depth -gt $MaxDepth) { return }
        # Mail folders only - skips calendar, contacts, notes.
        try {
            if ($Folder.DefaultItemType -eq $script:OlMailItemType) { $found.Add($Folder) }
        } catch { }
        try {
            foreach ($child in $Folder.Folders) { Walk -Folder $child -Depth ($Depth + 1) }
        } catch { }
    }

    Walk -Folder $root -Depth 0
    return $found.ToArray()
}

function Resolve-SmtpAddress {
    <#
    .SYNOPSIS
        Resolves an Outlook Recipient to a real SMTP address.
    .DESCRIPTION
        Exchange recipients expose a legacy X.500 DN in .Address rather than an
        address you can use. This tries the Exchange user object first, then the
        PR_SMTP_ADDRESS MAPI property, then falls back to whatever .Address held.
    .EXAMPLE
        Resolve-SmtpAddress -Recipient $mail.Recipients.Item(1)
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)] $Recipient
    )

    try {
        $entry = $Recipient.AddressEntry
        if ($entry) {
            # 0 = olExchangeUser, 1 = olExchangeDistributionList
            if ($entry.AddressEntryUserType -in 0, 1) {
                $exchangeUser = $entry.GetExchangeUser()
                if ($exchangeUser -and $exchangeUser.PrimarySmtpAddress) {
                    return [string]$exchangeUser.PrimarySmtpAddress
                }
            }
            try {
                $smtp = $entry.PropertyAccessor.GetProperty($script:PrSmtpAddress)
                if ($smtp) { return [string]$smtp }
            } catch { }
        }
    } catch { }

    return [string]$Recipient.Address
}

function Test-SmtpAddress {
    <#
    .SYNOPSIS
        Cheap shape check for an address string.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([string] $Address)

    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    return $Address -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

function Request-FolderSync {
    <#
    .SYNOPSIS
        Promotes a folder to the front of Outlook's sync queue.
    .DESCRIPTION
        Outlook syncs the Inbox to completion before other folders, so a large
        Sent Items can sit at zero items for a long time. Making the folder the
        active one - the programmatic equivalent of clicking it - then kicking a
        send/receive pulls it down immediately.
    .PARAMETER Namespace
        MAPI namespace from Connect-OutlookMapi.
    .PARAMETER Folder
        The folder to prioritise.
    .EXAMPLE
        $sent = $ns.GetDefaultFolder(5)
        Request-FolderSync -Namespace $ns -Folder $sent
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)] $Namespace,
        [Parameter(Mandatory)] $Folder
    )

    if (-not $PSCmdlet.ShouldProcess($Folder.Name, 'Promote in Outlook sync queue')) { return }

    $outlook = New-Object -ComObject Outlook.Application
    try {
        $explorer = $outlook.ActiveExplorer()
        if ($null -eq $explorer) {
            $Folder.Display()
            Start-Sleep -Seconds 5
            $explorer = $outlook.ActiveExplorer()
        }
        if ($explorer) { $explorer.CurrentFolder = $Folder }
    } catch {
        Write-Warning "Could not set active folder: $($_.Exception.Message)"
    }

    for ($i = 1; $i -le $Namespace.SyncObjects.Count; $i++) {
        try { $Namespace.SyncObjects.Item($i).Start() } catch { }
    }
}

function Get-OutlookCacheWindow {
    <#
    .SYNOPSIS
        Reports how many months of mail Outlook keeps cached offline.
    .DESCRIPTION
        Matters because the default is 12 months - ask for 18 months of history
        and you silently get 12. Note that the per-user key only applies to
        newly created accounts; an existing profile must be changed through the
        Outlook UI, or by policy.
    .EXAMPLE
        Get-OutlookCacheWindow
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string] $OfficeVersion = '16.0'
    )

    $userKey   = "HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Outlook\Cached Mode"
    $policyKey = "HKCU:\Software\Policies\Microsoft\Office\$OfficeVersion\Outlook\Cached Mode"

    function ReadSetting([string] $Path) {
        if (-not (Test-Path $Path)) { return $null }
        try { return (Get-ItemProperty -Path $Path -ErrorAction Stop).SyncWindowSetting } catch { return $null }
    }

    $userValue   = ReadSetting $userKey
    $policyValue = ReadSetting $policyKey
    $effective   = if ($null -ne $policyValue) { $policyValue } else { $userValue }

    $months = switch ($effective) {
        $null   { 12;  break }   # Outlook default when unset
        0       { 0;   break }   # 0 means "all mail"
        default { [int]$effective }
    }

    [pscustomobject]@{
        UserKeyValue     = $userValue
        PolicyKeyValue   = $policyValue
        EffectiveMonths  = $months
        MeansAllMail     = ($effective -eq 0)
        AppliesToExisting= ($null -ne $policyValue)
        Note             = 'The user key applies to NEW accounts only. Change an existing profile via Account Settings > "Mail to keep offline", or by policy.'
    }
}

Export-ModuleMember -Function `
    Connect-OutlookMapi,
    Get-OutlookMailFolder,
    Resolve-SmtpAddress,
    Test-SmtpAddress,
    Request-FolderSync,
    Get-OutlookCacheWindow
