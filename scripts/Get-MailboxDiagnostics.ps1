<#
.SYNOPSIS
    Reports mailbox stores, folder sync state and the offline cache window.
.DESCRIPTION
    Run this first. It answers the three questions that derail an extract:
    which stores are attached (delegate mailboxes show up here), how much of
    each folder has actually synced, and how far back the offline cache goes.
.PARAMETER StoreName
    Limit folder reporting to one store. Defaults to the primary mailbox.
.EXAMPLE
    .\Get-MailboxDiagnostics.ps1
.EXAMPLE
    .\Get-MailboxDiagnostics.ps1 -StoreName 'user@company.com'
#>
[CmdletBinding()]
param(
    [string] $StoreName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\src\OutlookMapiTools\OutlookMapiTools.psd1') -Force

$ns = Connect-OutlookMapi

Write-Host '=== stores attached to this profile ===' -ForegroundColor Cyan
foreach ($store in $ns.Stores) {
    # ExchangeStoreType 0 = primary mailbox, 1 = delegate/shared.
    $kind = switch ($store.ExchangeStoreType) {
        0       { 'primary' }
        1       { 'delegate/shared' }
        default { "type $($store.ExchangeStoreType)" }
    }
    '{0,-45} {1}' -f $store.DisplayName, $kind
}
Write-Host ''
Write-Host 'Delegate stores are other people''s mail. Scripts here read the default store only.' -ForegroundColor Yellow

Write-Host ''
Write-Host '=== mail folders with items ===' -ForegroundColor Cyan
$folderArgs = @{ Namespace = $ns }
if ($StoreName) { $folderArgs.StoreName = $StoreName }

foreach ($folder in (Get-OutlookMailFolder @folderArgs)) {
    try {
        $count = $folder.Items.Count
        if ($count -gt 0) { '{0,-40} {1,8}' -f $folder.Name, $count }
    } catch { }
}

Write-Host ''
Write-Host '=== offline cache window ===' -ForegroundColor Cyan
$cache = Get-OutlookCacheWindow
if ($cache.MeansAllMail) {
    'Effective: ALL mail cached offline'
} else {
    "Effective: $($cache.EffectiveMonths) months cached offline"
}
"User key   : $($cache.UserKeyValue)"
"Policy key : $($cache.PolicyKeyValue)"
''
$cache.Note
