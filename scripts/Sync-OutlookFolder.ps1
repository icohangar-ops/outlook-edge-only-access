<#
.SYNOPSIS
    Forces a named folder to sync ahead of the Inbox, then waits for it to fill.
.DESCRIPTION
    Outlook syncs the Inbox to completion before other folders. On a large
    mailbox that leaves Sent Items sitting at zero for a long time. Making the
    folder active promotes it in the queue; this then polls until the count
    stops changing.
.PARAMETER FolderId
    Default folder to sync. 5 = Sent Items, 6 = Inbox.
.PARAMETER TimeoutMinutes
    Give up after this long.
.EXAMPLE
    .\Sync-OutlookFolder.ps1 -FolderId 5
#>
[CmdletBinding()]
param(
    [int] $FolderId = 5,
    [int] $TimeoutMinutes = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\src\OutlookMapiTools\OutlookMapiTools.psd1') -Force

$ns     = Connect-OutlookMapi
$folder = $ns.GetDefaultFolder($FolderId)

Write-Host "Folder: $($folder.FolderPath)"
Write-Host "Before: $($folder.Items.Count) items"

Request-FolderSync -Namespace $ns -Folder $folder

$deadline    = (Get-Date).AddMinutes($TimeoutMinutes)
$previous    = -1
$stableTicks = 0

while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 20
    $count = $folder.Items.Count
    '{0:HH:mm:ss}  {1,8} items' -f (Get-Date), $count

    if ($count -eq $previous -and $count -gt 0) { $stableTicks++ } else { $stableTicks = 0 }
    if ($stableTicks -ge 3) {
        Write-Host "Stable at $count items."
        break
    }
    $previous = $count
}

if ($folder.Items.Count -gt 0) {
    $items = $folder.Items
    $items.Sort('[SentOn]', $false) | Out-Null
    $oldest = $items.GetFirst()
    $items.Sort('[SentOn]', $true) | Out-Null
    $newest = $items.GetFirst()

    try { Write-Host "Oldest cached: $($oldest.SentOn)" } catch { }
    try { Write-Host "Newest cached: $($newest.SentOn)" } catch { }
    Write-Host 'If the oldest date is ~12 months back, the offline cache window is capping you.'
}
