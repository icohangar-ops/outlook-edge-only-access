<#
.SYNOPSIS
    Finds mail attachments across the mailbox by filename, subject or body, and
    optionally saves them to disk.
.DESCRIPTION
    Searching only on filename misses things: a "tolling model" may be filed as
    "Unit Economics Model - Offtake.xlsx". Search the subject and body too, then
    look at what came attached.

    Reads mail; the only thing written is saved attachments under -SaveTo.
.PARAMETER FileNamePattern
    Regex matched against attachment filenames.
.PARAMETER SubjectPattern
    Substring matched against message subjects (SQL LIKE, case-insensitive).
.PARAMETER BodyPattern
    Substring matched against message bodies (SQL LIKE, case-insensitive).
.PARAMETER Extension
    Regex of file extensions to keep. Defaults to spreadsheets.
.PARAMETER SaveTo
    Directory to save matches into. Files are prefixed with the message date.
.EXAMPLE
    .\Find-MailAttachment.ps1 -SubjectPattern 'tolling'
.EXAMPLE
    .\Find-MailAttachment.ps1 -BodyPattern 'Glencore' -Extension '\.(xls[xmb]?|csv)$'
.EXAMPLE
    .\Find-MailAttachment.ps1 -FileNamePattern 'Unit Economics' -SaveTo .\out
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $FileNamePattern,
    [string] $SubjectPattern,
    [string] $BodyPattern,
    [string] $Extension = '\.(xls[xmb]?|csv)$',
    [string] $SaveTo
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\src\OutlookMapiTools\OutlookMapiTools.psd1') -Force

$OlMailItem = 43
$quote      = [char]34

$ns      = Connect-OutlookMapi
$folders = Get-OutlookMailFolder -Namespace $ns

# Build DASL restrictions. hasattachment alone is the cheapest broad filter.
$restrictions = [System.Collections.Generic.List[string]]::new()
if ($SubjectPattern) {
    $restrictions.Add("@SQL=$quote" + 'urn:schemas:httpmail:subject' + "$quote like '%$SubjectPattern%'")
}
if ($BodyPattern) {
    $restrictions.Add("@SQL=$quote" + 'urn:schemas:httpmail:textdescription' + "$quote like '%$BodyPattern%'")
}
if ($restrictions.Count -eq 0) {
    $restrictions.Add("@SQL=$quote" + 'urn:schemas:httpmail:hasattachment' + "$quote = 1")
}

$hits = [System.Collections.Generic.List[object]]::new()
$seen = @{}

foreach ($folder in $folders) {
    foreach ($restriction in $restrictions) {
        try { $results = $folder.Items.Restrict($restriction) } catch { continue }
        try { $message = $results.GetFirst() } catch { continue }

        while ($null -ne $message) {
            try {
                if ($message.Class -eq $OlMailItem -and -not $seen.ContainsKey($message.EntryID)) {
                    $seen[$message.EntryID] = $true

                    $count = 0
                    try { $count = $message.Attachments.Count } catch { }

                    for ($i = 1; $i -le $count; $i++) {
                        try {
                            $attachment = $message.Attachments.Item($i)
                            $name       = $attachment.FileName

                            if ($Extension -and $name -notmatch $Extension) { continue }
                            if ($FileNamePattern -and $name -notmatch $FileNamePattern) { continue }

                            # Resolve these before building the object - try/catch
                            # is not valid as an inline expression in a hashtable.
                            $date = $null
                            try {
                                $date = if ($message.SentOn) { $message.SentOn } else { $message.ReceivedTime }
                            } catch { }

                            $sender = ''
                            try { $sender = $message.SenderName } catch { }

                            $sizeKb = 0
                            try { $sizeKb = [math]::Round($attachment.Size / 1KB, 0) } catch { }

                            $hits.Add([pscustomobject]@{
                                Date     = $date
                                Folder   = $folder.Name
                                From     = $sender
                                Subject  = $message.Subject
                                FileName = $name
                                SizeKB   = $sizeKb
                                EntryID  = $message.EntryID
                                Index    = $i
                            })
                        } catch { }
                    }
                }
            } catch { }
            try { $message = $results.GetNext() } catch { break }
        }
    }
}

Write-Host "matches: $($hits.Count)"

if ($SaveTo -and $hits.Count -gt 0) {
    if (-not (Test-Path $SaveTo)) { New-Item -ItemType Directory -Path $SaveTo -Force | Out-Null }

    foreach ($hit in $hits) {
        $prefix = if ($hit.Date) { $hit.Date.ToString('yyyy-MM-dd') } else { 'undated' }
        $target = Join-Path $SaveTo "$prefix - $($hit.FileName)"
        if (Test-Path $target) { continue }

        if ($PSCmdlet.ShouldProcess($target, 'Save attachment')) {
            try {
                $message    = $ns.GetItemFromID($hit.EntryID)
                $attachment = $message.Attachments.Item($hit.Index)
                $attachment.SaveAsFile($target)
            } catch {
                Write-Warning "Could not save $($hit.FileName): $($_.Exception.Message)"
            }
        }
    }
}

$hits | Sort-Object Date -Descending
