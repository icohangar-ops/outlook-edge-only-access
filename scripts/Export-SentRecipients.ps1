<#
.SYNOPSIS
    Aggregates every recipient you have sent mail to, with first/last contact
    and message counts.
.DESCRIPTION
    Walks Sent Items in the default store, resolves each recipient to a real
    SMTP address, and emits one row per distinct address. Reads only.

    BCC recipients are deliberately excluded.

    Coverage is limited to what Outlook has cached offline - see
    Get-MailboxDiagnostics.ps1. If the window is 12 months you will get 12
    months, regardless of the -Since value.
.PARAMETER Since
    Earliest send date to include. Defaults to 18 months ago.
.PARAMETER OutCsv
    Optional path to write results as CSV.
.PARAMETER MaxItems
    Stop after this many messages. 0 means no cap. Useful for a quick smoke run.
.EXAMPLE
    .\Export-SentRecipients.ps1 -MaxItems 200
.EXAMPLE
    .\Export-SentRecipients.ps1 -Since '2025-02-01' -OutCsv .\recipients.csv
#>
[CmdletBinding()]
param(
    [datetime] $Since = (Get-Date).AddMonths(-18),
    [string]   $OutCsv,
    [int]      $MaxItems = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\src\OutlookMapiTools\OutlookMapiTools.psd1') -Force

$OlFolderSentMail = 5
$OlMailItem       = 43
$OlBcc            = 3

$ns    = Connect-OutlookMapi
$sent  = $ns.GetDefaultFolder($OlFolderSentMail)
$items = $sent.Items
$items.Sort('[SentOn]', $true) | Out-Null

# Restrict uses US-format dates regardless of locale.
$filter   = "[SentOn] >= '" + $Since.ToString('MM/dd/yyyy') + " 12:00 AM'"
$filtered = $items.Restrict($filter)

Write-Host "Sent Items total : $($items.Count)"
Write-Host "Since $($Since.ToString('yyyy-MM-dd')) : $($filtered.Count)"

$people  = @{}
$scanned = 0
$skipped = 0

$mail = $filtered.GetFirst()
while ($null -ne $mail) {
    if ($MaxItems -gt 0 -and $scanned -ge $MaxItems) { break }
    try {
        if ($mail.Class -eq $OlMailItem) {
            $sentOn  = $mail.SentOn
            $subject = $mail.Subject

            foreach ($recipient in $mail.Recipients) {
                if ($recipient.Type -eq $OlBcc) { continue }

                $address = Resolve-SmtpAddress -Recipient $recipient
                if (-not (Test-SmtpAddress -Address $address)) { continue }
                $key = $address.ToLowerInvariant().Trim()

                if (-not $people.ContainsKey($key)) {
                    $people[$key] = [pscustomobject]@{
                        Email        = $key
                        Name         = $recipient.Name
                        Domain       = $key.Split('@')[1]
                        FirstContact = $sentOn
                        LastContact  = $sentOn
                        Messages     = 0
                        Subjects     = [System.Collections.Generic.List[string]]::new()
                    }
                }

                $person = $people[$key]
                $person.Messages++
                if ($sentOn -lt $person.FirstContact) { $person.FirstContact = $sentOn }
                if ($sentOn -gt $person.LastContact)  { $person.LastContact  = $sentOn }
                if ($person.Subjects.Count -lt 5 -and $subject) { $person.Subjects.Add($subject) }
                if ([string]::IsNullOrWhiteSpace($person.Name) -and $recipient.Name) { $person.Name = $recipient.Name }
            }
        }
        $scanned++
    } catch {
        $skipped++
    }
    $mail = $filtered.GetNext()
}

Write-Host "Scanned          : $scanned"
Write-Host "Skipped          : $skipped"
Write-Host "Distinct people  : $($people.Count)"

$rows = $people.Values |
    Sort-Object -Property @{ Expression = 'Messages'; Descending = $true } |
    Select-Object Email, Name, Domain,
        @{ n = 'FirstContact';   e = { $_.FirstContact.ToString('yyyy-MM-dd') } },
        @{ n = 'LastContact';    e = { $_.LastContact.ToString('yyyy-MM-dd') } },
        Messages,
        @{ n = 'SampleSubjects'; e = { ($_.Subjects -join ' | ') } }

if ($OutCsv) {
    $rows | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Wrote: $OutCsv"
}

$rows
