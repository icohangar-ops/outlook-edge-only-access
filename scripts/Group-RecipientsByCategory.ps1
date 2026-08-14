<#
.SYNOPSIS
    Applies a domain -> category map to Export-SentRecipients output.
.DESCRIPTION
    Splits recipients into categorised contacts and a review pile. Domains not
    present in the map are routed to 'Review - Unclassified' rather than being
    guessed at, because guessing a category from a domain name is how people
    end up mislabelled in a document that matters.

    Free-mail domains go to 'Review - Personal address': plenty of real
    counterparties correspond from personal addresses, and the domain tells you
    nothing either way.
.PARAMETER InCsv
    CSV produced by Export-SentRecipients.ps1.
.PARAMETER MapPath
    .psd1 category map. See examples/category-map.sample.psd1.
.PARAMETER OutDir
    Directory for the output CSVs.
.PARAMETER InternalDomain
    Your own domains, excluded from the output entirely.
.EXAMPLE
    .\Group-RecipientsByCategory.ps1 -InCsv .\recipients.csv -MapPath .\map.psd1 -OutDir .\out -InternalDomain 'company.com'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]   $InCsv,
    [Parameter(Mandatory)][string]   $MapPath,
    [Parameter(Mandatory)][string]   $OutDir,
    [string[]] $InternalDomain = @(),
    [string[]] $PersonalDomain = @('gmail.com', 'hotmail.com', 'outlook.com', 'yahoo.com', 'icloud.com', 'me.com', 'aol.com', 'proton.me', 'pm.me')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$map = Import-PowerShellDataFile -Path $MapPath
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

$rows = Import-Csv -Path $InCsv | Where-Object { $InternalDomain -notcontains $_.Domain }

$classified = foreach ($row in $rows) {
    $domain = $row.Domain.ToLowerInvariant()

    if ($map.ContainsKey($domain)) {
        $category = $map[$domain][0]
        $firm     = $map[$domain][1]
    } elseif ($PersonalDomain -contains $domain) {
        $category = 'Review - Personal address'
        $firm     = ''
    } else {
        $category = 'Review - Unclassified'
        $firm     = ''
    }

    [pscustomobject]@{
        Category     = $category
        Firm         = $firm
        Name         = $row.Name
        Email        = $row.Email
        Domain       = $row.Domain
        FirstContact = $row.FirstContact
        LastContact  = $row.LastContact
        Messages     = [int]$row.Messages
        Evidence     = $row.SampleSubjects
    }
}

$sortByVolume = @{ Expression = 'Messages'; Descending = $true }

$capital = $classified |
    Where-Object { $_.Category -match '^(Investor|Advisor|Lender)' } |
    Sort-Object Category, $sortByVolume

$review = $classified |
    Where-Object { $_.Category -like 'Review*' } |
    Sort-Object $sortByVolume

$capital    | Export-Csv (Join-Path $OutDir 'capital-contacts.csv')  -NoTypeInformation -Encoding UTF8
$review     | Export-Csv (Join-Path $OutDir 'needs-review.csv')      -NoTypeInformation -Encoding UTF8
$classified | Sort-Object Category, $sortByVolume |
              Export-Csv (Join-Path $OutDir 'all-classified.csv')    -NoTypeInformation -Encoding UTF8

Write-Host '=== counts by category ==='
$classified | Group-Object Category | Sort-Object Count -Descending |
    ForEach-Object { '{0,5}  {1}' -f $_.Count, $_.Name }

Write-Host ''
Write-Host "capital contacts : $($capital.Count)"
Write-Host "distinct firms   : $(($capital | Select-Object -ExpandProperty Firm -Unique).Count)"
Write-Host "needs review     : $($review.Count)"
Write-Host ''
Write-Host 'Skim needs-review.csv before treating the capital list as complete.' -ForegroundColor Yellow
