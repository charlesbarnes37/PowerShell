<#
.SYNOPSIS
Fetches the latest cybersecurity articles from a set of RSS/Atom feeds.

.DESCRIPTION
This script downloads feeds from a collection of cybersecurity news sites,
selects the top articles from each site, and optionally opens them in the browser
or exports the results to CSV.

.PARAMETER Top
The number of articles to retrieve per site. Default is 5.

.PARAMETER OpenInBrowser
Switch to open each article in the default browser after fetching.

.PARAMETER ExportPath
Optional path to export the results as CSV. Defaults to a file on the desktop.
#>

param(
    [int]$Top = 5,
    [switch]$OpenInBrowser,
    [string]$ExportPath = "$HOME\Desktop\CyberNews_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$CyberFeeds = [ordered]@{
    "The Hacker News"   = "https://feeds.feedburner.com/TheHackersNews"
    "Dark Reading"      = "https://www.darkreading.com/rss.xml"
    "Bleeping Computer" = "https://www.bleepingcomputer.com/feed/"
    "Krebs on Security" = "https://krebsonsecurity.com/feed/"
    "SecurityWeek"      = "https://feeds.feedburner.com/Securityweek"
    "Sophos"            = "https://news.sophos.com/en-us/feed/"
    "CSO Online"        = "https://www.csoonline.com/feed/"
    "CyberScoop"        = "https://cyberscoop.com/feed/"
    "Help Net Security" = "https://www.helpnetsecurity.com/feed/"
}

function Get-ContentValue {
    param(
        [Parameter(Mandatory)]$Node,
        [Parameter(Mandatory)] [string[]]$PropertyNames
    )

    foreach ($propertyName in $PropertyNames) {
        $value = $Node.$propertyName
        if ($null -ne $value) {
            if ($value -is [System.Xml.XmlNodeList]) {
                if ($value.Count -gt 0) { return $value[0].InnerText.Trim() }
            }
            elseif ($value -is [System.Xml.XmlNode]) {
                return $value.InnerText.Trim()
            }
            else {
                return $value.ToString().Trim()
            }
        }
    }

    return $null
}

function Get-LinkValue {
    param(
        [Parameter(Mandatory)]$Item
    )

    if ($null -eq $Item.link) { return $null }

    if ($Item.link.href) {
        return $Item.link.href.ToString().Trim()
    }

    if ($Item.link -is [System.Xml.XmlNodeList]) {
        foreach ($linkNode in $Item.link) {
            if ($linkNode.href) { return $linkNode.href.ToString().Trim() }
            if ($linkNode.GetAttribute('href')) { return $linkNode.GetAttribute('href').Trim() }
            if ($linkNode.InnerText) { return $linkNode.InnerText.Trim() }
        }
    }

    if ($Item.link -is [System.Xml.XmlNode]) {
        return $Item.link.InnerText.Trim()
    }

    return $Item.link.ToString().Trim()
}

function Get-PubDate {
    param(
        [Parameter(Mandatory)]$Item
    )

    $dateCandidates = @(
        'pubDate',
        'published',
        'updated',
        'dc:date',
        'date'
    )

    foreach ($name in $dateCandidates) {
        $value = Get-ContentValue -Node $Item -PropertyNames @($name)
        if ($value) {
            try {
                return [DateTime]::Parse($value)
            }
            catch {
                continue
            }
        }
    }

    return $null
}

function Get-Description {
    param(
        [Parameter(Mandatory)]$Item
    )

    $description = Get-ContentValue -Node $Item -PropertyNames @('description', 'summary', 'content', 'subtitle')
    if ($description) {
        return ($description -replace '<[^>]+>', '').Trim()
    }

    return ''
}

function Get-FeedItems {
    param(
        [Parameter(Mandatory)] [xml]$Xml
    )

    if ($Xml.rss.channel.item) { return $Xml.rss.channel.item }
    if ($Xml.feed.entry) { return $Xml.feed.entry }
    if ($Xml.channel.item) { return $Xml.channel.item }
    return @()
}

Write-Host "🔍 Fetching top cybersecurity articles from $($CyberFeeds.Count) feeds..." -ForegroundColor Cyan

$allArticles = [System.Collections.Generic.List[PSObject]]::new()
$failedFeeds = [System.Collections.Generic.List[string]]::new()
$successfulFeeds = 0

foreach ($site in $CyberFeeds.GetEnumerator()) {
    try {
        Write-Host "   → Fetching from $($site.Key)..." -ForegroundColor Gray
        $response = Invoke-WebRequest -Uri $site.Value -TimeoutSec 20
        [xml]$xml = $response.Content

        $items = Get-FeedItems -Xml $xml | Select-Object -First $Top
        if (-not $items) {
            Write-Host "   ⚠️  No feed items found for $($site.Key)." -ForegroundColor Yellow
            continue
        }

        foreach ($item in $items) {
            $article = [PSCustomObject]@{
                Site        = $site.Key
                Title       = Get-ContentValue -Node $item -PropertyNames @('title')
                Link        = Get-LinkValue -Item $item
                PubDate     = Get-PubDate -Item $item
                Description = Get-Description -Item $item
            }

            $allArticles.Add($article)
        }

        $successfulFeeds++
    }
    catch {
        $failedFeeds.Add($site.Key)
        Write-Host "   ❌ Failed to fetch $($site.Key): $($_.Exception.Message)" -ForegroundColor Red
    }
}

if ($allArticles.Count -eq 0) {
    Write-Host "No articles were retrieved. Check your network connection or feed URLs." -ForegroundColor Yellow
    return
}

$allArticles = $allArticles | Sort-Object PubDate -Descending

Write-Host "`n🚀 Top Cybersecurity Articles" -ForegroundColor Green
Write-Host "====================================`n" -ForegroundColor Green

foreach ($article in $allArticles) {
    Write-Host "[$($article.Site)]" -ForegroundColor Cyan
    Write-Host "Title: $($article.Title)" -ForegroundColor White
    Write-Host "Link : $($article.Link)" -ForegroundColor Yellow

    if ($article.PubDate) {
        Write-Host "Date : $($article.PubDate.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Gray
    }

    if ($article.Description) {
        $shortDesc = if ($article.Description.Length -gt 150) {
            $article.Description.Substring(0, 147) + '...'
        } else {
            $article.Description
        }
        Write-Host "Desc : $shortDesc" -ForegroundColor Gray
    }

    Write-Host ('-' * 80) -ForegroundColor DarkGray
}

if ($OpenInBrowser) {
    Write-Host "`n🌐 Opening article links in the browser..." -ForegroundColor Magenta
    foreach ($article in $allArticles) {
        if (-not [string]::IsNullOrWhiteSpace($article.Link)) {
            try {
                Start-Process -FilePath $article.Link
                Start-Sleep -Milliseconds 750
            }
            catch {
                Write-Host "   ⚠️  Could not open link for '$($article.Title)': $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

try {
    $exportDirectory = Split-Path -Parent $ExportPath
    if (-not (Test-Path -Path $exportDirectory)) {
        New-Item -ItemType Directory -Path $exportDirectory -Force | Out-Null
    }

    $allArticles | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "`n💾 Results exported to: $ExportPath" -ForegroundColor Green
}
catch {
    Write-Host "`n⚠️  Unable to export CSV: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`nSummary: $($allArticles.Count) articles from $successfulFeeds/$($CyberFeeds.Count) successful feeds." -ForegroundColor Cyan
if ($failedFeeds.Count -gt 0) {
    Write-Host "Failed feeds: $($failedFeeds -join ', ')" -ForegroundColor Yellow
}
