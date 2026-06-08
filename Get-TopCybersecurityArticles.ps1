<#
.SYNOPSIS
Fetches the latest cybersecurity articles from a set of RSS feeds.

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

$CyberFeeds = @{
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

Write-Host "🔍 Fetching top cybersecurity articles from $($CyberFeeds.Count) feeds..." -ForegroundColor Cyan

$allArticles = [System.Collections.Generic.List[PSObject]]::new()

foreach ($site in $CyberFeeds.GetEnumerator()) {
    try {
        Write-Host "   → Fetching from $($site.Key)..." -ForegroundColor Gray

        $response = Invoke-WebRequest -Uri $site.Value -UseBasicParsing -TimeoutSec 20
        [xml]$xml = $response.Content

        $items = $xml.rss.channel.item | Select-Object -First $Top

        foreach ($item in $items) {
            $pubDate = $null
            if ($item.pubDate) {
                try {
                    $pubDate = [DateTime]::Parse($item.pubDate)
                }
                catch {
                    $pubDate = $null
                }
            }

            $description = if ($item.description) { ($item.description -replace '<[^>]+>', '').Trim() } else { '' }

            $article = [PSCustomObject]@{
                Site        = $site.Key
                Title       = $item.title
                Link        = $item.link
                PubDate     = $pubDate
                Description = $description
            }

            $allArticles.Add($article)
        }
    }
    catch {
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
    Write-Host "`n🌐 Opening all article links in the browser..." -ForegroundColor Magenta
    foreach ($article in $allArticles) {
        Start-Process -FilePath $article.Link
        Start-Sleep -Milliseconds 750
    }
}

try {
    $allArticles | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "`n💾 Results exported to: $ExportPath" -ForegroundColor Green
}
catch {
    Write-Host "`n⚠️  Unable to export CSV: $($_.Exception.Message)" -ForegroundColor Yellow
}
