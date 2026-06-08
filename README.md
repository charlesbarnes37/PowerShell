# Working file for VS
CMFB

param(
    [int]$Top = 5,                    # Number of articles per site
    [switch]$OpenInBrowser = $false   # Open all links in browser
)

# ==================== CYBERSECURITY RSS FEEDS ====================
$CyberFeeds = @{
    "The Hacker News"     = "https://feeds.feedburner.com/TheHackersNews"
    "Dark Reading"        = "https://www.darkreading.com/rss.xml"
    "Bleeping Computer"   = "https://www.bleepingcomputer.com/feed/"
    "Krebs on Security"   = "https://krebsonsecurity.com/feed/"
    "SecurityWeek"        = "https://feeds.feedburner.com/Securityweek"
    "Sophos"              = "https://news.sophos.com/en-us/feed/"
    "CSO Online"          = "https://www.csoonline.com/feed/"
    "CyberScoop"          = "https://cyberscoop.com/feed/"
    "Help Net Security"   = "https://www.helpnetsecurity.com/feed/"
}

Write-Host "🔍 Fetching latest cybersecurity articles..." -ForegroundColor Cyan

$allArticles = @()

foreach ($site in $CyberFeeds.GetEnumerator()) {
    try {
        Write-Host "   → Fetching from $($site.Key)..." -ForegroundColor Gray
        
        $response = Invoke-WebRequest -Uri $site.Value -UseBasicParsing -TimeoutSec 15
        [xml]$xml = $response.Content

        $items = $xml.rss.channel.item | Select-Object -First $Top

        foreach ($item in $items) {
            $article = [PSCustomObject]@{
                Site        = $site.Key
                Title       = $item.title
                Link        = $item.link
                PubDate     = if ($item.pubDate) { [DateTime]::Parse($item.pubDate) } else { $null }
                Description = if ($item.description) { ($item.description -replace '<[^>]+>', '').Trim() } else { "" }
            }
            $allArticles += $article
        }
    }
    catch {
        Write-Host "   ❌ Failed to fetch $($site.Key): $_" -ForegroundColor Red
    }
}

# Sort by publish date (newest first)
$allArticles = $allArticles | Sort-Object PubDate -Descending

# ==================== DISPLAY RESULTS ====================
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
            $article.Description.Substring(0, 147) + "..." 
        } else { 
            $article.Description 
        }
        Write-Host "Desc : $shortDesc" -ForegroundColor Gray
    }
    
    Write-Host "-" * 80 -ForegroundColor DarkGray
}

# Optional: Open all links in browser
if ($OpenInBrowser) {
    Write-Host "`n🌐 Opening all articles in browser..." -ForegroundColor Magenta
    foreach ($article in $allArticles) {
        Start-Process $article.Link
        Start-Sleep -Milliseconds 800  # Be gentle with browser
    }
}

# Export to CSV option
$exportPath = "$HOME\Desktop\CyberNews_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
$allArticles | Export-Csv -Path $exportPath -NoTypeInformation
Write-Host "`n💾 Results exported to: $exportPath" -ForegroundColor Green