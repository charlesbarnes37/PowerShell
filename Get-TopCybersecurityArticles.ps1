<#
.SYNOPSIS
Fetches the latest cybersecurity articles from a set of RSS/Atom feeds.

.DESCRIPTION
This script downloads feeds from a collection of cybersecurity news sites,
selects the top articles from each site, and optionally opens them in the browser,
exports the results to CSV, or sends them via email.

.PARAMETER Top
The number of articles to retrieve per site. Default is 5.

.PARAMETER OpenInBrowser
Switch to open each article in the default browser after fetching.

.PARAMETER ExportPath
Optional path to export the results as CSV. Defaults to a file on the desktop.

.PARAMETER SendEmail
Switch to send the articles via email.

.PARAMETER EmailTo
Recipient email address. Required when SendEmail is used. Defaults to ctbarnes37@gmail.com.

.PARAMETER SMTPServer
SMTP server address. Defaults to smtp.gmail.com (requires app password for Gmail).

.PARAMETER SMTPPort
SMTP port number. Defaults to 587 (TLS).

.PARAMETER FromEmail
Sender email address. Required when SendEmail is used.

.PARAMETER EmailPassword
Sender email password or app password. Required when SendEmail is used.
#>

param(
    [int]$Top = 5,
    [switch]$OpenInBrowser,
    [string]$ExportPath = "$HOME\Desktop\CyberNews_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [switch]$SendEmail,
    [string]$EmailTo = "ctbarnes37@gmail.com",
    [string]$SMTPServer = "smtp.gmail.com",
    [int]$SMTPPort = 587,
    [string]$FromEmail,
    [string]$EmailPassword
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

function Send-ArticlesEmail {
    param(
        [Parameter(Mandatory)][System.Collections.Generic.List[PSObject]]$Articles,
        [Parameter(Mandatory)][string]$To,
        [Parameter(Mandatory)][string]$From,
        [Parameter(Mandatory)][string]$Password,
        [Parameter(Mandatory)][string]$SMTPServer,
        [Parameter(Mandatory)][int]$SMTPPort,
        [int]$SuccessfulFeeds,
        [int]$TotalFeeds
    )

    try {
        # Build HTML email body
        $htmlBody = @"
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body { font-family: Arial, sans-serif; color: #333; }
                .header { background-color: #2c3e50; color: white; padding: 20px; text-align: center; }
                .summary { padding: 10px 20px; background-color: #ecf0f1; }
                .article { border-left: 4px solid #3498db; padding: 15px; margin: 10px 0; background-color: #f9f9f9; }
                .site { color: #3498db; font-weight: bold; font-size: 12px; }
                .title { color: #2c3e50; font-weight: bold; font-size: 16px; margin: 10px 0; }
                .date { color: #7f8c8d; font-size: 12px; }
                .description { color: #555; line-height: 1.6; margin: 10px 0; }
                .link { color: #3498db; text-decoration: none; }
                .link:hover { text-decoration: underline; }
                .footer { padding: 20px; background-color: #ecf0f1; text-align: center; font-size: 12px; color: #7f8c8d; }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>🔍 Top Cybersecurity Articles</h1>
                <p>Latest articles from $TotalFeeds cybersecurity news sources</p>
            </div>
            
            <div class="summary">
                <p><strong>Summary:</strong> $($Articles.Count) articles retrieved from $SuccessfulFeeds/$TotalFeeds feeds</p>
                <p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
            </div>
"@

        foreach ($article in $Articles) {
            $pubDateStr = if ($article.PubDate) { $article.PubDate.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
            $htmlBody += @"
            <div class="article">
                <div class="site">[$($article.Site)]</div>
                <div class="title">$([System.Web.HttpUtility]::HtmlEncode($article.Title))</div>
                <div class="date">📅 $pubDateStr</div>
"@
            
            if (-not [string]::IsNullOrEmpty($article.Description)) {
                $shortDesc = if ($article.Description.Length -gt 200) {
                    $article.Description.Substring(0, 197) + '...'
                } else {
                    $article.Description
                }
                $htmlBody += @"
                <div class="description">$([System.Web.HttpUtility]::HtmlEncode($shortDesc))</div>
"@
            }

            if (-not [string]::IsNullOrWhiteSpace($article.Link)) {
                $htmlBody += @"
                <p><a class="link" href="$($article.Link)" target="_blank">🔗 Read Full Article</a></p>
"@
            }

            $htmlBody += "</div>`n"
        }

        $htmlBody += @"
            <div class="footer">
                <p>Cybersecurity Articles Digest | Powered by PowerShell</p>
            </div>
        </body>
        </html>
"@

        # Setup email credentials
        $secPassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($From, $secPassword)

        # Send email
        $emailParams = @{
            To              = $To
            From            = $From
            Subject         = "🔐 Top Cybersecurity Articles - $(Get-Date -Format 'yyyy-MM-dd')"
            Body            = $htmlBody
            BodyAsHtml      = $true
            SmtpServer      = $SMTPServer
            Port            = $SMTPPort
            UseSsl          = $true
            Credential      = $credential
            ErrorAction     = 'Stop'
        }

        Send-MailMessage @emailParams
        Write-Host "`n📧 Email sent successfully to: $To" -ForegroundColor Green
    }
    catch {
        Write-Host "`n❌ Failed to send email: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "🔍 Fetching top cybersecurity articles from $($CyberFeeds.Count) feeds..." -ForegroundColor Cyan

$allArticles = [System.Collections.Generic.List[PSObject]]::new()
$failedFeeds = [System.Collections.Generic.List[string]]::new()
$successfulFeeds = 0

foreach ($site in $CyberFeeds.GetEnumerator()) {
    try {
        Write-Host "   → Fetching from $($site.Key)..." -ForegroundColor Gray
        $response = Invoke-WebRequest -Uri $site.Value -TimeoutSec 20 -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
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

$allArticles = $allArticles | Sort-Object @{Expression = {$_.PubDate -as [DateTime]}; Descending = $true}

Write-Host "`n🚀 Top Cybersecurity Articles" -ForegroundColor Green
Write-Host "====================================`n" -ForegroundColor Green

foreach ($article in $allArticles) {
    Write-Host "[$($article.Site)]" -ForegroundColor Cyan
    Write-Host "Title: $($article.Title)" -ForegroundColor White
    Write-Host "Link : $($article.Link)" -ForegroundColor Yellow

    if ($article.PubDate) {
        Write-Host "Date : $($article.PubDate.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Gray
    }

    if (-not [string]::IsNullOrEmpty($article.Description)) {
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
                Start-Process -Uri $article.Link
                Start-Sleep -Milliseconds 750
            }
            catch {
                Write-Host "   ⚠️  Could not open link for '$($article.Title)': $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

if ($SendEmail) {
    if ([string]::IsNullOrWhiteSpace($FromEmail) -or [string]::IsNullOrWhiteSpace($EmailPassword)) {
        Write-Host "`n❌ SendEmail requires -FromEmail and -EmailPassword parameters" -ForegroundColor Red
    } else {
        Send-ArticlesEmail -Articles $allArticles -To $EmailTo -From $FromEmail -Password $EmailPassword `
            -SMTPServer $SMTPServer -SMTPPort $SMTPPort -SuccessfulFeeds $successfulFeeds -TotalFeeds $CyberFeeds.Count
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
