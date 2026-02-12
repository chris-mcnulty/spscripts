<#
.SYNOPSIS
    Enumerate usage reports from all SharePoint sites in a tenant to a CSV file.

.DESCRIPTION
    This script connects to SharePoint Online and retrieves usage information for all sites 
    in the tenant, then exports the data to a CSV file. The script uses Microsoft Graph API
    or SharePoint Online Management Shell to collect site usage statistics.

.PARAMETER TenantName
    The name of your SharePoint Online tenant (e.g., 'contoso' for contoso.sharepoint.com).

.PARAMETER OutputPath
    The path where the CSV file will be saved. Defaults to current directory with timestamp.

.PARAMETER AuthMethod
    The authentication method to use: 'Interactive' (default), 'Certificate', or 'ClientSecret'.

.PARAMETER UseGraphAPI
    If specified, uses Microsoft Graph API instead of SharePoint Online Management Shell.
    Requires Microsoft.Graph PowerShell modules.

.PARAMETER UseCombined
    If specified, combines data from both SPO Management Shell and Graph API.
    Produces a single report with friendly site names/owners from SPO and
    page view/activity metrics from Graph. Requires both sets of modules.

.EXAMPLE
    .\Get-SPOSiteUsageReports.ps1 -TenantName "contoso"
    Connects to contoso.sharepoint.com and exports usage reports to a timestamped CSV file.

.EXAMPLE
    .\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -OutputPath "C:\Reports\usage.csv" -UseGraphAPI
    Uses Microsoft Graph API to get usage data and exports to specified path.

.EXAMPLE
    .\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -UseCombined
    Combines SPO Management Shell (friendly names) and Graph API (page views) into one report.

.NOTES
    Author: chris-mcnulty/spscripts
    Requirements:
    - For SharePoint method: SharePointPnPPowerShellOnline or Microsoft.Online.SharePoint.PowerShell module
    - For Graph method: Microsoft.Graph.Reports module
    - Appropriate permissions: SharePoint Admin or Global Admin
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter your SharePoint tenant name (e.g., 'contoso' for contoso.sharepoint.com)")]
    [string]$TenantName,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
    [string]$AuthMethod = 'Interactive',

    [Parameter(Mandatory = $false)]
    [switch]$UseGraphAPI,

    [Parameter(Mandatory = $false, HelpMessage = "Combines SPO Management Shell (friendly names) with Graph API (page views/activity) into a single report")]
    [switch]$UseCombined
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Function to check and install required modules
function Install-RequiredModules {
    param(
        [string[]]$ModuleNames
    )
    
    foreach ($moduleName in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            Write-Host "Module '$moduleName' not found. Attempting to install..." -ForegroundColor Yellow
            try {
                Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
                Write-Host "Module '$moduleName' installed successfully." -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to install module '$moduleName': $_"
                return $false
            }
        }
    }
    return $true
}

# Function to get usage reports using Microsoft Graph API
function Get-UsageReportsViaGraph {
    param(
        [string]$TenantName
    )
    
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    
    try {
        # Connect to Microsoft Graph — include ReportSettings.ReadWrite.All so the
        # script can detect and disable the report-privacy concealment setting.
        Connect-MgGraph -Scopes "Reports.Read.All", "Sites.Read.All", "ReportSettings.ReadWrite.All" -NoWelcome
        
        # Proactively check the admin concealment setting before pulling the report.
        $privacyResult = Resolve-GraphReportPrivacy
        
        Write-Host "Retrieving SharePoint site usage data from Microsoft Graph..." -ForegroundColor Cyan
        
        # Get SharePoint site usage details for the last 7 days
        $usageData = @()
        
        # Get site usage detail report — this calls the Graph getSharePointSiteUsageDetail
        # endpoint, which returns CSV data including Page View Count and Visited Page Count.
        # Use a temp file because the cmdlet requires -OutFile to save CSV output.
        # Suppress progress to avoid PercentComplete overflow bug in the Graph SDK.
        $tempFile = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + ".csv")
        $previousProgressPreference = $ProgressPreference
        try {
            $ProgressPreference = 'SilentlyContinue'
            Get-MgReportSharePointSiteUsageDetail -Period D7 -OutFile $tempFile
        }
        finally {
            $ProgressPreference = $previousProgressPreference
        }
        
        # Parse the CSV data returned by Graph API
        $sites = Import-Csv -Path $tempFile
        
        foreach ($site in $sites) {
            $siteInfo = [PSCustomObject]@{
                SiteUrl              = $site.'Site URL'
                SiteId               = $site.'Site Id'
                OwnerDisplayName     = $site.'Owner Display Name'
                OwnerPrincipalName   = $site.'Owner Principal Name'
                IsDeleted            = $site.'Is Deleted'
                LastActivityDate     = $site.'Last Activity Date'
                FileCount            = $site.'File Count'
                ActiveFileCount      = $site.'Active File Count'
                PageViewCount        = $site.'Page View Count'
                VisitedPageCount     = $site.'Visited Page Count'
                StorageUsedInBytes   = $site.'Storage Used (Byte)'
                StorageAllocatedInBytes = $site.'Storage Allocated (Byte)'
                RootWebTemplate      = $site.'Root Web Template'
                ReportRefreshDate    = $site.'Report Refresh Date'
                ReportPeriod         = $site.'Report Period'
            }
            $usageData += $siteInfo
        }
        
        Write-Host "Retrieved usage data for $($usageData.Count) sites." -ForegroundColor Green

        # --- Resolve blank URLs via Get-MgSite for sites with valid SiteIds ---
        # The Graph report may return real SiteIds but blank URLs when the
        # concealment setting is active.  Get-MgSite is NOT affected by report
        # concealment and returns real displayName and webUrl from a SiteId.
        $emptyGuid = '00000000-0000-0000-0000-000000000000'
        $blankUrlSites = $usageData | Where-Object {
            [string]::IsNullOrWhiteSpace($_.SiteUrl) -and
            $_.SiteId -and $_.SiteId -ne $emptyGuid
        }

        if ($blankUrlSites.Count -gt 0) {
            Write-Host "Resolving $($blankUrlSites.Count) sites with blank URLs via Get-MgSite..." -ForegroundColor Yellow
            $resolvedCount = 0

            foreach ($blankSite in $blankUrlSites) {
                try {
                    $mgSite = Get-MgSite -SiteId $blankSite.SiteId -Property "id,displayName,webUrl" -ErrorAction Stop
                    if ($mgSite) {
                        if ($mgSite.WebUrl)      { $blankSite.SiteUrl = $mgSite.WebUrl }
                        if ($mgSite.DisplayName)  { $blankSite.OwnerDisplayName = $mgSite.DisplayName }
                        $resolvedCount++
                    }
                }
                catch {
                    # Could not resolve this SiteId — leave as-is
                }
            }

            Write-Host "Resolved $resolvedCount of $($blankUrlSites.Count) blank-URL sites via Get-MgSite." -ForegroundColor Green
        }

        # --- Obfuscation detection and Sites API fallback ---
        if (Test-GraphDataObfuscated -ReportData $usageData) {
            Write-Warning "Graph report data is obfuscated — site URLs, IDs, and owner names are concealed."

            if ($privacyResult.WasEnabled -eq $true) {
                Write-Warning "The report-privacy setting may have been recently changed. Cached report data can take up to 48 hours to reflect the new setting."
            }

            # Fall back to the Sites API which is unaffected by report concealment.
            $realSites = Get-SiteMetadataViaGraph

            if ($realSites.Count -gt 0) {
                Write-Host "Enriching site metadata with per-site analytics (not affected by report concealment)..." -ForegroundColor Yellow

                # Get per-site analytics — these endpoints are NOT subject to report privacy.
                $siteAnalytics = Get-PerSiteAnalyticsViaGraph -Sites $realSites

                $usageData = @()
                foreach ($realSite in $realSites) {
                    $enrichment = if ($realSite.webUrl) { $siteAnalytics[$realSite.webUrl] } else { $null }
                    $usageData += [PSCustomObject]@{
                        SiteUrl                 = $realSite.webUrl
                        SiteId                  = $realSite.id
                        OwnerDisplayName        = $realSite.displayName
                        OwnerPrincipalName      = ''
                        IsDeleted               = $false
                        LastActivityDate        = if ($enrichment -and $null -ne $enrichment.LastActivityDate) { $enrichment.LastActivityDate } else { '' }
                        FileCount               = if ($enrichment -and $null -ne $enrichment.FileCount) { $enrichment.FileCount } else { '' }
                        ActiveFileCount         = ''
                        PageViewCount           = if ($enrichment -and $null -ne $enrichment.PageViewCount) { $enrichment.PageViewCount } else { '' }
                        VisitedPageCount        = ''
                        StorageUsedInBytes      = if ($enrichment -and $null -ne $enrichment.StorageUsedBytes) { $enrichment.StorageUsedBytes } else { '' }
                        StorageAllocatedInBytes = ''
                        RootWebTemplate         = ''
                        ReportRefreshDate       = ''
                        ReportPeriod            = ''
                    }
                }
                Write-Host "Rebuilt report with $($usageData.Count) sites using real site metadata and per-site analytics." -ForegroundColor Green
            }
            else {
                Write-Warning "Could not retrieve site metadata from Sites API. Report will contain obfuscated data."
                Write-Warning "Ensure the 'Conceal user, group, and site names in all reports' setting is disabled and re-run after 48 hours."
            }
        }

        return $usageData
    }
    catch {
        Write-Error "Error retrieving data from Microsoft Graph: $_"
        throw
    }
    finally {
        if ($tempFile) { Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue }
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}

# Function to get usage reports using SharePoint Online Management Shell
function Get-UsageReportsViaSPO {
    param(
        [string]$TenantName
    )
    
    $adminUrl = "https://$TenantName-admin.sharepoint.com"
    
    Write-Host "Connecting to SharePoint Online Admin Center: $adminUrl" -ForegroundColor Cyan
    
    try {
        # Connect to SharePoint Online
        Connect-SPOService -Url $adminUrl
        
        Write-Host "Retrieving all SharePoint sites..." -ForegroundColor Cyan
        
        # Get all site collections
        $sites = Get-SPOSite -Limit All -IncludePersonalSite $false
        
        Write-Host "Found $($sites.Count) sites. Collecting usage data..." -ForegroundColor Cyan
        
        $usageData = @()
        $counter = 0
        
        foreach ($site in $sites) {
            $counter++
            Write-Progress -Activity "Processing Sites" -Status "Processing $counter of $($sites.Count): $($site.Url)" -PercentComplete (($counter / $sites.Count) * 100)
            
            try {
                # Get site details
                $siteDetails = Get-SPOSite -Identity $site.Url -Detailed
                
                $siteInfo = [PSCustomObject]@{
                    SiteUrl                     = $site.Url
                    Title                       = $site.Title
                    Owner                       = $site.Owner
                    Template                    = $site.Template
                    Status                      = $site.Status
                    StorageQuotaMB              = [math]::Round($site.StorageQuota / 1024, 2)
                    StorageUsedMB               = [math]::Round($site.StorageUsageCurrent / 1024, 2)
                    StorageUsedPercentage       = if ($site.StorageQuota -gt 0) { [math]::Round(($site.StorageUsageCurrent / $site.StorageQuota) * 100, 2) } else { 0 }
                    LastContentModifiedDate     = $site.LastContentModifiedDate
                    SharingCapability           = $site.SharingCapability
                    LockState                   = $site.LockState
                    PWAEnabled                  = $site.PWAEnabled
                    ConditionalAccessPolicy     = $site.ConditionalAccessPolicy
                    AllowSelfServiceUpgrade     = $site.AllowSelfServiceUpgrade
                    LocaleId                    = $siteDetails.LocaleId
                    CompatibilityLevel          = $siteDetails.CompatibilityLevel
                    WebsCount                   = $siteDetails.WebsCount
                    IsHubSite                   = $siteDetails.IsHubSite
                    HubSiteId                   = $siteDetails.HubSiteId
                    SensitivityLabel            = $siteDetails.SensitivityLabel
                    CreatedDate                 = $site.CreatedDate
                }
                
                $usageData += $siteInfo
            }
            catch {
                Write-Warning "Error processing site $($site.Url): $_"
            }
        }
        
        Write-Progress -Activity "Processing Sites" -Completed
        Write-Host "Retrieved usage data for $($usageData.Count) sites." -ForegroundColor Green
        
        return $usageData
    }
    catch {
        Write-Error "Error retrieving data from SharePoint Online: $_"
        throw
    }
    finally {
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
}

# Helper to normalize a URL so SPO and Graph values match reliably.
# Decodes percent-encoded characters, strips trailing slashes, query strings,
# and fragments, and lowercases the result.
function Normalize-SiteUrl {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $null }
    $n = [System.Uri]::UnescapeDataString($Url.Trim())
    $n = $n.TrimEnd('/')
    $qi = $n.IndexOf('?'); if ($qi -ge 0) { $n = $n.Substring(0, $qi) }
    $fi = $n.IndexOf('#'); if ($fi -ge 0) { $n = $n.Substring(0, $fi) }
    return $n.ToLowerInvariant()
}

# Detect whether Graph report data has been obfuscated by the M365 privacy
# setting "Conceal user, group, and site names in all reports".  Obfuscated
# rows have zeroed-out SiteIds, empty/null Site URLs, and hashed owner names.
function Test-GraphDataObfuscated {
    param([array]$ReportData)
    if (-not $ReportData -or $ReportData.Count -eq 0) { return $false }
    $sampleSize = [Math]::Min(5, $ReportData.Count)
    for ($i = 0; $i -lt $sampleSize; $i++) {
        $entry = $ReportData[$i]
        if ([string]::IsNullOrWhiteSpace($entry.SiteUrl) -or
            $entry.SiteId -eq '00000000-0000-0000-0000-000000000000' -or
            ($entry.OwnerDisplayName -match '^[A-Fa-f0-9]{32}$')) {
            return $true
        }
    }
    return $false
}

# Check the tenant admin report-privacy setting using the Graph PowerShell SDK
# cmdlets (Get-MgAdminReportSetting / Update-MgAdminReportSetting) and, if
# concealment is enabled, attempt to disable it so that future reports contain
# real identifiers.  Returns a hashtable:
#   WasEnabled - $true/$false/$null (null = could not read the setting)
#   Fixed      - $true if the setting is now disabled
function Resolve-GraphReportPrivacy {
    try {
        # Use the typed Graph SDK cmdlet — more reliable than raw Invoke-MgGraphRequest.
        $reportSetting = Get-MgAdminReportSetting -ErrorAction Stop
        $concealed = $reportSetting.DisplayConcealedNames

        Write-Host "Current DisplayConcealedNames value: $concealed" -ForegroundColor Cyan

        if ($concealed -eq $true) {
            Write-Warning "Report privacy setting 'Conceal user, group, and site names in all reports' is ENABLED in your tenant (DisplayConcealedNames = True)."
            Write-Host "Attempting to disable the concealment setting via Update-MgAdminReportSetting..." -ForegroundColor Yellow
            try {
                Update-MgAdminReportSetting -DisplayConcealedNames:$false -ErrorAction Stop
                Write-Host "Concealment setting has been disabled (DisplayConcealedNames set to False)." -ForegroundColor Green
                Write-Host "Note: report data may take up to 48 hours to reflect this change." -ForegroundColor Yellow
                Write-Host "Re-run this script after the propagation period for fully de-obfuscated data." -ForegroundColor Yellow
                return @{ WasEnabled = $true; Fixed = $true }
            }
            catch {
                Write-Warning "Could not disable the concealment setting (requires ReportSettings.ReadWrite.All permission): $_"
                return @{ WasEnabled = $true; Fixed = $false }
            }
        }
        else {
            Write-Host "Report concealment setting is already disabled (DisplayConcealedNames = False)." -ForegroundColor Green
            return @{ WasEnabled = $false; Fixed = $true }
        }
    }
    catch {
        Write-Warning "Could not read report privacy settings via Get-MgAdminReportSetting (requires ReportSettings.Read.All permission): $_"
        return @{ WasEnabled = $null; Fixed = $false }
    }
}

# Retrieve per-site analytics and drive information via the Microsoft Graph
# Sites API.  These endpoints are NOT subject to the report-privacy concealment
# setting, so they always return real data.  For each site the function queries:
#   /sites/{id}/analytics/lastSevenDays  → access.actionCount (page views)
#   /sites/{id}/drive                    → quota/driveItem info
# Sites for which the call fails (permissions, missing drive, etc.) are silently
# skipped.  Returns a hashtable mapping webUrl → enrichment object.
function Get-PerSiteAnalyticsViaGraph {
    param([array]$Sites)

    $analytics = @{}
    $total = $Sites.Count
    $counter = 0

    Write-Host "Retrieving per-site analytics for $total sites (this may take a moment)..." -ForegroundColor Cyan

    foreach ($site in $Sites) {
        $counter++
        $siteId  = $site.id
        $webUrl  = $site.webUrl
        if (-not $siteId) { continue }

        Write-Progress -Activity "Retrieving per-site analytics" `
            -Status "Processing $counter of $total" `
            -PercentComplete (($counter / $total) * 100)

        $info = [PSCustomObject]@{
            PageViewCount    = $null
            FileCount        = $null
            StorageUsedBytes = $null
            LastActivityDate = $null
        }

        # --- Site analytics (page views) ---
        try {
            $resp = Invoke-MgGraphRequest -Method GET `
                -Uri "/v1.0/sites/$siteId/analytics/lastSevenDays" `
                -ErrorAction Stop
            if ($resp.access) {
                $info.PageViewCount = $resp.access.actionCount
            }
        }
        catch {
            # Analytics not available for this site — continue silently
        }

        # --- Drive info (file count + storage) ---
        try {
            $driveResp = Invoke-MgGraphRequest -Method GET `
                -Uri "/v1.0/sites/$siteId/drive?`$select=quota,lastModifiedDateTime" `
                -ErrorAction Stop
            if ($driveResp.quota) {
                $info.StorageUsedBytes = $driveResp.quota.used
            }
            if ($driveResp.lastModifiedDateTime) {
                $info.LastActivityDate = $driveResp.lastModifiedDateTime
            }
        }
        catch {
            # Drive not available for this site — continue silently
        }

        # --- File count via root children count (closest approximation to the
        #     report's FileCount; includes both files and sub-folders at root) ---
        try {
            $rootResp = Invoke-MgGraphRequest -Method GET `
                -Uri "/v1.0/sites/$siteId/drive/root?`$select=folder" `
                -ErrorAction Stop
            if ($rootResp.folder) {
                $info.FileCount = $rootResp.folder.childCount
            }
        }
        catch {
            # Root folder not available — continue silently
        }

        if ($webUrl) {
            $analytics[$webUrl] = $info
        }
    }

    Write-Progress -Activity "Retrieving per-site analytics" -Completed
    $populated = ($analytics.Values | Where-Object { $null -ne $_.PageViewCount }).Count
    Write-Host "Retrieved analytics for $populated of $total sites." -ForegroundColor Green

    return $analytics
}

# Retrieve real site metadata via the Microsoft Graph Sites API.  This endpoint
# is not subject to the report-privacy concealment setting, so it always returns
# real site IDs, URLs, and display names.
function Get-SiteMetadataViaGraph {
    Write-Host "Retrieving real site metadata from Microsoft Graph Sites API..." -ForegroundColor Cyan

    $allSites = @()

    # Try getAllSites first (works with Sites.Read.All application permission)
    $uri = '/v1.0/sites/getAllSites?$select=id,displayName,webUrl,createdDateTime&$top=999'
    try {
        while ($uri) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($response.value) { $allSites += $response.value }
            $uri = $response.'@odata.nextLink'
        }
        Write-Host "Retrieved metadata for $($allSites.Count) sites from Sites API." -ForegroundColor Green
        return $allSites
    }
    catch {
        Write-Warning "getAllSites endpoint failed, trying search-based approach: $_"
    }

    # Fallback: search-based approach (works with delegated Sites.Read.All)
    $uri = "/v1.0/sites?search=*&`$select=id,displayName,webUrl,createdDateTime&`$top=999"
    try {
        while ($uri) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($response.value) { $allSites += $response.value }
            $uri = $response.'@odata.nextLink'
        }
        Write-Host "Retrieved metadata for $($allSites.Count) sites from Sites API (search)." -ForegroundColor Green
        return $allSites
    }
    catch {
        Write-Warning "Sites search API also failed: $_"
        return @()
    }
}

# Function to combine Graph and SPO data for best-of-both-worlds reporting.
#
# Strategy: start from the Graph report (authoritative list with SiteIds and
# usage metrics for every site), then enrich each row with friendly metadata
# (Title, Owner, Template, etc.) from SPO.  This avoids the duplicate-row
# problem that occurred when the old approach started from SPO and appended
# unmatched Graph sites at the end.
#
# Matching order for each Graph site:
#   1. URL match — normalised SiteUrl (works when Graph URLs are populated).
#   2. SiteId match — resolve SPO URLs to Graph compound IDs and compare the
#      simple GUID portion (works even when Graph URLs are blank/obfuscated).
function Get-UsageReportsCombined {
    param(
        [string]$TenantName
    )

    Write-Host "Running combined mode: starting from Graph usage data, enriching with SPO site details..." -ForegroundColor Cyan

    # Step 1: Get Graph data first — this is the authoritative list of sites
    # with SiteIds and all usage metrics (page views, file counts, etc.).
    $graphData = Get-UsageReportsViaGraph -TenantName $TenantName

    Write-Host "Graph returned $($graphData.Count) site usage records." -ForegroundColor Cyan

    # Step 2: Get SPO data (friendly names, titles, owners, storage in MB).
    $spoData = Get-UsageReportsViaSPO -TenantName $TenantName

    Write-Host "SPO returned $($spoData.Count) site records." -ForegroundColor Cyan

    # Step 3: Build SPO lookup by normalised URL.
    $spoUrlLookup = @{}
    foreach ($spoSite in $spoData) {
        $key = Normalize-SiteUrl -Url $spoSite.SiteUrl
        if ($key) { $spoUrlLookup[$key] = $spoSite }
    }

    # Step 4: Build SPO lookup by SiteId (resolve each SPO URL to its Graph
    # SiteId via the /sites/{hostname}:/{path} endpoint).  This is needed when
    # Graph report URLs are blank due to the concealment setting.
    $spoIdLookup = @{}
    $emptyGuid = '00000000-0000-0000-0000-000000000000'

    # Check whether we even have usable Graph URLs for URL matching
    $graphUrlCount = ($graphData | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SiteUrl) }).Count
    $needSiteIdLookup = $graphUrlCount -lt ($graphData.Count * 0.5)

    if ($needSiteIdLookup -and $spoData.Count -gt 0) {
        Write-Host "Graph report URLs are mostly blank. Resolving SPO URLs to Graph SiteIds for matching..." -ForegroundColor Yellow
        $resolveCounter = 0
        $resolveTotal = $spoData.Count

        foreach ($spoSite in $spoData) {
            $resolveCounter++
            Write-Progress -Activity "Resolving SPO site IDs via Graph" `
                -Status "Processing $resolveCounter of $resolveTotal" `
                -PercentComplete (($resolveCounter / $resolveTotal) * 100)

            try {
                $spoUrl = [System.Uri]$spoSite.SiteUrl
                $hostname = $spoUrl.Host
                $path = $spoUrl.AbsolutePath.TrimEnd('/')
                $graphUri = if ($path -and $path -ne '/') {
                    "/v1.0/sites/${hostname}:${path}?`$select=id"
                } else {
                    "/v1.0/sites/${hostname}?`$select=id"
                }
                $resp = Invoke-MgGraphRequest -Method GET -Uri $graphUri -ErrorAction Stop
                if ($resp.id) {
                    # Extract the simple GUID from the compound id (hostname,siteGuid,webGuid)
                    $parts = $resp.id -split ','
                    if ($parts.Count -ge 2) {
                        $spoIdLookup[$parts[1]] = $spoSite
                    }
                    # Also key by the last segment for safety
                    $spoIdLookup[$parts[-1]] = $spoSite
                }
            }
            catch {
                # Could not resolve this SPO site — skip silently
            }
        }
        Write-Progress -Activity "Resolving SPO site IDs via Graph" -Completed
        Write-Host "Resolved $($spoIdLookup.Count) SPO site-to-SiteId mappings." -ForegroundColor Green
    }

    # Step 5: Loop through every Graph site, find its SPO match, and build the
    # combined output row.  Graph is the base — every Graph row gets one output
    # row, enriched with SPO metadata when a match is found.
    #
    # For unmatched sites with valid SiteIds but blank URLs, we call
    # Get-MgSite -SiteId to resolve displayName and webUrl directly from the
    # Graph Sites API — these endpoints are NOT affected by report concealment.
    $combinedData = @()
    $matchedCount = 0
    $resolvedViaSitesApi = 0
    $total = $graphData.Count
    $counter = 0

    Write-Host "Merging $total Graph sites with SPO metadata..." -ForegroundColor Cyan

    foreach ($graphSite in $graphData) {
        $counter++
        $spoSite = $null
        $resolvedDisplayName = $null
        $resolvedWebUrl = $null

        # Try 1: match by normalised URL
        $graphKey = Normalize-SiteUrl -Url $graphSite.SiteUrl
        if ($graphKey) {
            $spoSite = $spoUrlLookup[$graphKey]
        }

        # Try 2: match by SiteId (via the pre-resolved SPO→SiteId lookup)
        if (-not $spoSite -and $graphSite.SiteId -and $graphSite.SiteId -ne $emptyGuid) {
            $simpleId = ($graphSite.SiteId -split ',')[-1]
            $spoSite = $spoIdLookup[$simpleId]
        }

        # Try 3: resolve SiteId directly via Get-MgSite to get displayName & webUrl,
        # then try matching the resolved webUrl back to SPO for full metadata.
        if (-not $spoSite -and $graphSite.SiteId -and $graphSite.SiteId -ne $emptyGuid) {
            try {
                $mgSite = Get-MgSite -SiteId $graphSite.SiteId -Property "id,displayName,webUrl" -ErrorAction Stop
                if ($mgSite) {
                    $resolvedDisplayName = $mgSite.DisplayName
                    $resolvedWebUrl = $mgSite.WebUrl
                    $resolvedViaSitesApi++

                    # Try to match the resolved webUrl back to SPO for full metadata
                    $resolvedKey = Normalize-SiteUrl -Url $resolvedWebUrl
                    if ($resolvedKey) {
                        $spoSite = $spoUrlLookup[$resolvedKey]
                    }
                }
            }
            catch {
                # Get-MgSite failed for this SiteId — continue silently
            }
        }

        # Build combined row: Graph metrics + SPO friendly metadata (or resolved metadata)
        $siteUrl = if ($spoSite) { $spoSite.SiteUrl } elseif ($resolvedWebUrl) { $resolvedWebUrl } elseif ($graphSite.SiteUrl) { $graphSite.SiteUrl } else { '' }
        $title = if ($spoSite) { $spoSite.Title } elseif ($resolvedDisplayName) { $resolvedDisplayName } else { '' }
        $owner = if ($spoSite) { $spoSite.Owner } elseif ($resolvedDisplayName) { $graphSite.OwnerDisplayName } else { $graphSite.OwnerDisplayName }

        $combined = [PSCustomObject]@{
            SiteUrl                 = $siteUrl
            SiteId                  = $graphSite.SiteId
            Title                   = $title
            Owner                   = $owner
            OwnerPrincipalName      = $graphSite.OwnerPrincipalName
            Template                = if ($spoSite) { $spoSite.Template } else { '' }
            StorageUsedMB           = if ($spoSite) { $spoSite.StorageUsedMB } else { if ($graphSite.StorageUsedInBytes) { [math]::Round($graphSite.StorageUsedInBytes / 1MB, 2) } else { '' } }
            StorageQuotaMB          = if ($spoSite) { $spoSite.StorageQuotaMB } else { if ($graphSite.StorageAllocatedInBytes) { [math]::Round($graphSite.StorageAllocatedInBytes / 1MB, 2) } else { '' } }
            StorageUsedPercentage   = if ($spoSite) { $spoSite.StorageUsedPercentage } else { '' }
            LastContentModifiedDate = if ($spoSite) { $spoSite.LastContentModifiedDate } else { '' }
            LastActivityDate        = $graphSite.LastActivityDate
            FileCount               = $graphSite.FileCount
            ActiveFileCount         = $graphSite.ActiveFileCount
            PageViewCount           = $graphSite.PageViewCount
            VisitedPageCount        = $graphSite.VisitedPageCount
            SharingCapability       = if ($spoSite) { $spoSite.SharingCapability } else { '' }
            LockState               = if ($spoSite) { $spoSite.LockState } else { '' }
            IsHubSite               = if ($spoSite) { $spoSite.IsHubSite } else { '' }
            HubSiteId               = if ($spoSite) { $spoSite.HubSiteId } else { '' }
            SensitivityLabel        = if ($spoSite) { $spoSite.SensitivityLabel } else { '' }
            RootWebTemplate         = $graphSite.RootWebTemplate
            IsDeleted               = $graphSite.IsDeleted
            CreatedDate             = if ($spoSite) { $spoSite.CreatedDate } else { '' }
            ReportRefreshDate       = $graphSite.ReportRefreshDate
        }
        $combinedData += $combined

        if ($spoSite) {
            $matchedCount++
            # Remove from URL lookup so we can track SPO-only sites later
            if ($graphKey) { $spoUrlLookup.Remove($graphKey) }
        }
    }

    Write-Host "Combined report: $total Graph sites, $matchedCount enriched with SPO metadata, $resolvedViaSitesApi resolved via Get-MgSite, $($spoUrlLookup.Count) SPO-only sites not in Graph report." -ForegroundColor Green

    if ($matchedCount -eq 0 -and $spoData.Count -gt 0 -and $graphData.Count -gt 0) {
        Write-Warning "No Graph sites could be matched to SPO data. Title/Owner columns will be empty."
        Write-Warning "This can occur when the report-privacy concealment setting was recently changed. Re-run after 48 hours."
    }

    return $combinedData
}

# Main script execution
try {
    Write-Host "`n=== SharePoint Site Usage Report Generator ===" -ForegroundColor Cyan
    Write-Host "Tenant: $TenantName" -ForegroundColor Cyan
    Write-Host "Authentication Method: $AuthMethod" -ForegroundColor Cyan
    Write-Host "Mode: $(if ($UseCombined) { 'Combined (SPO + Graph)' } elseif ($UseGraphAPI) { 'Graph API' } else { 'SPO Management Shell' })`n" -ForegroundColor Cyan
    
    # Set default output path if not provided
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $OutputPath = Join-Path -Path $PSScriptRoot -ChildPath "SPO_SiteUsage_$($TenantName)_$timestamp.csv"
    }
    
    # Check and install required modules
    if ($UseCombined) {
        Write-Host "Checking for required modules (SPO + Graph)..." -ForegroundColor Cyan
        $modulesInstalled = Install-RequiredModules -ModuleNames @('Microsoft.Online.SharePoint.PowerShell', 'Microsoft.Graph.Reports', 'Microsoft.Graph.Authentication')
        if (-not $modulesInstalled) {
            throw "Failed to install required modules for combined mode."
        }

        # Get usage data via combined SPO + Graph approach
        $usageData = Get-UsageReportsCombined -TenantName $TenantName
    }
    elseif ($UseGraphAPI) {
        Write-Host "Checking for required Microsoft Graph modules..." -ForegroundColor Cyan
        $modulesInstalled = Install-RequiredModules -ModuleNames @('Microsoft.Graph.Reports', 'Microsoft.Graph.Authentication')
        if (-not $modulesInstalled) {
            throw "Failed to install required Microsoft Graph modules."
        }
        
        # Get usage data via Graph API
        $usageData = Get-UsageReportsViaGraph -TenantName $TenantName
    }
    else {
        Write-Host "Checking for required SharePoint Online Management Shell module..." -ForegroundColor Cyan
        $modulesInstalled = Install-RequiredModules -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')
        if (-not $modulesInstalled) {
            throw "Failed to install required SharePoint Online Management Shell module."
        }
        
        # Get usage data via SPO Management Shell
        $usageData = Get-UsageReportsViaSPO -TenantName $TenantName
    }
    
    # Export to CSV
    if ($usageData -and $usageData.Count -gt 0) {
        Write-Host "`nExporting data to CSV: $OutputPath" -ForegroundColor Cyan

        # Ensure the output directory exists
        $outputDir = Split-Path -Path $OutputPath -Parent
        if ($outputDir -and -not (Test-Path -Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }

        # Attempt to export with retry logic in case the file is locked
        $maxRetries = 3
        $retryDelay = 2
        $exported = $false
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                $usageData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
                $exported = $true
                break
            }
            catch [System.IO.IOException] {
                if ($attempt -lt $maxRetries) {
                    Write-Warning "File '$OutputPath' is in use. Retrying in $retryDelay seconds... (attempt $attempt of $maxRetries)"
                    Start-Sleep -Seconds $retryDelay
                    $retryDelay *= 2
                }
            }
        }

        if (-not $exported) {
            # All retries failed — write to a fallback file with a timestamp suffix
            $fallbackName = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath) +
                "_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv"
            $fallbackDir = Split-Path -Path $OutputPath -Parent
            $fallbackPath = if ($fallbackDir) { Join-Path -Path $fallbackDir -ChildPath $fallbackName } else { $fallbackName }
            Write-Warning "Could not write to '$OutputPath' because it is locked by another process (e.g., Excel)."
            Write-Warning "Saving report to fallback location: $fallbackPath"
            $usageData | Export-Csv -Path $fallbackPath -NoTypeInformation -Encoding UTF8
            $OutputPath = $fallbackPath
        }

        Write-Host "Export completed successfully!" -ForegroundColor Green
        Write-Host "Total sites exported: $($usageData.Count)" -ForegroundColor Green
        Write-Host "File location: $OutputPath" -ForegroundColor Green
        
        # Display summary statistics
        if ($UseCombined) {
            $totalStorageMB = ($usageData | Measure-Object -Property StorageUsedMB -Sum).Sum
            $totalFiles = ($usageData | Where-Object { $_.FileCount -ne '' } | Measure-Object -Property FileCount -Sum).Sum
            $totalPageViews = ($usageData | Where-Object { $_.PageViewCount -ne '' } | Measure-Object -Property PageViewCount -Sum).Sum
            
            Write-Host "`n=== Summary Statistics ===" -ForegroundColor Cyan
            Write-Host "Total Storage Used: $([math]::Round($totalStorageMB / 1024, 2)) GB" -ForegroundColor White
            Write-Host "Total Files: $totalFiles" -ForegroundColor White
            Write-Host "Total Page Views (last 7 days): $totalPageViews" -ForegroundColor White
        }
        elseif ($UseGraphAPI) {
            $totalStorage = ($usageData | Measure-Object -Property StorageUsedInBytes -Sum).Sum
            $totalFiles = ($usageData | Measure-Object -Property FileCount -Sum).Sum
            $totalPageViews = ($usageData | Measure-Object -Property PageViewCount -Sum).Sum
            
            Write-Host "`n=== Summary Statistics ===" -ForegroundColor Cyan
            Write-Host "Total Storage Used: $([math]::Round($totalStorage / 1GB, 2)) GB" -ForegroundColor White
            Write-Host "Total Files: $totalFiles" -ForegroundColor White
            Write-Host "Total Page Views: $totalPageViews" -ForegroundColor White
        }
        else {
            $totalStorageMB = ($usageData | Measure-Object -Property StorageUsedMB -Sum).Sum
            Write-Host "`n=== Summary Statistics ===" -ForegroundColor Cyan
            Write-Host "Total Storage Used: $([math]::Round($totalStorageMB / 1024, 2)) GB" -ForegroundColor White
        }
    }
    else {
        Write-Warning "No usage data retrieved. No CSV file created."
    }
}
catch {
    Write-Error "An error occurred: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}
