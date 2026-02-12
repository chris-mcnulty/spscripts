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
        # Connect to Microsoft Graph
        Connect-MgGraph -Scopes "Reports.Read.All", "Sites.Read.All" -NoWelcome
        
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
        $sites = Get-Content -Path $tempFile -Raw | ConvertFrom-Csv
        
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

# Function to combine SPO and Graph data for best-of-both-worlds reporting
function Get-UsageReportsCombined {
    param(
        [string]$TenantName
    )

    Write-Host "Running combined mode: collecting friendly site info from SPO and activity data from Graph API..." -ForegroundColor Cyan

    # Step 1: Get SPO data (friendly names, titles, owners)
    $spoData = Get-UsageReportsViaSPO -TenantName $TenantName

    # Step 2: Get Graph data (page views, active files, activity dates)
    $graphData = Get-UsageReportsViaGraph -TenantName $TenantName

    # Step 3: Build a lookup table of Graph data keyed by SiteUrl
    $graphLookup = @{}
    foreach ($graphSite in $graphData) {
        if (-not [string]::IsNullOrWhiteSpace($graphSite.SiteUrl)) {
            $graphLookup[$graphSite.SiteUrl.TrimEnd('/')] = $graphSite
        }
    }

    Write-Host "Merging $($spoData.Count) SPO sites with $($graphLookup.Count) Graph usage records..." -ForegroundColor Cyan

    # Step 4: Merge SPO and Graph data
    $combinedData = @()
    foreach ($spoSite in $spoData) {
        $url = $spoSite.SiteUrl.TrimEnd('/')
        $graphSite = $graphLookup[$url]

        $combined = [PSCustomObject]@{
            SiteUrl                 = $spoSite.SiteUrl
            Title                   = $spoSite.Title
            Owner                   = $spoSite.Owner
            Template                = $spoSite.Template
            StorageUsedMB           = $spoSite.StorageUsedMB
            StorageQuotaMB          = $spoSite.StorageQuotaMB
            StorageUsedPercentage   = $spoSite.StorageUsedPercentage
            LastContentModifiedDate = $spoSite.LastContentModifiedDate
            LastActivityDate        = if ($graphSite) { $graphSite.LastActivityDate } else { '' }
            FileCount               = if ($graphSite) { $graphSite.FileCount } else { '' }
            ActiveFileCount         = if ($graphSite) { $graphSite.ActiveFileCount } else { '' }
            PageViewCount           = if ($graphSite) { $graphSite.PageViewCount } else { '' }
            VisitedPageCount        = if ($graphSite) { $graphSite.VisitedPageCount } else { '' }
            SharingCapability       = $spoSite.SharingCapability
            LockState               = $spoSite.LockState
            IsHubSite               = $spoSite.IsHubSite
            HubSiteId               = $spoSite.HubSiteId
            SensitivityLabel        = $spoSite.SensitivityLabel
            RootWebTemplate         = if ($graphSite) { $graphSite.RootWebTemplate } else { '' }
            CreatedDate             = $spoSite.CreatedDate
        }
        $combinedData += $combined

        # Remove matched entry so we can track unmatched Graph sites
        if ($graphSite) {
            $graphLookup.Remove($url)
        }
    }

    # Step 5: Append any Graph-only sites not found in SPO (e.g. deleted sites)
    foreach ($remaining in $graphLookup.Values) {
        $combined = [PSCustomObject]@{
            SiteUrl                 = $remaining.SiteUrl
            Title                   = ''
            Owner                   = $remaining.OwnerDisplayName
            Template                = ''
            StorageUsedMB           = [math]::Round($remaining.StorageUsedInBytes / 1MB, 2)
            StorageQuotaMB          = [math]::Round($remaining.StorageAllocatedInBytes / 1MB, 2)
            StorageUsedPercentage   = if ($remaining.StorageAllocatedInBytes -gt 0) { [math]::Round(($remaining.StorageUsedInBytes / $remaining.StorageAllocatedInBytes) * 100, 2) } else { 0 }
            LastContentModifiedDate = ''
            LastActivityDate        = $remaining.LastActivityDate
            FileCount               = $remaining.FileCount
            ActiveFileCount         = $remaining.ActiveFileCount
            PageViewCount           = $remaining.PageViewCount
            VisitedPageCount        = $remaining.VisitedPageCount
            SharingCapability       = ''
            LockState               = ''
            IsHubSite               = ''
            HubSiteId               = ''
            SensitivityLabel        = ''
            RootWebTemplate         = $remaining.RootWebTemplate
            CreatedDate             = ''
        }
        $combinedData += $combined
    }

    $matchedCount = $combinedData.Count - $graphLookup.Count
    Write-Host "Combined report: $($combinedData.Count) total sites ($matchedCount matched, $($graphLookup.Count) Graph-only)." -ForegroundColor Green
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
