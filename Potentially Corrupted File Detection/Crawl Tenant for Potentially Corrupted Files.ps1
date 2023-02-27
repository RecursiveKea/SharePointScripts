$ModulePath = "C:\Users\Kris\Documents\GitHub\SharePointScripts\Potentially Corrupted File Detection";
. ($ModulePath + "\Module-SPOnlineDetectCorruptedFiles.ps1") -ModulePath $ModulePath;

$Tenant = "{TENANT}";

<# App Cred Cert Crawl Tenant
$AppCredName = "DevTenant";
$TenantUrl = ("https://" + $Tenant + ".sharepoint.com");
$AppCreds = (Get-SPOnlineHelperAppCredential -StoredCredentialName $AppCredName);
$PnPCredentialParams = @{
    ClientID = ($AppCreds.ClientID);
    Thumbprint = ($AppCreds.ClientSecret);
    Tenant = ($Tenant + ".onmicrosoft.com");
};

$CrawlParams = @{
    TenantUrl = $TenantUrl;
    PnPCredentialParams = $PnPCredentialParams;
    ExportPathAndName = "C:\TEMP\PotentiallyCorruptedFiles.csv";
    TempExportFolder = "C:\TEMP";
};

Crawl-SPTenantForPotentiallyCorruptedFiles @CrawlParams;
#>



<#Interactive Crawl Tenant
$PnPCredentialParams = @{ Interactive = $true; };
$CrawlParams = @{
    TenantUrl = $TenantUrl;
    PnPCredentialParams = $PnPCredentialParams;
    ExportPathAndName = "C:\TEMP\PotentiallyCorruptedFiles.csv";
    TempExportFolder = "C:\TEMP";
};

Crawl-SPTenantForPotentiallyCorruptedFiles @CrawlParams;
#>


<# Single Site Interactive
$SiteUrlToCheck = "{SiteUrl}";
$LibraryToCrawl = "{LibraryRootFolder}";
Connect-PnPOnline -Url $SiteUrlToCheck -Interactive -ErrorAction Stop;
Detect-SPOnlinePotentiallyCorruptedFiles -FolderPath $LibraryToCrawl -TempExportFolder "C:\TEMP";
#>