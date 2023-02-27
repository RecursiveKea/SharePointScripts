#OCR Script is in seperate file
Param (
     [Parameter(Mandatory=$true)][String]$ModulePath
)
. ($ModulePath + "\PsOcr.ps1");


#Wrapper function to get stored credentials
function Get-SPOnlineHelperAppCredential
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$StoredCredentialName
    );
    
    $Cred = (Get-PnPStoredCredential -Name $StoredCredentialName);
    $ClientId = ($Cred.UserName);
    $ClientSecret = ([System.Net.NetworkCredential]::new("", $Cred.Password).Password);    
    return (New-Object PSCustomObject -Property @{ 
        ClientID = $ClientId; 
        ClientSecret = $ClientSecret;
    });
}


#Generic Rest API caller for SharePoint
function Call-SPOnlineHelperRestApiMethod
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$RestApiUrl
    );

    $theResult = [System.Collections.ArrayList]::new();
    $theRestApiCallCount = 0;
    do
    {
        $RestApiResult = (Invoke-PnPSPRestMethod -Url $RestApiUrl -ErrorAction Stop);
        foreach ($currVal in $RestApiResult.value) {
            $theResult.add($currVal) | Out-Null;
        }        
        $RestApiUrl = ($RestApiResult.'odata.nextLink');

        $theRestApiCallCount++;
        if (($theRestApiCallCount % 10) -eq 0) {
            Start-Sleep -Seconds 120;
        }
    }
    while (-not [String]::IsNullOrEmpty($RestApiUrl));


    return $theResult;
}


#Iterate through a SharePoint folder and get all the files contained within along with the selected fields
function Recurse-SPOnlineHelperFolderIteratorForFiles
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$FolderPath
        ,[Parameter(Mandatory=$true)][String]$SiteUrl
        ,[Parameter(Mandatory=$true)][String]$RelativeUrl
        ,[Parameter(Mandatory=$true)][String]$SelectFields
    );

    $theResult = [System.Collections.ArrayList]::new();
    $theFolder = (Get-PnPFolder -Url $FolderPath -Includes Folders -ErrorAction Stop);

    $theFilesRestApiUrl = ($SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + $FolderPath + "')/Files?`$Select=" + $SelectFields + "&`$top=2000");
    $theFiles = (Call-SPOnlineHelperRestApiMethod -RestApiUrl $theFilesRestApiUrl);
    
    foreach ($currFile in $theFiles) {
        $theResult.Add($currFile) | Out-Null;
    }

    
    foreach ($currFolder in $theFolder.Folders) {
        if ($currFolder.Name -ne "Forms") {
            
            $SubFolderParams = @{
                FolderPath = ($currFolder.ServerRelativeUrl);
                SiteUrl = $SiteUrl
                RelativeUrl = $RelativeUrl;
                SelectFields = $SelectFields;
            };

            $subFolderFiles = (Recurse-SPOnlineHelperFolderIteratorForFiles @SubFolderParams);
            foreach ($currFile in $subFolderFiles) {
                $theResult.Add($currFile) | Out-Null;
            }
        }
    }

    return $theResult;
}



#OCR files and get all the text content from it in one string variable
function Get-AllTextFromImage
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$FilePath
    );

    $ImageText = (Convert-PsoImageToText -Path $FilePath);
    $ImageTextJoined = "";
    foreach ($currLine in $ImageText) {
        $ImageTextJoined += ($currLine.Words);
    }

    return $ImageTextJoined;
}


#Detect potentially corrupted files
function Detect-SPOnlinePotentiallyCorruptedFiles
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$FolderPath        #SharePoint Folder Path to iterate over and detect potentially corrupted files
        ,[Parameter(Mandatory=$true)][String]$TempExportFolder  #Temporary location to store thumbnails of emails to OCR
    );

    #Get the files and their properties to assist in getting their thumbnails and export results
    $Web = (Get-PnPWeb -ErrorAction Stop);
    $FolderParams = @{
        SiteUrl = ($Web.Url);
        RelativeUrl = ($web.ServerRelativeUrl);
        FolderPath = $FolderPath;
        SelectFields = "VroomDriveID,VroomItemID,ServerRelativeUrl,Name";
    };

    $LibraryFiles = (Recurse-SPOnlineHelperFolderIteratorForFiles @FolderParams);

    <# #SupportedExtensions: /_layouts/15/getpreview.ashx?action=supportedtypes
    $SupportedThumbnailExtensions = @{}
    $SupportedThumbnailExtensions[".docm"];
    $SupportedThumbnailExtensions[".docx"];
    $SupportedThumbnailExtensions[".dotx"];
    $SupportedThumbnailExtensions[".dotm"];
    $SupportedThumbnailExtensions[".bmp"];
    $SupportedThumbnailExtensions[".jpg"];
    $SupportedThumbnailExtensions[".jpeg"];
    $SupportedThumbnailExtensions[".tiff"];
    $SupportedThumbnailExtensions[".tif"];
    $SupportedThumbnailExtensions[".png"];
    $SupportedThumbnailExtensions[".gif"];
    $SupportedThumbnailExtensions[".emf"];
    $SupportedThumbnailExtensions[".wmf"];
    $SupportedThumbnailExtensions[".psd"];
    $SupportedThumbnailExtensions[".svg"];
    $SupportedThumbnailExtensions[".ai"];
    $SupportedThumbnailExtensions[".eps"];
    $SupportedThumbnailExtensions[".pdf"];
    $SupportedThumbnailExtensions[".pptm"];
    $SupportedThumbnailExtensions[".pptx"];
    $SupportedThumbnailExtensions[".potm"];
    $SupportedThumbnailExtensions[".potx"];
    $SupportedThumbnailExtensions[".ppsm"];
    $SupportedThumbnailExtensions[".ppsx"];
    $SupportedThumbnailExtensions[".xlsm"];
    $SupportedThumbnailExtensions[".xlsx"];
    $SupportedThumbnailExtensions[".aspx"];
    #>

    #Iterate through all the files checking if the thumbnail is valid
    $theResult = [System.Collections.ArrayList]::new();
    foreach ($currFile in $LibraryFiles) {   
            
        try {
            #Needed for downloading the image data / file
            $WebClient = [System.Net.WebClient]::new();

            #Get the extension of the file to check
            $Extension = [System.IO.Path]::GetExtension($currFile.Name);
            $fetchResult = "-";

            #Get the graph url for the retrieving the file thumbnail image
            $GraphThumbnailRequestUrl = "";
            if ($Extension -eq ".msg") {
                #Need the large image for OCR'ing
                $GraphThumbnailRequestUrl = ("drives/" + ($currFile.VroomDriveID) + "/items/" + ($currFile.VroomItemID) + "/thumbnails/0/large");
            } else {
                $GraphThumbnailRequestUrl = ("drives/" + ($currFile.VroomDriveID) + "/items/" + ($currFile.VroomItemID) + "/thumbnails/0/small");
            }

            #Call the graph method to get the image url for the current file
            $DocumentThumbnailUrl = (Invoke-PnPGraphMethod -Url $GraphThumbnailRequestUrl -ErrorAction Stop);
            $WebClient = [System.Net.WebClient]::new();

            #Emails get OCR'd
            if ($Extension -eq ".msg") {
                
                #SharePoint will generate a thumbnail with the January 1 that seem to be corrupted so download the file then OCR it and check if the content is January 1, 0001 and is empty
                $TempEmailThumbnailExportPath = ($TempExportFolder + "\TempEmailThumbnail.png");
                $WebClient.DownloadFile($DocumentThumbnailUrl.url, $TempEmailThumbnailExportPath);
                $ImageText = (Get-AllTextFromImage -FilePath $TempEmailThumbnailExportPath);

                #After getting the image text delete the thumbnail
                Remove-Item -LiteralPath $TempEmailThumbnailExportPath -Force;

                #If the file meets the criteria of being deemed empty throw an exception indicating it is likely corrupted
                if ($ImageText -eq "From:Sent on:Subject:Monday, January 1, AM") {
                    throw [System.Exception]::new("Email is likely corrupted");
                }

            } else {
                #If not an email download the data - if this fails SharePoint couln't generate an image for the file
                $ImageData = $WebClient.DownloadData($DocumentThumbnailUrl.url);
            }
            
            #If the script gets to here then the check was a success
            $fetchResult = "Success";
        } catch {
            
            #Otherwise it failed
            Write-Host "Failed" -ForegroundColor Red;
            Write-Host $currFile.Name;

            $fetchResult = "Error";
        }

        #If the thumbnail result is an error add it to the result set to be returned by the function
        if ($fetchResult -eq "Error") {
            $theResult.add(
                (New-Object PSCUstomObject -Property @{
                    Extension = $Extension;
                    Name = ($currFile.Name);
                    Result = $fetchResult;
                    ServerRelativeUrl = ($currFile.ServerRelativeUrl);
                })
            ) | Out-Null;
        }
    }

    #Return the collection of possibly corrupted files
    return $theResult;
}



#Iterate libraries in currently connected site to output the possibly corrupted files
function Export-SPSitePotentiallyCorruptedFiles
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$ExportPathAndName
        ,[Parameter(Mandatory=$true)][String]$TempExportFolder
    );

    $Web = (Get-PnPWeb -ErrorAction Stop); 
    $Lists = (Get-PnPList -Includes IsSystemList,RootFolder,BaseType -ErrorAction Stop);    

    foreach ($currList in $Lists) {
        if ((-not $currList.IsSystemList) -and ($currList.BaseType -eq "DocumentLibrary")) {
            
            $RootFolder = ($currList.RootFolder.ServerRelativeUrl);
            if ($Web.ServerRelativeUrl -ne "/") {
                $RootFolder = ($RootFolder.Substring($RootFolder.IndexOf($Web.ServerRelativeUrl) + $Web.ServerRelativeUrl.Length + 1))
            }
            
            if ($RootFolder -ne "SitePages") {
                Write-Host ("`t" + "-" + $RootFolder);
                $PotentiallyCorruptedFiles = (Detect-SPOnlinePotentiallyCorruptedFiles -FolderPath $RootFolder -TempExportFolder $TempExportFolder);
                $PotentiallyCorruptedFiles | Select @{N="SiteUrl";E={ $Web.Url}}, Extension, Name, Result, ServerRelativeUrl | Export-Csv -LiteralPath $ExportPathAndName -Append -NoTypeInformation;
            }
        }
    }
}


#Crawl the tenant of all libaries to detect potentially corrupted files
function Crawl-SPTenantForPotentiallyCorruptedFiles
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$TenantUrl
        ,[Parameter(Mandatory=$true)]$PnPCredentialParams
        ,[Parameter(Mandatory=$true)][String]$ExportPathAndName
        ,[Parameter(Mandatory=$true)][String]$TempExportFolder
        ,[Parameter(Mandatory=$false)][int]$SecondsToPauseBetweenSites = 0
    );
    
    Connect-PnPOnline -Url $TenantUrl @PnPCredentialParams;
    $AllSites = (Get-PnPTenantSite -ErrorAction Stop);    

    try {
        
        foreach ($currSite in $AllSites) {
        
            if (-not $currSite.Url.toLower().Contains("-my.sharepoint.com")) {
                
                Write-Host ($currSite.Url) -ForegroundColor Green;
                Connect-PnPOnline -Url ($currSite.Url) @PnPCredentialParams;
                Export-SPSitePotentiallyCorruptedFiles -ExportPathAndName $ExportPathAndName -TempExportFolder $TempExportFolder;

                $SubWebs = (Get-PnPSubWeb -Recurse -ErrorAction Stop);
                foreach ($currSubSite in $SubWebs) {

                    Write-Host ($currSubSite.Url) -ForegroundColor Green;
                    Connect-PnPOnline -Url ($currSubSite.Url) @PnPCredentialParams;
                    Export-SPSitePotentiallyCorruptedFiles -ExportPathAndName $ExportPathAndName -TempExportFolder $TempExportFolder;
                }

                Start-Sleep -Seconds $SecondsToPauseBetweenSites;
            }
        }
    } catch {
        Write-Host "ERROR" -ForegroundColor Red;
        Write-Host $_;
    }

    Write-Host "";
}