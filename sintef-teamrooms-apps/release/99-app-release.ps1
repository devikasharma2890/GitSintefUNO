<#
Title: 				99-app-Release.ps1
Description:		Release script to use in Azure DevOps
Created by:			Tom Solem
#>

[CmdletBinding()]
param(
    [Parameter(Position=1)]
    [string]$AppSecret
    <# 
    AppSecret need to be a paramenter. 
    Use a secret value in release defintion and add -AppSecret $(<name of variables>) in the argument list in the powershell setting #>
)
Install-PackageProvider -Name NuGet -Force -Scope "CurrentUser"
Install-Module SharePointPnPPowerShellOnline -Scope "CurrentUser" -Verbose -Force

# setting path use for release
$scriptDir = Get-Item -Filter Folder -Path "."
Write-Host "Script dir: $($scriptDir.FullName)"
$source = (get-item $scriptDir -Filter Folder).parent
Write-Host "source path: $($source.FullName)"

# setting values from environment settings
$domain = $env:Domain 
$GlobalPrimaryOwner = $env:GlobalPrimaryOwner
$appId = $env:AppId
$appSecret = $AppSecret

# setting sharepoint url based on domain
$rootSite = "https://$($domain).sharepoint.com"
$appSite = $rootSite + "/sites/apps"
$adminSite = "https://$($domain)-admin.sharepoint.com"
Connect-PnPOnline -Url $appSite -AppId $appId -AppSecret $appSecret

Write-Host "Getting artifacts from the source folder"
$packageItems = Get-ChildItem -File -Filter "*.sppkg" -path $source.FullName -Recurse
Write-Host "Found $($packageItems.Count) items in the source folder"
Write-Host "`tUpload packages"
foreach($item in $packageItems){
    try
    {
        Add-PnPApp -Path $item.FullName -Publish -Overwrite
    }
    catch
    {
        Write-Host "Missing APP file: '$($item.FullName)' Check build and path." -ForegroundColor Red -BackgroundColor White
    }
}

Write-Host "Disconnect SharePoint Online"
Disconnect-PnPOnline