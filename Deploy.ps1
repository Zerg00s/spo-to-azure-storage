param (
    [string]$Path
)

Get-ChildItem -Recurse | Unblock-File
# Legacy PowerShell PnP Module is used because 
# the new one (PnP.PowerShell) requires Tenant Admin consent for the PnP Azure App Registration
# and it's not possible to use it in the context of a non-admin user
Import-Module (Get-ChildItem -Recurse -Filter "*.psd1").FullName -DisableNameChecking

$ErrorActionPreference = "Stop"
$host.UI.RawUI.WindowTitle = "Move SharePoint Libraries to Azure Storage"
Write-Host $Path -ForegroundColor Green
Set-location $Path
Get-ChildItem -Recurse | Unblock-File
. .\Misc\PS-Forms.ps1
Import-Module (Get-ChildItem -Recurse -Filter "*.psd1").FullName
Clear-Host
                                                                    
# Created by Denis Molodtsov, 2023

$inputs = @{
    TARGET_SITE_URL   = "https://contoso.sharepoint.com/sites/Wishes"
}


$inputs = Get-FormItemProperties -item $inputs -dialogTitle "Enter target site" -propertiesOrder @("TARGET_SITE_URL") 
$TARGET_SITE_URL = $inputs.TARGET_SITE_URL


try{
    Connect-PnPOnline -UseWebLogin -Url $TARGET_SITE_URL -WarningAction Ignore
    
    $lists = Get-PnPList 
    # TODO: Display all SharePoint Libraries and allow to select the ones to move

    # TODO: Create a function that moves a single library

    # Connect to Azure Blob Storage

    # Create a container if does not exist. Decide if you need to create one container
    # or one container per SharePoint site
    # or one container per SharePoint library
    # Upload


    $StorageAccountName = ""
    $StorageAccountKey = ""
    $context = New-AzStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey

    ###############################
    $Body = "Hello World!"
    $blobName = "Sample.txt"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $stream = New-Object -TypeName System.IO.MemoryStream -ArgumentList (,$bytes)
    $Container = Get-AzStorageContainer -Name "results" -Context $context
    $BlockBlob = $Container.CloudBlobContainer.GetBlockBlobReference($blobName)
    $BlockBlob.Metadata.Add("Author", "Denis Molodtsov")
    $BlockBlob.Metadata.Add("Title", "Whatever")
    $BlockBlob.Metadata.Add("Date", "2023/01/01")
    $BlockBlob.UploadFromStream($stream);
    $stream.Dispose()
    # add metadata to the blob
    

}
catch {
    Write-Host
    Write-Host $_.Exception -ForegroundColor Red
    Write-host Please, share this error with the administrator


    Write-Host Press any key...
    [System.Console]::ReadKey()
    Write-Host Press any key...
    [System.Console]::ReadKey()
    return
}
