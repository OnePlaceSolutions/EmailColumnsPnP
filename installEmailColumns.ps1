 <#
        This script prompts you for the site collection URL you wish to install the Email Columns to
        and then applies the email columns template to this site collection.
        This is compatible with both classic PnP cmdlets and the new PnP.PowerShell cmdlets for SharePoint Online (no App consent required)   
#>
$ErrorActionPreference = 'Stop'

Try {    
    Set-ExecutionPolicy Bypass -Scope Process 
    Write-Host "Checking for PnP module..."
    $installedPnPModules = Get-InstalledModule -Name "*PnP*"
    If($installedPnPModules.Name -eq "PnP.PowerShell") {
        [boolean]$PnP_PowerShell = $true
    }
    Else {
        [boolean]$PnP_PowerShell = $false
    }
    #Prompt for SharePoint URL     
    $SharePointUrl = Read-Host -Prompt 'Enter your SharePoint Site Collection URL to install OnePlace Solutions Email Columns to'
    
    #Connect to site collection
    If(($SharePointUrl -match ".sharepoint.com/") -or $PnP_PowerShell) {
        Write-Host "Enter SharePoint credentials (your email address for SharePoint Online):" -ForegroundColor Green  
        Connect-PnPOnline -Url $SharePointUrl -UseWebLogin -WarningAction Ignore
    }
    Else {
        Write-Host "Enter SharePoint credentials (domain\username for on-premises):" -ForegroundColor Green  
        Connect-PnPOnline -Url $SharePointUrl 
    }

    #Download xml provisioning template
    $WebClient = New-Object System.Net.WebClient   
    $Url = "https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml"    
    $Path = "$env:temp\email-columns.xml"

    Write-Host "Downloading provisioning xml template:" $Path -ForegroundColor Green 
    $WebClient.DownloadFile( $Url, $Path )   
    #Apply xml provisioning template to SharePoint
    Write-Host "Applying email columns template to SharePoint:" $SharePointUrl -ForegroundColor Green 
    
    If($PnP_PowerShell) {
        Invoke-PnPSiteTemplate -Path $Path -Handlers 'Fields' -WarningAction Ignore
    }
    Else {
        Apply-PnPProvisioningTemplate -Path $Path -Handlers 'Fields' -WarningAction Ignore
    }
    
    Write-Host "`nSuccess! Please create your Email Content Type in SharePoint in the browser, and add the new columns to that Content Type." -ForegroundColor Green
    Pause
}
Catch {
    Write-Host $error[0].Message
}
