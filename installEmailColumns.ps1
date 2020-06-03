 <#
        This script prompts you to install the SharePoint PnP commands. Hit enter for SharePoint Online,
        6 for SharePoint 2016,3 for SharePoint 2013
        The script then prompts for the site collection URL you wish to install the Email Columns to
        and then applies the email columns template to this site collection        
#>

try {    
    Set-ExecutionPolicy Bypass -Scope Process 
    #Prompt for SharePoint URL     
    $SharePointUrl = Read-Host -Prompt 'Enter your SharePoint Site Collection URL to install OnePlace Solutions Email Columns to'
    
    #Connect to site collection
    If($SharePointUrl -match ".sharepoint.com/"){
        Write-Host "Enter SharePoint credentials (your email address for SharePoint Online):" -ForegroundColor Green  
        Connect-pnpOnline -url $SharePointUrl -UseWebLogin
    }
    Else{
        Write-Host "Enter SharePoint credentials (domain\username for on-premises):" -ForegroundColor Green  
        Connect-pnpOnline -url $SharePointUrl 
    }

    #Download xml provisioning template
    $WebClient = New-Object System.Net.WebClient   
    $Url = "https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml"    
    $Path = "$env:temp\email-columns.xml"

    Write-Host "Downloading provisioning xml template:" $Path -ForegroundColor Green 
    $WebClient.DownloadFile( $Url, $Path )   
    #Apply xml provisioning template to SharePoint
    Write-Host "Applying email columns template to SharePoint:" $SharePointUrl -ForegroundColor Green 
    Apply-PnPProvisioningTemplate -path $Path
    Write-Host "`nSuccess! Please add the new columns to your Email Content Type." -ForegroundColor Green
    Write-Host "Exiting script." -ForegroundColor Yellow
}
catch {
    Write-Host $error[0].Message
}
