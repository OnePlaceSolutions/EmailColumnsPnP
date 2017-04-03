 <#
        This script prompts you to install the SharePoint PnP commands. Hit enter for SharePoint Online,
        6 for SharePoint 2016,3 for SharePoint 2013
        The script then prompts for the site collectiuon url you wish to install the Email Columns to
        and then applies the email columns template to this site collection        
#>

try {
    Set-ExecutionPolicy Bypass -Scope Process       
    $SharePointUrl = Read-Host -Prompt 'Enter your SharePoint Site Collection Url to install OnePlace Solutions Email Columns to'
    Connect-pnpOnline -url $SharePointUrl
    $WebClient = New-Object System.Net.WebClient
   
    $Url = "https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml"    
    $Path = "$env:temp\emailcolumns.xml"
    Write-Host "Downloading provisioning xml template:" $Path -ForegroundColor Green 
    $WebClient.DownloadFile( $Url, $Path ) 
  
    #(New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml', '.\EmailColumns.xml')
    Write-Host "Applying email columns template to SharePoint:" $Url -ForegroundColor Green 
    Apply-PnPProvisioningTemplate -path $Path
   
}
catch {
    Write-Host $error[0].Message
}
