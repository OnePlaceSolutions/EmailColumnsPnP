 <#
        This script prompts you to install the SharePoint PnP commands. Hit enter for SharePoint Online,
        6 for SharePoint 2016,3 for SharePoint 2013
        The script then prompts for the site collectiuon url you wish to install the Email Columns to
        and then applies the email columns template to this site collection        
#>

try {   
    $SharePointUrl = Read-Host -Prompt 'Enter your SharePoint Site Collection Url to install OnePlace Solutions Email Columns to'
    Connect-pnpOnline -url $SharePointUrl
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml', '.\EmailColumns.xml')
    Apply-PnPProvisioningTemplate -path ".\EmailColumns.xml"
   
}
catch {
    Write-Host $error[0].Message
}
