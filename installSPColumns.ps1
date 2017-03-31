 <#
        This script prompts you to install the SharePoint PnP commands. Hit enter for SharePoint Online,
        6 for SharePoint 2016,3 for SharePoint 2013
        The script then prompts for the site collectiuon url you wish to install the Email Columns to
        and then applies the email columns template to this site collection        
#>

try {
    Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/OfficeDev/PnP-PowerShell/master/Samples/Modules.Install/Install-SharePointPnPPowerShell.ps1')
    $SharePointUrl = Read-Host -Prompt 'Enter your SharePoint Site Collection Url to install OnePlace Solutions Email Columns to'
    Connect-pnpOnline -url $SharePointUrl
    Apply-PnPProvisioningTemplate -path email-columns.xml 
}
catch {
    Write-Host $error[0].Message
}
