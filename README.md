Deploy OnePlace Solutions Email Columns to a Site Collection
============================================================

This document describes the steps to deploy the OnePlace Solutions Email Columns
to a site collection. The deployment requires the use of PowerShell and
SharePoint Patterns and Practices (PnP) PowerShell cmdlets.

This script does **not** deploy Content Types. For Deploying Email Columns and Content Types simultaneously please see [ContentTypeDeploymentPnP](https://github.com/OnePlaceSolutions/ContentTypeDeploymentPnP).

Pre-requisites
--------------

1.  SharePoint Online, SharePoint 2019 on-premise, SharePoint 2016 on-premise, or SharePoint 2013
    on-premise.

2.  PowerShell v3.0 or greater installed on the machine. Windows 10/8.1
    and Windows Server 2012 and greater are all ready to go. Windows 7
    is preinstalled with v2.0 of PowerShell. PowerShell needs to be
    upgraded on Windows 7 machines. This can be done by downloading and
    installing the Windows Management Framework 4.0 from here:
    <https://www.microsoft.com/en-au/download/details.aspx?id=40855> .
    Download and install either the x64 or x86 version based on your
    version of Windows 7:

    ![](./README-Images/image1.png)

3.  (SharePoint On-Premise Only) [The SharePoint PnP PowerShell cmdlets](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps). You will need to install only the the cmdlets that target your version of SharePoint on the machine you are running the script from. If you have installed the cmdlets previously using an MSI file these need to be uninstalled from Control Panel, but if you have installed the cmdlets previously using PowerShell Get you can update them with this command:
    ```
    Update-Module SharePointPnPPowerShell<version>
    ```
    
    This is the command pictured to install the PnP Cmdlets via PowerShell Get:
    ```
    Install-Module SharePointPnPPowerShell<version>
    ```
    > ![](./README-Images/installPnPClassic.png)

5.  (SharePoint Online Only) (Multi-Tenant supported) [The latest PnP.PowerShell](https://pnp.github.io/powershell/articles/installation.html) installed on the machine you are running the script from. You can run the below command in PowerShell (as Administrator) to install it. 

    Install new PnP.PowerShell Cmdlets:
    ```
    Install-Module -Name "PnP.PowerShell"
    ```
    Note that you will need to ensure you have uninstalled any previous 'Classic' PnP Cmdlets prior to installing this. If you have installed the cmdlets previously using an MSI file these need to be uninstalled from Control Panel, but if you have installed the cmdlets previously using PowerShell Get you can uninstall them with this command (as Administrator):

    ```
    Uninstall-Module 'SharePointPnPPowerShellOnline'
    ```
    
    *PnP Management Shell access does not need to be granted for this script, as it is operating in 'single site mode'.*
    
4.  (Optional, SharePoint Online Only) Content Type Hub Administrator Access
    If you wish to use the Email Site Columns in a Site Content Type deployed using the [Content Type Gallery in SharePoint Online](https://docs.microsoft.com/en-us/sharepoint/create-customize-content-type), you will need Administrative permissions on the Site Collection that supports this feature at 'https://<yourTenant>.sharepoint.com/sites/contenttypehub'. You can then enter this Site Collection URL in Step 3 of the script and continue on to using the Content Type Gallery after the script has finished.

Offline Scripting
--------------------------------------

If you have limited ability to run scripts from the internet in your environment, please download  [EmailColumnsPnPOfflineBundle.zip](https://github.com/OnePlaceSolutions/EmailColumnsPnP/raw/master/EmailColumnsPnPOfflineBundle.zip) above from this repo and extract all contents to one folder. Run the PowerShell script from that folder and it will recognize the XML file containing the Email Columns configuration is present. You can then continue from Step 3 below.

Installing Email Columns to SharePoint
--------------------------------------

1.  Start PowerShell on your machine:

    ![C:\\Users\\COLINW\~1\\AppData\\Local\\Temp\\SNAGHTML6ab5776c.PNG](./README-Images/image4.png)

2.  Copy and paste the following command into your PowerShell command
    window and hit enter:

    > **Invoke-Expression (New-Object
    > Net.WebClient).DownloadString(‘https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/installEmailColumns.ps1’)**
    >
    > Copy the text above, then in the PowerShell window right click at the
    > cursor and the command will be pasted into the window, then hit the
    > enter key to execute the command:

    ![](./README-Images/image5.png)

3.  The PowerShell script will execute and prompt you to enter the Site
    Collection Url for the Site Collection you wish to deploy the Email
    columns to. You can either type it in or copy and paste the url into
    the command window and hit enter:

    ![](./README-Images/image6.png)

    ![](./README-Images/image7.png)

4.  You will be asked to enter your credentials for SharePoint. For
    SharePoint Online it will be your email address, for on-premise it
    will your domain\\username:

    ![](./README-Images/image8.png)

5.  The email columns template will then be downloaded and applied to
    your site collection:

    ![](./README-Images/image9.png)

    ![](./README-Images/image10.png)

6.  Repeat from step 2 above for any other Site Collections. If at a
    later date you need to update the OnePlace Solutions Email columns,
    perform the steps from step 2 and any modifications to existing
    columns will be applied as well as the addition of any new email
    columns we add.


