 <#
        This script prompts you for the site collection URL you wish to install the Email Columns to
        and then applies the email columns template to this site collection.
        This is compatible with both classic PnP cmdlets and the new PnP.PowerShell cmdlets for SharePoint Online (no App consent required)   
#>
$ErrorActionPreference = 'Stop'

Try {
    Set-ExecutionPolicy Bypass -Scope Process
}
Catch {
    $_
    Write-Host "`nCouldn't set PowerShell Execution Policy to 'Bypass' for current session."
    Write-Host "Try entering this command and resolve as necessary, or allow unsigned scripts temporarily: Set-ExecutionPolicy Bypass -Scope Process"
    Write-Host "Please see this page for more information on PowerShell execution policies: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.1"
}

Try {
    Write-Host "Checking for PnP module..."
    Try{
        $installedPnPModules = Get-InstalledModule -Name "*PnP*"
        [boolean]$PnP_PowerShell = ($installedPnPModules.Name -eq "PnP.PowerShell") 
    }
    Catch {
        $PnP_PowerShell = $false
        Write-Host "No PnP Cmdlets detected. Please check pre-requisites before continuing." -ForegroundColor Yellow
        Pause
    }
    #Prompt for SharePoint URL     
    $SharePointUrl = Read-Host -Prompt 'Enter your SharePoint Site Collection URL to install OnePlace Solutions Email Columns to'
    
    #Connect to site collection
    If($SharePointUrl -match ".sharepoint.com/") {
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
    
    $rawXml = Get-Content $Path
        
    #To fix certain compatibility issues between site template types, we will just pull the Field XML entries from the template
    ForEach ($line in $rawXml) {
        Try {
            If (($line.ToString() -match 'Name="Em') -or ($line.ToString() -match 'Name="Doc')) {
                Add-PnPFieldFromXml -fieldxml $line -ErrorAction Stop
            }
        }
        Catch {
            If($($_.Exception.Message) -notmatch 'duplicate') {
                Throw $_
            }
        }
    }
    
    Start-Sleep -Seconds 3

    Try {
        $emailColumns = $null
        $emailColumns = Get-PnPField -Group "OnePlace Solutions"
        #Check if we have 35 columns in our Column Group
        If ($emailColumns.Count -eq 35) {
            Write-Host "All Email columns present!"
            Write-Host "`nSuccess! Please create your Email Content Type in SharePoint in the browser, and add the new columns to that Content Type." -ForegroundColor Green
            Pause
        }
        Else {
            Write-Host "Not all Email Columns detected after deployment. Please check Site Columns in the Site Collection."
            Write-Host "Number of columns detected in 'OnePlace Solutions' group: $($emailColumns.count)"
            Throw $_
        }
    }
    Catch {
        Throw $_
    }
}
Catch {
    $_
    Write-Host "`nException Message: $($error[0].Message)"
    Write-Host "`nPlease contact support for further assistance."
    Pause
}
