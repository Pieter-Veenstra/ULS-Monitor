<#
    .SYNOPSIS
    Monitor ULS logs
    .DESCRIPTION
    Monitors the SharePoint ULS logs on the local server and configures ULS Monitor the first time it is run
    .VERSION
    1.0.3
    .AUTHOR
    Pieter Veenstra
#>



Function Add-ULSMonitor
{
    <#
	    .SYNOPSIS
	    Adds the ULS Monitor process to the local server
	    .DESCRIPTION
	    Adds the ULS Monitor process to the local server
	    .EXAMPLE
	    Add-ULSMonitor	    
    #>

    begin
    {

        #$DebugPreference = "Continue"
        $DebugPreference = "SilentlyContinue"

        if ($env:PSModulePath -notlike "*$path\Modules\*")
        {
            "Adding ;$path\Modules to PSModulePath" | Write-Debug 
            $env:PSModulePath += ";$path\Modules\"
        }

        # Unloading the functions and reloading them.
        if (Get-Module -Name ULSMonitorModule)
        {
            Remove-SPPSSnapIn
            Remove-Module -Name ULSMonitorModule
        }

        Import-Module -Name ULSMonitorModule
        Add-SPPSSnapIn


        Start-SPAssignment -global

        #clearing all the variable that are used by the settings xml
        $config = $null
        $settings = $null
        $runfrequency = $null
        $monitorSiteUrl = $null
        $ULSListName = $null
        $ActionsListName =  $null
        $smtpServer = $null
        $emailAddress = $null
        $emailAddressFrom = $null

        # Other variables used by the ULS Monitor
        $hostname = hostname
        $now = Get-Date

        # reset the objects used by the ULS Monitor
        $site = $null
	    $web = $null
	    $ulsList = $null
	    $actionslist = $null

        #reset the variable that contains all of the errors from the last period 
        $results = $null;

    }


    process
    {
    
        Clear-Host

        $Error.Clear()

        Write-Debug "Start the work"

        
        [xml]$settings = Get-Settings "$path\settings.xml"


        # Script Settings
        $expiry = $settings.Settings.Monitor.Jobs.Job.ExpireAfter         #number of days to keep messages in the ULS list 
        $runfrequency = $settings.Settings.Monitor.Jobs.Job.Frequency #number of minutes to cover in the logs  
        $monitorSiteUrl = $settings.Settings.Monitor.Site
        $ULSListName = ($settings.Settings.Monitor.Lists.List |Where {$_.Name -eq "ULS"}).Url
        $ActionsListName = ($settings.Settings.Monitor.Lists.List |Where {$_.Name -eq "Actions"}).Url
        
        $smtpServer = $settings.Settings.SMTP.Server
        $emailAddress = $settings.Settings.Monitor.Jobs.Job.Alerts.Alert.Address.To
        $emailAddressFrom = $settings.Settings.Monitor.Jobs.Job.Alerts.Alert.Address.From

   
        try
        {
            if ($expiry -eq $null)
            {
                Write-Host "expiry is not set"
                exit
            }

            if ($now -eq $null)
            {
                Write-Host "now is not set"
                exit
            }

            if ($runfrequency -eq $null)
            {
                Write-Host "runfrequency is not set"
                exit
            }

            $expiryDate =$now.AddDays($expiry)
            $shortWhileBack = $now.AddMinutes(-$runfrequency)

            
            $web = Get-ULSMonitorWeb $monitorSiteUrl
            
            if ( $web -ne $null)
            {

                $ulsList = Get-ULSList $web $ULSListName            
			    $actionsList = Get-ActionsList $web $ActionsListName
            
                Process-ULSLogs $shortWhileBack $emailAddress $smtpServer $actionsList $ulsList

                Edit-CleanULSList $expiryDate $ulsList
              }
        }
        catch
        {
            Write-Host "Something has gone wrong"
            Write-Host $Error[$Error.Count -1]
        }
        

      

  
    }
    

    end
    {
        if ($web -ne $null)
        {
            $web.Dispose()
	    }

        Stop-SPAssignment -global
        Remove-SPPSSnapIn

        if (Get-Module -Name ULSMonitorModule)
        {
            Remove-Module -Name ULSMonitorModule
        }

    }

    
}    



#################
#Functions 
#################

Function Remove-SPPSSnapIn
{
    <#
	    .SYNOPSIS
	    Load the SharePoint snap in
	    .DESCRIPTION
	    Load the SharePoint snap in for PowerShell
	    .EXAMPLE
	    Remove-SPPSSnapIn
    #>
    begin
    {      
    }

    process
    {     
	   if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) 
	   {
          Remove-PSSnapin "Microsoft.SharePoint.PowerShell"
       }
    }

    end
    {
    }
}

Function Add-SPPSSnapIn
{
    <#
	    .SYNOPSIS
	    Load the SharePoint snap in
	    .DESCRIPTION
	    Load the SharePoint snap in for PowerShell
	    .EXAMPLE
	    Load-SnapIn	
    #>

    begin
    {
      $ver = $host | select version
    }

    process
    {
       if ($ver.Version.Major -gt 1) 
	   {
	      $host.Runspace.ThreadOptions = "ReuseThread"		  
	   }  
	   
	   if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) 
	   {
          Add-PSSnapin "Microsoft.SharePoint.PowerShell"
       }  
    }

    end
    {
    }

}


$path = Split-Path -parent $MyInvocation.MyCommand.Definition
        
Add-ULSMonitor