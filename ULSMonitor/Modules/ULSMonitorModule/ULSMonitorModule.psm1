function Add-ULSItem 
{
    <#
	    .SYNOPSIS
	    Creates an item in the ULS list
	    .DESCRIPTION
	    Creates an item in the ULS list
	    .EXAMPLE
	    Create-ULSItem $list $result
	    .PARAMETER $list
	    The ULS list
	    .PARAMETER $result
	    The result
    #>

    [CmdletBinding()]
        param
        (

            [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The SharePoint ULS list object')] [Microsoft.SharePoint.SPList]$list,            
            [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='This is a single result (list item) from a query of the ULs logs')] [Microsoft.SharePoint.Diagnostics.LogFileEntry[]]$result
        )
        
    begin
    {
        $now = Get-Date
        $shortWhileBack = $now.AddMinutes(-$runfrequency)
    }

    process
    {
    
        $item = $list.Items.Add();
        $item["Title"] = "Error found"
        $item["ULSLevel"] = $result.Level
        $item["ULSArea"] = $result.Area
        $item["ULSCategory"] = $result.Category
        $item["ULSMessage"] = $result.Message
        $item["ULSDateTime"] = $result.Timestamp
        $item["ULSEventID"] = $result.EventId
        $item["ULSServer"] = hostname

        $relatedItems = Get-SPLogEvent -StartTime $shortWhileBack | where-object { $_.correlation -eq $result.Correlation } 
        foreach ($relatedItem in $relatedItems)
        {
           $additonalInfo += $relatedItem.level + [Environment]::NewLine + $relatedItem.Message + [Environment]::NewLine + [Environment]::NewLine
        }
        
        $item["ULSRelatedInfo"] = $additonalInfo 
        $item.Update()

        return $item    
    }

    end
    {
    }

}


function Write-Log
{
    <#
	    .SYNOPSIS
	    Writes a message to the ULS logs
	    .DESCRIPTION
	    Writes a message to the ULS logs as Monitorable Error message	    
	    .EXAMPLE
	    Create-ULSItem $message
	    .PARAMETER $message
	    The message to send to the ULS logs	    
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The Messages to report in the ULS logs')] [string]$message            
    )     


    begin
    {
    }

    process
    {
        $diagSvc = [Microsoft.SharePoint.Administration.SPDiagnosticsService]::Local
        $category = New-Object Microsoft.SharePoint.Administration.SPDiagnosticsCategory(“ULS Monitor”,
                                        [Microsoft.SharePoint.Administration.TraceSeverity]::Monitorable,
                                        [Microsoft.SharePoint.Administration.EventSeverity]::Error )
        $diagSvc.WriteTrace(0, $category,  [Microsoft.SharePoint.Administration.TraceSeverity]::Monitorable, $message )
    }

    end
    {
    }
}




function Send-MailAlert
{
    <#
	    .SYNOPSIS
	    Sends an email with a ULS message and its details
	    .DESCRIPTION
	    Sends an email with a ULS message and its details
	    .EXAMPLE
	    Send-MailAlert $email $subject $level $message $event $time $smtpServer
	    .PARAMETER $email
        The email address to send the message to.
	    .PARAMETER $subject
        The subject line of the email
        .PARAMETER $level
        The level of the message in the ULS logs
        .PARAMETER $message
        The message from the ULS logs
        .PARAMETER $event
        The event of the message in the ULs logs
        .PARAMETER $time
        The time stamp the message appeared in the ULS logs
        .PARAMETER $smtpServer
        The SMTP server to use to send the email.	    
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The email address to send the message to')] [string]$email,    
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The subject line of the email')] [string]$subject,            
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The level of the message in the ULS logs')] [string]$level,            
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The Messages to report in the ULS logs')] [string]$message,            
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The event of the message in the ULs logs')] [string]$event,            
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The time stamp the message appeared in the ULS logs')] [string]$time,            
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True, HelpMessage='The SMTP server to use to send the email')] [string]$smtpServer                   
    )   

    if ($message.Trim() -ne "")
    {

        $subjectText = $subject + " - " + $level + " - " + $event.Process + " on " + $hostname + " at " + $time

        Send-MailMessage -To $email -From "noreply@mycorp.com" -Subject $subjectText -Body $message -SmtpServer $smtpServer
        "sent email: " + $message | Write-Debug 
    }
    else
    {
        Write-Debug "Ignoring empty message"
    }
}


function Get-ActionsList($web, $ActionListName)
{
    $actionList = Get-ActionsListTry $web $ActionListName
            
    $actionList = Get-ActionsListTry $web $ActionListName

    return $actionList
}


function Get-ActionsListTry($web, $ActionsListName)
{

		$actionsList = $web.Lists.TryGetList($ActionsListName)

        if ($actionsList -eq $null)
        {
            Write-Host "The actions list $actionslist is not found"
            $createlist = Read-Host "Do you want to create the actions list (Y/N)"

            if ($createlist -eq "Y")
            {
                
                $web.Lists.Add($ActionsListName, "", [Microsoft.SharePoint.SPListTemplateType]::GenericList)
                $actionsList = $web.Lists.TryGetList($ActionsListName)
                    

                # Add the fields to the web
				if ($web.Fields["ULSMessage"] -eq $null)
                {
                    $field = $web.Fields.Add("ULSMessage","Note", 0);
                }

				$choices = New-Object System.Collections.Specialized.StringCollection
				$choices.Add("None")
				$choices.Add("Send email")
				$choices.Add("Run PowerShell")
				$choices.Add("Create List Item Only")
				$spFieldType = [Microsoft.SharePoint.SPFieldType]::Choice
                if ($web.Fields["ULSAction"] -eq $null)
                {
					$web.Fields.Add("ULSAction",$spFieldType,$false,$false,$choices)
                }

                if ($web.Fields["ULSComment"] -eq $null )
                {
					$field = $web.Fields.Add("ULSComment","Text", 0);
                }
                if ($web.Fields["ULSPowerShell"] -eq $null)
                {
					$field = $web.Fields.Add("ULSPowerShell","Text", 0);
                }
                # Add the fields to the list
				$field = $actionsList.Fields.Add("ULSMessage","Note", 0);
                $field = $actionsList.Fields.Add("ULSAction",$spFieldType,$false,$false,$choices)
 				$field = $actionsList.Fields.Add("ULSComment","Text", 0);
				$field = $actionsList.Fields.Add("ULSPowerShell","Text", 0);

                $ctype = $web.AvailableContentTypes | Where {$_.Name -eq "ULSAction" }
                 
                if ($ctype -eq $null)
                {
                $ctype = new-object Microsoft.SharePoint.SPContentType($web.ContentTypes["Item"], $web.contenttypes, "ULSAction")
                $ctype.Group = "ULS Monitor"
                                                         

                # add the fields to the content type

                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $actionsList.Fields["ULSMessage"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $actionsList.Fields["ULSAction"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $actionsList.Fields["ULSComment"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $actionsList.Fields["ULSPowerShell"]))
                   
                $web.ContentTypes.Add($ctype)
                $ctype.Update()
                }

                $view = $actionsList.Views[0]                 
                $view.ViewFields.Add("ULSMessage")
                $view.ViewFields.Add("ULSAction")
                $view.ViewFields.Add("ULSComment")
                $view.ViewFields.Add("ULSPowerShell")
                $view.Update()

                $actionsList.ContentTypes.Add($ctype)
                $actionsList.Update();




            }
            else
            {
                exit
            }
        }

		return $actionsList
}


function Get-ULSList($web, $ULSListName)
{
    $ulsList = Get-ULSListTry $web $ULSListName
            
    $ulsList = Get-ULSListTry $web $ULSListName

    return $ulsList
}

function Get-ULSListTry($web, $ULSListName)
{
	$ulsList = $web.Lists.TryGetList($ULSListName)

    if ($ulsList -eq $null)
    {
        Write-Host "The ULS list $ulsList is not found"
           
		$createlist = Read-Host "Do you want to create the ULS list (Y/N)"

        if ($createlist -eq "Y")
        {
            # Create the list
                $web.Lists.Add($ULSListName, "", [Microsoft.SharePoint.SPListTemplateType]::GenericList)
                $ulsList = $web.Lists.TryGetList($ULSListName)

                # Add the fields to the web
                if ($web.Fields["ULSArea"] -eq $null)
                {
                $field = $web.Fields.Add("ULSArea","Text", 0);
                }
                if ($web.Fields["ULSCategory"] -eq $null)
                {
                $field = $web.Fields.Add("ULSCategory","Text", 0);
                }
                if ($web.Fields["ULSDateTime"] -eq $null)
                {
                $field = $web.Fields.Add("ULSDateTime","DateTime", 0);
                }
                if ($web.Fields["ULSEventID"] -eq $null)
                {
                $field = $web.Fields.Add("ULSEventID","Text", 0);
                }
                if ($web.Fields["ULSDateTime"] -eq $null)
                {
                $field = $web.Fields.Add("ULSLevel","Text", 0);
                }
                if ($web.Fields["ULSMessage"] -eq $null)
                {
                $field = $web.Fields.Add("ULSMessage","Note", 0);
                }
                if ($web.Fields["ULSRelatedInfo"] -eq $null)
                {
                $field = $web.Fields.Add("ULSRelatedInfo","Note", 0);
                }
                if ($web.Fields["ULSServer"] -eq $null)
                {
                $field = $web.Fields.Add("ULSServer","Text", 0);
                }
                # Associate the fields to the list.

                $field = $ulsList.Fields.Add("ULSArea","Text", 0);
                $field = $ulsList.Fields.Add("ULSCategory","Text", 0);
                $field = $ulsList.Fields.Add("ULSDateTime","DateTime", 0);
                $field = $ulsList.Fields.Add("ULSEventID","Text", 0);
                $field = $ulsList.Fields.Add("ULSLevel","Text", 0);
                $field = $ulsList.Fields.Add("ULSMessage","Note", 0);
                $field = $ulsList.Fields.Add("ULSRelatedInfo","Note", 0);
                $field = $ulsList.Fields.Add("ULSServer","Text", 0);


                $ctype = $web.AvailableContentTypes | Where {$_.Name -eq "ULSItem" }
                 
                if ($ctype -eq $null)
                {
                $ctype = new-object Microsoft.SharePoint.SPContentType($web.ContentTypes["Item"], $web.contenttypes, "ULSItem")
                $ctype.Group = "ULS Monitor"
                                                         

                # add the fields to the content type

                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $ulsList.Fields["ULSArea"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $ulsList.Fields["ULSCategory"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $ulsList.Fields["ULSDateTime"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $ulsList.Fields["ULSEventID"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $ulsList.Fields["ULSLevel"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $ulsList.Fields["ULSMessage"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $ulsList.Fields["ULSRelatedInfo"]))
                $ctype.FieldLinks.Add((new-object Microsoft.SharePoint.SPFieldLink $ulsList.Fields["ULSServer"]))


                $web.ContentTypes.Add($ctype)
                $ctype.Update()
                }

                $view = $ulsList.Views[0]
                $view.ViewFields.Add("ULSArea")
                $view.ViewFields.Add("ULSCategory")
                $view.ViewFields.Add("ULSDateTime")
                $view.ViewFields.Add("ULSEventID")
                $view.ViewFields.Add("ULSLevel")
                $view.ViewFields.Add("ULSMessage")
                $view.ViewFields.Add("ULSRelatedInfo")
                $view.ViewFields.Add("ULSServer")
                $view.Update()

                 

                $ulsList.ContentTypes.Add($ctype)
                $ulsList.Update();

                 
                 
                                  
            }
            else
            {
                exit
            }

    }
        
	return $ulsList
}


function Get-Settings($settingsFile)
{
    begin
    {
        $settings = "";
    }

    process
    {

      if (Test-Path $settingsFile)
      {
        $settings = Get-Content $settingsFile
      }
      else
      {
        Write-Error "$settingsFile is not found"
        exit
      }
    }

    end
    {
        return $settings
    }
    
}

Function Get-ULSMonitorWeb
{
    <#
	    .SYNOPSIS
	    Gets the web object for the monitor site.
	    .DESCRIPTION
	    Gets the web object for the monitor site and handles any errors in etting the object.
	    .EXAMPLE
	    Get-ULSMonitorWeb $url 
	    .PARAMETER $url
	    The url of the monitor site	    
    #>

    Param(
            [Parameter(Mandatory=$true)][System.String]$monitorSiteUrl 
        )

    begin
    {
        $web = $null
    }


    process
    {
        try
        {

            $web = Get-SPWeb $monitorSiteUrl
        }
        catch
        {
            if ($web -eq $null)
            { 
                Write-Error "Failed to find $monitorSiteUrl"            
            }
        }
    }
    
    end
    {
        return $web
    }

}





Function Edit-CleanULSList
{
    <#
        .DESCRIPTION
        cleanup of expired ULS items
    #>

    param(
           [Parameter(Mandatory=$true)][System.DateTime]$expiryDate,
           [Parameter(Mandatory=$true)][Microsoft.SharePoint.SPList]$ulsList
         )

    begin
    {
    }

    process
    {
    
                $spqQuery = New-Object Microsoft.SharePoint.SPQuery 
                $spqQuery.Query =  
                        "<Where> 
                            <Leq> 
                                <FieldRef Name='ULSDateTime' /> 
                                <Value Type='DateTime'>"+ $expiryDate.ToString("yyyy-MM-dd:hh:mm:ss") + "</Value> 
                           </Leq> 
                        </Where>" 
                $spqQuery.ViewFields = "<FieldRef Name='ID' /><FieldRef Name='ULSDateTime' />" 
                $spqQuery.ViewFieldsOnly = $true 
                $splListItems = $ulsList.GetItems($spqQuery)

                if ($splListItems.Count -eq 0)
                {
                    Write-Debug "No items to clean up"
                }
                else
                {
                   foreach ($splListItem in $splListItems)
                   {           
                      "about to expire message" + $splListItem["ID"] + ": " + $splListItem["ULSDateTime"] | Write-Debug 
                      $id = $splListItem["ID"]
                     $ulsList.getitembyid($id).Delete()
                   }
                }
    }

    end
    {

    }
}



Function Process-ULSLogs 
{
    <#
        .DESCRIPTION

    #>

    param(
        [Parameter(Mandatory=$true)][System.DateTime]$since,
        [Parameter(Mandatory=$true)][System.String]$emailAddressTo,
        [Parameter(Mandatory=$true)][System.String]$smtpServer,
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.SPList]$actionsList,
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.SPList]$ulsList
    )

    $results = Get-SPLogEvent -StartTime $since | Where-Object {$_.Level -ne "INFO" -and $_.Level -ne "High" -and $_.Level -ne "Information"  -and $_.Level -ne "DEBUG" -and $_.Level -ne "Medium" } 
	        
    if ($results.Count -eq 0)
	{
	    Write-Debug "No new errors found"
    }
	else
	{   
	    $results.Count.ToString() + " errors found" | Write-Debug 

	    foreach($result in $results)
		 
	    {
			   
			$message =  $result.Message 
			#"Looking for match in the action list: " + $message | Write-Debug 
			$action = $null

			
			    foreach ( $actionitem in $actionsList.items)
			    { 
			        $actionmsg = $actionitem["ULSMessage"]; 
			   
			        if ("$message" -like "*$actionmsg*")
			        {
				        $action = $actionitem["Action"]
				        break
			        }  
			    }

			    if ($action -ne "None")
			    {
			        $item = Add-ULSItem $ulsList $result
			        #  "Added error " + $shortWhileBack + " " + $message | Write-Debug 
			    }
                if ($message -eq "" -or $message -eq $null)
                {
                    Write-Debug "ignore"
                }
                else
                {
			        $message += [Environment]::NewLine + [Environment]::NewLine  + "For more information see: " + $web.Url + "/Lists/" + $ULSListName + "/DispForm.aspx?ID=" + $item.ID 
			    }
                $time = $result.Timestamp.ToString()
            
			    if ($action -eq $null)
			    {
				    #Write-Debug "Action not set"
	   
				    Send-MailAlert $emailAddressTo "Unknown ULS Message" $result.Level $message $result $time $smtpServer

				    "Added error" + $since + $message | Write-Debug 
			    }
			    else
			    {
				

				    switch ($actionitem["Action"])
				    {
				        "None" {  
                                    # "ignoring: " + $message | Write-Debug 
                                }
				        "Send email" {
								        Send-MailAlert $emailAddressTo "ULS Message" $result.Level $message $result $time

								    }
				        "Create List Item Only" { Write-Debug "Creating Item only"}
				        default { 
							        Write-Debug "Default Action"

							        Send-MailAlert $emailAddressTo "Unknown ULS Message" $result.Level $message $result $time
						        }
				    }

			   

			    }
			
			    	
			
	        }
        }

}
