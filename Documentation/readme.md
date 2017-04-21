# ULS Monitor
ULS Monitor scans your SharePoint ULS logs for all messages that are warnings/criticals/monitorable or unexpected errors. These errors are reported in a SharePoint list and emailed to selected users.
## Installation
To install this solution unzip the solution. This will give you:
* a PowerShell scripts 
* a PowerShell module
* a settings.xml
## Settings.xml
Within the settings.xml some settings will need to be changed.
This is the starter version of the settings.xml. 
<Settings>
    <Monitor Site="http://devsp.mydev.local/ULSMonitor">
    <Lists>
       <List Name="ULS" Url="ULS" />
       <List Name="Actions" Url="Actions" />
    </Lists>
    <Jobs>
      <Job Name="UpdateULSList" ExpireAfter="2" Frequency="5" >
        <Alerts>
         <Alert Type="email" >
          <Address From="NoReply@mydev.local" To="Test@mydev.local" />
         </Alert>
        </Alerts>
      </Job>
    </Jobs>
    </Monitor>
    <SMTP Server="devsp.mydev.local">
    </SMTP>
</Settings>
Most likely all you will need to update is your SMTP server, the from and to email addresses and the URL of the ULS Monitor site.
*Monitor Site:* The site where you want the lists used by ULS Monitor to be created. This should be a newly created site.
*List ULS:* This is a list used to store all ULS error messages into
*List Actions:* This list contains items describing what to do with each error. By default this list is empty, but unless you like to receive the same message many times, you might want to add some items here.*UpdateULSList Job:* Here you can configure how long to keep messages for (in days) in the ULS List. After the specified number of days items will be removed.*Alert Email address:* Specify the to and from for the alerts to be sent to.*SMTP Server*: specify an SMTP server to handle the emails sent out
## Script configuration
Once the script has been copied to a SharePoint server and the settings.xml has been adjusted the script should be run from a PowerShell window.
The script will ask you if you  want to create some lists. Type Y to create the lists for both the ULS and the Actions list.
Once these lists have been created the PowerShell script (MonitorULSLogs.ps1) can be configured to run as a scheduled task.
## Actions List
Items can be added to the actions list. To configure messages in the ULS logs to be ignored. Simply add part of the ULs message in the ULSMessage field of the Actions list item and configure  the action (None, Create List Item, Send email only). The Run PowerShell option is not supported yet.
