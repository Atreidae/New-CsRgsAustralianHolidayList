<#  
    .SYNOPSIS
    This script creates RGS holidaysets for Australian states based on live data from the Australian Government website


    .DESCRIPTION
    Created by James Arber. www.skype4badmin.com
    Although every effort has been made to ensure this list is correct, dates change and sometimes I goof.
    Please use at your own risk.

    Holiday Data taken from https://data.gov.au/data/dataset/b1bc6077-dadd-4f61-9f8c-002ab2cdff10

    .NOTES
    Version             : 3.0
    Date                : 06/01/2020
    Lync Version        : Tested against Skype4B Server 2015 and Lync Server 2013
    Author              : James Arber
    Header stolen from  : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"

    Revision History

    : v3.0: The ScoMo holiday build
    : ScoMo takes a holiday and so does the XML data I was using to get accurate dates
    : Updated, Changed to data.gov.au Datasource. Alot of rewriting
    : Updated, Functions replaced with newer versions
    : Added, Script has been signed, so you know what your running hasnt been fiddled with. Thanks DigiCert!
    : Added, Script now outputs the last date it ingested in proper region format
    : Added, non-domain joined detection
    : Added, Year tagging to Holiday names
    : Added, Better logging messages for writing to the database, esp when an FE is offline
    : Fixed, Script would always put the date 5/11/2018 in one of the last run place holders, now correctly uses ShortDate
    : Fixed, Sorted a bug with unattended mode not updating the last run flag



    : v2.30: The Feedback Build
    : Added display of dates to logs
    : Added a notification at the end of the script showing the last imported date.
    : Added last run date
    : Added RGS Update Time Stamp
    : Added Error handing for 0 FrontEnds
    : Fixed a bug with 1 FrontEnd Pools throwing errors when updating existing holidays




    : v2.2: Cleaned Up Code
    : Fixed a bug with logging system culture
    : Removed some old redundant code
    : Passed script through ISESteriods PSSharper and applied corrections
    : Fixed a few typos
    : Fixed a few bugs introduced cleaning up my dodgy code
    : Fixed a bug with multiple pools using the same holiday set names
    : Deprecated the ServiceID parameter, Specify the pool FQDN instead
    : Added warning for deprecated ServiceID
    : Updated Pat Richard's website
    : Removed PowerShell 5.1 cmdlet (Get-Timezone), now using a WMI query instead


    : v2.1: Added Script logging
    : Updated to use my new autoupdate code
    : Added ability to switch between devel/master branches
    : Added timezone offset detection / warning
    : Added SSL support for the new Govt website requirements


    : v2.01: Migrated to GitHub
    : Minor Typo corrections
    : Check for and prompt user for updates
    : Fixed a bug with multiple pool selection
    : Fixed issues with double spaced event names
    : Added better timeout handling to XML downloads
    : Added better user feedback when downloading XML file
    : Fixed bug with proxy detection failing to execute
    : Removed redundant code for XML lookup
    : Fixed an unattended run bug
    : Fixed commandline switch descriptions


    : v2.0: Update for XML Support
    : Added Autodetecton of single RGS pool
    : Complete Rewrite of existing rule rewrite code, Should make less red text now.
    : Added Region detection, Will prompt to change regions or try to use US date format
    : More user friendly and better instructions
    : Fixed a few typo's causing dates to be incorrect.
    : Fixed alot of grammatical errors
    : Added XML download and implementation with proxy support
    : Auto removes any dates not listed by the Australian Government (such as old dates) if the $RemoveExistingRules is set
    : Script no longer deletes existing timeframes, No need to re-assign to workflows!


    : v1.1: Fix for Typo in Victora Holiday set
    : Fix ForEach loop not correctly removing old time frames
    : Fix Documentation not including the SID for ServiceID parameter


    : v1.0: Initial Release

    Disclaimer: Whilst I take considerable effort to ensure this script is error free and wont harm your enviroment.
    I have no way to test every possible senario it may be used in. I provide these scripts free
    to the Lync and Skype4B community AS IS without any warranty on its appropriateness for use in
    your environment. I disclaim all implied warranties including,
    without limitation, any implied warranties of merchantability or of fitness for a particular
    purpose. The entire risk arising out of the use or performance of the sample scripts and
    documentation remains with you. In no event shall I be liable for any damages whatsoever
    (including, without limitation, damages for loss of business profits, business interruption,
    loss of business information, or other pecuniary loss) arising out of the use of or inability
    to use the script or documentation.

    Acknowledgements
    : Testing and Advice
    Greig Sheridan https://greiginsydney.com/about/ @greiginsydney

    : Testing
    Sean Werner https://www.linkedin.com/in/sean-werner-88ab126b/ @swerner1k

    : Auto Update Code
    Pat Richard https://ucunleashed.com @patrichard

    : Proxy Detection
    Michel de Rooij	http://eightwone.com

    : Code Signing Certificate
    DigiCert https://www.digicert.com/


    .INPUTS
    None. New-CsRgsAustralianHolidayList.ps1 does not accept pipelined input.

    .OUTPUTS
    New-CsRgsAustralianHolidayList.ps1 creates multiple new instances of the Microsoft.Rtc.Rgs.Management.WritableSettings.HolidaySet object and cannot be piped.

    .PARAMETER -ServiceID <RgsIdentity> 
    Service where the new holiday set will be hosted. For example: -ServiceID "service:ApplicationServer:SFBFE01.Skype4badmin.com/1987d3c2-4544-489d-bbe3-59f79f530a83".
    To obtain your service ID, run Get-CsRgsConfiguration -Identity FEPool01.skype4badmin.com
    If you dont specify a ServiceID or FrontEndPool, the script will try and guess the frontend to put the holidays on.

    .PARAMETER -FrontEndPool <FrontEnd FQDN> 
    Frontend Pool where the new holiday set will be hosted. 
    If you dont specify a ServiceID or FrontEndPool, the script will try and guess the frontend to put the holidays on.
    Specifiying this instead of ServiceID will cause the script to confirm the pool unless -Unattended is specified

    .PARAMETER -RGSPrepend <String>
    String to Prepend to Listnames to suit your environment

    .PARAMETER -DisableScriptUpdate
    Stops the script from checking online for an update and prompting the user to download. Ideal for scheduled tasks

    .PARAMETER -RemoveExistingRules
    Deprecated. Script now updates existing rulesets rather than removing them. Kept for backwards compatability

    .PARAMETER -Unattended
    Assumes yes for pool selection critera when multiple pools are present and Poolfqdn is specified.
    Also stops the script from checking for updates
    Check the script works before using this!

    .LINK
    http://www.UcMadScientist.com/australian-holiday-rulesets-for-response-group-service/


    .EXAMPLE

    PS C:\> New-CsRgsAustralianHolidayList.ps1 -ServiceID "service:ApplicationServer:SFBFE01.skype4badmin.com/1987d3c2-4544-489d-bbe3-59f79f530a83" -RGSPrepend "RGS-AU-"

    PS C:\> New-CsRgsAustralianHolidayList.ps1 

    PS C:\> New-CsRgsAustralianHolidayList.ps1 -DisableScriptUpdate -FrontEndPool AUMELSFBFE.Skype4badmin.local -Unattended

#>
# Script Config
#Requires -Version 3
[CmdletBinding(DefaultParametersetName = 'Common')]
param(
  [Parameter(Position = 1)] [string]$ServiceID,
  [Parameter(Position = 2)] [string]$RGSPrepend,
  [Parameter(Position = 3)] [string]$FrontEndPool,
  [Parameter(Position = 4)] [switch]$DisableScriptUpdate,
  [Parameter(Position = 4)] [switch]$Unattended,
  [Parameter(Position = 5)] [switch]$RemoveExistingRules,
  [Parameter(Position = 6)] [string]$LogFileLocation,
  [Parameter(Position = 7)] [switch]$DownloadOnly
)
#region config
[Net.ServicePointManager]::SecurityProtocol = 'tls12, tls11, tls'
$MaxCacheAge = 30 # Max age for XML cache, older than this # days will force info refresh
$DataSetCache = Join-Path -Path $PSScriptRoot -ChildPath 'b1bc6077-dadd-4f61-9f8c-002ab2cdff10.xml' #Filename for the XML data list
If (!$LogFileLocation) 
{
  $script:LogFileLocation = $PSCommandPath -replace '.ps1', '.log'
}
[float]$ScriptVersion = '3.0'
[string]$GithubRepo = 'New-CsRgsAustralianHolidayList'
[string]$GithubBranch = 'master' #todo
[string]$BlogPost = 'http://www.UcMadScientist.com/australian-holiday-rulesets-for-response-group-service/'
#endregion config


#region Functions
Function Write-Log
{
   <#
      .SYNOPSIS
      Function to output messages to the console based on their severity and create log files

      .DESCRIPTION
      It's a logger.

      .PARAMETER Message
      The message to write

      .PARAMETER Path
      The location of the logfile.

      .PARAMETER Severity
      Sets the severity of the log message, Higher severities will call Write-Warning or Write-Error

      .PARAMETER Component
      Used to track the module or function that called "Write-Log" 

      .PARAMETER LogOnly
      Forces Write-Log to not display anything to the user

      .EXAMPLE
      Write-Log -Message 'This is a log message' -Severity 3 -component 'Example Component'
      Writes a log file message and displays a warning to the user

      .NOTES
      N/A

      .LINK
      http://www.UcMadScientist.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>

  PARAM
  (
    [Parameter(Mandatory)][String]$Message,
    [String]$Path = $script:LogFileLocation,
    [int]$Severity = 1,
    [string]$Component = 'Default',
    [switch]$LogOnly

  )
  $Date             = Get-Date -Format 'HH:mm:ss'
  $Date2            = Get-Date -Format 'MM-dd-yyyy'
  $MaxLogFileSizeMB = 10
  
  If(Test-Path -Path $Path)
  {
    if(((Get-ChildItem -Path $Path).length/1MB) -gt $MaxLogFileSizeMB) # Check the size of the log file and archive if over the limit.
    {
      $ArchLogfile = $Path.replace('.log', "_$(Get-Date -Format dd-MM-yyy_hh-mm-ss).lo_")
      Rename-Item -Path ren -NewName $Path -Path $ArchLogfile
    }
  }
         
  "$env:ComputerName date=$([char]34)$Date2$([char]34) time=$([char]34)$Date$([char]34) component=$([char]34)$Component$([char]34) type=$([char]34)$Severity$([char]34) Message=$([char]34)$Message$([char]34)"| Out-File -FilePath $Path -Append -NoClobber -Encoding default
  If (!$LogOnly) 
  {
    #If LogOnly is set, we dont want to write anything to the screen as we are capturing data that might look bad onscreen
      
      
    #If the log entry is just Verbose (1), output it to verbose
    if ($Severity -eq 1) 
    {
      "$Date $Message"| Write-Verbose
    }
      
    #If the log entry is just informational (2), output it to write-host
    if ($Severity -eq 2) 
    {
      "Info: $Date $Message"| Write-Host -ForegroundColor Green
    }
    #If the log entry has a severity of 3 assume it's a warning and write it to write-warning
    if ($Severity -eq 3) 
    {
      "$Date $Message"| Write-Warning
    }
    #If the log entry has a severity of 4 or higher, assume it's an error and display an error message (Note, critical errors are caught by throw statements so may not appear here)
    if ($Severity -ge 4) 
    {
      "$Date $Message"| Write-Error
    }
  }
}

Function Get-IEProxy 
{
  Write-Log -Message 'Checking for proxy settings' -severity 1
  If ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0) 
  {
    $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
    if ($proxies) 
    {
      if ($proxies -ilike '*=*') 
      {
        return $proxies -replace '=', '://' -split (';') | Select-Object -First 1
      }
      Else 
      {
        return ('http://{0}' -f $proxies)
      }
    }
    Else 
    {
      return $null
    }
  }
  Else 
  {
    return $null
  }
}

Function Get-ScriptUpdate 
{
  if ($DisableScriptUpdate -eq $false) 
  {
    Write-Log -component 'Self Update' -Message 'Checking for Script Update' -severity 2
    Write-Log -component 'Self Update' -Message 'Checking for Proxy' -severity 1
    $ProxyURL = Get-IEProxy
    If ( $ProxyURL) 
    {
      Write-Log -component 'Self Update' -Message "Using proxy address $ProxyURL" -severity 1
    }
    Else 
    {
      Write-Log -component 'Self Update' -Message 'No proxy setting detected, using direct connection' -severity 1
    }
  }
  $GitHubScriptVersion = Invoke-WebRequest -Uri "https://raw.githubusercontent.com/atreidae/$GithubRepo/$GithubBranch/version" -TimeoutSec 10 -Proxy $ProxyURL
  If ($GitHubScriptVersion.Content.length -eq 0) 
  {
    Write-Log -component 'Self Update' -Message 'Error checking for new version. You can check manually here' -severity 3
    Write-Log -component 'Self Update' -Message $BlogPost -severity 2
    Write-Log -component 'Self Update' -Message 'Pausing for 5 seconds' -severity 2
    Start-Sleep -Seconds 5
  }
  else 
  { 
    if ([float]$GitHubScriptVersion.Content -gt [float]$ScriptVersion) 
    {
      Write-Log -component 'Self Update' -Message 'New Version Available' -severity 3
      #New Version available

      #Prompt user to download
      $title = 'Update Available'
      $Message = 'An update to this script is available, did you want to download it?'

      $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
      'Launches a browser window with the update'

      $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
      'No thanks.'

      $options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

      $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

      switch ($result)
      {
        0 
        {
          Write-Log -component 'Self Update' -Message 'User opted to download update' -severity 1
          Start-Process -FilePath $BlogPost #todo F
          Write-Log -component 'Self Update' -Message 'Exiting Script' -severity 3
          Exit
        }
        1 
        {
          Write-Log -component 'Self Update' -Message 'User opted to skip update' -severity 1
        }
							
      }
    }   
    Else
    {
      Write-Log -component 'Self Update' -Message "Script is up to date on $GithubBranch branch" -severity 2
    }
  }
}


Function Import-ManagementTools 
{
  <#
      .SYNOPSIS
      Function to check for and import Skype4B Management tools
      

      .DESCRIPTION
      Checks for and loads the approprate modules for Skype4B
      Will throw an error and abort script if they arent found

      Version                : 0.5
      Date                   : 20/11/2019 #todo
      Lync Version           : Tested against Skype4B 2015
      Author                 : James Arber
      Header stolen from     : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"

 
      .NOTES
      Version      	      : 0.1
      Date			          : 31/07/2018
      Lync Version		    : Tested against Skype4B 2015
      Author    			    : James Arber


      .LINK
      http://www.skype4badmin.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>

  $function = 'Import-ManagementTools'
  #Import the Skype for Business / Lync Modules and error if not found
  Write-Log -component $function -Message 'Checking for Lync/Skype management tools' -Severity 2
  $ManagementTools             =  $false
  if(!(Get-Module -Name 'SkypeForBusiness')) {Import-Module -Name SkypeForBusiness -Verbose:$false}
  if(!(Get-Module -Name 'Lync')) {Import-Module -Name Lync -Verbose:$false}
  if(Get-Module -Name 'SkypeForBusiness') {$ManagementTools = $true}
  if(Get-Module -Name 'Lync') {$ManagementTools = $true}
  if(!$ManagementTools) {
    Write-Log -Message 'Could not locate Lync/Skype4B Management tools. Script Exiting' -Severity 5 -Component $Function
    Throw 'Unable to load Skype4B/Lync management tools'
    Exit
  }
  Write-Log -component $function -Message 'Management tools loaded successfully' -Severity 2
}


Function Get-XmlFile
{
  $script:LoopCheck ++
  If ($script:LoopCheck -ge 3)
  {
    Write-Log -Message 'Loopcheck Triggered, Get-XMLFile has been called at least 3 times, Abort script' -severity 3
    Exit
  }

  $SessionCacheValid = $False
  If (!(Test-Path -Path $DataSetCache)) #File Missing, Download it
  {
    Write-Log -Message 'Downloading Date list from data.gov.au' -severity 2
    Try 
    {
      Invoke-WebRequest -Uri 'https://data.gov.au/data/dataset/b1bc6077-dadd-4f61-9f8c-002ab2cdff10/gmd' -TimeoutSec 20 -OutFile $DataSetCache -Proxy $ProxyURL #-PassThru
    }
  
    Catch 
    {
      Write-Log -Message 'An error occurred attempting to download XML file automatically' -severity 3
      Write-Log -Message 'Download the file from the URI below, name it "b1bc6077-dadd-4f61-9f8c-002ab2cdff10.xml" and place it in the same folder as this script' -severity 3
      Write-Log -Message 'https://data.gov.au/data/dataset/b1bc6077-dadd-4f61-9f8c-002ab2cdff10/gmd' 
      Throw ('Problem retrieving XML file {0}' -f $error[0])
      Exit 1
    }
  }
  
  If (Test-Path -Path $DataSetCache) #File Exists, Check it 
  {
    Try 
    {
      If ( (Get-ChildItem -Path $DataSetCache).LastWriteTime -ge (Get-Date).AddDays( - $MaxCacheAge)) #Check to see if the XML file is too old to be used
      {
        Write-Log -Message 'Dataset XML file found and not too old. Reading data' -severity 2
        
        #Import the file and check for good XML Data
        Try 
        {
          [xml]$XMLDataSets = Get-Content -Path $DataSetCache 
        }
        Catch
        {
          Write-Log -Message 'Error Reading XML File'  -severity 3
          Throw ('Error Reading XML File {0}' -f $error[0])
        }
        
        #check the file has events
        if ($XMLDataSets.MD_Metadata.distributionInfo.MD_Distribution.transferOptions.MD_DigitalTransferOptions.onLine.CI_OnlineResource.name.count -le 2) 
        {
          Write-Log -Message 'Downloaded file doesnt contain expected data, Less than 2 results' -severity 3
          Throw 'XML Index Check Failed - Too Few Events'
        }
          
       
        #Show the user the updated timestamp
        $IndexUpdated = ($XMLDataSets.MD_Metadata.dateStamp.Date)
        Write-Log -Message "Data.gov.au data last updated on $IndexUpdated"  -severity 2
        $SessionCacheValid = $true

      }
      Else 
      {
        Write-Log -Message 'XML file expired. Will re-download XML from datagov.au' -severity 2
        Try 
        {
          Remove-item -Path $DataSetCache
        }
        Catch 
        {
          Write-Log -Message "Error Removing old XML file, please manualy delete $DatasetCache and run this script again"  -severity 3
          Exit
        }
        #Now the file is gone, Call ourselves
        Get-XmlFile
        Return
               
      }
    }
     
    
    Catch 
    {
      Write-Log -Message 'Error reading XML file or CSV file invalid - Will re-download' -severity 3
      Try 
      {
        Remove-item -Path $DataSetCache
      }
      Catch 
      {
        Write-Log -Message "Error Removing old XML file, please manualy delete $DatasetCache and run this script again"  -severity 3
        Exit
      }
      #Now the file is gone, Call ourselves
      Get-XmlFile
      Return
      
    }
    
  }
  #Now we should have a good XML file, build an array of Data sets, their update times, file names and urls
  $script:XMLDataSetTable = @()
  ForEach($dataset in $XMLDataSets.MD_Metadata.distributionInfo)
  {
    $Filename = (Split-Path -Path $dataset.MD_Distribution.transferOptions.MD_DigitalTransferOptions.onLine.CI_OnlineResource.linkage.url -Leaf )
    $DataSetEntry = New-Object PSObject -property @{
      Name="$($dataset.MD_Distribution.transferOptions.MD_DigitalTransferOptions.onLine.CI_OnlineResource.name.innertext)";
      Url="$($dataset.MD_Distribution.transferOptions.MD_DigitalTransferOptions.onLine.CI_OnlineResource.linkage.url)";
      Filename=$Filename
    }#End of DataSetEntry codeblock
          
    #Add the dataset to the Array
    $script:XMLDataSetTable = ($script:XMLDataSetTable + $DataSetEntry)            
  }
          
 
}

Function Get-CsvFiles
{


  ForEach ($CSVDataSet in $script:XMLDataSetTable[1,2])
  {
   
    $CsvFilename = (Join-Path -Path $PSScriptRoot -ChildPath ($CSVDataSet.filename))#Filename for the CSV data Set
    Write-Log -Message "Checking CSV file $($CSVDataSet.filename)" -Severity 2
            
    If (!(Test-Path -Path $CsvFilename))
    {
      Write-Log -Message "CSV file $($CSVDataSet.filename) missing. Will attempt to download CSV from datagov.au" -severity 2
              
      Try 
      {
        Invoke-WebRequest -Uri ($CSVDataSet.url) -TimeoutSec 20 -OutFile $CsvFilename -Proxy $ProxyURL #-PassThru
        Write-Log -Message 'CSV file downloaded.' -severity 2
      }
      Catch 
      {
        Write-Log -Message 'An error occurred attempting to download CSV files automatically' -severity 3
        Write-Log -Message "Download the file from the URI below, name it $($CSVDataSet.filename) and place it in the same folder as this script" -severity 3
        Write-Log -Message "$($CSVDataSet.url)" 
        Throw ('Problem retrieving CSV file {0}' -f $error[0])
        Exit 1
      }
    }

    #We should have good CSV files
    Try 
    {
      $csvdata = Import-CSV -Path $CsvFilename
      $CSVCount = ($CSVData.Count)
      Write-Log -Message "Imported file with $CSVCount event tags"  -severity 2
      if ($CSVCount -le 10) 
      {
        Write-Log -Message 'Downloaded file doesnt appear to contain correct data'  -severity 3
        throw 'Imported file doesnt appear to contain correct data'
      }
                
    }
    Catch 
    {
      Write-Log -Message 'An error occurred attempting to Import CSV files automatically' -severity 3
      Write-Log -Message "Download the file from the URI below, name it $($CSVDataSet.filename) and place it in the same folder as this script" -severity 3
      Write-Log -Message " $($CSVDataSet.url)" 
      Throw ('Problem retrieving CSV file {0}' -f $error[0])
      Exit 1
    }
  }
      
}


Function Import-DateData
{
  $function = 'Import-DateData'
  #Import the CSV files and turn them into useful data
  Write-Log -component $function -Message 'Reading CSV Data and Generating Holiday list' -Severity 2
  $script:DateData = @()
  $tempdatedata = @()
  ForEach ($CSVDataSet in $script:XMLDataSetTable[1,2])
  {
    $CsvFilename = (Join-Path -Path $PSScriptRoot -ChildPath ($CSVDataSet.filename))#Filename for the CSV data Set
    $csvdata = Import-CSV -Path $CsvFilename
    $tempDateData = ($TempDateData + $csvdata)
  }
  
  #Clean up Data
  Foreach ($Event in $tempDateData)
  { 
    #Check for and fix missing Raw Dates
    if ($event.'raw date' -eq $null) 
    {
      Add-Member -InputObject $event -NotePropertyName 'Raw date' -NotePropertyValue ([datetimeoffset]::parseexact($event.date, 'yyyyMMdd', $null).tounixtimeseconds())
    }
    
    #Add Years to names to make thing easier
    $Event.'Holiday name' = ("$($Event.'Holiday name') $($Event.date.substring(0,4))")
    
    
    
    $script:DateData = ($script:DateData + $event)
  }
  
}
#endregion Functions



#Begin Main
Write-Log -Message "New-CsRgsAustralianHolidayList.ps1 Version $ScriptVersion" -severity 2

#Log everything important about the enviroment, skip if we are just downloading
If (!$DownloadOnly)
{

  $culture = (Get-Culture)
  $GMTOffset = (Get-WmiObject -Query 'Select Bias from Win32_TimeZone')
  Write-Log -Message 'Current system culture'
  Write-Log -Message $culture
  Write-Log -Message 'Current Timezone'
  Write-Log -Message $GMTOffset.bias
  Write-Log -Message 'Checking UTC Offset'
  
 
  
  If ($GMTOffset.bias -lt 480) 
  {
    Write-Log -Message 'UTC Base offset less than +8 hours' -Severity 3
    Write-Log -Message 'Your timezone appears to be misconfigured. This script may not function as expected' -severity 3
    If (!$Unattended)
    {
      #Skip the prompt in unattended mode
      Pause
    }
  }

  $National = $RGSPrepend+'National'

  if ($Unattended) 
  {
    $DisableScriptUpdate = $true
  }
  if ($RemoveExistingRules -eq $true) 
  {
    Write-Log -Message 'RemoveExistingRules parameter set to True. Script will automatically delete existing entries from rules' -severity 3
    Write-Log -Message 'Pausing for 5 seconds' -severity 2
    Start-Sleep -Seconds 5
  }
}


#Get Proxy Details
$ProxyURL = Get-IEProxy
If ($ProxyURL) 
{
  Write-Log -Message "Using proxy address $ProxyURL" -severity 2
}
Else 
{
  Write-Log -Message 'No proxy setting detected, using direct connection' -severity 1
}

#Check for Script update
if ($DisableScriptUpdate -eq $false) 
{
  Get-ScriptUpdate
}

#Import the Skype4B/Lync tools
If (!$DownloadOnly) 
{
  Import-ManagementTools
}


#Check for Data XML file and download it if its out of date / missing
$script:LoopCheck = 0
Write-Log -Message 'Checking for XML index file' -severity 2
Get-XmlFile


#We have the XML file, now grab the CSV files       
Write-Log -Message "XML File found. Checking for CSV cache" -severity 2
Write-Log -Message 'Checking for CSV Files' -severity 1

Get-CSVFiles

Write-Log -Message "CSV cache looks okay" -severity 2

If ($DownloadOnly)
{
  Write-Log -Message "Downloads completed. Exiting Script"  -severity 2
  Exit
}


#Now, turn this into holiday data
Import-DateData


#Now we check the enviroment
Write-Log -Message 'Gathering Front End Pool Data' -severity 2
try 
{
  $Pools = (Get-CsService -Registrar)
}
Catch
{
  Write-Log -Message "I couldn't execute Get-CsService correctly" -severity 3
  Write-Log -Message 'This usualy indicates the Skype4B management tools are missing, the CMS is unavailable, or this PC is not joined to the Skype4B domain' -severity 3
  Write-Log -Message 'Try running this script again from your management server or FrontEnd' -severity 3
  Exit
}


#Check for and warn the user if its not being run on an Australian configured server
Write-Log -Message 'Checking Region Info' -severity 1
$ConvertTime = $false
$region = (Get-Culture)
if ($region.Name -ne 'en-AU') 
{
  #We're not running en-AU region setting, Warn the user and prompt them to change
  Write-Log -Message 'This script is only supported on systems running the en-AU region culture' -severity 3
  Write-Log -Message 'This is due to the way the New-CsRgsHoliday cmdlet processes date strings' -severity 3
  Write-Log -Message 'More information is available at the url below' -severity 3
  Write-Log -Message 'https://docs.microsoft.com/en-us/powershell/module/skype/new-csrgsholiday?view=skype-ps' -severity 3
  Write-Log -Message 'Your timezone appears to be misconfigured. This script may not function as expected' -severity 3
  If (!$Unattended)
  {
    #Skip the prompt in unattended mode
   
    Write-Log -Message 'The script will now prompt you to change regions. If you continue without changing regions I will output everything in US date format and hope for the best.' -severity 3

	
    #Prompt user to switch culture
    Write-Log -Message 'prompting user to change region'
    $title = 'Switch Windows Region?'
    $Message = 'Update the Windows Region (Culture) to en-AU? This is not required but is often overlooked when building Aussie servers'

    $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
    'Changes the Region Settings to en-AU and exits'

    $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
    'No, I like my date format, please convert the values.'

    $options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

    switch ($result)
    {
      0 
      {
        Write-Log -Message 'Updating System Culture' -severity 1
        Set-Culture -CultureInfo en-AU
        Write-Log -Message 'System Culture Updated, Script will exit.' -severity 3
        Write-Log -Message 'Close any PowerShell windows and run the script again' -severity 3
        Exit
      }
      1 
      {
        Write-Log -Message 'Unsupported Region. Setting compatability mode' -severity 3
        Start-Sleep -Seconds 5
        $ConvertTime = $true
      }
							
    }
  }
  Else #We are unattended and dont match the system culture. Assume US date format
  {
    Write-Log -Message 'Unsupported Region. Setting compatability mode' -severity 3
    Start-Sleep -Seconds 5
    $ConvertTime = $true
  }
} #End region check


Write-Log -Message 'Parsing command line parameters' -severity 2

# Detect and deal with null service ID

If ($ServiceID.length -eq 0) 
{
  Write-Log -Message 'No ServiceID entered, Searching for valid ServiceID' -severity 3
  Write-Log -Message 'Looking for Front End Pools' -severity 2
  $PoolNumber = ($Pools).count
  if ($PoolNumber -eq 0) 
  { 

    Write-Log -Message "Couldn't locate any FrontEnd Pools! Aborting script" -severity 3
    Throw "Couldn't locate RGS pool. Abort script"
  }

  if ($PoolNumber -eq 1) 
  { 
    Write-Log -Message "Only found 1 Front End Pool, $Pools.poolfqdn, Selecting it" -severity 2
    $RGSIDs = (Get-CsRgsConfiguration -Identity $Pools.PoolFqdn)
    $Poolfqdn = $Pools.poolfqdn
    #Prompt user to confirm
    Write-Log -Message "Found RGS Service ID $RGSIDs" -severity 1
    $title = 'Use this Front End Pool?'
    $Message = "Use the Response Group Server on $Poolfqdn ?"

    $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
    'Continues using the selected Front End Pool.'

    $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
    'Aborts the script.'

    $options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

    switch ($result)
    {
      0 
      {
        Write-Log -Message 'Updating ServiceID parameter' -severity 2
        $ServiceID = $RGSIDs.Identity.tostring()
        $FrontEndPool = $Pools.poolfqdn
      }
      1 
      {
        Write-Log -Message "Couldn't Autolocate RGS pool. Aborting script" -severity 3
        Throw "Couldn't Autolocate RGS pool. Abort script"
      }				
    }
  }
	

  Else 
  {
    #More than 1 Pool Detected and the user didnt specify anything
    Write-Log -Message "Found $PoolNumber Front End Pools" -severity 2
	
    If ($FrontEndPool.length -eq 0) 
    {
      Write-Log -Message 'Prompting user to select Front End Pool' -severity 1
      Write-Log -Message "Couldn't Locate ServiceID or PoolFQDN on the command line and more than one Front End Pool was detected" -severity 3
			
      #Menu code thanks to Greig.
      #First figure out the maximum width of the pools name (for the tabular menu):
      $width = 0
      foreach ($Pool in ($Pools)) 
      {
        if ($Pool.Poolfqdn.Length -gt $width) 
        {
          $width = $Pool.Poolfqdn.Length
        }
      }

      #Provide an on-screen menu of Front End Pools for the user to choose from:
      $index = 0
      Write-Host -Object ('Index  '), ('Pool FQDN'.Padright($width + 1), ' '), 'Site ID'
      foreach ($Pool in ($Pools)) 
      {
        Write-Host -Object ($index.ToString()).PadRight(7, ' '), ($Pool.Poolfqdn.Padright($width + 1), ' '), $Pool.siteid.ToString()
        $index++
      }
      $index--	#Undo that last increment
      Write-Host
      Write-Host -Object 'Choose the Front End Pool you wish to use'
      $chosen = Read-Host -Prompt 'Or any other value to quit'
      Write-Log -Message "User input $chosen" -severity 2
      if ($chosen -notmatch '^\d$') 
      {
        Exit
      }
      if ([int]$chosen -lt 0) 
      {
        Exit
      }
      if ([int]$chosen -gt $index) 
      {
        Exit
      }
      $FrontEndPool = $Pools[$chosen].PoolFqdn 
      $Poolfqdn = $FrontEndPool 
      $RGSIDs = (Get-CsRgsConfiguration -Identity $FrontEndPool) 
    }


    #We should have a Valid front end by now
		
    Write-Log -Message "Using Front End Pool $FrontEndPool" -severity 1
    $RGSIDs = (Get-CsRgsConfiguration -Identity $FrontEndPool)

    $Poolfqdn = $FrontEndPool


    if ($Unattended)
    {
    $ServiceID = $RGSIDs.Identity.tostring()
    }
    if (!$Unattended) 
    {
      #Prompt user to confirm
      $title = 'Use this Pool?'
      $Message = "Use the Response Group Server on $Poolfqdn ?"

      $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
      'Continues using the selected Front End Pool.'

      $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
      'Aborts the script.'

      $options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

      $result = $host.ui.PromptForChoice($title, $Message, $options, 0) 

      switch ($result)
      {
        0 
        {
          Write-Log -Message 'Updating ServiceID'  -severity 1
          $ServiceID = $RGSIDs.Identity.tostring()
        }
        1 
        {
          Write-Log -Message 'Couldnt Autolocate RGS pool. Abort script' -severity 3
          Throw 'Couldnt Autolocate RGS pool. Abort script'
        }
      }
    }
  }
}
Else
{
  Write-Log -Message 'ServiceID parameter detected, display warning' -severity 1
  Write-Log -Message 'The ServiceID parameter is being deprecated. Please use the FrontEndPool parameter instead.' -severity 3
  Write-Log -Message 'If you have Multiple Frontend Pools that use the same Holiday set names you MUST use the FrontEndPool parameter' -severity 3
  Write-Log -Message 'Pausing for 10 seconds' -severity 1
  start-sleep -Seconds 10
}

#We should have a valid service ID by now

#Check for last run flag.
Write-Log -Message 'Checking to see if we have been run on this pool before' -severity 2
Write-Log -Message 'If the script errors here, check all your SQL/FE servers are online and accessible' -severity 2

$Lastrun = (Get-CsRgsHolidaySet | Where-Object {$_.Name -like "_(dont use)_ Aussie*" -and $_.Ownerpool -like $FrontEndPool} -ErrorAction silentlycontinue)
If ($Lastrun) {
  Write-Log -Message 'Looks like we have been run on this pool before, thankyou' -severity 2
  Write-Log -Message "Found existing RGS Object $($lastrun.name)" -severity 1
}
    

Write-Log -Message 'Parsing CSV data' -severity 2
$states = 'ACT','NSW','NT','QLD','SA','TAS','VIC','WA'
foreach ($State in $states) 
{
  switch ($State) 
  { 
    'ACT' 
    {
      $StateName = ($RGSPrepend+'Australian Capital Territory')
      $StateID = 0
    }

    'NSW' 
    {
      $StateName = ($RGSPrepend+'New South Wales')
      $StateID = 1
    } 

    'NT' 
    {
      $StateName = ($RGSPrepend+'Northern Territory')
      $StateID = 2
    }  

    'QLD' 
    {
      $StateName = ($RGSPrepend+'Queensland')
      $StateID = 3
    } 

    'SA' 
    {
      $StateName = ($RGSPrepend+'South Australia')
      $StateID = 4
    } 

    'TAS' 
    {
      $StateName = ($RGSPrepend+'Tasmania') 
      $StateID = 5
    } 

    'VIC' 
    {
      $StateName = ($RGSPrepend+'Victoria')
      $StateID = 6
    } 

    'WA' 
    {
      $StateName = ($RGSPrepend+'Western Australia')
      $StateID = 7
    } 
  }
  

  Write-Log -Message "Processing events in $StateName" -severity 1
  #Find and clear the existing RGS Object
  Try 
  {
    Write-Log -Message "Checking for existing $StateName Holiday Set" -severity 2
    $holidayset = (Get-CsRgsHolidaySet -Name "$StateName" | Where-Object {$_.ownerpool -like $FrontEndPool})
    Write-Log -Message "Removing old entries from $StateName" -severity 2
    $holidayset.HolidayList.clear()
    Write-Log -Message "Existing entries from Holiday Set $StateName removed" -severity 2
  }
  Catch 
  {
    Write-Log -Message "Didnt find $StateName Holiday Set. Creating" -severity 2
    $PlaceholderDate = (New-CsRgsHoliday -StartDate '11/11/1970 12:00 AM' -EndDate '12/11/1970 12:00 AM' -Name 'Placeholder. Shouldnt Exist')
    $holidayset = (New-CsRgsHolidaySet -Parent $ServiceID -Name "$StateName" -HolidayList $PlaceholderDate -ErrorAction silentlycontinue)
    Write-Log -Message 'Removing Placeholder Date' -severity 2
    $holidayset.HolidayList.clear()            
  }
 
  #Process Events in that State
  foreach ($event in ($script:DateData | Where-Object {$_.Jurisdiction -eq $state}))
  {
    #Deal with Unix date format
    $udate = Get-Date -Date '1/1/1970'
    if ($ConvertTime) 
    {
      #American Date format
      $StartDate = ($udate.AddSeconds($event.'Raw Date').ToLocalTime() | Get-Date -Format MM/dd/yyyy)
      $EndDate = ($udate.AddSeconds(([int]$event.'Raw Date'+86400)).ToLocalTime() | Get-Date -Format MM/dd/yyyy)     
    }
    else 
    {
      #Aussie Date format
      $StartDate = ($udate.AddSeconds($event.'Raw Date').ToLocalTime() | Get-Date -Format dd/MM/yyyy)
      $EndDate = ($udate.AddSeconds(([int]$event.'Raw Date'+86400)).ToLocalTime() | Get-Date -Format dd/MM/yyyy)
    }

    #Create the event in Skype format
    $EventName = ($event.'Holiday Name')      
    $EventName = ($EventName -replace '  ' , ' ') #Remove Double Spaces in eventname
    $EventName = $EventName.Trim()		  #Remove any leading or trailing whitespace
    $CurrentEvent = (New-CsRgsHoliday -StartDate "$StartDate 12:00 AM" -EndDate "$EndDate 12:00 AM" -Name "$StateName $EventName")
    #$CurrentEvent
    #add it to the variable.
    Write-Log -Message "Adding $EventName to $StateName" -severity 2
    $holidayset.HolidayList.Add($CurrentEvent)
  }
  Write-Log -Message 'Finished adding events' -severity 2
  Write-Log -Message "Writing $StateName to Database" -severity 2
  Try 
  {
    Set-CsRgsHolidaySet -Instance $holidayset
    Write-Log -Message "Write OK!" -severity 2
  }

  Catch 
  {
    Write-Log -Message 'Something went wrong attempting to commit holidayset to database' -severity 3
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Write-Log -Message "$FailedItem failed. The error message was $ErrorMessage" -severity 4
    Throw $ErrorMessage
  }
}


#Okay, now deal with National Holidays

try 
{
  Write-Log -Message "Checking for existing $National Holiday Set" -severity 2
  $holidayset = (Get-CsRgsHolidaySet -Name "$National" | Where-Object {$_.ownerpool -like $FrontEndPool})
  Write-Log -Message "Removing old entries from $National" -severity 2
  $holidayset.HolidayList.clear()
  Write-Log -Message "Existing entries from Holiday Set $National removed" -severity 2
}
catch 
{
  Write-Log -Message "Didnt find $National Holiday Set. Creating" -severity 2
  $PlaceholderDate = (New-CsRgsHoliday -StartDate '11/11/1970 12:00 AM' -EndDate '12/11/1970 12:00 AM' -Name 'Placeholder. Shouldnt Exist')
  $holidayset = (New-CsRgsHolidaySet -Parent $ServiceID -Name "$National" -HolidayList $PlaceholderDate -ErrorAction silentlycontinue)
  Write-Log -Message 'Removing Placeholder Date' -severity 2
  $holidayset.HolidayList.clear()            
}

#Find dates that are in every state

Write-Log -Message 'Finding National Holidays (This can take a while)' -severity 2
$i = 0
$RawNatHolidayset = $null
$NatHolidayset = $null

$RawNatHolidayset = @()

foreach ($State in $XMLdata.ausgovEvents.jurisdiction) 
{
  #Process Events in that State
  $states = 'ACT','NSW','NT','QLD','SA','TAS','VIC','WA'
  foreach ($State in $states) 
  {
    $RawNatHolidayset += ($event)
  }
  $i ++
}

$NatHolidayset = ($script:DateData | Sort-Object -Property 'raw Date' -Unique)
ForEach($Uniquedate in $NatHolidayset)
{
  $SEARCH_RESULT = $script:DateData|Where-Object -FilterScript {$_.'raw Date' -eq $Uniquedate.'raw date'}

  if ( $SEARCH_RESULT.Count -eq 8)
  {      
    $event = ($SEARCH_RESULT | Select-Object -First 1)
                  

    #Deal with Unix date format
    $udate = Get-Date -Date '1/1/1970'
    if ($ConvertTime) 
    {
      #American Date format
      $StartDate = ($udate.AddSeconds($event.'Raw Date').ToLocalTime() | Get-Date -Format MM/dd/yyyy)
      $EndDate = ($udate.AddSeconds(([int]$event.'Raw Date'+86400)).ToLocalTime() | Get-Date -Format MM/dd/yyyy)     
    }
    else 
    {
      #Aussie Date format
      $StartDate = ($udate.AddSeconds($event.'Raw Date').ToLocalTime() | Get-Date -Format dd/MM/yyyy)
      $EndDate = ($udate.AddSeconds(([int]$event.'Raw Date'+86400)).ToLocalTime() | Get-Date -Format dd/MM/yyyy)
    }
                                 
    #Create the event in Skype format
    Write-Log -Message "Found $EventName" -severity 2
    $EventName = ($event.'Holiday Name')
    $EventName = ($EventName -replace '  ' , ' ') #Remove Double Spaces in eventname
    $EventName = $EventName.Trim()		  #Remove any leading or trailing whitespace
    $CurrentEvent = (New-CsRgsHoliday -StartDate "$StartDate 12:00 AM" -EndDate "$EndDate 12:00 AM" -Name "$StateName $EventName")
    $holidayset.HolidayList.Add($CurrentEvent)
  }
}
Write-Log -Message 'Finished adding events' -severity 2
Write-Log -Message "Writing $National to Database" -severity 2
Try 
{
  #Update Database  
  Set-CsRgsHolidaySet -Instance $holidayset
  Write-Log -Message "Write OK!" -severity 2

}
Catch 
{
  Write-Log -Message 'Something went wrong attempting to commit holidayset to database' -severity 3
  $ErrorMessage = $_.Exception.Message
  $FailedItem = $_.Exception.ItemName
  Write-Log -Message "$FailedItem failed. The error message was $ErrorMessage" -ForegroundColor Red
  Throw $ErrorMessage
}

#Update Last Run Flag
Write-Log -Message "Updating Last Run Holiday Set" -severity 2
if ($Lastrun) {
  Get-CsRgsHolidaySet | Where-Object {$_.Name -like "_(dont use)_ Aussie*" -and $_.Ownerpool -like $FrontEndPool} | Remove-CsRgsHolidaySet
}
#Write a new Last Run Flag
$ShortDate = (Get-Date -Format dd/MM/yyyy)  
$PlaceholderDate = (New-CsRgsHoliday -StartDate '11/11/1970 12:00 AM' -EndDate '12/11/1970 12:00 AM' -Name "_(Dont Use)_ Aussie Holidays updated on $ShortDate by New-CsRgsAustralianHolidayList http://bit.ly/CsRgsAU UcMadScientist")
[void](New-CsRgsHolidaySet -Parent $ServiceID -Name "_(Dont Use)_ Aussie Holidays updated on $ShortDate by New-CsRgsAustralianHolidayList http://bit.ly/CsRgsAU" -HolidayList $PlaceholderDate)
Write-Log -Message "Last Run Flag Updated" -severity 2

#Find display the last holiday imported

Write-Host ''
Write-Host ''

$LastDate = (($script:DateData | Sort-Object -Property 'raw Date' -Unique | select-Object -Last 1).date )
$LastDate = ([datetimeoffset]::parseexact($LastDate, 'yyyyMMdd', $null)).tostring("D") 
$FirstDate = (($script:DateData | Sort-Object -Property 'raw Date' -Unique | select-Object -First 1).date )
$FirstDate = ([datetimeoffset]::parseexact($FirstDate, 'yyyyMMdd', $null)).tostring("D") 
$ReRunDate = (($script:DateData | Sort-Object -Property 'raw Date' -Unique | select-Object -Last 1).date )
$ReRunDate = ([datetimeoffset]::parseexact($ReRunDate, 'yyyyMMdd', $null)).tostring("Y") 
Write-Log -Message 'Looks like everything went okay. Here are your current RGS Holiday Sets' -severity 2
Get-CsRgsHolidaySet | Select-Object -Property OwnerPool, Name | Format-Table

Write-Host ''
Write-Host ''
Write-Log -Message "Imported $($script:DateData.count) events between $Firstdate and $LastDate. You will need to re-run this script before $ReRunDate" -severity 2


Write-Host ''
Write-Host ''
Write-Host "Did this script help you? Save you some time? I'd REALLY appreciate it if you voted for it in the TechNet Gallery or if you tweeted about it" -ForegroundColor Cyan -BackgroundColor Black
Write-Host "It only takes a few moments and helps give me the tools to continue developing these tools for you!" -ForegroundColor Cyan -BackgroundColor Black
Write-Host "TechNet Gallery: https://gallery.technet.microsoft.com/Australian-RGSResponse-22845230?redir=0" -ForegroundColor Cyan -BackgroundColor Black
Write-Host "URL to Share: http://bit.ly/CsRgsAU" -ForegroundColor Cyan -BackgroundColor Black



# SIG # Begin signature block
# MIINFwYJKoZIhvcNAQcCoIINCDCCDQQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/BF448ryg955E+PffDxd0UAA
# kPCgggpZMIIFITCCBAmgAwIBAgIQD274plv3rQv2N1HXnqk5jzANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDEwNTAwMDAwMFoXDTIyMDky
# ODEyMDAwMFowXjELMAkGA1UEBhMCQVUxETAPBgNVBAgTCFZpY3RvcmlhMRAwDgYD
# VQQHEwdCZXJ3aWNrMRQwEgYDVQQKEwtKYW1lcyBBcmJlcjEUMBIGA1UEAxMLSmFt
# ZXMgQXJiZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCVq3KHhsUn
# G0iP8Xv+EIRGhPEqceUcmXftvbWSXoEL+w8h79PVn9WZawPgyDlmAZvzlAWaPGSu
# tW7z0/XqkewTjFI4em2BxIsLr3enoB/OuBM11ktZVaMYWOHaUexj8CioBeoFTGYg
# H98cmoo6i3xQcBbFJauJcgAI8jDTTDHM1bvDE9ItyeTr63MGJx1rob4KXCr0Oi9R
# MVtk/TDVCNjG3IdK8dnrpKUE7s2grAiPJ2tmNkrk3R2pSRl1qx3d01LWKcV2tv4s
# fbWLCwdz2HVTdevl7PjhwUPhuLZVj/EctCiU+5UDDtAIIIvQ9uvbFngmF0QmE9Yb
# W1bgiyfr5GmFAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg
# +S32ZXUOWDAdBgNVHQ4EFgQUX+77NtBOxF+2arVa8Srnig2A/ocwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0
# dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWg
# M6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcx
# LmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRw
# czovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEE
# eDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYB
# BQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJB
# c3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3
# DQEBCwUAA4IBAQCfGaBR90KcYBczv5tVSBquFD0rP4h7oEE8ik+EOJQituu3m/nv
# X+fG8h8f8+cG+0O55g+P/iGPS1Uo/BUEKjfUvLQjg9gJN7ZZozqP5xU7pn270rFd
# chmu/vkSGh4waYoASiqJXvkQbVZcxV72j3+RBD1jsmgP05WaKMT5l9VZwGedVn40
# FHNarFpJoCsyQn6sQInWdDfi6X2cYi0x4U0ogWYYyR8bhBUlt6RhevYn6EfqHgV3
# oEZ7qwxApjyGpQIwwQUEs60/tO7bkH1futFDdogzsXFJO3cS9OykctpBucaPDrkH
# 1AcqMqpWVRcXGebpOHnW5zPoGFG9JblyuwBZMIIFMDCCBBigAwIBAgIQBAkYG1/V
# u2Z1U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYD
# VQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAw
# WhcNMjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdp
# Q2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/
# 5aid2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH
# 03sjlOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxK
# hwjfDPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr
# /mzLfnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi
# 6CxR93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCC
# AckwEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAww
# CgYIKwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8v
# b2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6
# MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
# ZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1s
# AAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMw
# CgYIYIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1Ud
# IwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+
# 7A1aJLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbR
# knUPUbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7
# uq+1UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7
# qPjFEmifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPa
# s7CM1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR
# 6mhsRDKyZqHnGKSaZFHvMYICKDCCAiQCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QQIQD274plv3rQv2N1HXnqk5jzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEK
# MAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQURW2WYrorkDak10CC
# QXvhPn9il90wDQYJKoZIhvcNAQEBBQAEggEANQBUCRiSL17b8LQzpYEF9Rd1DlRK
# ie61zJm1sA/ABVO9HTKAnKzY61U9Mb48PHLrtyerobV8ueoI3X9UmAVbA/660M5k
# tpV2R0aIq9hwqC0teRG5n5o6ztwsyZLpDE6X+dxkFR+Zli67aw50Qavu8lAIhJbe
# imFY5nbBum6ic2svStc3MMr0D8QT3M9l4g8xFSvslXPJ9hHehCllBWBQnnlcI15x
# 5CehCUzytUfoSY1x2cc5uPOEW40WP+S9odgYJtF2AdXKs+ikupnac4gj/brUANUG
# DDPGflg6e0WIKvogR3YJmpxjSaMT9DSf/4UNfcz2d99jY9n7Mgwgmvj0tA==
# SIG # End signature block