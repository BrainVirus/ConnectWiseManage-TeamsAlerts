<#
    AUTHOR:
        Brandon Fahnestock, INTELITECHS LLC

    CONTRIBUTORS:
        

    ORIGINAL DATE:
        8/14/2020
         
    DESCRIPTION:
      Checks Manage for Tickets and/or Time entry errors in the last 14 days and posts them to Teams via incoming webhook.

    DIRECTIONS:
       Install-Module 'ConnectWiseManageAPI'
       Install-Module 'PSTeams'
       Setup scheduled task on a server to run this script as needed, or run manually.
       Update Varibles as needed.

       To create a new search 
            Copy an existing elseif and add it to the Get-Problem Function and modify search parameters as needed.
                !!!!IMPORTANT!!!
                if it's a time entry search, copy an elseif that has -Time at the end. I.E 'Billable-Time' or 'MissingAgreement-Time'
                if it's a ticket entry search, copy an elseif that does NOT have -Time at the end. I.E. 'MissingType' or 'CatchAll'
            Add another function call near the end of the script with all the other function calls
                Get-Problems -Type NewTimeSearch-Time
                Get-Problems -Type NewTicketSearch
        
    VERSION HISTORY
        2021.07.02 - Sanitized script for general use

        2021.08.27 - Updated URL for $ticketEntryURl to v2021_2

        2022.01.17 - Fixed page size for $AgreementsCompanyID

        2022.01.25 - Fix Missing Agreement to only report on active agreements 
                     Updated URL for $ticketEntryURl to v2021_3

        2022.02.09 - Add "-condition" to get-CWMTimeEntry uses. Was throwing errors when
                     ran without it when using the latest ConnectWiseManageAPI (0.4.7.0)
                     
        2022.04.22 - Updated URL for $ticketEntryURl to v2022_1
        
        2022.07.22 - Replaced all instances of -pageSize 1000 with -all

    TO DO:

#>

<######################################################################>
<# Script Variables #>
$CWMConnectionInfo = @{
  # This is the URL to your manage server.
  Server     = ''
  
  # This is the company entered at login
  Company    = ''
  
  # Public key created for this integration
  pubKey     = ''
  
  # Private key created for this integration
  privateKey = ''
  
  # Your ClientID found at https://developer.connectwise.com/ClientID
  clientId   = ''
}

#set this to true to have the script post to the test channel in teams
$testing = $false

#Create a webhook for teams. Follow this https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook
$TeamsIDUrlLive = 'URL GOES HERE'
$TeamsIDUrlTest = 'TEST URL GOES HERE'

$emailsuffix = '*@companyemail.com' #example '*@gmail.com' or '*@mspcompany.com'

#if link breaks, change "v2022_1" in the following link to the current version and use Tinyurl.com to shorten it: https://na.myconnectwise.net/v2022_1/ConnectWise.aspx?locale=en_US&routeTo=ServiceFV
#make sure variable has ?srRecID= after the tinyURL
$ticketEntryURl = 'https://tinyurl.com/4tvrwx7s?srRecID='

#if link breaks, check https://developer.connectwise.com to see what the new URL is for the version and use Tinyurl.com to shorten it 
#make sure variable has ?RecID= after the tinyURL
#Current is (fill in your own company name): https://api-na.myconnectwise.net/v4_6_release/services/system_io/router/openrecord.rails?locale=en_US&companyName=PUTYOURCOMPANYNAMEHERE&recordType=TimeEntryFV
$timeEntryURL = 'makeyourowntinyurlwithyourcompanynameandputithere?RecID='   

#Used to detect if there were no issues found in any function. Leave this as true.
$noIssuesReported = $true

<# /Script Variables #>
<######################################################################>
<# Modules #>
Import-Module 'ConnectWiseManageAPI'
Import-Module 'PSTeams'
<# /Modules #>
<######################################################################>
<# Functions #>
Function Get-Problems([parameter(Mandatory = $true)]$Type) {
  
  #Prebuild the teams message
  $script:Section = New-TeamsSection {
    New-TeamsList -Name '' {
    }
  }
  
  $date = Get-Date -Format "MM/dd/yyyy"
  $days = 14
  $ticketDate = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-$days)

  #Set Search, message title, and member search based on which search we are doing.
  if ($Type -eq 'MustChange-Time') {
    $ticketSearch = get-CWMTimeEntry -condition "workType/name = `"MUST CHANGE!!!!`" and status != `"Billed`" and dateEntered > [$ticketDate]" -all
    $messageTitle = "$date - Time Work Type is Must Change"
    $uniqueMembers = $ticketSearch.member.identifier | Sort-Object | Get-Unique
  }
  elseif ($Type -eq 'NonBillable-Time') {
    $ticketSearch = get-CWMTimeEntry -condition "(workType/name = `"Admin`" or workType/name = `"Communication`" or workType/name = `"PTO`" or workType/name = `"RMM Agent`" or workType/name = `"Streamline IT`" ) and billableoption = `"Billable`" and status != `"Billed`" and dateEntered > [$ticketDate]" -all
    #filter to IGNORE billable RMM agent Work types for specific clients. Works, just commented out unless you need it. Make sure to change the "Company Name"
    #$ticketSearch = $ticketSearch | Where-Object { !($_.company.identifier -like "COMPANY NAME" -and $_.workType.Name -like "RMM Agent") }
    $messageTitle = "$date - Non-Billable Time Flagged as Billable"
    $uniqueMembers = $ticketSearch.member.identifier | Sort-Object | Get-Unique
  }
  elseif ($Type -eq 'Billable-Time') {
    $ticketSearch = get-CWMTimeEntry -condition "(workType/name = `"After Hours Project`" or workType/name = `"Emergency/Weekend`" or workType/name = `"Onsite`" or workType/name = `"Remote`" or workType/name = `"Travel`")  and billableoption = `"DoNotBill`" and status != `"Billed`" and dateEntered > [$ticketDate]" -all
    $messageTitle = "$date - Billable Time Flagged as Non-Billable"
    $uniqueMembers = $ticketSearch.member.identifier | Sort-Object | Get-Unique
  }
  elseif ($Type -eq 'MissingAgreement-Time') {
    $AgreementsCompanyID = (Get-CWMAgreement -condition "agreementstatus = `"Active`"" -all).company.id | Sort-Object | Get-Unique 
    $ticketSearch = get-CWMTimeEntry -condition "agreement/name = null and status != `"Billed`" and dateEntered > [$ticketDate]" -all
    #filter to IGNORE entries for companies that don't have an agreement
    $ticketSearch = $ticketSearch | Where-Object { ($_.company.id -in $AgreementsCompanyID) }
    $messageTitle = "$date - Missing Agreement"
    $uniqueMembers = $ticketSearch.member.identifier | Sort-Object | Get-Unique
  }
  elseif ($Type -eq '15MinuteCommunicationAlert-Time') {
    $ticketSearch = get-CWMTimeEntry -condition "workType/name = `"Communication`" and actualHours >= 0.25 and status != `"Billed`" and dateEntered > [$ticketDate]" -all
    $messageTitle = "$date - Communication Time Entry Over 15 Minutes"
    $uniqueMembers = $ticketSearch.member.identifier | Sort-Object | Get-Unique
  }
  elseif ($Type -eq 'MissingType') {
    $ticketSearch = Get-CWMTicket -condition "status/Name = `">Closed`" and closedDate > [$ticketDate] and type/name = null and ParentTicketId = null" -all
    $messageTitle = "$date - Ticket Type Missing"
    $uniqueMembers = $ticketSearch.closedBy | Sort-Object | Get-Unique
  }
  elseif ($Type -eq 'CatchAll') {
    $ticketSearch = Get-CWMTicket -condition "status/Name = `">Closed`" and closedDate > [$ticketDate] and company/name = `"Catchall`" and ParentTicketId = null" -all
    $messageTitle = "$date - Company set as Catchall"
    $uniqueMembers = $ticketSearch.closedBy | Sort-Object | Get-Unique
  }
 
  #Boolean to trigger how the first username is written to the message body
  $initalMember = $true

  #check if we found any ticket/time issues
  if ($ticketSearch.length -eq 0) {
    #exit function if there is no issues.
    Write-Verbose "$Type nothing found"
    return
  }

  #if there are issues found, parse them out and build the Teams Message Body
  else {
    foreach ($member in $uniqueMembers) {
      #get full info on member to check if they are your employee
      $memberInfo = Get-CWMMember -Condition "identifier = `"$member`""
      $memberFullName = $memberInfo.firstName.ToString() + " " + $memberInfo.lastName.ToString()
      
      #Skip non-MSP User
      if ($memberInfo.officeEmail -like $emailsuffix) {
        #write the the member name to the message body
        if ($initalMember) {
          $script:Section.facts[0].value = "- $memberFullName"
          $initalMember = $false
        }
        else {
          $text = $script:Section.facts[0].value.ToString()
          $script:Section.facts[0].value = "$text" + "`n- $memberFullName"
        }
    
        #Search for tickets based on what search we are doing.
        if ($Type -like '*-Time') {
          $problemTickets = $ticketSearch | where-Object { $_.member.identifier -contains "$member" }
        }
        else {
          $problemTickets = $ticketSearch | where-Object { $_.closedBy -contains "$member" }
        }
        
        foreach ($problemTicket in $problemTickets) {
          #grab ticket ID based on what search we are doing
          if ($Type -like '*-Time') {
            $ticketID = $problemTicket.ticket.id.ToString()
            $timeEntryID = $problemTicket.id.ToString()
          }
          else {
            $ticketID = $problemTicket.id.ToString()
          }
       
          #Set link for ticket in message body
          $text = $script:Section.facts[0].value.ToString()
          if ($Type -like '*-Time') {
            $script:Section.facts[0].value = "$text" + "`n`t- [#$ticketID - $timeEntryID]($timeEntryURL$timeEntryID)"  # Content
          }
          else {
            $script:Section.facts[0].value = "$text" + "`n`t- [#$ticketID]($ticketEntryURl$ticketID)"  # Content
          }
        }
      }
    } 
  }
  #post message to teams if issues are found and not skipped. If the only issues that were detected were non-MScompany users, it will post a blank message if we don't check the length)
  if ($script:Section.facts.value.Length -gt 0) {
    Send-TeamsMessage -URI $TeamsID -MessageSummary 'Daily Alerts' -MessageTitle "$messageTitle" -Sections $script:Section
    $script:noIssuesReported = $false
  }
  #clear search results to prep for next search
  $ticketSearch = ""
}
<# /Functions #>
<######################################################################>

<# Main Script #>

# Connect to your Manage server
Connect-CWM @CWMConnectionInfo

#Check if testing, and set the teamsID variable to the corresponding URL
if ($testing) { $TeamsID = $TeamsIDUrlTest }
else { $TeamsID = $TeamsIDUrlLive }

#Run Searches
Get-Problems -Type NonBillable-Time
Get-Problems -Type Billable-Time
Get-Problems -Type MissingAgreement-Time
Get-Problems -Type MustChange-Time
Get-Problems -Type 15MinuteCommunicationAlert-Time
Get-Problems -Type Catchall
Get-Problems -Type MissingType

#If there were no issues, post a no issues found message.
if ($noIssuesReported) {
  $date = Get-Date -Format "MM/dd/yyyy"
  $messageTitle = "$date"
  $script:Section.facts[0].value = "## No issues found!!! Awesome!"
  Send-TeamsMessage -URI $TeamsID -MessageSummary 'Daily Alerts' -MessageTitle "$messageTitle" -Sections $script:Section
}