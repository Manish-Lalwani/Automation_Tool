# Automation_Tool
I created this tool to automate the manual process of closing the Tickets generated in ServiceNow


#Basic overview of the scenario for which I developed the tool.


Little details of Our Prod Envoirment :

There was one Monitoring tool PRTG which monitors server health.
This tool was configured on all servers using sensors for each component like eg : CPU usage, Ram usage ,Storage usage, 
also for services like ping, ssh etc.

When something went wrong PRTG generated an alert on it's dashboard.
This alert got converted into an ServiceNow Ticket/Incident.



Our work was to Manually open each Ticket check which service of which server gone down, open the same service of the server, 
check the status of the service if it's up the close the ticket and if not wait for some time and if still it doesn't get up route it to the Particular Team(Linux, Windows,Backup,etc).

So to tackle this hectic of manually closing ticket.
I created this Automation Tool.


Languages and Libraries used : 
Python , Selenium , REST API's


Input for Tool :
An Excel with details of ServiceNow Incident in an formatted way.


Working and output of tool :

First it loaded the Excel with the Incident details.
Took one Incident and checked the Alert details (servr name,service).

Then fetched the relevant details from PRTG.

2 ways for fetvhing details fro PRTG.

Using Selenium.
API CALLS.

Both ways have been implemented, First had implemented using seenium but it took much time, so later implemented using API calls which was much faster.


After fetching the relevant details checked if the service is up or not.
  If the service is down moved to the next Incident.
  But if the service was up then,
  
  As Api calls were restricted in our Prod Env, 
  So had to Implement using selenium.
  
  It logged into ServiceNow opened the relevant ticket.
  Updated the Incident's with resolve note of the PRTG Service status,
  Changed the Incident status as resolved and closed the Incident,
  
  The details of if the Incient was closed or not were outputted in a New Excel file bu the tool.
  
  
  This is the workflow and a basic Insight of the tool or the script.
  
  
  
