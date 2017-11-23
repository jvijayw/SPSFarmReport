# SPSFarmReport

SPSFarmReport is a tool that can be used to gather topology-related details from SharePoint farms. This tool has been for close to a decade. It was first written for MOSS 2007, and we've had this script out of CodePlex for every subsequent release of SharePoint including 2010, 2013, and now 2016!

Roughly, we gather and build a report on the below infra of your farm:
+ Farm General Settings
	+ Central Admin URL
	+ Farm Build Version
	+ System Account
	+ Configuration Database details
+ Services on Server
	+ MinROLE and its Compliance
	+ Distributed Caching
+ Installed Products on Servers
+ Features Installed
+ Service Applications, Pools, and Proxies
	+ Search
	+ Project Services, etc.
+ Web Applications, AAMs
+ Content Databases
+ Content Deployment
+ Health Analyzer Reports
+ Timer Jobs and their Details

## Getting Started 
The tool comprises of a PowerShell Script and an XSL Stylesheet. You will also find an accompanying CSS. Running the PowerShell script is pretty straightfoward with no arguments. The output of that is written to an XML. Double-clicking the XML will open it in a browser. The XSL is HTML5 compliant and we recommend the use of __Microsoft Edge__ to view the report.

