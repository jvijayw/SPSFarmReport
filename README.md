# SPSFarmReport

SPSFarmReport is a scripted-tool that can be used to gather topology-related details from SharePoint farms. This tool has been for close to a decade. It was first written for MOSS 2007, and we've had this script out of CodePlex for every subsequent release of SharePoint including 2010, 2013, and now 2016!

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

For additional details on running the tool and usage, please refer to this nice [blog](https://blogs.technet.microsoft.com/brookswhite/2014/04/02/sps-farm-report/) written by Brooks White - Senior SharePoint PFE at Microsoft. *Follow the comments on that blog for FAQs*.

## Getting Started 
My sincere thanks to all the testers in Microsoft Support who've enabled the release of this tool. Special thanks to the below people:

+ __Ajith Jose__, *Support Escalation Engineer* from the *Microsoft Project Support team*. He coauthored and contributed code to gathering Project Server details.
+ __David Storey__, *Program Manager* from the *Microsoft Edge product team*. He contributed stylsheet code that'll open the report in Edge.
+ __Nirankush Panchbai__, *Principal Program Manager Lead* from the *Microsoft Edge product team*. He came first to offer help when I was struggling to get the <details / summary> tag to work in Edge.

--  J Vijay William
