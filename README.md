# SPSFarmReport

SPSFarmReport is a scripted-tool that can be used to gather topology-related details from SharePoint farms. This tool has been for close to a decade. It was first written for MOSS 2007, and we've had this script out of CodePlex for every subsequent release of SharePoint including 2010, 2013, 2016, 2019, and now 2022 !

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
Get it from https://github.com/jvijayw/SPSFarmReport/blob/master/SPSFarmReport.vNext.ps1.

--  J Vijay William
