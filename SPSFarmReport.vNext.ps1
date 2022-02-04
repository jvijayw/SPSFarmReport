if ( (Get-PSSnapin -Name Microsoft.Sharepoint.PowerShell -ErrorAction SilentlyContinue) -eq $null )
{
    # Add-PsSnapin Microsoft.Sharepoint.PowerShell -ErrorAction SilentlyContinue
}

[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
[void][System.Reflection.Assembly]::Load("System.Web, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Publishing, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
[void][System.Reflection.Assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")

# Declaring bool variables
$isFoundationOnlyInstalled = $false

# Declaring all string variables
$DSN, $confgDbServerName, $confgDbName, $BuildVersion, $systemAccount = "","","","","" | Out-Null
$adminURL, $adminDbName, $adminDbServerName, $enterpriseSearchServiceStatus, $enterpriseSearchServiceJobDefs  = "","","","","" | Out-Null
$exceptionDetails, $htmlContents, $global:HTMLpath, $HTMLHeaders, $HTMLTrail = "" | Out-Null

# Declaing all integer variables
$Servernum, $totalContentDBCount, $WebAppnum, $serviceAppPoolCount, $FeatureCount = 0,0,0,0,0,0,0 | Out-Null
$wFeatureCount, $solutionCount, $sFeatureCount, $timerJobCount = 0,0,0,0 | Out-Null
$serviceAppProxyCount, $serviceAppProxyGroupCount, $searchsvcAppsCount  = 0, 0, 0 | Out-Null
$SvcAppCount = 0 | Out-Null

# Declaring all string[] arrays 
[System.string[]] $Servers, $ServersId. $WebAppAAMs | Out-Null
[System.string[]] $searchServiceAppIds | Out-Null
# remove this line man [System.string[]] $delimiterChars, $delimiterChars2, $delimiterChars3, $delimiterChars4 | Out-Null

# Declaring all string[,] arrays 
[System.string[,]] $ServicesOnServers, $WebAppDetails, $WebAppIISPath | Out-Null
[System.string[,]] $WebAppExtended, $WebAppAuthProviders | Out-Null
[System.string[,]] $ContentDBs, $ContentDBSitesNum, $solutionProps, $ContentDBProps | Out-Null
[System.string[,]] $FarmFeatures, $SiteFeatures, $WebFeatures | Out-Null

# Declaring three dimensional arrays
[System.string[,,]] $serverProducts | Out-Null

# Declaring PowerShell environment settings
$FormatEnumerationLimit = 25

# Declaring XML data variables
[System.Xml.XmlDocument]$XMLToParse | Out-Null
[System.Xml.XmlDocument]$global:CDGI | Out-Null
[System.Xml.XmlNode]$XmlNode | Out-Null

# Declaring all Hash Tables
$global:ServerRoles = @{}
$global:ServiceApps = @{}
$global:SearchHostControllers = @{}
$global:SearchActiveTopologyComponents = @{}
$global:SearchActiveTopologyComponentsStatus = @{}
$global:SearchConfigAdminComponents = @{}
$global:SearchConfigLinkStores = @{}
$global:SearchConfigCrawlDatabases = @{}
$global:SearchConfigCrawlRules = @{}
$global:SearchConfigQuerySiteSettings = @{}
$global:SearchConfigContentSources = @{}
$global:SPServiceApplicationPools = @{}
$global:SPServiceAppProxies = @{}
$global:SPServiceAppProxyGroups = @{}
$global:CDPaths = @{}
$global:CDJobs = @{}
$global:HealthReport0 = @{}
$global:HealthReport1 = @{}
$global:HealthReport2 = @{}
$global:HealthReport3 = @{}
$global:timerJobs = @{} 
$global:projectInstances = @{} 
$global:projectPCSSettings = @{} 
$global:projectQueueSettings = @{} 
$global:projectsvcApps= @{} 
$global:_DCacheContainers= @{}
$global:_DCacheHosts= @{}

# Declaring all Hard-Coded values
$global:_maxServicesOnServers = 75 # This value indicates the maximum number of services that run on each server.
$global:_maxProductsonServer = 15 # This value indicates the maximum number of Products installed on each server.
$global:_maxItemsonServer = 200 # This value indicates the maximum number of Items installed per Product on each server.
$global:_maxContentDBs = 141 # This is the maximum number of content databases we enumerate per web application.
$global:_serviceTypeswithNames = @{"Microsoft.Office.Server.Search.Administration.SearchQueryAndSiteSettingsService" = "Search Query and Site Settings Service" ; "Microsoft.Office.Server.ApplicationRegistry.SharedService.ApplicationRegistryService" = "Application Registry Service"} # This varialble is used to translate service names to friendly names.
$global:_farmFeatureDefinitions = @{"AccSrvApplication" = "Access Services Farm Feature"; # This varialble is used to feature definition names to friendly names.
"GlobalWebParts" = "Global Web Parts";
"VisioServer" = "Visio Web Access";
"SpellChecking" = "Spell Checking";
"SocialRibbonControl" = "Social Tags and Note Board Ribbon Controls";
"VisioProcessRepositoryFeatureStapling" = "Visio Process Repository";
"DownloadFromOfficeDotCom" = "Office.com Entry Points from SharePoint";
"ExcelServerWebPartStapler" = "Excel Services Application Web Part Farm Feature";
"DataConnectionLibraryStapling" = "Data Connection Library";
"FastFarmFeatureActivation" = "FAST Search Server 2010 for SharePoint Master Job Provisioning";
"ExcelServer" = "Excel Services Application View Farm Feature";
"ObaStaple" = "`"Connect to Office`" Ribbon Controls";
"TemplateDiscovery" = "Offline Synchronization for External Lists" }
$global:_logpath = [Environment]::CurrentDirectory + "\2016SPSFarmReport{0}{1:d2}{2:d2}-{3:d2}{4:d2}" -f (Get-Date).Day,(Get-Date).Month,(Get-Date).Year,(Get-Date).Second,(Get-Date).Millisecond + ".LOG"
$global:_DCacheContainerNames = @("DistributedAccessCache", "DistributedActivityFeedCache", "DistributedActivityFeedLMTCache", "DistributedBouncerCache", "DistributedDefaultCache", "DistributedFileLockThrottlerCache", "DistributedHealthScoreCache", "DistributedLogonTokenCache", "DistributedResourceTallyCache", "DistributedSecurityTrimmingCache", "DistributedServerToAppServerAccessTokenCache", "DistributedSharedWithUserCache", "DistributedUnifiedGroupsCache", "DistributedSearchCache" )

function global:PSUsing 
{    param 
    (        
        [System.IDisposable] $inputObject = $(throw "The parameter -inputObject is required."),        
        [ScriptBlock] $scriptBlock = $(throw "The parameter -scriptBlock is required.")    
     )     
     Try 
     {        
        &$scriptBlock    
     } 
     Finally 
     {        
        if ($inputObject -ne $null) 
        {            
            if ($inputObject.psbase -eq $null) {                $inputObject.Dispose()            } 
            else {                $inputObject.psbase.Dispose()            }        
        }    
    }
}

function o16farmConfig()
{
    try
    {
            
<#            $global:DSN = Get-ItemProperty "hklm:SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\16.0\Secure\configdb" | select dsn | Format-Table -HideTableHeaders | Out-String -Width 1024
            if (-not $?)
            {
                Write-Host You will need to run this program on a server where SharePoint is installed.
                exit
            }
            $global:DSN = out-string -InputObject $global:DSN
            $confgDbServerNameTemp = $global:DSN -split '[=;]' 
#>            $global:configDbServerName = (Get-SPDatabase | ? { $_.type -match "configuration" }).Server.Address
            $global:configDbName = (Get-SPDatabase | ? { $_.type -match "configuration" }).Name

            [Microsoft.SharePoint.Administration.SPFarm] $mySPFarm = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.Farm
            $global:BuildVersion = [string] $mySPFarm.BuildVersion
            $global:systemAccount = $mySPFarm.TimerService.ProcessIdentity.Username.ToString()
            [Microsoft.SharePoint.Administration.SPServerCollection] $mySPServerCollection = $mySPFarm.Servers
            [Microsoft.SharePoint.Administration.SPWebApplicationCollection] $mySPAdminWebAppCollection = [Microsoft.SharePoint.Administration.SPWebService]::AdministrationService.WebApplications
            [Microsoft.SharePoint.Administration.SPTimerService] $spts = $mySPFarm.TimerService   
                
                if ($mySPAdminWebAppCollection -ne $null)
                {
                    $mySPAdminWebApp = new-object Microsoft.SharePoint.Administration.SPAdministrationWebApplication
                    foreach($mySPAdminWebApp in $mySPAdminWebAppCollection)
                    {
                        if ($mySPAdminWebApp.IsAdministrationWebApplication)
                        {
                            $mySPAlternateUrl = new-object Microsoft.SharePoint.Administration.SPAlternateUrl
                            foreach ($mySPAlternateUrl in $mySPAdminWebApp.AlternateUrls)
                            {
                                switch ($mySPAlternateUrl.UrlZone.ToString().Trim())
                                {
                                    default
                                    {
                                        $global:adminURL = $mySPAlternateUrl.IncomingUrl.ToString()
                                    }
                                }
                            }
                            [Microsoft.SharePoint.Administration.SPContentDatabaseCollection] $mySPContentDBCollection = $mySPAdminWebApp.ContentDatabases;
                            $mySPContentDB = new-object Microsoft.SharePoint.Administration.SPContentDatabase
                            foreach ( $mySPContentDB in $mySPContentDBCollection)
                            {
                                $global:adminDbName = $mySPContentDB.Name.ToString()
                                $global:adminDbServerName = $mySPContentDB.Server.ToString()
                            }
                        }
                    }
                }
				
				$productsNum = ((($mySPFarm.Products | ft -HideTableHeaders | Out-String).Trim()).Split("`n")).Count
				if($productsNum -eq 1)
				{ $global:isFoundationOnlyInstalled = $true }
				return 1
    } 
    
    catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		global:HandleException("o15farmConfig", $_)
		return 0
    } 
}

function o16enumServers()
{
    try
    {
        [Microsoft.SharePoint.Administration.SPFarm] $mySPFarm = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.Farm
        [Microsoft.SharePoint.Administration.SPServerCollection] $mySPServerCollection = $mySPFarm.Servers
        
        #To get the number of servers in farm
        $global:Servernum = $mySPServerCollection.Count
        $global:ServicesOnServers = new-object 'System.String[,]' $global:Servernum, $global:_maxServicesOnServers 
        $global:Servers = new-object 'System.String[]' $global:Servernum
        $global:ServersId = new-object 'System.String[]' $global:Servernum
        $local:count, $ServicesCount, $count2 = 0,0,0
        [Microsoft.SharePoint.Administration.SPServer] $ServerInstance  
        foreach ($ServerInstance in $mySPServerCollection)
        {
            $tempstr = ""
            $count2 = 0
            $Servers[$count] = $ServerInstance.Address
            $ServersId[$count] = $ServerInstance.Id.ToString()
            $ServicesCount = $ServerInstance.ServiceInstances.Count
            $global:ServicesOnServers[$local:count, ($global:_maxServicesOnServers - 1)] = $ServerInstance.ServiceInstances.Count.ToString()
            foreach ($serviceInstance in $ServerInstance.ServiceInstances)
            {
                if(($serviceInstance.Hidden -eq $FALSE) -and ($serviceInstance.Status.ToString().Trim().ToLower() -eq "online"))    
                {
					if($global:_serviceTypeswithNames.ContainsKey($serviceInstance.Service.TypeName))
					{
						$ServicesOnServers[$count, $count2] = $global:_serviceTypeswithNames.Get_Item($serviceInstance.Service.TypeName)
					}
					else
					{
						$ServicesOnServers[$count, $count2] = $serviceInstance.Service.TypeName
					}
					$count2++
				}                    
             }
             $tempstr = $ServerInstance.Role.ToString()
             $tempstr += ","
             if ($ServerInstance.CompliantWithMinRole -eq $null ) { $tempstr += "Unknown - Check in Central Admin" }
             else { $tempstr += $ServerInstance.CompliantWithMinRole }
             $global:ServerRoles.Add($ServerInstance.Name, $tempstr)
			 $count = $count + 1             
        }                   
        return 1
    }
    catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		global:HandleException("o15enumServers", $_)
		return 0
    }    
}

<# function o16enumSPDcacheConfig()
{
	try
	{
        $count = 0
        Use-SPCacheCluster 
        $hostdetails = Get-CacheHost            
        if($hostdetails.Length -lt 2)
        {
                $global:XMLToParse = New-Object System.Xml.XmlDocument
                $global:XMLToParse = [xml] ($hostdetails[$count]  | ConvertTo-Xml -notypeinformation)
                $tempstr = [System.String]$global:XMLToParse.Objects.Object.InnerXml
                $global:_DCacheHosts.Add($hostdetails.Hostname, $tempstr)
        }
        else
        {     
                $hostdetails.GetEnumerator() | ForEach-Object {        
                $global:XMLToParse = New-Object System.Xml.XmlDocument
                $global:XMLToParse = [xml] ($hostdetails[$count]  | ConvertTo-Xml -notypeinformation)
                $tempstr = [System.String]$global:XMLToParse.Objects.Object.InnerXml
                $global:_DCacheHosts.Add($_.Hostname, $tempstr)
                $count++
                } 
        }

        ForEach($container in $global:_DCacheContainerNames)
        {
			$global:XMLToParse = New-Object System.Xml.XmlDocument
			$global:XMLToParse = [xml] (get-SPDistributedCacheClientSetting $container  | ConvertTo-Xml -notypeinformation)
			$tempstr = [System.String]$global:XMLToParse.Objects.Object.InnerXml
			$global:_DCacheContainers.Add($container, $tempstr)
		}	
		return 1
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSPServiceApplicationPools", $_)
		return 0
    }
} #>

function o16enumProdVersions()
{
    try
    {
        $global:serverProducts = new-object 'System.String[,,]' $global:Servernum, $global:_maxProductsonServer, $global:_maxItemsonServer
        $count = $global:Servernum - 1
        $count2, $count3 = 0,0

        [Microsoft.SharePoint.Administration.SPProductVersions] $versions = [Microsoft.SharePoint.Administration.SPProductVersions]::GetProductVersions()
        $infos = New-Object 'System.Collections.Generic.List[Microsoft.SharePoint.Administration.SPServerProductInfo]' (,$versions.ServerInformation)
        
        foreach ($prodInfo in $infos)
        {
            $count2 = 0;
            $count3 = 0;
            $products = New-Object 'System.Collections.Generic.List[System.String]' (,$prodInfo.Products)
            $products.Sort()
            $global:serverProducts[$count, $count2, $count3] = $prodInfo.ServerName
            foreach ($str in $products)
            {
                $count2++
                $serverProducts[$count, $count2, $count3] = $str
                $singleProductInfo = $prodInfo.GetSingleProductInfo($str)
                $patchableUnitDisplayNames = New-Object 'System.Collections.Generic.List[System.String]' (,$singleProductInfo.PatchableUnitDisplayNames)
                $patchableUnitDisplayNames.Sort()
                foreach ($str2 in $patchableUnitDisplayNames)
                {
                    $patchableUnitInfoByDisplayName = New-Object 'System.Collections.Generic.List[Microsoft.SharePoint.Administration.SPPatchableUnitInfo]' (,$singleProductInfo.GetPatchableUnitInfoByDisplayName($str2))
                    foreach ($info in $patchableUnitInfoByDisplayName)
                    {
                        $count3++;
                        $version = [Microsoft.SharePoint.Utilities.SPHttpUtility]::HtmlEncode($info.BaseVersionOnServer($prodInfo.ServerId).ToString())
                        $serverProducts[$count, $count2, $count3] = $info.DisplayName + " : " + $info.LatestPatchOnServer($prodInfo.ServerId).Version.ToString() 
                     }
                     $serverProducts[$count, $count2, ($global:_maxItemsonServer - 1)] = $count3.ToString()
                }
                $serverProducts[$count, ($global:_maxProductsonServer - 1), ($_maxItemsonServer - 1)] = $count2.ToString()
                $count3 = 0
            }
            $count--
        }
        return 1
    }
    catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumProdVersions", $_)
		return 0
    }
}

function o16enumFeatures()
{
            try
            {
                $bindingFlags = [System.Reflection.BindingFlags] “NonPublic,Instance”
                [Microsoft.SharePoint.Administration.SPFarm] $mySPFarm = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.Farm
                $global:FeatureCount = 0
                $FeatureCount2 = 0
                #PropertyInfo pi;

                #to retrieve the number of features deployed in farm
                foreach ($FeatureDefinition in $mySPFarm.FeatureDefinitions)
                {
                    if (($FeatureDefinition.Hidden.ToString() -ne "true") -and ($FeatureDefinition.Scope.ToString() -eq "Farm"))
                    {
                        $global:FeatureCount++
                    }
                }
                $global:FarmFeatures = new-object 'System.String[,]' $global:FeatureCount, 4;

                #to retrieve the properties
                foreach ($FeatureDefinition in $mySPFarm.FeatureDefinitions)
                {
                    if (($FeatureDefinition.Hidden.ToString() -ne "true") -and ($FeatureDefinition.Scope.ToString() -eq "Farm"))
                    {
						if($global:_farmFeatureDefinitions.ContainsKey($FeatureDefinition.DisplayName))
						{ $global:FarmFeatures[$FeatureCount2, 1] = $global:_farmFeatureDefinitions.Get_Item($FeatureDefinition.DisplayName)	}
						else {	$FarmFeatures[$FeatureCount2, 1] = $FeatureDefinition.DisplayName }
					
						$FarmFeatures[$FeatureCount2, 0] = $FeatureDefinition.Id.ToString()
						$FarmFeatures[$FeatureCount2, 2] = $FeatureDefinition.SolutionId.ToString()
						$pi = $FeatureDefinition.GetType().GetProperty("HasActivations", $bindingFlags)
						$FarmFeatures[$FeatureCount2, 3] = $pi.GetValue($FeatureDefinition, $null).ToString()				
                        $FeatureCount2++
                    }
                }

                $global:sFeatureCount = 0
                $FeatureCount2 = 0
                foreach ($FeatureDefinition in $mySPFarm.FeatureDefinitions)
                {
                    if (($FeatureDefinition.Hidden.ToString() -ne "true") -and ($FeatureDefinition.Scope.ToString() -eq "Site"))
                    {
                        $global:sFeatureCount++
                    }
                }
                $global:SiteFeatures = new-object 'System.String[,]' $global:sFeatureCount, 4
                foreach ($FeatureDefinition in $mySPFarm.FeatureDefinitions)
                {
                    if (($FeatureDefinition.Hidden.ToString() -ne "true") -and ($FeatureDefinition.Scope.ToString() -eq "Site"))
                    {
                        $global:SiteFeatures[$FeatureCount2, 0] = $FeatureDefinition.Id.ToString()
                        $SiteFeatures[$FeatureCount2, 1] = $FeatureDefinition.DisplayName
                        $SiteFeatures[$FeatureCount2, 2] = $FeatureDefinition.SolutionId.ToString()
                        $pi = $FeatureDefinition.GetType().GetProperty("HasActivations", $bindingFlags)
                        $SiteFeatures[$FeatureCount2, 3] = $pi.GetValue($FeatureDefinition, $null).ToString()
                        $FeatureCount2++
                    }
                }

                $global:wFeatureCount = 0
                $FeatureCount2 = 0
                foreach ($FeatureDefinition in $mySPFarm.FeatureDefinitions)
                {
                    if (($FeatureDefinition.Hidden.ToString() -ne "true") -and ($FeatureDefinition.Scope.ToString() -eq "Web"))
                    {
                        $global:wFeatureCount++
                    }
                }
                $global:WebFeatures = new-object 'System.String[,]' $global:wFeatureCount, 4
                foreach ($FeatureDefinition in $mySPFarm.FeatureDefinitions)
                {
                    if (($FeatureDefinition.Hidden.ToString() -ne "true") -and ($FeatureDefinition.Scope.ToString() -eq "Web"))
                    {
                        $WebFeatures[$FeatureCount2, 0] = $FeatureDefinition.Id.ToString()
                        $WebFeatures[$FeatureCount2, 1] = $FeatureDefinition.DisplayName
                        $WebFeatures[$FeatureCount2, 2] = $FeatureDefinition.SolutionId.ToString()
                        $pi = $FeatureDefinition.GetType().GetProperty("HasActivations", $bindingFlags)
                        $WebFeatures[$FeatureCount2, 3] = $pi.GetValue($FeatureDefinition, $null).ToString()
                        $FeatureCount2++;
                    }
                }
        return 1
    }
    catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumFeatures", $_)
		return 0
    }
}

function o16enumSolutions()
{
	try
	{
		$global:solutionCount = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.Farm.Solutions.Count
        $global:solutionProps = New-Object 'System.String[,]' $global:solutionCount, 6
        $count = 0
        foreach ($solution in [Microsoft.SharePoint.Administration.SPWebService]::ContentService.Farm.Solutions)
        {
            $global:solutionProps[$count, 0] = $solution.DisplayName
            $global:solutionProps[$count, 1] = $solution.Deployed.ToString()
            $global:solutionProps[$count, 2] = $solution.LastOperationDetails
            $global:solutionProps[$count, 5] = $solution.Id.ToString()

            foreach ($deployedServer in $solution.DeployedServers)
            {
                if ($global:solutionProps[$count, 3] -eq $null)
                {
                    if ($deployedServer.Address -eq $null)
					{ $global:solutionProps[$count, 3] = "" }
                    else
                        { $global:solutionProps[$count, 3] = $deployedServer.Address }
                }
                else
                    { $global:solutionProps[$count, 3] = $global:solutionProps[$count, 3] + "<br>" + $deployedServer.Address }
            }
            $count = $count + 1
        }
		return 1
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSolutions", $_)
		return 0
    }
}

function o16enumSvcApps() 
{
    try
    {                
		$global:SvcAppCount = (Get-SPServiceApplication).Length

		$svcApps = Get-SPServiceApplication | Select Id | Out-String -Width 1000
		$delimitLines =  $svcApps.Split("`n")
	
        ForEach($ServiceAppID in $delimitLines)
        {
			$ServiceAppID = $ServiceAppID.Trim()
			if (($ServiceAppID -eq "") -or ($ServiceAppID -eq "Id") -or ($ServiceAppID -eq "--")) { continue }
			$global:XMLToParse = New-Object System.Xml.XmlDocument
			$global:XMLToParse = [xml](Get-SPServiceApplication | where {$_.Id -eq $ServiceAppID} | ConvertTo-XML -NoTypeInformation)
			
			$typeName = $global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "TypeName" } 
			if($typeName -eq $null)
			{
				$tempstr = ($global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "Name" }).InnerText
			}
			else
			{
				$tempstr = ($global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "TypeName" }).InnerText
			}
						
			$ServiceAppID = $ServiceAppID + "|" + $tempstr
			$tempstr2 = [System.String]$global:XMLToParse.Objects.Object.InnerXml
			$global:ServiceApps.Add($ServiceAppID, $tempstr2)
        }
		return 1
    }
    catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSvcApps", $_)
		return 0
    }
}

function o16enumSPSearchServiceApps()
{
	try
	{
		$searchsvcApps = Get-SPServiceApplication | where {$_.typename -eq "Search Service Application"} | select Id | fl | Out-String -Width 1000
		$global:searchsvcAppsCount = 0
		$delimitLines = $searchsvcApps.Trim().Split("`n")
		ForEach($Liner in $delimitLines) 
		{	
			if($liner.Trim().Length -eq 0) { continue } 
			$global:searchsvcAppsCount++	
		}
		$global:searchServiceAppIds = new-object 'System.String[]' $global:searchsvcAppsCount
		$x = $global:searchsvcAppsCount - 1 
        ForEach($Liner in $delimitLines)
        {
            $Liner = $Liner.Trim()
            if($Liner.Length -eq 0)
            { continue }
            if($Liner.Contains("Id"))
            {
				$tempstr = $Liner -split " : "
                $global:searchServiceAppIds[$x] = $tempstr[1]   
				$x--
				if($x -lt 0)
				{ break }
            }
        }
		return 1 | Out-Null
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSPSearchServiceApps", $_)
		return 0
    }		
}

function o16enumProjectServiceApps()
{
	try
	{
        $global:projectsvcApps = Get-SPServiceApplication | where {$_.typename -eq "Project Application Services"} 
        $instcount=(Get-SPProjectWebInstance).Length
		$prjInst = Get-SPProjectWebInstance | Select Id | Out-String -Width 1000
		$delimitLines =  $prjInst.Split("`n")
	    $global:projectPCSSettings = Get-SPProjectPCSSettings
        $global:projectQueueSettings=Get-SPProjectQueueSettings
        ForEach($instID in $delimitLines)
        {
			$instID = $instID.Trim()
			if (($instID -eq "") -or ($instID -eq "Id") -or ($instID -eq "--")) { continue }
			$global:XMLToParse = New-Object System.Xml.XmlDocument
			$global:XMLToParse = [xml](Get-SPProjectWebInstance | where {$_.Id -eq $instID} | ConvertTo-XML -NoTypeInformation)
			
			$typeName = $global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "TypeName" } 
			if($typeName -eq $null)
			{
				$tempstr = ($global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "Name" }).InnerText
			}
			else
			{
				$tempstr = ($global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "TypeName" }).InnerText
			}
						
			$instID = $instID + "|" + $tempstr
			$tempstr2 = [System.String]$global:XMLToParse.Objects.Object.InnerXml
			$global:projectInstances.Add($instID, $tempstr2)
        }
		return 1 | Out-Null
		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o16enumProjectServiceApps", $_)
		return 0
    }		
}

function o16enumSPSearchService()
{
	try
	{
		$searchsvc = Get-SPEnterpriseSearchService 
		$_ato = ($searchsvc.AcknowledgementTimeout).ToString()
		$_cto = ($searchsvc.ConnectionTimeout).ToString()
		$_wproxy = ($searchsvc.WebProxy).ToString()
		$_ucpf = ($searchsvc.UseCrawlProxyForFederation).ToString()
		$_pl = ($searchsvc.PerformanceLevel).ToString()
		$_pi = ($searchsvc.ProcessIdentity).ToString()

		foreach($jd in $searchsvc.JobDefinitions)
		{
			#$jdx = ($jd | Select NAme, Schedule, LastRunTime, Server | ConvertTo-Xml -NoTypeInformation)
			$jd_Name = $jd | Select Name | ft -HideTableHeaders | Out-String
			$jd_Schedule = $jd | Select Schedule | ft -HideTableHeaders | Out-String
			$jd_LastRunTime = $jd | Select LastRunTime | ft -HideTableHeaders | Out-String
			$jd_Server = $jd | Select Server | ft -HideTableHeaders | Out-String
			
			$jd_Name = $jd_Name.Trim()
			$jd_Schedule = $jd_Schedule.Trim()
			$jd_LastRunTime = $jd_LastRunTime.Trim()
			$jd_Server = $jd_Server.Trim()
			
			$ejob = $ejob + "<job Name=`"" + $jd_Name + "`" Schedule=`"" + $jd_Schedule + "`" LastRunTime=`"" + $jd_LastRunTime +  "`" Server=`""  + $jd_Server + "`" />"
		}
		$tempXML = "<Property Name = `"AcknowledgementTimeout`">" + $_ato + "</Property>" + "<Property Name =`"ConnectionTimeout`">" + $_cto + "</Property>" + "<Property Name =`"WebProxy`">" + $_wproxy + "</Property>" + "<Property Name =`"UseCrawlProxyForFederation`">" + $_ucpf + "</Property>" + "<Property Name =`"PerformanceLevel`">" + $_pl + "</Property>" + "<Property Name =`"ProcessIdentity`">" + $_pi + "</Property>"
		
		$global:enterpriseSearchServiceStatus = $tempXML 
		$global:enterpriseSearchServiceJobDefs = $ejob 
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSPSearchServiceApps", $_)
		return 0
    }		
}

function o16enumSearchActiveTopologies()
{
	try
	{
		if($global:searchsvcAppsCount -eq 0)
		{ 		return 		}
		
		for ($tempCnt = 0; $tempCnt -lt $global:searchsvcAppsCount ; $tempCnt ++)	
		{
			$esa = Get-SPEnterpriseSearchServiceApplication -Identity $global:searchServiceAppIds[$tempCnt]			
			$searchServiceAppID = $searchServiceAppIds[$tempCnt]
			$searchSatus = Get-SPEnterpriseSearchStatus -SearchApplication $searchServiceAppID -ErrorAction SilentlyContinue			
				$ATComponentNames = $esa.ActiveTopology.GetComponents() | Select Name | ft -HideTableHeaders | Out-String -Width 1000
				$ATComponentNames = $ATComponentNames.Trim().Split("`n")			
			for($i = 0; $i -lt $ATComponentNames.Length ; $i++)
			{
				$tempXML = [xml] ($esa.ActiveTopology.GetComponents() | where {$_.Name -eq $ATComponentNames[$i].Trim() } | ConvertTo-Xml -NoTypeInformation)
				if($searchSatus -ne $null)
				{
					$tempXML2 = [xml] (Get-SPEnterpriseSearchStatus -SearchApplication $searchServiceAppID | ? {$_.Name -eq $ATComponentNames[$i].Trim()} | select State | ConvertTo-Xml -NoTypeInformation)
					$tempXML3 = [xml] (Get-SPEnterpriseSearchStatus -SearchApplication $searchServiceAppID | ? {$_.Name -eq $ATComponentNames[$i].Trim()} | select Details | ConvertTo-Xml -NoTypeInformation)
				}
				
				$tempstr = [System.String] $tempXML.Objects.Object.InnerXML
				$tempstr2 = [System.String] $tempXML2.Objects.Object.InnerXML
				$tempstr3 = [System.String] $tempXML3.Objects.Object.InnerXML
				$tempstr4 = $searchServiceAppID + "|" + $ATComponentNames[$i]	
				$tempstr = $tempstr + $tempstr2 + $tempstr3
				$global:SearchActiveTopologyComponents.Add($tempstr4, $tempstr)
			}
		}		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSearchActiveTopologies", $_)
    }
}

function o16enumHostControllers()
{
	try
	{
		if($global:searchsvcAppsCount -eq 0)
		{ 		return 		}

		for ($tempCnt = 0; $tempCnt -lt $global:searchsvcAppsCount ; $tempCnt ++)	
		{
			$cmdstr = Get-SPEnterpriseSearchHostController | Select Server | ft -HideTableHeaders | Out-String -Width 1000
			$cmdstr = $cmdstr.Trim().Split("`n")
			
			for($i = 0; $i -lt $cmdstr.Length ; $i++)
			{
				$cmdstr2 = $cmdstr[$i].Trim() 
				$searchServiceAppID = $searchServiceAppIds[$tempCnt]
				$tempXML = [xml] ( Get-SPEnterpriseSearchHostController | where {$_.Server -match $cmdstr2 } | ConvertTo-Xml -NoTypeInformation)
				$tempstr = [System.String] $tempXML.Objects.Object.InnerXML
				$searchServiceAppID = $searchServiceAppID + "|" + $cmdstr2				 
				$global:SearchHostControllers.Add($searchServiceAppID, $tempstr)
			}			
		}		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		global:HandleException("o15enumHostControllers", $_)
		return 0
    } 
}

function o16enumSearchConfigAdminComponents()
{
	try
	{
		if($global:searchsvcAppsCount -eq 0)
		{ 		return 		}
		for ($tempCnt = 0; $tempCnt -lt $global:searchsvcAppsCount ; $tempCnt ++)	
		{
			$searchServiceAppID = $searchServiceAppIds[$tempCnt]			
			$tempXML = [xml] (Get-SPEnterpriseSearchAdministrationComponent -SearchApplication $searchServiceAppID | ConvertTo-Xml -NoTypeInformation )
			$tempstr = [System.String] $tempXML.Objects.Object.InnerXML
			$global:SearchConfigAdminComponents.Add($searchServiceAppID, $tempstr )
		}		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSearchConfigAdminComponents", $_)
    }
}

function o16enumSearchConfigLinkStores()
{
	try
	{
		if($global:searchsvcAppsCount -eq 0)
		{ 		return 		}
		for ($tempCnt = 0; $tempCnt -lt $global:searchsvcAppsCount ; $tempCnt ++)	
		{
			$searchServiceAppID = $searchServiceAppIds[$tempCnt]			
			$ssa = Get-SPEnterpriseSearchServiceApplication -Identity $searchServiceAppID
			$tempXML = [xml] ($ssa.LinksStores | ConvertTo-Xml -NoTypeInformation )
			$tempstr = [System.String] $tempXML.Objects.Object.InnerXML
			$global:SearchConfigLinkStores.Add($searchServiceAppID, $tempstr )
		}		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSearchConfigLinkStores", $_)
    }
}

function o16enumSearchConfigCrawlDatabases() 
{
	try
	{
		if($global:searchsvcAppsCount -eq 0)
		{ 		return 		}

		for ($tempCnt = 0; $tempCnt -lt $global:searchsvcAppsCount ; $tempCnt ++)	
		{
			$crawlDatabasesPerSearchApp = Get-SPEnterpriseSearchCrawlDatabase -SearchApplication $global:searchServiceAppIds[$tempCnt] | Select Id | ft -HideTableHeaders | Out-String -Width 1000
			$crawlDatabasesPerSearchApp = $crawlDatabasesPerSearchApp.Trim().Split("`n")
			for($i = 0; $i -lt $crawlDatabasesPerSearchApp.Length ; $i++)
			{
				$searchServiceAppID = $searchServiceAppIds[$tempCnt]
				$tempXML = [xml] (Get-SPEnterpriseSearchCrawlDatabase -SearchApplication $global:searchServiceAppIds[$tempCnt] | where {$_.Id -eq $crawlDatabasesPerSearchApp[$i] } | ConvertTo-Xml -NoTypeInformation)
				$tempstr = [System.String] $tempXML.Objects.Object.InnerXML
				$searchServiceAppID = $searchServiceAppID + "|" + $crawlDatabasesPerSearchApp[$i]				 
				$global:SearchConfigCrawlDatabases.Add($searchServiceAppID, $tempstr)
			}			
		}		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSearchConfigCrawlDatabases", $_)
    }
}

function o16enumSearchConfigCrawlRules() 
{
	try
	{
		if($global:searchsvcAppsCount -eq 0)
		{ 		return 		}
		
		for ($tempCnt = 0; $tempCnt -lt $global:searchsvcAppsCount ; $tempCnt ++)	
		{
			$CrawlRuleNames = Get-SPEnterpriseSearchCrawlRule -SearchApplication $global:searchServiceAppIds[$tempCnt] | Select AccountName | ft -HideTableHeaders | Out-String -Width 1000
			$CrawlRuleNames = $CrawlRuleNames.Trim().Split("`n")
			for($i = 0; $i -lt $CrawlRuleNames.Length ; $i++)
			{
				$searchServiceAppID = $searchServiceAppIds[$tempCnt]
				$tempXML = [xml] (Get-SPEnterpriseSearchCrawlRule -SearchApplication $global:searchServiceAppIds[$tempCnt] | ? {$_.AccountName -eq $CrawlRuleNames[$i]}| ConvertTo-Xml -NoTypeInformation)
				$tempstr = [System.String] $tempXML.Objects.Object.InnerXML
				$searchServiceAppID = $searchServiceAppID + "|" + $CrawlRuleNames[$i]					
				$global:SearchConfigCrawlRules.Add($searchServiceAppID, $tempstr)
			}
		}		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSearchConfigCrawlRules", $_)
    }
}

function o16enumSearchConfigQuerySiteSettings() 
{
	try
	{
		if($global:searchsvcAppsCount -eq 0)
		{ 		return 		}
		
		for ($tempCnt = 0; $tempCnt -lt $global:searchsvcAppsCount ; $tempCnt ++)	
		{
			$querySiteSettingsId = Get-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance | ? {$_.status -ne "Disabled"} | Select Id | ft -HideTableHeaders | Out-String -Width 1000
			$querySiteSettingsId = $querySiteSettingsId.Trim().Split("`n")
			for($i = 0; $i -lt $querySiteSettingsId.Length ; $i++)
			{
				$searchServiceAppID = $searchServiceAppIds[$tempCnt]
				$tempXML = [xml] (Get-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance | ? {$_.status -ne "Disabled"} | where {$_.Id -eq $querySiteSettingsId[$i] } | ConvertTo-Xml -NoTypeInformation)
				$tempstr = [System.String] $tempXML.Objects.Object.InnerXML
				$searchServiceAppID = $searchServiceAppID + "|" + $querySiteSettingsId[$i]				 
				$global:SearchConfigQuerySiteSettings.Add($searchServiceAppID, $tempstr)
			}
		}		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSearchConfigQuerySiteSettings", $_)
    }
}

function o16enumSearchConfigContentSources()
{
	try
	{
		if($global:searchsvcAppsCount -eq 0)
		{ 		return 		}

		for ($tempCnt = 0; $tempCnt -lt $global:searchsvcAppsCount ; $tempCnt ++)	
		{
			$cmdstr = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $searchServiceAppIds[$tempCnt] | Select Id | ft -HideTableHeaders | Out-String -Width 1000
			$cmdstr = $cmdstr.Trim().Split("`n")
			
			for($i = 0; $i -lt $cmdstr.Length ; $i++)
			{
				$cmdstr2 = $cmdstr[$i].Trim() 
				$searchServiceAppID = $searchServiceAppIds[$tempCnt]
				$tempXML = [xml] (Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $searchServiceAppIds[$tempCnt] | select Name, Type, DeleteCount, ErrorCount, SuccessCount, WarningCount, StartAddresses, Id, CrawlStatus, CrawlStarted, CrawlCompleted, CrawlState | where {$_.Id -eq $cmdstr2 } | ConvertTo-Xml -NoTypeInformation)
				$tempstr = [System.String] $tempXML.Objects.Object.InnerXML
				$searchServiceAppID = $searchServiceAppID + "|" + $cmdstr2				 
				$global:SearchConfigContentSources.Add($searchServiceAppID, $tempstr)
			}			
		}		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSearchConfigContentSources", $_)
    }
}

function o16enumSPServiceApplicationPools()
{
	try
	{
		$svcAppPoolIDs = Get-SPServiceApplicationPool | select Id | Out-String -Width 1000
		$delimitLines =  $svcAppPoolIDs.Split("`n")
		$global:serviceAppPoolCount = (Get-SPServiceApplicationPool).Length		
		
        ForEach($svcAppPoolID in $delimitLines)
        {
			$svcAppPoolID = $svcAppPoolID.Trim()
			if (($svcAppPoolID -eq "") -or ($svcAppPoolID -eq "Id") -or ($svcAppPoolID -eq "--")) { continue }
			
			$global:XMLToParse = New-Object System.Xml.XmlDocument
			$global:XMLToParse = [xml](Get-SPServiceApplicationPool | select Id, Name, ProcessAccountName | where {$_.Id -eq $svcAppPoolID} | select Name, ProcessAccountName | ConvertTo-XML -NoTypeInformation)
			$tempstr = [System.String]$global:XMLToParse.Objects.Object.InnerXml
			$global:SPServiceApplicationPools.Add($svcAppPoolID, $tempstr)
		}	
		return 1
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSPServiceApplicationPools", $_)
		return 0
    }
}

function o16enumSPServiceApplicationProxies()
{
	try
	{
		$global:serviceAppProxyCount = (Get-SPServiceApplicationProxy).Length
		$svcApps = Get-SPServiceApplicationProxy | Select Id | Out-String -Width 1000
		$delimitLines =  $svcApps.Split("`n")
		
		ForEach($ServiceAppProxyID in $delimitLines)
        {
			$ServiceAppProxyID = $ServiceAppProxyID.Trim()
			if (($ServiceAppProxyID -eq "") -or ($ServiceAppProxyID -eq "Id") -or ($ServiceAppProxyID -eq "--")) { continue }
			$global:XMLToParse = New-Object System.Xml.XmlDocument
			$global:XMLToParse = [xml](Get-SPServiceApplicationProxy | where {$_.Id -eq $ServiceAppProxyID} | ConvertTo-XML -NoTypeInformation)
			
			$typeName = $global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "TypeName" } 
			if($typeName -eq $null)
			{
				$tempstr = ($global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "Name" }).InnerText
			}
			else
			{
				$tempstr = ($global:XMLToParse.Objects.Object.Property | where { $_.Name -eq "TypeName" }).InnerText
			}
						
			$ServiceAppProxyID = $ServiceAppProxyID + "|" + $tempstr
			$tempstr2 = [System.String]$global:XMLToParse.Objects.Object.InnerXml 
			$global:SPServiceAppProxies.Add($ServiceAppProxyID, $tempstr2)
        }	
		return 1
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSPServiceApplicationProxies", $_)
		return 0
    }
}

function o16enumSPServiceApplicationProxyGroups()
{
	try
	{
		$global:SvcAppProxyGroupCount = 0
		$groupstr = Get-SPServiceApplicationProxyGroup | select Id | fl | Out-String
		$delimitLines =  $groupstr.Split("`n")
		
		ForEach($GroupID in $delimitLines)
        {
			$GroupID = $GroupID.Trim()
			if (($GroupID -eq "") -or ($GroupID -eq "Id") -or ($GroupID -eq "--")) { continue }
			$global:SvcAppProxyGroupCount ++
			$GroupID =  ($GroupID.Split(":"))[1]
			$GroupID = $GroupID.Trim()
			$global:XMLToParse = New-Object System.Xml.XmlDocument
			$global:XMLToParse = [xml](Get-SPServiceApplicationProxyGroup | Select-Object * -Exclude Proxies, DefaultProxies | where {$_.Id -eq $GroupID} | Out-String -Width 2000 | ConvertTo-XML )			
			$ProxyGroups = Get-SPServiceApplicationProxyGroup | where {$_.Id -eq $GroupID} | select Proxies
			$ProxiesXML = [xml]($ProxyGroups.Proxies | select DisplayName | ConvertTo-Xml -NoTypeInformation)
			$FriendlyName = Get-SPServiceApplicationProxyGroup | where {$_.Id -eq $GroupID} | select FriendlyName | fl | Out-String
			$FriendlyName =  ($FriendlyName.Split(":"))[1]
			$FriendlyName = $FriendlyName.Trim()
			$ProxiesStr = [System.String]$ProxiesXML.Objects.OuterXML
			$tempstr1 = $GroupID + "|" + $FriendlyName 
			$tempstr2 = [System.String]$global:XMLToParse.Objects.Object.InnerXml 
			$tempstr2 = $tempstr2.Trim() + $ProxiesStr 
			$global:SPServiceAppProxyGroups.Add($tempstr1, $tempstr2)
        }
		return 1
	}	
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumSPServiceApplicationProxyGroups", $_)
		return 0
    }
}

function o16enumWebApps()
{
	try
	{
        $global:totalContentDBCount = 0
        [Microsoft.SharePoint.Administration.SPFarm] $mySPFarm = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.Farm
        [Microsoft.SharePoint.Administration.SPServerCollection] $mySPServerCollection = $mySPFarm.Servers
        [Microsoft.SharePoint.Administration.SPWebApplicationCollection] $mySPWebAppCollection = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.WebApplications
        $global:WebAppnum = $mySPWebAppCollection.Count;
        $global:WebAppDetails = New-Object 'System.String[,]' $global:WebAppnum, 10
        $global:WebAppExtended = New-Object 'System.String[,]' $global:WebAppnum, 6
        $global:WebAppAuthProviders = New-Object 'System.String[,]' $global:WebAppnum, 6
        $global:WebAppIISPath = New-Object 'System.String[,]' $global:WebAppnum, 6
        $global:WebAppAAMs = New-Object 'System.String[]' $global:WebAppnum
        
        $count = 0;
        $global:ContentDBs = New-Object 'System.String[,]' $global:WebAppnum, $global:_maxContentDBs
        $global:WebAppDbID = New-Object 'System.String[,]' $global:WebAppnum, $global:_maxContentDBs
        $global:ContentDBSitesNum = New-Object 'System.String[,]' $global:WebAppnum, $global:_maxContentDBs
		
		if ($mySPWebAppCollection -ne $null)
        {
            foreach ($mySPWebApp in $mySPWebAppCollection)
            {
                $mySiteCollectionNum = 0
                $global:WebAppDetails[$count, 0] = $count.ToString()
				$global:WebAppExtended[$count, 0] = $count.ToString()
                $global:WebAppDetails[$count, 2] = $mySPWebApp.Name
                $global:WebAppDetails[$count, 6] = $mySPWebApp.ApplicationPool.Name
                $global:WebAppDetails[$count, 7] = $mySPWebApp.ApplicationPool.Username
                $global:WebAppDetails[$count, 3] = $mySPWebApp.ContentDatabases.Count.ToString()
                if ($mySPWebApp.ServiceApplicationProxyGroup.Name -eq "")
                    { $global:WebAppDetails[$count, 9] = "[default]" }
                else
                    { $global:WebAppDetails[$count, 9] = $mySPWebApp.ServiceApplicationProxyGroup.Name }
                $global:ContentDBcount = $mySPWebApp.ContentDatabases.Count
                $global:totalContentDBCount = $global:totalContentDBCount + $mySPWebApp.ContentDatabases.Count				
				
				$AllZones = [Microsoft.SharePoint.Administration.SPUrlZone]::Default, 
							[Microsoft.SharePoint.Administration.SPUrlZone]::Intranet, 
							[Microsoft.SharePoint.Administration.SPUrlZone]::Internet, 
							[Microsoft.SharePoint.Administration.SPUrlZone]::Custom, 
							[Microsoft.SharePoint.Administration.SPUrlZone]::Extranet
				#foreach($CurrentZone in $AllZones)
                for($i = 0; $i -le 4; $i++)
				{

                    $global:WebAppAuthProviders[$count, $i] = Get-SPAuthenticationProvider -WebApplication $mySPWebApp -Zone $i -ErrorAction SilentlyContinue | convertto-html -fragment
                    $global:WebAppIISPath[$count, $i] = $mySPWebApp.IisSettings[$i].Path.FullName
				}
				
				#finding out the content dbs for the web app
                [Microsoft.SharePoint.Administration.SPContentDatabaseCollection] $mySPContentDBCollection = $mySPWebApp.ContentDatabases
                foreach ($mySPContentDB in $mySPContentDBCollection)
                {
					$global:ContentDBcount--;
                    $global:ContentDBs[$count, $global:ContentDBcount] = $mySPContentDB.Name
                    $mySiteCollectionNum = $mySiteCollectionNum + $mySPContentDB.Sites.Count
                    $ContentDBSitesNum[$count, $ContentDBcount] = $mySPContentDB.Sites.Count.ToString()					
                }
                $global:WebAppDetails[$count, 4] = $mySiteCollectionNum.ToString()
				
				#enumerating alternateURLs - finished rewrite
                $global:WebAppAAMs[$count] = Get-SPAlternateURL -WebApplication $mySPWebApp | convertto-html -fragment -Property Zone, PublicURL, IncomingURL, Uri 
                					
                #enumerating contentDBs - finished rewrite


				$count++
			}
		}
		return 1
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumWebApps", $_)
		return 0
    }
}

function o16enumContentDBs()
{
	try
	{
	    $DiskSpaceReq = 0.000
	    $global:ContentDBProps = new-object 'System.String[,]' $global:totalContentDBCount, 10
	    $count = 0
	    $queryString = ""
		foreach ($webApplication in [Microsoft.SharePoint.Administration.SPWebService]::ContentService.WebApplications)
	    {
	        $contentDBs = $webApplication.ContentDatabases

	        foreach ($contentDB in $contentDBs)
	        {
	            $global:ContentDBProps[$count, 0] = $contentDB.Name
	            $global:ContentDBProps[$count, 1] = $webApplication.Name
	            $global:ContentDBProps[$count, 2] = $contentDB.Id.ToString()
	            $global:ContentDBProps[$count, 3] = $contentDB.ServiceInstance.DisplayName
	            $global:ContentDBProps[$count, 4] = $contentDB.Sites.Count.ToString()
	            $DiskSpaceReq = [double] $contentDB.DiskSizeRequired / 1048576
	            $global:ContentDBProps[$count, 5] = $DiskSpaceReq.ToString() + " MB"
				$DBConnectionString = $contentDB.DatabaseConnectionString
		    	PSUsing ($sqlConnection = New-Object System.Data.SqlClient.SqlConnection $DBConnectionString)  {   
			        try 
			        {      
						$queryString = "select LockedBy from timerlock with (nolock)"
			            $sqlCommand = $sqlConnection.CreateCommand()      
			            $sqlCommand.CommandText = $queryString       
			            $sqlConnection.Open() | Out-Null      
			            $reader = $sqlcommand.ExecuteReader() 
			            while ($reader.Read())
			            {
			                $global:ContentDBProps[$count, 6] = $reader[0].ToString()
			            }
			            $reader.Close()
			        }    
			        catch [Exception] 
			        {      
			            Write-Host "Exception while running SQL query." -ForegroundColor Cyan      
			            Write-Host "$($_.exception)" -ForegroundColor Black      return -1    
			        }
					
					$global:ContentDBProps[$count, 7] = $contentDB.NeedsUpgrade.ToString()
                    $global:ContentDBProps[$count, 8] = $contentDB.RemoteBlobStorageSettings.Enabled.ToString()
                    if($contentDB.RemoteBlobStorageSettings.ActiveProviderName) { $global:ContentDBProps[$count, 9] = $contentDB.RemoteBlobStorageSettings.ActiveProviderName.ToString() }
    			}            

	            $count = $count + 1
	        }
	    }
	    return 1
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumContentDBs", $_)
		return 0
    }
}

function o16enumCDConfig()
{
	try
	{
		$CDInstance = [Microsoft.SharePoint.Publishing.Administration.ContentDeploymentConfiguration]::GetInstance()
		#Obtaining General Information about the CDInstance
		$global:CDGI = [xml] ($CDInstance | ConvertTo-Xml)
		
		#Obtaining information about deployment paths
		$global:Paths = [Microsoft.SharePoint.Publishing.Administration.ContentDeploymentPath]::GetAllPaths()
		foreach($CDPath in $global:Paths)
		{
			$PathName = $CDPath.Name | Out-String
			$global:XMLToParse = [xml] ($CDPath | ConvertTo-Xml)
			$PathGI = [System.String]$global:XMLToParse.Objects.Object.InnerXml 
			$PathId = $CDPath.Id | fl | Out-String
			$PathId = ($PathId.Split(':'))[1]
			$PathId = $PathId.Trim()
			$tempstr = $PathId + "|" + $PathName
			$global:CDPaths.Add($tempstr, $PathGI)
			
			foreach($Job in $CDPath.Jobs)
			{
				$JobId = $Job.Id | Out-String
				$JobName = $Job.Name | Out-String
				$XMLToParse2 = [xml] ($Job | ConvertTo-Xml)
				$tempstr2 = $PathId + "|" + $JobId + "|" + $JobName
				$tempstr3 = [System.String]$XMLToParse2.Objects.Object.InnerXml 
				$global:CDJobs.Add($tempstr2, $tempstr3)
			}			
		}
		return 1
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		global:HandleException("o15CDConfig", $_)
		return 0
    } 
}

function o16enumHealthReport()
{
	try
	{
		$site = Get-SPSite $global:adminURL
		$web = $site.RootWeb

		$list = $web.Lists["Review problems and solutions"]
		foreach($item in $list.Items)
		{
			$id = $item["ID"]
			$tempstr = $item["Title"] + "||" + $item["Failing Servers"] + "||" + $item["Failing Services"] + "||" + $item["Modified"]
			switch($item["Severity"])
			{
				"0 - Rule Execution Failure"  
				{ $global:HealthReport0.Add($id, $tempstr) }
				"1 - Error" 
				{ $global:HealthReport1.Add($id, $tempstr) }
				"2 - Warning" 
				{ $global:HealthReport2.Add($id, $tempstr) }
				"3 - Information" 
				{ $global:HealthReport3.Add($id, $tempstr) }
				default { }
			}
		}
		return 1
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		global:HandleException("o15enumHealthReport", $_)
		return 0
    } 
}

function o16enumTimerJobs()
{
      try
      {
            $jobs = Get-SPTimerJob | Select Id, Title, Server, WebApplication, Schedule, LastRunTime, IsDisabled, LockType 
            $global:timerJobCount = $jobs.Length
            
            ForEach($Job in $jobs)
            {
                  $timerJobId = ($Job | Select Id | ft -HideTableHeaders | Out-String).Trim()
                  $title = ($Job | Select Title | ft -HideTableHeaders | Out-String).Trim()
                  $server = ($Job | Select Server | ft -HideTableHeaders | Out-String).Trim()
                  $webapplication = ($Job | Select WebApplication | ft -HideTableHeaders | Out-String).Trim()
                  $schedule = ($Job | Select Schedule | ft -HideTableHeaders | Out-String).Trim()
                  $lastruntime = ($Job | Select LastRunTime | ft -HideTableHeaders | Out-String).Trim()
                  $isdisabled = ($Job | Select IsDisabled | ft -HideTableHeaders | Out-String).Trim()
                  $locktype = ($Job | Select LockType | ft -HideTableHeaders | Out-String).Trim()
                  
                  $tempstr2 = $timerJobId + "||" + $title + "||" + $webapplication + "||" + $schedule + "||" + $lastruntime + "||" + $isdisabled + "||" + $locktype
                  $global:timerJobs.Add($timerJobId, $tempstr2)
            }           
            return 1
            }
      catch [System.Exception] 
    {
            Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15enumTimerJobs", $_)
            return 0
    }
}


function o16WriteInitialHTML()
{    
    $global:HTMLHeaders = "<html>
<head>
<title>SPSFarmReport vNext</title>
<!-- Styles -->
<style type=`"text/css`">
                body    { background-color:#FFFFFF; border:1px solid #666666; color:#000000; font-size:68%; font-family:MS Shell Dlg; margin:0,0,10px,0; word-break:normal; word-wrap:break-word; }
                table   { font-size:100%; table-layout:fixed; width:100%; }
                td,th   { overflow:visible; text-align:left; vertical-align:top; white-space:normal; }
                .title  { background:#FFFFFF; border:none; color:#333333; display:block; height:24px; margin:0px,0px,0px,0px; padding-top:0px; position:relative; table-layout:fixed; z-index:5; }
                .he0_expanded    { background-color:#FEF7D6; border:1px solid #BBBBBB; color:#3333CC; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:0px; margin-right:0px; padding-left:8px; padding-right:5em; padding-top:4px; position:relative; }
                .he1_expanded    { background-color:#A0BACB; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:20px; margin-right:0px; padding-left:8px; padding-right:5em; padding-top:4px; position:relative; }
                .he0h_expanded   { background-color: #FEF0D0; border: 1px solid #BBBBBB; color: #000000; cursor: hand; display: block; font-family: MS Shell Dlg; font-size: 100%; font-weight: bold; height: 2.25em; margin-bottom: -1px; margin-left: 5px; margin-right: 0px; padding-left: 8px; padding-right: 5em; padding-top: 4px; position: relative;  }
                .he1h_expanded   { background-color: #7197B3; border: 1px solid #BBBBBB; color: #000000; cursor: hand; display: block; font-family: MS Shell Dlg; font-size: 100%; font-weight: bold; height: 2.25em; margin-bottom: -1px; margin-left: 10px; margin-right: 0px; padding-left: 8px; padding-right: 5em; padding-top: 4px; position: relative; }
                .he1    { background-color:#A0BACB; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:20px; margin-right:0px; padding-left:8px; padding-right:5em; padding-top:4px; position:relative; }
                .he2    { background-color:#C0D2DE; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:30px; margin-right:0px; padding-left:8px; padding-right:5em; padding-top:4px; position:relative; }
                .he3    { background-color:#D9E3EA; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:40px; margin-right:0px; padding-left:11px; padding-right:5em; padding-top:4px; position:relative; }
                .he4    { background-color:#E8E8E8; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:50px; margin-right:0px; padding-left:11px; padding-right:5em; padding-top:4px; position:relative; }
                .he4h   { background-color:#E8E8E8; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:55px; margin-right:0px; padding-left:11px; padding-right:5em; padding-top:4px; position:relative; }
                .he4i   { background-color:#F9F9F9; border:1px solid #BBBBBB; color:#000000; display:block; font-family:MS Shell Dlg; font-size:100%; margin-bottom:-1px; margin-left:55px; margin-right:0px; padding-bottom:5px; padding-left:21px; padding-top:4px; position:relative; }
                .he2i   { background-color:#F9F9F9; border:1px solid #BBBBBB; color:#000000; display:block; font-family:MS Shell Dlg; font-size:100%; margin-bottom:-1px; margin-left:35px; margin-right:0px; padding-bottom:5px; padding-left:21px; padding-top:4px; position:relative;}
                .he5    { background-color:#E8E8E8; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:60px; margin-right:0px; padding-left:11px; padding-right:5em; padding-top:4px; position:relative; }
                .he5h   { background-color:#E8E8E8; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; padding-left:11px; padding-right:5em; padding-top:4px; margin-bottom:-1px; margin-left:65px; margin-right:0px; position:relative; }
                .he5i   { background-color:#F9F9F9; border:1px solid #BBBBBB; color:#000000; display:block; font-family:MS Shell Dlg; font-size:100%; margin-bottom:-1px; margin-left:65px; margin-right:0px; padding-left:21px; padding-bottom:5px; padding-top: 4px; position:relative; }
                div .expando { color:#000000; text-decoration:none; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:normal; position:absolute; right:10px; text-decoration:underline; z-index: 0; }
                .he0 .expando { font-size:100%; }
                .info, .info3, .info4, .disalign  { line-height:1.6em; padding:0px,0px,0px,0px; margin:0px,0px,0px,0px; }
                .disalign TD                      { padding-bottom:5px; padding-right:10px; }
                .info TD                          { padding-right:10px; width:50%; }
                .info3 TD                         { padding-right:10px; width:33%; }
                .info4 TD, .info4 TH              { padding-right:10px; width:25%; }
                .info TH, .info3 TH, .info4 TH, .disalign TH { border-bottom:1px solid #CCCCCC; padding-right:10px; }
                .subtable, .subtable3             { border:1px solid #CCCCCC; margin-left:0px; background:#FFFFFF; margin-bottom:10px; }
                .subtable TD, .subtable3 TD       { padding-left:10px; padding-right:5px; padding-top:3px; padding-bottom:3px; line-height:1.1em; }
                .subtable TH, .subtable3 TH       { border-bottom:1px solid #CCCCCC; font-weight:normal; padding-left:10px; line-height:1.6em;  }
                .subtable .footnote               { border-top:1px solid #CCCCCC; }
                .subtable3 .footnote, .subtable .footnote { border-top:1px solid #CCCCCC; }
                .subtable_frame     { background:#D9E3EA; border:1px solid #CCCCCC; margin-bottom:10px; margin-left:15px; }
                .subtable_frame TD  { line-height:1.1em; padding-bottom:3px; padding-left:10px; padding-right:15px; padding-top:3px; }
                .subtable_frame TH  { border-bottom:1px solid #CCCCCC; font-weight:normal; padding-left:10px; line-height:1.6em; }
                .subtableInnerHead { border-bottom:1px solid #CCCCCC; border-top:1px solid #CCCCCC; }
                .explainlink            { color:#0000FF; text-decoration:none; cursor:hand; }
                .explainlink:hover      { color:#0000FF; text-decoration:underline; }
                .spacer { background:transparent; border:1px solid #BBBBBB; color:#FFFFFF; display:block; font-family:MS Shell Dlg; font-size:100%; height:10px; margin-bottom:-1px; margin-left:43px; margin-right:0px; padding-top: 4px; position:relative; }
                .filler { background:transparent; border:none; color:#FFFFFF; display:block; font:100% MS Shell Dlg; line-height:8px; margin-bottom:-1px; margin-left:53px; margin-right:0px; padding-top:4px; position:relative; }
                .container { display:block; position:relative; }
                .spsfrheader { background-color:#F9F9F9; border-bottom:1px solid black; color:#333333; font-family:MS Shell Dlg; font-size:130%; font-weight:bold; padding-bottom:5px; text-align:center; }
                .rsopname { color:#333333; font-family:MS Shell Dlg; font-size:130%; font-weight:bold; padding-left:11px; }
                .gponame{ color:#333333; font-family:MS Shell Dlg; font-size:130%; font-weight:bold; padding-left:11px; }
                .gpotype{ color:#333333; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; padding-left:11px; }
                #uri    { color:#333333; font-family:MS Shell Dlg; font-size:100%; padding-left:11px; }
                #dtstamp{ color:#333333; font-family:MS Shell Dlg; font-size:100%; padding-left:11px; text-align:left; width:30%; }
                #objshowhide { color:#000000; cursor:hand; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; margin-right:0px; padding-right:10px; text-align:right; text-decoration:underline; z-index:2; word-wrap:normal; }
                #gposummary { display:block; }
                #gpoinformation { display:block; }
                @media print {
                    #objshowhide{ display:none; }
                    body    { color:#000000; border:1px solid #000000; }
                    .title  { color:#000000; border:1px solid #000000; }
                    .he0_expanded    { color:#000000; border:1px solid #000000; }
                    .he1h_expanded   { color:#000000; border:1px solid #000000; }
                    .he1_expanded    { color:#000000; border:1px solid #000000; }
                    .he1    { color:#000000; border:1px solid #000000; }
                    .he2    { color:#000000; background:#EEEEEE; border:1px solid #000000; }
                    .he3    { color:#000000; border:1px solid #000000; }
                    .he4    { color:#000000; border:1px solid #000000; }
                    .he4h   { color:#000000; border:1px solid #000000; }
                    .he4i   { color:#000000; border:1px solid #000000; }
                    .he5    { color:#000000; border:1px solid #000000; }
                    .he5h   { color:#000000; border:1px solid #000000; }
                    .he5i   { color:#000000; border:1px solid #000000; }
                    }
</style>
<!-- Scripts -->
<script type=`"text/javascript`" language=`"javascript`">
/*
String `"strShowHide(0/1)`"
0 = Hide all mode.
1 = Show all mode.
*/

var windowsArray = new Array();
var strShowHide = 0;

//Localized strings

var strShow = `"show`";
var strHide = `"hide`";
var strShowAll = `"show all`";
var strHideAll = `"hide all`";
var strShown = `"shown`";
var strHidden = `"hidden`";
var strExpandoNumPixelsFromEdge = `"10px`";


function IsSectionHeader(obj) {
    return (obj.className === `"he0_expanded`") || (obj.className === `"he0h_expanded`") || (obj.className === `"he1h_expanded`") || (obj.className === `"he1_expanded`") || (obj.className === `"he1`") || (obj.className === `"he2`") || (obj.className === `"he3`") || (obj.className === `"he4`") || (obj.className === `"he4h`") || (obj.className === `"he5`") || (obj.className === `"he5h`");
}

function IsSectionExpandedByDefault(objHeader) {
    if (objHeader === null) {
        return false;
    } else {
        return (objHeader.className.slice(objHeader.className.lastIndexOf(`"_`")) === `"_expanded`");
    }
}

function SetSectionState(objHeader, strState) {
    var i = 0;
    var j;
    var all = objHeader.parentElement.ownerDocument.all;

    if (all === null) {
        return;
    }

    for (j = 0; j < all.length; j++) {
        if (all[j] === objHeader) {
            break;
        }
        i = i + 1;
    }

    for (j = i; j < all.length; j++) {
        if (all[i].className === `"container`") {
            break;
        }
        i = i + 1;
    }

    var objContainer = all[i];

    if (strState === `"toggle`") {
        if (objContainer.style.display === `"none`") {
            SetSectionState(objHeader, `"show`");
        }
        else {
            SetSectionState(objHeader, `"hide`");
        }
    }
    else {
        var objExpando = objHeader.children[1];

        if (strState === `"show`") {
            objContainer.style.display = `"block`";
            objExpando.innerText = strHide;
        }
        else if (strState === `"hide`") {
            objContainer.style.display = `"none`";
            objExpando.innerText = strShow;
        }
    }
}

function ShowSection(objHeader) {
    SetSectionState(objHeader, `"show`");
}

function HideSection(objHeader) {
    SetSectionState(objHeader, `"hide`");
}

function ToggleSection(objHeader) {
    SetSectionState(objHeader, `"toggle`");
}

/*================================================================================
' link at the top of the page to collapse/expand all collapsable elements
'================================================================================
*/
function objshowhide_onClick() {
    var obji;
    var objBody = document.body.getElementsByTagName(`"*`");

    if (objBody === null) {
        return;
    }
    
    switch (strShowHide) {
        case 0:
            strShowHide = 1;
            window.objshowhide.innerText = strShowAll;
            for (obji = 0; obji < objBody.length; obji++) {
                if (objBody[obji].className !== 'undefined' && IsSectionHeader(objBody[obji])) {
                    HideSection(objBody[obji]);
                }
            }
            break;
        case 1:
            strShowHide = 0;
            window.objshowhide.innerText = strHideAll;
            for (obji = 0; obji < objBody.length; obji++) {
                if (objBody[obji].className !== 'undefined' && IsSectionHeader(objBody[obji])) {
                    ShowSection(objBody[obji]);
                }
            }
            break;
    }
}

/*================================================================================
' onload collapse all except the first two levels of headers (he0, he1)
'================================================================================*/
    function window_onload() {
    // Only initialize once.  The UI may reinsert a report into the webbrowser control,
    // firing onLoad multiple times.
        objshowhide_onClick();
}

/*'================================================================================
' When direction (LTR/RTL) changes, change adjust for readability
'================================================================================
*/
function document_onPropertyChange() {
    if (window.event.propertyName === `"dir`") {
        fDetDir(document.dir.toUpperCase());
    }
}

function fDetDir(strDir) {
    var colRules;
    var nug;
    var i;
    var strClass;

    switch (strDir.toUpperCase()) {
        case `"LTR`":
            colRules = document.styleSheets[0].cssRules;
            if (colRules !== null && colRules !== undefined ) {            
                for (i = 0; i < colRules.length - 1; i++) {
                    nug = colRules[i];
                    strClass = nug.selectorText;
                    if (nug.style.textAlign === `"right`") {
                        nug.style.textAlign = `"left`";
                    }
                    switch (strClass) {
                        case `"div .expando`":
                            nug.style.Left = `"`";
                            nug.style.Right = strExpandoNumPixelsFromEdge;
                            break;
                        case `"#objshowhide`":
                            nug.style.textAlign = `"right`";
                            break;
                    }
                }
            }
            break;
        case `"RTL`":
            colRules = document.styleSheets[0].cssRules;
            if (colRules !== null && colRules !== undefined ) {            
                for (i = 0; i < colRules.length - 1; i++) {
                    nug = colRules[i];
                    strClass = nug.selectorText;
                    if (nug.style.textAlign === `"left`") {
                        nug.style.textAlign = `"right`";
                    }
                    switch (strClass) {
                        case `"div .expando`":
                            nug.style.Left = strExpandoNumPixelsFromEdge;
                            nug.style.Right = `"`";
                            break;
                        case `"#objshowhide`":
                            nug.style.textAlign = `"left`";
                            break;
                    }
                }
            }
            break;
    }
}

/*'================================================================================
'When printing reports, if a given section is expanded, let's says `"shown`" (instead of `"hide`" in the UI).
'================================================================================
*/
function window_onbeforeprint() {
    var obji;
    for (obji in document.all) {
        if (document.all.hasOwnProperty(obji)) {
            if (obji.className === `"expando`") {
                if (obji.innerText === strHide) {
                    obji.innerText = strShown;
                }
                if (obji.innerText === strShow) {
                    obji.innerText = strHidden;
                }
            }
        }
    }
}

/*================================================================================
'If a section is collapsed, change to `"hidden`" in the printout (instead of `"show`").
'================================================================================
*/
function window_onafterprint() {
    var obji;
    for (obji in document.all) {
        if (document.all.hasOwnProperty(obji)) {
            if (obji.className === `"expando`") {
                if (obji.innerText === strShown) {
                    obji.innerText = strHide;
                }
                if (obji.innerText === strHidden) {
                    obji.innerText = strShow;
                }
            }
        }
    }
}

/*================================================================================
' Adding keypress support for accessibility
'================================================================================
*/
function document_onkeypress(event) {
    var chCode = ('charCode' in event) ? event.charCode : event.keyCode;

    //space bar (32) or carriage return (13) or line feed (10)
    if (chCode == `"32`" || chCode == `"13`" || chCode == `"10`") {
        if (event.srcElement.className === `"expando`") {
            document_onclick();
            event.returnValue = false;
        }
        if (event.srcElement.className === `"sectionTitle`") {
            document_onclick();
            event.returnValue = false;
        }
        if (event.srcElement.id === `"objshowhide`") {
            objshowhide_onClick();
            event.returnValue = false;
        }
    }
}

/*================================================================================
' When user clicks anywhere in the document body, determine if user is clicking
' on a header element.
'================================================================================
*/
function document_onclick() {
    var strsrc = window.event.srcElement;

    while (strsrc.className === `"sectionTitle`" || strsrc.className === `"expando`") {
        strsrc = strsrc.parentElement;
    }

    // Only handle clicks on headers.
    if (!IsSectionHeader(strsrc)) {
        return;
    }

    ToggleSection(strsrc);

    window.event.returnValue = false;
}

function ToggleState(e) {
    var objParentDisplayItem;
    var objDisplayItem;
    var i;

    if (e.innerText === strShow) {
        e.innerText = strHide;
        objParentDisplayItem = e.parentNode;
        objDisplayItem = objParentDisplayItem.childNodes;
        for (i = 0; i < objDisplayItem.length; i++) {
            if (objDisplayItem[i].id === `"showItem`") {
                objDisplayItem[i].style.display = `"Block`";
            }
        }
    }
    else {
        e.innerText = strShow;
        objParentDisplayItem = e.parentNode;
        objDisplayItem = objParentDisplayItem.childNodes;
        for (i = 0; i < objDisplayItem.length; i++) {
            if (objDisplayItem[i].id === `"showItem`") {
                objDisplayItem[i].style.display = `"None`";
            }
        }
    }
}

function traverseToURL(url) {
    if (url != null) {
        var urlInitialSubstr = url.substring(0, 4).toLowerCase();
        if (urlInitialSubstr === `"http`") {
            window.open(url, `"_blank`");
        }
    }
}

function getExplainWindowTitle() {
    return document.getElementById(`"explainText_windowTitle`").innerHTML;
}

function getExplainWindowStyles() {
    return document.getElementById(`"explainText_windowStyles`").innerHTML;
}

function getExplainWindowSettingPathLabel() {
    return document.getElementById(`"explainText_settingPathLabel`").innerHTML;
}

function getExplainWindowExplainTextLabel() {
    return document.getElementById(`"explainText_explainTextLabel`").innerHTML;
}

function getExplainWindowPrintButton() {
    return document.getElementById(`"explainText_printButton`").innerHTML;
}

function getExplainWindowCloseButton() {
    return document.getElementById(`"explainText_closeButton`").innerHTML;
}

function getNoExplainTextAvailable() {
    return document.getElementById(`"explainText_noExplainTextAvailable`").innerHTML;
}

function getExplainWindowSupportedLabel() {
    return document.getElementById(`"explainText_supportedLabel`").innerHTML;
}

function getNoSupportedTextAvailable() {
    return document.getElementById(`"explainText_noSupportedTextAvailable`").innerHTML;
}

function showExplainText(srcElement)
{
    var strDiagArgs;

    var strSettingName = srcElement.getAttribute(`"gpmc_settingName`");
    var strSettingPath = srcElement.getAttribute(`"gpmc_settingPath`");
    var strSettingDescription = srcElement.getAttribute(`"gpmc_settingDescription`");

    if (strSettingDescription === `"`")
    {
        strSettingDescription = getNoExplainTextAvailable();
    }

    var strSupported = srcElement.getAttribute(`"gpmc_supported`");

    if (strSupported === `"`")
    {
        strSupported = getNoSupportedTextAvailable();
    }

    var strHtml = `"<html dir=`" + document.dir +  `">\n`";
    strHtml += `"<head>\n`";
    strHtml += `"<title>`" + getExplainWindowTitle() + `"</title>\n`";
    strHtml += `"<style type='text/css'>\n`" + getExplainWindowStyles() + `"</style>\n`";
    strHtml += `"</head>\n`";
    strHtml += `"<body>\n`";
    strHtml += `"<div class='head'>`" + strSettingName +`"</div>\n`";
    strHtml += `"<div class='path'><b>`" + getExplainWindowSettingPathLabel() + `"</b><br/>`" + strSettingPath +`"</div>\n`";
    strHtml += `"<div class='path'><b>`" + getExplainWindowSupportedLabel() + `"</b><br/>`" + strSupported +`"</div>\n`";
    strHtml += `"<div class='info'>\n`";
    strHtml += `"<div class='hdr'>`" + getExplainWindowExplainTextLabel() + `"</div>\n`";
    strHtml += `"<div class='bdy'>`" + strSettingDescription + `"</div>\n`";
    strHtml += `"<div class='btn'>`";
    strHtml += getExplainWindowPrintButton();
    strHtml += getExplainWindowCloseButton();
    strHtml += `"</div></body></html>`";

    // IE specific method for showing the popup.
    if(navigator.userAgent.indexOf(`"MSIE`") > 0 && window.location.toString().indexOf(`"file:`") === -1)
    {
        strDiagArgs = `"dialogHeight=360px;dialogWidth=630px;status=no;scroll=yes;resizable=yes;minimize=yes;maximize=yes;`";

        var vModeless = window.showModelessDialog(`"about:blank`", window, strDiagArgs);
        vModeless.document.write(strHtml);
        vModeless.document.close();
        vModeless.location.reload(false);
                        
        window.event.returnValue = false;
    }
    else
    {
        strDiagArgs = `"height=360px, width=630px, status=no, toolbar=no, scrollbars=yes, resizable=yes `";
        
        var expWin = window.open(`"`", `"expWin`", strDiagArgs);
        expWin.document.write(`"`");
        expWin.document.close();
        expWin.document.write(strHtml);
        expWin.document.close();
        expWin.focus();
    }
    
    return false;
}

function showEvents(srcElement,bVerbose,bInformational,bWarning,bError)
{
    var strWindowId = `"EventDetails_`" + srcElement.getAttribute(`"eventLogActivityId`");
    if((windowsArray[strWindowId]) && (windowsArray[strWindowId].closed === false)) {
        windowsArray[strWindowId].focus();
    } else {
        var eventIdLabelNode, eventTimeLabelNode, eventDescriptionLabelNode, eventDetailsLabelNode, eventXmlLabelNode, gpEventsTitleNode;
        var eventIdLabelNodeText, eventTimeLabelNodeText, eventDescriptionLabelNodeText, eventXmlLabelNodeText, gpEventsTitleNodeText, eventDetailsLabelNodeText;
        var singlePassEventsDetailsNode, eventRecordArray;
        var dataNotFoundWarningLabelNode, dataNotFoundWarningLabelNodeText;
        var mainSection;
        var attributeValue;
        var singlePassEventsDetails;
        var singlePassEventsDetailsChildren;
        var node;
        var children;
        var xmlDocumentRoot;
        var xmlDocument;
        var serializer;
        var itemSub;
        var doc;

        if (window.XMLSerializer) 
        {
           serializer = new XMLSerializer();
        }

        if (window.DOMParser) 
        {
           // This browser appears to support DOMParser
           parser = new DOMParser();

           doc = document.getElementById(`"data-island`").textContent;
           xmlDocumentRoot = parser.parseFromString(doc, `"application/xml`");
           xmlDocument = xmlDocumentRoot.documentElement;
           itemSub = 1;
        } 
        else 
        { 
           // Internet Explorer, create a new XML document using ActiveX 
           // and use loadXML as a DOM parser. 
           try 
           {
              doc = document.getElementById(`"data-island`");

              xmlDocumentRoot = new ActiveXObject(`"Msxml2.DOMDocument.6.0`"); 
              xmlDocumentRoot.async = false; 
              xmlDocumentRoot.loadXML(doc.innerHTML);
              xmlDocument = xmlDocumentRoot.documentElement;
              itemSub = 0;
           } 
           catch(e) 
           {
              // Not supported.
           }
        }

        if (xmlDocument != null) {
            mainSection = xmlDocument.getElementsByTagName(`"MainSection`")[0].childNodes;

            if (mainSection != null) {
                for (children = 0; children < mainSection.length; children++) {
                    node = mainSection[children];
                    if (node.nodeType === 1 && node.nodeName === 'Label') {
                        attributeValue = node.getAttribute(`"Name`");
                        if (attributeValue != null) {
                            if (attributeValue === 'ComponentStatus_EventId') {
                                eventIdLabelNode = node.childNodes[1];
                            }
                            if (attributeValue === 'ComponentStatus_EventTime') {
                                eventTimeLabelNode = node.childNodes[1];
                            }
                            if (attributeValue === 'ComponentStatus_EventDescription') {
                                eventDescriptionLabelNode = node.childNodes[1];
                            }
                            if (attributeValue === 'ComponentStatus_EventXml') {
                                eventXmlLabelNode = node.childNodes[1];
                            }
                            if (attributeValue === 'ComponentStatus_EventDetails') {
                                eventDetailsLabelNode = node.childNodes[1];
                            }
                            if (attributeValue === 'ComponentStatus_GPEvents') {
                                gpEventsTitleNode = node.childNodes[1];
                            }
                            if (attributeValue === 'Warning_DataNotFound') {
                                dataNotFoundWarningLabelNode = node.childNodes[1];
                            }
                        }
                    }
                }
            }

            singlePassEventsDetails = xmlDocument.getElementsByTagName(`"SinglePassEventsDetails`");
            if (singlePassEventsDetails != null) {
                for (singlePassEventsDetailsChildren = 0; singlePassEventsDetailsChildren < singlePassEventsDetails.length; singlePassEventsDetailsChildren++) {
                    node = singlePassEventsDetails[singlePassEventsDetailsChildren];
                    attributeValue = node.getAttribute(`"ActivityId`");
                    if (attributeValue === srcElement.getAttribute(`"eventLogActivityId`")) {
                        singlePassEventsDetailsNode = node;
                    }
                }
            }
        }
        
        eventIdLabelNodeText = null;
        if (eventIdLabelNode != null) {
            if (eventIdLabelNode.childNodes.length > 0) {
                eventIdLabelNodeText = eventIdLabelNode.childNodes[0].nodeValue;
            }
        }
        if (eventIdLabelNodeText == null) {
            eventIdLabelNodeText = `"Event ID`";
        }

        eventTimeLabelNodeText = null;
        if (eventTimeLabelNode != null) {
            if (eventTimeLabelNode.firstChild.childNodes.length > 0) {
                eventTimeLabelNodeText = eventTimeLabelNode.childNodes[0].nodeValue;
            }
        }
        if (eventTimeLabelNodeText == null) {
            eventTimeLabelNodeText = `"Event Time`";
        }

        eventDescriptionLabelNodeText = null;
        if (eventDescriptionLabelNode != null) {
            if (eventDescriptionLabelNode.childNodes.length > 0) {
                eventDescriptionLabelNodeText = eventDescriptionLabelNode.childNodes[0].nodeValue;
            }
        }
        if (eventDescriptionLabelNodeText == null) {
            eventDescriptionLabelNodeText = `"Event Description`";
        }

        eventXmlLabelNodeText = null;
        if (eventXmlLabelNode != null) {
            if (eventXmlLabelNode.childNodes.length > 0) {
                eventXmlLabelNodeText = eventXmlLabelNode.childNodes[0].nodeValue;
            }
        }
        if (eventXmlLabelNodeText == null) {
            eventXmlLabelNodeText = `"Event XML`";
        }

        gpEventsTitleNodeText = null;
        if (gpEventsTitleNode != null) {
            if (gpEventsTitleNode.childNodes.length > 0) {
                gpEventsTitleNodeText = gpEventsTitleNode.childNodes[0].nodeValue;
            }
        }
        if (gpEventsTitleNodeText == null) {
            gpEventsTitleNodeText = `"Group Policy Events`";
        }

        eventDetailsLabelNodeText = null;
        if (eventDetailsLabelNode != null) {
            if (eventDetailsLabelNode.childNodes.length > 0) {
                eventDetailsLabelNodeText = eventDetailsLabelNode.childNodes[0].nodeValue;
            }
        }
        if (eventDetailsLabelNodeText == null) {
            eventDetailsLabelNodeText = `"Event Details`";
        }

        dataNotFoundWarningLabelNodeText = null;
        if (dataNotFoundWarningLabelNode != null) {
            if (dataNotFoundWarningLabelNode.childNodes.length > 0) {
                dataNotFoundWarningLabelNodeText = dataNotFoundWarningLabelNode.childNodes[0].nodeValue;
            }
        }
        if (dataNotFoundWarningLabelNodeText == null) {
            dataNotFoundWarningLabelNodeText = `"Data Not Found`";
        }
                
        if(singlePassEventsDetailsNode != null)
        {
            eventRecordArray = singlePassEventsDetailsNode.getElementsByTagName(`"EventRecord`");
        }
        
        var htmlText = `"<html dir=`" + document.dir +  `">`";
        htmlText = htmlText + `"<head>`";
        htmlText = htmlText + `"<meta http-equiv=\`"X-UA-Compatible\`" content=\`"IE=edge\`" />`";
        htmlText = htmlText + `"<meta http-equiv=\`"Content-Type\`" content=\`"text/html; charset=UTF-16\`" />`";
        htmlText = htmlText + `"<title>`" + gpEventsTitleNodeText + `"</title>`";
        htmlText = htmlText + `"</head><style type=\`"text/css\`">`";
        htmlText = htmlText + `"body    { background-color:#FFFFFF; color:#000000; font-size:68%; font-family:MS Shell Dlg; margin:0,0,10px,0; word-break:normal; word-wrap:break-word; }`";
        htmlText = htmlText + `"table   { font-size:100%; table-layout:fixed; width:100%; }`";
        htmlText = htmlText + `"td,th   { overflow:visible; text-align:left; vertical-align:top; white-space:normal; }`";
        htmlText = htmlText + `".he1    { text-align: center; vertical-align: middle; background-color:#C0D2DE; border:1px solid #BBBBBB; color:#000000; cursor:hand; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:4em; position:relative; }`";
        htmlText = htmlText + `".centerTxt { text-align: center; }`";
        htmlText = htmlText + `".txtFormat1 { text-align: left; vertical-align:top; white-space:pre-line; }`";
        htmlText = htmlText + `"</style>`";
      
        htmlText = htmlText + `"<script> function toggle(e) {`";
        htmlText = htmlText + `"if (e.style.display === \`"none\`"){ e.style.display = \`"\`"; }`";
        htmlText = htmlText + `"else { e.style.display = \`"none\`"; }`";
        htmlText = htmlText + `"}</`";
        htmlText = htmlText + `"script`";
        htmlText = htmlText + `">`";
      
        htmlText = htmlText + `"<body><table border=1><tr>`";
        htmlText = htmlText + `"<th class=\`"he1\`"><strong>`" + eventIdLabelNodeText + `"</strong></th>`";
        htmlText = htmlText + `"<th class=\`"he1\`"><strong>`" + eventTimeLabelNodeText + `"</strong></th>`";
        htmlText = htmlText + `"<th class=\`"he1\`"><strong>`" + eventDescriptionLabelNodeText + `"</strong></th>`";
        htmlText = htmlText + `"<th class=\`"he1\`"><strong>`" + eventDetailsLabelNodeText + `"</strong></th>`";

        htmlText = htmlText + `"</tr>`";
        var i;
        var eventId;
        var eventTime;
        var eventDescription;
        var eventXml;
        var eventType;
        var displayEvent;
        var eventXmlId;
        var displayBgColor;

        if(eventRecordArray != null && eventRecordArray.length > 0)
        {
            for (i=0; i < eventRecordArray.length; i++)
            {
                displayEvent = false;
                var eventIdElements = eventRecordArray[i].getElementsByTagName(`"EventId`");        
                if((eventIdElements != null) && (eventIdElements.length > 0) && (eventIdElements[0].firstChild != null))
                {
                    eventId =  eventIdElements[0].firstChild.nodeValue;
                }
                else
                {
                    eventId =  dataNotFoundWarningLabelNodeText;
                }
                var eventTimeElements = eventRecordArray[i].getElementsByTagName(`"EventTime`");
                if((eventTimeElements != null) && (eventTimeElements.length > 0) && (eventTimeElements[0].firstChild != null))
                {
                    eventTime = eventTimeElements[0].firstChild.nodeValue;
                }
                else
                {
                    eventTime = dataNotFoundWarningLabelNodeText;
                }
                var eventDescriptionElements = eventRecordArray[i].getElementsByTagName(`"EventDescription`");
                if((eventDescriptionElements != null) && (eventDescriptionElements.length > 0) && (eventDescriptionElements[0].firstChild != null))
                {
                        eventDescription = eventDescriptionElements[0].firstChild.nodeValue;
                }
                else
                {
                    eventDescription = dataNotFoundWarningLabelNodeText;
                }
                var eventXmlElements = eventRecordArray[i].getElementsByTagName(`"EventXml`");
                if((eventXmlElements != null) && (eventXmlElements.length > 0) && (eventXmlElements[0].firstChild != null))
                {
                    if (window.XMLSerializer) 
                    {
                       var xml = serializer.serializeToString(eventXmlElements[0].firstChild);
                       eventXml = xml;
                    } 
                    else 
                    {
                       if (typeof eventXmlElements[0].firstChild.xml != `"undefined`") 
                       {
                          eventXml = eventXmlElements[0].firstChild.xml;
                       }
                    }
                }
                else
                {
                    eventXml = dataNotFoundWarningLabelNodeText;
                }
                var eventLevelElements = eventRecordArray[i].getElementsByTagName(`"EventLevel`");
                if((eventLevelElements != null) && (eventLevelElements.length > 0) && (eventLevelElements[0].firstChild != null))
                {
                    eventType = eventLevelElements[0].firstChild.nodeValue;
                }
                else
                {
                    eventType = 5;
                }
                
                if((bVerbose === true)&&(eventType == 5))
                {
                    displayEvent = true;
                }
                else if((bInformational === true)&&(eventType == 4))
                {
                    displayEvent = true;
                }
                else if((bWarning === true)&&(eventType == 3))
                {
                    displayEvent = true;
                }
                else if((bError === true)&&((eventType == 1)||(eventType == 2)))
                {
                    displayEvent = true;
                }
                
                if (displayEvent === true)
                {
                    eventXmlId = `"EventXml`" + (i+`"`");
                    htmlText = htmlText + `"<tr>`";
                    htmlText = htmlText + `"<td class=\`"centerTxt\`" style=\`"background:`" + displayBgColor +`"\`">`" + eventId + `"</td>`";
                    htmlText = htmlText + `"<td class=\`"centerTxt\`" style=\`"background:`" + displayBgColor +`"\`">`" + eventTime + `"</td>`";
                    htmlText = htmlText + `"<td class=\`"txtFormat1\`" style=\`"background:`" + displayBgColor +`"\`">`" + eventDescription + `"</td>`";
                    htmlText = htmlText + `"<td style=\`"background:`" + displayBgColor +`"\`"><span style=\`"color:blue; cursor:hand\`" onclick=\`"toggle(`" + eventXmlId +`");\`" onKeyPress=\`"toggle(`" + eventXmlId + `");\`" tabIndex=1 >`";
                    htmlText = htmlText + eventXmlLabelNodeText + `"</span><br/>`";
                    htmlText = htmlText + `"<span style=\`"display:none\`" id=`" + eventXmlId +`">`";
                    htmlText = htmlText + eventXml + `"</span>`";
                    htmlText = htmlText + `"</td>`";
                    htmlText = htmlText + `"</tr>`";
                }
            }
        }
        htmlText = htmlText + `"</table></body></html>`";

        if(windowsArray[strWindowId])
        {
            delete windowsArray[strWindowId];
        }
        
        // IE specific method for showing the popup.
        if(navigator.userAgent.indexOf(`"MSIE`") > 0 && window.location.toString().indexOf(`"file:`") === -1)
        {
            var strDiagArgs = `"dialogHeight=360px;dialogWidth=630px;status=no;scroll=yes;resizable=yes;minimize=yes;maximize=yes;`";

            var vModeless = window.showModelessDialog(`"about:blank`", window, strDiagArgs);
            vModeless.document.write(htmlText);
            vModeless.document.close();
            vModeless.location.reload(false);
            windowsArray[strWindowId] = vModeless;            
        }
        else
        {
            var strDiagArgs = `"height=360px, width=630px, status=no, toolbar=no, scrollbars=yes, resizable=yes`";
        
            windowsArray[strWindowId] = window.open(`"`", `"`", strDiagArgs);
            windowsArray[strWindowId].document.write(htmlText);
            windowsArray[strWindowId].focus();
        } 
    }

    xmlDocumentRoot = null;
}

function cleanUp() {
    var windowsArray = this.windowsArray;
    for (var currentWindow in windowsArray) {
        if (windowsArray.hasOwnProperty(currentWindow)) {
            windowsArray[currentWindow].close();
        }
    }
}

function getMessageText(messageNode) {
    if (messageNode != null) {
        if (messageNode.firstChild != null) {
            if (messageNode.firstChild.nodeType === 3) {
                for (var i = 0; i < messageNode.childNodes.length; i++) 
                {
                    var curNode = messageNode.childNodes[i];
                    if(curNode.nodeType === 1){
                        return curNode.childNodes[0].nodeValue;
                    }
                }
            } else {
                return messageNode.firstChild.childNodes[0].nodeValue;
            }
        }
    }
    return null;
}

function showComponentProcessingDetails(srcElement) {
    var strWindowId = `"ProcessingDetails_`" + srcElement.getAttribute(`"eventLogActivityId`");
    if ((windowsArray[strWindowId]) && (windowsArray[strWindowId].closed === false)) {
        windowsArray[strWindowId].focus();
    } else {
        var doc;
        var parser;
        var xmlDocumentRoot;
        var xmlDocument;

        var extensionsProcessedLabelNode, slowLinkThresholdLabelNode, linkSpeedLabelNode, extensionsProcessedTimeTakenNode;
        var domainControllerIpLabelNode, domainControllerNameLabelNode, processingTypeLabelNode, loopbackModeLabelNode;
        var processingTriggerLabelNode, extensionNameLabelNode, timeTakenLabelNode;
        var dataNotFoundWarningLabelNode;
        var singlePassEventsDetailsNode, totalProcessingTimeLabelNode, refreshMessageLabelNode;
        var processingDetailsUserTitleNode, processingDetailsComputerTitleNode;
        var policySectionNode;
        var policyEventsDetailsNode, detailsLabelNode;

        var extensionsProcessedLabelNodeText, slowLinkThresholdLabelNodeText, linkSpeedLabelNodeText, extensionsProcessedTimeTakenNodeText;
        var domainControllerIpLabelNodeText, domainControllerNameLabelNodeText, processingTypeLabelNodeText, loopbackModeLabelNodeText;
        var processingTriggerLabelNodeText, extensionNameLabelNodeText, timeTakenLabelNodeText;
        var dataNotFoundWarningLabelNodeText, totalProcessingTimeLabelNodeText, refreshMessageLabelNodeText;
        var processingDetailsUserTitleNodeText, processingDetailsComputerTitleNodeText;
        var detailsLabelNodeText;

        var slowLinkThresholdValue, linkSpeedValue, domainControllerIpValue, domainControllerNameValue;
        var processingTypeValue, loopbackModeValue, processingTriggerValue, totalPolicyProcessingTime, extensionProcessingTimeArray;
        var cseNameArray = new Array();
        var cseElapsedTimeArray = new Array();
        var policyApplicationFinishedTime;

        var isComputerProcessing;
        var strDiagArgs;
        var mainSection;
        var attributeValue;
        var singlePassEventsDetails;
        var singlePassEventsDetailsChildren;
        var node;
        var children;
        var itemSub;

        if (window.DOMParser) 
        {
           // This browser appears to support DOMParser
           parser = new DOMParser();
           doc = document.getElementById(`"data-island`").textContent;

           xmlDocumentRoot = parser.parseFromString(doc, `"application/xml`");

           xmlDocument = xmlDocumentRoot.documentElement;

           itemSub = 1;
        } 
        else 
        { 
           // Internet Explorer, create a new XML document using ActiveX 
           // and use loadXML as a DOM parser. 
           try 
           {
              doc = document.getElementById(`"data-island`");

              xmlDocumentRoot = new ActiveXObject(`"Msxml2.DOMDocument.6.0`"); 
              xmlDocumentRoot.async = false; 
              xmlDocumentRoot.loadXML(doc.innerHTML);
              xmlDocument = xmlDocumentRoot.documentElement;
              itemSub = 0;
           } 
           catch(e) 
           {
              // Not supported.
           }
        }

        if (xmlDocument != null) {
            mainSection = xmlDocument.getElementsByTagName(`"MainSection`")[0].childNodes;

            if (mainSection != null) {
                for (children = 0; children < mainSection.length; children++) {
                    node = mainSection[children];
                    if (node.nodeType === 1 && node.nodeName === 'Label') {
                        attributeValue = node.getAttribute(`"Name`")
                        if (attributeValue != null) {
                            if (attributeValue === 'ComponentStatus_ExtensionsProcessed') {
                                extensionsProcessedLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_SlowLinkThreshold') {
                                slowLinkThresholdLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_LinkSpeed') {
                                linkSpeedLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_TimeTaken') {
                                extensionsProcessedTimeTakenNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_DomainControllerIP') {
                                domainControllerIpLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_DomainControllerName') {
                                domainControllerNameLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_ProcessingTrigger') {
                                processingTriggerLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_ExtensionName') {
                                extensionNameLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_TimeTaken') {
                                timeTakenLabelNode = node;
                            }
                            if (attributeValue === 'Warning_DataNotFound') {
                                dataNotFoundWarningLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_TotalProcessingTime') {
                                totalProcessingTimeLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_RefreshMessage') {
                                refreshMessageLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_UserProcessingDetails') {
                                processingDetailsUserTitleNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_ComputerProcessingDetails') {
                                detailsLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_ProcessingType') {
                                processingTypeLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_LoopbackMode') {
                                loopbackModeLabelNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_UserProcessingDetails') {
                                processingDetailsUserTitleNode = node;
                            }
                            if (attributeValue === 'ComponentStatus_ComputerProcessingDetails') {
                                processingDetailsComputerTitleNode = node;
                            }
                        }
                    }
                }
            }

            singlePassEventsDetails = xmlDocument.getElementsByTagName(`"SinglePassEventsDetails`");
            if (singlePassEventsDetails != null) {
                for (singlePassEventsDetailsChildren = 0; singlePassEventsDetailsChildren < singlePassEventsDetails.length; singlePassEventsDetailsChildren++) {
                    node = singlePassEventsDetails[singlePassEventsDetailsChildren];
                    if (node.getAttribute(`"ActivityId`") === srcElement.getAttribute(`"eventLogActivityId`")) {
                        singlePassEventsDetailsNode = node;
                    }
                }
            }

            if (singlePassEventsDetailsNode) {
                policyEventsDetailsNode = singlePassEventsDetailsNode.parentNode;
                if (policyEventsDetailsNode) {
                    policySectionNode = policyEventsDetailsNode.parentNode;
                    if (policySectionNode) {
                        if (policySectionNode.nodeName === 'UserPolicySection') {
                            isComputerProcessing = false;
                        }
                        if (policySectionNode.nodeName === 'ComputerPolicySection') {
                            isComputerProcessing = true;
                        }
                    }
                }
            }
        }

        
        extensionsProcessedLabelNodeText = getMessageText(extensionsProcessedLabelNode);
        slowLinkThresholdLabelNodeText = getMessageText(slowLinkThresholdLabelNode);
        linkSpeedLabelNodeText = getMessageText(linkSpeedLabelNode);
        domainControllerIpLabelNodeText = getMessageText(domainControllerIpLabelNode);
        domainControllerNameLabelNodeText = getMessageText(domainControllerNameLabelNode);
        processingTypeLabelNodeText = getMessageText(processingTypeLabelNode);
        loopbackModeLabelNodeText = getMessageText(loopbackModeLabelNode);
        processingTriggerLabelNodeText = getMessageText(processingTriggerLabelNode);
        extensionNameLabelNodeText = getMessageText(extensionNameLabelNode);
        timeTakenLabelNodeText = getMessageText(timeTakenLabelNode);
        processingDetailsUserTitleNodeText = getMessageText(processingDetailsUserTitleNode);
        processingDetailsComputerTitleNodeText = getMessageText(processingDetailsComputerTitleNode);
        dataNotFoundWarningLabelNodeText = getMessageText(dataNotFoundWarningLabelNode);
        totalProcessingTimeLabelNodeText = getMessageText(totalProcessingTimeLabelNode);
        refreshMessageLabelNodeText = getMessageText(refreshMessageLabelNode);
        detailsLabelNodeText = getMessageText(detailsLabelNode);
     

        slowLinkThresholdValue = null;
        linkSpeedValue = null;
        domainControllerIpValue = null;
        domainControllerNameValue = null;
        processingTypeValue = null;
        loopbackModeValue = null;
        processingTriggerValue = null;

        if (singlePassEventsDetailsNode != null) {
            slowLinkThresholdValue = singlePassEventsDetailsNode.getAttribute(`"SlowLinkThresholdInKbps`");
            linkSpeedValue = singlePassEventsDetailsNode.getAttribute(`"LinkSpeedInKbps`");
            domainControllerIpValue = singlePassEventsDetailsNode.getAttribute(`"DomainControllerIPAddress`");
            domainControllerNameValue = singlePassEventsDetailsNode.getAttribute(`"DomainControllerName`");
            processingTypeValue = singlePassEventsDetailsNode.getAttribute(`"ProcessingAppMode`");
            loopbackModeValue = singlePassEventsDetailsNode.getAttribute(`"PolicyProcessingMode`");
            processingTriggerValue = singlePassEventsDetailsNode.getAttribute(`"ProcessingTrigger`");
            totalPolicyProcessingTime = singlePassEventsDetailsNode.getAttribute(`"PolicyElapsedTime`");
            extensionProcessingTimeArray = singlePassEventsDetailsNode.getElementsByTagName(`"ExtensionProcessingTime`");
        }
        if (slowLinkThresholdValue == null) {
            slowLinkThresholdValue = dataNotFoundWarningLabelNodeText;
        }
        if (linkSpeedValue == null) {
            linkSpeedValue = dataNotFoundWarningLabelNodeText;
        }
        if (domainControllerIpValue == null) {
            domainControllerIpValue = dataNotFoundWarningLabelNodeText;
        }
        else {
            domainControllerIpValue = domainControllerIpValue.replace(/^\\\\/, `"`");
        }
        if (domainControllerNameValue == null) {
            domainControllerNameValue = dataNotFoundWarningLabelNodeText;
        }
        else {
            domainControllerNameValue = domainControllerNameValue.replace(/^\\\\/, `"`");
        }
        if (processingTypeValue == null) {
            processingTypeValue = dataNotFoundWarningLabelNodeText;
        }
        if (loopbackModeValue == null) {
            loopbackModeValue = dataNotFoundWarningLabelNodeText;
        }
        if (processingTriggerValue == null) {
            processingTriggerValue = dataNotFoundWarningLabelNodeText;
        }

        if (extensionProcessingTimeArray != null && extensionProcessingTimeArray.length > 0) {
            var cseName;
            var cseElapsedTime;
            var cseProcessedTime;
            var cseId;
            var i;
            var index = 0;
            for (i = 0; i < extensionProcessingTimeArray.length; i++) {
                var cseNameElements = extensionProcessingTimeArray[i].getElementsByTagName(`"ExtensionName`");
                var cseElapsedTimeElements = extensionProcessingTimeArray[i].getElementsByTagName(`"ElapsedTime`");
                var cseProcessedTimeElements = extensionProcessingTimeArray[i].getElementsByTagName(`"ProcessedTime`");
                var cseIdElements = extensionProcessingTimeArray[i].getElementsByTagName(`"ExtensionGuid`");
                if ((cseNameElements.length > 0) && (cseElapsedTimeElements.length > 0) && (cseProcessedTimeElements.length > 0) && (cseIdElements.length > 0)) {
                    if ((cseNameElements[0].firstChild != null) && (cseElapsedTimeElements[0].firstChild != null) && (cseProcessedTimeElements[0].firstChild != null) && (cseIdElements[0].firstChild != null)) {
                        cseName = cseNameElements[0].firstChild.nodeValue;
                        cseElapsedTime = cseElapsedTimeElements[0].firstChild.nodeValue;
                        cseProcessedTime = cseProcessedTimeElements[0].firstChild.nodeValue;
                        cseId = cseIdElements[0].firstChild.nodeValue;
                        if ((cseName != null) && (cseElapsedTime != null) && (cseProcessedTime != null) && (cseId != null)) {
                            cseNameArray[index] = cseName;
                            cseElapsedTimeArray[index] = cseElapsedTime;
                            index = index + 1;
                            if (cseId === '{00000000-0000-0000-0000-000000000000}') {
                                policyApplicationFinishedTime = cseProcessedTime;
                            }
                        }
                    }
                }
            }
        }
          
        var htmlText = `"<html dir=`" + document.dir +  `">`";
        htmlText = htmlText + `"<head>`";
        htmlText = htmlText + `"<meta http-equiv=\`"X-UA-Compatible\`" content=\`"IE=edge\`" />`";
        htmlText = htmlText + `"<meta http-equiv=\`"Content-Type\`" content=\`"text/html; charset=UTF-16\`" />`";
        if(isComputerProcessing != null)
        {
            if(isComputerProcessing === true)
            {
                htmlText = htmlText + `"<title>`" + processingDetailsComputerTitleNodeText + `"</title>`";
            }
            else
            {
                htmlText = htmlText + `"<title>`" + processingDetailsUserTitleNodeText + `"</title>`";
            }
        }
        

        htmlText = htmlText + `"</head><style type=\`"text/css\`">`";
        htmlText = htmlText + `"body    { background-color:#FFFFFF; color:#000000; font-size:68%; font-family:MS Shell Dlg; margin:0,0,10px,0; word-break:normal; word-wrap:break-word; }`";
        htmlText = htmlText + `"table   { font-size:100%; table-layout:fixed; width:100%; }`";
        htmlText = htmlText + `"td,th   { overflow:visible; text-align:left; vertical-align:top; white-space:normal; }`";
        htmlText = htmlText + `".he0    { background-color:#FEF7D6; border:1px solid #BBBBBB; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2.25em; margin-bottom:-1px; margin-left:0px; margin-right:0px; padding-left:8px; padding-right:5em; padding-top:4px; position:relative; width:100%; }`";
        htmlText = htmlText + `".he1    { color:#000000; display:block; font-family:MS Shell Dlg; font-size:100%; font-weight:bold; height:2em;margin-left: 5 px; margin-top: 5 px; position:relative; width:100%; }`";
        htmlText = htmlText + `".tblspecialfmt { border:1px solid black;border-collapse:collapse; }`";
        htmlText = htmlText + `".tblfirstcolfmt { border-left-width: 1px;border-top-width: 1px;border-bottom-width: 1px;border-right-width: 0px;border-style: solid; border-color: black; }`";
        htmlText = htmlText + `".tblsecondcolfmt { border-left-width: 0px;border-top-width: 1px;border-bottom-width: 1px;border-right-width: 1px;border-style: solid; border-color: black; }`";
        htmlText = htmlText + `"</style>`";
        htmlText = htmlText + `"<body>`";
        htmlText = htmlText + `"<span class=\`"he1\`">`" + refreshMessageLabelNodeText + `" `" + policyApplicationFinishedTime + `"</span>`" ;
        htmlText = htmlText + `"<div class=\`"he0\`">`" + detailsLabelNodeText + `"</div>`"
        htmltext = htmlText + `"<table><tr>`";

        htmlText = htmlText + `"<td>`";
        htmlText = htmlText + `"<table>`";
        htmlText = htmlText + `"<tr><td colspan=\`"2\`">&nbsp;</td></tr>`";


        htmlText = htmlText + `"<tr><td style=\`"width: 50%\`"><strong>`" + processingTypeLabelNodeText + `"</strong></td>`";
        htmlText = htmlText + `"<td>`" + processingTypeValue + `"</td></tr>`";

        htmlText = htmlText + `"<tr><td style=\`"width: 50%\`"><strong>`" + loopbackModeLabelNodeText + `"</strong></td>`";
        htmlText = htmlText + `"<td>`" + loopbackModeValue + `"</td></tr>`";

        htmlText = htmlText + `"<tr><td style=\`"width: 50%\`"><strong>`" + linkSpeedLabelNodeText + `"</strong></td>`";
        htmlText = htmlText + `"<td>`" + linkSpeedValue + `"</td></tr>`";

        htmlText = htmlText + `"<tr><td style=\`"width: 50%\`"><strong>`" + slowLinkThresholdLabelNodeText + `"</strong></td>`";
        htmlText = htmlText + `"<td>`" + slowLinkThresholdValue + `"</td></tr>`";

        htmlText = htmlText + `"<tr><td style=\`"width: 50%\`"><strong>`" + domainControllerNameLabelNodeText + `"</strong></td>`";
        htmlText = htmlText + `"<td>`" + domainControllerNameValue +`"</td></tr>`";

        htmlText = htmlText + `"<tr><td style=\`"width: 50%\`"><strong>`" + domainControllerIpLabelNodeText + `"</strong></td>`";
        htmlText = htmlText + `"<td>`" + domainControllerIpValue +`"</td></tr>`";

        htmlText = htmlText + `"<tr><td style=\`"width: 50%\`"><strong>`" + processingTriggerLabelNodeText + `"</strong></td>`";
        htmlText = htmlText + `"<td>`" + processingTriggerValue + `"</td></tr>`";

        htmlText = htmlText + `"</table></td></tr>`";
        htmlText = htmlText + `"<tr ><td ><table>`";


        htmlText = htmlText + `"<tr><td><span class=\`"he1\`" >`" + extensionsProcessedLabelNodeText +`"</span></td></tr>`";
        htmlText = htmlText + `"<tr><td><table class=\`"tblspecialfmt\`" >`";
        htmlText = htmlText + `"<tr><td class=\`"tblfirstcolfmt\`" style=\`"width: 50%;background-color:#FEF7D6;\`"><strong>`" + extensionNameLabelNodeText + `"</strong></td>`";
        htmlText = htmlText + `"<td class=\`"tblsecondcolfmt\`" style=\`"background-color:#FEF7D6;\`" ><strong>`" + timeTakenLabelNodeText + `"</strong></td></tr>`";

        for (var idx in cseNameArray)
        {
            htmlText = htmlText + `"<tr><td style=\`"width: 50%\`">`" + cseNameArray[idx] + `"</td>`";                   
            htmlText = htmlText + `"<td>`" + cseElapsedTimeArray[idx] + `"</td></tr>`";
        }

        if (totalPolicyProcessingTime != null)
        {
            htmlText = htmlText + `"<tr><td class=\`"tblfirstcolfmt\`" style=\`"width: 50%\`" >`" + totalProcessingTimeLabelNodeText +`":</td>`";
            htmlText = htmlText + `"<td class=\`"tblsecondcolfmt\`">`" + totalPolicyProcessingTime + `"</td></tr>`";
        }
        htmlText = htmlText + `"</table></td></tr></table></td></tr></table></body></html>`";

        if(windowsArray[strWindowId])
        {
            delete windowsArray[strWindowId];
        }
         
        // IE specific method for showing the popup.
        if(navigator.userAgent.indexOf(`"MSIE`") > 0 && window.location.toString().indexOf(`"file:`") === -1)
        {
            strDiagArgs = `"dialogHeight=360px;dialogWidth=630px;status=no;scroll=yes;resizable=yes;minimize=yes;maximize=yes;`";

            var vModeless = window.showModelessDialog(`"about:blank`", window, strDiagArgs);
            vModeless.document.write(htmlText);
            vModeless.document.close();
            vModeless.location.reload(false);
            windowsArray[strWindowId] = vModeless;                      
        }
        else
        {
            strDiagArgs = `"height=360px, width=630px, status=no, toolbar=no, scrollbars=yes, resizable=yes`";
        
            windowsArray[strWindowId] = window.open(`"`", `"`" , strDiagArgs);
            windowsArray[strWindowId].document.write(htmlText);
            windowsArray[strWindowId].focus();
        }
    }

    xmlDocumentRoot = null;
}
</script>
</head> <body onload=`"window_onload();`" onclick=`"document_onclick();`" onkeypress=`"document_onkeypress(event);`" onunload=`"cleanUp();`"> "

    $global:HTMLTrail = "</body></html>"
    $global:htmlContents  = $global:HTMLHeaders

}

function o16WriteFarmGenSettings()
{
	try
	{
        # Farm General Settings - StartHTML
        $global:htmlContents += "<table class=`"title`" > <tr><td colspan=`"2`" class=`"spsfrheader`">SPSFarmReport</td></tr> <tr><td><div id=`"objshowhide`" tabindex=`"0`" onclick=`"objshowhide_onClick();return false;`"></div></td></tr> </table>"
        $global:htmlContents +=  "<div class=`"spsfrsettings`"> <div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Topology Details</span><a class=`"expando`" href=`"#`"></a></div> <div class=`"container`"><div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">General</span><a class=`"expando`" href=`"#`"></a> </div><div class=`"container`"><div class=`"he2i`">"
        $global:htmlContents += "<table class=`"info`" > <tr><td><strong>Central Admin URL</strong></td><td>" + $global:adminURL + "</td></tr><tr><td><strong>Farm Build Version</strong></td><td>" + $global:BuildVersion + "</td></tr><tr><td><strong>System Account</strong></td><td>" + $global:systemAccount 
        $global:htmlContents += "</td></tr><tr><td><strong>Configuration Database Name</strong></td><td>" + $global:configDbName + "</td></tr><tr><td><strong>Configuration Database Server Name</strong></td><td>" + $global:configDbServerName        
        $global:htmlContents +=  "</td></tr> <tr><td><strong>Admin Content Database Name</strong></td><td>" + $global:adminDbName + "</td></tr> </table>"
        $global:htmlContents += "</div></div></div>"
        $global:htmlContents += "<div class=`"filler`"></div>"
        # Farm General Settings - EndHTML
		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15WriteFarmGenSettings", $_)
    }
}


function o16writeServers()
{
	try
	{
    # Write Server Info - StartHTML
    $global:htmlContents += "<div class=`"spsfrsettings`"> <div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`"> Servers in Farm</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
    for($i = $global:Servernum; $i -gt 0; $i--)
    {
        $global:ServerRoles.GetEnumerator() | ForEach-Object {		
                if($_.key -eq $Servers[$i - 1]) {                
		                $Role = ($_.value.Split(','))[0]
		                $Compliance = ($_.value.Split(','))[1]	    }
                        }
        $global:htmlContents +=  "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ $Servers[$i - 1] + " Role:" + $Role + " Compliance:" + $Compliance+ "</span><a class=`"expando`" href=`"#`"></a></div>  <div class=`"container`"> "
        $global:htmlContents += "<div class=`"he2i`">"
        $global:htmlContents += "<table class=`"info`" >"
        for($j = [System.Int16]::Parse(($ServicesOnServers[($i - 1), ($global:_maxServicesOnServers - 1)])); $j -ge 0 ; $j--)
	    	{
               
	           if ($ServicesOnServers[($i - 1), $j] -ne $null)
               {
                if ($j -eq 0) { $global:htmlContents += "<tr><td><strong>"+ $ServicesOnServers[($i - 1), $j] +"</strong></td><tr>" }
                else { $global:htmlContents += "<tr><td><strong>"+ $ServicesOnServers[($i - 1), $j] +"</strong></td><tr>" }
                }
                
			}
        $global:htmlContents += "</table>"
        $global:htmlContents += "</div></div>"
        $global:htmlContents += "<div class=`"filler`"></div>"        
    }
    $global:htmlContents += "</div></div>"
    # Write Server Info - EndHTML

	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeServers", $_)
    }
}

function o16writeProdVersions2()
{ 
	$thCount = $global:_maxItemsonServer - 1
	$writtenItem, $itemVal2Found, $allProductsConsistent = [Boolean] "false", [Boolean] "false", [Boolean] "false"
	$totalProducts = 0
	try
	{
        for ($count = ($global:Servernum - 1); $count -ge 0; $count--)
        {
            if ($global:serverProducts[$count, 0, 0] -eq $null)
                { continue }
				
				
            if ( [System.Convert]::ToInt32(($global:serverProducts[$count, ($global:_maxProductsonServer - 1), ($global:_maxItemsonServer - 1)])) -gt $totalProducts)
                { $totalProducts = [System.Convert]::ToInt32(($global:serverProducts[$count, ($global:_maxProductsonServer - 1), ($global:_maxItemsonServer - 1) ])) }
		}
		
		# get names of the installed products 
        $productsInstalled = New-Object System.Collections.ArrayList
        $itemsInstalled = New-Object System.Collections.ArrayList
        $itemsWriter = New-Object System.Collections.ArrayList
		
        for ($count = ($global:Servernum - 1); $count -ge 0; $count--)
        {
            for ($count2 = ($global:_maxProductsonServer - 1); $count2 -ge 1; $count2--)
            {
                if (!$productsInstalled.Contains(($global:serverProducts[$count, $count2, 0])) -and ($serverProducts[$count, $count2, 0] -ne $null))
                    { $productsInstalled.Add(($serverProducts[$count, $count2, 0])) | Out-Null }

                for ($count3 = 1; $count3 -le ($global:_maxItemsonServer - 2); $count3++)
                {
                    $itemVal2Found = [boolean] "false"
                    if ($serverProducts[$count, $count2, $count3] -ne $null)
                    {
                        if($itemsInstalled -notcontains ($serverProducts[$count, $count2, 0] + " : " + ($serverProducts[$count, $count2, $count3].Split(':')[0])) )
                        { 

                            $itemsInstalled.Add(($serverProducts[$count, $count2, 0]) + " : " + ($serverProducts[$count, $count2, $count3].Split(':')[0])) | Out-Null 
                        } 
                    } 
                }
            }
        }
		
        # let us get the max number of items per product 
        $count = $Servernum - 1
        $temptotalProducts = $totalProducts | Out-Null
        while ($count -ge 0)
        {
            while ($temptotalProducts -ge 0)
            {
                while ($thCount -ge 0)
                {
                    if ($serverProducts[$count, $temptotalProducts, $thCount] -ne $null)
                        { $itemCount = $itemCount + 1 }
                    $thCount--
                }
                if ($maxitemCount -lt $itemCount)
                    { $maxitemCount = $itemCount | Out-Null }
                $itemCount = 0
                $temptotalProducts = $temptotalProducts - 1
            }
            $count = $count - 1
        }
		
		# Now, the writing part $Write-Host 
		$global:HTMLContents += "<div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Installed Products on Servers</span><a class=`"expando`" href=`"#`"></a> </div>" 
		
		foreach ($tcp in $productsInstalled)
        {
            $star = [boolean] "false"
            # writing the Products in XML and HTML
            $global:HTMLContents += "<div class=`"container`"><div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Product Name: "+ $tcp +"</span><a class=`"expando`" href=`"#`"></a></div>" 
            $global:HTMLContents += "<div class=`"container`">"        
			
            foreach ($tcp0 in $itemsInstalled)
            {
                $global:HTMLContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">Item: "+ $tcp0.Split(':')[1].Trim() +"</span><a class=`"expando`" href=`"#`"></a></div>"
                $global:HTMLContents += "<div class=`"container`"><div class=`"he4i`"><table class=`"info`" >"
				
                for ($count = ($global:Servernum - 1); $count -ge 0; $count--)
                {
                    for ($count2 = ($global:_maxProductsonServer - 1); $count2 -ge 1; $count2--)
                    {
                        for ($count3 = ($global:_maxItemsonServer - 2); $count3 -ge 1; $count3--)
                        {
                            if ($global:serverProducts[$count, $count2, $count3] -ne $null)
							{
                                if ($tcp0.Split(':')[1].ToLower().Trim() -eq $serverProducts[$count,$count2, $count3].Split(':')[0].Trim().ToLower())
                                {
                                    $global:HTMLContents +=  "<tr><td><strong>" + $global:serverProducts[$count,0,0] + "</strong></td><td>"+ $serverProducts[$count, $count2, $count3].Split(':')[1] +"</td></tr>"
                                }
							}
                        }
                    }
                }

                $global:HTMLContents += "</table></div></div>"
            }
            $global:HTMLContents += "</div></div>"        
        }     
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeProdVersions2", $_)
    }
}

function o16writeFeatures
{
	try
	{
		$global:HTMLContents += "<div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Features</span><a class=`"expando`" href=`"#`"></a> </div><div class=`"container`">" 
		
		# where scope is farm
        $global:HTMLContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Scope: Farm</span><a class=`"expando`" href=`"#`"></a></div>"
        $global:HTMLContents +=  "<div class=`"container`"> <div class=`"he2i`">"
        $global:HTMLContents +=  "<table class=`"info`">"
        $global:HTMLContents +=  "<tr><td><strong>Id</strong></td><td><strong>Name</strong></td><td><strong>SolutionId</strong></td><td><strong>IsActive</strong></td><tr>"

        for ($i = 0; $i -lt $global:FeatureCount; $i++)
        {            
            $global:HTMLContents += "<tr><td>"+ $global:FarmFeatures[$i, 0] +"</td><td>"+ $global:FarmFeatures[$i, 1] +"</td><td>"+ $global:FarmFeatures[$i, 2] +"</td><td>"+ $global:FarmFeatures[$i, 3] +"</td><tr>"
        }

        $global:HTMLContents += "</table></div></div><div class=`"filler`"></div>"
		
		# where scope is site
        $global:HTMLContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Scope: Site</span><a class=`"expando`" href=`"#`"></a></div>"
        $global:HTMLContents +=  "<div class=`"container`"> <div class=`"he2i`">"
        $global:HTMLContents +=  "<table class=`"info`">"
        $global:HTMLContents +=  "<tr><td><strong>Id</strong></td><td><strong>Name</strong></td><td><strong>SolutionId</strong></td><td><strong>IsActive</strong></td><tr>"
		
        for ($i = 0; $i -lt $global:sFeatureCount; $i++)
        {
            $global:HTMLContents += "<tr><td>"+ $global:SiteFeatures[$i, 0] +"</td><td>"+ $global:SiteFeatures[$i, 1] +"</td><td>"+ $global:SiteFeatures[$i, 2] +"</td><td>"+ $global:SiteFeatures[$i, 3] +"</td><tr>"
        }
        $global:HTMLContents += "</table></div></div><div class=`"filler`"></div>"
		
		# where scope is web
        $global:HTMLContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Scope: Web</span><a class=`"expando`" href=`"#`"></a></div>"
        $global:HTMLContents +=  "<div class=`"container`"> <div class=`"he2i`">"
        $global:HTMLContents +=  "<table class=`"info`">"
        $global:HTMLContents +=  "<tr><td><strong>Id</strong></td><td><strong>Name</strong></td><td><strong>SolutionId</strong></td><td><strong>IsActive</strong></td><tr>"

        for ($i = 0; $i -lt $global:wFeatureCount; $i++)
        {
            $global:HTMLContents += "<tr><td>"+ $global:WebFeatures[$i, 0] +"</td><td>"+ $global:WebFeatures[$i, 1] +"</td><td>"+ $global:WebFeatures[$i, 2] +"</td><td>"+ $global:WebFeatures[$i, 3] +"</td><tr>"
        }

        $global:HTMLContents += "</table></div></div><div class=`"filler`"></div>"
		$global:HTMLContents += "</div>"

	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeFeatures", $_)
    }
}

function o16writeSolutions
{
	try
	{
            
            # Writing HTML            
            $global:htmlContents +=  "<div class=`"spsfrsettings`"> <div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Custom Solutions</span><a class=`"expando`" href=`"#`"></a></div>"
            $global:htmlContents += "<div class=`"container`"><div class=`"he2i`">"		
            # End of HTML


        if ($global:solutionCount -eq 0)
        {
            # Writing HTML            
            $global:htmlContents += "<table class=`"info`"><tr><td> There are no Custom solutions to report ! </td></tr></table>"
            $global:htmlContents += "</div></div><div class=`"filler`"></div>"
            $global:HTMLContents += "</div>"
            # End of HTML

			return
        }
        # Writing HTML Headers
        $global:htmlContents += "<table class=`"info`"><tr><td><strong>No.</strong></td><td><strong>Id</strong></td><td><strong>Name</strong></td><td><strong>Deployed on Web Apps</strong></td><td><strong>Last Operation Details</strong></td><td><strong>Deployed on Servers</strong></td></tr></table>"            		

        for ($count = 0; $count -le ($global:solutionCount - 1); $count++)
        {
            # Writing HTML            
            $global:htmlContents += "<table class=`"info`"><tr><td>"+ ($count + 1).ToString() +"</td><td>"+ $global:solutionProps[$count, 5] +"</td><td>" + $global:solutionProps[$count, 0] + "</td><td>" + $global:solutionProps[$count, 1] + "</td><td>" + $global:solutionProps[$count, 2] + "</td><td> " + $global:solutionProps[$count, 3] + " </td></tr></table>"            		
            # End of HTML	
        }

        # Writing HTML            		
        $global:htmlContents += "</div></div><div class=`"filler`"></div>"
        $global:HTMLContents += "</div>"
        # End of HTML
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeSolutions", $_)
    }	
}

function o16writeServiceApps
{
	try
	{ 
        
        #HTML
        $global:htmlContents += "<div class=`"spsfrsettings`"> <div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Service Applications</span><a class=`"expando`" href=`"#`"></a></div> <div class=`"container`">"

		$global:ServiceApps.GetEnumerator() | ForEach-Object {
		
		$isSearchSvcApp = 0	
        $isProjectSvcApp = 0          
        $prjApp=$_.Value
		$ServiceAppID = ($_.key.Split('|'))[0]
		$typeName = ($_.key.Split('|'))[1]		
        if($global:projectsvcApps.Id -eq $ServiceAppID) { $isProjectSvcApp=1}  	
	
		ForEach($searchAppId in $searchServiceAppIds)
		{
			if($searchAppId -eq $ServiceAppID) { $isSearchSvcApp = 1 }
		}
		
		if($isSearchSvcApp -eq 1)		
		{ 			

            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ $typeName +"</span><a class=`"expando`" href=`"#`"></a> </div> <div class=`"container`">"

            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Service Application Properties</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
            $tempHTMLstr = "<properties>" + $_.value + "</properties>"
            $xmldoc = New-Object -TypeName xml
            $xmldoc.LoadXml($tempHTMLstr)            
            $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText  
            $global:htmlContents += "</div></div>"

			
			#Writing the Search Service Status
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Search Service Status</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			try
			{ 
                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">Service Configuration</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $global:enterpriseSearchServiceStatus + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
                    $global:htmlContents += "</div></div>"
					
                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">Search Timer Job Definitions</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $global:enterpriseSearchServiceJobDefs + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, Schedule, Server                         
                    $global:htmlContents += "</div></div>"
			  }
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }
            $global:htmlContents += "</div>"
			
			#Writing the Active Topology
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Active Topology</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			try
			{
				$global:SearchActiveTopologyComponents.GetEnumerator() | ForEach-Object {
				$searchServiceAppID = ($_.key.Split('|'))[0]
				if($ServiceAppID -eq ($searchServiceAppID)) 
				{ 
					$props = $_.value
					$compName = ($_.key.Split('|'))[1]

                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">"+ $compName.Trim() +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $props + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
                    $global:htmlContents += "</div></div>"
				}
				}
			}
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }
            $global:htmlContents += "</div>"
			
			#Writing the Host Controllers
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Host Controllers</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			try
			{
				$global:SearchHostControllers.GetEnumerator() | ForEach-Object {
				$searchServiceAppID = ($_.key.Split('|'))[0]
				if($ServiceAppID -eq ($searchServiceAppID)) 
				{ 
					$props = $_.value
					$serverName = ($_.key.Split('|'))[1]
                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">"+ $serverName.Trim() +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $props + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
                    $global:htmlContents += "</div></div>"

				}
				}
			}
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }
            $global:htmlContents += "</div>"
            

			#Writing the Admin Component
			try
			{
				$global:SearchConfigAdminComponents.GetEnumerator() | ForEach-Object {
				if($ServiceAppID -eq ($_.key)) { $adminComponent = $_.value}
				}

                    $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Admin Component</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $adminComponent + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)            
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText  
                    $global:htmlContents += "</div></div>"
			}
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }

			# Writing Link Stores
			try
			{
				$global:SearchConfigLinkStores.GetEnumerator() | ForEach-Object {
				if($ServiceAppID -eq ($_.key)) { $storeValue = $_.value}
				}
                    $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Link Stores</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $storevalue + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)            
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText  
                    $global:htmlContents += "</div></div>"
			}
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }
			
			#Writing the Crawl Databases
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Crawl Databases</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			try
			{
				$global:SearchConfigCrawlDatabases.GetEnumerator() | ForEach-Object {
					$crawlComponent = ""
					$searchServiceAppID = ($_.key.Split('|'))[0]
					$crawlDatabaseID = ($_.key.Split('|'))[1]	
					if($ServiceAppID -eq $searchServiceAppID) 
					{ 
						$crawlComponent = $_.value				

                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Id: " + $crawlDatabaseID +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $crawlComponent + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
                    $global:htmlContents += "</div></div>"

					}
				}
			}
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }
            $global:htmlContents += "</div>"
			
			#Writing crawl rules
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Crawl Rules</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			try
			{
				$global:SearchConfigCrawlRules.GetEnumerator() | ForEach-Object {
					$searchServiceAppID = ($_.key.Split('|'))[0]
					$crawlRuleName = ($_.key.Split('|'))[1]	
					if($ServiceAppID -eq $searchServiceAppID) 
					{ 
						$crawlRule = $_.value				
                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Id: " + $crawlRuleName +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $crawlRule + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
                    $global:htmlContents += "</div></div>"

					}
				}
			}
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }
            $global:htmlContents += "</div>"
			
			#Writing the Query Site Settings
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Query Site Settings</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			try
			{
				$global:SearchConfigQuerySiteSettings.GetEnumerator() | ForEach-Object {
					$queryComponent = ""
					$searchServiceAppID = ($_.key.Split('|'))[0]
					$queryComponentID = ($_.key.Split('|'))[1]	
					if($ServiceAppID -eq $searchServiceAppID) 
					{ 
						$queryComponent = $_.value

                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Query Component ID: " + $queryComponentID +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $queryComponent + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
                    $global:htmlContents += "</div></div>"
					}
				}
			}
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }
            $global:htmlContents += "</div>"
			
			#Writing the Content Sources
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Content Sources</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			try
			{
				$global:SearchConfigContentSources.GetEnumerator() | ForEach-Object {
					$contentSource = ""
					$searchServiceAppID = ($_.key.Split('|'))[0]
					$contentSourceID = ($_.key.Split('|'))[1]	
					if($ServiceAppID -eq $searchServiceAppID) 
					{ 
						$contentSource = $_.value

                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Content Source ID: " + $contentSourceID +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    $tempHTMLstr = "<properties>" + $contentSource + "</properties>"
                    $xmldoc = New-Object -TypeName xml
                    $xmldoc.LoadXml($tempHTMLstr)
                    $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
                    $global:htmlContents += "</div></div>"

					}
				}
			}
			catch [System.Exception] 
		    {
				Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
		    }
            $global:htmlContents += "</div>"
            $global:htmlContents += "</div>" #This one needs to be pushed down	

	      }		
        elseif($isProjectSvcApp -eq 1)		
        {	 		        
            $cnt=0 
           
           #Writing Project Server Instance Information                   
            try
            {                                        
                        $global:projectInstances.GetEnumerator() | ForEach-Object {
					    $prjAppID = ($_.key.Split('|'))[0]
					    $prjName = ($_.key.Split('|'))[1]						
						$prjInst = $_.value				
                        $cnt++					
            }                                  	                                          
                                
                         										
        }
        catch [System.Exception] 
        {
		                Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		                Write-Output $_ | Out-File -FilePath $global:_logpath -Append
   		}

            #Writing Project Server PCS Settings
            try
            {    
            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Project Server Application" +"</span><a class=`"expando`" href=`"#`"></a> </div> <div class=`"container`"><div class=`"he2i`">"
            $global:htmlContents += "<table><tr><td>There appears to be a Project Server service application. Details about the service application can be found in the accompanying XML. </td></tr></table>"       
            $global:htmlContents += "</div></div><div class=`"filler`"></div> "                               
                                                     
            }
            catch [System.Exception] 
            {
		                Write-Host " ******** Exception caught. Check the log file for more details. ******** "
		                Write-Output $_ | Out-File -FilePath $global:_logpath -Append
       		}
        }
		elseif($isSearchSvcApp -eq 0 -and $isProjectSvcApp -eq 0)
		{ 
            $tempHTMLstr = "<properties>" + $_.value + "</properties>"
            $xmldoc = New-Object -TypeName xml
            $xmldoc.LoadXml($tempHTMLstr)

            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ $typeName +"</span><a class=`"expando`" href=`"#`"></a> </div> <div class=`"container`"><div class=`"he2i`">"
            $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
            $global:htmlContents += "</div></div><div class=`"filler`"></div> "
        }
		}		
        $global:HTMLContents += "</div>"       
        
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeServiceApps", $_)
    }
}

function o16writeSPServiceApplicationPools
{
	try
	{
		if($global:serviceAppPoolCount -eq 0 ) { return }
            
        $global:htmlContents += "<div class=`"spsfrsettings`"><div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`"> Service Application Pools </span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
		$global:SPServiceApplicationPools.GetEnumerator() | ForEach-Object {
		$serviceAppPoolID = $_.key
		$serviceAppPoolstr = $_.value
	
        $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "App Pool ID: " + $serviceAppPoolID + "</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"> <div class=`"he2i`">"
        $tempHTMLstr = "<properties>" + $serviceAppPoolstr + "</properties>"
        $xmldoc = New-Object -TypeName xml
        $xmldoc.LoadXml($tempHTMLstr)
        $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
        $global:htmlContents += "</div></div>"

		}
        $global:htmlContents += "</div></div>"
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeSPServiceApplicationPools", $_)
    }

}

function o16writeSPServiceApplicationProxies
{
	try
	{
		if($global:serviceAppProxyCount -eq 0)
		{ 		return 		}

		try
		{
            
            $global:htmlContents += "<div class=`"spsfrsettings`"><div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`"> Service Application Proxies </span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			$global:SPServiceAppProxies.GetEnumerator() | ForEach-Object {
			$searchServiceAppProxyID = ($_.key.Split('|'))[0]
			$TypeName = ($_.key.Split('|'))[1]	
			$proxy = $_.value				

            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "App Proxy ID: " + $searchServiceAppProxyID + "     TypeName: " + $TypeName +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"> <div class=`"he2i`">"
            $tempHTMLstr = "<properties>" + $proxy + "</properties>"
            $xmldoc = New-Object -TypeName xml
            $xmldoc.LoadXml($tempHTMLstr)
            $global:htmlContents += $xmldoc.properties.GetEnumerator() | ConvertTo-Html -Fragment -Property Name, InnerText                         
            $global:htmlContents += "</div></div>"
			}
		}
		catch [System.Exception] 
	    {
			Write-Host " ******** Exception caught. Check the log file for more details. ******** "
	        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
	    }		
        $global:htmlContents += "</div></div>"
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeSPServiceApplicationProxies", $_)
    }
}

function o16writeSPServiceApplicationProxyGroups
{
	try
	{
		if($global:SvcAppProxyGroupCount -eq 0)		{ 		return 		}
		
		try
		{
            $global:htmlContents += "<div class=`"spsfrsettings`"><div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`"> Service Application Proxy Groups </span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
			$global:SPServiceAppProxyGroups.GetEnumerator() | ForEach-Object {
			$GroupID = ($_.key.Split("|"))[0]
			$FriendlyName = ($_.key.Split("|"))[1]
			$GroupXML = $_.value				

            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Group ID: " + $GroupID + "     FriendlyName: " + $FriendlyName +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"> <div class=`"he2i`">"
            $tempHTMLstr =  get-SPServiceApplicationProxyGroup | ? { $_.id -match "c8c88b49-1fe1-41d6-9bbd-3bd6ccac3c21" } | select proxies | ft -Hide -AutoSize | Out-String -width 1000
            $global:htmlContents += "<table> <tr> <th> Proxy Group ID </th> <th>Associated Service Applications</th> </tr> <tr><td>" + $GroupID + "</td><td>" + $tempHTMLstr + "</td></tr>></table>" 
            $global:htmlContents += "</div></div>"

			}
		}
		catch [System.Exception] 
	    {
			Write-Host " ******** Exception caught. Check the log file for more details. ******** "
	        Write-Output $_ | Out-File -FilePath $global:_logpath -Append
	    }
		$global:htmlContents += "</div></div>"
		
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeSPServiceApplicationProxyGroups", $_)
    }
}

function o16writeWebApps()
{
	try
	{
			$contentDBs = ""
            $zonestring = ""
			$global:htmlContents += "<div class=`"spsfrsettings`"> <div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Web Applications</span><a class=`"expando`" href=`"#`"></a></div> <div class=`"container`">"

			for($wcount = 0; $wcount -lt $global:WebAppnum; $wcount++ )
			{
                        for ($dbcount = 0; $dbcount -le [System.Convert]::ToInt32(($global:WebAppDetails[$wcount, 3])); $dbcount++)
                        {
                            if (($contentDBs -eq "") -and (($dbcount + 1) -eq [System.Convert]::ToInt32(($global:WebAppDetails[$wcount, 3]))))
                            {    
								$contentDB = $global:ContentDBs[$wcount, $dbcount] 
							}								
                        }
               
                $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ $global:WebAppDetails[$wcount, 2] +"</span><a class=`"expando`" href=`"#`"></a> </div> <div class=`"container`">"
                $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Web Application Properties</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                $global:htmlContents += "<table><tr> <th>App Pool</th><th>App Pool ID</th><th>Service App Proxy Group</th> </tr> <tr><td>"+ $global:WebAppDetails[$wcount, 6] +" </td><td>" + $global:WebAppDetails[$wcount, 7] + "</td><td>" + $global:WebAppDetails[$wcount, 9] + "</td></table>"
                $global:htmlContents += "</div></div>"
                
                
                $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Authentication Providers</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"
                for($i = 0; $i -le 4; $i++) 
                {  
                    switch($i)
					{
						0  { $zonestring = "Default" }
						1 { $zonestring = "Intranet" }
                        2 { $zonestring = "Internet" }
                        3 { $zonestring = "Custom" }
                        4 { $zonestring = "Extranet" }
						default { $zonestring = "Default" }
					} 

                    $global:htmlContents += "<div class=`"he2`"><span class=`"sectionTitle`" tabindex=`"0`">"+ $zonestring + " Zone" +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
                    if($global:WebAppAuthProviders[$wcount, $i].Length -gt 16) 
                    { $global:htmlContents += $global:WebAppAuthProviders[$wcount, $i]  }
                    else {  $global:htmlContents += "<table><tr><td>Not configured.</td></tr></table>" }
                    $global:htmlContents += "</div></div>"
                }
                $global:htmlContents += "</div>"
                
            #Writing AAMs
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Alternate Access Mappings</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
            $global:htmlContents += $global:WebAppAAMs[$wcount]
            $global:htmlContents += "</div></div>"

            #Writing Content DB basic info
            $global:htmlContents += "<div class=`"he1h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">Content Databases</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he4i`">"
            $global:htmlContents += get-spcontentdatabase -WebApplication $global:WebAppDetails[$wcount, 2] | select id, DisplayName, CurrentSiteCount, buildversion, NeedsUpgrade | convertto-html -fragment 
            $global:htmlContents += "</div></div></div>"
			}	
            $global:HTMLContents += "</div></div>"
            $global:HTMLContents += "</div>" # this one needs to remain at bottom
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeWebApps", $_)
    }
}

function o16writeAAMsnAPs()
{
	try
	{
		$AllZones = [Microsoft.SharePoint.Administration.SPUrlZone]::Default, 
							[Microsoft.SharePoint.Administration.SPUrlZone]::Intranet, 
							[Microsoft.SharePoint.Administration.SPUrlZone]::Internet, 
							[Microsoft.SharePoint.Administration.SPUrlZone]::Custom, 
							[Microsoft.SharePoint.Administration.SPUrlZone]::Extranet

		for($wcount = 0; $wcount -lt $global:WebAppnum; $wcount++)
		{
		
			for ($zones = 1; $zones -le 5; $zones++)
			{				
				$tempstr = $WebAppPublicAAM[$wcount, $zones]
				if($tempstr -ne $null) 
				{ 
					$tempstr = $tempstr.Trim()
					if($tempstr -ne "")
					{
					}
				}
				
				$tempstr = $WebAppInternalAAMURL[$wcount, $zones]
				if($tempstr -ne $null)
				{
					$tempstr = $tempstr.Trim()
					if($tempstr -ne "")
					{
					}
				}
				
				$tempstr = $WebAppAuthProviders[$wcount, $zones]
				if($tempstr -ne $null)
				{
					$tempstr = $tempstr -split ']'
				}
			}		
		}
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeAAMsnAPs", $_)
    }
}

function o16writeContentDBs()
{
	try
	{
        $global:htmlContents += "<div class=`"spsfrsettings`"><div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`"> Content Databases [Full List] </span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he2i`">"
        $global:htmlContents += "<table> <tr> <th>ID</th> <th>Name</th><th>Web Application</th><th>SQL Service Instance</th><th>Site Count</th><th>Disk Space Required for Backup</th><th>Timer Locked By</th> <th>NeedUpgrade</th><th>RBS Enabled</th><th>RBS ActiveProviderName</th></tr>"  
		for($count = 0; $count -lt $global:totalContentDBCount; $count++)
		{
			for ($count2 = ($global:Servernum - 1); $count2 -ge 0; $count2--)
			{
				if ($global:ServersId[$count2] -eq $global:ContentDBProps[$count, 6]) 
				{ 
				}
			}
			
            $global:htmlContents += "<tr><td>" + $global:ContentDBProps[$count, 2] + "</td><td>" + $global:ContentDBProps[$count, 0] + "</td><td>" + $global:ContentDBProps[$count, 1] + "</td><td>" + $global:ContentDBProps[$count, 3] + "</td><td>" + $global:ContentDBProps[$count, 4] + "</td><td>" + $global:ContentDBProps[$count, 5] + "</td><td>" + $global:ContentDBProps[$count, 6] + "</td><td>" + $global:ContentDBProps[$count, 7] + "</td><td>" + $global:ContentDBProps[$count, 8] + "</td><td>" + $global:ContentDBProps[$count, 9] + "</td></tr>"		            
		}        
        $global:htmlContents += "</table>"
        $global:htmlContents += "</div></div></div>"
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15writeContentDBs", $_)
    }    
}

function o16writeCDConfig()
{
	try
	{
		#Writing General Information

		#Writing Paths
		$global:CDPaths.GetEnumerator() | ForEach-Object {
		$PathId = ($_.key.Split('|'))[0]
		$PathName = ($_.key.Split('|'))[1]
		$PathName = $PathName.Trim()
		$tempstr = $_.value
		
		$global:CDJobs.GetEnumerator() | ForEach-Object {
		$PathId2 = ($_.key.Split('|'))[0]
		$JobName = ($_.key.Split('|'))[2]
		$JobName = $JobName.Trim()
		
		if($PathId2 -eq $PathId)
		{
		}
		
		}
		
		}	
	}
		catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        #global:HandleException("o15writeSPServiceApplicationPools", $_)
    }
}

function o16writeHealthReport()
{
	try
	{
	
		if (($global:HealthReport0.Count -eq 0) -and ($global:HealthReport1.Count -eq 0) -and ($global:HealthReport2.Count -eq 0) -and ($global:HealthReport3.Count -eq 0))
			{ exit }
				
		$global:htmlContents += "<div class=`"spsfrsettings`"><div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`"> Health Analyzer Report </span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`">"

		# We iterate through each Severity separately because the rules (their ids) are run and presented sporadic in CA.
		#Writing 0 - Rule Execution Failures
		if($global:HealthReport0.Count -gt 0)
		{
            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Rule Execution Errors" +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"> <div class=`"he2i`">"
            $global:htmlContents += "<table>"
            $global:htmlContents += "<tr> <th>ID</th><th>Title</th><th>Failing Servers</th><th>Failing Services</th><th>Modified</th> </tr>"
			$global:HealthReport0.GetEnumerator() | ForEach-Object {
			$id = $_.key
			$title = $_.value.Split("||")[0]
			$failingServers = $_.value.Split("||")[2]
			$failingServices = $_.value.Split("||")[4]
			$Modified = $_.value.Split("||")[6]

            $global:htmlContents += "<tr> <td>"+ $id +"</td> <td>"+ $title +"</td> <td>"+ $failingServers +"</td> <td>"+ $failingServices +"</td> <td>"+ $Modified +"</td> </tr>"

			}	
            $global:htmlContents += "</table>"
            $global:htmlContents += "</div></div>"	
		}
		
		#Writing 1 - Errors
		if($global:HealthReport1.Count -gt 0)
		{
            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Errors" +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"> <div class=`"he2i`">"
            $global:htmlContents += "<table>"
            $global:htmlContents += "<tr> <th>ID</th><th>Title</th><th>Failing Servers</th><th>Failing Services</th><th>Modified</th> </tr>"
			$global:HealthReport1.GetEnumerator() | ForEach-Object {
			$id = $_.key
			$title = $_.value.Split("||")[0]
			$failingServers = $_.value.Split("||")[2]
			$failingServices = $_.value.Split("||")[4]
			$Modified = $_.value.Split("||")[6]
			
            $global:htmlContents += "<tr> <td>"+ $id +"</td> <td>"+ $title +"</td> <td>"+ $failingServers +"</td> <td>"+ $failingServices +"</td> <td>"+ $Modified +"</td> </tr>"
			}		
            $global:htmlContents += "</table>"
            $global:htmlContents += "</div></div>"	
		}
		
		#Writing 2 - Warning
		if($global:HealthReport2.Count -gt 0)
		{
            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Warnings" +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"> <div class=`"he2i`">"
            $global:htmlContents += "<table>"
            $global:htmlContents += "<tr> <th>ID</th><th>Title</th><th>Failing Servers</th><th>Failing Services</th><th>Modified</th> </tr>"
			$global:HealthReport2.GetEnumerator() | ForEach-Object {
			$id = $_.key
			$title = $_.value.Split("||")[0]
			$failingServers = $_.value.Split("||")[2]
			$failingServices = $_.value.Split("||")[4]
			$Modified = $_.value.Split("||")[6]
            
            $global:htmlContents += "<tr> <td>"+ $id +"</td> <td>"+ $title +"</td> <td>"+ $failingServers +"</td> <td>"+ $failingServices +"</td> <td>"+ $Modified +"</td> </tr>"

			}		
            $global:htmlContents += "</table>"
            $global:htmlContents += "</div></div>"
		}
		
		#Writing 3 - Information
		if($global:HealthReport3.Count -gt 0)
		{
            $global:htmlContents += "<div class=`"he0h_expanded`"><span class=`"sectionTitle`" tabindex=`"0`">"+ "Information" +"</span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"> <div class=`"he2i`">"
            $global:htmlContents += "<table>"
            $global:htmlContents += "<tr> <th>ID</th><th>Title</th><th>Failing Servers</th><th>Failing Services</th><th>Modified</th> </tr>"
			$global:HealthReport3.GetEnumerator() | ForEach-Object {
			$id = $_.key
			$title = $_.value.Split("||")[0]
			$failingServers = $_.value.Split("||")[2]
			$failingServices = $_.value.Split("||")[4]
			$Modified = $_.value.Split("||")[6]

            $global:htmlContents += "<tr> <td>"+ $id +"</td> <td>"+ $title +"</td> <td>"+ $failingServers +"</td> <td>"+ $failingServices +"</td> <td>"+ $Modified +"</td> </tr>"
			}		
            $global:htmlContents += "</table>"
            $global:htmlContents += "</div></div>"
		}
		
        $global:htmlContents += "</div></div>"
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        #global:HandleException("o15writeSPServiceApplicationPools", $_)
    }
}

function o16writeTimerJobs()
{
	try
	{
    
		if ($global:timerJobs.Count -eq 0)
		{ exit }
        $global:htmlContents += "<div class=`"spsfrsettings`"><div class=`"he0_expanded`"><span class=`"sectionTitle`" tabindex=`"0`"> Timer Jobs </span><a class=`"expando`" href=`"#`"></a></div><div class=`"container`"><div class=`"he2i`">"
        $global:htmlContents += "<table> <tr> <th>ID</th> <th>Title</th><th>WebApplication</th><th>Schedule</th><th>LastRunTime</th><th>isDisabled</th><th>LockType</th></tr>"  
				
		#Writing them
			$global:timerJobs.GetEnumerator() | ForEach-Object {
		
			$id = $_.value.Split("||")[0]
			$title = $_.value.Split("||")[2]
			$webapplication = $_.value.Split("||")[4]
			$schedule = $_.value.Split("||")[6]
			$lastruntime = $_.value.Split("||")[8]
			$isdisabled = $_.value.Split("||")[10]
			$locktype = $_.value.Split("||")[12]
            
            $global:htmlContents += "<tr><td>" + $id + "</td><td>" + $title + "</td><td>" + $webapplication + "</td><td>" + $schedule + "</td><td>" + $lastruntime + "</td><td>" + $isdisabled + "</td><td>" + $locktype + "</td></tr>"		            

			}
			$global:htmlContents += "</table>"
        $global:htmlContents += "</div></div></div>"
	}
	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        #global:HandleException("o15writeSPServiceApplicationPools", $_)
    }
}

function o16WriteEndHTML()
{
    try
    {
        #Completing HTML Write
        $global:HTMLpath = [Environment]::CurrentDirectory + "\o16SPSFarmReport{0}{1:d2}{2:d2}-{3:d2}{4:d2}" -f (Get-Date).Day,(Get-Date).Month,(Get-Date).Year,(Get-Date).Second,(Get-Date).Millisecond + ".HTML"
        $global:htmlContents += $global:HTMLTrail
        Write-Output $global:htmlContents | Out-File -FilePath $global:HTMLpath
    }
    	catch [System.Exception] 
    {
		Write-Host " ******** Exception caught. Check the log file for more details. ******** "
        global:HandleException("o15WriteEndXML", $_)
    }
}

function global:HandleException([string]$functionName,[System.Exception]$err)
{
	$global:exceptionDetails = $global:exceptionDetails + "********* Exception caught:" + $functionName + " , " + $err
	Write-Output $_ | Out-File -FilePath $global:_logpath -Append
}



$dtime = " Starting run of SPSFarmReport at " + (Get-Date).ToString()
Write-Output "---------------------------------------------------------------------------------" | Out-File -FilePath $global:_logpath -Append
Write-Output  $dtime | Out-File -FilePath $global:_logpath -Append
Write-Output "---------------------------------------------------------------------------------" | Out-File -FilePath $global:_logpath -Append

o16WriteInitialHTML
Write-Host o16WriteInitialHTML

$status = o16farmConfig
$dtime = " Completed running o16farmConfig at " + (Get-Date).ToString()
Write-Host o16farmConfig
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16WriteFarmGenSettings 
	$dtime = " Completed running o16WriteFarmGenSettings at " + (Get-Date).ToString()
	Write-Host o16WriteFarmGenSettings
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
}

$status = o16enumServers
$dtime = " Completed running o16enumServers at " + (Get-Date).ToString()
Write-Output o16enumServers
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeServers 
	$dtime = " Completed running o16writeServers at " + (Get-Date).ToString()
	Write-Host o16writeServers
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
}

$status = o16enumProdVersions
$dtime = " Completed running o16enumProdVersions at " + (Get-Date).ToString()
Write-Host o16enumProdVersions
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeProdVersions2 
	$dtime = " Completed running o16writeProdVersions2 at " + (Get-Date).ToString()
	Write-Host o16writeProdVersions2
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append	
}

$status = o16enumFeatures
$dtime = " Completed running o16enumFeatures at " + (Get-Date).ToString()
Write-Host o16enumFeatures
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeFeatures 
	$dtime = " Completed running o16writeFeatures at " + (Get-Date).ToString()
	Write-Host o16writeFeatures
	Write-Output $dtime	| Out-File -FilePath $global:_logpath -Append	
}


$status = o16enumSolutions
$dtime = " Completed running o16enumSolutions at " + (Get-Date).ToString()
Write-Host o16enumSolutions
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeSolutions 
	$dtime = " Completed running o16writeSolutions at " + (Get-Date).ToString()
	Write-Host o16writeSolutions
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
}


$status = o16enumSvcApps 
$dtime = " Completed running o16enumSvcApps at " + (Get-Date).ToString()
Write-Host  o16enumSvcApps
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append

if($status -eq 1) { o16enumSPSearchServiceApps }
Write-Host o16enumSPSearchServiceApps
if($status -eq 1) { o16enumSPSearchService }
Write-Host o16enumSPSearchService
if($status -eq 1) { o16enumHostControllers }
Write-Host o16enumHostControllers
if($status -eq 1) { o16enumSearchActiveTopologies }
Write-Host o16enumSearchActiveTopologies
if($status -eq 1) { o16enumSearchConfigAdminComponents }
Write-Host o16enumSearchConfigAdminComponents
if($status -eq 1) { o16enumSearchConfigLinkStores }
Write-Host o16enumSearchConfigLinkStores
if($status -eq 1) { o16enumSearchConfigCrawlDatabases }
Write-Host o16enumSearchConfigCrawlDatabases
if($status -eq 1) { o16enumSearchConfigCrawlRules }
Write-Host o16enumSearchConfigCrawlRules
if($status -eq 1) { o16enumSearchConfigQuerySiteSettings }
Write-Host o16enumSearchConfigQuerySiteSettings
if($status -eq 1) { o16enumSearchConfigContentSources }
Write-Host o16enumSearchConfigContentSources
if($status -eq 1) { o16enumProjectServiceApps }
Write-Host o16enumProjectServiceApps


if($status -eq 1) 
{ 
	o16writeServiceApps 
	$dtime = " Completed running o16writeServiceApps at " + (Get-Date).ToString()
	Write-Host o16writeServiceApps
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
} 


$status = o16enumSPServiceApplicationPools
$dtime = " Completed running o16enumSPServiceApplicationPools at " + (Get-Date).ToString()
Write-Host o16enumSPServiceApplicationPools
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeSPServiceApplicationPools 
	$dtime = " Completed running o16writeSPServiceApplicationPools at " + (Get-Date).ToString()
	Write-Host o16writeSPServiceApplicationPools	
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
}


$status = o16enumSPServiceApplicationProxies
$dtime = " Completed running o16enumSPServiceApplicationProxies at " + (Get-Date).ToString()
Write-Host o16enumSPServiceApplicationProxies
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeSPServiceApplicationProxies 
	$dtime = " Completed running o16writeSPServiceApplicationProxies at " + (Get-Date).ToString()
	Write-Host o16writeSPServiceApplicationProxies 
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
}

$status = o16enumSPServiceApplicationProxyGroups
$dtime = " Completed running o16enumSPServiceApplicationProxyGroups at " + (Get-Date).ToString()
Write-Host o16enumSPServiceApplicationProxyGroups
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeSPServiceApplicationProxyGroups 
	$dtime = " Completed running o16writeSPServiceApplicationProxyGroups at " + (Get-Date).ToString()
	Write-Host o16writeSPServiceApplicationProxyGroups 
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
}



$status = o16enumWebApps
$dtime = " Completed running o16enumWebApps at " + (Get-Date).ToString()
Write-Host o16enumWebApps
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeWebApps 
	$dtime = " Completed running o16writeWebApps at " + (Get-Date).ToString()
	Write-Host o16writeWebApps  
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
	
	<#o16writeAAMsnAPs
	$dtime = " Completed running o16writeAAMsnAPs at " + (Get-Date).ToString()
	Write-Host o16writeAAMsnAPs  
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append #>
}

$status = o16enumContentDBs
$dtime = " Completed running o16enumContentDBs at " + (Get-Date).ToString()
Write-Host o16enumContentDBs
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeContentDBs 
	$dtime = " Completed running o16writeContentDBs at " + (Get-Date).ToString()
	Write-Host o16writeContentDBs   
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
}

if($isFoundationOnlyInstalled -eq $false)
{
	$status = o16enumCDConfig
	$dtime = " Completed running o16enumCDConfig at " + (Get-Date).ToString()
	Write-Host o16enumCDConfig
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
	if($status -eq 1) 
	{ 
		o16writeCDConfig 
		$dtime = " Completed running o16writeCDConfig at " + (Get-Date).ToString()
		Write-Host o16writeCDConfig   
		Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
	}
}

$status = o16enumHealthReport
$dtime = " Completed running o16enumHealthReport at " + (Get-Date).ToString()
Write-Host o16enumHealthReport
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeHealthReport
	$dtime = " Completed running o16writeHealthReport at " + (Get-Date).ToString()
	Write-Host o16writeHealthReport - this usually takes longer to complete
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
}

$status = o16enumTimerJobs
$dtime = " Completed running o16enumTimerJobs at " + (Get-Date).ToString()
Write-Host o16enumTimerJobs
Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
if($status -eq 1) 
{ 
	o16writeTimerJobs
	$dtime = " Completed running o16writeTimerJobs at " + (Get-Date).ToString()
	Write-Host o16writeTimerJobs 
	Write-Output $dtime | Out-File -FilePath $global:_logpath -Append
} 

Write-Output $dtime | Out-File -FilePath $global:_logpath -Append

o16WriteEndHTML
Write-Host o16WriteEndHTML


$dtime = " Ending run of SPSFarmReport at " + (Get-Date).ToString()
Write-Output "---------------------------------------------------------------------------------" | Out-File -FilePath $global:_logpath -Append
Write-Output  $dtime | Out-File -FilePath $global:_logpath -Append
Write-Output "---------------------------------------------------------------------------------" | Out-File -FilePath $global:_logpath  -Append

Write-Host "---------------------------------------------------------------------------------" 
Write-Host  $dtime.Trim()
Write-Host HTML is generated in $global:HTMLpath -ForegroundColor DarkGreen
Write-Host Log file written to $global:_logpath -ForegroundColor DarkGreen
Write-Host The path to write files is based on the value of the [Environment]::CurrentDirectory environment variable.
Write-Host "---------------------------------------------------------------------------------" 