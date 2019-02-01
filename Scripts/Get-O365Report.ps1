<#
.SYNOPSIS
 Returns a Microsoft Graph Reporting API report for an Office365 Tenant
 
.DESCRIPTION
 Using an native App registered in Azure AD and an authorized Office 365 admin this script calls the Microsoft Graph Reporting API
 and returns the desired report type as a system.array object. The types of reports available are documented at https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report

.Example
$cred = Get-Credential

.\Get-Office365Report.ps1 `
    -TenantName "contoso.onmicrosoft.com" `
    -ClientID "df4d5697-2465-49e3-90b1-d029e609e33f" `
    -RedirectURI "urn:foo" `
    -WorkLoad OneDrive `
    -ReportType getOneDriveUsageStorage `
    -Cred $cred `
    -Period D180 `
    -Date 2017-10-26 `
    -Verbose
 
.PARAMETER TenantName
Tenant name in the format contoso.onmicrosoft.com

.PARAMETER ClientID
AppID for the App registered in AzureAD for the purpose of accessing the reporting API

.PARAMETER RedirectURI
ReplyURL for the App registered in AzureAD for the purpose of accessing the reporting API

.PARAMETER WorKload
Service in Office365 for which to provide report options. Used to provide a usable parameter set for the ReportType parameter

.PARAMETER ReportType
Report to retrieve see details at https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report.
This script does not currently support the following script types:
getMailboxUsageUserDetail
getOffice365GroupsActivityUserDetail
getOneDriveUsageUserDetail
getSharePointSiteUsageUserDetail
getYammerGroupsActivityUserDetail

.PARAMETER Period
Time period for the report in days. Allowed values: D7,D30,D90,D180
Period is not supported for reports starting with getOffice365Activations and will be ignored

.PARAMETER Credential
PSCredential object for a with access to view reports. If this is not provided the user will be prompted to enter their credentials

.PARAMETER Date
Specifies the day to a view of the users that performed an activity on that day. Must have a format of YYYY-MM-DD.
Only available for the last 30 days and is ignored unless view type is Detail
Date is not supported for the following report types: "getMailboxUsage*","getOffice365Activations*", "getSfbOrganizerActivity*" and will be ignored

.OUTPUTS
Returns an system.array object that is a representation of a Microsoft Graph API Report Object

.NOTES
To register the App (ClientID)
1) Login to Portal.Azure.Com
2) Navigate to "Azure Active Directory" > "App Registrations"
3) Click "New Application Registration"
4) Give your application a friendly name, Select application type "native", and enter a redirect URL i the format urn:foo and click create
5) Click on the App > Required Permissions
6) Click Add and select the "Microsoft Graph" API
7) Grant the App the "Read All Usage Reports" permission
8) Copy the Application ID and use that for ClientID parameter in this script
9) Copy the Redirect URI and use that for the RedirectURI parameter in this script

This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree:
(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and 
(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.

Version History:
## 4/18/2017 ##
Intial release

## 10/26/2017 ##
Updated script to use indivdiual APIs for each view as per documentation update (https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report#changes-to-the-reports-apis).
This script does not currently support the following report types types:
getMailboxUsageUserDetail
getOffice365GroupsActivityUserDetail
getOneDriveUsageUserDetail
getSharePointSiteUsageUserDetail
getYammerGroupsActivityUserDetail

## 10/30/2017 ##
Updated script for new report names noted in documentation 10/26
OLD                                  --> NEW
getMailboxUsageUserDetail            --> getMailboxUsageDetail
getOneDriveUsageUserDetail           --> getOneDriveUsageAccountDetail
getSharePointSiteUsageUserDetail     --> getSharePointSiteUsageDetail
getYammerGroupsActivityUserDetail    --> getYammerGroupsActivityDetail
getOffice365GroupsActivityUserDetail --> getOffice365GroupsActivityDetail 

#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true)]
    $TenantName,

    [Parameter(Mandatory=$true)]
    $ClientID,

    [Parameter(Mandatory=$true)]
    $RedirectURI,
     
    [Parameter(Mandatory=$true,Position=2)]
    [ValidateSet(
    "Exchange",
    "Groups",
    "OneDrive",
    "SharePoint",
    "Skype",
    "Tenant",
    "Yammer"
    )]
    $WorkLoad,  

    [Parameter(Mandatory=$false,Position=3)]
    [ValidateSet(
    "D7",
    "D30",
    "D90",
    "D180")]
    $Period,

    [Parameter(Mandatory=$false,Position=5)]
    $Date,

    [Parameter(Mandatory=$false,Position=4)]
    [PSCredential]$Credential

)
DynamicParam {
            # Set the dynamic parameters' name
            $ParameterName = 'ReportType'
            
            # Create the dictionary 
            $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

            # Create the collection of attributes
            $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            
            # Create and set the parameters' attributes
            $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
            $ParameterAttribute.Mandatory = $true
            $ParameterAttribute.Position = 1

            # Add the attributes to the attributes collection
            $AttributeCollection.Add($ParameterAttribute)

            # Generate and set the ValidateSet
            If ($Workload -eq "Exchange"){$arrSet = @("EmailActivity","getEmailActivityUserDetail","getEmailActivityCounts","getEmailActivityUserCounts","getEmailAppUsageUserDetail","getEmailAppUsageAppsUserCounts","getEmailAppUsageUserCounts","getEmailAppUsageVersionsUserCounts","getMailboxUsageDetail","getMailboxUsageMailboxCounts","getMailboxUsageQuotaMailboxStatusCounts","getMailboxUsageStorage")}
            If ($Workload -eq "Groups"){$arrSet = @("getOffice365GroupsActivityDetail","getOffice365GroupsActivityCounts","getOffice365GroupsActivityGroupCounts","getOffice365GroupsActivityStorage","getOffice365GroupsActivityFileCounts")}
            If ($Workload -eq "OneDrive"){$arrSet = @("getOneDriveActivityUserDetail","getOneDriveActivityUserCounts","getOneDriveActivityFileCounts","getOneDriveUsageAccountDetail","getOneDriveUsageAccountCounts","getOneDriveUsageFileCounts","getOneDriveUsageStorage")}
            If ($Workload -eq "SharePoint"){$arrSet = @("getSharePointActivityUserDetail","getSharePointActivityFileCounts","getSharePointActivityUserCounts","getSharePointActivityPages","getSharePointSiteUsageDetail","getSharePointSiteUsageFileCounts","getSharePointSiteUsageSiteCounts","getSharePointSiteUsageStorage","getSharePointSiteUsagePages")}
            If ($Workload -eq "Skype"){$arrSet = @("getSkypeForBusinessActivityUserDetail","getSkypeForBusinessActivityCounts","getSkypeForBusinessActivityUserCounts","getSkypeForBusinessDeviceUsageUserDetail","getSkypeForBusinessDeviceUsageDistributionUserCounts","getSkypeForBusinessDeviceUsageUserCounts","getSkypeForBusinessOrganizerActivityCounts","getSkypeForBusinessOrganizerActivityUserCounts","getSkypeForBusinessOrganizerActivityMinuteCounts","getSkypeForBusinessParticipantActivityCounts","getSkypeForBusinessParticipantActivityUserCounts","getSkypeForBusinessParticipantActivityMinuteCounts","getSkypeForBusinessPeerToPeerActivityCounts","getSkypeForBusinessPeerToPeerActivityUserCounts","getSkypeForBusinessPeerToPeerActivityMinuteCounts")}
            If ($Workload -eq "Tenant"){$arrSet = @("getOffice365ActivationsUserDetail","getOffice365ActivationCounts","getOffice365ActivationsUserCounts","getOffice365ActiveUserDetail","getOffice365ActiveUserCounts","getOffice365ServicesUserCounts")}
            If ($Workload -eq "Yammer"){$arrSet = @("getYammerActivityUserDetail","getYammerActivityCounts","getYammerActivityUserCounts","getYammerDeviceUsageUserDetail","getYammerDeviceUsageDistributionUserCounts","getYammerDeviceUsageUserCounts","getYammerGroupsActivityDetail","getYammerGroupsActivityGroupCounts","getYammerGroupsActivityCounts")}
            

            #$arrSet = Get-ChildItem -Path .\ -Directory | Select-Object -ExpandProperty FullName
            $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)

            # Add the ValidateSet to the attributes collection
            $AttributeCollection.Add($ValidateSetAttribute)

            # Create and return the dynamic parameter
            $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
            $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
            return $RuntimeParameterDictionary
    }

Begin {
    # Bind the parameter to a friendly variable
    $Report = $PsBoundParameters[$ParameterName]
}

#Start the loading of the rest of the script
Process{

    #If the credential object is empty, prompt the user for credentials
    if(!$Credential) {$Credential = Get-Credential}
    
    function Get-AuthToken
    {
    <#
    .SYNOPSIS
     Gets an OAuth token for use with the Microsoft Graph API
 
    .DESCRIPTION
     Gets an OAuth token for use with the Microsoft Graph API

    .EXAMPLE
     Get-AuthToken -TenantName "contoso.onmicrosoft.com" -clientId "74f0e6c8-0a8e-4a9c-9e0e-4c8223013eb9" -redirecturi "urn:ietf:wg:oauth:2.0:oob" -resourceAppIdURI "https://graph.microsoft.com"
 
    .PARAMETER TentantName
    Tenant name in the format <tenantname>.onmicrosoft.com

    .PARAMETER clientID
    The clientID or AppID of the native app created in AzureAD to grant access to the reporting API

    .Parameter redirecturi
    The replyURL of the native app created in AzureAD to grant access to the reporting API

    .Parameter resourceAppIDURI
    protocol and hostname for the endpoint you are accessing. For the Graph API enter "https://graph.microsoft.com"
 
    .NOTES
    Inital authentication sample from:
    https://blogs.technet.microsoft.com/paulomarques/2016/03/21/working-with-azure-active-directory-graph-api-from-powershell/

    #>
           param
           (
                  [Parameter(Mandatory=$true)]
                  $TenantName,
              
                  [Parameter(Mandatory=$true)]
                  $clientId,
              
                  [Parameter(Mandatory=$true)]
                  $redirecturi,

                  [Parameter(Mandatory=$true)]
                  $resourceAppIdURI
           )

            #Import the MSOnline module so we can lookup the directory for Microsoft.IdentityModel.Clients.ActiveDirectory.dll and Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll
            #MSOnline module documentation: https://www.powershellgallery.com/packages/MSOnline/1.1.166.0
            Try
                {
                    Write-Debug "Importing MSONline Module for ADAL assemblies"
                    Import-Module MSOnline -ErrorAction Stop
                }
            Catch [System.IO.FileNotFoundException]
                {
                    Write-Warning "The module MSOnline is not installed.`nPlease run Install-Module MSOnline from an elevated window to install it from the PowerShell Gallery"
                    Throw "MSOnline module not installed"
                }
            #Get the module folder so we can load the DLLs we want
            $modulebase = (Get-Module MSONline | Sort Version -Descending | Select -First 1).ModuleBase
            $adal = "{0}\Microsoft.IdentityModel.Clients.ActiveDirectory.dll" -f $modulebase
            $adalforms = "{0}\Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll" -f $modulebase

            #Attempt to load the assemblies. Without these we cannot continue so we need the user to stop and take an action
            Try
                {
                    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
                    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
                }
            Catch
                {
                    #MSOnline Version 1.0 does not contain the DLLs that we need, a minimum version of 1.1.166.0 is required
                    Write-Warning "Unable to load ADAL assemblies.`nUpdate the MSOnline module by running Install-Module MSOnline -Force -AllowClobber"
                    Throw $error[0]
                }
       
           #Build the logon URL with the tenant name
           $authority = "https://login.windows.net/$TenantName"
           Write-Verbose "Logon Authority: $authority"
       
           #Build the auth context and get the result
           Write-Verbose "Creating AuthContext"
           $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
           Write-Verbose "Creating AD UserCredential Object"
           $AdUserCred = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential" -ArgumentList $Credential.username, $Credential.Password
            Try
                {
                    Write-Verbose "Attempting passive authentication"
                    $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$AdUserCred)
                }
            Catch [System.Management.Automation.MethodInvocationException]
                {
                    #The first that the the user runs this, they must open an interactive window to grant permissions to the app
                    If ($error[0].Exception.Message -like "*Send an interactive authorization request for this user and resource*")
                        {
                            Write-Warning "The app has not been granted permissions by the user. Opening an interactive prompt to grant permissions"
                            $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$redirectUri, "Always") #Always prompt for user credentials so we don't use Windows Integrated Auth
                        }
                    Else
                        {
                            Throw
                        }
                }
           
       
           #Return the authentication token
           return $authResult
    }

    #Getting the authorization token
    $token = Get-AuthToken -TenantName $TenantName -clientId $ClientID -redirecturi $RedirectURI -resourceAppIdURI "https://graph.microsoft.com"
 
    #Build REST API header with authorization token
    $authHeader = @{
       'Content-Type'='application\json'
       'Authorization'=$token.CreateAuthorizationHeader()
    }

    #Build Parameter String

    #If period is specified then add that to the parameters unless it is not supported
    if($period -and $Report -notlike "*Office365Activation*")
        {
            $str = "period='{0}'," -f $Period
            $parameterset += $str
        }
    
    #If the date is specified then add that to the parameters unless it is not supported
    if($date -and !($report -eq "MailboxUsage" -or $report -notlike "*Office365Activation*" -or $report -notlike "*getSkypeForBusinessOrganizerActivity*"))
        {
            $str = "date='{0}'" -f $Date
            $parameterset += $str
        }
    #Trim a trailing comma off the ParameterSet if needed
    if($parameterset)
        {
            $parameterset = $parameterset.TrimEnd(",")
        }
    Write-Verbose "Parameter set is: $parameterset"

    #Build the request URL and invoke
    $uri = "https://graph.microsoft.com/beta/reports/{0}({1})/content" -f $report, $parameterset
    Write-Host $uri
    Write-Host "Retrieving Report $report, please wait" -ForegroundColor Green
    $result = Invoke-RestMethod -Uri $uri –Headers $authHeader –Method Get
               
    
    #Convert the stream result to an array
    $resultarray = ConvertFrom-Csv -InputObject $result

}

End{
    Return $resultarray
   }

 