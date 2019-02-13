# Requirements
The following items are required for the tool to function

## Modules
1. **Azure RM module**
2. **Sharepoint management shell**: https://www.microsoft.com/en-us/download/details.aspx?id=35588
3. **Azure AD Module**: Install-Module AzureAD
4. **MS Online module**: Install-Module MSOnline
5. **Sharepoint Online module**: Install-Module Microsoft.Online.SharePoint.PowerShell
6. **Microsoft Teams Module**: Install-Module MicrosoftTeams
7. **PowerApps Modules**" Install-Module -Name Microsoft.PowerApps.Administration.PowerShell & Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber
8. **PowerApps Plan 2 license or trial license**: This is required for the powerapps/flow admin actions
9. **Office 365 Global Administrator**: Due to the extent of actions we take, a lot of items that are touched will require global admin permissions.
10. **PowerBi Module**: Install-Module -Name MicrosoftPowerBIMgmt


## Graph API
In order to retrieve Graph reports, a graph api application is needed:

1. Login to Portal.Azure.Com
2. Navigate to "Azure Active Directory" > "App Registrations"
3. Click "New Application Registration"
4. Give your application a friendly name, Select application type "native", and enter a redirect URL and click create
5. Click on the App > Required Permissions
6. Click Add and select the** "Microsoft Graph"** API
7. Grant the App the "Read All Usage Reports" permission
8. Copy the Application ID
9. Copy the Redirect URI*

*Note: The redirect URI does not have to be a valid URI. I typically use "urn:" for this field.

Once the Graph application has been created, authorization to the tenant must be granted. In order to do this, adapt the following URI:

__https://login.microsoftonline.com/common/adminconsent?client_id=<-Application Id->&state=12345&redirect_uri=<-Redirect URL->__

And paste it a browser window. You will be prompted to login, and grant permissions. Be aware that, because the redirect URI is garbage, the webpage will go to a 401 or you can hit accept eternally without anything happening...
