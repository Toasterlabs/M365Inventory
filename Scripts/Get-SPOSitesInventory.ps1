$Creds = Get-Credential 
$Creds.UserName 
 
Connect-SPOService -Url https://tenant-Admin.sharepoint.com -Credential $Creds 
Connect-MsolService -Credential $Creds 
$Users = Get-MsolUser -All 
$UnLicensedUsers = Get-MsolUser -UnlicensedUsersOnly 
$Users.Count 
$UnLicensedUsers.Count 
$Users.Count - $UnLicensedUsers.Count 
 
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 
 
$SitesIncludingPersonal = Get-SPOSite -IncludePersonalSite $true -Limit All -Detailed 
$SitesIncludingPersonal | Select * | Export-Csv -Path C:\temp\PondSites1.csv 
$Sites = Get-SPOSite -Limit All -Detailed 
foreach($asite in $Sites) 
{ 
  Set-SPOUser -Site $asite.Url -LoginName $Creds.UserName -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue 
} 
$Sites | Export-Csv -Path C:\temp\PondSites.csv -NoClobber -NoTypeInformation 
foreach($asite in $Sites) 
{ 
  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($asite.Url) 
  #Authenticate 
  $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName , $Creds.Password) 
  $ctx.Credentials = $credentials 
 
  #Fetch the users in Site Collection 
  $ctx.Load($ctx.Web.Webs) 
  $ctx.ExecuteQuery() 
  $Count=0 
  foreach($aWEb in $ctx.Web.Webs) 
  { 
    $Count++; 
  } 
  Write-Host $asite.Url $Count 
} 
 
$Databases = $null 
$Databases = @(); 
foreach($aSite in $Sites) 
{ 
  $Users = Get-SPOUser -Site $aSite.Url -Limit All | Select * -Verbose 
  foreach($User in $Users) 
  { 
    $DB = New-Object PSObject 
    Add-Member -input $DB noteproperty 'SiteUrl' $aSite.Url  
    Add-Member -input $DB noteproperty 'DisplayName' $User.DisplayName 
    Add-Member -input $DB noteproperty 'LoginName' $User.LoginName 
    Add-Member -input $DB noteproperty 'IsSiteAdmin' $User.IsSiteAdmin 
    Add-Member -input $DB noteproperty 'IsGroup' $User.IsGroup 
    $Databases += $DB 
  } 
} 
$UsersOutput = "C:\temp\AllSitesUsers.csv" 
$Databases | Export-Csv -Path $UsersOutput -NoTypeInformation -Force 
                                                                       
 
$Databases = $null 
$Databases = @(); 
foreach($asite in $Sites) 
{ 
  $Groups = Get-SPOSiteGroup -Site $asite.Url -Limit 100 | Select * 
  foreach($Group in $Groups) 
  { 
    $DB = New-Object PSObject 
    Add-Member -input $DB noteproperty 'SiteUrl' $aSite.Url  
    Add-Member -input $DB noteproperty 'DisplayName' $Group.Title 
    Add-Member -input $DB noteproperty 'LoginName' $Group.LoginName 
    Add-Member -input $DB noteproperty 'OwnerLoginName' $Group.OwnerLoginName 
    Add-Member -input $DB noteproperty 'OwnerTitle' $Group.OwnerTitle 
    $RolesString = "" 
    foreach($Role in $Group.Roles) 
    { 
      $RolesString+=$Role 
      $RolesString+="," 
    } 
    Add-Member -input $DB noteproperty 'Roles' $RolesString 
    $Databases += $DB 
  } 
} 
 
$UsersOutput = "C:\temp\AllSiteGroups.csv" 
$Databases | Export-Csv -Path $UsersOutput -NoTypeInformation -Force 
 
$Databases = $null 
$Databases = @(); 
foreach($asite in $Sites) 
{ 
  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($asite.Url) 
  #Authenticate 
  $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName , $Creds.Password) 
  $ctx.Credentials = $credentials 
 
  #Fetch the users in Site Collection 
  $ctx.Load($ctx.Web.Webs) 
  $Lists = $ctx.Web.Lists 
  $ctx.Load($Lists) 
  $ctx.ExecuteQuery() 
  foreach($List in $Lists) 
  { 
    if($List.Hidden -eq $false) 
    { 
      if($List.ItemCount -gt 100) 
      { 
        $DB = New-Object PSObject 
        Add-Member -input $DB noteproperty 'SiteUrl' $asite.Url  
        Add-Member -input $DB noteproperty 'Title' $List.Title 
        Add-Member -input $DB noteproperty 'ListType' $List.BaseType 
        Add-Member -input $DB noteproperty 'ItemCount' $List.ItemCount 
        $Databases += $DB 
        Write-Host $aSite.Url $List.Title $List.ItemCount 
      } 
       
    } 
  } 
 
  foreach($aWeb in $ctx.Web.Webs) 
  { 
    $Lists = $aWeb.Lists 
    $ctx.Load($Lists) 
    $ctx.ExecuteQuery() 
    foreach($List in $Lists) 
    { 
      if($List.Hidden -eq $false) 
      { 
        if($List.ItemCount -gt 100) 
        { 
          $DB = New-Object PSObject 
          Add-Member -input $DB noteproperty 'SiteUrl' $aWeb.Url  
          Add-Member -input $DB noteproperty 'Title' $List.Title 
          Add-Member -input $DB noteproperty 'ListType' $List.BaseType 
          Add-Member -input $DB noteproperty 'ItemCount' $List.ItemCount 
          $Databases += $DB 
          Write-Host $aSite.Url $List.Title $List.ItemCount 
        } 
       
      } 
    } 
  } 
  Write-Host $asite.Url 
} 
 
$UsersOutput = "C:\temp\Libraries.csv" 
$Databases | Export-Csv -Path $UsersOutput -NoTypeInformation -Force 
$Databases | Out-GridView 
 
 
$Databases = $null 
$Databases = @(); 
foreach($asite in $Sites) 
{ 
  #$asite = "https://leapthepond.sharepoint.com" 
  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($asite.url) 
  #Authenticate 
  $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName , $Creds.Password) 
  $ctx.Credentials = $credentials 
 
  #Fetch the users in Site Collection 
  $ctx.Load($ctx.Web.Webs) 
  $Lists = $ctx.Web.Lists 
  $ctx.Load($Lists) 
  $ctx.ExecuteQuery() 
  foreach($List in $Lists) 
  { 
    if($List.Hidden -eq $false) 
    { 
      $ctx.Load($List) 
      $ctx.ExecuteQuery() 
      if($List.WorkflowAssociations.Count -gt 0) 
      { 
        $DB = New-Object PSObject 
        Add-Member -input $DB noteproperty 'SiteUrl' $asite.Url  
        Add-Member -input $DB noteproperty 'Title' $List.Title 
        Add-Member -input $DB noteproperty 'ListType' $List.BaseType 
        Add-Member -input $DB noteproperty 'WorkflowsCount' $List.WorkflowAssociations.Count 
        $Databases += $DB 
        Write-Host $List.Title $List.ItemCount 
      } 
       
    } 
  } 
 
  foreach($aWeb in $ctx.Web.Webs) 
  { 
    $Lists = $aWeb.Lists 
    $ctx.Load($Lists) 
    $ctx.ExecuteQuery() 
    foreach($List in $Lists) 
    { 
      if($List.Hidden -eq $false) 
      { 
        if($List.ItemCount -gt 100) 
        { 
          $DB = New-Object PSObject 
          Add-Member -input $DB noteproperty 'SiteUrl' $aWeb.Url  
          Add-Member -input $DB noteproperty 'Title' $List.Title 
          Add-Member -input $DB noteproperty 'ListType' $List.BaseType 
          Add-Member -input $DB noteproperty 'WorkflowsCount' $List.WorkflowAssociations.Count 
          $Databases += $DB 
          Write-Host $List.Title $List.ItemCount 
        } 
       
      } 
    } 
  } 
  Write-Host $asite.Url 
} 
 
$UsersOutput = "C:\temp\Libraries.csv" 
$Databases | Export-Csv -Path $UsersOutput -NoTypeInformation -Force 
$Databases | Out-GridView