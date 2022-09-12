# Script will create a SharePoint Online list to store ownerless M365 groups
#
# created by Thorsten Pickhan
# Initial script created on 07.06.2022 (06/07/2022)
#
# Version 1.0

# PNP PowerShell is required in Version

# Install-Module -Name PnPOnline
# Import-Module PNPOnline

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

# Define the SharePoint teamsite Url where the list should be created
$RootURL = "https://<Your Tenant name>.sharepoint.com/teams/TeamsAutomation/"

# Define the list Url
$SharePointListName = "Lists/OwnerlessTeams"

# Define the list display name 
$SharePointListDisplayName = "Ownerless Teams"

# Connect to SharePoint Online
$RootConnection = Connect-PnPOnline -Url $RootUrl -Interactive -ReturnConnection -ErrorAction Stop

# Create a new generic List
$item = New-PnPList -Title $SharePointListDisplayName -Template GenericList -Url $SharePointListName -EnableVersioning -OnQuickLaunch

# Create required columns
$item = Add-PnPField -List $SharePointListName -DisplayName "Report Refresh Date" -InternalName "ReportRefreshDate" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Group Id" -InternalName "GroupId" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Group Display Name" -InternalName "GroupDisplayName" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Group Type" -InternalName "GroupType" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Is Deleted" -InternalName "IsDeleted" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Member Count" -InternalName "MemberCount" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "External Member Count" -InternalName "ExternalMemberCount" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Last Activity Date" -InternalName "LastActivityDate" -Type DateTime -AddToDefaultView -Connection $RootConnection
