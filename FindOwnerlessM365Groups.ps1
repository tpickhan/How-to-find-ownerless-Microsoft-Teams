# Script will use the M365 Activity Report to find ownerless M365 groups
# if ownerless groups are found, they will be added to a SharePoint Online list
# an App registration is required
#
# more details can be found in my blog
# https://office365.thorpick.de/?p=826
#
# created by Thorsten Pickhan
# Initial script created on 07.06.2022 (06/07/2022)
#
# Version 1.0

# PNP PowerShell is required in Version

# Install-Module -Name PnPOnline
# Import-Module PNPOnline


# Set App registration data
#
###
$AppId = "<Your registered appliacation ID>"
$AppSecret = "<Your app secrect>"
$TenantId = "<Your Tenant ID>"

# Define some parameter like SharePont Online list and period of Report time
# following Periof of Report Time are supported
# [D7,D30,D90,D180]
#
###
$RootUrl = "<Your SharePoint Root URL>" # for example "https://contoso.sharepoint.com/teams/TeamsAutomation
$List = "<List Url>" # for example "lists/OwnerlessTeams" 
$PeriodOfReport = "D7" # Period of Report time [D7,D30,D90,D180]

$csvfilename = ".\report.csv"

# Connect to SharePoint Online
#
###
try {
    $RootConnection = Connect-PnPOnline -Url $RootUrl -Interactive -ReturnConnection -ErrorAction Stop 
    #if you use Azure Automation, use the follwoing cmd to connectt to SharePoint Online
    #$RootConnection = Connect-PnPOnline -Url $RootUrl -ErrorAction Stop -ClientId $AppId -Thumbprint $CertificateThumbprint -Tenant $TenantId -ReturnConnection
}
catch {
    Write-Output "Could not connect to PnPOnline - $($_.Exception.Message)"
    break
}

# Get OAuth token
#
###
$Scope = "https://graph.microsoft.com/.default"
$Url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

# Add System.Web for urlencode
#
###
Add-Type -AssemblyName System.Web

# Create body
#
###
$Body = @{
    client_id = $AppId
	client_secret = $AppSecret
	scope = $Scope
	grant_type = 'client_credentials'
}

# Splat the parameters for Invoke-Restmethod for cleaner code
#
###
$PostSplat = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method = 'POST'
    # Create string by joining bodylist with '&'
    Body = $Body
    Uri = $Url
}

# Request the token
#
###
$Request = Invoke-RestMethod @PostSplat

# Create token header
#
###
$Header = @{
    Authorization = "$($Request.token_type) $($Request.access_token)"
}

# GraphAPI URL to get M365 group activities
#
###
$GraphApiUrl = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='$($PeriodOfReport)')"

# Get Usage Report and store them in temporary variable
#
###
$Report = Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $GraphApiUrl -outfile $csvfilename

# Import data
#
###
$UsageData = Import-Csv $csvfilename
$Counter = 1

# Filter dataset on M365 groups without owner
#
###
$OwnerLessTeams = $UsageData | Where-Object {$_."Owner Principal Name" -eq ""}

# Loop throught dataset and add it to SharePoint Online List
#
###
if ($OwnerLessTeams) {
    try {
        $Count = $OwnerLessTeams.Count
    }
    catch {
        $Count = 1
    }
    ForEach ($OwnerLessTeam in $OwnerLessTeams) {
        Write-Output "Proceed list entry $($Counter) from $($Count)..."

	    $GroupId = $OwnerLessTeam."Group Id"
        $ReportRefreshDate = (Get-Date $OwnerLessTeam."Report Refresh Date").AddDays(0)
        $GroupDisplayName = $OwnerLessTeam."Group Display Name"
        $Title = $GroupDisplayName
        if ($OwnerLessTeam."Last Activity Date") {
            $LastActivityDate = (Get-Date $OwnerLessTeam."Last Activity Date").AddDays(0)
        }
        $GroupType = $OwnerLessTeam."Group Type"
        $MemberCount = $OwnerLessTeam."Member Count"
        $ExternalMemberCount = $OwnerLessTeam."External Member Count"
        $ReportPeriod = $OwnerLessTeam."Report Period"
        $IsDeleted = $OwnerLessTeam.'Is Deleted'

        # no owner found, add this M365 group to the SharePoint Online list
        #
        ###
	    try {
            Write-output "Write entry to SharePoint..."
		    $AddSPListPerm = Add-PnPListItem -List $List -Values @{"Title" = $Title; "ReportRefreshDate" = $ReportRefreshDate ; "GroupId" = $GroupId; "GroupDisplayName" = $GroupDisplayName; "GroupType" = $GroupType; "IsDeleted" = $IsDeleted; "MemberCount" = $MemberCount; "ExternalMemberCount" = $ExternalMemberCount; "LastActivityDate" = $LastActivityDate; } -Connection $RootConnection -ErrorAction Stop
	    }
	    catch {
		    Write-Output $GroupDisplayName
		    Write-Output "Could not add entry to SharePoint List - $($_.Exception.Message)"
		    break;
	    }
        $Counter++
    }
}
else {
    Write-Output "You are good to go!"
    Write-Output "No ownerless M365 groups found in the report from $($ReportRefreshDate)"
}
