# Set console output encoding to UTF-8 for Latin characters
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the output directory for working files, report and other variabled
$outputDirectory = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Read tenant IDs from a text file
$tenantIdsFilePath = ".\tenantIds.txt"
if (-not (Test-Path $tenantIdsFilePath)) {
    Write-Host "Tenant IDs file not found. Please create a file named 'tenantIds.txt' with the tenant IDs, one per line."
    pause
    exit
}
$tenantIds = Get-Content -Path $tenantIdsFilePath

# Prompt the user to select a tenant
Write-Host "Select a Tenant ID to connect:"
for ($i = 0; $i -lt $tenantIds.Count; $i++) {
    Write-Host "$($i + 1). $($tenantIds[$i])"
}
$tenantSelection = Read-Host "Enter the number of the Tenant ID you want to connect to"
$selectedTenantId = $tenantIds[$tenantSelection - 1]

# Check if AzureAD module is installed, if not, ask the user to install it
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "AzureAD module is not installed. Please install it using the following command: Install-Module -Name AzureAD -Force -AllowClobber"
    pause
    exit
}

# Check if ImportExcel module is installed, if not, ask the user to install it
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module is not installed. Please install it using the following command: Install-Module -Name ImportExcel -Force -AllowClobber"
    pause
    exit
}

# Check if Microsoft Teams module is installed, if not, ask the user to install it
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "Microsoft Teams module is not installed. Please install it using the following command: Install-Module -Name MicrosoftTeams -Force -AllowClobber"
    pause
    exit
}

# Import the required modules
Import-Module AzureAD
Import-Module MicrosoftTeams
Import-Module ImportExcel

# Connect to Microsoft 365
Write-Host
Write-Host "Connecting to Microsoft 365..."
Connect-AzureAD -TenantId $selectedTenantId
Write-Host "Connected to Microsoft 365."

# Connect to Microsoft Teams
Write-Host
Write-Host "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams
Write-Host "Connected to Microsoft Teams."

# Create the output directory if it doesn't exist
if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory
}

# Export Users with account status and licenses
Write-Host
Write-Host "Exporting users..."
$users = Get-AzureADUser -All $true

$usersArray = $users | ForEach-Object {
    $licenses = ($_ | Get-AzureADUserLicenseDetail).SkuPartNumber -join ", "
    [PSCustomObject]@{
        ObjectId        = $_.ObjectId
        DisplayName     = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        Mail            = $_.Mail
        AccountEnabled  = if ($_.AccountEnabled) { "Enabled" } else { "Blocked" }
        Licenses        = $licenses
    }
}

$usersArray | Export-Csv -Path "$outputDirectory\Users.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Users exported to $outputDirectory\Users.csv."

# Export Groups with all columns available
Write-Host
Write-Host "Exporting groups..."
$groups = Get-AzureADGroup -All $true

# Export Microsoft Teams
Write-Host
Write-Host "Exporting Microsoft Teams..."
$teams = Get-Team

$teams | Export-Csv -Path "$outputDirectory\Teams.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Teams exported to $outputDirectory\Teams.csv."

# Export Teams Channels and Memberships
Write-Host
Write-Host "Exporting Teams channels and memberships..."
$teamsChannels = @()

foreach ($team in $teams) {
    $channels = Get-TeamChannel -GroupId $team.GroupId
    foreach ($channel in $channels) {
        $channelMembers = Get-TeamChannelUser -GroupId $team.GroupId -DisplayName $channel.DisplayName
        foreach ($member in $channelMembers) {
            $teamsChannels += [PSCustomObject]@{
                TeamName    = $team.DisplayName
                TeamId      = $team.GroupId
                ChannelName = $channel.DisplayName
                ChannelId   = $channel.Id
                UserName    = $member.User
                Role        = $member.Role
            }
        }
    }
}

$teamsChannels | Export-Csv -Path "$outputDirectory\TeamsChannels.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Teams channels and memberships exported to $outputDirectory\TeamsChannels.csv."

# Export Domains
Write-Host
Write-Host "Exporting domains..."
$domains = Get-AzureADDomain

$domains | Export-Csv -Path "$outputDirectory\Domains.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Domains exported to $outputDirectory\Domains.csv."

# Convert Teams and Groups data to arrays for easier comparison
$teamsArray = $teams | ForEach-Object {
    [PSCustomObject]@{
        GroupId = $_.GroupId
    }
}

$groupsArray = $groups | ForEach-Object {
    $aliases = (Get-AzureADGroup -ObjectId $_.ObjectId).ProxyAddresses -join ", "
    [PSCustomObject]@{
        ObjectId        = $_.ObjectId
        DisplayName     = $_.DisplayName
        MailEnabled     = $_.MailEnabled
        Mail            = $_.Mail
        GroupTypes      = $_.GroupTypes -join ", "
        SecurityEnabled = $_.SecurityEnabled
        SharePointGroup = if ($teamsArray.GroupId -contains $_.ObjectId -or $aliases -match "SPO:") { "Yes" } else { "No" }
        Aliases         = $aliases
    }
}

$groupsArray | Export-Csv -Path "$outputDirectory\Groups.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Groups exported to $outputDirectory\Groups.csv."

# Export Group Memberships to separate CSV files based on SharePoint group status
Write-Host
Write-Host "Exporting group memberships..."
$sharePointMembershipsCsvPath = "$outputDirectory\SharePointMemberships.csv"
$dlMembershipsCsvPath = "$outputDirectory\DLMemberships.csv"
$sharePointMemberships = @()
$dlMemberships = @()

foreach ($group in $groupsArray) {
    $owners = Get-AzureADGroupOwner -ObjectId $group.ObjectId | Select-Object ObjectId, DisplayName, UserPrincipalName, Mail
    $members = Get-AzureADGroupMember -ObjectId $group.ObjectId | Select-Object ObjectId, DisplayName, UserPrincipalName, Mail

    foreach ($owner in $owners) {
        $membership = [PSCustomObject]@{
            GroupName        = $group.DisplayName
            GroupObjectId    = $group.ObjectId
            ObjectId         = $owner.ObjectId
            DisplayName      = $owner.DisplayName
            UserPrincipalName= $owner.UserPrincipalName
            Mail             = $owner.Mail
            MembershipStatus = "Owner"
        }
        if ($group.SharePointGroup -eq "Yes") {
            $sharePointMemberships += $membership
        } else {
            $dlMemberships += $membership
        }
    }

    foreach ($member in $members) {
        $membership = [PSCustomObject]@{
            GroupName        = $group.DisplayName
            GroupObjectId    = $group.ObjectId
            ObjectId         = $member.ObjectId
            DisplayName      = $member.DisplayName
            UserPrincipalName= $member.UserPrincipalName
            Mail             = $member.Mail
            MembershipStatus = "Member"
        }
        if ($group.SharePointGroup -eq "Yes") {
            $sharePointMemberships += $membership
        } else {
            $dlMemberships += $membership
        }
    }
}

$sharePointMemberships | Export-Csv -Path $sharePointMembershipsCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "SharePoint memberships exported to $sharePointMembershipsCsvPath."

$dlMemberships | Export-Csv -Path $dlMembershipsCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "DL memberships exported to $dlMembershipsCsvPath."

# Disconnect from Microsoft 365 and Microsoft Teams
Write-Host
Write-Host "Disconnecting from Microsoft 365 and Microsoft Teams..."
Disconnect-AzureAD
Disconnect-MicrosoftTeams
Write-Host "Disconnected from Microsoft 365 and Microsoft Teams."

# Consolidate all CSVs into an Excel file with each CSV in a separate sheet
Write-Host
Write-Host "Consolidating CSVs into an Excel file..."
$excelFilePath = "$outputDirectory\$selectedTenantId-report_$timestamp.xlsx"

$excelSheets = @(
    @{Name="Users"; Path="$outputDirectory\Users.csv"},
    @{Name="Groups"; Path="$outputDirectory\Groups.csv"},
    @{Name="SharePointMemberships"; Path="$sharePointMembershipsCsvPath"},
    @{Name="DLMemberships"; Path="$dlMembershipsCsvPath"},
    @{Name="Teams"; Path="$outputDirectory\Teams.csv"},
    @{Name="Domains"; Path="$outputDirectory\Domains.csv"},
    @{Name="TeamsChannels"; Path="$outputDirectory\TeamsChannels.csv"}
)

foreach ($sheet in $excelSheets) {
    Import-Csv -Path $sheet.Path | Export-Excel -Path $excelFilePath -WorksheetName $sheet.Name -AutoSize -TableName $sheet.Name -TableStyle Light1
}

Write-Host "CSV consolidation complete. Excel file saved to $excelFilePath."

# Delete the CSV files after consolidation
foreach ($sheet in $excelSheets) {
    Remove-Item -Path $sheet.Path -Force
    Write-Host "Deleted $($sheet.Path)"
}

# Clear the prompt
Clear-Host

# Clear the prompt
cls

# Calculate MD5 hash of the Excel file
Write-Host "Calculating MD5 hash of the Excel file..."
$md5Hash = Get-FileHash -Path $excelFilePath -Algorithm MD5 | Select-Object -ExpandProperty Hash
Write-Host "MD5 hash of the Excel file: $md5Hash"

# Archive the Excel file and the script itself in a zip file
Write-Host "Archiving the Excel file and the script itself in a zip file..."
$zipFilePath = "$outputDirectory\$selectedTenantId-report_$timestamp.zip"
$scriptPath = $MyInvocation.MyCommand.Path
Compress-Archive -Path $excelFilePath, $scriptPath -DestinationPath $zipFilePath
Write-Host "Files archived to $zipFilePath"

# Delete the Excel file
Remove-Item -Path $excelFilePath -Force
Write-Host "Deleted $excelFilePath"

# Pause at the end
pause
