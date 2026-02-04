# Set console output encoding to UTF-8 for Latin characters
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the output directory for working files, report and other variables
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
Write-Host
Write-Host "Select a Tenant ID to connect:"
for ($i = 0; $i -lt $tenantIds.Count; $i++) {
    Write-Host "$($i + 1). $($tenantIds[$i])"
}
$tenantSelection = Read-Host "Enter the number of the Tenant ID you want to connect to"
$selectedTenantId = $tenantIds[$tenantSelection - 1]

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

# Check if Microsoft Graph module is installed, if not, ask the user to install it
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft Graph module is not installed. Please install it using the following command: Install-Module -Name Microsoft.Graph -Force -AllowClobber"
    pause
    exit
}

# Check if Exchange Online module is installed, if not, ask the user to install it
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module is not installed. Please install it using the following command: Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber"
    pause
    exit
}

# Import the required modules
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users # Get-MgUser Get-MgUserLicenseDetail
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Teams # Get-MgTeam
Import-Module Microsoft.Graph.Identity.DirectoryManagement # Get-MgDomain
Import-Module MicrosoftTeams # Get-TeamChannel Get-TeamChannelUser
Import-Module ImportExcel # Export-Excel
Import-Module ExchangeOnlineManagement # Get-Mailbox

# Connect to Microsoft Graph
Write-Host
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -TenantId $selectedTenantId -Scopes "User.Read.All Directory.Read.All Group.Read.All Team.ReadBasic.All TeamMember.Read.All ChannelMember.Read.All Domain.Read.All" -NoWelcome
Write-Host
Write-Host "Connected to Microsoft Graph."

# Connect to Microsoft Teams
Write-Host
Write-Host "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams
Write-Host "Connected to Microsoft Teams."

# Connect to Exchange Online
Write-Host
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -ShowBanner:$false
Write-Host
Write-Host "Connected to Exchange Online."

# Create the output directory if it doesn't exist
if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory
}


# Export Groups with all columns available
Write-Host
Write-Host "Exporting groups..."
$groups = Get-MgGroup -All


# Export Users with account status and licenses using Graph (temporary)

# Comprehensive license mapping including Visio SKUs
$licenseMap = @{
    'a403ebcc-fae0-4ca2-8c8c-7a907fd6c235' = 'Microsoft Fabric (Free)'
	'f245ecc8-75af-4f8e-b61f-27d8114de5f3' = 'Microsoft 365 Business Standard'
	'3b555118-da6a-4418-894f-7df1e2096870' = 'Microsoft 365 Business Basic'
	'f30db892-07e9-47e9-837c-80727f46fd3d' = 'Microsoft Power Automate Free'
	'4b9405b0-7788-4568-add1-99614e613b69' = 'Exchange Online (Plan 1)'
	'19ec0d23-8335-4cbd-94ac-6050e30712fa' = 'Exchange Online (Plan 2)'
	'cdd28e44-67e3-425e-be4c-737fab2899d3' = 'Microsoft 365 Apps for business'
	'50509a35-f0bd-4c5e-89ac-22f0e16a00f8' = 'Microsoft Teams Rooms Basic without Audio Conferencing'
	'6af4b3d6-14bb-4a2a-960c-6c902aad34f3' = 'Microsoft Teams Rooms Basic'
	'e0dfc8b9-9531-4ec8-94b4-9fec23b05fc8' = 'Microsoft Teams Exploratory Dept (unlisted)'
	'f8a1db68-be16-40ed-86d5-cb42ce701560' = 'Power BI Pro'
	'cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46' = 'Microsoft 365 Business Premium'
	'c1d032e0-5619-4761-9b5c-75b6831e1711' = 'Power BI Premium Per User'
	'53818b1b-4a27-454b-8896-0dba576410e6' = 'Planner and Project Plan 3'
	'295a8eb0-f78d-45c7-8b5b-1eed5ed02dff' = 'Microsoft Teams Shared Devices'
	'c5928f49-12ba-48f7-ada3-0d743a3601d5' = 'Vision Plan 2'
	'1f2f344a-700d-42c9-9427-5cea1d5d7ba6' = 'Microsoft Stream Trial'
	'52ea0e27-ae73-4983-a08f-13561ebdb823' = 'Teams Premium (for Departments)'
	'639dec6b-bb19-468b-871c-c5c441c4b0cb' = 'Microsoft 365 Copilot'
	'5b631642-bd26-49fe-bd20-1daaa972ef80' = 'Microsoft Power Apps for Developer'
	'beb6439c-caad-48d3-bf46-0c82871e12be' = 'Planner Plan 1'
	'3f9f06f5-3c31-472c-985f-62d9c10ec167' = 'Power Pages vTrial for Makers'
}


# Get users and process license information
$usersGraph = Get-MgUser -All -Property `
    Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,UserType, `
    CreatedDateTime,AssignedLicenses,Department,JobTitle,OfficeLocation, `
    UserPrincipalName,ProxyAddresses


$usersGraphReport = $usersGraph | ForEach-Object {
    $user = $_
    $skuPartNumbers = @()

    try {
        $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id
        foreach ($license in $licenseDetails) {
            $skuPartNumbers += $license.SkuPartNumber
        }
    } catch {
        Write-Warning "Could not get license details for $($user.UserPrincipalName)"
        $skuPartNumbers += "Error"
    }

    [PSCustomObject]@{
        Id              = $user.Id
        DisplayName     = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        Mail            = $user.Mail
        AccountEnabled  = $user.AccountEnabled
        UserType        = $user.UserType
        Licenses        = ($user.AssignedLicenses | ForEach-Object {
            if ($licenseMap.ContainsKey($_.SkuId)) {
                $licenseMap[$_.SkuId]
            } else {
                $_.SkuId
            }
        }) -join ", "
        LicenseSKUs     = ($user.AssignedLicenses.SkuId -join ", ")
        SkuPartNumbers  = $skuPartNumbers -join ", "
        CreatedDateTime = $user.CreatedDateTime
        Department      = $user.Department
        JobTitle        = $user.JobTitle
        OfficeLocation  = $user.OfficeLocation
    }
}

$usersGraphReport | Export-Csv -Path "$outputDirectory\UsersGraph.csv" -NoTypeInformation -Encoding UTF8


# Export Shared Mailboxes
Write-Host
Write-Host "Exporting Shared Mailboxes..."
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Select-Object DisplayName,PrimarySmtpAddress,ExternalDirectoryObjectId

$sharedMailboxes | Export-Csv -Path "$outputDirectory\SharedMailboxes.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Shared Mailboxes exported to $outputDirectory\SharedMailboxes.csv."


# Get all Teams
$teams = Get-MgTeam -All

# Export Teams using original file name
$teams | Select-Object Id, DisplayName, Description, Visibility | `
    Export-Csv -Path "$outputDirectory\Teams.csv" -NoTypeInformation -Encoding UTF8 -Force
Write-Host "Teams exported to $outputDirectory\Teams.csv."

# Export Teams Channels and Memberships
Write-Host
Write-Host "Exporting Teams channels and memberships..."
$teamsChannels = @()

# Loop through Teams
foreach ($team in $teams) {
    # Check if Team Id is null or empty
    if ([string]::IsNullOrEmpty($team.Id)) {
        Write-Host "Skipping team '$($team.DisplayName)' because Id is null or empty."
        continue
    }

    # Get channels for the team
    try {
        $channels = Get-TeamChannel -GroupId $team.Id
    } catch {
        Write-Warning "Could not get channels for team $($team.DisplayName): $_"
        continue
    }

    foreach ($channel in $channels) {
        $channelMembers = @()
        $membershipNote = ""

        if ($channel.MembershipType -eq "shared") {
            $membershipNote = "Shared channel – membership not accessible via Teams module"
            Write-Warning "Skipping members of shared channel '$($channel.DisplayName)' in team '$($team.DisplayName)'"
        } else {
            try {
                $channelMembers = Get-TeamChannelUser -GroupId $team.Id -DisplayName $channel.DisplayName
            } catch {
                $membershipNote = "Error retrieving users: $_"
                Write-Warning "Could not get users for channel $($channel.DisplayName) in team $($team.DisplayName): $_"
            }
        }

        if (-not $channelMembers -or $channelMembers.Count -eq 0) {
            $teamsChannels += [PSCustomObject]@{
                TeamName       = $team.DisplayName
                TeamId         = $team.Id
                ChannelName    = $channel.DisplayName
                ChannelId      = $channel.Id
                MembershipType = $channel.MembershipType
                UserName       = ""
                Role           = ""
                Notes          = $membershipNote
            }
        } else {
            foreach ($member in $channelMembers) {
                $teamsChannels += [PSCustomObject]@{
                    TeamName       = $team.DisplayName
                    TeamId         = $team.Id
                    ChannelName    = $channel.DisplayName
                    ChannelId      = $channel.Id
                    MembershipType = $channel.MembershipType
                    UserName       = $member.User
                    Role           = $member.Role
                    Notes          = ""
                }
            }
        }
    }
}


$teamsChannels | Export-Csv -Path "$outputDirectory\TeamsChannels.csv" -NoTypeInformation -Encoding UTF8 -Force
Write-Host "Teams channels and memberships exported to $outputDirectory\TeamsChannels.csv."

# Wait
Start-Sleep -Seconds 2


# Export Domains
Write-Host
Write-Host "Exporting domains..."
$domains = Get-MgDomain

$domains | Export-Csv -Path "$outputDirectory\Domains.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Domains exported to $outputDirectory\Domains.csv."


# Convert Teams and Groups data to arrays for easier comparison
$teamsArray = $teams | ForEach-Object {
    [PSCustomObject]@{
        GroupId = $_.GroupId
    }
}

$groupsArray = $groups | ForEach-Object {
    $aliases = (Get-MgGroup -GroupId $_.Id).ProxyAddresses -join ", "
    [PSCustomObject]@{
        ObjectId        = $_.Id
        DisplayName     = $_.DisplayName
        MailEnabled     = $_.MailEnabled
        Mail            = $_.Mail
        SecurityEnabled = $_.SecurityEnabled
        SharePointGroup = if ($teamsArray.GroupId -contains $_.Id -or $aliases -match "SPO:") { "Yes" } else { "No" }
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
    $owners = Get-MgGroupOwner -GroupId $group.ObjectId
    $members = Get-MgGroupMember -GroupId $group.ObjectId

    foreach ($owner in $owners) {
        $userInfo = $usersGraph | Where-Object { $_.Id -eq $owner.Id }
        $membership = [PSCustomObject]@{
            GroupName        = $group.DisplayName
            GroupObjectId    = $group.ObjectId
            ObjectId         = $owner.Id
            DisplayName      = $userInfo.DisplayName
            UserPrincipalName= $userInfo.UserPrincipalName
            MembershipStatus = "Owner"
        }
        if ($group.SharePointGroup -eq "Yes") {
            $sharePointMemberships += $membership
        } else {
            $dlMemberships += $membership
        }
    }

    foreach ($member in $members) {
        $userInfo = $usersGraph | Where-Object { $_.Id -eq $member.Id }
        $membership = [PSCustomObject]@{
            GroupName        = $group.DisplayName
            GroupObjectId    = $group.ObjectId
            ObjectId         = $member.Id
            DisplayName      = $userInfo.DisplayName
            UserPrincipalName= $userInfo.UserPrincipalName
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






# Get all distribution groups
$groups = Get-DistributionGroup -ResultSize Unlimited

# Prepare an array to store results
$dlOwnersCsvPath = "$outputDirectory\DLOwners.csv"
$dlOwners = @()

foreach ($group in $groups) {
    foreach ($owner in $group.ManagedBy) {
        $ownerRecipient = Get-Recipient -Identity $owner -ErrorAction SilentlyContinue
        # Defensive: Only output if we get a valid recipient
        if ($ownerRecipient) {
            $dlOwners += [PSCustomObject]@{
                GroupDisplayName = $group.DisplayName
                GroupPrimarySmtp = $group.PrimarySmtpAddress
                OwnerDisplayName = $ownerRecipient.DisplayName
                OwnerEmail       = $ownerRecipient.PrimarySmtpAddress
            }
        }
    }
}

# Export to CSV
$dlOwners | Export-Csv -Path $dlOwnersCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "Distribution list owners exported to $outputDirectory\DLOwners.csv."







# Disconnect from Microsoft Teams (if connected)
try {
    Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
    Write-Host "Disconnected from Microsoft Teams." -ForegroundColor Green
} catch {
    Write-Host "No active Teams session to disconnect." -ForegroundColor Yellow
}

# Disconnect from Microsoft Graph (if connected)
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
} catch {
    Write-Host "No active Graph session to disconnect." -ForegroundColor Yellow
}

# Disconnect from Exchange Online (if connected)
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
} catch {
    Write-Host "No active Exchange Online session to disconnect." -ForegroundColor Yellow
}


# Consolidate all CSVs into an Excel file with each CSV in a separate sheet
Write-Host
Write-Host "Consolidating CSVs into an Excel file..."
$excelFilePath = "$outputDirectory\$selectedTenantId-report_$timestamp.xlsx"

$excelSheets = @(
    @{Name="Users"; Path="$outputDirectory\UsersGraph.csv"},
    @{Name="SharedMailboxes"; Path="$outputDirectory\SharedMailboxes.csv"},
    @{Name="Groups"; Path="$outputDirectory\Groups.csv"},
    @{Name="SharePointMemberships"; Path="$sharePointMembershipsCsvPath"},
    @{Name="DLMemberships"; Path="$dlMembershipsCsvPath"},
    @{Name="DLOwners"; Path="$dlOwnersCsvPath"},
    @{Name="Teams"; Path="$outputDirectory\Teams.csv"},
    @{Name="TeamsChannels"; Path="$outputDirectory\TeamsChannels.csv"},
    @{Name="Domains"; Path="$outputDirectory\Domains.csv"}
)

foreach ($sheet in $excelSheets) {
    Import-Csv -Path $sheet.Path | Export-Excel -Path $excelFilePath -WorksheetName $sheet.Name -AutoSize -TableName $sheet.Name -TableStyle Light1
}

Write-Host "CSV consolidation complete. Excel file saved to $excelFilePath."

# Wait
Start-Sleep -Seconds 2


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
Write-Host "File name: $selectedTenantId-report_$timestamp.xlsx"
$md5Hash = Get-FileHash -Path $excelFilePath -Algorithm MD5 | Select-Object -ExpandProperty Hash
Write-Host "MD5 hash of the Excel file: $md5Hash"
Write-Host

# Wait to print info on screen for print screen
Start-Sleep -Seconds 2

# Print screen program
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Send alt + printscreen to capture the active window
[System.Windows.Forms.SendKeys]::SendWait("%{PRTSC}")
Start-Sleep -Milliseconds 500  # Give clipboard time to update

# Try to get image from clipboard
$image = [System.Windows.Forms.Clipboard]::GetImage()
if ($image -ne $null) {
    $bitmap = New-Object System.Drawing.Bitmap $image
    $screenshotPath = "$outputDirectory\screenshot_$timestamp.png"
    $bitmap.Save($screenshotPath)
    Write-Host "Screenshot saved to $screenshotPath"
} else {
    Write-Host "No image found in clipboard. Screenshot not saved."
    $screenshotPath = $null
}

Write-Host
Start-Sleep -Seconds 2

# Archive the Excel file, the script itself, and the screenshot in a zip file
Write-Host "Archiving the Excel file, the script itself, and the screenshot in a zip file..."
$zipFilePath = "$outputDirectory\$selectedTenantId-report_$timestamp.zip"
$scriptPath = $MyInvocation.MyCommand.Path
Write-Host "$zipFilePath"
Write-Host "$scriptPath"
Write-Host "$excelFilePath"
Write-Host

Start-Sleep -Seconds 2

# Prepare list of files to archive
$filesToArchive = @($excelFilePath, $scriptPath)
if ($screenshotPath) {
    $filesToArchive += $screenshotPath
}

Compress-Archive -Path $filesToArchive -DestinationPath $zipFilePath
Write-Host "Files archived to $zipFilePath"

Start-Sleep -Seconds 2

# Delete the Excel file and the screenshot if it exists
Remove-Item -Path $excelFilePath -Force
if ($screenshotPath) {
    Remove-Item -Path $screenshotPath -Force
}
Write-Host "Deleted $selectedTenantId-report_$timestamp.xlsx and screenshot (if created)"

Start-Sleep -Seconds 2

# Pause at the end
pause
