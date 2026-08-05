#Requires -Version 5.1
# Recommended: Windows PowerShell 5.1 x64
# If needed for WinForms, launch with: powershell.exe -STA

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()


Connect-ExchangeOnline

# ============================================================
# GLOBAL CONFIG
# ============================================================
$TenantDomain          = "#Domain#"
$DefaultUsageLocation = "#Country"
$DefaultCompany       = "#Companyname#"
$script:ServicesConnected = $false
$MaximumFunctionCount = 32768
$TenantId = $TenantDomain
$script:AdminUPN = $null

# ============================================================
# TRANSFER USER VARIABLES
# ============================================================

$script:TransferUser                = $null
$script:TransferCurrentAccess       = $null
$script:TransferRoleAccess          = $null
$script:TransferComparison          = $null
$script:TransferChangesReviewed     = $false

# ============================================================
# TRANSFER EXPORT SETTINGS
# ============================================================

$TransferExportPath = "C:\Temp"

if (-not (Test-Path $TransferExportPath)) {
    New-Item `
        -Path $TransferExportPath `
        -ItemType Directory `
        -Force | Out-Null
}


# ============================================================
# UTILITY FUNCTIONS
# ============================================================
function Is-Blank {
    param([AllowNull()][string]$Value)
    return [string]::IsNullOrWhiteSpace($Value)
}

function Convert-CellToArray {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return @() }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }

    return @(
        $text -split ";" |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Convert-SharedMailboxes {
    param([AllowNull()]$Value)

    $mailboxes = Convert-CellToArray -Value $Value
    if (-not $mailboxes -or $mailboxes.Count -eq 0) { return @() }

    $result = @()
    foreach ($mb in $mailboxes) {
        $result += @{
            Mailbox      = $mb
            AccessRights = @("FullAccess")
            SendAs       = $false
        }
    }

    return $result
}

function Ensure-RoleModules {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "Missing module: PnP.PowerShell. Install with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "Missing module: ImportExcel. Install with: Install-Module ImportExcel -Scope CurrentUser"
    }

    Import-Module PnP.PowerShell -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop
}

function Get-SharedRoleWorkbook {
    param(
        [string]$SiteUrl,
        [string]$ServerRelativeUrl,
        [string]$DownloadPath
    )

    Ensure-RoleModules
    Connect-PnPOnline -Url $SiteUrl -Interactive

    $folder = Split-Path $DownloadPath -Parent
    if (-not (Test-Path $folder)) {
        $null = New-Item -Path $folder -ItemType Directory -Force
    }

    Get-PnPFile -Url $ServerRelativeUrl -Path $folder -FileName (Split-Path $DownloadPath -Leaf) -AsFile -Force

    if (-not (Test-Path $DownloadPath)) {
        throw "Role workbook download failed: $DownloadPath"
    }

    return $DownloadPath
}

function Load-RoleMapFromCsv {
    param([Parameter(Mandatory = $true)][string]$CsvPath)

    if (-not (Test-Path $CsvPath)) {
        throw "Roles file not found: $CsvPath"
    }

    $rows = Import-Csv -Path $CsvPath
    $map = @{}

    foreach ($row in $rows) {
        if ([string]::IsNullOrWhiteSpace($row.Role)) { continue }

        $roleName = ([string]$row.Role).Trim()
        if ([string]::IsNullOrWhiteSpace($roleName)) { continue }

        $map[$roleName] = @{
            Description        = ""
            EntraGroups        = Convert-CellToArray -Value $row.Entragroup
            M365Groups         = Convert-CellToArray -Value $row.M365
            DistributionGroups = Convert-CellToArray -Value $row.Dlist
            SharedMailboxes    = Convert-SharedMailboxes -Value $row.Shared
            CalendarAccess     = Convert-CellToArray -Value $row.Calendar
            AzureRbac          = @()
            Licenses           = @()
            UnifiedGroups      = @()
        }
    }

    return $map
}

function Load-RoleMapFromWorkbook {
    param([string]$WorkbookPath)

    $rows = Import-Excel -Path $WorkbookPath
    $map = @{}

    foreach ($row in $rows) {
        if ([string]::IsNullOrWhiteSpace($row.Role)) { continue }

        $roleName = ([string]$row.Role).Trim()
        if ([string]::IsNullOrWhiteSpace($roleName)) { continue }

        $map[$roleName] = @{
            Description        = ""
            EntraGroups        = Convert-CellToArray -Value $row.Entragroup
            M365Groups         = Convert-CellToArray -Value $row.M365
            DistributionGroups = Convert-CellToArray -Value $row.Dlist
            SharedMailboxes    = Convert-SharedMailboxes -Value $row.Shared
            CalendarAccess     = Convert-CellToArray -Value $row.Calendar
            AzureRbac          = @()
            Licenses           = @()
            UnifiedGroups      = @()
        }
    }

    return $map
}

function Ensure-Module {
    param([Parameter(Mandatory = $true)][string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        throw "Required module '$ModuleName' is not installed. Install it with: Install-Module $ModuleName -Scope CurrentUser"
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.TextBox]$Box,
        [AllowNull()][string]$Message
    )

    if (-not $Box) { return }
    if ([string]::IsNullOrWhiteSpace($Message)) { $Message = "(no message)" }

    $time = Get-Date -Format "HH:mm:ss"
    $line = "$time - $Message`r`n"
    $Box.AppendText($line)
    $Box.SelectionStart = $Box.Text.Length
    $Box.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Parse-SemicolonList {
    param([string]$Text)
    if (Is-Blank $Text) { return @() }
    return @(
        $Text.Split(";") |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Get-CombinedUniqueList {
    param([array]$BaseList, [array]$ExtraList)
    $combined = @()
    if ($BaseList)  { $combined += $BaseList }
    if ($ExtraList) { $combined += $ExtraList }
    return @(
        $combined |
            Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Select-Object -Unique
    )
}

function Get-MailAlias {
    param([string]$FirstName, [string]$LastName)
    if ((Is-Blank $FirstName) -or (Is-Blank $LastName)) { return "" }
    $firstInitial = $FirstName.Trim().Substring(0,1).ToLower()
    $surname      = $LastName.Trim().ToLower()
    $firstInitial = $firstInitial -replace "[^a-z0-9]", ""
    $surname      = $surname -replace "[^a-z0-9]", ""
    return "$firstInitial.$surname"
}

function Get-UserPrincipalName {
    param([string]$FirstName, [string]$LastName)
    $alias = Get-MailAlias -FirstName $FirstName -LastName $LastName
    if (Is-Blank $alias) { return "" }
    return "$alias@$TenantDomain"
}

function Test-UpnExists {
    param([string]$UPN)
    try {
        $null = Get-MgUser -UserId $UPN -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Get-NextAvailableUpn {
    param([string]$FirstName, [string]$LastName)
    $alias = Get-MailAlias -FirstName $FirstName -LastName $LastName
    if (Is-Blank $alias) { throw "Cannot generate alias from first name / last name." }
    $candidate = "$alias@$TenantDomain"
    if (-not (Test-UpnExists -UPN $candidate)) { return $candidate }
    for ($i = 2; $i -le 99; $i++) {
        $candidate = "$alias$i@$TenantDomain"
        if (-not (Test-UpnExists -UPN $candidate)) { return $candidate }
    }
    throw "Could not find an available UPN for alias '$alias'."
}

function Resolve-GroupId {
    param([string]$GroupIdentifier)
    if ([string]::IsNullOrWhiteSpace($GroupIdentifier)) { throw "Blank group name/id supplied." }
    if ($GroupIdentifier -match '^[0-9a-fA-F-]{36}$') { return $GroupIdentifier }
    $escapedName = $GroupIdentifier.Replace("'","''")
    $groups = Get-MgGroup -Filter "displayName eq '$escapedName'" -All -ErrorAction SilentlyContinue
    if (-not $groups) { throw "Could not resolve group '$GroupIdentifier'" }
    return ($groups | Select-Object -First 1).Id
}

function Resolve-LicenseSkuIds {
    param([string[]]$SkuPartNumbers)
    if (-not $SkuPartNumbers -or $SkuPartNumbers.Count -eq 0) { return @() }
    $allSkus = Get-MgSubscribedSku -All
    $skuIds  = @()
    foreach ($sku in $SkuPartNumbers) {
        $match = $allSkus | Where-Object { $_.SkuPartNumber -eq $sku } | Select-Object -First 1
        if (-not $match) { throw "Could not find subscribed SKU '$sku'" }
        $skuIds += $match.SkuId
    }
    return $skuIds
}

function Test-RequiredConnections {
    $graphOk = $false
    $exoOk   = $false
    try {
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($graphContext) { $graphOk = $true }
    } catch {}
    try {
        $exoConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($exoConnection) { $exoOk = $true }
    } catch {}
    return ($graphOk -and $exoOk)
}

function Show-AdminSignInPrompt {
    # No longer used. Sign-in is now handled directly by each Microsoft module:
    # Connect-ExchangeOnline and Connect-MgGraph.
    return $null
}

function Connect-Services {
    param(
        [bool]$UseDeviceCode = $false
    )

    # IMPORTANT:
    # No custom credential prompt is used here.
    # Each Microsoft module opens its own Microsoft sign-in / MFA prompt when needed.
    # Azure/Az has been removed because this tool only needs Exchange Online and Microsoft Graph.

    Ensure-Module -ModuleName "ExchangeOnlineManagement"
    Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop

    $exoConnected = $null
    try {
        $exoConnected = Get-ConnectionInformation -ErrorAction SilentlyContinue
    }
    catch {}

    if (-not $exoConnected) {
        Write-Log $txtLogHome "Connecting to Exchange Online..."
        Write-Log $txtLogOnboard "Connecting to Exchange Online..."

        Connect-ExchangeOnline

        Write-Log $txtLogHome "Exchange Online connected."
        Write-Log $txtLogOnboard "Exchange Online connected."
    }

    Ensure-Module -ModuleName "Microsoft.Graph.Authentication"
    Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Force -ErrorAction Stop

    $graphContext = $null
    try {
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    }
    catch {}

    if (-not $graphContext) {
        Write-Log $txtLogHome "Connecting to Microsoft Graph..."
        Write-Log $txtLogOnboard "Connecting to Microsoft Graph..."

        $graphScopes = @(
            "User.ReadWrite.All",
            "Group.ReadWrite.All"   )

        if ($UseDeviceCode) {
            Connect-MgGraph `
                -Scopes $graphScopes `
                -UseDeviceCode `
                -NoWelcome `
                -ContextScope Process `
                -ErrorAction Stop
        }
        else {
            Connect-MgGraph `
                -Scopes $graphScopes `
                -NoWelcome `
                -ContextScope Process `
                -ErrorAction Stop
        }



        Write-Log $txtLogHome "Microsoft Graph connected."
        Write-Log $txtLogOnboard "Microsoft Graph connected."
    }

    # Load Graph workload modules after authentication is established.
    Ensure-Module -ModuleName "Microsoft.Graph.Users"
    Ensure-Module -ModuleName "Microsoft.Graph.Groups"

    Import-Module Microsoft.Graph.Users -Force


    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.Account) {
            $script:AdminUPN = $ctx.Account
        }
    }
    catch {}
}


function Add-UserToEntraGroups {
    param([string]$UserId, [string[]]$Groups, [System.Windows.Forms.TextBox]$LogBox)
    foreach ($groupItem in $Groups) {
        try {
            if ([string]::IsNullOrWhiteSpace($groupItem)) { continue }
            $groupId = Resolve-GroupId -GroupIdentifier $groupItem
            New-MgGroupMemberByRef -GroupId $groupId -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" } -ErrorAction Stop
            Write-Log $LogBox "Added user to Entra/M365 group: $groupItem"
        } catch {
            Write-Log $LogBox ("WARN - Failed to add group '{0}' : {1}" -f $groupItem, $_.Exception.Message)
        }
    }
}

function Add-UserToDistributionGroups {
    param([string]$PrimarySmtpAddress, [string[]]$DistributionGroups, [System.Windows.Forms.TextBox]$LogBox)
    foreach ($dl in $DistributionGroups) {
        try {
            if ([string]::IsNullOrWhiteSpace($dl)) { continue }
            Add-DistributionGroupMember -Identity $dl -Member $PrimarySmtpAddress -ErrorAction Stop
            Write-Log $LogBox "Added user to distribution group: $dl"
        } catch {
            Write-Log $LogBox ("WARN - Failed to add distribution group '{0}' : {1}" -f $dl, $_.Exception.Message)
        }
    }
}

function Add-UserToUnifiedGroups {
    param([string]$PrimarySmtpAddress, [string[]]$UnifiedGroups, [System.Windows.Forms.TextBox]$LogBox)
    foreach ($grp in $UnifiedGroups) {
        try {
            if ([string]::IsNullOrWhiteSpace($grp)) { continue }
            Add-UnifiedGroupLinks -Identity $grp -LinkType Members -Links $PrimarySmtpAddress -ErrorAction Stop
            Write-Log $LogBox "Added user to unified group: $grp"
        } catch {
            Write-Log $LogBox ("WARN - Failed to add unified group '{0}' : {1}" -f $grp, $_.Exception.Message)
        }
    }
}

function Add-SharedMailboxAccess {
    param([string]$PrimarySmtpAddress, [array]$SharedMailboxes, [System.Windows.Forms.TextBox]$LogBox)
    foreach ($mbx in $SharedMailboxes) {
        try {
            if (-not $mbx -or [string]::IsNullOrWhiteSpace($mbx.Mailbox)) { continue }
            $mailbox = $mbx.Mailbox
            $rights  = $mbx.AccessRights
            $sendAs  = $mbx.SendAs
            if ($rights -and $rights.Count -gt 0) {
                foreach ($right in $rights) {
                    Add-MailboxPermission -Identity $mailbox -User $PrimarySmtpAddress -AccessRights $right -InheritanceType All -AutoMapping:$true -ErrorAction Stop | Out-Null
                }
                Write-Log $LogBox ("Granted mailbox rights [{0}] to {1}" -f ($rights -join ', '), $mailbox)
            }
            if ($sendAs) {
                Add-RecipientPermission -Identity $mailbox -Trustee $PrimarySmtpAddress -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Log $LogBox "Granted SendAs to mailbox: $mailbox"
            }
        } catch {
            Write-Log $LogBox ("WARN - Failed mailbox permission for '{0}' : {1}" -f $mbx.Mailbox, $_.Exception.Message)
        }
    }
}

function Add-AzureRbacAssignments {
    param([string]$UserPrincipalName, [array]$AzureRbac, [System.Windows.Forms.TextBox]$LogBox)
    foreach ($rbac in $AzureRbac) {
        try {
            New-AzRoleAssignment -SignInName $UserPrincipalName -RoleDefinitionName $rbac.Role -Scope $rbac.Scope -ErrorAction Stop | Out-Null
            Write-Log $LogBox ("Assigned Azure RBAC role '{0}' at scope '{1}'" -f $rbac.Role, $rbac.Scope)
        } catch {
            Write-Log $LogBox ("WARN - Failed Azure RBAC role '{0}' : {1}" -f $rbac.Role, $_.Exception.Message)
        }
    }
}

# ============================================================
# OFFICE / LOCATION MAP

#To add a new location copy another one and change the details.
   #     "New Site name" = @{
   #     StreetAddress = "Street Address"
   #    City          = "City"
   #    State         = "Canterbury?"
   #     Country       = "New Zealand"
   #  }

# ============================================================
$OfficeMap = @{
    "Site 1" = @{
        StreetAddress = "Street Address 1"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
    "Site 2" = @{
        StreetAddress = "Street Address 2"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
    "Site 3" = @{
        StreetAddress = "Street Address 3"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
    "Site 4" = @{
        StreetAddress = "Street Address 4"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
    "Site 5" = @{
        StreetAddress = "Street Address 5"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
}

# ============================================================
# SHARED ROLE SOURCE CONFIG
#You need to map this folder for the script to work. This is the location of your Userroles.csv file
# ============================================================
$RoleCsvPath = Join-Path $env:USERPROFILE "#Sharepointlocation#\userroles.csv"

# ============================================================
# ROLE MAP
# ============================================================
if (-not (Test-Path $RoleCsvPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Roles file not found. Please ensure SharePoint is synced via OneDrive.`r`n`r`nExpected:`r`n$RoleCsvPath",
        "Missing Roles File"
    ) | Out-Null
    return
}

try {
    $RoleMap = Load-RoleMapFromCsv -CsvPath $RoleCsvPath
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to load roles file:`r`n$($_.Exception.Message)", "Role Load Error") | Out-Null
    return
}

function Get-RoleNames { return @($RoleMap.Keys | Sort-Object) }
function Get-OfficeNames { return @($OfficeMap.Keys | Sort-Object) }

function Create-NewStarter {
    param([hashtable]$FormData, [System.Windows.Forms.TextBox]$LogBox)

    if (-not $script:ServicesConnected) { throw "Modules/services are not connected. Please click 'Connect Modules' first." }
    $roleConfig = $RoleMap[$FormData.Role]
    if (-not $roleConfig) { throw "Role '$($FormData.Role)' is not defined in RoleMap." }
    if (Is-Blank $FormData.Password) { throw "Password is required." }

    $hasAssignments =
        ($roleConfig.EntraGroups.Count -gt 0) -or
        ($roleConfig.M365Groups.Count -gt 0) -or
        ($roleConfig.DistributionGroups.Count -gt 0) -or
        ($roleConfig.SharedMailboxes.Count -gt 0) -or
        ($roleConfig.AzureRbac.Count -gt 0) -or
        ($roleConfig.Licenses.Count -gt 0) -or
        ($roleConfig.UnifiedGroups.Count -gt 0)

    Write-Log $LogBox ("Role selected: {0}" -f $FormData.Role)
    if (-not $hasAssignments) {
        Write-Log $LogBox "NOTE: Selected role currently has no permissions configured."
    }

    $upn           = Get-NextAvailableUpn -FirstName $FormData.FirstName -LastName $FormData.LastName
    $mailNickname = ($upn.Split("@")[0])
    $officeDetails = $OfficeMap[$FormData.Office]
    if (-not $officeDetails) { throw "Office '$($FormData.Office)' is not configured." }

    Write-Log $LogBox ("Creating user: {0}" -f $upn)

    $newUserParams = @{
        AccountEnabled    = $true
        DisplayName       = $FormData.DisplayName
        GivenName         = $FormData.FirstName
        Surname           = $FormData.LastName
        UserPrincipalName = $upn
        MailNickname      = $mailNickname
        JobTitle          = $FormData.JobTitle
        Department        = $FormData.Department
        CompanyName       = $FormData.Company
        EmployeeId        = $FormData.EmployeeId
        OfficeLocation    = $FormData.Office
        MobilePhone       = $FormData.Mobile
        BusinessPhones    = @($FormData.BusinessPhone)
        UsageLocation     = $FormData.UsageLocation
        StreetAddress     = $officeDetails.StreetAddress
        City              = $officeDetails.City
        State             = $officeDetails.State
        Country           = $officeDetails.Country
        PasswordProfile   = @{ ForceChangePasswordNextSignIn = $true; Password = $FormData.Password }
    }

    $newUser = New-MgUser @newUserParams -ErrorAction Stop
    Write-Log $LogBox ("User created with Object ID: {0}" -f $newUser.Id)

    if (-not (Is-Blank $FormData.ManagerUpn)) {
        try {
            $mgr = Get-MgUser -UserId $FormData.ManagerUpn -ErrorAction Stop
            Set-MgUserManagerByRef -UserId $newUser.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)" } -ErrorAction Stop
            Write-Log $LogBox ("Manager set to {0}" -f $FormData.ManagerUpn)
        } catch {
            Write-Log $LogBox ("WARN - Failed to set manager: {0}" -f $_.Exception.Message)
        }
    }

    try {
        Update-MgUser -UserId $newUser.Id -MobilePhone $FormData.Mobile -BusinessPhones @($FormData.BusinessPhone) -ErrorAction Stop | Out-Null
    } catch {
        Write-Log $LogBox ("WARN - Phone update failed: {0}" -f $_.Exception.Message)
    }


   try {
        Write-Log $LogBox "Assigning standard licence group..."

        Add-UserToEntraGroups `
            -UserId $newUser.Id `
            -Groups @("LIC-M365-Business-Premium-Defender-Suite") `
            -LogBox $LogBox

        Write-Log $LogBox "Waiting 120 seconds for licence and mailbox provisioning..."

        Start-Sleep -Seconds 120

        Write-Log $LogBox "Continuing with access assignment..."
    }
    catch {
        throw "Failed to assign licence group: $($_.Exception.Message)"
    }



    $allEntraGroups   = Get-CombinedUniqueList -BaseList $roleConfig.EntraGroups        -ExtraList $FormData.ManualEntraGroups
    $allM365Groups    = Get-CombinedUniqueList -BaseList $roleConfig.M365Groups         -ExtraList $FormData.ManualM365Groups
    $allDLs           = Get-CombinedUniqueList -BaseList $roleConfig.DistributionGroups -ExtraList $FormData.ManualDistributionGroups
    $allUnifiedGroups = Get-CombinedUniqueList -BaseList $roleConfig.UnifiedGroups      -ExtraList $FormData.ManualUnifiedGroups

    if ($allEntraGroups.Count -gt 0)   { Add-UserToEntraGroups -UserId $newUser.Id -Groups $allEntraGroups -LogBox $LogBox }
    if ($allM365Groups.Count -gt 0)    { Add-UserToEntraGroups -UserId $newUser.Id -Groups $allM365Groups -LogBox $LogBox }
    if ($allDLs.Count -gt 0)           { Add-UserToDistributionGroups -PrimarySmtpAddress $upn -DistributionGroups $allDLs -LogBox $LogBox }
    if ($allUnifiedGroups.Count -gt 0) { Add-UserToUnifiedGroups -PrimarySmtpAddress $upn -UnifiedGroups $allUnifiedGroups -LogBox $LogBox }

    if ($roleConfig.SharedMailboxes.Count -gt 0) {
        Add-SharedMailboxAccess -PrimarySmtpAddress $upn -SharedMailboxes $roleConfig.SharedMailboxes -LogBox $LogBox
    }

    if ($FormData.ManualSharedMailboxes.Count -gt 0) {
        $manualMailboxObjects = @()
        foreach ($s in $FormData.ManualSharedMailboxes) {
            $manualMailboxObjects += @{ Mailbox = $s; AccessRights = @("FullAccess"); SendAs = $false }
        }
        Add-SharedMailboxAccess -PrimarySmtpAddress $upn -SharedMailboxes $manualMailboxObjects -LogBox $LogBox
    }

    if ($roleConfig.AzureRbac.Count -gt 0) {
        Write-Log $LogBox "NOTE: Azure RBAC values are present in the role data, but Azure/Az connection has been removed from this version. No Azure RBAC assignments were applied."
    }

    if ($roleConfig.CalendarAccess.Count -gt 0) {
        Write-Log $LogBox "NOTE: Calendar access values are currently preview-only. No calendar permissions are applied by this script yet."
    }

    Write-Log $LogBox ("Completed onboarding for {0}" -f $upn)
    return $upn
}


function Get-RoleConfig {
    param([string]$RoleName)

    if ($null -eq $RoleName -or $RoleName.Trim() -eq "") {
        return $null
    }

    if (-not $RoleMap.ContainsKey($RoleName)) {
        return $null
    }

    return $RoleMap[$RoleName]
}

function Compare-TransferLists {
    param(
        [array]$Current,
        [array]$Target
    )

    @{
        Add = @(
            $Target |
                Where-Object { $_ -notin $Current } |
                Sort-Object -Unique
        )

        Remove = @(
            $Current |
                Where-Object { $_ -notin $Target } |
                Sort-Object -Unique
        )
    }
}

function Build-TransferComparison {
    param(
        [hashtable]$Current,
        [hashtable]$Target
    )

    return @{
        Entra   = Compare-TransferLists $Current.EntraGroups        $Target.EntraGroups
        M365    = Compare-TransferLists $Current.M365Groups         $Target.M365Groups
        DLists  = Compare-TransferLists $Current.DistributionGroups $Target.DistributionGroups
        Shared  = Compare-TransferLists $Current.SharedMailboxes    $Target.SharedMailboxes
    }
}

# ============================================================
# SECTION 2 - EXPORT CURRENT ACCESS AND REMOVE FROM CSV
# Place with helper/transfer functions
# ============================================================

function Export-CurrentUserAccess {
    param(
        [Parameter(Mandatory = $true)][object]$User,
        [Parameter(Mandatory = $true)][string]$ExportFile,
        [Parameter(Mandatory = $true)][System.Windows.Forms.TextBox]$LogBox
    )

    $upn = $User.UserPrincipalName
    $exportData = @()

    Write-Log $LogBox "Exporting current access for $upn..."

    # ------------------------------------------------------------
    # Entra security groups via Microsoft Graph
    # This captures non-mail-enabled security groups only.
    # Mail-enabled groups are handled by Exchange Online sections.
    # ------------------------------------------------------------
    try {
        $groups = Get-MgUserMemberOf `
            -UserId $User.Id `
            -All `
            -ErrorAction Stop

        foreach ($group in $groups) {
            try {
                if ($group.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.group') {
                    continue
                }

                $groupObj = Get-MgGroup `
                    -GroupId $group.Id `
                    -ErrorAction Stop

                if ($groupObj.DisplayName -eq "LIC-M365-Business-Premium-Defender-Suite") {
                    $exportData += [pscustomobject]@{
                        Type = "LicenceGroup"
                        Name = $groupObj.DisplayName
                    }
                    continue
                }

                if ($groupObj.MailEnabled) {
                    continue
                }

                $exportData += [pscustomobject]@{
                    Type = "EntraGroup"
                    Name = $groupObj.DisplayName
                }
            }
            catch {
                Write-Log $LogBox ("WARN - Could not inspect Entra group membership: {0}" -f $_.Exception.Message)
            }
        }
    }
    catch {
        Write-Log $LogBox ("WARN - Could not export Entra groups: {0}" -f $_.Exception.Message)
    }

    # ------------------------------------------------------------
    # Microsoft 365 Groups via Exchange Online
    # ------------------------------------------------------------
    try {
        Get-UnifiedGroup -ResultSize Unlimited | ForEach-Object {
            $ug = $_
            try {
                $members = Get-UnifiedGroupLinks `
                    -Identity $ug.Identity `
                    -LinkType Members `
                    -ResultSize Unlimited `
                    -ErrorAction Stop

                $isMember = @($members | Where-Object {
                    ([string]$_.PrimarySmtpAddress).ToLower() -eq $upn.ToLower() -or
                    ([string]$_.WindowsEmailAddress).ToLower() -eq $upn.ToLower() -or
                    ([string]$_.Name).ToLower() -eq ([string]$User.DisplayName).ToLower()
                }).Count -gt 0

                if ($isMember) {
                    $exportData += [pscustomobject]@{
                        Type = "M365Group"
                        Name = $ug.Identity
                    }
                }
            }
            catch {
                # Ignore groups that cannot be queried.
            }
        }
    }
    catch {
        Write-Log $LogBox ("WARN - Could not export M365 group memberships: {0}" -f $_.Exception.Message)
    }

    # ------------------------------------------------------------
    # Distribution Lists via Exchange Online
    # ------------------------------------------------------------
    try {
        Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {
            $dl = $_
            try {
                $members = Get-DistributionGroupMember `
                    -Identity $dl.Identity `
                    -ResultSize Unlimited `
                    -ErrorAction Stop

                $isMember = @($members | Where-Object {
                    ([string]$_.PrimarySmtpAddress).ToLower() -eq $upn.ToLower() -or
                    ([string]$_.WindowsEmailAddress).ToLower() -eq $upn.ToLower() -or
                    ([string]$_.Name).ToLower() -eq ([string]$User.DisplayName).ToLower()
                }).Count -gt 0

                if ($isMember) {
                    $exportData += [pscustomobject]@{
                        Type = "DList"
                        Name = $dl.Identity
                    }
                }
            }
            catch {
                # Ignore DLs that cannot be queried.
            }
        }
    }
    catch {
        Write-Log $LogBox ("WARN - Could not export Distribution List memberships: {0}" -f $_.Exception.Message)
    }

    # ------------------------------------------------------------
    # Shared Mailboxes via role profile CSV catalogue only
    # This avoids scanning/removing against every shared mailbox in tenant.
    # ------------------------------------------------------------
    try {
        $managed = Get-ManagedRoleAccessCatalog

        foreach ($mailbox in @($managed.SharedMailboxes)) {
            if ($null -eq $mailbox -or [string]::IsNullOrWhiteSpace([string]$mailbox.Mailbox)) {
                continue
            }

            $mailboxName = $mailbox.Mailbox

            try {
                $fullAccess = Get-MailboxPermission `
                    -Identity $mailboxName `
                    -User $upn `
                    -ErrorAction SilentlyContinue |
                    Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'FullAccess' }

                if ($fullAccess) {
                    $exportData += [pscustomobject]@{
                        Type = "SharedMailboxFullAccess"
                        Name = $mailboxName
                    }
                }
            }
            catch {}

            try {
                $sendAs = Get-RecipientPermission `
                    -Identity $mailboxName `
                    -Trustee $upn `
                    -ErrorAction SilentlyContinue |
                    Where-Object { $_.AccessRights -contains 'SendAs' }

                if ($sendAs) {
                    $exportData += [pscustomobject]@{
                        Type = "SharedMailboxSendAs"
                        Name = $mailboxName
                    }
                }
            }
            catch {}
        }
    }
    catch {
        Write-Log $LogBox ("WARN - Could not export shared mailbox permissions: {0}" -f $_.Exception.Message)
    }

    $folder = Split-Path $ExportFile -Parent
    if (-not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }

    $exportData |
        Sort-Object Type,Name -Unique |
        Export-Csv -Path $ExportFile -NoTypeInformation

    Write-Log $LogBox "Current access exported to: $ExportFile"
    Write-Log $LogBox ("Exported {0} access entrie(s)." -f @($exportData).Count)
}

function Remove-AccessFromCsv {
    param(
        [Parameter(Mandatory = $true)][object]$User,
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][System.Windows.Forms.TextBox]$LogBox
    )

    if (-not (Test-Path $CsvPath)) {
        throw "Transfer access CSV not found: $CsvPath"
    }

    $upn = $User.UserPrincipalName
    $accessEntries = Import-Csv -Path $CsvPath

    Write-Log $LogBox "Removing access using CSV: $CsvPath"

    foreach ($entry in $accessEntries) {
        switch ($entry.Type) {

            "LicenceGroup" {
                Write-Log $LogBox "Keeping licence group: $($entry.Name)"
            }

            "EntraGroup" {
                if ($entry.Name -eq "LIC-M365-Business-Premium-Defender-Suite") {
                    Write-Log $LogBox "Skipped licence group: $($entry.Name)"
                    continue
                }

                try {
                    $groupId = Resolve-GroupId -GroupIdentifier $entry.Name

                    Remove-MgGroupMemberByRef `
                        -GroupId $groupId `
                        -DirectoryObjectId $User.Id `
                        -ErrorAction Stop

                    Write-Log $LogBox "Removed Entra Group: $($entry.Name)"
                }
                catch {
                    Write-Log $LogBox ("WARN - Failed removing Entra Group '{0}' : {1}" -f $entry.Name, $_.Exception.Message)
                }
            }

            "M365Group" {
                try {
                    Remove-UnifiedGroupLinks `
                        -Identity $entry.Name `
                        -LinkType Members `
                        -Links $upn `
                        -Confirm:$false `
                        -ErrorAction Stop

                    Write-Log $LogBox "Removed M365 Group: $($entry.Name)"
                }
                catch {
                    Write-Log $LogBox ("WARN - Failed removing M365 Group '{0}' : {1}" -f $entry.Name, $_.Exception.Message)
                }
            }

            "DList" {
                try {
                    Remove-DistributionGroupMember `
                        -Identity $entry.Name `
                        -Member $upn `
                        -Confirm:$false `
                        -BypassSecurityGroupManagerCheck `
                        -ErrorAction Stop

                    Write-Log $LogBox "Removed Distribution List: $($entry.Name)"
                }
                catch {
                    Write-Log $LogBox ("WARN - Failed removing Distribution List '{0}' : {1}" -f $entry.Name, $_.Exception.Message)
                }
            }

            "SharedMailboxFullAccess" {
                try {
                    Remove-MailboxPermission `
                        -Identity $entry.Name `
                        -User $upn `
                        -AccessRights FullAccess `
                        -Confirm:$false `
                        -ErrorAction Stop | Out-Null

                    Write-Log $LogBox "Removed FullAccess from Shared Mailbox: $($entry.Name)"
                }
                catch {
                    Write-Log $LogBox ("WARN - Failed removing FullAccess from Shared Mailbox '{0}' : {1}" -f $entry.Name, $_.Exception.Message)
                }
            }

            "SharedMailboxSendAs" {
                try {
                    Remove-RecipientPermission `
                        -Identity $entry.Name `
                        -Trustee $upn `
                        -AccessRights SendAs `
                        -Confirm:$false `
                        -ErrorAction Stop | Out-Null

                    Write-Log $LogBox "Removed SendAs from Shared Mailbox: $($entry.Name)"
                }
                catch {
                    Write-Log $LogBox ("WARN - Failed removing SendAs from Shared Mailbox '{0}' : {1}" -f $entry.Name, $_.Exception.Message)
                }
            }
        }
    }

    Write-Log $LogBox "CSV access cleanup completed."
}




function Get-ManagedRoleAccessCatalog {
    $catalog = @{
        EntraGroups        = @()
        M365Groups         = @()
        DistributionGroups = @()
        SharedMailboxes    = @()
    }

    foreach ($roleName in $RoleMap.Keys) {
        $role = $RoleMap[$roleName]

        if ($role.EntraGroups) {
            $catalog.EntraGroups += $role.EntraGroups
        }

        if ($role.M365Groups) {
            $catalog.M365Groups += $role.M365Groups
        }

        if ($role.DistributionGroups) {
            $catalog.DistributionGroups += $role.DistributionGroups
        }

        if ($role.SharedMailboxes) {
            foreach ($mailbox in $role.SharedMailboxes) {
                if ($mailbox.Mailbox) {
                    $catalog.SharedMailboxes += $mailbox
                }
            }
        }
    }

    return @{
        EntraGroups        = @($catalog.EntraGroups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        M365Groups         = @($catalog.M365Groups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        DistributionGroups = @($catalog.DistributionGroups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        SharedMailboxes    = @($catalog.SharedMailboxes)
    }
}

function Remove-AllUserAccess {
    param(
        [object]$User,
        [System.Windows.Forms.TextBox]$LogBox
    )

    $upn = $User.UserPrincipalName

    Write-Log $LogBox "Removing all existing access..."

    # Remove all Entra / M365 groups except licence group

$groups = Get-MgUserMemberOf `
    -UserId $User.Id `
    -All



foreach ($group in $groups) {

    if ($group.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.group') {
        continue
    }

    $groupObj = Get-MgGroup `
        -GroupId $group.Id `
        -ErrorAction Stop

    if ($groupObj.DisplayName -eq "LIC-M365-Business-Premium-Defender-Suite") {
        continue
    }

    Remove-MgGroupMemberByRef `
        -GroupId $groupObj.Id `
        -DirectoryObjectId $User.Id `
        -ErrorAction Stop

    Write-Log $LogBox "Removed Entra Group: $($groupObj.DisplayName)"
}

    # Remove all Distribution Lists

    try {

        Get-DistributionGroup -ResultSize Unlimited |
        ForEach-Object {

            try {

                Remove-DistributionGroupMember `
                    -Identity $_.Identity `
                    -Member $upn `
                    -Confirm:$false `
                    -BypassSecurityGroupManagerCheck `
                    -ErrorAction Stop

            }
            catch {}

        }

        Write-Log $LogBox "Distribution Lists checked."

    }
    catch {
        Write-Log $LogBox "WARN - DL cleanup failed."
    }

    }

 # --------------------------------------------------
# SHARED MAILBOXES (FROM ROLE CSV ONLY)
# --------------------------------------------------

try {
    $managedRoles = Get-ManagedRoleAccessCatalog

    foreach ($mailbox in $managedRoles.SharedMailboxes) {
        # Check if the mailbox object or its Mailbox property is null/empty
        if ($null -eq $mailbox -or [string]::IsNullOrEmpty($mailbox.Mailbox)) {
            continue
        }

        # Remove FullAccess Permissions
        try {
            Remove-MailboxPermission `
                -Identity $mailbox.Mailbox `
                -User $upn `
                -AccessRights FullAccess `
                -Confirm:$false `
                -ErrorAction SilentlyContinue | Out-Null

            Write-Log $LogBox "Checked FullAccess: $($mailbox.Mailbox)"
        }
        catch {
            Write-Log $LogBox "ERROR - FullAccess removal failed for: $($mailbox.Mailbox)"
        }

        # Remove SendAs Permissions
        try {
            Remove-RecipientPermission `
                -Identity $mailbox.Mailbox `
                -Trustee $upn `
                -AccessRights SendAs `
                -Confirm:$false `
                -ErrorAction SilentlyContinue | Out-Null

            Write-Log $LogBox "Checked SendAs: $($mailbox.Mailbox)"
        }
        catch {
            Write-Log $LogBox "ERROR - SendAs removal failed for: $($mailbox.Mailbox)"
        }
    }

    Write-Log $LogBox "Shared mailbox cleanup completed."
}
catch {

    $msg = "WARN - Shared mailbox cleanup failed: $($_.Exception.Message)"

    if ($null -ne $LogBox) {
        Write-Log $LogBox $msg
    }
    else {
        Write-Host $msg
    }

}

function Apply-TransferRoleAccess {
    param(
        [object]$User,
        [hashtable]$RoleConfig,
        [System.Windows.Forms.TextBox]$LogBox
    )

    if ($null -eq $User) {
        throw "No transfer user loaded."
    }

    if ($null -eq $RoleConfig) {
        throw "No transfer role selected."
    }

    $upn = $User.UserPrincipalName

    Write-Log $LogBox "Transfer User Id: $($User.Id)"
Write-Log $LogBox "Transfer UPN: $upn"

    Write-Log $LogBox "Applying new role access to $upn..."

    if ($RoleConfig.EntraGroups.Count -gt 0) {
        Add-UserToEntraGroups `
            -UserId $User.Id `
            -Groups $RoleConfig.EntraGroups `
            -LogBox $LogBox
    }

    if ($RoleConfig.M365Groups.Count -gt 0) {
        Add-UserToEntraGroups `
            -UserId $User.Id `
            -Groups $RoleConfig.M365Groups `
            -LogBox $LogBox
    }

    if ($RoleConfig.DistributionGroups.Count -gt 0) {
        Add-UserToDistributionGroups `
            -PrimarySmtpAddress $upn `
            -DistributionGroups $RoleConfig.DistributionGroups `
            -LogBox $LogBox
    }

    if ($RoleConfig.SharedMailboxes.Count -gt 0) {
        Add-SharedMailboxAccess `
            -PrimarySmtpAddress $upn `
            -SharedMailboxes $RoleConfig.SharedMailboxes `
            -LogBox $LogBox
    }

    Write-Log $LogBox "Finished applying new role access."
}



# ============================================================
# UI Controls
# ============================================================
$script:font     = New-Object System.Drawing.Font("Segoe UI", 9.5)
$script:fontBold = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$script:fontMono = New-Object System.Drawing.Font("Consolas", 9)

function New-TableLabel {
    param([string]$Text)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.AutoSize = $true
    $lbl.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $lbl.Margin = New-Object System.Windows.Forms.Padding(10,8,10,8)
    $lbl.Font = $script:font
    return $lbl
}

function New-TableTextBox {
    param([bool]$ReadOnly = $false, [bool]$Password = $false)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tb.Margin = New-Object System.Windows.Forms.Padding(10,4,10,4)
    $tb.Font = $script:font
    $tb.ReadOnly = $ReadOnly
    $tb.UseSystemPasswordChar = $Password
    return $tb
}

function New-TableComboBox {
    param([string[]]$Items)
    $cb = New-Object System.Windows.Forms.ComboBox
    $cb.Dock = [System.Windows.Forms.DockStyle]::Fill
    $cb.Margin = New-Object System.Windows.Forms.Padding(10,4,10,4)
    $cb.Font = $script:font
    $cb.DropDownStyle = "DropDownList"
    if ($Items -and $Items.Count -gt 0) {
        [void]$cb.Items.AddRange([string[]]$Items)
        $cb.SelectedIndex = 0
    }
    return $cb
}

function New-MultiTextBox {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tb.Multiline = $true
    $tb.ScrollBars = "Vertical"
    $tb.Margin = New-Object System.Windows.Forms.Padding(10,4,10,8)
    $tb.Font = $script:font
    return $tb
}

function Add-TableRow {
    param([System.Windows.Forms.TableLayoutPanel]$Table, [int]$Row, [string]$LabelText, [System.Windows.Forms.Control]$Control)
    $lbl = New-TableLabel -Text $LabelText
    $Table.Controls.Add($lbl, 0, $Row)
    $Table.Controls.Add($Control, 1, $Row)
}

function Set-PanelVisible {
    param([System.Windows.Forms.Panel]$ShowPanel, [System.Windows.Forms.Panel]$HidePanel)
    $HidePanel.Visible = $false
    $ShowPanel.Visible = $true
    $ShowPanel.BringToFront()
}

function Set-FlowChildWidths {
    param([System.Windows.Forms.FlowLayoutPanel]$Flow)
    $usableWidth = [Math]::Max(400, $Flow.ClientSize.Width - 35)
    foreach ($ctl in $Flow.Controls) {
        $ctl.Width = $usableWidth
    }
}

# ============================================================
# MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "#Companyname# - Entra / Microsoft 365 User Management"
$form.Size = New-Object System.Drawing.Size(1240, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1180, 840)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# HOME PANEL
$panelHome = New-Object System.Windows.Forms.Panel
$panelHome.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($panelHome)

$titleHome = New-Object System.Windows.Forms.Label
$titleHome.Text = "User Management"
$titleHome.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$titleHome.AutoSize = $true
$titleHome.Location = New-Object System.Drawing.Point(40, 30)
$panelHome.Controls.Add($titleHome)

$subtitleHome = New-Object System.Windows.Forms.Label
$subtitleHome.Text = "Choose an action to continue."
$subtitleHome.Font = $script:font
$subtitleHome.AutoSize = $true
$subtitleHome.Location = New-Object System.Drawing.Point(42, 70)
$panelHome.Controls.Add($subtitleHome)

$homeActions = New-Object System.Windows.Forms.GroupBox
$homeActions.Text = "Actions"
$homeActions.Font = $script:fontBold
$homeActions.Size = New-Object System.Drawing.Size(420, 320)
$homeActions.Location = New-Object System.Drawing.Point(40, 120)
$panelHome.Controls.Add($homeActions)

$btnConnectModules = New-Object System.Windows.Forms.Button
$btnConnectModules.Text = "Connect Modules"
$btnConnectModules.Size = New-Object System.Drawing.Size(360, 42)
$btnConnectModules.Location = New-Object System.Drawing.Point(25, 35)
$btnConnectModules.Font = $script:fontBold
$homeActions.Controls.Add($btnConnectModules)

$btnGoOnboard = New-Object System.Windows.Forms.Button
$btnGoOnboard.Text = "Onboard User"
$btnGoOnboard.Size = New-Object System.Drawing.Size(360, 42)
$btnGoOnboard.Location = New-Object System.Drawing.Point(25, 90)
$btnGoOnboard.Font = $script:fontBold
$homeActions.Controls.Add($btnGoOnboard)

$btnTransferUser = New-Object System.Windows.Forms.Button
$btnTransferUser.Text = "Transfer User"
$btnTransferUser.Size = New-Object System.Drawing.Size(360, 42)
$btnTransferUser.Location = New-Object System.Drawing.Point(25, 145)
$btnGoOnboard.Font = $script:fontBold
$homeActions.Controls.Add($btnTransferUser)

$btnOffboardUser = New-Object System.Windows.Forms.Button
$btnOffboardUser.Text = "Offboard User"
$btnOffboardUser.Size = New-Object System.Drawing.Size(360, 42)
$btnOffboardUser.Location = New-Object System.Drawing.Point(25, 200)
$btnOffboardUser.Font = $script:font
$homeActions.Controls.Add($btnOffboardUser)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Size = New-Object System.Drawing.Size(360, 32)
$btnExit.Location = New-Object System.Drawing.Point(25, 258)
$btnExit.Font = $script:font
$homeActions.Controls.Add($btnExit)

$grpConnectionStatusHome = New-Object System.Windows.Forms.GroupBox
$grpConnectionStatusHome.Text = "Connection Status"
$grpConnectionStatusHome.Font = $script:fontBold
$grpConnectionStatusHome.Size = New-Object System.Drawing.Size(420, 215)
$grpConnectionStatusHome.Location = New-Object System.Drawing.Point(40, 465)
$panelHome.Controls.Add($grpConnectionStatusHome)

$lblSignedInAdminHome = New-Object System.Windows.Forms.Label
$lblSignedInAdminHome.Text = "Admin: Not signed in yet"
$lblSignedInAdminHome.AutoSize = $true
$lblSignedInAdminHome.Font = $script:font
$lblSignedInAdminHome.Location = New-Object System.Drawing.Point(15, 27)
$grpConnectionStatusHome.Controls.Add($lblSignedInAdminHome)

$statusHomeLabels = @{}
$statusItemsHome = @(
    @{ Name="Microsoft Graph"; X=15; Y=55 },
    @{ Name="Exchange Online"; X=15; Y=80 },
    @{ Name="Graph Auth"; X=15; Y=125 },
    @{ Name="Graph Users"; X=15; Y=150 },
    @{ Name="Graph Groups"; X=15; Y=175 },
    @{ Name="Graph Directory"; X=215; Y=125 },
    @{ Name="Exchange"; X=215; Y=150 }
)

$lblModuleHeaderHome = New-Object System.Windows.Forms.Label
$lblModuleHeaderHome.Text = "Module dependencies"
$lblModuleHeaderHome.AutoSize = $true
$lblModuleHeaderHome.Font = $script:fontBold
$lblModuleHeaderHome.Location = New-Object System.Drawing.Point(15, 105)
$grpConnectionStatusHome.Controls.Add($lblModuleHeaderHome)

foreach ($item in $statusItemsHome) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $true
    $lbl.Font = $script:font
    $lbl.Location = New-Object System.Drawing.Point($item.X, $item.Y)
    $grpConnectionStatusHome.Controls.Add($lbl)
    $statusHomeLabels[$item.Name] = $lbl
}
$script:StatusHomeLabels = $statusHomeLabels

$txtLogHome = New-Object System.Windows.Forms.TextBox
$txtLogHome.Location = New-Object System.Drawing.Point(520, 120)
$txtLogHome.Size = New-Object System.Drawing.Size(650, 520)
$txtLogHome.Multiline = $true
$txtLogHome.ScrollBars = "Vertical"
$txtLogHome.ReadOnly = $true
$txtLogHome.Font = $script:fontMono
$panelHome.Controls.Add($txtLogHome)
Write-Log $txtLogHome "Landing page loaded. Click 'Connect Modules' first."
Write-Log $txtLogHome ("Loaded {0} role(s) from CSV." -f $RoleMap.Count)

# ONBOARD PANEL
$panelOnboard = New-Object System.Windows.Forms.Panel
$panelOnboard.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelOnboard.Visible = $false
$form.Controls.Add($panelOnboard)

$panelTransfer = New-Object System.Windows.Forms.Panel
$panelTransfer.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelTransfer.Visible = $false
$form.Controls.Add($panelTransfer)

$lblTransferUser = New-Object System.Windows.Forms.Label
$lblTransferUser.AutoSize = $true
$lblTransferUser.Text = "User Email Address"
$lblTransferUser.Location = New-Object System.Drawing.Point(20,55)
$panelTransfer.Controls.Add($lblTransferUser)


$lblTransferRole = New-Object System.Windows.Forms.Label
$lblTransferRole.Text = "Role"
$lblTransferRole.AutoSize = $true
$lblTransferRole.Location = New-Object System.Drawing.Point(20,135)
$panelTransfer.Controls.Add($lblTransferRole)


$txtTransferEmail = New-Object System.Windows.Forms.TextBox
$txtTransferEmail.Location = New-Object System.Drawing.Point(20,80)
$txtTransferEmail.Width = 300
$panelTransfer.Controls.Add($txtTransferEmail)

$btnTransferGet = New-Object System.Windows.Forms.Button
$btnTransferGet.Text = "Get User"
$btnTransferGet.Location = New-Object System.Drawing.Point(340,78)
$panelTransfer.Controls.Add($btnTransferGet)

$btnReviewTransfer = New-Object System.Windows.Forms.Button
$btnReviewTransfer.Text = "Review Changes"
$btnReviewTransfer.Location = New-Object System.Drawing.Point(460,78)
$btnReviewTransfer.Size   = New-Object System.Drawing.Size(140,30)

$panelTransfer.Controls.Add($btnReviewTransfer)

$btnApplyTransfer = New-Object System.Windows.Forms.Button
$btnApplyTransfer.Text = "Apply Transfer"
$btnApplyTransfer.Location = New-Object System.Drawing.Point(600,78)
$btnApplyTransfer.Size    = New-Object System.Drawing.Size(140,30)
$btnApplyTransfer.Enabled = $false
$panelTransfer.Controls.Add($btnApplyTransfer)

$cmbTransferRole = New-Object System.Windows.Forms.ComboBox
$cmbTransferRole.DropDownStyle = 'DropDownList'
$cmbTransferRole.Width = 300
$cmbTransferRole.Location = New-Object System.Drawing.Point(20,155)

[void]$cmbTransferRole.Items.AddRange(
    [string[]](Get-RoleNames)
)

$panelTransfer.Controls.Add($cmbTransferRole)

$lblTransferUserInfo = New-Object System.Windows.Forms.Label
$lblTransferUserInfo.Location = New-Object System.Drawing.Point(20,125)
$lblTransferUserInfo.Size = New-Object System.Drawing.Size(700,40)
$panelTransfer.Controls.Add($lblTransferUserInfo)


$lblTransferRole = New-Object System.Windows.Forms.Label
$lblTransferRole.Text = "Role"
$lblTransferRole.AutoSize = $true
$lblTransferRole.Location = New-Object System.Drawing.Point(20,135)
$panelTransfer.Controls.Add($lblTransferRole)


$lblTransferLog = New-Object System.Windows.Forms.Label
$lblTransferLog.Text = "Transfer Log"
$lblTransferLog.AutoSize = $true
$lblTransferLog.Location = New-Object System.Drawing.Point(20,590)
$panelTransfer.Controls.Add($lblTransferLog)

$txtTransferLog = New-Object System.Windows.Forms.TextBox
$txtTransferLog.Location = New-Object System.Drawing.Point(20,615)
$txtTransferLog.Size = New-Object System.Drawing.Size(900,150)
$txtTransferLog.Multiline = $true
$txtTransferLog.ScrollBars = "Vertical"
$txtTransferLog.ReadOnly = $true
$txtTransferLog.Font = $script:fontMono
$panelTransfer.Controls.Add($txtTransferLog)


$lstTransferAdd = New-Object System.Windows.Forms.ListBox
$lstTransferAdd.Location = New-Object System.Drawing.Point(20,180)
$lstTransferAdd.Size = New-Object System.Drawing.Size(900,350)
$panelTransfer.Controls.Add($lstTransferAdd)

$lstTransferKeep = New-Object System.Windows.Forms.ListBox
$lstTransferKeep.Location = New-Object System.Drawing.Point(290,180)
$lstTransferKeep.Size = New-Object System.Drawing.Size(250,250)
$panelTransfer.Controls.Add($lstTransferKeep)

$lstTransferRemove = New-Object System.Windows.Forms.ListBox
$lstTransferRemove.Location = New-Object System.Drawing.Point(560,180)
$lstTransferRemove.Size = New-Object System.Drawing.Size(250,250)
$panelTransfer.Controls.Add($lstTransferRemove)

$btnTransferBack = New-Object System.Windows.Forms.Button
$btnTransferBack.Text = "Back"
$btnTransferBack.Size = New-Object System.Drawing.Size(75,30)
$btnTransferBack.Location = New-Object System.Drawing.Point(20,10)
$panelTransfer.Controls.Add($btnTransferBack)

$topBar = New-Object System.Windows.Forms.Panel
$topBar.Dock = [System.Windows.Forms.DockStyle]::Top
$topBar.Height = 115
$panelOnboard.Controls.Add($topBar)

$btnBackHome = New-Object System.Windows.Forms.Button
$btnBackHome.Text = "Back"
$btnBackHome.Size = New-Object System.Drawing.Size(100, 32)
$btnBackHome.Location = New-Object System.Drawing.Point(20, 14)
$btnBackHome.Font = $script:font
$topBar.Controls.Add($btnBackHome)

$btnCreateUser = New-Object System.Windows.Forms.Button
$btnCreateUser.Text = "Create User"
$btnCreateUser.Size = New-Object System.Drawing.Size(130, 32)
$btnCreateUser.Location = New-Object System.Drawing.Point(135, 14)
$btnCreateUser.Font = $script:fontBold
$btnCreateUser.Enabled = $false
$topBar.Controls.Add($btnCreateUser)


$grpConnectionStatusOnboard = New-Object System.Windows.Forms.GroupBox
$grpConnectionStatusOnboard.Text = "Connection Status"
$grpConnectionStatusOnboard.Font = $script:fontBold
$grpConnectionStatusOnboard.Size = New-Object System.Drawing.Size(870, 95)
$grpConnectionStatusOnboard.Location = New-Object System.Drawing.Point(290, 10)
$topBar.Controls.Add($grpConnectionStatusOnboard)

$lblSignedInAdminOnboard = New-Object System.Windows.Forms.Label
$lblSignedInAdminOnboard.Text = "Admin: Not signed in yet"
$lblSignedInAdminOnboard.AutoSize = $true
$lblSignedInAdminOnboard.Font = $script:font
$lblSignedInAdminOnboard.Location = New-Object System.Drawing.Point(12, 22)
$grpConnectionStatusOnboard.Controls.Add($lblSignedInAdminOnboard)

$statusOnboardLabels = @{}
$statusItemsOnboard = @(
    @{ Name="Microsoft Graph"; X=12; Y=48 },
    @{ Name="Exchange Online"; X=12; Y=70 },
    @{ Name="Graph Auth"; X=210; Y=48 },
    @{ Name="Graph Users"; X=210; Y=70 },
    @{ Name="Graph Groups"; X=405; Y=48 },
    @{ Name="Graph Directory"; X=405; Y=70 },
    @{ Name="Exchange"; X=640; Y=48 }
)
foreach ($item in $statusItemsOnboard) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $true
    $lbl.Font = $script:font
    $lbl.Location = New-Object System.Drawing.Point($item.X, $item.Y)
    $grpConnectionStatusOnboard.Controls.Add($lbl)
    $statusOnboardLabels[$item.Name] = $lbl
}
$script:StatusOnboardLabels = $statusOnboardLabels

$grpLogOnboard = New-Object System.Windows.Forms.GroupBox
$grpLogOnboard.Text = "Log Output"
$grpLogOnboard.Font = $script:fontBold
$grpLogOnboard.Dock = [System.Windows.Forms.DockStyle]::Bottom
$grpLogOnboard.Height = 80
$panelOnboard.Controls.Add($grpLogOnboard)

$txtLogOnboard = New-Object System.Windows.Forms.TextBox
$txtLogOnboard.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLogOnboard.Multiline = $true
$txtLogOnboard.ScrollBars = "Vertical"
$txtLogOnboard.ReadOnly = $true
$txtLogOnboard.Font = $script:fontMono
$grpLogOnboard.Controls.Add($txtLogOnboard)

$flowOnboard = New-Object System.Windows.Forms.FlowLayoutPanel
$flowOnboard.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowOnboard.AutoScroll = $true
$flowOnboard.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$flowOnboard.WrapContents = $false
$flowOnboard.Padding = New-Object System.Windows.Forms.Padding(20,115,20,10)
$panelOnboard.Controls.Add($flowOnboard)

$grpAccount = New-Object System.Windows.Forms.GroupBox
$grpAccount.Text = "User Account"
$grpAccount.Font = $script:fontBold
$grpAccount.AutoSize = $false
$grpAccount.Height = 255
$flowOnboard.Controls.Add($grpAccount)

$tblAccount = New-Object System.Windows.Forms.TableLayoutPanel
$tblAccount.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblAccount.ColumnCount = 2
$tblAccount.RowCount = 6
$tblAccount.Padding = New-Object System.Windows.Forms.Padding(10)
[void]$tblAccount.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 28)))
[void]$tblAccount.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 72)))
for ($i=0; $i -lt 6; $i++) { [void]$tblAccount.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) }
$grpAccount.Controls.Add($tblAccount)

$txtFirstName = New-TableTextBox
$txtLastName = New-TableTextBox
$txtDisplayName = New-TableTextBox
$txtPassword = New-TableTextBox -Password $true
$txtPreviewUpn = New-TableTextBox -ReadOnly $true
$txtPreviewAlias = New-TableTextBox -ReadOnly $true
Add-TableRow -Table $tblAccount -Row 0 -LabelText "First Name" -Control $txtFirstName
Add-TableRow -Table $tblAccount -Row 1 -LabelText "Last Name" -Control $txtLastName
Add-TableRow -Table $tblAccount -Row 2 -LabelText "Display Name" -Control $txtDisplayName
Add-TableRow -Table $tblAccount -Row 3 -LabelText "Password" -Control $txtPassword
Add-TableRow -Table $tblAccount -Row 4 -LabelText "Preview Email / UPN" -Control $txtPreviewUpn
Add-TableRow -Table $tblAccount -Row 5 -LabelText "Mail Alias" -Control $txtPreviewAlias

$grpEmployment = New-Object System.Windows.Forms.GroupBox
$grpEmployment.Text = "Employment Details"
$grpEmployment.Font = $script:fontBold
$grpEmployment.AutoSize = $false
$grpEmployment.Height = 405
$flowOnboard.Controls.Add($grpEmployment)

$tblEmployment = New-Object System.Windows.Forms.TableLayoutPanel
$tblEmployment.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblEmployment.ColumnCount = 2
$tblEmployment.RowCount = 10
$tblEmployment.Padding = New-Object System.Windows.Forms.Padding(10)
[void]$tblEmployment.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 28)))
[void]$tblEmployment.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 72)))
for ($i=0; $i -lt 10; $i++) { [void]$tblEmployment.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) }
$grpEmployment.Controls.Add($tblEmployment)

$txtJobTitle = New-TableTextBox
$txtDepartment = New-TableTextBox
$cmbOffice = New-TableComboBox -Items (Get-OfficeNames)
$txtManagerUpn = New-TableTextBox
$txtUsageLocation = New-TableTextBox
$txtUsageLocation.Text = $DefaultUsageLocation
$txtCompany = New-TableTextBox
$txtCompany.Text = $DefaultCompany
$txtEmployeeId = New-TableTextBox
$txtBusinessPhone = New-TableTextBox
$txtMobile = New-TableTextBox
$cmbRole = New-TableComboBox -Items (Get-RoleNames)

Add-TableRow -Table $tblEmployment -Row 9 -LabelText "Job Title" -Control $txtJobTitle
Add-TableRow -Table $tblEmployment -Row 1 -LabelText "Department" -Control $txtDepartment
Add-TableRow -Table $tblEmployment -Row 2 -LabelText "Site" -Control $cmbOffice
Add-TableRow -Table $tblEmployment -Row 3 -LabelText "Manager UPN" -Control $txtManagerUpn
Add-TableRow -Table $tblEmployment -Row 4 -LabelText "Usage Location" -Control $txtUsageLocation
Add-TableRow -Table $tblEmployment -Row 5 -LabelText "Company" -Control $txtCompany
Add-TableRow -Table $tblEmployment -Row 6 -LabelText "Employee ID" -Control $txtEmployeeId
Add-TableRow -Table $tblEmployment -Row 7 -LabelText "Business Phone" -Control $txtBusinessPhone
Add-TableRow -Table $tblEmployment -Row 8 -LabelText "Mobile" -Control $txtMobile
Add-TableRow -Table $tblEmployment -Row 0 -LabelText "Role" -Control $cmbRole

$grpManual = New-Object System.Windows.Forms.GroupBox
$grpManual.Text = "Template Role access"
$grpManual.Font = $script:fontBold
$grpManual.AutoSize = $false
$grpManual.Height = 610
$flowOnboard.Controls.Add($grpManual)

$tblManual = New-Object System.Windows.Forms.TableLayoutPanel
$tblManual.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblManual.ColumnCount = 1
$tblManual.RowCount = 12
$tblManual.Padding = New-Object System.Windows.Forms.Padding(10)
[void]$tblManual.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$manualHeights = @(30,90,30,90,30,90,30,90,30,90,35,40)
foreach ($h in $manualHeights) { [void]$tblManual.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $h))) }
$grpManual.Controls.Add($tblManual)

function New-ManualLabel {
    param([string]$Text)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding(10,6,10,2)
    $lbl.Font = $script:font
    return $lbl
}

$txtManualEntra    = New-MultiTextBox; $txtManualEntra.ReadOnly = $true
$txtManualM365     = New-MultiTextBox; $txtManualM365.ReadOnly = $true
$txtManualDL       = New-MultiTextBox; $txtManualDL.ReadOnly = $true
$txtManualShared   = New-MultiTextBox; $txtManualShared.ReadOnly = $true
$txtManualCalendar = New-MultiTextBox; $txtManualCalendar.ReadOnly = $true

$tblManual.Controls.Add((New-ManualLabel "Entra Groups"), 0, 0)
$tblManual.Controls.Add($txtManualEntra, 0, 1)
$tblManual.Controls.Add((New-ManualLabel "M365 Groups"), 0, 2)
$tblManual.Controls.Add($txtManualM365, 0, 3)
$tblManual.Controls.Add((New-ManualLabel "Distribution Groups"), 0, 4)
$tblManual.Controls.Add($txtManualDL, 0, 5)
$tblManual.Controls.Add((New-ManualLabel "Shared Mailboxes"), 0, 6)
$tblManual.Controls.Add($txtManualShared, 0, 7)
$tblManual.Controls.Add((New-ManualLabel "Calendar Access"), 0, 8)
$tblManual.Controls.Add($txtManualCalendar, 0, 9)

$lblManualHint = New-Object System.Windows.Forms.Label
$lblManualHint.Text = "Values are loaded from the selected role in userroles.csv. Calendar access is preview-only in this version."
$lblManualHint.AutoSize = $true
$lblManualHint.Margin = New-Object System.Windows.Forms.Padding(10,8,10,2)
$lblManualHint.Font = $script:font
$tblManual.Controls.Add($lblManualHint, 0, 10)

$lblRoleInfo = New-Object System.Windows.Forms.Label
$lblRoleInfo.Text = "Selected role description will show here."
$lblRoleInfo.AutoSize = $true
$lblRoleInfo.Margin = New-Object System.Windows.Forms.Padding(10,6,10,2)
$lblRoleInfo.Font = $script:font
$tblManual.Controls.Add($lblRoleInfo, 0, 11)

# ============================================================
# UI STATE FUNCTIONS
# ============================================================
function Set-StatusLabel {
    param(
        [System.Windows.Forms.Label]$Label,
        [string]$Name,
        [bool]$IsOk,
        [string]$OkText = "Connected",
        [string]$BadText = "Not connected"
    )

    if (-not $Label) { return }

    if ($IsOk) {
        $Label.Text = ("{0}: {1}" -f $Name, $OkText)
        $Label.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    else {
        $Label.Text = ("{0}: {1}" -f $Name, $BadText)
        $Label.ForeColor = [System.Drawing.Color]::DarkRed
    }
}

function Test-ModuleAvailable {
    param([string]$ModuleName)
    return [bool](Get-Module -ListAvailable -Name $ModuleName)
}

function Get-ModuleLoadedSafe {
    param([string]$ModuleName)
    return [bool](Get-Module -Name $ModuleName -ErrorAction SilentlyContinue)
}

function Update-AdminLabels {
    $adminText = if ([string]::IsNullOrWhiteSpace($script:AdminUPN)) { "Admin: Not signed in yet" } else { "Admin: $script:AdminUPN" }
    if ($lblSignedInAdminHome)    { $lblSignedInAdminHome.Text = $adminText }
    if ($lblSignedInAdminOnboard) { $lblSignedInAdminOnboard.Text = $adminText }
}

function Set-DashboardStatus {
    param([string]$Name, [bool]$IsOk, [string]$OkText = "Connected", [string]$BadText = "Not connected")

    if ($script:StatusHomeLabels -and $script:StatusHomeLabels.ContainsKey($Name)) {
        Set-StatusLabel -Label $script:StatusHomeLabels[$Name] -Name $Name -IsOk $IsOk -OkText $OkText -BadText $BadText
    }
    if ($script:StatusOnboardLabels -and $script:StatusOnboardLabels.ContainsKey($Name)) {
        Set-StatusLabel -Label $script:StatusOnboardLabels[$Name] -Name $Name -IsOk $IsOk -OkText $OkText -BadText $BadText
    }
}

function Update-ConnectionLabels {
    $graphConnected = $false
    $exoConnected   = $false

    # Do not autoload Graph from the timer. Only check Graph context if the auth module is loaded.
    if (Get-ModuleLoadedSafe -ModuleName "Microsoft.Graph.Authentication") {
        try {
            $graphContext = Get-MgContext -ErrorAction SilentlyContinue
            if ($graphContext) {
                $graphConnected = $true
                if ($graphContext.Account) { $script:AdminUPN = $graphContext.Account }
            }
        }
        catch {}
    }

    if (Get-ModuleLoadedSafe -ModuleName "ExchangeOnlineManagement") {
        try {
            $exoConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($exoConnection) { $exoConnected = $true }
        }
        catch {}
    }

    Update-AdminLabels

    Set-DashboardStatus -Name "Microsoft Graph" -IsOk $graphConnected
    Set-DashboardStatus -Name "Exchange Online" -IsOk $exoConnected

    Set-DashboardStatus -Name "Graph Auth"      -IsOk (Test-ModuleAvailable -ModuleName "Microsoft.Graph.Authentication") -OkText "Installed" -BadText "Missing"
       Set-DashboardStatus -Name "Exchange"        -IsOk (Test-ModuleAvailable -ModuleName "ExchangeOnlineManagement") -OkText "Installed" -BadText "Missing"



    $script:ServicesConnected = ($graphConnected -and $exoConnected)
    $btnCreateUser.Enabled = $script:ServicesConnected
}


function Update-PreviewFields {
    $alias = Get-MailAlias -FirstName $txtFirstName.Text -LastName $txtLastName.Text
    $upn   = Get-UserPrincipalName -FirstName $txtFirstName.Text -LastName $txtLastName.Text

    $txtPreviewAlias.Text = $alias
    $txtPreviewUpn.Text   = $upn

    if ((-not (Is-Blank $txtFirstName.Text)) -and (-not (Is-Blank $txtLastName.Text))) {
        $txtDisplayName.Text = "$($txtFirstName.Text.Trim()) $($txtLastName.Text.Trim())"
    }
}

function Set-PreviewText {
    param(
        [System.Windows.Forms.TextBox]$TextBox,
        [array]$Values,
        [string]$EmptyText
    )

    if ($Values -and $Values.Count -gt 0) {
        $TextBox.Text = ($Values | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique) -join "`r`n"
    } else {
        $TextBox.Text = $EmptyText
    }
}

function Update-RoleDescription {
    $selectedRole = $cmbRole.Text

    if (-not [string]::IsNullOrWhiteSpace($selectedRole) -and $RoleMap.ContainsKey($selectedRole)) {
        $role = $RoleMap[$selectedRole]

        
        # Auto-fill Job Title from selected role
        $txtJobTitle.Text = $selectedRole

        $lblRoleInfo.Text = "Role: $selectedRole"

        Set-PreviewText -TextBox $txtManualEntra    -Values $role.EntraGroups        -EmptyText "(No Entra Groups)"
        Set-PreviewText -TextBox $txtManualM365     -Values $role.M365Groups         -EmptyText "(No M365 Groups)"
        Set-PreviewText -TextBox $txtManualDL       -Values $role.DistributionGroups -EmptyText "(No Distribution Groups)"
        Set-PreviewText -TextBox $txtManualShared   -Values @($role.SharedMailboxes | ForEach-Object { $_.Mailbox }) -EmptyText "(No Shared Mailboxes)"
        Set-PreviewText -TextBox $txtManualCalendar -Values $role.CalendarAccess     -EmptyText "(No Calendar Access)"
    } else {
        $lblRoleInfo.Text = "Selected role not found."
        $txtManualEntra.Text = ""
        $txtManualM365.Text = ""
        $txtManualDL.Text = ""
        $txtManualShared.Text = ""
        $txtManualCalendar.Text = ""
    }
}

$txtFirstName.Add_TextChanged({ Update-PreviewFields })
$txtLastName.Add_TextChanged({ Update-PreviewFields })
$cmbRole.Add_SelectedIndexChanged({ Update-RoleDescription })
$flowOnboard.Add_Resize({ Set-FlowChildWidths -Flow $flowOnboard })
$form.Add_Shown({ Set-FlowChildWidths -Flow $flowOnboard })
Update-ConnectionLabels
Update-RoleDescription

$statusTimer = New-Object System.Windows.Forms.Timer
$statusTimer.Interval = 5000
$statusTimer.Add_Tick({ Update-ConnectionLabels })
$statusTimer.Start()

# Force HOME as default landing page
$panelHome.Visible = $true
$panelHome.BringToFront()
$panelOnboard.Visible = $false
$panelOnboard.SendToBack()

# ============================================================
# BUTTON EVENTS
# ============================================================
$btnConnectModules.Add_Click({

    try {

        if ($statusTimer) { $statusTimer.Stop() }

        $script:AdminUPN = $null
        Update-ConnectionLabels

        Write-Log $txtLogHome "Checking required modules..."
        Write-Log $txtLogOnboard "Checking required modules..."

        $requiredModules = @(
            "ExchangeOnlineManagement",
            "Microsoft.Graph.Authentication"
        )

        foreach ($m in $requiredModules) {

            if (Get-Module -ListAvailable -Name $m) {
                Write-Log $txtLogHome "Module OK: $m"
                Write-Log $txtLogOnboard "Module OK: $m"
            }
            else {
                throw "Required module missing: $m. Install with Install-Module $m -Scope CurrentUser"
            }
        }

        Connect-Services -UseDeviceCode:$false

        Update-ConnectionLabels

        if (Test-RequiredConnections) {

            $script:ServicesConnected = $true
            Update-ConnectionLabels

            Write-Log $txtLogHome "Modules/services connected successfully."
            Write-Log $txtLogOnboard "Modules/services connected successfully."

            [System.Windows.Forms.MessageBox]::Show(
                "Exchange Online and Microsoft Graph connected successfully."
            ) | Out-Null
        }
        else {
            $script:ServicesConnected = $false
            Update-ConnectionLabels
            throw "Connection did not complete successfully."
        }
    }
    catch {

        $script:ServicesConnected = $false
        Update-ConnectionLabels

        Write-Log $txtLogHome ("ERROR - " + $_.Exception.Message)
        Write-Log $txtLogOnboard ("ERROR - " + $_.Exception.Message)

        [System.Windows.Forms.MessageBox]::Show(
            "Connect Modules failed:`r`n$($_.Exception.Message)"
        ) | Out-Null
    }
    finally {
        Update-ConnectionLabels
        if ($statusTimer) { $statusTimer.Start() }
        if ($script:ServicesConnected) { $btnCreateUser.Enabled = $true }
    }
})
$btnGoOnboard.Add_Click({ Set-PanelVisible -ShowPanel $panelOnboard -HidePanel $panelHome })
$btnBackHome.Add_Click({ Set-PanelVisible -ShowPanel $panelHome -HidePanel $panelOnboard })

$btnTransferBack.Add_Click({

    $txtTransferEmail.Clear()
    $lblTransferUserInfo.Text = ""

    $lstTransferAdd.Items.Clear()

    $script:TransferUser = $null
    $script:TransferCurrentAccess = $null
    $script:TransferRoleAccess = $null
    $script:TransferChangesReviewed = $false

    $btnApplyTransfer.Enabled = $false

    $panelTransfer.Visible = $false
    $panelHome.Visible = $true
    $panelHome.BringToFront()

})

$btnCreateUser.Add_Click({
    try {
        if (-not $script:ServicesConnected) {
            [System.Windows.Forms.MessageBox]::Show("Please click 'Connect Modules' first.") | Out-Null
            return
        }
        if ((Is-Blank $txtFirstName.Text) -or (Is-Blank $txtLastName.Text) -or (Is-Blank $txtPassword.Text) -or (Is-Blank $txtDisplayName.Text) -or (Is-Blank $cmbOffice.Text) -or (Is-Blank $cmbRole.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please complete First Name, Last Name, Password, Display Name, Site, and Role.") | Out-Null
            return
        }
        $companyValue = if (Is-Blank $txtCompany.Text) { $DefaultCompany } else { $txtCompany.Text.Trim() }
        $usageValue   = if (Is-Blank $txtUsageLocation.Text) { $DefaultUsageLocation } else { $txtUsageLocation.Text.Trim() }
        $formData = @{
            FirstName                = $txtFirstName.Text.Trim()
            LastName                 = $txtLastName.Text.Trim()
            DisplayName              = $txtDisplayName.Text.Trim()
            Password                 = $txtPassword.Text
            JobTitle                 = $txtJobTitle.Text.Trim()
            Department               = $txtDepartment.Text.Trim()
            EmployeeId               = $txtEmployeeId.Text.Trim()
            Company                  = $companyValue
            BusinessPhone            = $txtBusinessPhone.Text.Trim()
            Mobile                   = $txtMobile.Text.Trim()
            ManagerUpn               = $txtManagerUpn.Text.Trim()
            Office                   = $cmbOffice.Text
            UsageLocation            = $usageValue
            Role                     = $cmbRole.Text

            # These are blank for now because the role preview boxes are read-only.
            # If you later add editable manual override boxes, wire those controls into these arrays.
            ManualEntraGroups        = @()
            ManualM365Groups         = @()
            ManualDistributionGroups = @()
            ManualUnifiedGroups      = @()
            ManualSharedMailboxes    = @()
        }
        $btnCreateUser.Enabled = $false
        $upn = Create-NewStarter -FormData $formData -LogBox $txtLogOnboard
        $txtPreviewUpn.Text   = $upn
        $txtPreviewAlias.Text = $upn.Split("@")[0]
        Write-Log $txtLogHome ("Onboarded user: {0}" -f $upn)
        [System.Windows.Forms.MessageBox]::Show("User onboarded successfully: $upn") | Out-Null
    } catch {
        Write-Log $txtLogOnboard ("ERROR - {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show("Onboarding failed: $($_.Exception.Message)") | Out-Null
    } finally {
        if ($script:ServicesConnected) { $btnCreateUser.Enabled = $true }
    }
})


$btnTransferUser.Add_Click({
    $panelHome.Visible = $false
    $panelOnboard.Visible = $false
    $panelTransfer.Visible = $true
    $panelTransfer.BringToFront()
})

$btnTransferGet.Add_Click({

    if ([string]::IsNullOrWhiteSpace($txtTransferEmail.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a user email address.") | Out-Null
        return
    }

    try {
        $userUpn = $txtTransferEmail.Text.Trim()

        $script:TransferUser = Get-MgUser `
            -UserId $userUpn `
            -Property Id,DisplayName,UserPrincipalName `
            -ErrorAction Stop

        $script:TransferCurrentAccess = @{
            User = $script:TransferUser
        }

        $script:TransferAccessFile = Join-Path `
            $TransferExportPath `
            "$($script:TransferUser.UserPrincipalName).csv"

        $lblTransferUserInfo.Text = "$($script:TransferUser.DisplayName) ($($script:TransferUser.UserPrincipalName))"

        Write-Log $txtTransferLog "Transfer user loaded: $($script:TransferUser.UserPrincipalName)"

        Export-CurrentUserAccess `
            -User $script:TransferUser `
            -ExportFile $script:TransferAccessFile `
            -LogBox $txtTransferLog

        [System.Windows.Forms.MessageBox]::Show("User loaded successfully and current access was exported.") | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Unable to locate or export user.`r`n$($_.Exception.Message)") | Out-Null

        $script:TransferUser = $null
        $script:TransferCurrentAccess = $null
        $script:TransferAccessFile = $null
    }
})




$btnReviewTransfer.Add_Click({

    if ($null -eq $script:TransferCurrentAccess) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please load a user first."
        ) | Out-Null
        return
    }

    if ($null -eq $cmbTransferRole.Text -or $cmbTransferRole.Text.Trim() -eq "") {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a target role."
        ) | Out-Null
        return
    }

    $targetRole = Get-RoleConfig -RoleName $cmbTransferRole.Text

    $lstTransferAdd.Items.Clear()

$targetRole = Get-RoleConfig -RoleName $cmbTransferRole.Text

$lstTransferAdd.Items.Clear()

[void]$lstTransferAdd.Items.Add("=== ENTRA GROUPS ===")

foreach ($group in $targetRole.EntraGroups) {
    [void]$lstTransferAdd.Items.Add($group)
}

[void]$lstTransferAdd.Items.Add("")

[void]$lstTransferAdd.Items.Add("=== M365 GROUPS ===")

foreach ($group in $targetRole.M365Groups) {
    [void]$lstTransferAdd.Items.Add($group)
}

[void]$lstTransferAdd.Items.Add("")

[void]$lstTransferAdd.Items.Add("=== DISTRIBUTION LISTS ===")

foreach ($group in $targetRole.DistributionGroups) {
    [void]$lstTransferAdd.Items.Add($group)
}

[void]$lstTransferAdd.Items.Add("")

[void]$lstTransferAdd.Items.Add("=== SHARED MAILBOXES ===")

foreach ($mailbox in $targetRole.SharedMailboxes) {
    [void]$lstTransferAdd.Items.Add($mailbox.Mailbox)
}

$script:TransferRoleAccess = $targetRole
$script:TransferChangesReviewed = $true
$btnApplyTransfer.Enabled = $true

})

$btnApplyTransfer.Add_Click({

    if (-not $script:TransferChangesReviewed) {
        [System.Windows.Forms.MessageBox]::Show(
            "Review Changes must be completed first."
        ) | Out-Null
        return
    }

    if ($null -eq $script:TransferUser) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please load a user first."
        ) | Out-Null
        return
    }

    if ($null -eq $script:TransferRoleAccess) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please review the selected role first."
        ) | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:TransferAccessFile) -or -not (Test-Path $script:TransferAccessFile)) {
        [System.Windows.Forms.MessageBox]::Show(
            "The transfer access export file could not be found. Please click Get User again before applying the transfer."
        ) | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will remove the user's exported current access, keep the licence group, then apply the selected role access. Continue?",
        "Confirm Transfer",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $btnApplyTransfer.Enabled = $false

        Write-Log $txtTransferLog "Starting transfer process..."
        Write-Log $txtTransferLog "User: $($script:TransferUser.UserPrincipalName)"
        Write-Log $txtTransferLog "Role: $($cmbTransferRole.Text)"
        Write-Log $txtTransferLog "Using access file: $script:TransferAccessFile"

        Remove-AccessFromCsv `
            -User $script:TransferUser `
            -CsvPath $script:TransferAccessFile `
            -LogBox $txtTransferLog

        Write-Log $txtTransferLog "Finished removal phase."

        Apply-TransferRoleAccess `
            -User $script:TransferUser `
            -RoleConfig $script:TransferRoleAccess `
            -LogBox $txtTransferLog

        Write-Log $txtTransferLog "Finished apply phase."

        [System.Windows.Forms.MessageBox]::Show(
            "Transfer completed. Please allow a few minutes for Entra and Exchange changes to appear."
        ) | Out-Null

        Write-Log $txtTransferLog "Transfer completed for $($script:TransferUser.UserPrincipalName)."
        Write-Log $txtLogHome "Transfer completed for $($script:TransferUser.UserPrincipalName)."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Transfer failed:`r`n$($_.Exception.Message)"
        ) | Out-Null

        Write-Log $txtTransferLog ("ERROR - Transfer failed: {0}" -f $_.Exception.Message)
        Write-Log $txtLogHome ("ERROR - Transfer failed: {0}" -f $_.Exception.Message)
    }
    finally {
        $btnApplyTransfer.Enabled = $true
    }
})


$btnOffboardUser.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Offboard User is not implemented yet.") | Out-Null
    Write-Log $txtLogHome "Offboard User clicked (placeholder only)."
})

$btnExit.Add_Click({ $form.Close() })

[void]$form.ShowDialog()
