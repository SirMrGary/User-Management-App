#Script Created by Gary Mundt.

# Company Role Builder - Read Only - Skips Dynamic Groups
#Update Line 15 with your storage location for the Userroles.csv  before first use.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

Connect-ExchangeOnline

# ============================================================
# CONFIG
# ============================================================
$TenantDomain = "Your Domain" #eg  Google.com or Companyname.co.nz
$DefaultRolePath = Join-Path $env:USERPROFILE "#Enter Location you are storing the#\userroles.csv"
$script:GraphConnected = $false
$script:ExchangeConnected = $false
$script:LastSnapshot = $null

# ============================================================
# COMMON HELPERS
# ============================================================
function Is-Blank {
    param([AllowNull()][string]$Value)
    return [string]::IsNullOrWhiteSpace($Value)
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
    $Box.AppendText("$time - $Message`r`n")
    $Box.SelectionStart = $Box.Text.Length
    $Box.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Join-RoleValues {
    param([array]$Values)
    if (-not $Values -or $Values.Count -eq 0) { return "" }
    return (
        $Values |
            Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) } |
            ForEach-Object { ([string]$_).Trim() } |
            Sort-Object -Unique
    ) -join ";"
}

function Set-StatusLabel {
    param(
        [System.Windows.Forms.Label]$Label,
        [bool]$Connected,
        [string]$ConnectedText,
        [string]$DisconnectedText
    )
    if ($Connected) {
        $Label.Text = $ConnectedText
        $Label.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    else {
        $Label.Text = $DisconnectedText
        $Label.ForeColor = [System.Drawing.Color]::DarkRed
    }
}

function Set-PreviewText {
    param(
        [System.Windows.Forms.TextBox]$TextBox,
        [array]$Values,
        [string]$EmptyText = ""
    )
    if ($Values -and $Values.Count -gt 0) {
        $TextBox.Text = ($Values | Sort-Object -Unique) -join "`r`n"
    }
    else {
        $TextBox.Text = $EmptyText
    }
}

# ============================================================
# CONNECTIONS
# ============================================================
function Test-RoleBuilderConnections {
    $script:GraphConnected = $false
    $script:ExchangeConnected = $false
    try { if (Get-MgContext -ErrorAction SilentlyContinue) { $script:GraphConnected = $true } } catch {}
    try { if (Get-ConnectionInformation -ErrorAction SilentlyContinue) { $script:ExchangeConnected = $true } } catch {}
}

function Connect-RoleBuilderServices {
    param([System.Windows.Forms.TextBox]$LogBox)

    Write-Log $LogBox "Checking required modules..."
    Ensure-Module -ModuleName "ExchangeOnlineManagement"
    Ensure-Module -ModuleName "Microsoft.Graph.Authentication"

    Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop

    try {
        if (Get-ConnectionInformation -ErrorAction SilentlyContinue) { $script:ExchangeConnected = $true }
    }
    catch {}

    if (-not $script:ExchangeConnected) {
        Write-Log $LogBox "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ErrorAction Stop
        $script:ExchangeConnected = $true
        Write-Log $LogBox "Exchange Online connected."
    }
    else {
        Write-Log $LogBox "Exchange Online already connected."
    }

    try {
        if (Get-MgContext -ErrorAction SilentlyContinue) { $script:GraphConnected = $true }
    }
    catch {}

    if (-not $script:GraphConnected) {
        Write-Log $LogBox "Connecting to Microsoft Graph..."
        # Read-only scopes only. This role builder does not write to Microsoft Graph.
        Connect-MgGraph -Scopes @("User.Read.All", "Group.Read.All") -NoWelcome -ContextScope Process -ErrorAction Stop
        $script:GraphConnected = $true
        Write-Log $LogBox "Microsoft Graph connected."
    }
    else {
        Write-Log $LogBox "Microsoft Graph already connected."
    }
}

# ============================================================
# GRAPH READ-ONLY SNAPSHOT
# Uses Invoke-MgGraphRequest to avoid importing Microsoft.Graph.Users/Groups workload modules.
# Dynamic groups are skipped.
# ============================================================
function Invoke-GraphGetAllPages {
    param([Parameter(Mandatory = $true)][string]$Uri)
    $items = @()
    $next = $Uri
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if ($response.value) { $items += @($response.value) }
        $next = $null
        if ($response.'@odata.nextLink') { $next = $response.'@odata.nextLink' }
    }
    return $items
}

function Get-GraphRoleMembershipSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$UserPrincipalName,
        [System.Windows.Forms.TextBox]$LogBox
    )

    $result = [ordered]@{
        EntraGroups = @()
        M365Groups  = @()
        DynamicGroupsSkipped = @()
    }

    Write-Log $LogBox "Reading Graph group memberships for $UserPrincipalName..."
    $encodedUser = [System.Uri]::EscapeDataString($UserPrincipalName)

    # membershipRule is requested as an additional dynamic-group indicator where Graph returns it.
    $uri = "https://graph.microsoft.com/v1.0/users/$encodedUser/memberOf?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,membershipRule"
    $memberObjects = Invoke-GraphGetAllPages -Uri $uri

    foreach ($obj in $memberObjects) {
        try {
            $odataType = [string]$obj.'@odata.type'
            if ($odataType -and ($odataType -notmatch 'group')) { continue }

            $displayName = [string]$obj.displayName
            if ([string]::IsNullOrWhiteSpace($displayName)) { continue }

            $groupTypes = @()
            if ($obj.groupTypes) { $groupTypes = @($obj.groupTypes) }

            $membershipRule = ""
            if ($null -ne $obj.membershipRule) { $membershipRule = [string]$obj.membershipRule }

            $isDynamicGroup = ($groupTypes -contains "DynamicMembership") -or (-not [string]::IsNullOrWhiteSpace($membershipRule))
            if ($isDynamicGroup) {
                $result.DynamicGroupsSkipped += $displayName
                Write-Log $LogBox ("Skipping dynamic group: {0}" -f $displayName)
                continue
            }

            $securityEnabled = $false
            if ($null -ne $obj.securityEnabled) { $securityEnabled = [bool]$obj.securityEnabled }

            if ($groupTypes -contains "Unified") {
                $result.M365Groups += $displayName
            }
            elseif ($securityEnabled) {
                $result.EntraGroups += $displayName
            }
        }
        catch {
            Write-Log $LogBox ("WARN - Could not process Graph membership object: {0}" -f $_.Exception.Message)
        }
    }

    $result.EntraGroups = @($result.EntraGroups | Sort-Object -Unique)
    $result.M365Groups  = @($result.M365Groups  | Sort-Object -Unique)
    $result.DynamicGroupsSkipped = @($result.DynamicGroupsSkipped | Sort-Object -Unique)
    return $result
}

# ============================================================
# EXCHANGE READ-ONLY SNAPSHOT
# ============================================================
function Get-ExchangeCompareValuesForUser {
    param([Parameter(Mandatory = $true)][string]$UserPrincipalName)
    try {
        $recipient = Get-Recipient -Identity $UserPrincipalName -ErrorAction Stop
        return @(
            $UserPrincipalName,
            $recipient.PrimarySmtpAddress,
            $recipient.WindowsEmailAddress,
            $recipient.Name,
            $recipient.DisplayName,
            $recipient.Alias,
            $recipient.ExternalDirectoryObjectId
        ) | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLower().Trim() } | Sort-Object -Unique
    }
    catch {
        return @($UserPrincipalName.ToLower().Trim())
    }
}

function Get-DistributionGroupsForUser {
    param(
        [Parameter(Mandatory = $true)][string]$UserPrincipalName,
        [System.Windows.Forms.TextBox]$LogBox
    )
    $distributionGroups = @()
    $compareValues = Get-ExchangeCompareValuesForUser -UserPrincipalName $UserPrincipalName
    Write-Log $LogBox "Reading Exchange distribution group memberships for $UserPrincipalName..."
    $allDls = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop

    foreach ($dl in $allDls) {
        try {
            $members = Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited -ErrorAction Stop
            foreach ($member in $members) {
                $memberValues = @(
                    $member.PrimarySmtpAddress,
                    $member.WindowsEmailAddress,
                    $member.Name,
                    $member.DisplayName,
                    $member.Alias,
                    $member.ExternalDirectoryObjectId
                ) | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLower().Trim() }
                if ($memberValues | Where-Object { $compareValues -contains $_ }) {
                    $distributionGroups += $dl.DisplayName
                    break
                }
            }
        }
        catch {
            Write-Log $LogBox ("WARN - Could not read members for distribution group '{0}': {1}" -f $dl.DisplayName, $_.Exception.Message)
        }
    }
    return @($distributionGroups | Sort-Object -Unique)
}

function Get-SharedMailboxAccessForUser {
    param(
        [Parameter(Mandatory = $true)][string]$UserPrincipalName,
        [System.Windows.Forms.TextBox]$LogBox
    )
    $sharedMailboxes = @()
    Write-Log $LogBox "Reading shared mailbox access for $UserPrincipalName..."
    $shared = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop

    foreach ($mbx in $shared) {
        $mailboxIdentity = $mbx.PrimarySmtpAddress
        if (-not $mailboxIdentity) { $mailboxIdentity = $mbx.Identity }
        $hasAccess = $false

        try {
            $perms = Get-MailboxPermission -Identity $mailboxIdentity -User $UserPrincipalName -ErrorAction SilentlyContinue
            foreach ($perm in $perms) {
                if (($perm.AccessRights -contains "FullAccess") -and (-not $perm.Deny)) { $hasAccess = $true }
            }
        }
        catch {
            Write-Log $LogBox ("WARN - Could not read FullAccess for mailbox '{0}': {1}" -f $mailboxIdentity, $_.Exception.Message)
        }

        try {
            $sendAsPerms = Get-RecipientPermission -Identity $mailboxIdentity -Trustee $UserPrincipalName -ErrorAction SilentlyContinue
            foreach ($perm in $sendAsPerms) {
                if (($perm.AccessRights -contains "SendAs") -and (-not $perm.Deny)) { $hasAccess = $true }
            }
        }
        catch {
            Write-Log $LogBox ("WARN - Could not read SendAs for mailbox '{0}': {1}" -f $mailboxIdentity, $_.Exception.Message)
        }

        if ($hasAccess) { $sharedMailboxes += ([string]$mailboxIdentity) }
    }
    return @($sharedMailboxes | Sort-Object -Unique)
}

function Build-RoleSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$RoleName,
        [Parameter(Mandatory = $true)][string]$SourceUser,
        [System.Windows.Forms.TextBox]$LogBox
    )
    if (-not $script:GraphConnected -or -not $script:ExchangeConnected) { throw "Connect to Exchange Online and Microsoft Graph first." }
    Write-Log $LogBox "Building role '$RoleName' from source user $SourceUser..."
    $graph = Get-GraphRoleMembershipSnapshot -UserPrincipalName $SourceUser -LogBox $LogBox
    $dls = Get-DistributionGroupsForUser -UserPrincipalName $SourceUser -LogBox $LogBox
    $shared = Get-SharedMailboxAccessForUser -UserPrincipalName $SourceUser -LogBox $LogBox

    $snapshot = [ordered]@{
        Role                 = $RoleName
        SourceUser           = $SourceUser
        EntraGroups          = @($graph.EntraGroups)
        M365Groups           = @($graph.M365Groups)
        DistributionGroups   = @($dls)
        SharedMailboxes      = @($shared)
        CalendarAccess       = @()
        DynamicGroupsSkipped = @($graph.DynamicGroupsSkipped)
    }

    Write-Log $LogBox "Snapshot complete. Entra=$($snapshot.EntraGroups.Count), M365=$($snapshot.M365Groups.Count), DL=$($snapshot.DistributionGroups.Count), Shared=$($snapshot.SharedMailboxes.Count), Dynamic skipped=$($snapshot.DynamicGroupsSkipped.Count)."
    return $snapshot
}

# ============================================================
# SAVE ROLE FILE
# ============================================================
function New-RoleRowObject {
    param([hashtable]$Snapshot)
    return [pscustomobject]@{
        Role       = $Snapshot.Role
        Entragroup = Join-RoleValues -Values $Snapshot.EntraGroups
        M365       = Join-RoleValues -Values $Snapshot.M365Groups
        Dlist      = Join-RoleValues -Values $Snapshot.DistributionGroups
        Shared     = Join-RoleValues -Values $Snapshot.SharedMailboxes
        Calendar   = Join-RoleValues -Values $Snapshot.CalendarAccess
    }
}

function Save-GeneratedRoleToCsv {
    param([hashtable]$Snapshot, [string]$RolePath, [System.Windows.Forms.TextBox]$LogBox)
    if ([string]::IsNullOrWhiteSpace($RolePath)) { throw "Role file path is blank." }
    $newRow = New-RoleRowObject -Snapshot $Snapshot
    $folder = Split-Path $RolePath -Parent
    if (-not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }

    if (Test-Path $RolePath) {
        $rows = @(Import-Csv -Path $RolePath)
        $existing = $rows | Where-Object { $_.Role -eq $Snapshot.Role }
        if ($existing) {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Role '$($Snapshot.Role)' already exists. Replace it?",
                "Replace Existing Role",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log $LogBox "Save cancelled. Existing role was not replaced."
                return
            }
            $rows = @($rows | Where-Object { $_.Role -ne $Snapshot.Role })
        }
    }
    else { $rows = @() }

    $rows += $newRow
    $rows | Sort-Object Role | Export-Csv -Path $RolePath -NoTypeInformation -Encoding UTF8
    Write-Log $LogBox "Saved role '$($Snapshot.Role)' to $RolePath."
}

function Save-GeneratedRoleToExcel {
    param([hashtable]$Snapshot, [string]$RolePath, [System.Windows.Forms.TextBox]$LogBox)
    Ensure-Module -ModuleName "ImportExcel"
    Import-Module ImportExcel -Force -ErrorAction Stop
    $newRow = New-RoleRowObject -Snapshot $Snapshot
    $folder = Split-Path $RolePath -Parent
    if (-not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }

    if (Test-Path $RolePath) {
        $rows = @(Import-Excel -Path $RolePath)
        $existing = $rows | Where-Object { $_.Role -eq $Snapshot.Role }
        if ($existing) {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Role '$($Snapshot.Role)' already exists. Replace it?",
                "Replace Existing Role",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log $LogBox "Save cancelled. Existing role was not replaced."
                return
            }
            $rows = @($rows | Where-Object { $_.Role -ne $Snapshot.Role })
        }
    }
    else { $rows = @() }

    $rows += $newRow
    $rows = $rows | Sort-Object Role
    $rows | Export-Excel -Path $RolePath -WorksheetName "Roles" -AutoSize -ClearSheet
    Write-Log $LogBox "Saved role '$($Snapshot.Role)' to $RolePath."
}

function Save-GeneratedRole {
    param([hashtable]$Snapshot, [string]$RolePath, [System.Windows.Forms.TextBox]$LogBox)
    $ext = [System.IO.Path]::GetExtension($RolePath).ToLower()
    switch ($ext) {
        ".csv"  { Save-GeneratedRoleToCsv -Snapshot $Snapshot -RolePath $RolePath -LogBox $LogBox }
        ".xlsx" { Save-GeneratedRoleToExcel -Snapshot $Snapshot -RolePath $RolePath -LogBox $LogBox }
        ".xlsm" { Save-GeneratedRoleToExcel -Snapshot $Snapshot -RolePath $RolePath -LogBox $LogBox }
        default { throw "Unsupported role file type '$ext'. Use .csv, .xlsx, or .xlsm." }
    }
}

# ============================================================
# UI
# ============================================================
$font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$fontBold = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$fontMono = New-Object System.Drawing.Font("Consolas", 9)

$form = New-Object System.Windows.Forms.Form
$form.Text = "Company - Role Builder (Read Only - Skips Dynamic Groups)"
$form.Size = New-Object System.Drawing.Size(1180, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1120, 800)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

$title = New-Object System.Windows.Forms.Label
$title.Text = "Role Builder"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(25, 20)
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Build role templates by reading permissions from an existing user. Dynamic groups are skipped."
$subtitle.Font = $font
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(28, 60)
$form.Controls.Add($subtitle)

$grpInputs = New-Object System.Windows.Forms.GroupBox
$grpInputs.Text = "Role Source"
$grpInputs.Font = $fontBold
$grpInputs.Location = New-Object System.Drawing.Point(25, 95)
$grpInputs.Size = New-Object System.Drawing.Size(520, 205)
$form.Controls.Add($grpInputs)

$lblRoleName = New-Object System.Windows.Forms.Label
$lblRoleName.Text = "New Role Name"
$lblRoleName.AutoSize = $true
$lblRoleName.Location = New-Object System.Drawing.Point(18, 32)
$lblRoleName.Font = $font
$grpInputs.Controls.Add($lblRoleName)

$txtRoleName = New-Object System.Windows.Forms.TextBox
$txtRoleName.Location = New-Object System.Drawing.Point(170, 28)
$txtRoleName.Size = New-Object System.Drawing.Size(330, 24)
$txtRoleName.Font = $font
$grpInputs.Controls.Add($txtRoleName)

$lblSourceUser = New-Object System.Windows.Forms.Label
$lblSourceUser.Text = "Source User UPN"
$lblSourceUser.AutoSize = $true
$lblSourceUser.Location = New-Object System.Drawing.Point(18, 70)
$lblSourceUser.Font = $font
$grpInputs.Controls.Add($lblSourceUser)

$txtSourceUser = New-Object System.Windows.Forms.TextBox
$txtSourceUser.Location = New-Object System.Drawing.Point(170, 66)
$txtSourceUser.Size = New-Object System.Drawing.Size(330, 24)
$txtSourceUser.Font = $font
$txtSourceUser.Text = "@$TenantDomain"
$grpInputs.Controls.Add($txtSourceUser)

$lblRolePath = New-Object System.Windows.Forms.Label
$lblRolePath.Text = "Role File"
$lblRolePath.AutoSize = $true
$lblRolePath.Location = New-Object System.Drawing.Point(18, 108)
$lblRolePath.Font = $font
$grpInputs.Controls.Add($lblRolePath)

$txtRolePath = New-Object System.Windows.Forms.TextBox
$txtRolePath.Location = New-Object System.Drawing.Point(170, 104)
$txtRolePath.Size = New-Object System.Drawing.Size(280, 24)
$txtRolePath.Font = $font
$txtRolePath.Text = $DefaultRolePath
$grpInputs.Controls.Add($txtRolePath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Location = New-Object System.Drawing.Point(455, 103)
$btnBrowse.Size = New-Object System.Drawing.Size(55, 26)
$btnBrowse.Font = $font
$grpInputs.Controls.Add($btnBrowse)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect Services"
$btnConnect.Location = New-Object System.Drawing.Point(20, 150)
$btnConnect.Size = New-Object System.Drawing.Size(145, 36)
$btnConnect.Font = $fontBold
$grpInputs.Controls.Add($btnConnect)

$btnBuild = New-Object System.Windows.Forms.Button
$btnBuild.Text = "Build Role"
$btnBuild.Location = New-Object System.Drawing.Point(180, 150)
$btnBuild.Size = New-Object System.Drawing.Size(145, 36)
$btnBuild.Font = $fontBold
$grpInputs.Controls.Add($btnBuild)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save Role"
$btnSave.Location = New-Object System.Drawing.Point(340, 150)
$btnSave.Size = New-Object System.Drawing.Size(145, 36)
$btnSave.Font = $fontBold
$btnSave.Enabled = $false
$grpInputs.Controls.Add($btnSave)

$grpStatus = New-Object System.Windows.Forms.GroupBox
$grpStatus.Text = "Connection Status"
$grpStatus.Font = $fontBold
$grpStatus.Location = New-Object System.Drawing.Point(565, 95)
$grpStatus.Size = New-Object System.Drawing.Size(570, 205)
$form.Controls.Add($grpStatus)

$lblGraph = New-Object System.Windows.Forms.Label
$lblGraph.AutoSize = $true
$lblGraph.Location = New-Object System.Drawing.Point(18, 32)
$lblGraph.Font = $font
$grpStatus.Controls.Add($lblGraph)

$lblExchange = New-Object System.Windows.Forms.Label
$lblExchange.AutoSize = $true
$lblExchange.Location = New-Object System.Drawing.Point(18, 62)
$lblExchange.Font = $font
$grpStatus.Controls.Add($lblExchange)

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "Read-only against M365: discovers access and saves a role row. Dynamic groups are skipped and logged."
$lblNote.AutoSize = $false
$lblNote.Size = New-Object System.Drawing.Size(530, 80)
$lblNote.Location = New-Object System.Drawing.Point(18, 105)
$lblNote.Font = $font
$grpStatus.Controls.Add($lblNote)

$grpPreview = New-Object System.Windows.Forms.GroupBox
$grpPreview.Text = "Generated Role Preview"
$grpPreview.Font = $fontBold
$grpPreview.Location = New-Object System.Drawing.Point(25, 320)
$grpPreview.Size = New-Object System.Drawing.Size(1110, 310)
$form.Controls.Add($grpPreview)

$previewTable = New-Object System.Windows.Forms.TableLayoutPanel
$previewTable.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewTable.ColumnCount = 5
$previewTable.RowCount = 2
$previewTable.Padding = New-Object System.Windows.Forms.Padding(10)
for ($i=0; $i -lt 5; $i++) { [void]$previewTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20))) }
[void]$previewTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
[void]$previewTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$grpPreview.Controls.Add($previewTable)

function New-PreviewLabel([string]$Text) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.Font = $fontBold
    $lbl.AutoSize = $true
    return $lbl
}

function New-PreviewTextbox {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.ScrollBars = "Vertical"
    $tb.ReadOnly = $true
    $tb.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tb.Font = $fontMono
    return $tb
}

$txtEntra = New-PreviewTextbox
$txtM365 = New-PreviewTextbox
$txtDlist = New-PreviewTextbox
$txtShared = New-PreviewTextbox
$txtDynamicSkipped = New-PreviewTextbox

$previewTable.Controls.Add((New-PreviewLabel "Entra Groups"), 0, 0)
$previewTable.Controls.Add((New-PreviewLabel "M365 Groups"), 1, 0)
$previewTable.Controls.Add((New-PreviewLabel "Distribution Groups"), 2, 0)
$previewTable.Controls.Add((New-PreviewLabel "Shared Mailboxes"), 3, 0)
$previewTable.Controls.Add((New-PreviewLabel "Dynamic Groups Skipped"), 4, 0)
$previewTable.Controls.Add($txtEntra, 0, 1)
$previewTable.Controls.Add($txtM365, 1, 1)
$previewTable.Controls.Add($txtDlist, 2, 1)
$previewTable.Controls.Add($txtShared, 3, 1)
$previewTable.Controls.Add($txtDynamicSkipped, 4, 1)

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = "Log Output"
$grpLog.Font = $fontBold
$grpLog.Location = New-Object System.Drawing.Point(25, 650)
$grpLog.Size = New-Object System.Drawing.Size(1110, 145)
$form.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Font = $fontMono
$grpLog.Controls.Add($txtLog)

function Refresh-Status {
    Test-RoleBuilderConnections
    Set-StatusLabel -Label $lblGraph -Connected $script:GraphConnected -ConnectedText "Microsoft Graph: Connected" -DisconnectedText "Microsoft Graph: Not connected"
    Set-StatusLabel -Label $lblExchange -Connected $script:ExchangeConnected -ConnectedText "Exchange Online: Connected" -DisconnectedText "Exchange Online: Not connected"
}

function Refresh-Preview {
    if ($null -eq $script:LastSnapshot) {
        Set-PreviewText -TextBox $txtEntra -Values @() -EmptyText ""
        Set-PreviewText -TextBox $txtM365 -Values @() -EmptyText ""
        Set-PreviewText -TextBox $txtDlist -Values @() -EmptyText ""
        Set-PreviewText -TextBox $txtShared -Values @() -EmptyText ""
        Set-PreviewText -TextBox $txtDynamicSkipped -Values @() -EmptyText ""
        $btnSave.Enabled = $false
        return
    }
    Set-PreviewText -TextBox $txtEntra -Values $script:LastSnapshot.EntraGroups -EmptyText "(No Entra Groups found)"
    Set-PreviewText -TextBox $txtM365 -Values $script:LastSnapshot.M365Groups -EmptyText "(No M365 Groups found)"
    Set-PreviewText -TextBox $txtDlist -Values $script:LastSnapshot.DistributionGroups -EmptyText "(No Distribution Groups found)"
    Set-PreviewText -TextBox $txtShared -Values $script:LastSnapshot.SharedMailboxes -EmptyText "(No Shared Mailboxes found)"
    Set-PreviewText -TextBox $txtDynamicSkipped -Values $script:LastSnapshot.DynamicGroupsSkipped -EmptyText "(No Dynamic Groups skipped)"
    $btnSave.Enabled = $true
}

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = "Choose role file"
    $dlg.Filter = "Role file (*.csv;*.xlsx;*.xlsm)|*.csv;*.xlsx;*.xlsm|CSV file (*.csv)|*.csv|Excel workbook (*.xlsx)|*.xlsx|Macro workbook (*.xlsm)|*.xlsm"
    $dlg.FileName = [System.IO.Path]::GetFileName($txtRolePath.Text)
    $initial = Split-Path $txtRolePath.Text -Parent
    if (Test-Path $initial) { $dlg.InitialDirectory = $initial }
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $txtRolePath.Text = $dlg.FileName }
})

$btnConnect.Add_Click({
    try {
        Connect-RoleBuilderServices -LogBox $txtLog
        Refresh-Status
    }
    catch {
        Write-Log $txtLog ("ERROR - Connect failed: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show("Connect failed:`r`n$($_.Exception.Message)", "Connection Error") | Out-Null
        Refresh-Status
    }
})

$btnBuild.Add_Click({
    try {
        Refresh-Status
        if (-not $script:GraphConnected -or -not $script:ExchangeConnected) { throw "Please connect services first." }
        if (Is-Blank $txtRoleName.Text) { throw "Role name is required." }
        if (Is-Blank $txtSourceUser.Text) { throw "Source user UPN is required." }
        $script:LastSnapshot = Build-RoleSnapshot -RoleName $txtRoleName.Text.Trim() -SourceUser $txtSourceUser.Text.Trim() -LogBox $txtLog
        Refresh-Preview
        [System.Windows.Forms.MessageBox]::Show("Role snapshot built. Review the preview, then click Save Role.", "Role Builder") | Out-Null
    }
    catch {
        Write-Log $txtLog ("ERROR - Build role failed: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show("Build role failed:`r`n$($_.Exception.Message)", "Build Role Error") | Out-Null
    }
})

$btnSave.Add_Click({
    try {
        if ($null -eq $script:LastSnapshot) { throw "No role snapshot has been built yet." }
        if (Is-Blank $txtRolePath.Text) { throw "Role file path is required." }
        Save-GeneratedRole -Snapshot $script:LastSnapshot -RolePath $txtRolePath.Text.Trim() -LogBox $txtLog
        [System.Windows.Forms.MessageBox]::Show("Role saved successfully.", "Role Builder") | Out-Null
    }
    catch {
        Write-Log $txtLog ("ERROR - Save role failed: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show("Save role failed:`r`n$($_.Exception.Message)", "Save Role Error") | Out-Null
    }
})

$statusTimer = New-Object System.Windows.Forms.Timer
$statusTimer.Interval = 5000
$statusTimer.Add_Tick({ Refresh-Status })
$statusTimer.Start()

Refresh-Status
Refresh-Preview
Write-Log $txtLog "Role Builder loaded. Click Connect Services first. Dynamic groups will be skipped and logged."
[void]$form.ShowDialog()
