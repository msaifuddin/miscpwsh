# Run: pwsh -NoLogo -STA -ExecutionPolicy Bypass -File .\gui.ps1

$ErrorActionPreference = 'Stop'

# Preload WinForms so we can show UI prompts early
Add-Type -AssemblyName System.Windows.Forms

# Require PowerShell 7+ (pwsh)
if (-not $PSVersionTable -or $PSVersionTable.PSVersion.Major -lt 7 -or $PSVersionTable.PSEdition -ne 'Core') {
    [System.Windows.Forms.MessageBox]::Show(
        "This tool requires PowerShell 7+ (pwsh).`r`n`r`nInstall from: https://aka.ms/pwsh",
        "PowerShell 7 Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    exit 1
}

# Ensure STA mode for Windows Forms
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    if ($PSCommandPath) {
        Start-Process pwsh -ArgumentList "-NoLogo -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`"" -WindowStyle Normal
        exit
    }
}

# Safe string
function S([object]$v) { if ($null -eq $v) { '' } else { [string]$v } }

# --- Splash: non-TopMost, roomy, owner-aware ---
$script:LoadingForm  = $null
$script:LoadingLabel = $null

function Show-Loading([string]$message, [System.Windows.Forms.IWin32Window]$Owner = $null) {
    if ($script:LoadingForm -and -not $script:LoadingForm.IsDisposed) {
        $script:LoadingLabel.Text = $message
        [System.Windows.Forms.Application]::DoEvents()
        return
    }

    $f = New-Object System.Windows.Forms.Form
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
    $f.ControlBox = $false
    $f.ShowInTaskbar = $false
    $f.TopMost = $false
    $f.StartPosition = if ($Owner) { [System.Windows.Forms.FormStartPosition]::CenterParent } else { [System.Windows.Forms.FormStartPosition]::CenterScreen }
    $f.Size = New-Object System.Drawing.Size(520,170)
    $f.Text = 'Working...'

    $tlp = New-Object System.Windows.Forms.TableLayoutPanel
    $tlp.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tlp.Padding = New-Object System.Windows.Forms.Padding(12)
    $tlp.ColumnCount = 1
    $tlp.RowCount = 2
    $null = $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $null = $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $message
    $lbl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lbl.AutoEllipsis = $true

    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $pb.MarqueeAnimationSpeed = 30
    $pb.Dock = [System.Windows.Forms.DockStyle]::Fill

    $tlp.Controls.Add($lbl, 0, 0)
    $tlp.Controls.Add($pb, 0, 1)
    $f.Controls.Add($tlp)

    $script:LoadingForm  = $f
    $script:LoadingLabel = $lbl

    if ($Owner) { $f.Show($Owner) } else { $f.Show() }
    [System.Windows.Forms.Application]::DoEvents()
}

function Hide-Loading {
    try {
        if ($script:LoadingForm -and -not $script:LoadingForm.IsDisposed) {
            $script:LoadingForm.Close()
            $script:LoadingForm.Dispose()
        }
    } catch {}
    $script:LoadingForm  = $null
    $script:LoadingLabel = $null
}

# WinForms + Drawing
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Optionally trust PSGallery (first run)
try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}

# Graph modules (install if missing) — show splash while doing it
$required = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Identity.Governance'
)

Show-Loading "Loading Microsoft Graph modules...`r`n(First run may take a minute)"
foreach ($m in $required) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        try {
            if ($script:LoadingLabel) { $script:LoadingLabel.Text = "Installing: $m`r`n(This can take a minute on first run)"; [System.Windows.Forms.Application]::DoEvents() }
            Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Hide-Loading
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to install ${m}: $($_.Exception.Message)",
                "Setup Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            exit 1
        }
    }
    if (-not (Get-Module -Name $m)) {
        if ($script:LoadingLabel) { $script:LoadingLabel.Text = "Importing: $m"; [System.Windows.Forms.Application]::DoEvents() }
        Import-Module $m -ErrorAction Stop
    }
}
Hide-Loading

# Graph config
$clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$graphScopes = @('User.Read.All','DeviceManagementManagedDevices.Read.All','RoleManagement.ReadWrite.Directory')

# State
$script:IsConnected = $false
$script:SignedInUPN = $null
$script:CurrentUserId = $null
$script:lastUsers = @()
$script:lastDevices = @()
$script:lastPIMEligible = @()
$script:lastPIMActive = @()

# ===== Graph helpers =====
function Connect-GraphInteractive {
    try {
        Connect-MgGraph -ClientId $clientId -Scopes $graphScopes -NoWelcome -ErrorAction Stop
        $me = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me' -ErrorAction Stop
        if (-not $me.id -or -not $me.userPrincipalName) { throw 'Signed in but /me returned no id/UPN.' }
        $script:SignedInUPN = $me.userPrincipalName
        $script:CurrentUserId = $me.id
        $script:IsConnected = $true
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Sign-in Failed',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $script:IsConnected = $false
        return $false
    }
}
function Disconnect-GraphSafe {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $script:IsConnected = $false
    $script:SignedInUPN = $null
    $script:CurrentUserId = $null
    $script:lastUsers = @()
    $script:lastDevices = @()
    $script:lastPIMEligible = @()
    $script:lastPIMActive = @()
}
function Search-UsersOnce([string]$Text) {
    $q = (S $Text).Trim()
    if ($q -eq '') { return @() }
    $escapedQ = $q.Replace("'", "''")
    $filter = "startswith(userPrincipalName,'$escapedQ') or startswith(displayName,'$escapedQ') or startswith(givenName,'$escapedQ') or startswith(surname,'$escapedQ')"
    try {
        $users = Get-MgUser -Filter $filter -All -ErrorAction Stop
        if ($users -and $users.Count -gt 0) { return $users }
        $headers = @{ 'ConsistencyLevel' = 'eventual' }
        $searchUri = "https://graph.microsoft.com/v1.0/users?`$search=`"$q`""
        (Invoke-MgGraphRequest -Method GET -Uri $searchUri -Headers $headers -ErrorAction Stop).value
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'User Search Error',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        @()
    }
}
function Get-User-ById([string]$UserId) {
    if ([string]::IsNullOrWhiteSpace($UserId)) { return $null }
    try {
        $select = 'id,displayName,userPrincipalName,jobTitle,mail,mobilePhone,businessPhones,officeLocation,companyName'
        $uri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($UserId))?`$select=$select"
        Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    } catch { $null }
}
function Get-User-Manager([string]$UserId) {
    if ([string]::IsNullOrWhiteSpace($UserId)) { return $null }
    try {
        $select = 'id,displayName,jobTitle,mail,userPrincipalName'
        $uri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($UserId))/manager?`$select=$select"
        Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    } catch { $null }
}
function Get-User-ManagedDevicesOnce([string]$UserId) {
    if ([string]::IsNullOrWhiteSpace($UserId)) { return @() }
    $uri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($UserId))/managedDevices"
    $all = @()
    try {
        do {
            $res = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($res.value) { $all += $res.value }
            $uri = $res.'@odata.nextLink'
        } while ($null -ne $uri)
    } catch { $all = @() }
    $all
}
function Search-DevicesOnce([string]$Text) {
    $q = (S $Text).Trim()
    if ($q -eq '') { return @() }
    try {
        $escapedQ = $q.Replace("'", "''")
        $bySerial = Get-MgDeviceManagementManagedDevice -All -Filter "serialNumber eq '$escapedQ'" -ErrorAction Stop
        if ($bySerial -and $bySerial.Count -gt 0) { return $bySerial }
        Get-MgDeviceManagementManagedDevice -All -Filter "startswith(deviceName,'$escapedQ')" -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Device Search Error',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        @()
    }
}
function Load-PIMOnce {
    try {
        $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "principalId eq '$($script:CurrentUserId)'" -ExpandProperty RoleDefinition -All -ErrorAction Stop
    } catch { $eligible = @() }
    try {
        $active = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter "principalId eq '$($script:CurrentUserId)'" -ExpandProperty RoleDefinition -All -ErrorAction Stop
    } catch { $active = @() }
    @($eligible, $active)
}
function PIM-ActivateOnce($Eligibility, [string]$Just, [int]$Hours, [string]$Tno = 'None', [string]$Tsys = 'JIRA') {
    try {
        $body = @{
            action           = 'selfActivate'
            principalId      = $script:CurrentUserId
            roleDefinitionId = $Eligibility.RoleDefinition.Id
            directoryScopeId = $Eligibility.DirectoryScopeId
            justification    = $Just
            scheduleInfo     = @{ startDateTime = (Get-Date).ToUniversalTime().ToString('o'); expiration = @{ type = 'AfterDuration'; duration = "PT${Hours}H" } }
            ticketInfo       = @{ ticketNumber = $Tno; ticketSystem = $Tsys }
        }
        New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body -ErrorAction Stop | Out-Null
        $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PIM Activation Error',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $false
    }
}
function PIM-DeactivateOnce($ActiveInst) {
    try {
        $body = @{
            action               = 'selfDeactivate'
            assignmentScheduleId = $ActiveInst.AssignmentScheduleId
            principalId          = $script:CurrentUserId
            roleDefinitionId     = $ActiveInst.RoleDefinition.Id
            directoryScopeId     = $ActiveInst.DirectoryScopeId
            justification        = 'User requested deactivation'
        }
        New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body -ErrorAction Stop | Out-Null
        $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'PIM Deactivation Error',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $false
    }
}

# ===== UI helpers =====
function Set-ListViewColumns([System.Windows.Forms.ListView]$lv, [int[]]$widths) {
    if (-not $lv -or $lv.Columns.Count -eq 0) { return }
    $lv.BeginUpdate()
    try {
        for ($i=0; $i -lt $widths.Count -and $i -lt $lv.Columns.Count; $i++) {
            if ($widths[$i] -ge 0) { $lv.Columns[$i].Width = [Math]::Max(24, $widths[$i]) }
        }
        if ($widths.Count -gt 0 -and $widths[$widths.Count-1] -lt 0) {
            $total = 0
            for ($i=0; $i -lt $lv.Columns.Count - 1; $i++) { $total += [Math]::Max(24, $lv.Columns[$i].Width) }
            $fill = [Math]::Max(60, $lv.ClientSize.Width - $total - 4)
            $lv.Columns[$lv.Columns.Count - 1].Width = $fill
        }
    } finally { $lv.EndUpdate() }
}
function Set-SplitterSafe([System.Windows.Forms.SplitContainer]$split, [int]$p1, [int]$p2, [double]$ratio) {
    $split.Panel1MinSize = $p1
    $split.Panel2MinSize = $p2
    $dim = if ($split.Orientation -eq [System.Windows.Forms.Orientation]::Horizontal) { $split.Height } else { $split.Width }
    if ($dim -le 0) { $dim = 1 }
    $min = $split.Panel1MinSize
    $max = $dim - $split.Panel2MinSize
    if ($max -lt $min) { $max = $min + 1 }
    $target = [int]($dim * $ratio)
    if ($target -lt $min) { $target = $min }
    if ($target -gt $max) { $target = $max }
    $split.SplitterDistance = $target
}

# ===== Main form =====
$form = New-Object System.Windows.Forms.Form
$form.Text = 'IT Support Helper v2.0'
$form.StartPosition = 'CenterScreen'
$form.Size        = New-Object System.Drawing.Size(1060, 680)
$form.MinimumSize = New-Object System.Drawing.Size(900, 560)
$form.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor   = [System.Drawing.Color]::FromArgb(240,240,240)

# Top bar
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = 'Top'
$topPanel.Height = 44
$topPanel.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
$topPanel.ForeColor = [System.Drawing.Color]::White

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Connect'
$btnConnect.Location = New-Object System.Drawing.Point(12, 8)
$btnConnect.Size = New-Object System.Drawing.Size(100, 28)
$btnConnect.FlatStyle = 'Flat'
$btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0,122,204)
$btnConnect.ForeColor = [System.Drawing.Color]::White
$btnConnect.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text = 'Disconnect'
$btnDisconnect.Location = New-Object System.Drawing.Point(120, 8)
$btnDisconnect.Size = New-Object System.Drawing.Size(100, 28)
$btnDisconnect.FlatStyle = 'Flat'
$btnDisconnect.BackColor = [System.Drawing.Color]::FromArgb(202,81,81)
$btnDisconnect.ForeColor = [System.Drawing.Color]::White
$btnDisconnect.Enabled = $false
$btnDisconnect.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Not connected'
$lblStatus.Location = New-Object System.Drawing.Point(240, 12)
$lblStatus.Size = New-Object System.Drawing.Size(700, 20)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,193,7)
$lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$topPanel.Controls.AddRange(@($btnConnect,$btnDisconnect,$lblStatus))

# Tabs
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock       = 'Fill'
$tabControl.Font       = New-Object System.Drawing.Font('Segoe UI', 9)
$tabControl.Alignment  = [System.Windows.Forms.TabAlignment]::Top
$tabControl.SizeMode   = [System.Windows.Forms.TabSizeMode]::Fixed
$tabControl.ItemSize   = New-Object System.Drawing.Size(120, 24)

$form.Controls.Add($tabControl)
$form.Controls.Add($topPanel)

# ===== Users tab =====
$tabUsers = New-Object System.Windows.Forms.TabPage
$tabUsers.Text = 'Users'
$tabUsers.Padding = [System.Windows.Forms.Padding]::Empty
$tabControl.TabPages.Add($tabUsers)

$usersMainSplit = New-Object System.Windows.Forms.SplitContainer
$usersMainSplit.Dock = 'Fill'
$usersMainSplit.Orientation = 'Vertical'
$tabUsers.Controls.Add($usersMainSplit)

$usersSearchPanel = New-Object System.Windows.Forms.Panel
$usersSearchPanel.Dock = 'Top'
$usersSearchPanel.Height = 40
$usersSearchPanel.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
$tabUsers.Controls.Add($usersSearchPanel)

$lblUserSearch = New-Object System.Windows.Forms.Label
$lblUserSearch.Text = 'Search Users:'
$lblUserSearch.Location = New-Object System.Drawing.Point(12, 11)
$lblUserSearch.Size = New-Object System.Drawing.Size(100, 20)

$txtUserSearch = New-Object System.Windows.Forms.TextBox
$txtUserSearch.Location = New-Object System.Drawing.Point(110, 9)
$txtUserSearch.Size = New-Object System.Drawing.Size(360, 22)

$btnUserSearch = New-Object System.Windows.Forms.Button
$btnUserSearch.Text = 'Search'
$btnUserSearch.Location = New-Object System.Drawing.Point(476, 8)
$btnUserSearch.Size = New-Object System.Drawing.Size(90, 24)
$btnUserSearch.FlatStyle = 'Flat'
$btnUserSearch.BackColor = [System.Drawing.Color]::FromArgb(0,122,204)
$btnUserSearch.ForeColor = [System.Drawing.Color]::White

$lblUserResults = New-Object System.Windows.Forms.Label
$lblUserResults.Text = 'Total: 0 results'
$lblUserResults.Location = New-Object System.Drawing.Point(572, 11)
$lblUserResults.AutoSize = $true

$usersSearchPanel.Controls.AddRange(@($lblUserSearch,$txtUserSearch,$btnUserSearch,$lblUserResults))

$lvUsers = New-Object System.Windows.Forms.ListView
$lvUsers.Dock = 'Fill'
$lvUsers.View = 'Details'
$lvUsers.FullRowSelect = $true
$lvUsers.GridLines = $true
$lvUsers.HideSelection = $false
[void]$lvUsers.Columns.Add('No.', 56)
[void]$lvUsers.Columns.Add('Display Name', 240)
[void]$lvUsers.Columns.Add('User Principal Name', 300)
$usersMainSplit.Panel1.Controls.Add($lvUsers)

$usersRightSplit = New-Object System.Windows.Forms.SplitContainer
$usersRightSplit.Dock = 'Fill'
$usersRightSplit.Orientation = 'Horizontal'
$usersMainSplit.Panel2.Controls.Add($usersRightSplit)

$grpUserDetails = New-Object System.Windows.Forms.GroupBox
$grpUserDetails.Text = 'User Details'
$grpUserDetails.Dock = 'Fill'
$grpUserDetails.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grpUserDetails.ForeColor = [System.Drawing.Color]::FromArgb(0,122,204)
$usersRightSplit.Panel1.Controls.Add($grpUserDetails)

$txtUserDetails = New-Object System.Windows.Forms.TextBox
$txtUserDetails.Dock = 'Fill'
$txtUserDetails.Multiline = $true
$txtUserDetails.ReadOnly = $true
$txtUserDetails.ScrollBars = 'Vertical'
$txtUserDetails.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtUserDetails.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
$grpUserDetails.Controls.Add($txtUserDetails)

$grpUserDevices = New-Object System.Windows.Forms.GroupBox
$grpUserDevices.Text = 'Managed Devices'
$grpUserDevices.Dock = 'Fill'
$grpUserDevices.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grpUserDevices.ForeColor = [System.Drawing.Color]::FromArgb(0,122,204)
$usersRightSplit.Panel2.Controls.Add($grpUserDevices)

$txtUserDevices = New-Object System.Windows.Forms.TextBox
$txtUserDevices.Dock = 'Fill'
$txtUserDevices.Multiline = $true
$txtUserDevices.ReadOnly = $true
$txtUserDevices.ScrollBars = 'Vertical'
$txtUserDevices.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtUserDevices.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
$grpUserDevices.Controls.Add($txtUserDevices)

# ===== Devices tab =====
$tabDevices = New-Object System.Windows.Forms.TabPage
$tabDevices.Text = 'Devices'
$tabDevices.Padding = [System.Windows.Forms.Padding]::Empty
$tabControl.TabPages.Add($tabDevices)

$devicesMainSplit = New-Object System.Windows.Forms.SplitContainer
$devicesMainSplit.Dock = 'Fill'
$devicesMainSplit.Orientation = 'Vertical'
$tabDevices.Controls.Add($devicesMainSplit)

$devicesSearchPanel = New-Object System.Windows.Forms.Panel
$devicesSearchPanel.Dock = 'Top'
$devicesSearchPanel.Height = 40
$devicesSearchPanel.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
$tabDevices.Controls.Add($devicesSearchPanel)

$lblDeviceSearch = New-Object System.Windows.Forms.Label
$lblDeviceSearch.Text = 'Search Devices:'
$lblDeviceSearch.Location = New-Object System.Drawing.Point(12, 11)
$lblDeviceSearch.Size = New-Object System.Drawing.Size(120, 20)

$txtDeviceSearch = New-Object System.Windows.Forms.TextBox
$txtDeviceSearch.Location = New-Object System.Drawing.Point(130, 9)
$txtDeviceSearch.Size = New-Object System.Drawing.Size(360, 22)

$btnDeviceSearch = New-Object System.Windows.Forms.Button
$btnDeviceSearch.Text = 'Search'
$btnDeviceSearch.Location = New-Object System.Drawing.Point(496, 8)
$btnDeviceSearch.Size = New-Object System.Drawing.Size(90, 24)
$btnDeviceSearch.FlatStyle = 'Flat'
$btnDeviceSearch.BackColor = [System.Drawing.Color]::FromArgb(0,122,204)
$btnDeviceSearch.ForeColor = [System.Drawing.Color]::White

$lblDeviceResults = New-Object System.Windows.Forms.Label
$lblDeviceResults.Text = 'Total: 0 results'
$lblDeviceResults.Location = New-Object System.Drawing.Point(600, 11)
$lblDeviceResults.AutoSize = $true

$devicesSearchPanel.Controls.AddRange(@($lblDeviceSearch,$txtDeviceSearch,$btnDeviceSearch,$lblDeviceResults))

$lvDevices = New-Object System.Windows.Forms.ListView
$lvDevices.Dock = 'Fill'
$lvDevices.View = 'Details'
$lvDevices.FullRowSelect = $true
$lvDevices.GridLines = $true
$lvDevices.HideSelection = $false
[void]$lvDevices.Columns.Add('No.', 56)
[void]$lvDevices.Columns.Add('Device Name', 280)
[void]$lvDevices.Columns.Add('Serial Number', 200)
$devicesMainSplit.Panel1.Controls.Add($lvDevices)

$devicesRightSplit = New-Object System.Windows.Forms.SplitContainer
$devicesRightSplit.Dock = 'Fill'
$devicesRightSplit.Orientation = 'Horizontal'
$devicesMainSplit.Panel2.Controls.Add($devicesRightSplit)

$grpDeviceDetails = New-Object System.Windows.Forms.GroupBox
$grpDeviceDetails.Text = 'Device Details'
$grpDeviceDetails.Dock = 'Fill'
$grpDeviceDetails.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grpDeviceDetails.ForeColor = [System.Drawing.Color]::FromArgb(0,122,204)
$devicesRightSplit.Panel1.Controls.Add($grpDeviceDetails)

$txtDeviceDetails = New-Object System.Windows.Forms.TextBox
$txtDeviceDetails.Dock = 'Fill'
$txtDeviceDetails.Multiline = $true
$txtDeviceDetails.ReadOnly = $true
$txtDeviceDetails.ScrollBars = 'Vertical'
$txtDeviceDetails.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtDeviceDetails.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
$grpDeviceDetails.Controls.Add($txtDeviceDetails)

$grpDeviceUser = New-Object System.Windows.Forms.GroupBox
$grpDeviceUser.Text = 'Associated User'
$grpDeviceUser.Dock = 'Fill'
$grpDeviceUser.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grpDeviceUser.ForeColor = [System.Drawing.Color]::FromArgb(0,122,204)
$devicesRightSplit.Panel2.Controls.Add($grpDeviceUser)

$txtDeviceUser = New-Object System.Windows.Forms.TextBox
$txtDeviceUser.Dock = 'Fill'
$txtDeviceUser.Multiline = $true
$txtDeviceUser.ReadOnly = $true
$txtDeviceUser.ScrollBars = 'Vertical'
$txtDeviceUser.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtDeviceUser.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
$grpDeviceUser.Controls.Add($txtDeviceUser)

# ===== PIM tab =====
$tabPIM = New-Object System.Windows.Forms.TabPage
$tabPIM.Text = 'PIM'
$tabPIM.Padding = [System.Windows.Forms.Padding]::Empty
$tabControl.TabPages.Add($tabPIM)

$pimMainSplit = New-Object System.Windows.Forms.SplitContainer
$pimMainSplit.Dock = 'Fill'
$pimMainSplit.Orientation = 'Vertical'
$tabPIM.Controls.Add($pimMainSplit)

$pimTopPanel = New-Object System.Windows.Forms.Panel
$pimTopPanel.Dock = 'Top'
$pimTopPanel.Height = 40
$pimTopPanel.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
$tabPIM.Controls.Add($pimTopPanel)

$btnLoadPIM = New-Object System.Windows.Forms.Button
$btnLoadPIM.Text = 'Load Eligible/Active Roles'
$btnLoadPIM.Location = New-Object System.Drawing.Point(12, 8)
$btnLoadPIM.Size = New-Object System.Drawing.Size(210, 24)
$btnLoadPIM.FlatStyle = 'Flat'
$btnLoadPIM.BackColor = [System.Drawing.Color]::FromArgb(0,122,204)
$btnLoadPIM.ForeColor = [System.Drawing.Color]::White
$btnLoadPIM.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$pimTopPanel.Controls.Add($btnLoadPIM)

$lblPIMCounts = New-Object System.Windows.Forms.Label
$lblPIMCounts.Text = '0 Eligible | 0 Activated'
$lblPIMCounts.Location = New-Object System.Drawing.Point(230, 11)
$lblPIMCounts.AutoSize = $true
$pimTopPanel.Controls.Add($lblPIMCounts)

$lvPIM = New-Object System.Windows.Forms.ListView
$lvPIM.Dock = 'Fill'
$lvPIM.View = 'Details'
$lvPIM.FullRowSelect = $true
$lvPIM.GridLines = $true
$lvPIM.HideSelection = $false
$lvPIM.CheckBoxes = $true
[void]$lvPIM.Columns.Add('No.', 56)
[void]$lvPIM.Columns.Add('Role Name', 240)
[void]$lvPIM.Columns.Add('Status', 100)
[void]$lvPIM.Columns.Add('Role ID', 320)
$pimMainSplit.Panel1.Controls.Add($lvPIM)

$pimRightPanel = New-Object System.Windows.Forms.Panel
$pimRightPanel.Dock = 'Fill'
$pimRightPanel.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
$pimMainSplit.Panel2.Controls.Add($pimRightPanel)

$grpPIMControls = New-Object System.Windows.Forms.GroupBox
$grpPIMControls.Text = 'Activation Controls'
$grpPIMControls.Dock = 'Top'
$grpPIMControls.Padding = New-Object System.Windows.Forms.Padding(10)
$grpPIMControls.AutoSize = $true
$grpPIMControls.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$pimRightPanel.Controls.Add($grpPIMControls)

$tlp = New-Object System.Windows.Forms.TableLayoutPanel
$tlp.Dock = 'Top'
$tlp.AutoSize = $true
$tlp.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$tlp.ColumnCount = 2
$tlp.RowCount = 4
$null = $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,130)))
$null = $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
$grpPIMControls.Controls.Add($tlp)

$lblJustification = New-Object System.Windows.Forms.Label
$lblJustification.Text = 'Justification:'
$lblJustification.Dock = 'Fill'
$tlp.Controls.Add($lblJustification, 0, 0)

$txtJustification = New-Object System.Windows.Forms.TextBox
$txtJustification.Dock = 'Fill'
$tlp.Controls.Add($txtJustification, 1, 0)

$lblDuration = New-Object System.Windows.Forms.Label
$lblDuration.Text = 'Duration (hours):'
$lblDuration.Dock = 'Fill'
$tlp.Controls.Add($lblDuration, 0, 1)

$numDuration = New-Object System.Windows.Forms.NumericUpDown
$numDuration.Minimum = 1
$numDuration.Maximum = 8
$numDuration.Value   = 4
$numDuration.Dock    = 'Left'
$numDuration.Width   = 80
$tlp.Controls.Add($numDuration, 1, 1)

$lblTicketNumber = New-Object System.Windows.Forms.Label
$lblTicketNumber.Text = 'Ticket Number:'
$lblTicketNumber.Dock = 'Fill'
$tlp.Controls.Add($lblTicketNumber, 0, 2)

$txtTicketNumber = New-Object System.Windows.Forms.TextBox
$txtTicketNumber.Dock = 'Fill'
$tlp.Controls.Add($txtTicketNumber, 1, 2)

$lblTicketSystem = New-Object System.Windows.Forms.Label
$lblTicketSystem.Text = 'Ticket System:'
$lblTicketSystem.Dock = 'Fill'
$tlp.Controls.Add($lblTicketSystem, 0, 3)

$txtTicketSystem = New-Object System.Windows.Forms.TextBox
$txtTicketSystem.Text = 'JIRA'
$txtTicketSystem.Dock = 'Fill'
$tlp.Controls.Add($txtTicketSystem, 1, 3)

$btnRow = New-Object System.Windows.Forms.FlowLayoutPanel
$btnRow.FlowDirection = 'LeftToRight'
$btnRow.Dock = 'Top'
$btnRow.AutoSize = $true
$btnRow.Padding = New-Object System.Windows.Forms.Padding(0,8,0,0)
$grpPIMControls.Controls.Add($btnRow)

$btnActivate = New-Object System.Windows.Forms.Button
$btnActivate.Text = 'Activate Selected Roles'
$btnActivate.AutoSize = $true
$btnActivate.FlatStyle = 'Flat'
$btnActivate.BackColor = [System.Drawing.Color]::FromArgb(40,167,69)
$btnActivate.ForeColor = [System.Drawing.Color]::White
$btnRow.Controls.Add($btnActivate)

$btnDeactivate = New-Object System.Windows.Forms.Button
$btnDeactivate.Text = 'Deactivate Selected Roles'
$btnDeactivate.AutoSize = $true
$btnDeactivate.FlatStyle = 'Flat'
$btnDeactivate.BackColor = [System.Drawing.Color]::FromArgb(220,53,69)
$btnDeactivate.ForeColor = [System.Drawing.Color]::White
$btnRow.Controls.Add($btnDeactivate)

# ===== UI logic =====
function Update-ConnectionStatus {
    if ($script:IsConnected) {
        $lblStatus.Text = "Connected as: $($script:SignedInUPN)"
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(40,167,69)
        $btnConnect.Enabled = $false
        $btnDisconnect.Enabled = $true
    } else {
        $lblStatus.Text = 'Not connected'
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,193,7)
        $btnConnect.Enabled = $true
        $btnDisconnect.Enabled = $false
        $lvUsers.Items.Clear();   $txtUserDetails.Clear(); $txtUserDevices.Clear()
        $lvDevices.Items.Clear(); $txtDeviceDetails.Clear(); $txtDeviceUser.Clear()
        $lvPIM.Items.Clear()
        $lblUserResults.Text = 'Total: 0 results'
        $lblDeviceResults.Text = 'Total: 0 results'
        $lblPIMCounts.Text = '0 Eligible | 0 Activated'
    }
}

# Events: connect/disconnect (with splash)
$btnConnect.Add_Click({
    if (-not $script:IsConnected) {
        $lblStatus.Text = "Connecting to Microsoft Graph..."
        $form.UseWaitCursor = $true
        try {
            if (Connect-GraphInteractive) { Update-ConnectionStatus }
        }
        finally {
            $form.UseWaitCursor = $false
            $lblStatus.Text = if ($script:IsConnected) { "Connected as: $($script:SignedInUPN)" } else { "Not connected" }
        }
    }
})

$btnDisconnect.Add_Click({
    if ($script:IsConnected) {
        $result = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to disconnect?','Confirm Disconnect',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { Disconnect-GraphSafe; Update-ConnectionStatus }
    }
})

# Users search
$btnUserSearch.Add_Click({
    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show('Please connect to Microsoft Graph first.','Not Connected',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
    }
    $q = $txtUserSearch.Text.Trim()
    if ($q -eq '') {
        [System.Windows.Forms.MessageBox]::Show('Please enter a search term.','Search Users',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $script:lastUsers = Search-UsersOnce -Text $q
        $lvUsers.BeginUpdate(); $lvUsers.Items.Clear()
        $i = 1
        foreach ($u in $script:lastUsers) {
            $it = New-Object System.Windows.Forms.ListViewItem([string]$i)
            [void]$it.SubItems.Add((S $u.displayName))
            [void]$it.SubItems.Add((S $u.userPrincipalName))
            $it.Tag = $u
            [void]$lvUsers.Items.Add($it); $i++
        }
        $lvUsers.EndUpdate()
        $lblUserResults.Text = "Total: $($script:lastUsers.Count) results"
        Set-ListViewColumns $lvUsers @(56,240,-1)
        if ($script:lastUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No users found matching your search.','Search Results',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$txtUserSearch.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $btnUserSearch.PerformClick() } })

# Users selection -> details + devices (with Manager)
$lvUsers.Add_SelectedIndexChanged({
    if ($lvUsers.SelectedItems.Count -lt 1) { $txtUserDetails.Clear(); $txtUserDevices.Clear(); return }
    $sel = $lvUsers.SelectedItems[0].Tag
    if (-not $sel) { return }
    $user = Get-User-ById -UserId $sel.id; if (-not $user) { $user = $sel }
    $mgr = Get-User-Manager -UserId $sel.id
    $businessPhones = ''; try { if ($user.businessPhones) { $businessPhones = ($user.businessPhones -join ', ') } } catch {}
    $lines = @(
        "Display Name       : $(S $user.displayName)"
        "User Principal Name: $(S $user.userPrincipalName)"
        "Job Title          : $(S $user.jobTitle)"
        "Email              : $(S $user.mail)"
        "Mobile Phone       : $(S $user.mobilePhone)"
        "Business Phones    : $businessPhones"
        "Office Location    : $(S $user.officeLocation)"
        "Company Name       : $(S $user.companyName)"
    )
    if ($mgr) {
        $lines += @(
            "Manager            : $(S $mgr.displayName)"
            "Manager's Title    : $(S $mgr.jobTitle)"
            "Manager's Email    : $(S $mgr.mail)"
        )
    }
    $txtUserDetails.Text = $lines -join [Environment]::NewLine

    $txtUserDevices.Text = 'Loading devices...'
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $ds = Get-User-ManagedDevicesOnce -UserId $sel.id
        if (-not $ds -or $ds.Count -eq 0) {
            $txtUserDevices.Text = 'No managed devices found for this user.'
        } else {
            $txtUserDevices.Text = ($ds | ForEach-Object {
                @(
                    "Device Name     : $(S $_.deviceName)"
                    "Model           : $(S $_.model)"
                    "Serial Number   : $(S $_.serialNumber)"
                    "Operating System: $(S $_.operatingSystem)"
                    "Compliance State: $(S $_.complianceState)"
                    "Last Sync       : $(S $_.lastSyncDateTime)"
                    ('-' * 50)
                ) -join [Environment]::NewLine
            }) -join [Environment]::NewLine
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# Devices search
$btnDeviceSearch.Add_Click({
    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show('Please connect to Microsoft Graph first.','Not Connected',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
    }
    $q = $txtDeviceSearch.Text.Trim()
    if ($q -eq '') {
        [System.Windows.Forms.MessageBox]::Show('Please enter a device name or serial number.','Search Devices',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $script:lastDevices = Search-DevicesOnce -Text $q
        $lvDevices.BeginUpdate(); $lvDevices.Items.Clear()
        $i = 1
        foreach ($d in $script:lastDevices) {
            $it = New-Object System.Windows.Forms.ListViewItem([string]$i)
            [void]$it.SubItems.Add((S $d.deviceName))
            [void]$it.SubItems.Add((S $d.serialNumber))
            $it.Tag = $d
            [void]$lvDevices.Items.Add($it); $i++
        }
        $lvDevices.EndUpdate()
        $lblDeviceResults.Text = "Total: $($script:lastDevices.Count) results"
        Set-ListViewColumns $lvDevices @(56,280,-1)
        if ($script:lastDevices.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No devices found matching your search.','Search Results',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$txtDeviceSearch.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $btnDeviceSearch.PerformClick() } })

# Device selection -> details + associated user
$lvDevices.Add_SelectedIndexChanged({
    if ($lvDevices.SelectedItems.Count -lt 1) { $txtDeviceDetails.Clear(); $txtDeviceUser.Clear(); return }
    $device = $lvDevices.SelectedItems[0].Tag
    if (-not $device) { return }
    $ownerType = ''; try { $ownerType = S $device.managedDeviceOwnerType } catch {}
    $txtDeviceDetails.Text = @(
        "Device Name      : $(S $device.deviceName)"
        "Model            : $(S $device.model)"
        "Manufacturer     : $(S $device.manufacturer)"
        "Serial Number    : $(S $device.serialNumber)"
        "Operating System : $(S $device.operatingSystem)"
        "Owner Type       : $ownerType"
        "Compliance State : $(S $device.complianceState)"
        "Enrollment Date  : $(S $device.enrolledDateTime)"
        "Last Sync        : $(S $device.lastSyncDateTime)"
    ) -join [Environment]::NewLine

    $txtDeviceUser.Text = 'Resolving user...'
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $u = $null
        if ($device.userId) { $u = Get-User-ById -UserId $device.userId }
        if ($u) {
            $txtDeviceUser.Text = @(
                "User Principal Name: $(S $u.userPrincipalName)"
                "Display Name       : $(S $u.displayName)"
                "Job Title          : $(S $u.jobTitle)"
            ) -join [Environment]::NewLine
        } else {
            $txtDeviceUser.Text = 'No associated user found for this device.'
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# PIM load
$btnLoadPIM.Add_Click({
    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show('Please connect to Microsoft Graph first.','Not Connected',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $eligible, $active = Load-PIMOnce
        $script:lastPIMEligible = $eligible
        $script:lastPIMActive = $active

        $eligibleList = ($eligible + @())
        $activeIds = @(); if ($active) { $activeIds = $active | ForEach-Object { $_.RoleDefinition.Id } }

        $lvPIM.BeginUpdate(); $lvPIM.Items.Clear()
        $i = 1
        foreach ($role in $eligibleList) {
            $isActive = $activeIds -contains $role.RoleDefinition.Id
            $status = if ($isActive) { 'Active' } else { 'Eligible' }
            $it = New-Object System.Windows.Forms.ListViewItem([string]$i)
            [void]$it.SubItems.Add((S $role.RoleDefinition.DisplayName))
            [void]$it.SubItems.Add($status)
            [void]$it.SubItems.Add((S $role.RoleDefinition.Id))
            $it.Tag = @{ eligible = $role; isActive = $isActive }
            $it.Checked = $false
            if ($isActive) { $it.BackColor = [System.Drawing.Color]::FromArgb(232,245,233) }
            [void]$lvPIM.Items.Add($it); $i++
        }
        $lvPIM.EndUpdate()

        $eligibleCount = $eligibleList.Count
        $activeCount   = ($active + @()).Count
        $lblPIMCounts.Text = "$eligibleCount Eligible | $activeCount Activated"

        # Keep No. visible even with checkboxes
        Set-ListViewColumns $lvPIM @(56,240,100,-1)

        if ($eligibleCount -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No PIM eligible roles found for your account.','PIM Roles',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# PIM activate/deactivate
$btnActivate.Add_Click({
    $checkedItems = @(); foreach ($item in $lvPIM.Items) { if ($item.Checked) { $checkedItems += $item } }
    if ($checkedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Please select at least one eligible role to activate.','PIM Activation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
    }
    $just = $txtJustification.Text.Trim()
    if ($just -eq '') {
        [System.Windows.Forms.MessageBox]::Show('Justification is required for role activation.','PIM Activation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null; return
    }
    $hours = [int]$numDuration.Value
    $tno = $txtTicketNumber.Text.Trim(); if ($tno -eq '') { $tno = 'None' }
    $tsys = $txtTicketSystem.Text.Trim(); if ($tsys -eq '') { $tsys = 'JIRA' }
    $activeIds = @(); if ($script:lastPIMActive) { $activeIds = $script:lastPIMActive | ForEach-Object { $_.RoleDefinition.Id } }
    $ok = 0; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        foreach ($it in $checkedItems) {
            $elig = $it.Tag.eligible
            if ($activeIds -contains $elig.RoleDefinition.Id) { continue }
            if (PIM-ActivateOnce -Eligibility $elig -Just $just -Hours $hours -Tno $tno -Tsys $tsys) { $ok++ }
        }
        if ($ok -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("$ok role activation request(s) submitted successfully.",'PIM Activation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $btnLoadPIM.PerformClick()
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})
$btnDeactivate.Add_Click({
    $checkedItems = @(); foreach ($item in $lvPIM.Items) { if ($item.Checked -and $item.Tag.isActive) { $checkedItems += $item } }
    if ($checkedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Please select at least one active role to deactivate.','PIM Deactivation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
    }
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to deactivate $($checkedItems.Count) role(s)?",'Confirm Deactivation',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $activeMap = @{}; foreach ($a in ($script:lastPIMActive + @())) { $activeMap[$a.RoleDefinition.Id] = $a }
    $ok = 0; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        foreach ($it in $checkedItems) {
            $elig = $it.Tag.eligible
            $inst = $activeMap[$elig.RoleDefinition.Id]
            if ($inst -and (PIM-DeactivateOnce -ActiveInst $inst)) { $ok++ }
        }
        if ($ok -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("$ok role deactivation request(s) submitted successfully.",'PIM Deactivation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $btnLoadPIM.PerformClick()
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# Form init
$form.Add_Shown({
    Set-SplitterSafe $usersMainSplit     250 260 0.42
    Set-SplitterSafe $usersRightSplit    140 140 0.52
    Set-SplitterSafe $devicesMainSplit   300 260 0.52
    Set-SplitterSafe $devicesRightSplit  130 130 0.52
    Set-SplitterSafe $pimMainSplit       320 320 0.50

    Update-ConnectionStatus
    $tabControl.SelectedIndex = 0

    # Initial column widths so index is visible straight away
    Set-ListViewColumns $lvUsers   @(56,240,-1)
    Set-ListViewColumns $lvDevices @(56,280,-1)
    Set-ListViewColumns $lvPIM     @(56,240,100,-1)
})

# Show
[void]$form.ShowDialog()
