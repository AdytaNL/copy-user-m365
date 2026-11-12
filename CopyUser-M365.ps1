Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# CopyUser-M365.ps1
# Copy an M365 user
#
# Architecture:
#   Everything runs synchronously on the main thread.
#   Graph AND Exchange auth are both bound to the runspace you signed in from.
#   Wait-ForUserInExchange uses short sleeps + DoEvents so the UI stays responsive.
#
# Main behaviour:
#   - Shared mailbox permissions are only set once the new user has a real UserMailbox,
#     so Exchange records the AutoMapping link.
#   - Optional license transfer from a source who has left (confirmation with an overview first;
#     licenses go back to the source if assigning fails).
#   - Optional Teams phone number transfer, run last, with rollback where possible.
#   - Waits stop early when the new user's license is in an error state.
# ============================================================

$FixedLogName = 'CopyUser-M365.log'

# Exchange provisioning wait times (seconds)
$UserWaitTimeoutSeconds     = 180   # Exchange recipient visible (for distribution groups)
$MailboxWaitTimeoutSeconds  = 900   # UserMailbox created after license (for mailbox permissions + AutoMapping)
$MailboxWaitIntervalSeconds = 15

# License transfer: after removal from the source it can take a while before the
# seat is free. Assignment to the new user is therefore attempted a few times.
$LicenseTransferRetryCount   = 6
$LicenseTransferRetrySeconds = 10
$SharedMailboxWaitSeconds    = 60    # Source mailbox visible as SharedMailbox after conversion
$GroupReleaseWaitSeconds     = 600   # source out of license group after blocking/removal
$LicenseReprocessRetryCount  = 6     # reprocess licenses of the new user (30s in between)

# Phone number transfer: wait until the new user has an active Teams Phone
# in Teams (license via group + Teams provisioning can take a while).
$TeamsUserWaitTimeoutSeconds  = 1800
$TeamsUserWaitIntervalSeconds = 30
$NumberAssignRetryCount       = 10
$NumberAssignRetrySeconds     = 30
$VoiceCheckTimeoutSeconds     = 120   # check EnterpriseVoiceEnabled after assignment
$LicenseVerifySeconds         = 300   # check whether the new user really has a released group license active
$LicenseBlockerGraceSeconds   = 180   # only after this time may a wait loop stop on a license in an error state
$GraphUserWaitSeconds         = 60    # new user findable in Graph before groups/licenses are set

$IsSystem         = [System.Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
$CurrentUserLabel = if ($IsSystem) { 'SYSTEM' } else { $env:USERNAME }

$LogDir = if ($IsSystem) {
    Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
} else {
    Join-Path $env:LOCALAPPDATA 'Microsoft\IntuneManagementExtension\Logs'
}

if ([string]::IsNullOrWhiteSpace($LogDir)) {
    $LogDir = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'
}

$LogFile = Join-Path $LogDir $FixedLogName

$script:Form                 = $null
$script:LogBox               = $null
$script:StatusLabel          = $null
$script:ControlsToToggle     = @()
$script:SelectedSourceUser   = $null
$script:SourceUserResults    = [System.Collections.Generic.List[object]]::new()
# Group results of the last copy run (memberships added through Graph and Exchange).
$script:CopiedGroupResults   = [System.Collections.Generic.List[object]]::new()
$script:GraphConnectedMain   = $false
$script:GraphTenantId        = $null
$script:GraphModulesLoaded  = $false   # Idempotency flag: Graph modules are loaded only once
$script:ExchangeModuleLoaded = $false  # Idempotency flag: Exchange module is loaded only once
$script:TeamsModuleLoaded    = $false  # Idempotency flag: Teams module is loaded only once
$script:UpdatingSuggestions    = $false
$script:AutoUpdateDisplayName  = $true
$script:AutoUpdateUpn          = $true
$script:PasswordVisible        = $false
$script:IsBusy                 = $false
$script:LastWaitBlocker        = $null
$script:SkuServicePlans        = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR')][string]$Level = 'INFO'
    )

    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Line      = "$Timestamp [$Level] [$CurrentUserLabel] $Message"

    # Logging must never make an action fail (e.g. log file temporarily locked by AV).
    try {
        if (-not (Test-Path -LiteralPath $LogDir)) {
            New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
        }
        Add-Content -LiteralPath $LogFile -Value $Line -Encoding UTF8 -ErrorAction Stop
    } catch {
    }

    try { if ($Host.UI -and $Host.UI.RawUI) { Write-Host $Line } } catch { }

    try {
        if ($script:LogBox -and -not $script:LogBox.IsDisposed -and -not $script:LogBox.Disposing -and $script:LogBox.FindForm()) {
            $appendAction = {
                param($TextBox, $Text)
                if ($TextBox -and -not $TextBox.IsDisposed -and -not $TextBox.Disposing) {
                    $TextBox.AppendText($Text + [Environment]::NewLine)
                    $TextBox.SelectionStart = $TextBox.TextLength
                    $TextBox.ScrollToCaret()
                }
            }

            if ($script:LogBox.InvokeRequired) {
                $null = $script:LogBox.BeginInvoke($appendAction, @($script:LogBox, $Line))
            } else {
                & $appendAction $script:LogBox $Line
            }
        }
    } catch {
    }
}

function Log-Info    { param([Parameter(Mandatory)][string]$Message); Write-Log -Message $Message -Level 'INFO' }
function Log-Warning { param([Parameter(Mandatory)][string]$Message); Write-Log -Message $Message -Level 'WARNING' }
function Log-Error   { param([Parameter(Mandatory)][string]$Message); Write-Log -Message $Message -Level 'ERROR' }

function Set-UiBusyState {
    param(
        [Parameter(Mandatory)][bool]$Busy,
        [string]$StatusText
    )

    $script:IsBusy = $Busy

    foreach ($control in $script:ControlsToToggle) {
        if ($null -ne $control -and -not $control.IsDisposed) {
            $control.Enabled = -not $Busy
        }
    }

    if ($script:StatusLabel -and -not $script:StatusLabel.IsDisposed) {
        $script:StatusLabel.Text = if ($StatusText) { $StatusText } else { '' }
    }

    if ($script:Form -and -not $script:Form.IsDisposed) {
        $script:Form.UseWaitCursor = $Busy
        [System.Windows.Forms.Application]::UseWaitCursor = $Busy
    }
}

function Show-UiError {
    param([Parameter(Mandatory)][string]$Message)
    Log-Error $Message
    [System.Windows.Forms.MessageBox]::Show(
        $Message, 'Error',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Show-UiInfo {
    param([Parameter(Mandatory)][string]$Message)
    Log-Info $Message
    [System.Windows.Forms.MessageBox]::Show(
        $Message, 'Information',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Show-UiWarning {
    param([Parameter(Mandatory)][string]$Message)
    Log-Warning $Message
    [System.Windows.Forms.MessageBox]::Show(
        $Message, 'Attention',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

function Show-UiQuestion {
    # Returns the chosen DialogResult as text: 'Yes', 'No' or 'Cancel'.
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('YesNo','YesNoCancel')][string]$Buttons = 'YesNo'
    )
    $Answer = [System.Windows.Forms.MessageBox]::Show(
        $Message, 'Confirm',
        [System.Windows.Forms.MessageBoxButtons]$Buttons,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    return [string]$Answer
}

function Wait-UiSeconds {
    # Wait without freezing the UI.
    param([Parameter(Mandatory)][int]$Seconds)
    $End = [DateTime]::UtcNow.AddSeconds($Seconds)
    while ([DateTime]::UtcNow -lt $End) {
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Escape-ODataString {
    param([Parameter(Mandatory)][string]$Value)
    return $Value.Replace("'", "''")
}

function Get-MailNickname {
    param(
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName
    )

    $Nickname = ($FirstName.Trim() + $LastName.Trim()) -replace '[^a-zA-Z0-9._-]', ''
    if ([string]::IsNullOrWhiteSpace($Nickname)) {
        throw 'MailNickname cannot be empty after cleaning up the input.'
    }

    return $Nickname.ToLower()
}


function New-RandomPassword {
    param(
        [ValidateRange(12, 128)][int]$Length = 18
    )

    # Deliberately no confusing characters such as O/0 and I/l/1.
    # Quotes, backticks and backslashes are also avoided so the password
    # can easily be copied into admin consoles and documentation.
    $Upper   = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $Lower   = 'abcdefghijkmnopqrstuvwxyz'
    $Digits  = '23456789'
    $Special = '!@$%*-_+=?'
    $All     = $Upper + $Lower + $Digits + $Special

    $Rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()

    try {
        $GetIndex = {
            param([Parameter(Mandatory)][int]$MaxExclusive)

            if ($MaxExclusive -le 0) {
                throw 'MaxExclusive must be greater than 0.'
            }

            $Bytes = New-Object byte[] 4
            $Rng.GetBytes($Bytes)
            $Value = [System.BitConverter]::ToUInt32($Bytes, 0)
            return [int]($Value % [uint32]$MaxExclusive)
        }

        $Chars = [System.Collections.Generic.List[char]]::new()

        # At least one character from each required category.
        [void]$Chars.Add($Upper[(& $GetIndex $Upper.Length)])
        [void]$Chars.Add($Lower[(& $GetIndex $Lower.Length)])
        [void]$Chars.Add($Digits[(& $GetIndex $Digits.Length)])
        [void]$Chars.Add($Special[(& $GetIndex $Special.Length)])

        while ($Chars.Count -lt $Length) {
            [void]$Chars.Add($All[(& $GetIndex $All.Length)])
        }

        # Cryptographic Fisher-Yates shuffle, so the required categories
        # are not always in the first four positions.
        for ($i = $Chars.Count - 1; $i -gt 0; $i--) {
            $j = & $GetIndex ($i + 1)
            $Temp      = $Chars[$i]
            $Chars[$i] = $Chars[$j]
            $Chars[$j] = $Temp
        }

        return -join $Chars
    } finally {
        if ($Rng) { $Rng.Dispose() }
    }
}

function Test-ModuleAvailable {
    param([Parameter(Mandatory)][string]$ModuleName)
    return [bool](Get-Module -ListAvailable -Name $ModuleName)
}

function Import-GraphModules {
    if ($script:GraphModulesLoaded) { return }

    $Modules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Users.Actions',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    foreach ($module in $Modules) {
        if (-not (Test-ModuleAvailable -ModuleName $module)) {
            throw "Module missing: $module. Install with: Install-Module $module -Scope CurrentUser"
        }

        Import-Module $module -Scope Global -ErrorAction Stop | Out-Null
    }

    $script:GraphModulesLoaded = $true
}

function Import-ExchangeModule {
    if ($script:ExchangeModuleLoaded) { return }

    $module = 'ExchangeOnlineManagement'

    if (-not (Test-ModuleAvailable -ModuleName $module)) {
        throw "Module missing: $module. Install with: Install-Module $module -Scope CurrentUser"
    }

    Import-Module $module -Scope Global -ErrorAction Stop | Out-Null
    $script:ExchangeModuleLoaded = $true
}

function Import-TeamsModule {
    if ($script:TeamsModuleLoaded) { return }

    $module = 'MicrosoftTeams'

    if (-not (Test-ModuleAvailable -ModuleName $module)) {
        throw "Module missing: $module. Install with: Install-Module $module -Scope CurrentUser"
    }

    Import-Module $module -Scope Global -ErrorAction Stop | Out-Null
    $script:TeamsModuleLoaded = $true
}

function Ensure-TeamsConnection {
    Import-TeamsModule
    Ensure-GraphConnection

    $GraphTenantId = [string]$script:GraphTenantId
    if ([string]::IsNullOrWhiteSpace($GraphTenantId)) {
        throw 'Graph TenantId missing; Teams tenant cannot be validated safely.'
    }

    $CurrentTenantId = $null
    try { $CurrentTenantId = [string](Get-CsTenant -ErrorAction Stop).TenantId } catch { $CurrentTenantId = $null }
    if ($CurrentTenantId -eq $GraphTenantId) { return }

    if (-not [string]::IsNullOrWhiteSpace($CurrentTenantId)) {
        Log-Warning ("Existing Teams connection belongs to tenant {0}, not to Graph tenant {1}. Reconnecting." -f $CurrentTenantId, $GraphTenantId)
        try { Disconnect-MicrosoftTeams -ErrorAction Stop | Out-Null } catch {}
    }

    Log-Info ("Connecting to Microsoft Teams for Graph tenant: {0}" -f $GraphTenantId)
    # Device code: the Windows broker (WAM) fails alongside Graph/Exchange ('Only some brokers
    # (WAM) can log in the current OS account') and with -DisableWAM no sign-in page
    # appears. The code shows up in the PowerShell console window behind the GUI.
    Show-UiInfo 'Sign in to Microsoft Teams: after OK a sign-in code appears in the PowerShell window behind this screen. Open https://login.microsoft.com/device, enter that code and sign in with an admin account of this tenant. The script then continues automatically.'
    Connect-MicrosoftTeams -TenantId $GraphTenantId -UseDeviceAuthentication -ErrorAction Stop | Out-Null

    $CurrentTenantId = [string](Get-CsTenant -ErrorAction Stop).TenantId
    if ($CurrentTenantId -ne $GraphTenantId) {
        throw ("Microsoft Teams is not connected to the same tenant as Graph. Graph: {0}, Teams: {1}." -f $GraphTenantId, $CurrentTenantId)
    }
    Log-Info ("Microsoft Teams connected: tenant {0}" -f $CurrentTenantId)
}

function Ensure-GraphConnectionInteractive {
    Import-GraphModules

    $RequiredScopes = @(
        'User.Read.All', 'User.ReadWrite.All',
        'Group.Read.All', 'GroupMember.ReadWrite.All',
        'Directory.Read.All', 'Directory.ReadWrite.All',
        'Organization.Read.All'
    )

    $NeedsConnect = $true
    $Context = Get-MgContext -ErrorAction SilentlyContinue

    if ($Context) {
        $CurrentScopes = @($Context.Scopes)
        $MissingScopes = @($RequiredScopes | Where-Object { $_ -notin $CurrentScopes })
        if ($MissingScopes.Count -eq 0) {
            $NeedsConnect = $false
        }
    }

    if ($NeedsConnect) {
        Log-Info 'Starting interactive Graph sign-in.'
        Connect-MgGraph -Scopes $RequiredScopes -NoWelcome | Out-Null
        $Context = Get-MgContext -ErrorAction Stop
        Log-Info ("Graph connected: tenant {0}" -f $Context.TenantId)
    }

    if (-not $Context -or [string]::IsNullOrWhiteSpace([string]$Context.TenantId)) {
        throw 'Graph sign-in is not valid: TenantId missing.'
    }

    $script:GraphTenantId      = [string]$Context.TenantId
    $script:GraphConnectedMain = $true
}

function Ensure-GraphConnection {
    Import-GraphModules

    $Context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $Context) {
        throw 'No active Graph session. Sign in interactively first.'
    }

    $CurrentTenantId = [string]$Context.TenantId
    if ([string]::IsNullOrWhiteSpace($CurrentTenantId)) {
        throw 'Active Graph session has no TenantId.'
    }

    if ($script:GraphTenantId -and $script:GraphTenantId -ne $CurrentTenantId) {
        throw ("Graph tenant changed during the session ({0} -> {1}). Sign in again before continuing." -f $script:GraphTenantId, $CurrentTenantId)
    }

    $script:GraphTenantId = $CurrentTenantId
}

function Ensure-ExchangeConnection {
    Import-ExchangeModule
    Ensure-GraphConnection

    $GraphTenantId = [string]$script:GraphTenantId
    if ([string]::IsNullOrWhiteSpace($GraphTenantId)) {
        throw 'Graph TenantId missing; Exchange tenant cannot be validated safely.'
    }

    if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        throw 'ExchangeOnlineManagement v3.0.0 or newer is required (Get-ConnectionInformation is missing).'
    }

    $Connections = @()
    try {
        $Connections = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object {
            $_.State -eq 'Connected' -and -not $_.IsEopSession
        })
    } catch {
        $Connections = @()
    }

    $MatchingConnection = $Connections | Where-Object {
        [string]$_.TenantID -eq $GraphTenantId
    } | Select-Object -First 1

    if ($MatchingConnection) {
        return
    }

    if ($Connections.Count -gt 0) {
        $OtherTenants = @($Connections | ForEach-Object { [string]$_.TenantID } | Where-Object { $_ } | Select-Object -Unique)
        Log-Warning ("Existing Exchange Online connection does not belong to Graph tenant {0}. Reconnecting. EXO tenant(s): {1}" -f $GraphTenantId, ($OtherTenants -join ', '))
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
    }

    Log-Info ("Connecting to Exchange Online for Graph tenant: {0}" -f $GraphTenantId)
    # -DisableWAM: with Graph and Teams in the same session the Windows broker (WAM)
    # sometimes fails with 'Object reference not set'; signing in via the browser does work.
    Connect-ExchangeOnline -ShowBanner:$false -DisableWAM -ErrorAction Stop | Out-Null

    $NewConnection = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object {
        $_.State -eq 'Connected' -and -not $_.IsEopSession -and [string]$_.TenantID -eq $GraphTenantId
    }) | Select-Object -First 1

    if (-not $NewConnection) {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
        throw ("Exchange Online is not connected to the same tenant as Graph. Graph TenantId: {0}. Check the admin account/GDAP connection used." -f $GraphTenantId)
    }

    Log-Info ("Exchange Online connected: tenant {0}" -f $NewConnection.TenantID)
}

function Search-SourceUsers {
    param([Parameter(Mandatory)][string]$SearchTerm)

    Ensure-GraphConnection

    $Trimmed = $SearchTerm.Trim()
    if ([string]::IsNullOrWhiteSpace($Trimmed)) { return @() }

    $Escaped = Escape-ODataString -Value $Trimmed
    $Filter  = "startswith(displayName,'$Escaped') or startswith(userPrincipalName,'$Escaped') or startswith(givenName,'$Escaped') or startswith(surname,'$Escaped')"

    $Users = Get-MgUser `
        -Filter $Filter -ConsistencyLevel eventual -CountVariable userCount -Top 25 `
        -Property Id,DisplayName,UserPrincipalName,GivenName,Surname,Mail,UsageLocation,JobTitle,Department,OfficeLocation,City,Country,CompanyName |
        Sort-Object DisplayName

    return @($Users)
}

function Get-UserByUpn {
    param([Parameter(Mandatory)][string]$UserPrincipalName)

    Ensure-GraphConnection
    return Get-MgUser -UserId $UserPrincipalName `
        -Property Id,DisplayName,UserPrincipalName,Mail,ProxyAddresses,GivenName,Surname,AccountEnabled,AssignedLicenses,LicenseAssignmentStates,UsageLocation,JobTitle,Department,OfficeLocation,City,Country,CompanyName `
        -ErrorAction Stop
}

function Get-TenantVerifiedDomains {
    Ensure-GraphConnection
    $Org = Get-MgOrganization -Property VerifiedDomains | Select-Object -First 1
    return @($Org.VerifiedDomains | ForEach-Object { $_.Name.ToLower() })
}

function Test-UpnDomainAllowed {
    param([Parameter(Mandatory)][string]$UserPrincipalName)

    $Parts = $UserPrincipalName.Split('@')
    if ($Parts.Count -ne 2) { throw 'UPN does not have a valid format.' }

    $Domain = $Parts[1].ToLower()
    $VerifiedDomains = Get-TenantVerifiedDomains

    if ($Domain -notin $VerifiedDomains) {
        throw "UPN domain not found in verified domains: $Domain"
    }
}

function Find-ExistingProxyAddressOwner {
    param([Parameter(Mandatory)][string]$SmtpAddress)

    Ensure-GraphConnection

    $Target  = $SmtpAddress.Trim().ToLower()
    $Escaped = Escape-ODataString -Value $Target
    $Filter  = "proxyAddresses/any(x:x eq 'smtp:$Escaped') or proxyAddresses/any(x:x eq 'SMTP:$Escaped')"
    $EncodedFilter = [System.Uri]::EscapeDataString($Filter)

    $Lookups = @(
        @{
            ObjectType = 'user'
            Uri        = "https://graph.microsoft.com/v1.0/users?`$filter=$EncodedFilter&`$select=id,displayName,userPrincipalName,mail&`$top=1"
            Identity   = 'userPrincipalName'
        },
        @{
            ObjectType = 'group'
            Uri        = "https://graph.microsoft.com/v1.0/groups?`$filter=$EncodedFilter&`$select=id,displayName,mail&`$top=1"
            Identity   = 'mail'
        },
        @{
            ObjectType = 'organizational contact'
            Uri        = "https://graph.microsoft.com/v1.0/contacts?`$filter=$EncodedFilter&`$select=id,displayName,mail&`$top=1"
            Identity   = 'mail'
        }
    )

    foreach ($Lookup in $Lookups) {
        $Response = Invoke-MgGraphRequest -Method GET -Uri $Lookup.Uri -ErrorAction Stop
        $Items = @($Response.value)
        if ($Items.Count -gt 0) {
            $Item = $Items[0]
            $Identity = [string]$Item[$Lookup.Identity]
            if ([string]::IsNullOrWhiteSpace($Identity)) { $Identity = [string]$Item.displayName }

            return [pscustomobject]@{
                ObjectType  = $Lookup.ObjectType
                DisplayName = [string]$Item.displayName
                Identity    = $Identity
            }
        }
    }

    return $null
}

function Test-PasswordStrength {
    param([Parameter(Mandatory)][string]$Password)

    if ($Password.Length -lt 12) {
        throw 'Password must contain at least 12 characters.'
    }

    $Checks = @(
        ($Password -match '[A-Z]'),
        ($Password -match '[a-z]'),
        ($Password -match '[0-9]'),
        ($Password -match '[^a-zA-Z0-9]')
    )

    if (@($Checks | Where-Object { $_ }).Count -lt 4) {
        throw 'Password must contain an uppercase letter, a lowercase letter, a digit and a special character.'
    }
}

function Test-NewUserInput {
    param(
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$NewUpn,
        [Parameter(Mandatory)][string]$Password
    )

    Ensure-GraphConnection

    $FirstName   = $FirstName.Trim()
    $LastName    = $LastName.Trim()
    $DisplayName = $DisplayName.Trim()
    $NewUpn      = $NewUpn.Trim().ToLower()

    if ([string]::IsNullOrWhiteSpace($FirstName))   { throw 'First name is required.' }
    if ([string]::IsNullOrWhiteSpace($LastName))    { throw 'Last name is required.' }
    if ([string]::IsNullOrWhiteSpace($DisplayName)) { throw 'Display name is required.' }
    if ([string]::IsNullOrWhiteSpace($NewUpn))      { throw 'New UPN is required.' }

    if ($NewUpn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        throw 'New UPN does not have a valid format.'
    }

    Test-UpnDomainAllowed -UserPrincipalName $NewUpn
    Test-PasswordStrength -Password $Password

    $EscapedUpn  = Escape-ODataString -Value $NewUpn
    $ExistingUser = Get-MgUser -Filter "userPrincipalName eq '$EscapedUpn'" -Property Id,UserPrincipalName -Top 1 -ErrorAction Stop
    if ($ExistingUser) {
        throw "A user with this UPN already exists: $NewUpn"
    }

    $ExistingProxyOwner = Find-ExistingProxyAddressOwner -SmtpAddress $NewUpn
    if ($ExistingProxyOwner) {
        throw ("The address is already in use as a proxyAddress by {0}: {1}" -f $ExistingProxyOwner.ObjectType, $ExistingProxyOwner.Identity)
    }

    $EscapedDisplayName  = Escape-ODataString -Value $DisplayName
    $DisplayNameConflict = Get-MgUser -Filter "displayName eq '$EscapedDisplayName'" -Property Id,DisplayName,UserPrincipalName -Top 5 -ErrorAction Stop
    if ($DisplayNameConflict) {
        Log-Warning ("DisplayName already exists: {0}" -f ($DisplayNameConflict | Select-Object -First 1).UserPrincipalName)
    }

    $MailNickname     = Get-MailNickname -FirstName $FirstName -LastName $LastName
    $EscapedNickname  = Escape-ODataString -Value $MailNickname
    $NicknameConflict = Get-MgUser -Filter "mailNickname eq '$EscapedNickname'" -Property Id,DisplayName,UserPrincipalName -Top 5 -ErrorAction Stop
    if ($NicknameConflict) {
        Log-Warning ("MailNickname already exists: {0}" -f ($NicknameConflict | Select-Object -First 1).UserPrincipalName)
    }

    [pscustomobject]@{
        FirstName    = $FirstName
        LastName     = $LastName
        DisplayName  = $DisplayName
        NewUpn       = $NewUpn
        MailNickname = $MailNickname
    }
}

function New-CopiedM365User {
    param(
        [Parameter(Mandatory)][string]$SourceUpn,
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$NewUpn,
        [Parameter(Mandatory)][string]$Password,
        [bool]$RequireUsageLocation = $false
    )

    Ensure-GraphConnection

    $SourceUser   = Get-UserByUpn -UserPrincipalName $SourceUpn
    $MailNickname = Get-MailNickname -FirstName $FirstName -LastName $LastName

    $ActiveLicenseAssignments = @($SourceUser.LicenseAssignmentStates | Where-Object { $_.State -eq 'Active' })
    if ($RequireUsageLocation -and $ActiveLicenseAssignments.Count -gt 0 -and [string]::IsNullOrWhiteSpace([string]$SourceUser.UsageLocation)) {
        throw 'Source user has active licenses but no UsageLocation. License assignment can therefore not be copied reliably.'
    }

    $Params = @{
        AccountEnabled    = $true
        GivenName         = $FirstName
        Surname           = $LastName
        DisplayName       = $DisplayName
        UserPrincipalName = $NewUpn
        MailNickname      = $MailNickname
        UsageLocation     = $SourceUser.UsageLocation
        JobTitle          = $SourceUser.JobTitle
        Department        = $SourceUser.Department
        OfficeLocation    = $SourceUser.OfficeLocation
        City              = $SourceUser.City
        Country           = $SourceUser.Country
        CompanyName       = $SourceUser.CompanyName
        PasswordProfile   = @{
            ForceChangePasswordNextSignIn = $true
            Password                      = $Password
        }
    }

    $NullableFields = @('UsageLocation','JobTitle','Department','OfficeLocation','City','Country','CompanyName')
    foreach ($field in $NullableFields) {
        if ($null -eq $Params[$field] -or [string]::IsNullOrWhiteSpace([string]$Params[$field])) {
            $Params.Remove($field)
        }
    }

    $NewUser = New-MgUser @Params -ErrorAction Stop
    Log-Info ("User created: {0} {1}" -f $DisplayName, $NewUpn)

    return [pscustomobject]@{
        SourceUser = $SourceUser
        NewUser    = $NewUser
    }
}

function Wait-ForGraphUser {
    # Right after New-MgUser, Graph may not yet find the user in other calls
    # (Request_ResourceNotFound on group membership or license).
    param(
        [Parameter(Mandatory)][string]$UserId,
        [int]$TimeoutSeconds = 60
    )
    $StartTime = [DateTime]::UtcNow
    while (([DateTime]::UtcNow - $StartTime).TotalSeconds -lt $TimeoutSeconds) {
        try {
            $null = Get-MgUser -UserId $UserId -Property Id -ErrorAction Stop
            return $true
        } catch {
            Wait-UiSeconds -Seconds 5
        }
    }
    Log-Warning ("New user still not findable in Graph after {0}s; follow-up steps will be attempted anyway." -f $TimeoutSeconds)
    return $false
}

function Invoke-GraphWithRetry {
    # Retries a Graph write action only on 'object not found (yet)' (replication delay).
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [int]$Attempts = 4,
        [int]$DelaySeconds = 10
    )
    for ($Attempt = 1; $Attempt -le $Attempts; $Attempt++) {
        try {
            return (& $Action)
        } catch {
            $Message = [string]$_.Exception.Message
            $Transient = $Message -match 'ResourceNotFound|does not exist|not present|could not be found|404'
            if (-not $Transient -or $Attempt -eq $Attempts) { throw }
            Wait-UiSeconds -Seconds $DelaySeconds
        }
    }
}

function Get-LicenseBlocker {
    # Returns $null if waiting for a service makes sense, otherwise the reason why it
    # will not come. Exact data: which SKUs contain a service plan with this name
    # (subscribedSkus), which plans the user has active (assignedPlans) and the status
    # of the license assignments. Only if ALL assignments of such SKUs are in an error state
    # is it a blocker; no license yet = possibly on its way = simply keep waiting.
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$ServicePlanPattern   # e.g. 'EXCHANGE_S_*' or 'MCOEV*'
    )
    try {
        if (-not $script:SkuServicePlans) {
            $Map = @{}
            foreach ($Sku in @(Get-MgSubscribedSku -All -Property SkuId,ServicePlans -ErrorAction Stop)) {
                $Map[[string]$Sku.SkuId] = @(@($Sku.ServicePlans) | Where-Object { $null -ne $_ } | ForEach-Object { [pscustomobject]@{ Id = [string]$_.ServicePlanId; Name = [string]$_.ServicePlanName } })
            }
            $script:SkuServicePlans = $Map
        }
        $User = Get-MgUser -UserId $UserId -Property AssignedPlans,LicenseAssignmentStates -ErrorAction Stop
    } catch {
        return $null
    }

    $PlanIds = @{}
    $SkuIds  = @{}
    foreach ($Key in @($script:SkuServicePlans.Keys)) {
        foreach ($Plan in @($script:SkuServicePlans[$Key] | Where-Object { $_.Name -like $ServicePlanPattern })) {
            $PlanIds[$Plan.Id] = $true
            $SkuIds[$Key] = $true
        }
    }
    if ($SkuIds.Count -eq 0) { return $null }

    $Enabled = @(@($User.AssignedPlans) | Where-Object { $null -ne $_ -and $PlanIds.ContainsKey([string]$_.ServicePlanId) -and [string]$_.CapabilityStatus -eq 'Enabled' }).Count -gt 0
    if ($Enabled) { return $null }

    $Relevant = @(@($User.LicenseAssignmentStates) | Where-Object { $null -ne $_ -and $SkuIds.ContainsKey([string]$_.SkuId) })
    if ($Relevant.Count -eq 0) { return $null }
    $Errors = @($Relevant | Where-Object { [string]$_.State -eq 'Error' })
    if ($Errors.Count -eq $Relevant.Count) {
        return ('license assignment of the new user is in an error state (' + ((@($Errors | ForEach-Object { [string]$_.Error }) | Select-Object -Unique) -join ', ') + ')')
    }
    return $null
}

function Copy-GroupMemberships {
    param(
        [Parameter(Mandatory)][string]$SourceUserId,
        [Parameter(Mandatory)][string]$NewUserId
    )

    Ensure-GraphConnection

    $Results = [System.Collections.Generic.List[object]]::new()
    $Memberships = Get-MgUserMemberOf -UserId $SourceUserId -All -ErrorAction Stop | Where-Object {
        $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group'
    }

    foreach ($Membership in $Memberships) {
        try {
            $Group = Get-MgGroup -GroupId $Membership.Id `
                -Property Id,DisplayName,Mail,GroupTypes,MembershipRule,SecurityEnabled,MailEnabled,OnPremisesSyncEnabled,IsAssignableToRole `
                -ErrorAction Stop

            if ($Group.MembershipRule) {
                $Results.Add([pscustomobject]@{
                    GroupId     = $Group.Id
                    DisplayName = $Group.DisplayName
                    Mail        = $Group.Mail
                    GroupType   = 'Dynamic'
                    Status      = 'SkippedDynamic'
                    Detail      = 'Dynamic group; membership is determined by the rule.'
                })
                Log-Info ("Skipped dynamic group: {0}" -f $Group.DisplayName)
                continue
            }

            if ($Group.OnPremisesSyncEnabled) {
                $Results.Add([pscustomobject]@{
                    GroupId     = $Group.Id
                    DisplayName = $Group.DisplayName
                    Mail        = $Group.Mail
                    GroupType   = 'OnPremises'
                    Status      = 'SkippedOnPremises'
                    Detail      = 'On-premises synchronized group; cloud membership is not writable.'
                })
                Log-Warning ("On-premises group not copied automatically: {0}" -f $Group.DisplayName)
                continue
            }

            if ($Group.IsAssignableToRole) {
                $Results.Add([pscustomobject]@{
                    GroupId     = $Group.Id
                    DisplayName = $Group.DisplayName
                    Mail        = $Group.Mail
                    GroupType   = 'RoleAssignable'
                    Status      = 'SkippedRoleAssignable'
                    Detail      = 'Role-assignable group requires additional privileged permissions.'
                })
                Log-Warning ("Role-assignable group not copied automatically: {0}" -f $Group.DisplayName)
                continue
            }

            $IsMicrosoft365Group       = @($Group.GroupTypes) -contains 'Unified'
            $IsSecurityGroup           = (-not $Group.MailEnabled -and $Group.SecurityEnabled)
            $IsMailEnabledSecurity     = ($Group.MailEnabled -and $Group.SecurityEnabled -and -not $IsMicrosoft365Group)
            $IsDistributionGroup       = ($Group.MailEnabled -and -not $Group.SecurityEnabled -and -not $IsMicrosoft365Group)

            # Distribution groups and mail-enabled security groups are read-only
            # in Microsoft Graph, but can be managed via Exchange Online.
            if ($IsMailEnabledSecurity -or $IsDistributionGroup) {
                $GroupType = if ($IsMailEnabledSecurity) { 'MailEnabledSecurity' } else { 'Distribution' }
                $Results.Add([pscustomobject]@{
                    GroupId     = $Group.Id
                    DisplayName = $Group.DisplayName
                    Mail        = $Group.Mail
                    GroupType   = $GroupType
                    Status      = 'PendingExchange'
                    Detail      = 'Will be added via Exchange Online.'
                })
                Log-Info ("Mail-enabled group passed on to Exchange phase: {0} <{1}>" -f $Group.DisplayName, $Group.Mail)
                continue
            }

            if (-not $IsMicrosoft365Group -and -not $IsSecurityGroup) {
                $Results.Add([pscustomobject]@{
                    GroupId     = $Group.Id
                    DisplayName = $Group.DisplayName
                    Mail        = $Group.Mail
                    GroupType   = 'Unsupported'
                    Status      = 'SkippedReadOnly'
                    Detail      = 'Unknown or unsupported group type; not changed automatically.'
                })
                Log-Warning ("Unsupported group type not copied automatically: {0}" -f $Group.DisplayName)
                continue
            }

            try {
                $null = Invoke-GraphWithRetry -Action { New-MgGroupMemberByRef -GroupId $Group.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$NewUserId" } -ErrorAction Stop }
                $Results.Add([pscustomobject]@{
                    GroupId     = $Group.Id
                    DisplayName = $Group.DisplayName
                    Mail        = $Group.Mail
                    GroupType   = if ($IsMicrosoft365Group) { 'Microsoft365' } else { 'Security' }
                    Status      = 'Added'
                    Detail      = ''
                })
                Log-Info ("Added to group via Graph: {0}" -f $Group.DisplayName)
            } catch {
                $Message = $_.Exception.Message
                $Results.Add([pscustomobject]@{
                    GroupId     = $Group.Id
                    DisplayName = $Group.DisplayName
                    Mail        = $Group.Mail
                    GroupType   = if ($IsMicrosoft365Group) { 'Microsoft365' } else { 'Security' }
                    Status      = 'Failed'
                    Detail      = $Message
                })
                Log-Warning ("Could not add to group via Graph: {0} {1}" -f $Group.DisplayName, $Message)
            }
        } catch {
            $Message = $_.Exception.Message
            $Results.Add([pscustomobject]@{
                GroupId     = $Membership.Id
                DisplayName = $Membership.Id
                Mail        = ''
                GroupType   = 'Unknown'
                Status      = 'Failed'
                Detail      = $Message
            })
            Log-Warning ("Could not retrieve group details: {0} {1}" -f $Membership.Id, $Message)
        }
    }

    return $Results
}

function Wait-ForUserInExchange {
    param(
        [Parameter(Mandatory)][string]$Upn,
        [int]$TimeoutSeconds = 180,
        [int]$IntervalSeconds = 10
    )

    Ensure-ExchangeConnection

    # Exchange provisioning after creation in Graph can take 1–3 minutes.
    # DoEvents() every 200ms keeps the UI responsive; Get-User is only tried every IntervalSeconds.
    $MaxPogingen   = [Math]::Ceiling($TimeoutSeconds / $IntervalSeconds)
    $StartTime     = [DateTime]::UtcNow
    $LastCheckTime = $StartTime.AddSeconds(-$IntervalSeconds)  # Force first check immediately

    Log-Info ("Waiting for Exchange provisioning for: {0} (max {1}s, check every {2}s)" -f $Upn, $TimeoutSeconds, $IntervalSeconds)

    while (([DateTime]::UtcNow - $StartTime).TotalSeconds -lt $TimeoutSeconds) {
        $ElapsedSeconds = [Math]::Round(([DateTime]::UtcNow - $StartTime).TotalSeconds)
        $SecondsSinceCheck = ([DateTime]::UtcNow - $LastCheckTime).TotalSeconds

        if ($SecondsSinceCheck -ge $IntervalSeconds) {
            $Poging = [Math]::Floor($ElapsedSeconds / $IntervalSeconds) + 1
            try {
                $User = Get-User -Identity $Upn -ErrorAction Stop
                if ($User) {
                    Log-Info ("User visible in Exchange after {0}s (attempt {1}/{2}): {3}" -f $ElapsedSeconds, $Poging, $MaxPogingen, $Upn)
                    return
                }
            } catch {
                # Get-User throws an exception as long as the Exchange recipient is not yet visible — normal behavior.
                Log-Info ("Exchange provisioning not ready yet. Attempt {0}/{1} — {2}s elapsed." -f $Poging, $MaxPogingen, $ElapsedSeconds)
            }
            $LastCheckTime = [DateTime]::UtcNow
        }

        # Short sleep + DoEvents: the UI stays responsive during the wait
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }

    $TotalElapsed = [Math]::Round(([DateTime]::UtcNow - $StartTime).TotalSeconds)
    throw ("Exchange provisioning timeout after {0}s for: {1}. Exchange group memberships and/or mailbox permissions could not be processed. Check the Exchange admin center once the user is visible." -f $TotalElapsed, $Upn)
}

function Wait-ForMailboxInExchange {
    # Waits until the new user has a real UserMailbox.
    # Get-User succeeds earlier than the mailbox exists (the mailbox is only created after
    # license assignment). Add-MailboxPermission then succeeds, but
    # Exchange does not record the AutoMapping link and returns no error.
    # Returns $true as soon as the mailbox exists, $false on timeout.
    param(
        [Parameter(Mandatory)][string]$Upn,
        [int]$TimeoutSeconds = 900,
        [int]$IntervalSeconds = 15,
        [string]$GraphUserId = ''   # optional: stop early if this user's license is in an error state
    )

    $script:LastWaitBlocker = $null
    $BlockerHits      = 0
    $LastLicenseCheck = [DateTime]::UtcNow

    Ensure-ExchangeConnection

    $MaxPogingen   = [Math]::Ceiling($TimeoutSeconds / $IntervalSeconds)
    $StartTime     = [DateTime]::UtcNow
    $LastCheckTime = $StartTime.AddSeconds(-$IntervalSeconds)  # Force first check immediately

    Log-Info ("Waiting for UserMailbox for: {0} (max {1}s, check every {2}s)" -f $Upn, $TimeoutSeconds, $IntervalSeconds)

    while (([DateTime]::UtcNow - $StartTime).TotalSeconds -lt $TimeoutSeconds) {
        $ElapsedSeconds    = [Math]::Round(([DateTime]::UtcNow - $StartTime).TotalSeconds)
        $SecondsSinceCheck = ([DateTime]::UtcNow - $LastCheckTime).TotalSeconds

        if ($SecondsSinceCheck -ge $IntervalSeconds) {
            $Poging = [Math]::Floor($ElapsedSeconds / $IntervalSeconds) + 1
            $Status = 'no mailbox'
            try {
                $Mailbox = Get-Mailbox -Identity $Upn -ErrorAction Stop
                $Status  = [string]$Mailbox.RecipientTypeDetails
                if ($Status -eq 'UserMailbox') {
                    Log-Info ("UserMailbox present after {0}s (attempt {1}/{2}): {3}" -f $ElapsedSeconds, $Poging, $MaxPogingen, $Upn)
                    return $true
                }
            } catch {
                # Get-Mailbox fails as long as the mailbox has not been created yet - normal behavior.
            }
            Log-Info ("Mailbox not ready yet ({0}). Attempt {1}/{2} - {3}s elapsed." -f $Status, $Poging, $MaxPogingen, $ElapsedSeconds)

            if ($GraphUserId -and $ElapsedSeconds -ge $LicenseBlockerGraceSeconds -and ([DateTime]::UtcNow - $LastLicenseCheck).TotalSeconds -ge 60) {
                $LastLicenseCheck = [DateTime]::UtcNow
                $Blocker = Get-LicenseBlocker -UserId $GraphUserId -ServicePlanPattern 'EXCHANGE_S_*'
                if ($Blocker) { $BlockerHits++ } else { $BlockerHits = 0 }
                if ($BlockerHits -ge 2) {
                    $script:LastWaitBlocker = $Blocker
                    Log-Warning ("Stopped waiting for mailbox after {0}s: {1}." -f $ElapsedSeconds, $Blocker)
                    return $false
                }
            }
            $LastCheckTime = [DateTime]::UtcNow
        }

        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }

    $TotalElapsed = [Math]::Round(([DateTime]::UtcNow - $StartTime).TotalSeconds)
    Log-Warning ("Mailbox provisioning timeout after {0}s for: {1}. Check whether a license with Exchange Online has been assigned." -f $TotalElapsed, $Upn)
    return $false
}

function Copy-ExchangeGroupMemberships {
    param(
        [Parameter(Mandatory)][object[]]$Groups,
        [Parameter(Mandatory)][string]$NewUpn
    )

    Ensure-ExchangeConnection

    $Results = [System.Collections.Generic.List[object]]::new()

    foreach ($Group in @($Groups)) {
        $Identity = [string]$Group.Mail

        if ([string]::IsNullOrWhiteSpace($Identity)) {
            $Results.Add([pscustomobject]@{
                GroupId     = $Group.GroupId
                DisplayName = $Group.DisplayName
                Mail        = $Group.Mail
                GroupType   = $Group.GroupType
                Status      = 'Failed'
                Detail      = 'Mail-enabled group has no usable primary SMTP address.'
            })
            Log-Warning ("Mail-enabled group has no usable SMTP address: {0}" -f $Group.DisplayName)
            continue
        }

        try {
            $DistributionGroup = Get-DistributionGroup -Identity $Identity -ErrorAction Stop

            # Extra safety net: synced groups must be managed on-premises.
            # The PSObject check prevents an error under StrictMode if an EXO version
            # does not return this property in the default output.
            $IsDirSynced = $false
            if ($DistributionGroup.PSObject.Properties['IsDirSynced']) {
                $IsDirSynced = [bool]$DistributionGroup.IsDirSynced
            }
            if ($IsDirSynced) {
                $Results.Add([pscustomobject]@{
                    GroupId     = $Group.GroupId
                    DisplayName = $Group.DisplayName
                    Mail        = $Identity
                    GroupType   = $Group.GroupType
                    Status      = 'SkippedOnPremises'
                    Detail      = 'Exchange reports that this group is directory-synced; membership must be managed on-premises.'
                })
                Log-Warning ("Directory-synced distribution group not changed automatically: {0}" -f $Group.DisplayName)
                continue
            }

            Add-DistributionGroupMember `
                -Identity $DistributionGroup.Identity `
                -Member $NewUpn `
                -BypassSecurityGroupManagerCheck `
                -ErrorAction Stop

            $Results.Add([pscustomobject]@{
                GroupId     = $Group.GroupId
                DisplayName = $Group.DisplayName
                Mail        = $Identity
                GroupType   = $Group.GroupType
                Status      = 'Added'
                Detail      = ''
            })
            Log-Info ("Added to mail-enabled group via Exchange: {0} <{1}>" -f $Group.DisplayName, $Identity)
        } catch {
            $Message = $_.Exception.Message
            $Results.Add([pscustomobject]@{
                GroupId     = $Group.GroupId
                DisplayName = $Group.DisplayName
                Mail        = $Identity
                GroupType   = $Group.GroupType
                Status      = 'Failed'
                Detail      = $Message
            })
            Log-Warning ("Could not add to mail-enabled group via Exchange: {0} <{1}> {2}" -f $Group.DisplayName, $Identity, $Message)
        }
    }

    return $Results
}

function Copy-SharedMailboxPermissions {
    param(
        [Parameter(Mandatory)][string]$SourceUpn,
        [Parameter(Mandatory)][string]$NewUpn,
        [switch]$ListOnly   # only read which permissions the source has (for manual follow-up steps)
    )

    Ensure-ExchangeConnection

    $Results = [System.Collections.Generic.List[object]]::new()
    $SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop

    foreach ($Mailbox in $SharedMailboxes) {
        $MailboxIdentity = [string]$Mailbox.PrimarySmtpAddress
        if ([string]::IsNullOrWhiteSpace($MailboxIdentity)) {
            $Message = 'Shared mailbox has no usable PrimarySmtpAddress; permissions were not processed.'
            Log-Warning ("{0}: {1}" -f $Mailbox.DisplayName, $Message)
            $Results.Add([pscustomobject]@{
                Mailbox           = $Mailbox.DisplayName
                MailboxAddress    = ''
                FullAccess        = $false
                SendAs            = $false
                FullAccessReadOk  = $false
                SendAsReadOk      = $false
                FullAccessAdded   = $false
                SendAsAdded       = $false
                Status            = 'Partial'
                Detail            = $Message
            })
            continue
        }

        $HasFullAccess    = $false
        $HasSendAs        = $false
        $FullAccessReadOk = $true
        $SendAsReadOk     = $true
        $PermissionErrors = [System.Collections.Generic.List[string]]::new()

        try {
            $MailboxPermissions = Get-MailboxPermission -Identity $MailboxIdentity -User $SourceUpn -ErrorAction Stop | Where-Object {
                $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and [string]$_.Deny -ne 'True'
            }
            $HasFullAccess = [bool]$MailboxPermissions
        } catch {
            $FullAccessReadOk = $false
            $PermissionErrors.Add("Reading FullAccess: $($_.Exception.Message)")
            Log-Warning ("Could not read FullAccess for: {0} <{1}> {2}" -f $Mailbox.DisplayName, $MailboxIdentity, $_.Exception.Message)
        }

        try {
            $RecipientPermissions = Get-RecipientPermission -Identity $MailboxIdentity -Trustee $SourceUpn -ErrorAction Stop | Where-Object {
                $_.AccessRights -contains 'SendAs'
            }
            $HasSendAs = [bool]$RecipientPermissions
        } catch {
            $SendAsReadOk = $false
            $PermissionErrors.Add("Reading SendAs: $($_.Exception.Message)")
            Log-Warning ("Could not read SendAs for: {0} <{1}> {2}" -f $Mailbox.DisplayName, $MailboxIdentity, $_.Exception.Message)
        }

        if ($FullAccessReadOk -and $SendAsReadOk -and -not $HasFullAccess -and -not $HasSendAs) {
            continue
        }

        $FullAccessAdded = $false
        $SendAsAdded     = $false

        if ($ListOnly) {
            $Results.Add([pscustomobject]@{
                Mailbox = $Mailbox.DisplayName; MailboxAddress = $MailboxIdentity
                FullAccess = $HasFullAccess; SendAs = $HasSendAs
                FullAccessReadOk = $FullAccessReadOk; SendAsReadOk = $SendAsReadOk
                FullAccessAdded = $false; SendAsAdded = $false
                Status = 'Listed'; Detail = ($PermissionErrors -join ' | ')
            })
            continue
        }

        if ($HasFullAccess) {
            try {
                Add-MailboxPermission -Identity $MailboxIdentity -User $NewUpn -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -ErrorAction Stop | Out-Null
                $FullAccessAdded = $true
                Log-Info ("FullAccess (AutoMapping) added on: {0} <{1}>" -f $Mailbox.DisplayName, $MailboxIdentity)
            } catch {
                $PermissionErrors.Add("Adding FullAccess: $($_.Exception.Message)")
                Log-Warning ("Could not add FullAccess on: {0} <{1}> {2}" -f $Mailbox.DisplayName, $MailboxIdentity, $_.Exception.Message)
            }
        }

        if ($HasSendAs) {
            try {
                Add-RecipientPermission -Identity $MailboxIdentity -Trustee $NewUpn -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                $SendAsAdded = $true
                Log-Info ("SendAs added on: {0} <{1}>" -f $Mailbox.DisplayName, $MailboxIdentity)
            } catch {
                $PermissionErrors.Add("Adding SendAs: $($_.Exception.Message)")
                Log-Warning ("Could not add SendAs on: {0} <{1}> {2}" -f $Mailbox.DisplayName, $MailboxIdentity, $_.Exception.Message)
            }
        }

        $Status = 'Success'
        if (-not $FullAccessReadOk -or -not $SendAsReadOk -or
            ($HasFullAccess -and -not $FullAccessAdded) -or
            ($HasSendAs -and -not $SendAsAdded)) {
            $Status = 'Partial'
        }

        $Results.Add([pscustomobject]@{
            Mailbox           = $Mailbox.DisplayName
            MailboxAddress    = $MailboxIdentity
            FullAccess        = $HasFullAccess
            SendAs            = $HasSendAs
            FullAccessReadOk  = $FullAccessReadOk
            SendAsReadOk      = $SendAsReadOk
            FullAccessAdded   = $FullAccessAdded
            SendAsAdded       = $SendAsAdded
            Status            = $Status
            Detail            = ($PermissionErrors -join ' | ')
        })
    }

    return $Results
}

function Copy-UserLicenses {
    param(
        [Parameter(Mandatory)][string]$SourceUpn,
        [Parameter(Mandatory)][string]$NewUserId,
        [object]$SourceUserObject = $null
    )

    Ensure-GraphConnection

    $SourceUser = if ($SourceUserObject) { $SourceUserObject } else { Get-UserByUpn -UserPrincipalName $SourceUpn }
    $States = @($SourceUser.LicenseAssignmentStates)
    $Results = [System.Collections.Generic.List[object]]::new()

    if ($States.Count -eq 0) {
        Log-Info 'Source user has no license assignments.'
        return @()
    }

    $InheritedStates = @($States | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.AssignedByGroup) -and $_.State -eq 'Active' })
    $InheritedSkuIds = @{}
    foreach ($InheritedState in $InheritedStates) {
        $InheritedSkuId = [string]$InheritedState.SkuId
        if (-not [string]::IsNullOrWhiteSpace($InheritedSkuId)) {
            $InheritedSkuIds[$InheritedSkuId] = $true
        }
    }

    if ($InheritedSkuIds.Count -gt 0) {
        Log-Info ("Group-based licenses are not assigned directly; they follow via the group memberships. SKUs: {0}" -f $InheritedSkuIds.Count)
    }

    $DirectStates = @($States | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.AssignedByGroup) })
    if ($DirectStates.Count -eq 0) {
        Log-Info 'Source user has no directly assigned licenses.'
        return @()
    }

    # Inventory is retrieved up front so that a missing seat results in a clean Partial
    # instead of a Graph Request_BadRequest. If the inventory check itself
    # fails, we still try to assign the license and Graph remains authoritative.
    $SkuInventory = @{}
    try {
        $SubscribedSkus = @(Get-MgSubscribedSku -All -Property SkuId,SkuPartNumber,ConsumedUnits,PrepaidUnits -ErrorAction Stop)
        foreach ($Sku in $SubscribedSkus) {
            $InventoryKey = [string]$Sku.SkuId
            if (-not [string]::IsNullOrWhiteSpace($InventoryKey)) {
                $SkuInventory[$InventoryKey] = $Sku
            }
        }
    } catch {
        Log-Warning ("License inventory could not be read up front; assignment will be attempted without a precheck. {0}" -f $_.Exception.Message)
    }

    $ProcessedSkuIds = @{}
    foreach ($State in $DirectStates) {
        $SkuIdText = [string]$State.SkuId
        if ([string]::IsNullOrWhiteSpace($SkuIdText) -or $ProcessedSkuIds.ContainsKey($SkuIdText)) {
            continue
        }
        $ProcessedSkuIds[$SkuIdText] = $true

        $Sku = if ($SkuInventory.ContainsKey($SkuIdText)) { $SkuInventory[$SkuIdText] } else { $null }
        $SkuPartNumber = if ($Sku -and -not [string]::IsNullOrWhiteSpace([string]$Sku.SkuPartNumber)) { [string]$Sku.SkuPartNumber } else { $SkuIdText }
        $SkuLabel = if ($SkuPartNumber -eq $SkuIdText) { $SkuIdText } else { "$SkuPartNumber ($SkuIdText)" }

        # If the same SKU is also delivered to the source actively via a group,
        # the direct assignment does not need to be copied again.
        if ($InheritedSkuIds.ContainsKey($SkuIdText)) {
            $Results.Add([pscustomobject]@{
                SkuId         = $SkuIdText
                SkuPartNumber = $SkuPartNumber
                Status        = 'SkippedCoveredByGroup'
                Detail        = 'The same SKU is also active via a group and is delivered via group membership.'
            })
            Log-Info ("Direct license not assigned twice because the same SKU is delivered via a group: {0}" -f $SkuLabel)
            continue
        }

        if ($State.State -ne 'Active') {
            $Detail = "Source license has status '$($State.State)' and error '$($State.Error)'."
            $Results.Add([pscustomobject]@{
                SkuId         = $SkuIdText
                SkuPartNumber = $SkuPartNumber
                Status        = 'SkippedSourceError'
                Detail        = $Detail
            })
            Log-Warning ("Direct source license not copied because of source status {0}: {1}" -f $State.State, $SkuLabel)
            continue
        }

        if ($Sku) {
            $Enabled  = 0
            $Consumed = 0
            if ($Sku.PSObject.Properties['PrepaidUnits'] -and $Sku.PrepaidUnits -and $Sku.PrepaidUnits.PSObject.Properties['Enabled']) {
                $Enabled = [int]$Sku.PrepaidUnits.Enabled
            }
            if ($Sku.PSObject.Properties['ConsumedUnits'] -and $null -ne $Sku.ConsumedUnits) {
                $Consumed = [int]$Sku.ConsumedUnits
            }
            $Available = [Math]::Max(0, ($Enabled - $Consumed))

            Log-Info ("Direct license inventory: {0} — available {1} of {2} (in use {3})." -f $SkuLabel, $Available, $Enabled, $Consumed)

            if ($Available -lt 1) {
                $Detail = "No free license available. Available: $Available of $Enabled; in use: $Consumed."
                $Results.Add([pscustomobject]@{
                    SkuId         = $SkuIdText
                    SkuPartNumber = $SkuPartNumber
                    Status        = 'NoAvailable'
                    Detail        = $Detail
                })
                Log-Warning ("No free license available for {0}. Assignment skipped; the other copy steps continue." -f $SkuLabel)
                continue
            }
        }

        $DisabledPlans = @($State.DisabledPlans | Where-Object { $_ } | ForEach-Object { [Guid][string]$_ })
        $LicenseToAdd = @{
            SkuId         = [Guid]$SkuIdText
            DisabledPlans = $DisabledPlans
        }

        try {
            $null = Invoke-GraphWithRetry -Action { Set-MgUserLicense -UserId $NewUserId -AddLicenses @($LicenseToAdd) -RemoveLicenses @() -ErrorAction Stop }
            $Results.Add([pscustomobject]@{
                SkuId         = $SkuIdText
                SkuPartNumber = $SkuPartNumber
                Status        = 'Assigned'
                Detail        = if ($DisabledPlans.Count -gt 0) { "DisabledPlans copied: $($DisabledPlans.Count)" } else { '' }
            })
            Log-Info ("Direct license assigned: {0} (disabled plans: {1})" -f $SkuLabel, $DisabledPlans.Count)
        } catch {
            $Message = $_.Exception.Message
            $Results.Add([pscustomobject]@{
                SkuId         = $SkuIdText
                SkuPartNumber = $SkuPartNumber
                Status        = 'Failed'
                Detail        = $Message
            })
            Log-Warning ("Could not assign direct license: {0} {1}" -f $SkuLabel, $Message)
        }
    }

    return $Results
}

# ── License transfer (source has left) ──────────────────────────────────────

function Get-SkuInventory {
    $Inventory = @{}
    try {
        foreach ($Sku in @(Get-MgSubscribedSku -All -Property SkuId,SkuPartNumber,ConsumedUnits,PrepaidUnits -ErrorAction Stop)) {
            $Key = [string]$Sku.SkuId
            if (-not [string]::IsNullOrWhiteSpace($Key)) { $Inventory[$Key] = $Sku }
        }
    } catch {
        Log-Warning ("License inventory could not be read: {0}" -f $_.Exception.Message)
    }
    return $Inventory
}

function Get-SkuAvailable {
    # Returns the number of free seats, or $null if the SKU is not in the inventory.
    param([object]$Sku)
    if (-not $Sku) { return $null }
    $Enabled  = 0
    $Consumed = 0
    if ($Sku.PSObject.Properties['PrepaidUnits'] -and $Sku.PrepaidUnits -and $Sku.PrepaidUnits.PSObject.Properties['Enabled']) {
        $Enabled = [int]$Sku.PrepaidUnits.Enabled
    }
    if ($Sku.PSObject.Properties['ConsumedUnits'] -and $null -ne $Sku.ConsumedUnits) {
        $Consumed = [int]$Sku.ConsumedUnits
    }
    return [Math]::Max(0, ($Enabled - $Consumed))
}

function Get-ObjectPropertyValue {
    # StrictMode-safe reading of a property that does not always exist.
    param([object]$InputObject, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $InputObject) { return $null }
    $Property = $InputObject.PSObject.Properties[$Name]
    if ($Property) { return $Property.Value }
    return $null
}

function Get-LicenseTransferPlan {
    # Collects up front (before anything is created) what will happen to the source,
    # so the administrator can confirm or cancel the whole operation.
    param([Parameter(Mandatory)][string]$SourceUpn)

    Ensure-GraphConnection

    $SourceUser = Get-UserByUpn -UserPrincipalName $SourceUpn
    $Inventory  = Get-SkuInventory
    $States     = @(@($SourceUser.LicenseAssignmentStates) | Where-Object { $null -ne $_ })

    $GroupStates = @($States | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.AssignedByGroup) -and $_.State -eq 'Active' })
    $GroupSkuIds = @{}
    foreach ($GroupState in $GroupStates) {
        $GroupSkuIds[[string]$GroupState.SkuId] = $true
    }

    $DirectSkus = [System.Collections.Generic.List[object]]::new()
    $Processed  = @{}
    foreach ($State in @($States | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.AssignedByGroup) })) {
        $SkuIdText = [string]$State.SkuId
        if ([string]::IsNullOrWhiteSpace($SkuIdText) -or $Processed.ContainsKey($SkuIdText)) { continue }
        $Processed[$SkuIdText] = $true

        $Sku = if ($Inventory.ContainsKey($SkuIdText)) { $Inventory[$SkuIdText] } else { $null }
        $DirectSkus.Add([pscustomobject]@{
            SkuId          = $SkuIdText
            SkuPartNumber  = if ($Sku -and -not [string]::IsNullOrWhiteSpace([string]$Sku.SkuPartNumber)) { [string]$Sku.SkuPartNumber } else { $SkuIdText }
            State          = [string]$State.State
            Error          = [string]$State.Error
            DisabledPlans  = @($State.DisabledPlans | Where-Object { $_ } | ForEach-Object { [Guid][string]$_ })
            CoveredByGroup = $GroupSkuIds.ContainsKey($SkuIdText)
        })
    }

    # Group licenses cannot be transferred: the new user joins the same
    # group and then needs a seat of its own. Without a free seat this becomes an attention point.
    $GroupSkusWithoutSeat = [System.Collections.Generic.List[object]]::new()
    foreach ($SkuIdText in @($GroupSkuIds.Keys)) {
        $Sku = if ($Inventory.ContainsKey($SkuIdText)) { $Inventory[$SkuIdText] } else { $null }
        $Available = Get-SkuAvailable -Sku $Sku
        if ($null -eq $Available -or $Available -ge 1) { continue }

        $GroupNames = foreach ($GroupState in @($GroupStates | Where-Object { [string]$_.SkuId -eq $SkuIdText })) {
            $GroupId = [string]$GroupState.AssignedByGroup
            try { [string](Get-MgGroup -GroupId $GroupId -Property DisplayName -ErrorAction Stop).DisplayName } catch { $GroupId }
        }
        $GroupSkusWithoutSeat.Add([pscustomobject]@{
            SkuId         = $SkuIdText
            SkuPartNumber = if ($Sku -and -not [string]::IsNullOrWhiteSpace([string]$Sku.SkuPartNumber)) { [string]$Sku.SkuPartNumber } else { $SkuIdText }
            Groups        = (@($GroupNames | Select-Object -Unique) -join ', ')
            GroupIds      = @($GroupStates | Where-Object { [string]$_.SkuId -eq $SkuIdText } | ForEach-Object { [string]$_.AssignedByGroup } | Select-Object -Unique)
        })
    }

    # The source holds on to those seats as long as it is a member of the license group. Per group, determine
    # whether that can be resolved: assigned group = remove the source; dynamic group with
    # accountEnabled in the rule = block the source's sign-in; otherwise manual.
    $GroupsHoldingSeats = [System.Collections.Generic.List[object]]::new()
    foreach ($GroupId in @($GroupSkusWithoutSeat | ForEach-Object { $_.GroupIds } | Select-Object -Unique)) {
        $Group = $null
        try { $Group = Get-MgGroup -GroupId $GroupId -Property Id,DisplayName,GroupTypes,MembershipRule,MembershipRuleProcessingState,MailEnabled,OnPremisesSyncEnabled -ErrorAction Stop } catch { $Group = $null }
        $IsDynamic  = [bool]($Group -and (@($Group.GroupTypes) -contains 'DynamicMembership'))
        $Rule       = if ($Group) { [string]$Group.MembershipRule } else { '' }
        $RuleOn     = $Group -and [string]$Group.MembershipRuleProcessingState -eq 'On'
        # Blocking only helps if the rule hard-excludes disabled accounts: accountEnabled
        # as a condition in a pure and-chain (an -or can still let the source qualify).
        $BlockWorks = $IsDynamic -and $RuleOn -and ($Rule -match 'user\.accountEnabled\s+-eq\s+true') -and ($Rule -notmatch '\s-or\s')
        $IsUnified  = [bool]($Group -and (@($Group.GroupTypes) -contains 'Unified'))
        $MailOrSync = $Group -and ((([bool]$Group.MailEnabled) -and -not $IsUnified) -or [bool]$Group.OnPremisesSyncEnabled)
        $Manual     = if (-not $Group) { 'group could not be read' }
                      elseif ($IsDynamic -and -not $RuleOn) { 'dynamic group with paused rule processing' }
                      elseif ($IsDynamic -and -not $BlockWorks) { 'dynamic group whose rule does not (unambiguously) exclude disabled accounts; adjust the rule or source attributes' }
                      elseif ($MailOrSync) { 'mail-enabled or on-premises synchronized group; change membership via Exchange/AD' }
                      else { '' }
        $GroupsHoldingSeats.Add([pscustomobject]@{
            Id                 = $GroupId
            DisplayName        = if ($Group) { [string]$Group.DisplayName } else { $GroupId }
            IsDynamic          = $IsDynamic
            ReleasableByBlock  = [bool]$BlockWorks
            ReleasableByRemove = [bool]($Group -and -not $IsDynamic -and -not $MailOrSync)
            ManualReason       = $Manual
            Skus               = (@($GroupSkusWithoutSeat | Where-Object { $_.GroupIds -contains $GroupId } | ForEach-Object { $_.SkuPartNumber }) -join ', ')
        })
    }

    # Per SKU: is the seat freed once all releasable groups have been handled?
    foreach ($Item in $GroupSkusWithoutSeat) {
        $AllReleasable = @($Item.GroupIds | Where-Object { $Gid = $_; @($GroupsHoldingSeats | Where-Object { $_.Id -eq $Gid -and ($_.ReleasableByBlock -or $_.ReleasableByRemove) }).Count -eq 0 }).Count -eq 0
        $Item | Add-Member -NotePropertyName AllGroupsReleasable -NotePropertyValue $AllReleasable
    }

    # Source mailbox: without a license a UserMailbox is deleted after 30 days.
    $MailboxInfo = [pscustomobject]@{
        Checked        = $false
        Type           = ''
        SizeGB         = $null
        ArchiveActive  = $false
        LitigationHold = $false
        Detail         = ''
    }
    try {
        Ensure-ExchangeConnection
        $EscapedUpn = Escape-ODataString -Value $SourceUpn
        $Mailbox = Get-Mailbox -Filter "UserPrincipalName -eq '$EscapedUpn'" -ResultSize 1 -ErrorAction Stop | Select-Object -First 1
        $MailboxInfo.Checked = $true
        if ($Mailbox) {
            $MailboxInfo.Type           = [string]$Mailbox.RecipientTypeDetails
            $MailboxInfo.ArchiveActive  = ([string](Get-ObjectPropertyValue -InputObject $Mailbox -Name 'ArchiveStatus') -eq 'Active')
            $MailboxInfo.LitigationHold = ([string](Get-ObjectPropertyValue -InputObject $Mailbox -Name 'LitigationHoldEnabled') -eq 'True')
            try {
                $Stats    = Get-MailboxStatistics -Identity ([string]$Mailbox.PrimarySmtpAddress) -ErrorAction Stop
                $SizeText = [string](Get-ObjectPropertyValue -InputObject $Stats -Name 'TotalItemSize')
                if ($SizeText -match '\(([\d,\.]+) bytes\)') {
                    $MailboxInfo.SizeGB = [Math]::Round(([double]($Matches[1] -replace '[,\.]', '')) / 1GB, 1)
                }
            } catch {
                Log-Warning ("Mailbox size of source could not be determined: {0}" -f $_.Exception.Message)
            }
        }
    } catch {
        $MailboxInfo.Detail = $_.Exception.Message
        Log-Warning ("Source mailbox could not be checked: {0}" -f $MailboxInfo.Detail)
    }

    return [pscustomobject]@{
        SourceUser           = $SourceUser
        DirectSkus           = $DirectSkus
        GroupSkusWithoutSeat = $GroupSkusWithoutSeat
        GroupsHoldingSeats   = $GroupsHoldingSeats
        InventoryUnknown     = ($Inventory.Count -eq 0)
        Mailbox              = $MailboxInfo
    }
}

function Confirm-LicenseTransfer {
    # Shows the overview. Returns $null on cancel, otherwise an object with
    # ConvertMailbox and ReleaseGroups ($true/$false).
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][string]$NewUpn
    )

    $NL = [Environment]::NewLine
    $SourceUpn     = [string]$Plan.SourceUser.UserPrincipalName
    $Transferable  = @($Plan.DirectSkus | Where-Object { $_.State -eq 'Active' -and -not $_.CoveredByGroup })
    $CoveredDirect = @($Plan.DirectSkus | Where-Object { $_.State -eq 'Active' -and $_.CoveredByGroup })
    $NotTransfered = @($Plan.DirectSkus | Where-Object { $_.State -ne 'Active' })
    $Releasable    = @($Plan.GroupsHoldingSeats | Where-Object { $_.ReleasableByBlock -or $_.ReleasableByRemove })
    $ReleasableIds = @($Releasable | ForEach-Object { $_.Id })
    $Mailbox       = $Plan.Mailbox
    $NeedsLicenseAsShared = (($null -ne $Mailbox.SizeGB -and $Mailbox.SizeGB -gt 50) -or $Mailbox.ArchiveActive -or $Mailbox.LitigationHold)

    $Lines = [System.Collections.Generic.List[string]]::new()
    $Lines.Add(("Transfer licenses from {0} to {1}." -f $SourceUpn, $NewUpn))
    $Lines.Add('')

    if ($Plan.SourceUser.AccountEnabled) {
        $Lines.Add('NOTE: the source account is still ENABLED. After the transfer this user will no longer have these licenses.')
        $Lines.Add('')
    }
    if ($Plan.InventoryUnknown) {
        $Lines.Add('NOTE: the license inventory could not be read. Whether there are free seats is unknown; licenses via groups are therefore not assessed.')
        $Lines.Add('')
    }

    if ($Transferable.Count -gt 0) {
        $Lines.Add('Will be removed from the source and assigned to the new user:')
        foreach ($Item in $Transferable) { $Lines.Add(("  - {0}" -f $Item.SkuPartNumber)) }
    } else {
        $Lines.Add('The source has no directly assigned licenses that can be transferred.')
    }

    if ($CoveredDirect.Count -gt 0) {
        $Lines.Add('')
        $Lines.Add('Assigned directly and via a group (the new user gets these via the group; the direct assignment on the source is only removed if the source also leaves the group):')
        foreach ($Item in $CoveredDirect) { $Lines.Add(("  - {0}" -f $Item.SkuPartNumber)) }
    }

    if ($NotTransfered.Count -gt 0) {
        $Lines.Add('')
        $Lines.Add('Not transferred:')
        foreach ($Item in $NotTransfered) { $Lines.Add(("  - {0} (source status {1})" -f $Item.SkuPartNumber, $Item.State)) }
    }

    if (@($Plan.GroupsHoldingSeats).Count -gt 0) {
        $Lines.Add('')
        $Lines.Add('License groups without a free seat; the source holds the seat as long as it is a member:')
        foreach ($Group in $Plan.GroupsHoldingSeats) {
            $How = if ($Group.ReleasableByRemove) { 'assigned group: freed by removing the source from the group' }
                   elseif ($Group.ReleasableByBlock) { 'dynamic, rule excludes disabled accounts: freed by blocking the source''s sign-in' }
                   else { "manual: $($Group.ManualReason)" }
            $Lines.Add(("  - {0} ({1}): {2}" -f $Group.DisplayName, $Group.Skus, $How))
        }
        $Partial = @($Plan.GroupSkusWithoutSeat | Where-Object { -not $_.AllGroupsReleasable -and @($_.GroupIds | Where-Object { $ReleasableIds -contains $_ }).Count -gt 0 })
        foreach ($Item in $Partial) {
            $Lines.Add(("  Note: {0} is also in a group that cannot be freed automatically; that seat will not be freed." -f $Item.SkuPartNumber))
        }
    }

    $RemovesFromSource = ($Transferable.Count -gt 0 -or $Releasable.Count -gt 0)
    $AskMailbox = ($RemovesFromSource -and $Mailbox.Type -eq 'UserMailbox')

    $Lines.Add('')
    if ($AskMailbox) {
        $Details = [System.Collections.Generic.List[string]]::new()
        if ($null -ne $Mailbox.SizeGB) { $Details.Add(("{0} GB" -f $Mailbox.SizeGB)) }
        if ($Mailbox.ArchiveActive)    { $Details.Add('archive active') }
        if ($Mailbox.LitigationHold)   { $Details.Add('litigation hold') }
        $DetailText = if ($Details.Count -gt 0) { ' (' + ($Details -join ', ') + ')' } else { '' }

        $Lines.Add(("The source has a personal mailbox{0}. Without a license it will be deleted after 30 days." -f $DetailText))
        if ($null -eq $Mailbox.SizeGB -and -not $NeedsLicenseAsShared) {
            $Lines.Add('NOTE: the mailbox size could not be determined. A shared mailbox can only be used without a license up to 50 GB, without an archive and without a hold; check the size.')
            $Lines.Add('')
            $Lines.Add('Yes = first convert the source mailbox to a shared mailbox, then transfer (if the mailbox is smaller than 50 GB)')
        } elseif ($NeedsLicenseAsShared) {
            $Lines.Add('NOTE: above 50 GB, with an archive or with a hold, a SHARED mailbox also needs a license. Converting alone does not prevent data loss in that case; arrange a license for the shared mailbox first or cancel.')
            $Lines.Add('')
            $Lines.Add('Yes = convert the source mailbox to a shared mailbox and transfer (arrange a license for the shared mailbox yourself)')
        } else {
            $Lines.Add('A shared mailbox does not need a license, provided it is smaller than 50 GB, without an archive and without a hold.')
            $Lines.Add('')
            $Lines.Add('Yes = first convert the source mailbox to a shared mailbox, then transfer (recommended)')
        }
        $Lines.Add('No = transfer without converting the mailbox')
        $Lines.Add('Cancel = stop, nothing will be created')
    } else {
        if ($RemovesFromSource -and -not $Mailbox.Checked) {
            $Lines.Add('NOTE: the source mailbox could not be checked. If the source has a personal mailbox, it will be deleted 30 days after the license is removed. Check this first; if in doubt, choose No.')
        } elseif (-not $Mailbox.Checked) {
            $Lines.Add('Source mailbox could not be checked.')
        } elseif ($Mailbox.Type) {
            $Lines.Add(("Source mailbox: {0}." -f $Mailbox.Type))
            if ($RemovesFromSource -and $Mailbox.Type -eq 'SharedMailbox' -and $NeedsLicenseAsShared) {
                $Lines.Add('NOTE: this shared mailbox is larger than 50 GB or has an archive/hold and therefore needs a license of its own.')
            }
        } else {
            $Lines.Add('The source has no mailbox.')
        }
        $Lines.Add('')
        $Lines.Add('Continue? (No = stop, nothing will be created)')
    }

    $Message = $Lines -join $NL
    Log-Info ("License transfer confirmation requested for source {0}." -f $SourceUpn)

    if ($AskMailbox) {
        $Answer = Show-UiQuestion -Message $Message -Buttons 'YesNoCancel'
        if ($Answer -eq 'Cancel') { return $null }
        $Convert = ($Answer -eq 'Yes')
    } else {
        $Answer = Show-UiQuestion -Message $Message -Buttons 'YesNo'
        if ($Answer -ne 'Yes') { return $null }
        $Convert = $false
    }

    # Second question only if seats can be freed via groups.
    $Release = $false
    if ($Releasable.Count -gt 0) {
        $Block  = @($Releasable | Where-Object { $_.ReleasableByBlock }).Count -gt 0
        $Remove = @($Releasable | Where-Object { $_.ReleasableByRemove })
        $Steps  = [System.Collections.Generic.List[string]]::new()
        if ($Block -and $Plan.SourceUser.AccountEnabled) { $Steps.Add(("block sign-in of {0} and revoke sessions" -f $SourceUpn)) }
        foreach ($Group in $Remove) { $Steps.Add(("remove {0} from group {1}" -f $SourceUpn, $Group.DisplayName)) }
        if ($CoveredDirect.Count -gt 0) { $Steps.Add('then remove the direct assignment of the same license(s) from the source') }
        if ($Steps.Count -gt 0) {
            $ReleaseLines = @(
                'Phase out the source so the group licenses become available for the new user?'
                ''
            ) + @($Steps | ForEach-Object { "  - $_" }) + @(
                ''
                'The script then waits until the source has lost the licenses, has the new user''s licenses reprocessed and checks whether they are active.'
                ''
                'Yes = phase out'
                'No = skip (the new user will not get these licenses yet; the source keeps them)'
            )
            $Release = ((Show-UiQuestion -Message ($ReleaseLines -join $NL) -Buttons 'YesNo') -eq 'Yes')
        } else {
            $Release = $true   # source already blocked; only wait, reprocess and check
        }
    }

    Log-Info ("License transfer confirmed. Convert source mailbox to shared: {0}. Phase out source for group licenses: {1}" -f $Convert, $Release)
    return [pscustomobject]@{ ConvertMailbox = $Convert; ReleaseGroups = $Release }
}

function Convert-SourceMailboxToShared {
    param(
        [Parameter(Mandatory)][string]$SourceUpn,
        [int]$TimeoutSeconds = 60
    )

    Ensure-ExchangeConnection
    Set-Mailbox -Identity $SourceUpn -Type Shared -ErrorAction Stop
    Log-Info ("Source mailbox converted to shared mailbox: {0}" -f $SourceUpn)

    # Only remove licenses once Exchange really reports the mailbox as shared.
    $EscapedUpn = Escape-ODataString -Value $SourceUpn
    $StartTime  = [DateTime]::UtcNow
    while (([DateTime]::UtcNow - $StartTime).TotalSeconds -lt $TimeoutSeconds) {
        $Mailbox = Get-Mailbox -Filter "UserPrincipalName -eq '$EscapedUpn'" -ResultSize 1 -ErrorAction Stop | Select-Object -First 1
        if ($Mailbox -and [string]$Mailbox.RecipientTypeDetails -eq 'SharedMailbox') { return }
        Wait-UiSeconds -Seconds 10
    }
    throw ("Source mailbox {0} WAS converted to shared (Set-Mailbox succeeded), but after {1}s it is not yet visible as a shared mailbox." -f $SourceUpn, $TimeoutSeconds)
}

function Release-SourceGroupLicenses {
    # Removes the source from license groups that hold seats: block sign-in
    # (dynamic group that excludes disabled accounts) and/or remove from assigned groups.
    # Waits until the source is no longer a member. Returns per group whether that succeeded, plus
    # the issues the administrator needs to know about.
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][object[]]$Groups
    )

    $Source    = $Plan.SourceUser
    $SourceUpn = [string]$Source.UserPrincipalName
    $Issues    = [System.Collections.Generic.List[string]]::new()
    $Attempted = [System.Collections.Generic.List[object]]::new()

    $BlockGroups = @($Groups | Where-Object { $_.ReleasableByBlock })
    if ($BlockGroups.Count -gt 0) {
        $Blocked = $false
        try {
            $Enabled = [bool](Get-MgUser -UserId $Source.Id -Property AccountEnabled -ErrorAction Stop).AccountEnabled
            if ($Enabled) {
                Update-MgUser -UserId $Source.Id -AccountEnabled:$false -ErrorAction Stop
                Log-Info ("Sign-in blocked for source {0}." -f $SourceUpn)
            }
            $Blocked = $true
        } catch {
            $Issues.Add(("Blocking sign-in of {0} failed: {1}" -f $SourceUpn, $_.Exception.Message))
            Log-Warning $Issues[$Issues.Count - 1]
        }
        if ($Blocked) {
            foreach ($Group in $BlockGroups) { $Attempted.Add($Group) }
            try {
                Revoke-MgUserSignInSession -UserId $Source.Id -ErrorAction Stop | Out-Null
                Log-Info ("Sessions revoked for source {0}." -f $SourceUpn)
            } catch {
                $Issues.Add(("Revoking sessions of {0} failed: {1}" -f $SourceUpn, $_.Exception.Message))
                Log-Warning $Issues[$Issues.Count - 1]
            }
        }
    }

    foreach ($Group in @($Groups | Where-Object { $_.ReleasableByRemove })) {
        try {
            Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $Source.Id -ErrorAction Stop
            $Attempted.Add($Group)
            Log-Info ("Source {0} removed from group {1}." -f $SourceUpn, $Group.DisplayName)
        } catch {
            $Issues.Add(("Removing source {0} from group {1} failed: {2}" -f $SourceUpn, $Group.DisplayName, $_.Exception.Message))
            Log-Warning $Issues[$Issues.Count - 1]
        }
    }

    # Only wait for groups for which the action was actually carried out.
    $Still = @($Attempted)
    if ($Attempted.Count -gt 0) {
        $StartTime = [DateTime]::UtcNow
        while ($true) {
            $MemberOf = @($Attempted | ForEach-Object { $_.Id })
            try { $MemberOf = @(Get-MgUserMemberOf -UserId $Source.Id -All -ErrorAction Stop | ForEach-Object { [string]$_.Id }) } catch { }
            $Still = @($Attempted | Where-Object { $MemberOf -contains $_.Id })
            if ($Still.Count -eq 0) { break }
            if (([DateTime]::UtcNow - $StartTime).TotalSeconds -ge $GroupReleaseWaitSeconds) { break }
            Log-Info ("Source still a member of: {0}. Waiting..." -f (($Still | ForEach-Object { $_.DisplayName }) -join ', '))
            Wait-UiSeconds -Seconds 30
        }
    }

    $StillIds  = @($Still | ForEach-Object { $_.Id })
    $DoneIds   = @($Attempted | ForEach-Object { $_.Id })
    $Released  = @{}
    foreach ($Group in $Groups) { $Released[$Group.Id] = (($DoneIds -contains $Group.Id) -and -not ($StillIds -contains $Group.Id)) }
    foreach ($Group in $Still) {
        $Issues.Add(("Source still a member of group {1} after {0}s; the seat is not free yet." -f $GroupReleaseWaitSeconds, $Group.DisplayName))
        Log-Warning $Issues[$Issues.Count - 1]
    }

    return [pscustomobject]@{ Released = $Released; Issues = $Issues }
}

function Complete-GroupLicenseHandover {
    # After releasing: wait until the source has really lost the SKUs, have the new user's
    # licenses reprocessed and check per SKU whether they are active via the group.
    # Returns per SkuId: 'Active' or a description of what is still missing.
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][object[]]$Skus,
        [Parameter(Mandatory)][string]$NewUserId
    )

    $SourceId = [string]$Plan.SourceUser.Id
    $SkuIds   = @($Skus | ForEach-Object { [string]$_.SkuId })
    $Result   = @{}

    # 1. Source lost them? (the membership is gone, but revoking the license runs asynchronously)
    $StartTime = [DateTime]::UtcNow
    while (([DateTime]::UtcNow - $StartTime).TotalSeconds -lt $LicenseVerifySeconds) {
        $Held = @()
        try {
            $States = @(@((Get-MgUser -UserId $SourceId -Property LicenseAssignmentStates -ErrorAction Stop).LicenseAssignmentStates) | Where-Object { $null -ne $_ })
            $Held = @($States | Where-Object { $SkuIds -contains [string]$_.SkuId })
        } catch { }
        if ($Held.Count -eq 0) { break }
        Log-Info 'Source still has the released licenses; waiting...'
        Wait-UiSeconds -Seconds 20
    }

    # 2. Reprocess and 3. check, with retries (the seat sometimes only counts later).
    $StartTime = [DateTime]::UtcNow
    $LastReprocess = [DateTime]::MinValue
    $Pending = @($Skus)
    while ($true) {
        if (([DateTime]::UtcNow - $LastReprocess).TotalSeconds -ge 60) {
            try {
                Invoke-MgLicenseUser -UserId $NewUserId -ErrorAction Stop | Out-Null
                Log-Info 'Requested license reprocessing for the new user.'
            } catch {
                Log-Info ("Reprocessing not successful yet: {0}" -f $_.Exception.Message)
            }
            $LastReprocess = [DateTime]::UtcNow
        }

        $NewStates = @()
        try { $NewStates = @(@((Get-MgUser -UserId $NewUserId -Property LicenseAssignmentStates -ErrorAction Stop).LicenseAssignmentStates) | Where-Object { $null -ne $_ }) } catch { }
        foreach ($Sku in $Pending) {
            $State = @($NewStates | Where-Object { [string]$_.SkuId -eq [string]$Sku.SkuId -and $Sku.GroupIds -contains [string]$_.AssignedByGroup }) | Select-Object -First 1
            $Result[[string]$Sku.SkuId] = if ($State -and [string]$State.State -eq 'Active') { 'Active' }
                                          elseif ($State) { "status $($State.State) ($($State.Error))" }
                                          else { 'not yet assigned via the group (is the new user a member?)' }
        }
        $Pending = @($Pending | Where-Object { $Result[[string]$_.SkuId] -ne 'Active' })
        if ($Pending.Count -eq 0) { break }
        if (([DateTime]::UtcNow - $StartTime).TotalSeconds -ge $LicenseVerifySeconds) { break }
        Wait-UiSeconds -Seconds 20
    }

    foreach ($Sku in $Pending) {
        $Member = $null
        try { $Member = @(Get-MgUserMemberOf -UserId $NewUserId -All -ErrorAction Stop | ForEach-Object { [string]$_.Id }) } catch { }
        if ($null -ne $Member -and @($Sku.GroupIds | Where-Object { $Member -contains $_ }).Count -eq 0) {
            $Result[[string]$Sku.SkuId] = "the new user is not a member of $($Sku.Groups) (check whether the group rule matches this user)"
        }
    }
    return $Result
}

function Move-UserLicenses {
    # Order:
    #  1. remove direct licenses from the source (one call) and 2. assign them to the new user;
    #     3. on failure, return them to the source.
    #  4. only if 1-3 succeeded and the administrator wants it: source out of license groups
    #     (block / remove from assigned group), 5. only then remove the direct assignment of the same SKU
    #     from the source, 6. check whether the new user has the license active.
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][string]$NewUserId,
        [bool]$ConvertMailbox = $false,
        [bool]$ReleaseGroups = $false
    )

    Ensure-GraphConnection

    $SourceUser = $Plan.SourceUser
    $SourceUpn  = [string]$SourceUser.UserPrincipalName
    $Results    = [System.Collections.Generic.List[object]]::new()

    foreach ($Item in @($Plan.DirectSkus | Where-Object { $_.State -ne 'Active' })) {
        $Results.Add([pscustomobject]@{ SkuId = $Item.SkuId; SkuPartNumber = $Item.SkuPartNumber; Status = 'SkippedSourceError'; Detail = "source license has status '$($Item.State)' ($($Item.Error)) and was not transferred." })
        Log-Warning ("Not transferred because of source status {0}: {1}" -f $Item.State, $Item.SkuPartNumber)
    }

    $Transferable  = @($Plan.DirectSkus | Where-Object { $_.State -eq 'Active' -and -not $_.CoveredByGroup })
    $CoveredDirect = @($Plan.DirectSkus | Where-Object { $_.State -eq 'Active' -and $_.CoveredByGroup })
    $Releasable    = @(if ($ReleaseGroups) { $Plan.GroupsHoldingSeats | Where-Object { $_.ReleasableByBlock -or $_.ReleasableByRemove } })

    if ($Transferable.Count -eq 0 -and $Releasable.Count -eq 0) {
        Log-Info 'No licenses to transfer or release.'
    } elseif ($ConvertMailbox) {
        try {
            Convert-SourceMailboxToShared -SourceUpn $SourceUpn -TimeoutSeconds $SharedMailboxWaitSeconds
        } catch {
            throw ("{0} No licenses have been removed yet. Check the source mailbox and then transfer manually." -f $_.Exception.Message)
        }
    }

    # Steps 1-3: direct licenses.
    $TransferOk = $true
    if ($Transferable.Count -gt 0) {
        $Labels      = ($Transferable | ForEach-Object { $_.SkuPartNumber }) -join ', '
        $LicensesAdd = @($Transferable | ForEach-Object { @{ SkuId = [Guid]$_.SkuId; DisabledPlans = @($_.DisabledPlans) } })
        $Removed     = $false
        try {
            Set-MgUserLicense -UserId $SourceUser.Id -AddLicenses @() -RemoveLicenses @($Transferable | ForEach-Object { [Guid]$_.SkuId }) -ErrorAction Stop | Out-Null
            $Removed = $true
            Log-Info ("Direct licenses removed from source {0}: {1}" -f $SourceUpn, $Labels)
        } catch {
            $TransferOk = $false
            $Message = $_.Exception.Message
            Log-Warning ("Licenses could not be removed from the source; the source is unchanged. {0}" -f $Message)
            foreach ($Item in $Transferable) {
                $Results.Add([pscustomobject]@{ SkuId = $Item.SkuId; SkuPartNumber = $Item.SkuPartNumber; Status = 'Failed'; Detail = "removal from the source failed; the source still has the license and nothing was transferred ($Message)." })
            }
        }

        if ($Removed) {
            $Assigned  = $false
            $LastError = ''
            for ($Attempt = 1; $Attempt -le $LicenseTransferRetryCount; $Attempt++) {
                try {
                    Set-MgUserLicense -UserId $NewUserId -AddLicenses $LicensesAdd -RemoveLicenses @() -ErrorAction Stop | Out-Null
                    $Assigned = $true
                    break
                } catch {
                    $LastError = $_.Exception.Message
                    Log-Info ("Assignment to new user not successful yet (attempt {0}/{1}): {2}" -f $Attempt, $LicenseTransferRetryCount, $LastError)
                    if ($Attempt -lt $LicenseTransferRetryCount) { Wait-UiSeconds -Seconds $LicenseTransferRetrySeconds }
                }
            }

            if ($Assigned) {
                foreach ($Item in $Transferable) {
                    $Results.Add([pscustomobject]@{
                        SkuId         = $Item.SkuId
                        SkuPartNumber = $Item.SkuPartNumber
                        Status        = 'Transferred'
                        Detail        = if ($Item.DisabledPlans.Count -gt 0) { "DisabledPlans copied: $($Item.DisabledPlans.Count)" } else { '' }
                    })
                }
                Log-Info ("Licenses transferred to new user: {0}" -f $Labels)
            } else {
                $TransferOk = $false
                Log-Warning ("Assignment to new user failed; licenses are being returned to the source. {0}" -f $LastError)
                try {
                    Set-MgUserLicense -UserId $SourceUser.Id -AddLicenses $LicensesAdd -RemoveLicenses @() -ErrorAction Stop | Out-Null
                    Log-Info ("Licenses returned to source {0}: {1}" -f $SourceUpn, $Labels)
                    $Status = 'RolledBack'
                    $Detail = "assignment to the new user failed; returned to the source ($LastError)."
                } catch {
                    $RollbackError = $_.Exception.Message
                    Log-Error ("Returning to source failed. Source {0} now has NO licenses {1}. {2}" -f $SourceUpn, $Labels, $RollbackError)
                    $Status = 'RollbackFailed'
                    $Detail = "NOT assigned and NOT returned to the source; assign manually ($LastError / $RollbackError)."
                }
                foreach ($Item in $Transferable) {
                    $Results.Add([pscustomobject]@{ SkuId = $Item.SkuId; SkuPartNumber = $Item.SkuPartNumber; Status = $Status; Detail = $Detail })
                }
            }
        }
    }

    # Step 4: source out of license groups, only if the direct transfer succeeded.
    $Released = @{}
    if ($Releasable.Count -gt 0 -and -not $TransferOk) {
        Log-Warning 'Phasing out the source skipped because the transfer of direct licenses did not succeed.'
        $Results.Add([pscustomobject]@{ SkuId = ''; SkuPartNumber = 'group licenses'; Status = 'ReleaseIssue'; Detail = 'source not phased out (not blocked, not removed from groups) because the transfer of direct licenses failed.' })
    } elseif ($Releasable.Count -gt 0) {
        $Release = Release-SourceGroupLicenses -Plan $Plan -Groups $Releasable
        $Released = $Release.Released
        foreach ($Issue in $Release.Issues) {
            $Results.Add([pscustomobject]@{ SkuId = ''; SkuPartNumber = 'phase-out'; Status = 'ReleaseIssue'; Detail = $Issue })
        }
    }

    # A SKU is only freed once ALL groups through which the source receives it have been released.
    $ReleasedSkus = @($Plan.GroupSkusWithoutSeat | Where-Object { $Item = $_; @($Item.GroupIds | Where-Object { -not $Released[$_] }).Count -eq 0 -and $Item.GroupIds.Count -gt 0 })
    $ReleasedIds  = @($ReleasedSkus | ForEach-Object { [string]$_.SkuId })

    # Step 5: remove a direct assignment that also came via the group only now (otherwise no seat is freed).
    foreach ($Item in $CoveredDirect) {
        if ($ReleasedIds -contains [string]$Item.SkuId) {
            try {
                Set-MgUserLicense -UserId $SourceUser.Id -AddLicenses @() -RemoveLicenses @([Guid]$Item.SkuId) -ErrorAction Stop | Out-Null
                Log-Info ("Direct assignment {0} removed from source." -f $Item.SkuPartNumber)
            } catch {
                $Results.Add([pscustomobject]@{ SkuId = $Item.SkuId; SkuPartNumber = $Item.SkuPartNumber; Status = 'ReleaseIssue'; Detail = "removing the direct assignment from the source failed; the seat remains occupied ($($_.Exception.Message))." })
                $ReleasedIds = @($ReleasedIds | Where-Object { $_ -ne [string]$Item.SkuId })
            }
        } else {
            Log-Info ("Direct assignment {0} on source is kept (source not removed from the group)." -f $Item.SkuPartNumber)
        }
    }

    # Step 6: check whether the new user really has the released licenses.
    $Verify = @{}
    $ToVerify = @($ReleasedSkus | Where-Object { $ReleasedIds -contains [string]$_.SkuId })
    if ($ToVerify.Count -gt 0) {
        $Verify = Complete-GroupLicenseHandover -Plan $Plan -Skus $ToVerify -NewUserId $NewUserId
    }

    foreach ($Item in @($Plan.GroupSkusWithoutSeat)) {
        $Key = [string]$Item.SkuId
        if ($Verify.ContainsKey($Key) -and $Verify[$Key] -eq 'Active') {
            $Results.Add([pscustomobject]@{ SkuId = $Item.SkuId; SkuPartNumber = $Item.SkuPartNumber; Status = 'GroupReleased'; Detail = "via group $($Item.Groups); active for the new user." })
            Log-Info ("Group license active for new user: {0} via group {1}" -f $Item.SkuPartNumber, $Item.Groups)
        } elseif ($Verify.ContainsKey($Key)) {
            $Results.Add([pscustomobject]@{ SkuId = $Item.SkuId; SkuPartNumber = $Item.SkuPartNumber; Status = 'ReprocessPending'; Detail = ("the source has released the seat, but for the new user: {0}. Reprocess later: Invoke-MgLicenseUser -UserId {1}" -f $Verify[$Key], $NewUserId) })
            Log-Warning ("Group license {0} not yet active for new user: {1}" -f $Item.SkuPartNumber, $Verify[$Key])
        } else {
            $Advice = @($Plan.GroupsHoldingSeats | Where-Object { $Item.GroupIds -contains $_.Id } | ForEach-Object {
                if ($_.ReleasableByRemove) { "remove the source from group $($_.DisplayName)" }
                elseif ($_.ReleasableByBlock) { "block the source's sign-in (group $($_.DisplayName) excludes disabled accounts)" }
                else { "group $($_.DisplayName): $($_.ManualReason)" }
            }) -join '; '
            $Results.Add([pscustomobject]@{ SkuId = $Item.SkuId; SkuPartNumber = $Item.SkuPartNumber; Status = 'GroupNoSeat'; Detail = ("no free seat and the source holds it via {0}. To resolve: {1}; then Invoke-MgLicenseUser -UserId {2}, or buy an additional seat." -f $Item.Groups, $Advice, $NewUserId) })
            Log-Warning ("Group license without a free seat: {0} via group {1}" -f $Item.SkuPartNumber, $Item.Groups)
        }
    }

    return $Results
}

# ── Phone number transfer (Teams Phone) ─────────────────────────────────────

function Test-PrimaryNumber {
    param([Parameter(Mandatory)][object]$Number)
    return ([string]::IsNullOrWhiteSpace([string]$Number.AssignmentCategory) -or [string]$Number.AssignmentCategory -eq 'Primary')
}

function Get-NumberTransferPlan {
    param([Parameter(Mandatory)][object]$SourceUser)

    Ensure-TeamsConnection

    $Assignments = @(Get-CsPhoneNumberAssignment -AssignedPstnTargetId ([string]$SourceUser.Id) -ErrorAction Stop)
    $Numbers = foreach ($Assignment in $Assignments) {
        $LocationId = [string](Get-ObjectPropertyValue -InputObject $Assignment -Name 'LocationId')
        if ($LocationId -match '^(0{8}-0{4}-0{4}-0{4}-0{12}|null)$') { $LocationId = '' }
        [pscustomobject]@{
            TelephoneNumber    = [string]$Assignment.TelephoneNumber
            NumberType         = [string]$Assignment.NumberType
            LocationId         = $LocationId
            AssignmentCategory = [string](Get-ObjectPropertyValue -InputObject $Assignment -Name 'AssignmentCategory')
        }
    }

    Log-Info ("Phone numbers on source {0}: {1}" -f $SourceUser.UserPrincipalName, @($Numbers).Count)
    return [pscustomobject]@{
        SourceUser = $SourceUser
        Numbers    = @($Numbers)
    }
}

function Confirm-NumberTransfer {
    # Returns $true to continue, $false to cancel.
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][string]$NewUpn,
        [bool]$WithLicenseTransfer = $false
    )

    $Lines = [System.Collections.Generic.List[string]]::new()
    $SourceUpn = [string]$Plan.SourceUser.UserPrincipalName

    if (@($Plan.Numbers).Count -eq 0) {
        $Lines.Add(("The source {0} has no phone number in Teams. No number will be transferred." -f $SourceUpn))
        $Lines.Add('')
        $Lines.Add('Continue with the rest of the copy?')
    } else {
        $Lines.Add(("Transfer phone number from {0} to {1}:" -f $SourceUpn, $NewUpn))
        foreach ($Number in $Plan.Numbers) {
            $Category = if (Test-PrimaryNumber -Number $Number) { '' } else { ", $($Number.AssignmentCategory)" }
            $Location = if ($Number.LocationId) { 'emergency location is transferred' } else { 'no emergency location' }
            $Lines.Add(("  - {0} ({1}{2}, {3})" -f $Number.TelephoneNumber, $Number.NumberType, $Category, $Location))
        }
        $Lines.Add('')
        $Lines.Add(("As soon as the new user is ready to call in Teams (waits up to {0} minutes) the number is removed from the source, assigned to the new user and Enterprise Voice is enabled. If assignment fails, the number is returned to the source." -f [Math]::Round($TeamsUserWaitTimeoutSeconds / 60)))
        if ($WithLicenseTransfer) {
            $Lines.Add('Note: if the source loses Teams Phone through the license transfer or the phase-out, Teams releases the number earlier. The script then assigns it as soon as the new user is ready; returning it to the source is not possible in that case.')
        }
        $Lines.Add('')
        $Lines.Add('Continue?')
    }

    Log-Info ("Number transfer confirmation requested for source {0}." -f $SourceUpn)
    $Answer = Show-UiQuestion -Message ($Lines -join [Environment]::NewLine) -Buttons 'YesNo'
    return ($Answer -eq 'Yes')
}

function Wait-ForTeamsVoiceUser {
    # Ready as soon as Get-CsOnlineUser knows the user and Teams Phone (PhoneSystem) is active.
    # Stops early if the new user's license is in an error state, or if Teams
    # repeatedly returns an error other than 'not found' (e.g. an expired sign-in).
    param(
        [Parameter(Mandatory)][string]$Upn,
        [int]$TimeoutSeconds = 1800,
        [int]$IntervalSeconds = 30,
        [string]$GraphUserId = ''
    )

    $script:LastWaitBlocker = $null
    Ensure-TeamsConnection
    Log-Info ("Waiting until Teams Phone is active for: {0} (max {1}s, check every {2}s)" -f $Upn, $TimeoutSeconds, $IntervalSeconds)

    $StartTime   = [DateTime]::UtcNow
    $BlockerHits = 0
    $OtherErrors = 0
    while (([DateTime]::UtcNow - $StartTime).TotalSeconds -lt $TimeoutSeconds) {
        $Elapsed = [Math]::Round(([DateTime]::UtcNow - $StartTime).TotalSeconds)
        try {
            $CsUser   = Get-CsOnlineUser -Identity $Upn -ErrorAction Stop
            $OtherErrors = 0
            $Features = @(Get-ObjectPropertyValue -InputObject $CsUser -Name 'FeatureTypes')
            if ($Features -contains 'PhoneSystem') {
                Log-Info ("Teams Phone active after {0}s: {1}" -f $Elapsed, $Upn)
                return $true
            }
            Log-Info ("User visible in Teams, Teams Phone not active yet ({0}s)." -f $Elapsed)
        } catch {
            $Message = [string]$_.Exception.Message
            if ($Message -match 'not found|NotFound|could not be found|niet gevonden|404') {
                $OtherErrors = 0
                Log-Info ("User not yet visible in Teams ({0}s)." -f $Elapsed)
            } else {
                $OtherErrors++
                Log-Warning ("Teams check returned an error ({0}s): {1}" -f $Elapsed, $Message)
                if ($OtherErrors -ge 5) {
                    $script:LastWaitBlocker = "Teams repeatedly returns an error ($Message); check the Teams sign-in"
                    return $false
                }
            }
        }

        if ($GraphUserId -and $Elapsed -ge $LicenseBlockerGraceSeconds) {
            # Only stop if the Teams Phone license (service plan MCOEV*) is in an error state.
            $Blocker = Get-LicenseBlocker -UserId $GraphUserId -ServicePlanPattern 'MCOEV*'
            if ($Blocker) { $BlockerHits++ } else { $BlockerHits = 0 }
            if ($BlockerHits -ge 2) {
                $script:LastWaitBlocker = $Blocker
                Log-Warning ("Stopped waiting for Teams Phone after {0}s: {1}." -f $Elapsed, $Blocker)
                return $false
            }
        }
        Wait-UiSeconds -Seconds $IntervalSeconds
    }

    Log-Warning ("Teams Phone still not active after {0}s for: {1}" -f $TimeoutSeconds, $Upn)
    return $false
}

function Get-NumberHolder {
    # Returns 'Source', 'Free', 'Other' or 'Missing' for a number, based on AssignedPstnTargetId.
    # The number may already be free: if the source loses Teams Phone (e.g. via the license group),
    # Teams releases the number automatically.
    param(
        [Parameter(Mandatory)][string]$TelephoneNumber,
        [Parameter(Mandatory)][string]$SourceId
    )
    $Assignment = Get-CsPhoneNumberAssignment -TelephoneNumber $TelephoneNumber -ErrorAction Stop | Select-Object -First 1
    if ($null -eq $Assignment) { return 'Missing' }
    $Target = [string](Get-ObjectPropertyValue -InputObject $Assignment -Name 'AssignedPstnTargetId')
    if ([string]::IsNullOrWhiteSpace($Target)) { return 'Free' }
    if ($Target -eq $SourceId) { return 'Source' }
    return 'Other'
}

function Move-PhoneNumbers {
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][string]$NewUpn,
        [string]$NewUserId = ''
    )

    $Results   = [System.Collections.Generic.List[object]]::new()
    $SourceId  = [string]$Plan.SourceUser.Id
    $SourceUpn = [string]$Plan.SourceUser.UserPrincipalName
    $Numbers   = @($Plan.Numbers)
    if ($Numbers.Count -eq 0) { return $Results }

    # Manual commands per number (incl. AssignmentCategory for private/alternate lines).
    $ManualAssign = {
        param($Number)
        $Extra = ''
        if ($Number.LocationId) { $Extra += " -LocationId $($Number.LocationId)" }
        if (-not (Test-PrimaryNumber -Number $Number)) { $Extra += " -AssignmentCategory $($Number.AssignmentCategory)" }
        "Set-CsPhoneNumberAssignment -Identity $NewUpn -PhoneNumber $($Number.TelephoneNumber) -PhoneNumberType $($Number.NumberType)$Extra; Set-CsPhoneNumberAssignment -Identity $NewUpn -EnterpriseVoiceEnabled `$true"
    }

    # Only remove anything from the source if the new user can also receive the number.
    $Ready = Wait-ForTeamsVoiceUser -Upn $NewUpn -TimeoutSeconds $TeamsUserWaitTimeoutSeconds -IntervalSeconds $TeamsUserWaitIntervalSeconds -GraphUserId $NewUserId
    if (-not $Ready) {
        $Reason = if ($script:LastWaitBlocker) { $script:LastWaitBlocker } else { ("the new user still had no active Teams Phone after {0} min" -f [Math]::Round($TeamsUserWaitTimeoutSeconds / 60)) }
        foreach ($Number in $Numbers) {
            $Holder = 'Source'
            try { $Holder = Get-NumberHolder -TelephoneNumber $Number.TelephoneNumber -SourceId $SourceId } catch {}
            $Where  = switch ($Holder) {
                'Free'    { 'The number has already been released from the source (it is free).' }
                'Other'   { 'NOTE: the number is now assigned to another user.' }
                'Missing' { 'NOTE: the number was not found (anymore) in the tenant.' }
                default   { 'The number is still assigned to the source.' }
            }
            $Remove = if ($Holder -eq 'Source') { "Remove-CsPhoneNumberAssignment -Identity $SourceUpn -PhoneNumber $($Number.TelephoneNumber) -PhoneNumberType $($Number.NumberType); " } else { '' }
            $Results.Add([pscustomobject]@{
                TelephoneNumber = $Number.TelephoneNumber
                Status          = 'NotReady'
                Detail          = ("not transferred: {0}. {1} Later manually: {2}{3}" -f $Reason, $Where, $Remove, (& $ManualAssign $Number))
            })
        }
        return $Results
    }

    # Can the source still get the number back? Not without Teams Phone.
    $SourceHasPhone = $true
    try {
        $SourceCs = Get-CsOnlineUser -Identity $SourceId -ErrorAction Stop
        $SourceHasPhone = @(Get-ObjectPropertyValue -InputObject $SourceCs -Name 'FeatureTypes') -contains 'PhoneSystem'
    } catch { }

    # Step 1: remove from the source (non-primary numbers first). Already free = assign directly;
    # assigned to someone else or not found = leave untouched.
    $Removed    = [System.Collections.Generic.List[object]]::new()
    $NoRollback = @{}
    foreach ($Number in @($Numbers | Sort-Object { Test-PrimaryNumber -Number $_ })) {
        $Holder = 'Source'
        try { $Holder = Get-NumberHolder -TelephoneNumber $Number.TelephoneNumber -SourceId $SourceId } catch {
            Log-Warning ("Status of number {0} could not be retrieved; assuming source. {1}" -f $Number.TelephoneNumber, $_.Exception.Message)
        }
        if ($Holder -eq 'Other' -or $Holder -eq 'Missing') {
            $Detail = if ($Holder -eq 'Other') { 'is now assigned to another user; not transferred.' } else { 'not found (anymore) in the tenant; not transferred.' }
            $Results.Add([pscustomobject]@{ TelephoneNumber = $Number.TelephoneNumber; Status = 'Failed'; Detail = $Detail })
            Log-Warning ("Number {0}: {1}" -f $Number.TelephoneNumber, $Detail)
            continue
        }
        if ($Holder -eq 'Free') {
            $Removed.Add($Number)
            $NoRollback[$Number.TelephoneNumber] = 'the number had already been released from the source'
            Log-Info ("Number {0} had already been released from the source; assigning directly." -f $Number.TelephoneNumber)
            continue
        }
        try {
            Remove-CsPhoneNumberAssignment -Identity $SourceId -PhoneNumber $Number.TelephoneNumber -PhoneNumberType $Number.NumberType -ErrorAction Stop | Out-Null
            $Removed.Add($Number)
            if (-not $SourceHasPhone) { $NoRollback[$Number.TelephoneNumber] = 'the source no longer has Teams Phone' }
            Log-Info ("Number {0} removed from source {1}." -f $Number.TelephoneNumber, $SourceUpn)
        } catch {
            $Message = $_.Exception.Message
            $Recheck = 'Source'
            try { $Recheck = Get-NumberHolder -TelephoneNumber $Number.TelephoneNumber -SourceId $SourceId } catch {}
            if ($Recheck -eq 'Free') {
                $Removed.Add($Number)
                $NoRollback[$Number.TelephoneNumber] = 'the number was released by Teams in the meantime'
                Log-Info ("Number {0} released in the meantime; assigning directly." -f $Number.TelephoneNumber)
            } else {
                $Where = switch ($Recheck) { 'Other' { 'the number is now assigned to another user' } 'Missing' { 'the number was not found (anymore) in the tenant' } default { 'the number is still assigned to the source' } }
                $Results.Add([pscustomobject]@{ TelephoneNumber = $Number.TelephoneNumber; Status = 'Failed'; Detail = "removal from the source failed; $Where ($Message)." })
                Log-Warning ("Number {0} could not be removed from the source: {1}" -f $Number.TelephoneNumber, $Message)
            }
        }
    }

    # Step 2: assign to the new user (primary number first), with retries.
    # If the primary number fails, the others are not assigned but returned.
    $PrimaryFailed = $false
    foreach ($Number in @($Removed | Sort-Object { -not (Test-PrimaryNumber -Number $_) })) {
        $AssignParams = @{
            Identity        = $NewUpn
            PhoneNumber     = $Number.TelephoneNumber
            PhoneNumberType = $Number.NumberType
            ErrorAction     = 'Stop'
        }
        if ($Number.LocationId) { $AssignParams.LocationId = $Number.LocationId }
        if (-not (Test-PrimaryNumber -Number $Number)) { $AssignParams.AssignmentCategory = $Number.AssignmentCategory }

        $Assigned  = $false
        $LastError = ''
        if ($PrimaryFailed) {
            $LastError = 'not assigned because the primary number could not be assigned'
        } else {
            for ($Attempt = 1; $Attempt -le $NumberAssignRetryCount; $Attempt++) {
                try {
                    Set-CsPhoneNumberAssignment @AssignParams | Out-Null
                    $Assigned = $true
                    break
                } catch {
                    $LastError = $_.Exception.Message
                    Log-Info ("Assigning number {0} not successful yet (attempt {1}/{2}): {3}" -f $Number.TelephoneNumber, $Attempt, $NumberAssignRetryCount, $LastError)
                    if ($Attempt -lt $NumberAssignRetryCount) { Wait-UiSeconds -Seconds $NumberAssignRetrySeconds }
                }
            }
        }

        if ($Assigned) {
            $Results.Add([pscustomobject]@{ TelephoneNumber = $Number.TelephoneNumber; Status = 'Assigned'; Detail = '' })
            Log-Info ("Number {0} assigned to {1}." -f $Number.TelephoneNumber, $NewUpn)
            continue
        }
        if (Test-PrimaryNumber -Number $Number) { $PrimaryFailed = $true }

        if ($NoRollback.ContainsKey($Number.TelephoneNumber)) {
            $Results.Add([pscustomobject]@{ TelephoneNumber = $Number.TelephoneNumber; Status = 'Failed'; Detail = ("NOT assigned and the number IS FREE ({0}; returning it to the source is not possible). Manually: {1} ({2})" -f $NoRollback[$Number.TelephoneNumber], (& $ManualAssign $Number), $LastError) })
            Log-Warning ("Number {0} not assigned; it is free. {1}" -f $Number.TelephoneNumber, $LastError)
            continue
        }

        # Step 3: back to the source so the number is not left free and unmanaged.
        $AssignParams.Identity = $SourceId
        try {
            Set-CsPhoneNumberAssignment @AssignParams | Out-Null
            $Results.Add([pscustomobject]@{ TelephoneNumber = $Number.TelephoneNumber; Status = 'RolledBack'; Detail = "not assigned, returned to the source ($LastError)." })
            Log-Warning ("Number {0} not assigned; returned to source. {1}" -f $Number.TelephoneNumber, $LastError)
        } catch {
            $RollbackError = $_.Exception.Message
            $Results.Add([pscustomobject]@{ TelephoneNumber = $Number.TelephoneNumber; Status = 'RollbackFailed'; Detail = ("NOT assigned and NOT returned to the source; the number IS FREE. Manually: {0} ({1} / {2})" -f (& $ManualAssign $Number), $LastError, $RollbackError) })
            Log-Error ("Number {0} not assigned and returning it to the source failed. {1} / {2}" -f $Number.TelephoneNumber, $LastError, $RollbackError)
        }
    }

    if (-not @($Results | Where-Object { $_.Status -eq 'Assigned' })) { return $Results }

    # Step 4: explicitly enable Enterprise Voice. This should happen automatically when
    # a number is assigned, but in practice it turned out not to always happen.
    try {
        Set-CsPhoneNumberAssignment -Identity $NewUpn -EnterpriseVoiceEnabled $true -ErrorAction Stop | Out-Null
        Log-Info ("Enterprise Voice enabled for {0}." -f $NewUpn)
    } catch {
        Log-Warning ("Enabling Enterprise Voice failed for {0}: {1}" -f $NewUpn, $_.Exception.Message)
    }

    $VoiceEnabled = $false
    $StartTime    = [DateTime]::UtcNow
    while ($true) {
        try {
            $CsUser = Get-CsOnlineUser -Identity $NewUpn -ErrorAction Stop
            $VoiceEnabled = ([string](Get-ObjectPropertyValue -InputObject $CsUser -Name 'EnterpriseVoiceEnabled') -eq 'True')
            if ($VoiceEnabled) {
                Log-Info ("Teams check: EnterpriseVoiceEnabled=True, LineUri={0}" -f (Get-ObjectPropertyValue -InputObject $CsUser -Name 'LineUri'))
                break
            }
        } catch {
            Log-Info ("Enterprise Voice check not possible yet: {0}" -f $_.Exception.Message)
        }
        if (([DateTime]::UtcNow - $StartTime).TotalSeconds -ge $VoiceCheckTimeoutSeconds) { break }
        Wait-UiSeconds -Seconds 15
    }

    if (-not $VoiceEnabled) {
        $Results.Add([pscustomobject]@{
            TelephoneNumber = ''
            Status          = 'VoiceNotEnabled'
            Detail          = ("Enterprise Voice was still off at the check. Check again in fifteen minutes or enable it with: Set-CsPhoneNumberAssignment -Identity {0} -EnterpriseVoiceEnabled `$true" -f $NewUpn)
        })
        Log-Warning ("Enterprise Voice still off at the check for {0}." -f $NewUpn)
    }

    return $Results
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:Form               = New-Object System.Windows.Forms.Form
$script:Form.Text          = 'Copy Microsoft 365 user'
$script:Form.StartPosition = 'CenterScreen'
$script:Form.Size          = New-Object System.Drawing.Size(760, 725)
$script:Form.MinimumSize   = New-Object System.Drawing.Size(760, 725)

$lblSearch          = New-Object System.Windows.Forms.Label
$lblSearch.Text     = 'Search source user'
$lblSearch.Location = New-Object System.Drawing.Point(15, 15)
$lblSearch.Size     = New-Object System.Drawing.Size(160, 20)
$script:Form.Controls.Add($lblSearch)

$txtSearch          = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(15, 38)
$txtSearch.Size     = New-Object System.Drawing.Size(500, 24)
$script:Form.Controls.Add($txtSearch)

$btnSearch          = New-Object System.Windows.Forms.Button
$btnSearch.Text     = 'Search'
$btnSearch.Location = New-Object System.Drawing.Point(525, 36)
$btnSearch.Size     = New-Object System.Drawing.Size(90, 28)
$script:Form.Controls.Add($btnSearch)

$cmbResults               = New-Object System.Windows.Forms.ComboBox
$cmbResults.Location      = New-Object System.Drawing.Point(15, 72)
$cmbResults.Size          = New-Object System.Drawing.Size(710, 24)
$cmbResults.DropDownStyle = 'DropDownList'
$script:Form.Controls.Add($cmbResults)

$grpNewUser          = New-Object System.Windows.Forms.GroupBox
$grpNewUser.Text     = 'New user'
$grpNewUser.Location = New-Object System.Drawing.Point(15, 110)
$grpNewUser.Size     = New-Object System.Drawing.Size(710, 200)
$script:Form.Controls.Add($grpNewUser)

$labels = @(
    @{ Text = 'First name';  Name = 'txtFirstName';   Y = 30  },
    @{ Text = 'Last name';   Name = 'txtLastName';    Y = 65  },
    @{ Text = 'Display name'; Name = 'txtDisplayName'; Y = 100 },
    @{ Text = 'New UPN';     Name = 'txtNewUpn';      Y = 135 },
    @{ Text = 'Password';    Name = 'txtPassword';    Y = 170 }
)

$InputControls = @{}

foreach ($item in $labels) {
    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = $item.Text
    $lbl.Location = New-Object System.Drawing.Point(15, [int]$item.Y)
    $lbl.Size     = New-Object System.Drawing.Size(120, 20)
    $grpNewUser.Controls.Add($lbl)

    $txt          = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(150, ([int]$item.Y - 2))

    if ($item.Name -eq 'txtPassword') {
        $txt.Size = New-Object System.Drawing.Size(290, 24)
        $txt.UseSystemPasswordChar = $false
        $txt.PasswordChar = [char]'*'
    } else {
        $txt.Size = New-Object System.Drawing.Size(530, 24)
    }

    $grpNewUser.Controls.Add($txt)
    $InputControls[$item.Name] = $txt
}

$btnRandomPassword          = New-Object System.Windows.Forms.Button
$btnRandomPassword.Text     = 'Random password'
$btnRandomPassword.Location = New-Object System.Drawing.Point(450, 166)
$btnRandomPassword.Size     = New-Object System.Drawing.Size(145, 28)
$grpNewUser.Controls.Add($btnRandomPassword)

$btnShowPassword          = New-Object System.Windows.Forms.Button
$btnShowPassword.Text     = [char]0xE890
$btnShowPassword.Location = New-Object System.Drawing.Point(605, 166)
$btnShowPassword.Size     = New-Object System.Drawing.Size(35, 28)
$btnShowPassword.Font     = New-Object System.Drawing.Font('Segoe MDL2 Assets', 10)
$grpNewUser.Controls.Add($btnShowPassword)

$btnCopyPassword          = New-Object System.Windows.Forms.Button
$btnCopyPassword.Text     = [char]0xE8C8
$btnCopyPassword.Location = New-Object System.Drawing.Point(645, 166)
$btnCopyPassword.Size     = New-Object System.Drawing.Size(35, 28)
$btnCopyPassword.Font     = New-Object System.Drawing.Font('Segoe MDL2 Assets', 10)
$grpNewUser.Controls.Add($btnCopyPassword)

$PasswordToolTip = New-Object System.Windows.Forms.ToolTip
$PasswordToolTip.SetToolTip($btnRandomPassword, 'Generate a strong random password.')
$PasswordToolTip.SetToolTip($btnShowPassword, 'Show or hide password.')
$PasswordToolTip.SetToolTip($btnCopyPassword, 'Copy password to clipboard.')

$grpOptions          = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text     = 'Options'
$grpOptions.Location = New-Object System.Drawing.Point(15, 320)
$grpOptions.Size     = New-Object System.Drawing.Size(710, 95)
$script:Form.Controls.Add($grpOptions)

$chkMailbox          = New-Object System.Windows.Forms.CheckBox
$chkMailbox.Text     = 'Copy mailbox permissions'
$chkMailbox.Location = New-Object System.Drawing.Point(15, 30)
$chkMailbox.Size     = New-Object System.Drawing.Size(220, 24)
$grpOptions.Controls.Add($chkMailbox)

$chkLicense          = New-Object System.Windows.Forms.CheckBox
$chkLicense.Text     = 'Copy licenses'
$chkLicense.Location = New-Object System.Drawing.Point(250, 30)
$chkLicense.Size     = New-Object System.Drawing.Size(180, 24)
$grpOptions.Controls.Add($chkLicense)

$chkTransfer          = New-Object System.Windows.Forms.CheckBox
$chkTransfer.Text     = 'Transfer licenses from source'
$chkTransfer.Location = New-Object System.Drawing.Point(450, 30)
$chkTransfer.Size     = New-Object System.Drawing.Size(250, 24)
$grpOptions.Controls.Add($chkTransfer)

$chkNumber          = New-Object System.Windows.Forms.CheckBox
$chkNumber.Text     = 'Transfer phone number from source (Teams)'
$chkNumber.Location = New-Object System.Drawing.Point(15, 58)
$chkNumber.Size     = New-Object System.Drawing.Size(400, 24)
$grpOptions.Controls.Add($chkNumber)

$OptionsToolTip = New-Object System.Windows.Forms.ToolTip
$OptionsToolTip.SetToolTip($chkNumber, 'The source''s Teams phone number moves to the new user as soon as they have Teams Phone. Enterprise Voice is enabled as part of this. Requires the MicrosoftTeams module.')
$OptionsToolTip.SetToolTip($chkTransfer, 'For a source who has left the organization: directly assigned licenses are removed from the source and assigned to the new user. An overview for confirmation is shown first.')

# Transferring is a way of taking licenses along: both checkboxes belong together.
$chkTransfer.Add_CheckedChanged({
    if ($chkTransfer.Checked) { $chkLicense.Checked = $true }
})
$chkLicense.Add_CheckedChanged({
    if (-not $chkLicense.Checked) { $chkTransfer.Checked = $false }
})

$btnCreate          = New-Object System.Windows.Forms.Button
$btnCreate.Text     = 'Start copy'
$btnCreate.Location = New-Object System.Drawing.Point(15, 425)
$btnCreate.Size     = New-Object System.Drawing.Size(120, 34)
$script:Form.Controls.Add($btnCreate)

$btnPrefill          = New-Object System.Windows.Forms.Button
$btnPrefill.Text     = 'Fill from source'
$btnPrefill.Location = New-Object System.Drawing.Point(145, 425)
$btnPrefill.Size     = New-Object System.Drawing.Size(140, 34)
$script:Form.Controls.Add($btnPrefill)

$script:StatusLabel          = New-Object System.Windows.Forms.Label
$script:StatusLabel.Text     = 'Ready'
$script:StatusLabel.Location = New-Object System.Drawing.Point(300, 433)
$script:StatusLabel.Size     = New-Object System.Drawing.Size(425, 20)
$script:Form.Controls.Add($script:StatusLabel)

$script:LogBox            = New-Object System.Windows.Forms.TextBox
$script:LogBox.Location   = New-Object System.Drawing.Point(15, 470)
$script:LogBox.Size       = New-Object System.Drawing.Size(710, 200)
$script:LogBox.Multiline  = $true
$script:LogBox.ScrollBars = 'Vertical'
$script:LogBox.ReadOnly   = $true
$script:LogBox.Font       = New-Object System.Drawing.Font('Consolas', 9)
$script:Form.Controls.Add($script:LogBox)

$script:ControlsToToggle = @(
    $txtSearch, $btnSearch, $cmbResults,
    $btnCreate, $btnPrefill, $btnRandomPassword, $btnShowPassword, $btnCopyPassword, $chkMailbox, $chkLicense, $chkTransfer, $chkNumber,
    $InputControls['txtFirstName'],
    $InputControls['txtLastName'],
    $InputControls['txtDisplayName'],
    $InputControls['txtNewUpn'],
    $InputControls['txtPassword']
)


function Update-NewUserSuggestions {
    if (-not $InputControls) { return }

    $FirstName = ([string]$InputControls['txtFirstName'].Text).Trim()
    $LastName  = ([string]$InputControls['txtLastName'].Text).Trim()

    $SuggestedDisplayName = ("{0} {1}" -f $FirstName, $LastName).Trim()

    $DomainPart = ''
    if ($script:SelectedSourceUser -and
        ([string]$script:SelectedSourceUser.UserPrincipalName) -match '@(.+)$') {
        $DomainPart = $Matches[1]
    }

    $SuggestedLocal = ("{0}.{1}" -f $FirstName, $LastName).Trim('.').ToLower()
    $SuggestedLocal = $SuggestedLocal -replace '\s+', '.'
    $SuggestedLocal = $SuggestedLocal -replace '[^a-z0-9._-]', ''
    $SuggestedLocal = $SuggestedLocal -replace '\.{2,}', '.'
    $SuggestedLocal = $SuggestedLocal.Trim('.')

    $script:UpdatingSuggestions = $true
    try {
        if ($script:AutoUpdateDisplayName) {
            $InputControls['txtDisplayName'].Text = $SuggestedDisplayName
        }

        if ($script:AutoUpdateUpn) {
            if (-not [string]::IsNullOrWhiteSpace($SuggestedLocal) -and
                -not [string]::IsNullOrWhiteSpace($DomainPart)) {
                $InputControls['txtNewUpn'].Text = "${SuggestedLocal}@${DomainPart}"
            } else {
                $InputControls['txtNewUpn'].Text = ''
            }
        }
    } finally {
        $script:UpdatingSuggestions = $false
    }
}

$InputControls['txtFirstName'].Add_TextChanged({
    Update-NewUserSuggestions
})

$InputControls['txtLastName'].Add_TextChanged({
    Update-NewUserSuggestions
})

$InputControls['txtDisplayName'].Add_TextChanged({
    if (-not $script:UpdatingSuggestions) {
        $script:AutoUpdateDisplayName = $false
    }
})

$InputControls['txtNewUpn'].Add_TextChanged({
    if (-not $script:UpdatingSuggestions) {
        $script:AutoUpdateUpn = $false
    }
})

$btnRandomPassword.Add_Click({
    # A newly generated password is always shown masked first.
    $script:PasswordVisible = $false
    $InputControls['txtPassword'].UseSystemPasswordChar = $false
    $InputControls['txtPassword'].PasswordChar = [char]'*'
    $PasswordToolTip.SetToolTip($btnShowPassword, 'Show password.')

    $InputControls['txtPassword'].Text = New-RandomPassword -Length 18
    $InputControls['txtPassword'].SelectionStart = $InputControls['txtPassword'].TextLength
    Log-Info 'Random initial password generated.'
})

$btnShowPassword.Add_Click({
    $script:PasswordVisible = -not $script:PasswordVisible
    $InputControls['txtPassword'].UseSystemPasswordChar = $false

    if ($script:PasswordVisible) {
        $InputControls['txtPassword'].PasswordChar = [char]0
        $PasswordToolTip.SetToolTip($btnShowPassword, 'Hide password.')
    } else {
        $InputControls['txtPassword'].PasswordChar = [char]'*'
        $PasswordToolTip.SetToolTip($btnShowPassword, 'Show password.')
    }

    $InputControls['txtPassword'].Focus()
    $InputControls['txtPassword'].SelectionStart = $InputControls['txtPassword'].TextLength
})

$btnCopyPassword.Add_Click({
    $PasswordValue = [string]$InputControls['txtPassword'].Text

    if ([string]::IsNullOrWhiteSpace($PasswordValue)) {
        Show-UiWarning 'There is no password in the field yet.'
        return
    }

    [System.Windows.Forms.Clipboard]::SetText($PasswordValue)
    Log-Info 'Initial password copied to the Windows clipboard.'
    $script:StatusLabel.Text = 'Password copied'
})

$btnSearch.Add_Click({
    $SearchTerm = $txtSearch.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
        Show-UiWarning 'Enter a search term.'
        return
    }

    try {
        Set-UiBusyState -Busy $true -StatusText 'Signing in to Graph...'
        [System.Windows.Forms.Application]::DoEvents()
        Ensure-GraphConnectionInteractive

        Set-UiBusyState -Busy $true -StatusText 'Searching...'
        [System.Windows.Forms.Application]::DoEvents()
        Log-Info ("Searching for users with term: {0}" -f $SearchTerm)

        $Users = @(Search-SourceUsers -SearchTerm $SearchTerm)

        $cmbResults.Items.Clear()
        $script:SourceUserResults.Clear()
        $script:SelectedSourceUser = $null

        foreach ($User in $Users) {
            [void]$script:SourceUserResults.Add($User)
            [void]$cmbResults.Items.Add(('{0} <{1}>' -f $User.DisplayName, $User.UserPrincipalName))
        }

        if ($Users.Count -eq 0) {
            Log-Warning 'No users found.'
        } else {
            $cmbResults.SelectedIndex = 0
            Log-Info ("Number of users found: {0}" -f $Users.Count)
        }
    } catch {
        Show-UiError $_.Exception.Message
    } finally {
        Set-UiBusyState -Busy $false -StatusText 'Ready'
    }
})

$cmbResults.Add_SelectedIndexChanged({
    $SelectedIndex = $cmbResults.SelectedIndex
    if ($SelectedIndex -lt 0 -or $SelectedIndex -ge $script:SourceUserResults.Count) {
        $script:SelectedSourceUser = $null
        return
    }

    $script:SelectedSourceUser = $script:SourceUserResults[$SelectedIndex]
    Update-NewUserSuggestions
    Log-Info ("Source user selected: {0}" -f $script:SelectedSourceUser.UserPrincipalName)
})

$btnPrefill.Add_Click({
    if (-not $script:SelectedSourceUser) {
        Show-UiError 'Select a source user first.'
        return
    }

    # "Fill from source" turns the automatic suggestions back on.
    # As soon as DisplayName or UPN is then changed manually, the
    # automatic completion leaves that field alone.
    $script:AutoUpdateDisplayName = $true
    $script:AutoUpdateUpn         = $true

    $InputControls['txtFirstName'].Text = [string]$script:SelectedSourceUser.GivenName
    $InputControls['txtLastName'].Text  = [string]$script:SelectedSourceUser.Surname

    Update-NewUserSuggestions
    Log-Info 'Source details filled in; DisplayName and UPN suggested automatically.'
})

function Invoke-CopyRun {
    if (-not $script:SelectedSourceUser) {
        Show-UiError 'Select a source user first.'
        return
    }

    # Read the form fields on the UI thread before we disable the controls
    $SourceUpn   = $script:SelectedSourceUser.UserPrincipalName
    $FirstName   = $InputControls['txtFirstName'].Text
    $LastName    = $InputControls['txtLastName'].Text
    $DisplayName = $InputControls['txtDisplayName'].Text
    $NewUpn      = $InputControls['txtNewUpn'].Text
    $Password    = $InputControls['txtPassword'].Text
    $CopyMailbox = $chkMailbox.Checked
    $TransferLicense = $chkTransfer.Checked
    $CopyLicense     = $chkLicense.Checked -or $TransferLicense

    $TransferNumber  = $chkNumber.Checked

    $TransferPlan     = $null
    $TransferDecision = $null
    $NumberPlan       = $null
    $NumberResults    = @()
    $Validated      = $null
    $Created        = $null
    $GroupResults   = @()
    $LicenseResults = @()
    $DistributionResults = @()
    $MailboxResults      = @()
    $RunIssues           = [System.Collections.Generic.List[string]]::new()

    # ── Phase 1: connection, validation and user creation ───────────────────
    # Up to and including New-MgUser an error is fatal: there are no follow-up actions yet
    # that would have to be reported as Partial.
    try {
        Set-UiBusyState -Busy $true -StatusText 'Signing in to Graph...'
        [System.Windows.Forms.Application]::DoEvents()
        Ensure-GraphConnectionInteractive

        Set-UiBusyState -Busy $true -StatusText 'Validating input...'
        [System.Windows.Forms.Application]::DoEvents()
        $Validated = Test-NewUserInput `
            -FirstName   $FirstName `
            -LastName    $LastName `
            -DisplayName $DisplayName `
            -NewUpn      $NewUpn `
            -Password    $Password

        # Have the transfer confirmed up front, so that cancelling has not created anything yet.
        if ($TransferLicense) {
            Set-UiBusyState -Busy $true -StatusText 'Preparing license transfer...'
            [System.Windows.Forms.Application]::DoEvents()
            $TransferPlan = Get-LicenseTransferPlan -SourceUpn $SourceUpn

            Set-UiBusyState -Busy $false -StatusText 'Waiting for confirmation...'
            $TransferDecision = Confirm-LicenseTransfer -Plan $TransferPlan -NewUpn $Validated.NewUpn
            if (-not $TransferDecision) {
                Log-Info 'Cancelled at the license transfer confirmation; nothing has been created or changed.'
                Set-UiBusyState -Busy $false -StatusText 'Ready'
                return
            }
        }

        if ($TransferNumber) {
            Set-UiBusyState -Busy $true -StatusText 'Looking up phone number of source...'
            [System.Windows.Forms.Application]::DoEvents()
            $SourceForNumber = if ($TransferPlan) { $TransferPlan.SourceUser } else { Get-UserByUpn -UserPrincipalName $SourceUpn }
            $NumberPlan = Get-NumberTransferPlan -SourceUser $SourceForNumber

            Set-UiBusyState -Busy $false -StatusText 'Waiting for confirmation...'
            if (-not (Confirm-NumberTransfer -Plan $NumberPlan -NewUpn $Validated.NewUpn -WithLicenseTransfer:$TransferLicense)) {
                Log-Info 'Cancelled at the number transfer confirmation; nothing has been created or changed.'
                Set-UiBusyState -Busy $false -StatusText 'Ready'
                return
            }
        }

        Set-UiBusyState -Busy $true -StatusText 'Creating user...'
        [System.Windows.Forms.Application]::DoEvents()
        $Created = New-CopiedM365User `
            -SourceUpn            $SourceUpn `
            -FirstName            $Validated.FirstName `
            -LastName             $Validated.LastName `
            -DisplayName          $Validated.DisplayName `
            -NewUpn               $Validated.NewUpn `
            -Password             $Password `
            -RequireUsageLocation:$CopyLicense
    } catch {
        Set-UiBusyState -Busy $false -StatusText 'Ready'
        Show-UiError $_.Exception.Message
        return
    }

    # Right after New-MgUser, Graph sometimes does not yet know the user in other calls.
    Set-UiBusyState -Busy $true -StatusText 'Waiting until new user can be found...'
    $null = Wait-ForGraphUser -UserId $Created.NewUser.Id -TimeoutSeconds $GraphUserWaitSeconds

    # ── Phase 2: groups ─────────────────────────────────────────────────────
    try {
        Set-UiBusyState -Busy $true -StatusText 'Copying groups...'
        [System.Windows.Forms.Application]::DoEvents()
        $GroupResults = @(Copy-GroupMemberships `
            -SourceUserId $Created.SourceUser.Id `
            -NewUserId    $Created.NewUser.Id)
    } catch {
        $Message = $_.Exception.Message
        Log-Error ("Group copy failed completely: {0}" -f $Message)
        $RunIssues.Add('Group memberships could not be processed completely.')
    }

    $GroupAddedGraph       = @($GroupResults | Where-Object { $_.Status -eq 'Added' }).Count
    $GroupPendingExchange  = @($GroupResults | Where-Object { $_.Status -eq 'PendingExchange' })
    $GroupDynamic          = @($GroupResults | Where-Object { $_.Status -eq 'SkippedDynamic' }).Count
    $GroupAttention        = @($GroupResults | Where-Object { $_.Status -in @('Failed','SkippedReadOnly','SkippedOnPremises','SkippedRoleAssignable') }).Count

    Log-Info ("Groups added via Graph: {0}" -f $GroupAddedGraph)
    if ($GroupPendingExchange.Count -gt 0) { Log-Info ("Distribution/mail-enabled security groups for Exchange phase: {0}" -f $GroupPendingExchange.Count) }
    if ($GroupDynamic -gt 0)                { Log-Info ("Dynamic groups skipped: {0}" -f $GroupDynamic) }
    if ($GroupAttention -gt 0) {
        Log-Warning ("Groups requiring manual review before Exchange phase: {0}" -f $GroupAttention)
        foreach ($ProblemGroup in @($GroupResults | Where-Object { $_.Status -in @('Failed','SkippedReadOnly','SkippedOnPremises','SkippedRoleAssignable') })) {
            $RunIssues.Add(("Group {0}: {1}" -f $ProblemGroup.DisplayName, $ProblemGroup.Detail))
        }
    }

    $script:CopiedGroupResults.Clear()
    foreach ($r in $GroupResults) { $script:CopiedGroupResults.Add($r) }

    # ── Phase 3: directly assigned licenses only ────────────────────────────
    if ($TransferLicense) {
        try {
            Set-UiBusyState -Busy $true -StatusText 'Transferring licenses from source...'
            [System.Windows.Forms.Application]::DoEvents()
            $LicenseResults = @(Move-UserLicenses `
                -Plan           $TransferPlan `
                -NewUserId      $Created.NewUser.Id `
                -ConvertMailbox:$TransferDecision.ConvertMailbox `
                -ReleaseGroups:$TransferDecision.ReleaseGroups)
        } catch {
            $Message = $_.Exception.Message
            Log-Error ("License transfer failed: {0}" -f $Message)
            $RunIssues.Add(("Licenses not transferred: {0}" -f $Message))
        }

        $LicenseMoved = @($LicenseResults | Where-Object { $_.Status -eq 'Transferred' }).Count
        if ($LicenseMoved -gt 0) { Log-Info ("License SKUs transferred: {0}" -f $LicenseMoved) }
        foreach ($LicenseProblem in @($LicenseResults | Where-Object { $_.Status -in @('Failed','SkippedSourceError','RolledBack','RollbackFailed','GroupNoSeat','ReprocessPending','ReleaseIssue') })) {
            $LicenseName = [string]$LicenseProblem.SkuPartNumber
            switch ($LicenseProblem.Status) {
                'RollbackFailed'     { $RunIssues.Add(("License {0}: NOT assigned and NOT returned to the source. Assign manually!" -f $LicenseName)) }
                'RolledBack'         { $RunIssues.Add(("License {0}: assignment failed, returned to the source." -f $LicenseName)) }
                'SkippedSourceError' { $RunIssues.Add(("License {0}: source license has an error status and was not transferred." -f $LicenseName)) }
                'GroupNoSeat'        { $RunIssues.Add(("License {0}: {1}" -f $LicenseName, $LicenseProblem.Detail)) }
                'ReleaseIssue'       { $RunIssues.Add(("Phasing out source: {0}" -f $LicenseProblem.Detail)) }
                'ReprocessPending'   { $RunIssues.Add(("License {0}: {1}" -f $LicenseName, $LicenseProblem.Detail)) }
                default              { $RunIssues.Add(("License {0}: {1}" -f $LicenseName, $LicenseProblem.Detail)) }
            }
        }
    } elseif ($CopyLicense) {
        try {
            Set-UiBusyState -Busy $true -StatusText 'Copying licenses...'
            [System.Windows.Forms.Application]::DoEvents()
            $LicenseResults = @(Copy-UserLicenses `
                -SourceUpn        $SourceUpn `
                -NewUserId        $Created.NewUser.Id `
                -SourceUserObject $Created.SourceUser)
        } catch {
            $Message = $_.Exception.Message
            Log-Error ("License copy failed completely: {0}" -f $Message)
            $RunIssues.Add('Direct licenses could not be processed completely.')
        }

        $LicenseAssigned = @($LicenseResults | Where-Object { $_.Status -eq 'Assigned' }).Count
        $LicenseProblems = @($LicenseResults | Where-Object { $_.Status -in @('Failed','SkippedSourceError','NoAvailable') })
        if ($LicenseAssigned -gt 0) { Log-Info ("Direct license SKUs assigned: {0}" -f $LicenseAssigned) }
        if ($LicenseProblems.Count -gt 0) {
            Log-Warning ("Direct licenses needing attention: {0}" -f $LicenseProblems.Count)
            foreach ($LicenseProblem in $LicenseProblems) {
                $LicenseName = if (-not [string]::IsNullOrWhiteSpace([string]$LicenseProblem.SkuPartNumber)) { [string]$LicenseProblem.SkuPartNumber } else { [string]$LicenseProblem.SkuId }
                if ($LicenseProblem.Status -eq 'NoAvailable') {
                    $RunIssues.Add(("License {0}: no free license available." -f $LicenseName))
                } elseif ($LicenseProblem.Status -eq 'SkippedSourceError') {
                    $RunIssues.Add(("License {0}: source license has an error status and was not copied." -f $LicenseName))
                } else {
                    $RunIssues.Add(("License {0}: assignment failed. See the log for details." -f $LicenseName))
                }
            }
        }
    }

    # ── Phase 4: Exchange groups and optional mailbox permissions ─────────────
    # An Exchange sign-in is also needed when the source is a member of a
    # distribution group or mail-enabled security group. These group types are
    # read-only via Microsoft Graph and are processed here via Exchange.
    $NeedsExchange = ($GroupPendingExchange.Count -gt 0 -or $CopyMailbox)

    if ($NeedsExchange) {
        try {
            Set-UiBusyState -Busy $true -StatusText 'Connecting to Exchange Online...'
            [System.Windows.Forms.Application]::DoEvents()
            Ensure-ExchangeConnection

            Set-UiBusyState -Busy $true -StatusText 'Waiting until new user is visible in Exchange...'
            [System.Windows.Forms.Application]::DoEvents()
            Wait-ForUserInExchange -Upn $Validated.NewUpn -TimeoutSeconds $UserWaitTimeoutSeconds

            if ($GroupPendingExchange.Count -gt 0) {
                Set-UiBusyState -Busy $true -StatusText 'Copying distribution groups...'
                [System.Windows.Forms.Application]::DoEvents()

                $DistributionResults = @(Copy-ExchangeGroupMemberships `
                    -Groups $GroupPendingExchange `
                    -NewUpn $Validated.NewUpn)

                $DistributionAdded   = @($DistributionResults | Where-Object { $_.Status -eq 'Added' }).Count
                $DistributionProblem = @($DistributionResults | Where-Object { $_.Status -in @('Failed','SkippedOnPremises') }).Count

                if ($DistributionAdded -gt 0) {
                    Log-Info ("Distribution/mail-enabled security groups added via Exchange: {0}" -f $DistributionAdded)
                }
                if ($DistributionProblem -gt 0) {
                    Log-Warning ("Mail-enabled groups needing attention: {0}" -f $DistributionProblem)
                    foreach ($DistributionIssue in @($DistributionResults | Where-Object { $_.Status -in @('Failed','SkippedOnPremises') })) {
                        $RunIssues.Add(("Mail-enabled group {0}: {1}" -f $DistributionIssue.DisplayName, $DistributionIssue.Detail))
                    }
                }

                foreach ($r in $DistributionResults) { $script:CopiedGroupResults.Add($r) }
            }

            $MailboxReady = $false
            if ($CopyMailbox) {
                # Only set mailbox permissions once the UserMailbox exists; otherwise
                # FullAccess is recorded without AutoMapping (see Wait-ForMailboxInExchange).
                Set-UiBusyState -Busy $true -StatusText 'Waiting until mailbox of new user has been created...'
                [System.Windows.Forms.Application]::DoEvents()
                $MailboxReady = Wait-ForMailboxInExchange -Upn $Validated.NewUpn `
                    -TimeoutSeconds  $MailboxWaitTimeoutSeconds `
                    -IntervalSeconds $MailboxWaitIntervalSeconds `
                    -GraphUserId     $Created.NewUser.Id

                if (-not $MailboxReady) {
                    $Reason = if ($script:LastWaitBlocker) { $script:LastWaitBlocker } else { 'mailbox not created yet' }
                    Log-Warning ("Shared mailbox permissions NOT set: {0}." -f $Reason)
                    # Only read which permissions the source has, so the follow-up steps are concrete.
                    $Pending = @()
                    try {
                        Set-UiBusyState -Busy $true -StatusText 'Reading mailbox permissions of the source...'
                        $Pending = @(Copy-SharedMailboxPermissions -SourceUpn $SourceUpn -NewUpn $Validated.NewUpn -ListOnly | Where-Object { $_.FullAccess -or $_.SendAs })
                    } catch {
                        Log-Warning ("Mailbox permissions of the source could not be read: {0}" -f $_.Exception.Message)
                    }
                    if ($Pending.Count -eq 0) {
                        $RunIssues.Add(("Shared mailbox permissions not set ({0}). Once the mailbox exists: copy the source's permissions again (Add-MailboxPermission -AutoMapping `$true / Add-RecipientPermission)." -f $Reason))
                    } else {
                        $RunIssues.Add(("Shared mailbox permissions not set ({0}). Once the mailbox exists, add these:" -f $Reason))
                        foreach ($Item in $Pending) {
                            if ($Item.FullAccess) { $RunIssues.Add(("  Add-MailboxPermission -Identity {0} -User {1} -AccessRights FullAccess -InheritanceType All -AutoMapping `$true" -f $Item.MailboxAddress, $Validated.NewUpn)) }
                            if ($Item.SendAs)     { $RunIssues.Add(("  Add-RecipientPermission -Identity {0} -Trustee {1} -AccessRights SendAs -Confirm:`$false" -f $Item.MailboxAddress, $Validated.NewUpn)) }
                        }
                    }
                }
            }

            if ($CopyMailbox -and $MailboxReady) {
                Set-UiBusyState -Busy $true -StatusText 'Copying shared mailbox permissions...'
                [System.Windows.Forms.Application]::DoEvents()

                $MailboxResults = @(Copy-SharedMailboxPermissions -SourceUpn $SourceUpn -NewUpn $Validated.NewUpn)

                $MailboxSuccess = @($MailboxResults | Where-Object { $_.Status -eq 'Success' }).Count
                $MailboxPartial = @($MailboxResults | Where-Object { $_.Status -eq 'Partial' }).Count
                if ($MailboxResults.Count -eq 0) { Log-Info 'No shared mailbox permissions found to copy.' }
                if ($MailboxSuccess -gt 0) { Log-Info ("Mailbox permissions fully copied: {0}" -f $MailboxSuccess) }
                if ($MailboxPartial -gt 0) {
                    Log-Warning ("Mailboxes with incomplete result or scan error: {0}" -f $MailboxPartial)
                    foreach ($MailboxIssue in @($MailboxResults | Where-Object { $_.Status -eq 'Partial' })) {
                        $IssueParts = [System.Collections.Generic.List[string]]::new()
                        if (-not $MailboxIssue.FullAccessReadOk) { $IssueParts.Add('FullAccess could not be read completely') }
                        if (-not $MailboxIssue.SendAsReadOk)     { $IssueParts.Add('SendAs could not be read completely') }
                        if ($MailboxIssue.FullAccess -and -not $MailboxIssue.FullAccessAdded) { $IssueParts.Add('FullAccess not added') }
                        if ($MailboxIssue.SendAs -and -not $MailboxIssue.SendAsAdded)         { $IssueParts.Add('SendAs not added') }
                        if ($IssueParts.Count -eq 0) { $IssueParts.Add('incomplete mailbox result') }
                        $MailboxLabel = if (-not [string]::IsNullOrWhiteSpace([string]$MailboxIssue.MailboxAddress)) {
                            "$($MailboxIssue.Mailbox) <$($MailboxIssue.MailboxAddress)>"
                        } else {
                            [string]$MailboxIssue.Mailbox
                        }
                        $RunIssues.Add(("Mailbox {0}: {1}." -f $MailboxLabel, ($IssueParts -join ', ')))
                        if (-not [string]::IsNullOrWhiteSpace([string]$MailboxIssue.Detail)) {
                            Log-Warning ("Mailbox detail {0}: {1}" -f $MailboxLabel, $MailboxIssue.Detail)
                        }
                    }
                }
            }
        } catch {
            $Message = $_.Exception.Message
            Log-Error ("Exchange phase failed: {0}" -f $Message)

            if ($GroupPendingExchange.Count -gt 0 -and $DistributionResults.Count -eq 0) {
                $RunIssues.Add(("Distribution/mail-enabled groups not processed ({0}). Manually:" -f $Message))
                foreach ($PendingGroup in $GroupPendingExchange) {
                    $GroupIdentity = if (-not [string]::IsNullOrWhiteSpace([string]$PendingGroup.Mail)) { [string]$PendingGroup.Mail } else { [string]$PendingGroup.DisplayName }
                    $RunIssues.Add(("  Add-DistributionGroupMember -Identity {0} -Member {1} -BypassSecurityGroupManagerCheck" -f $GroupIdentity, $Validated.NewUpn))
                }
            }
            if ($CopyMailbox -and $MailboxResults.Count -eq 0) {
                $RunIssues.Add('Mailbox permissions could not be processed completely.')
            }
        }
    }

    # ── Phase 5: transfer Teams phone number ────────────────────────────────
    # Last of all: waiting until the new user has Teams Phone can take a long time.
    if ($TransferNumber -and $NumberPlan -and @($NumberPlan.Numbers).Count -gt 0) {
        try {
            Set-UiBusyState -Busy $true -StatusText 'Transferring phone number (waiting until Teams Phone is active)...'
            [System.Windows.Forms.Application]::DoEvents()
            $NumberResults = @(Move-PhoneNumbers -Plan $NumberPlan -NewUpn $Validated.NewUpn -NewUserId $Created.NewUser.Id)
        } catch {
            $Message = $_.Exception.Message
            Log-Error ("Number transfer failed: {0}" -f $Message)
            $RunIssues.Add(("Phone number not transferred ({0}). Use Get-CsPhoneNumberAssignment -TelephoneNumber <number> to check whether it is still assigned to the source." -f $Message))
        }

        foreach ($NumberIssue in @($NumberResults | Where-Object { $_.Status -ne 'Assigned' })) {
            if ([string]::IsNullOrWhiteSpace([string]$NumberIssue.TelephoneNumber)) {
                $RunIssues.Add(("Teams: {0}" -f $NumberIssue.Detail))
            } else {
                $RunIssues.Add(("Phone number {0}: {1}" -f $NumberIssue.TelephoneNumber, $NumberIssue.Detail))
            }
        }
    }

    # ── Final status ─────────────────────────────────────────────────────────
    $FinalStatus = if ($RunIssues.Count -gt 0) { 'Partial' } else { 'Success' }
    Log-Info ("Final status: {0}" -f $FinalStatus)
    Set-UiBusyState -Busy $false -StatusText 'Ready'

    if ($FinalStatus -eq 'Success') {
        Show-UiInfo ("Copy fully completed for: {0}" -f $Validated.NewUpn)
    } else {
        $IssueText = ($RunIssues | Select-Object -Unique | ForEach-Object { if ($_ -like '  *') { "  $_" } else { "- $_" } }) -join [Environment]::NewLine
        Show-UiWarning ("Copy completed; some items need attention for: {0}`n`n{1}`n`nThe user has been created; the other phases were carried out as far as possible. See the log for technical details." -f $Validated.NewUpn, $IssueText)
    }
}

$btnCreate.Add_Click({
    # Overarching safety net: an unexpected error must never leave the GUI disabled.
    try {
        Invoke-CopyRun
    } catch {
        try { Show-UiError ("Unexpected error during the copy: {0}`n`nCheck the log to see which phases have already been carried out." -f $_.Exception.Message) } catch { }
    } finally {
        Set-UiBusyState -Busy $false -StatusText 'Ready'
    }
})

$txtSearch.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
        $btnSearch.PerformClick()
    }
})

$script:Form.Add_FormClosing({
    # Do not close during a copy: the run would continue invisibly, and closing the
    # console can abort it between 'remove from source' and 'assign to new user'.
    if ($script:IsBusy) {
        $_.Cancel = $true
        [System.Windows.Forms.MessageBox]::Show(
            'A copy is still running. Wait until the final message appears. Do not close the PowerShell window either: that would abort the run halfway.',
            'Busy', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
})

$script:Form.Add_FormClosed({
    # Everything runs on the main thread, so there is no background cleanup
})

$script:Form.Add_Shown({
    Log-Info 'Application started. Modules are loaded on first action.'
})

[void]$script:Form.ShowDialog()
